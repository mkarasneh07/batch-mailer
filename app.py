"""
Batch Mailer v2: send personalized emails in batches from your own mailbox.

Built for non-technical users:
  - Type your email address. The tool finds your provider by itself.
  - Microsoft 365 / Outlook: sends through the Outlook app on this computer (Windows, classic Outlook):
    no password, no sign-in, no Azure. "Sign in with Microsoft" remains as a fallback.
  - Google, Zoho, Yahoo, GoDaddy, Hostinger, company mail servers: just the password
    (or app password), with a link to the exact page where it's created.
  - The connection is remembered, so it's a one-time step.

Every email is sent from the user's real mailbox, so replies land in their normal inbox.
Run:  streamlit run app.py   (or use the start / setup scripts)
"""

from __future__ import annotations

import csv
import io
import json
import random
import re
import smtplib
import socket
import ssl
import sys
import time
import xml.etree.ElementTree as ET
from datetime import datetime
from email.headerregistry import Address
from email.message import EmailMessage
from email.utils import formatdate, make_msgid
from pathlib import Path

import msal
import pandas as pd
import requests
import streamlit as st

APP_DIR = Path(__file__).resolve().parent
LOG_FILE = APP_DIR / "sent_log.csv"
DNC_FILE = APP_DIR / "do_not_contact.txt"
TEMPLATE_FILE = APP_DIR / "last_template.json"
ACCOUNT_FILE = APP_DIR / "account.json"      # which mailbox is connected (no password in here)
SETTINGS_FILE = APP_DIR / "settings.json"    # admin settings, e.g. the Microsoft app ID
MS_CACHE_FILE = APP_DIR / ".ms_token_cache.json"
KEYRING_SERVICE = "BatchMailer"

LOG_FIELDS = ["timestamp", "campaign", "sender", "to", "subject", "status", "error"]
EMAIL_RE = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")
PLACEHOLDER_RE = re.compile(r"\{\{\s*([^}|]+?)\s*(?:\|([^}]*))?\}\}")
GRAPH_SCOPES = ["Mail.Send"]
GRAPH_SEND_URL = "https://graph.microsoft.com/v1.0/me/sendMail"
MAX_ERRORS_IN_A_ROW = 3

IS_WINDOWS = sys.platform == "win32"
OUTLOOK_NAME = "Outlook on this computer"
OL_MAIL_ITEM = 0                    # Outlook's code for "new email"
OL_FORMAT_PLAIN = 1                 # plain-text body
SEND_USING_ACCOUNT_DISPID = 64209   # Outlook's internal id for the "send from" account property
OUTLOOK_HELP = ("Make sure the classic Outlook app is installed and signed in to your email. "
                "If Outlook shows a \"New Outlook\" switch at the top right, turn it off, then try again.")

DEFAULT_CAMPAIGN = "Mailing 1"
DEFAULT_SUBJECT = "Quick question about {{Company}}"
DEFAULT_BODY = """Hi {{FirstName|there}},

[One line on why you're writing to {{Company}} specifically.]

[One line on what you do and the result it gets.]

Open to a 15-minute call next week?

[Your name]
[Title, Company]

P.S. Not relevant? Reply "no" and I won't email again."""


# ================================================================ provider detection

CONSUMER_DOMAINS = {
    "gmail.com": "google", "googlemail.com": "google",
    "outlook.com": "microsoft", "hotmail.com": "microsoft", "live.com": "microsoft", "msn.com": "microsoft",
    "yahoo.com": "yahoo", "ymail.com": "yahoo",
    "icloud.com": "icloud", "me.com": "icloud", "mac.com": "icloud",
    "zoho.com": "zoho", "zohomail.com": "zoho",
}

# Mail-server (MX) name endings that reveal the provider behind a company domain.
MX_RULES = [
    ("google.com", "google"), ("googlemail.com", "google"),
    ("outlook.com", "microsoft"),
    ("zoho.com", "zoho"), ("zoho.eu", "zoho"), ("zoho.in", "zoho"), ("zoho.com.au", "zoho"),
    ("yahoodns.net", "yahoo"),
    ("icloud.com", "icloud"),
    ("secureserver.net", "godaddy"),
    ("hostinger.com", "hostinger"),
]

GENERIC_HELP = "Use the password you normally sign in to your email with."
GENERIC_AUTH_HINT = ("The password didn't work. If your account uses 2-step verification, "
                     "your provider may need an app password instead.")

PROVIDERS = {
    "google": {
        "name": "Google (Gmail / Workspace)", "host": "smtp.gmail.com", "port": 587, "security": "STARTTLS",
        "password_label": "App password (16 letters)",
        "help": ("Google needs an **app password** here, not your normal password.\n\n"
                 "1. Click the button below.\n"
                 "2. Type *Batch Mailer* as the name and click **Create**.\n"
                 "3. Copy the 16 letters into the box below.\n\n"
                 "If the page says it isn't available, 2-Step Verification needs to be turned on first."),
        "help_url": "https://myaccount.google.com/apppasswords", "help_button": "Create a Google app password",
        "auth_hint": "Google didn't accept that. It needs an app password (16 letters), not your normal password.",
    },
    "zoho": {
        "name": "Zoho Mail", "password_label": "Zoho password",
        "help": ("Use your Zoho password. If you sign in to Zoho with 2-step verification, create an "
                 "**app password** instead: Zoho Accounts → Security → App Passwords."),
        "help_url": "https://accounts.zoho.com", "help_button": "Open Zoho Accounts",
        "auth_hint": "Zoho didn't accept that. If you use 2-step verification, you need a Zoho app password.",
    },
    "yahoo": {
        "name": "Yahoo Mail", "host": "smtp.mail.yahoo.com", "port": 465, "security": "SSL",
        "password_label": "App password",
        "help": "Yahoo needs an **app password**: Account security → Generate app password.",
        "help_url": "https://login.yahoo.com/account/security", "help_button": "Open Yahoo account security",
        "auth_hint": "Yahoo didn't accept that. It needs an app password, not your normal password.",
    },
    "icloud": {
        "name": "iCloud Mail", "host": "smtp.mail.me.com", "port": 587, "security": "STARTTLS",
        "password_label": "App-specific password",
        "help": "Apple needs an **app-specific password**: Sign-In and Security → App-Specific Passwords.",
        "help_url": "https://account.apple.com", "help_button": "Open your Apple account",
        "auth_hint": "Apple didn't accept that. It needs an app-specific password.",
    },
    "godaddy": {
        "name": "GoDaddy email", "host": "smtpout.secureserver.net", "port": 465, "security": "SSL",
        "password_label": "Email password", "help": GENERIC_HELP,
    },
    "hostinger": {
        "name": "Hostinger email", "host": "smtp.hostinger.com", "port": 465, "security": "SSL",
        "password_label": "Email password", "help": GENERIC_HELP,
    },
}


def lookup_mx(domain: str) -> list[str]:
    try:
        import dns.resolver
        answers = dns.resolver.resolve(domain, "MX", lifetime=5)
        return [str(r.exchange).rstrip(".").lower() for r in sorted(answers, key=lambda r: r.preference)]
    except Exception:  # noqa: BLE001 - no MX / no internet: fall back to other checks
        return []


def match_mx(mx_hosts: list[str]) -> tuple[str | None, str | None]:
    for host in mx_hosts:
        for suffix, key in MX_RULES:
            if host == suffix or host.endswith("." + suffix):
                return key, suffix
    return None, None


def parse_autoconfig(xml_text: str, email: str) -> dict | None:
    """Read Thunderbird-style mail settings (the same database Thunderbird uses to set up accounts)."""
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError:
        return None
    local, domain = email.split("@", 1)
    candidates = []
    for server in root.iter("outgoingServer"):
        if server.get("type") != "smtp":
            continue
        security = {"SSL": "SSL", "STARTTLS": "STARTTLS"}.get((server.findtext("socketType") or "").strip().upper())
        host = (server.findtext("hostname") or "").strip()
        if not security or not host:
            continue
        username = ((server.findtext("username") or "%EMAILADDRESS%").strip()
                    .replace("%EMAILADDRESS%", email).replace("%EMAILLOCALPART%", local)
                    .replace("%EMAILDOMAIN%", domain))
        candidates.append({"host": host.replace("%EMAILDOMAIN%", domain),
                           "port": int(server.findtext("port") or 587), "security": security, "username": username})
    if not candidates:
        return None
    return {"name": (root.findtext(".//displayName") or "").strip(), "candidates": candidates}


def lookup_autoconfig(email: str, domain: str, mx_hosts: list[str]) -> dict | None:
    urls = [f"https://autoconfig.thunderbird.net/v1.1/{domain}"]
    if mx_hosts:
        mx_base = ".".join(mx_hosts[0].split(".")[-2:])
        if mx_base != domain:
            urls.append(f"https://autoconfig.thunderbird.net/v1.1/{mx_base}")
    urls.append(f"https://autoconfig.{domain}/mail/config-v1.1.xml?emailaddress={email}")
    for url in urls:
        try:
            response = requests.get(url, timeout=5)
            if response.ok:
                found = parse_autoconfig(response.text, email)
                if found:
                    return found
        except requests.RequestException:
            continue
    return None


def zoho_server(domain: str, mx_suffix: str | None) -> str:
    region = (mx_suffix or "zoho.com").split("zoho.", 1)[1]           # com, eu, in, com.au
    personal = domain in ("zoho.com", "zohomail.com")
    return f"{'smtp' if personal else 'smtppro'}.zoho.{region}"


def detect_provider(email: str) -> dict:
    """Work out the provider from the email address. Returns what the connect screen needs."""
    domain = email.split("@", 1)[1].lower().strip()
    key, suffix = CONSUMER_DOMAINS.get(domain), None
    mx_hosts: list[str] = []
    if not key:
        mx_hosts = lookup_mx(domain)
        key, suffix = match_mx(mx_hosts)

    if key == "microsoft":
        return {"email": email, "kind": "microsoft", "name": "Microsoft 365 / Outlook"}

    if key in PROVIDERS:
        p = dict(PROVIDERS[key])
        host = zoho_server(domain, suffix) if key == "zoho" else p["host"]
        port, security = (465, "SSL") if key == "zoho" else (p["port"], p["security"])
        return {"email": email, "kind": "smtp", "name": p["name"], "password_label": p["password_label"],
                "help": p["help"], "help_url": p.get("help_url"), "help_button": p.get("help_button"),
                "auth_hint": p.get("auth_hint", GENERIC_AUTH_HINT),
                "candidates": [{"host": host, "port": port, "security": security, "username": email}]}

    found = lookup_autoconfig(email, domain, mx_hosts)
    guesses = [
        {"host": f"smtp.{domain}", "port": 587, "security": "STARTTLS", "username": email},
        {"host": f"mail.{domain}", "port": 465, "security": "SSL", "username": email},
        {"host": f"mail.{domain}", "port": 587, "security": "STARTTLS", "username": email},
    ]
    return {"email": email, "kind": "smtp", "name": (found or {}).get("name") or "your company's email server",
            "password_label": "Email password", "help": GENERIC_HELP, "auth_hint": GENERIC_AUTH_HINT,
            "candidates": (found["candidates"] if found else []) + guesses, "guessed": not found}


# ================================================================ saved account, password, settings

def read_json(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
    except (ValueError, OSError):
        return {}


def write_json(path: Path, data: dict) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def save_password(email: str, password: str) -> bool:
    """Store in Windows Credential Manager / macOS Keychain. Falls back to this session only."""
    st.session_state[f"pw:{email.lower()}"] = password
    try:
        import keyring
        keyring.set_password(KEYRING_SERVICE, email.lower(), password)
        return True
    except Exception:  # noqa: BLE001 - no secure store available on this computer
        return False


def get_password(email: str) -> str | None:
    if f"pw:{email.lower()}" in st.session_state:
        return st.session_state[f"pw:{email.lower()}"]
    try:
        import keyring
        return keyring.get_password(KEYRING_SERVICE, email.lower())
    except Exception:  # noqa: BLE001
        return None


def forget_password(email: str) -> None:
    st.session_state.pop(f"pw:{email.lower()}", None)
    try:
        import keyring
        keyring.delete_password(KEYRING_SERVICE, email.lower())
    except Exception:  # noqa: BLE001
        pass


# ================================================================ message templates

def render(template: str, row: dict) -> str:
    """Replace {{Column}} and {{Column|backup}} with values from the row."""
    def fill(match: re.Match) -> str:
        value = str(row.get(match.group(1).strip(), "") or "").strip()
        return value or (match.group(2) or "").strip()
    return PLACEHOLDER_RE.sub(fill, template)


def placeholders_in(*templates: str) -> set[str]:
    return {m.group(1).strip() for t in templates for m in PLACEHOLDER_RE.finditer(t)}


# ================================================================ log + do-not-email list

def read_log() -> pd.DataFrame:
    if not LOG_FILE.exists():
        return pd.DataFrame(columns=LOG_FIELDS, dtype=str)
    return pd.read_csv(LOG_FILE, dtype=str, keep_default_na=False)


def append_log(**entry: str) -> None:
    is_new = not LOG_FILE.exists()
    with LOG_FILE.open("a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=LOG_FIELDS)
        if is_new:
            writer.writeheader()
        writer.writerow({"timestamp": datetime.now().isoformat(timespec="seconds"), **entry})


def sent_today(log: pd.DataFrame, sender: str) -> int:
    if log.empty or not sender:
        return 0
    today = datetime.now().date().isoformat()
    mask = ((log["status"] == "sent") & (log["sender"].str.lower() == sender.lower())
            & log["timestamp"].str.startswith(today))
    return int(mask.sum())


def already_emailed(log: pd.DataFrame, campaign: str | None) -> set[str]:
    if log.empty:
        return set()
    mask = log["status"] == "sent"
    if campaign is not None:
        mask &= log["campaign"] == campaign
    return set(log.loc[mask, "to"].str.lower())


def load_do_not_contact() -> tuple[set[str], set[str]]:
    if not DNC_FILE.exists():
        return set(), set()
    lines = {ln.strip().lower() for ln in DNC_FILE.read_text(encoding="utf-8").splitlines() if ln.strip()}
    return {ln for ln in lines if not ln.startswith("@")}, {ln[1:] for ln in lines if ln.startswith("@")}


def build_queue(df: pd.DataFrame, email_col: str, campaign: str, log: pd.DataFrame, skip_any: bool):
    emailed = already_emailed(log, None if skip_any else campaign)
    dnc_emails, dnc_domains = load_do_not_contact()
    seen: set[str] = set()
    queue: list[dict] = []
    skipped = {"empty or invalid email": 0, "listed twice": 0, "already emailed": 0, "on the do-not-email list": 0}
    for row in df.to_dict("records"):
        email = str(row.get(email_col, "")).strip()
        key = email.lower()
        if not EMAIL_RE.match(email):
            skipped["empty or invalid email"] += 1
        elif key in seen:
            skipped["listed twice"] += 1
        elif key in emailed:
            skipped["already emailed"] += 1
        elif key in dnc_emails or key.split("@")[1] in dnc_domains:
            skipped["on the do-not-email list"] += 1
        else:
            queue.append(row)
        seen.add(key)
    return queue, skipped


# ================================================================ sending

def build_mime(from_email: str, from_name: str, reply_to: str, to: str, subject: str, body: str) -> EmailMessage:
    msg = EmailMessage()
    msg["From"] = Address(display_name=from_name or "", addr_spec=from_email)
    msg["To"] = to
    msg["Subject"] = subject
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain=from_email.split("@")[-1])
    if reply_to:
        msg["Reply-To"] = reply_to
    msg.set_content(body)  # plain text: best inbox placement for one-to-one email
    return msg


class SmtpSender:
    def __init__(self, host, port, security, username, password, from_email, from_name="", reply_to="", timeout=30):
        self.host, self.port, self.security = host, int(port), security
        self.username, self.password = username or from_email, password
        self.from_email, self.from_name, self.reply_to = from_email, from_name, reply_to
        self.timeout = timeout
        self.server: smtplib.SMTP | None = None

    def connect(self) -> None:
        self.close()
        context = ssl.create_default_context()
        if self.security == "SSL":
            server = smtplib.SMTP_SSL(self.host, self.port, timeout=self.timeout, context=context)
        else:
            server = smtplib.SMTP(self.host, self.port, timeout=self.timeout)
            if self.security == "STARTTLS":
                server.starttls(context=context)
        if self.password:
            server.login(self.username, self.password)
        self.server = server

    def _is_alive(self) -> bool:
        if self.server is None:
            return False
        try:
            return self.server.noop()[0] == 250
        except (smtplib.SMTPException, OSError):
            return False

    def send(self, to: str, subject: str, body: str) -> None:
        if not self._is_alive():
            self.connect()
        self.server.send_message(build_mime(self.from_email, self.from_name, self.reply_to, to, subject, body))

    def close(self) -> None:
        if self.server is not None:
            try:
                self.server.quit()
            except (smtplib.SMTPException, OSError):
                pass
            self.server = None


@st.cache_resource(show_spinner=False)
def microsoft_client(client_id: str, tenant: str = "common"):
    cache = msal.SerializableTokenCache()
    if MS_CACHE_FILE.exists():
        cache.deserialize(MS_CACHE_FILE.read_text(encoding="utf-8"))
    app = msal.PublicClientApplication(
        client_id, authority=f"https://login.microsoftonline.com/{tenant or 'common'}", token_cache=cache)
    return app, cache


def save_microsoft_cache(cache: msal.SerializableTokenCache) -> None:
    if cache.has_state_changed:
        MS_CACHE_FILE.write_text(cache.serialize(), encoding="utf-8")


def microsoft_account(email: str):
    """The signed-in Microsoft account for this email, or None."""
    settings = read_json(SETTINGS_FILE)
    if not settings.get("microsoft_client_id"):
        return None, None, None
    try:
        app, cache = microsoft_client(settings["microsoft_client_id"], settings.get("microsoft_tenant", "common"))
        accounts = app.get_accounts(username=email) or app.get_accounts()
    except Exception:  # noqa: BLE001 - e.g. no internet right now
        return None, None, None
    return app, cache, (accounts[0] if accounts else None)


class GraphSender:
    """Sends through Microsoft Graph as the signed-in user."""

    def __init__(self, email: str, reply_to: str = ""):
        self.app, self.cache, self.account = microsoft_account(email)
        if self.account is None:
            raise ValueError("Microsoft sign-in is missing. Click Disconnect at the top and sign in again.")
        self.from_email = self.account.get("username", email)
        self.reply_to = reply_to

    def _token(self) -> str:
        result = self.app.acquire_token_silent(GRAPH_SCOPES, account=self.account)
        save_microsoft_cache(self.cache)
        if not result or "access_token" not in result:
            raise ValueError("Your Microsoft sign-in expired. Click Disconnect at the top and sign in again.")
        return result["access_token"]

    def connect(self) -> None:
        self._token()

    def send(self, to: str, subject: str, body: str) -> None:
        message = {"subject": subject, "body": {"contentType": "Text", "content": body},
                   "toRecipients": [{"emailAddress": {"address": to}}]}
        if self.reply_to:
            message["replyTo"] = [{"emailAddress": {"address": self.reply_to}}]
        for _ in range(3):
            response = requests.post(GRAPH_SEND_URL, headers={"Authorization": f"Bearer {self._token()}"},
                                     json={"message": message, "saveToSentItems": True}, timeout=30)
            if response.status_code == 202:
                return
            if response.status_code in (429, 503, 504):  # Microsoft says slow down: wait, then retry
                time.sleep(int(response.headers.get("Retry-After", "10")))
                continue
            raise RuntimeError(f"Microsoft didn't send it ({response.status_code}): {response.text[:300]}")
        raise RuntimeError("Microsoft is limiting this mailbox right now. Try again later.")

    def close(self) -> None:
        pass


# ---------------------------------------------------------------- Outlook on this computer (Windows)

def outlook_available() -> bool:
    """True when classic Outlook is installed and can be driven by other programs."""
    if not IS_WINDOWS:
        return False
    try:
        import winreg
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Outlook.Application\CLSID"))
        import win32com.client  # noqa: F401
        return True
    except Exception:  # noqa: BLE001 - not installed, or new Outlook only
        return False


def open_outlook():
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()  # each Streamlit run is its own thread; Outlook needs this once per thread
    return win32com.client.Dispatch("Outlook.Application")


def outlook_accounts(app) -> list:
    accounts = app.Session.Accounts
    return [accounts.Item(i) for i in range(1, accounts.Count + 1)]


def account_address(account) -> str:
    try:
        return (account.SmtpAddress or "").strip()
    except Exception:  # noqa: BLE001
        return ""


def outlook_addresses() -> list[str]:
    return [a for a in (account_address(acc) for acc in outlook_accounts(open_outlook())) if a]


def set_send_account(mail, account, email: str) -> None:
    try:
        mail.SendUsingAccount = account
    except Exception:  # noqa: BLE001 - some setups refuse the normal way; this is the known workaround
        mail._oleobj_.Invoke(SEND_USING_ACCOUNT_DISPID, 0, 8, 0, account)
    try:
        current = mail.SendUsingAccount
        chosen = account_address(current) if current is not None else ""
    except Exception:  # noqa: BLE001 - can't read it back; trust the setting
        chosen = ""
    if chosen and chosen.lower() != email.lower():
        raise ValueError(f"Outlook wouldn't send from {email}. In Outlook, make it the default account "
                         "(File → Account Settings), then try again.")


class OutlookSender:
    """Hands each email to the classic Outlook app, which sends it from the user's own account."""

    def __init__(self, email: str, reply_to: str = ""):
        self.from_email, self.reply_to = email, reply_to
        self.app = None
        self.account = None
        self.several_accounts = False

    def connect(self) -> None:
        self.app = open_outlook()
        accounts = outlook_accounts(self.app)
        self.account = next((a for a in accounts if account_address(a).lower() == self.from_email.lower()), None)
        if self.account is None:
            raise ValueError(f"Outlook on this computer is no longer set up with {self.from_email}. "
                             "Click Disconnect at the top and connect again.")
        self.several_accounts = len(accounts) > 1

    def send(self, to: str, subject: str, body: str) -> None:
        if self.app is None:
            self.connect()
        mail = self.app.CreateItem(OL_MAIL_ITEM)
        mail.To = to
        mail.Subject = subject
        mail.BodyFormat = OL_FORMAT_PLAIN
        mail.Body = body
        if self.reply_to:
            mail.ReplyRecipients.Add(self.reply_to)
        if self.several_accounts:
            set_send_account(mail, self.account, self.from_email)
        mail.Send()

    def close(self) -> None:
        self.app = None
        self.account = None


def outlook_error_text(err: Exception) -> str:
    if isinstance(err, ModuleNotFoundError):
        return "The Outlook connector isn't installed. Close Batch Mailer and open it again from the desktop shortcut."
    code = err.args[0] if err.args else None
    if code in (-2147221005, -2147221164):  # Outlook isn't registered on this computer
        return "Classic Outlook isn't available on this computer. " + OUTLOOK_HELP
    if code == -2146959355:  # Outlook refused to start for another program
        return "Outlook couldn't be opened by Batch Mailer. Close Outlook completely, open it again, then try again."
    info = getattr(err, "excepinfo", None)
    detail = info[2] if info and len(info) > 2 and info[2] else str(err)
    return f"Outlook said: {detail}"


def sender_for(account: dict, reply_to: str = ""):
    if account["kind"] == "outlook":
        return OutlookSender(account["email"], reply_to)
    if account["kind"] == "microsoft":
        return GraphSender(account["email"], reply_to)
    password = get_password(account["email"])
    return SmtpSender(account["host"], account["port"], account["security"], account.get("username"),
                      password or "", account["email"], account.get("name", ""), reply_to)


def try_connect(candidates: list[dict], email: str, password: str, name: str):
    """Try each server until one works. Returns (working server, error)."""
    last_error: Exception | None = None
    for server in candidates:
        sender = SmtpSender(server["host"], server["port"], server["security"], server.get("username") or email,
                            password, email, name, timeout=12)
        try:
            sender.connect()
            sender.close()
            return server, None
        except smtplib.SMTPAuthenticationError as err:  # right server, wrong password: stop here
            return None, err
        except Exception as err:  # noqa: BLE001 - try the next server
            last_error = err
    return None, last_error


def friendly_error(err: Exception, auth_hint: str = GENERIC_AUTH_HINT) -> str:
    if err.__class__.__name__ == "com_error":
        return outlook_error_text(err)
    if isinstance(err, smtplib.SMTPAuthenticationError):
        return auth_hint
    if isinstance(err, socket.gaierror):
        return "Couldn't find that email server."
    if isinstance(err, ssl.SSLError):
        return "The secure connection failed. Try 587 with STARTTLS, or 465 with SSL."
    if isinstance(err, (TimeoutError, ConnectionRefusedError, ConnectionResetError)):
        return "Couldn't reach the email server. Check the internet connection, or the server details."
    if isinstance(err, smtplib.SMTPRecipientsRefused):
        return "The email server says this address doesn't exist."
    return str(err) or err.__class__.__name__


# ================================================================ UI: connect

def connected_bar(account: dict) -> bool:
    """Shows who we're sending as. Returns True if ready to send."""
    if account["kind"] == "outlook":
        ready = True
    elif account["kind"] == "microsoft":
        _, _, ms_account = microsoft_account(account["email"])
        ready = ms_account is not None
    else:
        ready = get_password(account["email"]) is not None

    col1, col2 = st.columns([5, 1])
    who = f"{account.get('name')} <{account['email']}>" if account.get("name") else account["email"]
    col1.markdown(f"Sending as **{who}**  \n{account['provider_name']}")
    if col2.button("Disconnect"):
        forget_password(account["email"])
        if account["kind"] == "microsoft":
            app, cache, ms_account = microsoft_account(account["email"])
            if ms_account:
                app.remove_account(ms_account)
                save_microsoft_cache(cache)
        ACCOUNT_FILE.unlink(missing_ok=True)
        st.session_state.pop("detected", None)
        st.session_state.pop("outlook_found", None)
        st.rerun()

    if account["kind"] == "outlook":
        st.caption("Keep Outlook open while sending. Emails leave from Outlook's Outbox and appear in your "
                   "Sent Items. If Outlook asks whether to allow a program to send email, click **Allow**.")

    if not ready and account["kind"] == "smtp":
        password = st.text_input(f"Enter the password for {account['email']} to continue", type="password")
        if password:
            save_password(account["email"], password)
            st.rerun()
    elif not ready:
        st.warning("The Microsoft sign-in is missing. Click Disconnect, then sign in again.")
    return ready


def connect_screen() -> None:
    st.subheader("Connect your email")
    st.write("Type the email address you want to send from. The tool will work out the rest.")
    email = st.text_input("Your email address", key="connect_email", placeholder="name@company.com").strip()
    if st.button("Continue", type="primary"):
        if not EMAIL_RE.match(email):
            st.error("That doesn't look like an email address.")
            return
        with st.spinner("Finding your email provider…"):
            st.session_state.detected = detect_provider(email)

    detected = st.session_state.get("detected")
    if not detected or detected["email"].lower() != email.lower():
        return

    has_outlook = outlook_available()
    if detected["kind"] == "microsoft" and has_outlook:
        outlook_connect(detected["email"], primary=True)
        if read_json(SETTINGS_FILE).get("microsoft_client_id"):
            st.divider()
            st.caption("Or, instead of Outlook, sign in with Microsoft:")
            microsoft_connect(detected, show_intro=False)
    elif detected["kind"] == "microsoft":
        microsoft_connect(detected)
    else:
        smtp_connect(detected)
        if has_outlook:
            outlook_connect(detected["email"], primary=False)


def save_outlook_account(address: str) -> None:
    write_json(ACCOUNT_FILE, {"kind": "outlook", "email": address, "provider_name": OUTLOOK_NAME})
    st.session_state.pop("outlook_found", None)
    st.rerun()


def outlook_connect(email: str, primary: bool) -> None:
    if primary:
        st.success("Your email is on Microsoft 365. Batch Mailer can send through the Outlook app on this "
                   "computer, with no password or sign-in.")
        clicked = st.button("Use Outlook on this computer", type="primary")
    else:
        st.divider()
        st.caption("Already use the Outlook app on this computer for this email? You can send through it instead.")
        clicked = st.button("Use Outlook on this computer instead")

    if clicked:
        with st.spinner("Opening Outlook…"):
            try:
                st.session_state.outlook_found = outlook_addresses()
            except Exception as err:  # noqa: BLE001
                st.session_state.pop("outlook_found", None)
                st.error(outlook_error_text(err))
                return

    found = st.session_state.get("outlook_found")
    if found is None:
        return
    match = next((a for a in found if a.lower() == email.lower()), None)
    if match:
        save_outlook_account(match)
    if not found:
        st.error("Outlook opened, but it has no email account set up. Open Outlook, sign in to your email, "
                 "then try again.")
        return
    st.warning(f"Outlook on this computer isn't set up with {email}. It has: {', '.join(found)}.")
    choice = st.selectbox("Send from one of these Outlook accounts instead?", found, key="outlook_choice")
    if st.button("Use this account"):
        save_outlook_account(choice)


def microsoft_connect(detected: dict, show_intro: bool = True) -> None:
    if show_intro:
        st.success("Your email is on Microsoft 365 / Outlook. Sign in the way you normally do.")
    settings = read_json(SETTINGS_FILE)
    if not settings.get("microsoft_client_id"):
        st.warning("Microsoft sign-in isn't switched on for this computer yet. On a Windows computer with the "
                   "classic Outlook app, Batch Mailer can send through Outlook instead, with no setup. "
                   "Otherwise, ask the person who installed this tool to add the Microsoft app ID in "
                   "**Admin settings** (bottom of the left panel).")
        return
    if not st.button("Sign in with Microsoft", type="primary" if show_intro else "secondary"):
        return
    app, cache = microsoft_client(settings["microsoft_client_id"], settings.get("microsoft_tenant", "common"))
    st.info("A Microsoft sign-in page opened in your browser. Finish signing in there, then come back to this tab.")
    try:
        result = app.acquire_token_interactive(GRAPH_SCOPES, login_hint=detected["email"],
                                               prompt="select_account", timeout=300)
    except Exception as err:  # noqa: BLE001
        st.error(f"Sign-in didn't finish: {err}")
        return
    if "access_token" not in result:
        st.error(f"Sign-in didn't finish: {result.get('error_description', result.get('error'))}")
        return
    save_microsoft_cache(cache)
    signed_in_as = result.get("id_token_claims", {}).get("preferred_username") or detected["email"]
    write_json(ACCOUNT_FILE, {"kind": "microsoft", "email": signed_in_as, "provider_name": detected["name"]})
    st.rerun()


def smtp_connect(detected: dict) -> None:
    st.success(f"Your email is on {detected['name']}.")
    st.markdown(detected["help"])
    if detected.get("help_url"):
        st.link_button(detected["help_button"], detected["help_url"])

    name = st.text_input("Your name, as people will see it", key="connect_name", placeholder="e.g. Sara Haddad")
    password = st.text_input(detected["password_label"], type="password", key="connect_password")

    with st.expander("Enter server details yourself", expanded=st.session_state.get("show_manual", False)):
        st.caption("Only needed if connecting fails. Your IT team or email host can give you these.")
        manual_host = st.text_input("Outgoing (SMTP) server", key="manual_host", placeholder="smtp.yourcompany.com")
        col1, col2 = st.columns(2)
        manual_port = col1.number_input("Port", value=587, min_value=1, max_value=65535, step=1, key="manual_port")
        manual_security = col2.selectbox("Security", ["STARTTLS", "SSL", "None"], key="manual_security")

    if not st.button("Connect", type="primary"):
        return
    candidates = ([{"host": manual_host.strip(), "port": int(manual_port), "security": manual_security,
                    "username": detected["email"]}] if manual_host.strip() else detected["candidates"])
    with st.spinner("Connecting…"):
        server, error = try_connect(candidates, detected["email"], password, name.strip())
    if server is None:
        if isinstance(error, smtplib.SMTPAuthenticationError):
            st.error(detected["auth_hint"])
        else:
            st.session_state.show_manual = True
            st.error("Couldn't connect automatically. Open **Enter server details yourself** above "
                     f"and fill in the server your IT team gives you. ({friendly_error(error)})")
        return

    remembered = save_password(detected["email"], password)
    write_json(ACCOUNT_FILE, {"kind": "smtp", "email": detected["email"], "name": name.strip(),
                              "provider_name": detected["name"], **server})
    if not remembered:
        st.session_state.flash = ("This computer couldn't store the password securely, "
                                  "so you'll be asked for it each time you open the tool.")
    st.rerun()


# ================================================================ UI: sidebar settings

def sidebar_settings() -> dict:
    sb = st.sidebar
    sb.header("Sending settings")
    sb.text_input("Mailing name", key="campaign",
                  help="Saved with every email. Use a new name for a follow-up round.")
    sb.number_input("Most emails per day", min_value=1, max_value=2000, step=5, key="daily_limit",
                    help="For sales emails to people who don't know you, 30 to 50 a day is safest.")
    sb.slider("Seconds between emails", 1, 300, key="delay",
              help="Each pause is a random length inside this range, so the rhythm isn't machine-regular.")
    if st.session_state.delay[0] < 10:
        sb.caption("Fast pace. Fine for people who know you. For cold sales email, many emails in a few "
                   "minutes is a pattern spam filters notice, so keep the daily limit low.")
    sb.checkbox("Skip anyone I've already emailed", key="skip_any",
                help="Untick only for a follow-up round, after adding everyone who replied to the do-not-email list.")
    sb.text_input("Replies go to (optional)", key="reply_to",
                  help="Leave empty so replies come back to the address you send from.")

    with sb.expander("Admin settings"):
        st.caption("For the person who installs the tool. Not needed when sending through Outlook. "
                   "Only for Microsoft sign-in on computers without the classic Outlook app.")
        settings = read_json(SETTINGS_FILE)
        client_id = st.text_input("Microsoft app (client) ID", value=settings.get("microsoft_client_id", ""))
        tenant = st.text_input("Microsoft tenant", value=settings.get("microsoft_tenant", "common"))
        if st.button("Save admin settings"):
            write_json(SETTINGS_FILE, {**settings, "microsoft_client_id": client_id.strip(),
                                       "microsoft_tenant": tenant.strip() or "common"})
            st.success("Saved.")

    return {"campaign": st.session_state.campaign.strip() or DEFAULT_CAMPAIGN,
            "daily_limit": int(st.session_state.daily_limit), "delay": st.session_state.delay,
            "skip_any": st.session_state.skip_any, "reply_to": st.session_state.reply_to.strip()}


# ================================================================ UI: steps

@st.cache_data(show_spinner=False)
def read_table(name: str, data: bytes) -> pd.DataFrame:
    if name.lower().endswith(".csv"):
        for encoding in ("utf-8-sig", "cp1256", "latin-1"):  # cp1256 = Arabic CSVs saved by Excel
            try:
                df = pd.read_csv(io.BytesIO(data), dtype=str, keep_default_na=False, encoding=encoding)
                break
            except UnicodeDecodeError:
                continue
    else:
        df = pd.read_excel(io.BytesIO(data), dtype=str).fillna("")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def send_batch(sender, batch, email_col, campaign, subject, body, delay) -> None:
    st.caption("Keep this tab open while it sends. To stop, click **Stop** at the top right. "
               "People already emailed are skipped next time.")
    progress = st.progress(0.0)
    status = st.empty()
    table = st.empty()
    results: list[dict] = []
    sent = failed = errors_in_a_row = 0
    stopped_early = False

    for i, row in enumerate(batch, start=1):
        to = str(row[email_col]).strip()
        subj = render(subject, row)
        status.write(f"Sending {i} of {len(batch)} to {to}…")
        try:
            sender.send(to, subj, render(body, row))
            append_log(campaign=campaign, sender=sender.from_email, to=to, subject=subj, status="sent", error="")
            sent += 1
            errors_in_a_row = 0
            results.append({"To": to, "Result": "Sent"})
        except Exception as err:  # noqa: BLE001
            message = friendly_error(err)
            append_log(campaign=campaign, sender=sender.from_email, to=to, subject=subj,
                       status="failed", error=message[:300])
            failed += 1
            errors_in_a_row += 1
            results.append({"To": to, "Result": f"Not sent: {message[:150]}"})
            if isinstance(err, (smtplib.SMTPAuthenticationError, ValueError)) or errors_in_a_row >= MAX_ERRORS_IN_A_ROW:
                stopped_early = True

        progress.progress(i / len(batch))
        table.dataframe(pd.DataFrame(results[::-1]), hide_index=True)
        if stopped_early:
            status.error(f"Stopped after {errors_in_a_row} email(s) in a row didn't send. "
                         f"Last problem: {results[-1]['Result']}")
            break
        if i < len(batch):
            end = time.time() + random.uniform(*delay)
            while (left := end - time.time()) > 0:
                status.write(f"Sent {sent} of {len(batch)}. Next email in {int(left) + 1}s…")
                time.sleep(min(1.0, left))

    sender.close()
    if not stopped_early:
        status.success(f"Finished. Sent {sent}" + (f", {failed} didn't send." if failed else "."))


def steps(account: dict, opts: dict) -> None:
    st.subheader("1. Your list")
    upload = st.file_uploader("Upload your list (Excel or CSV)", type=["xlsx", "csv"])
    if upload is None:
        st.info("Use a file with one row per person and a column of email addresses. "
                "Other columns, like FirstName or Company, can be added to your message.")
        return
    df = read_table(upload.name, upload.getvalue())
    if df.empty:
        st.warning("That file has no rows.")
        return
    columns = list(df.columns)
    guess = next((i for i, c in enumerate(columns) if "mail" in c.lower()), 0)
    email_col = st.selectbox("Column with the email addresses", columns, index=guess)
    with st.expander(f"See your list ({len(df)} rows)"):
        st.dataframe(df.head(50), hide_index=True)

    st.subheader("2. Your message")
    subject = st.text_input("Subject", key="subject")
    body = st.text_area("Message", key="body", height=320)
    st.caption("To add each person's details, type a column name in double curly brackets, like {{FirstName}}. "
               "If a cell might be empty, add a backup word: {{FirstName|there}}.  \nYour columns: "
               + ", ".join("{{" + c + "}}" for c in columns))
    unknown = placeholders_in(subject, body) - set(columns)
    if unknown:
        st.warning("There's no column called " + ", ".join("{{" + u + "}}" for u in sorted(unknown))
                   + ". Check the spelling; it must match the column name exactly.")

    st.subheader("3. Check it")
    log = read_log()
    queue, skipped = build_queue(df, email_col, opts["campaign"], log, opts["skip_any"])
    skipped_text = ", ".join(f"{n} {reason}" for reason, n in skipped.items() if n)
    if skipped_text:
        st.caption(f"Left out: {skipped_text}.")
    if not queue:
        st.info("Everyone on this list has already been emailed, or is on the do-not-email list.")
        return
    pick = st.number_input(f"Preview person (1 to {len(queue)})", min_value=1, max_value=len(queue), value=1, step=1)
    row = queue[int(pick) - 1]
    with st.container(border=True):
        st.markdown(f"**To:** {row[email_col]}  \n**Subject:** {render(subject, row)}")
        st.text(render(body, row))
    if st.button("Send this as a test to myself"):
        try:
            sender = sender_for(account, opts["reply_to"])
            sender.connect()
            sender.send(sender.from_email, "[TEST] " + render(subject, row), render(body, row))
            sender.close()
            write_json(TEMPLATE_FILE, {"subject": subject, "body": body})
            st.success(f"Test sent to {sender.from_email}. Check that it arrived in your inbox (not spam) "
                       "and the details filled in correctly.")
        except Exception as err:  # noqa: BLE001
            st.error(friendly_error(err))

    st.subheader("4. Send")
    done_today = sent_today(log, account["email"])
    to_send = min(len(queue), max(opts["daily_limit"] - done_today, 0))
    people = "1 person" if len(queue) == 1 else f"{len(queue)} people"
    if to_send == 0:
        st.write(f"{people} ready, but today's limit of {opts['daily_limit']} is reached. "
                 "Come back tomorrow, or raise the limit in the left panel.")
    elif to_send < len(queue):
        st.write(f"{people} ready. Today's limit lets you send {to_send} now; the rest can go tomorrow.")
    else:
        st.write(f"{people} ready to send.")
    if st.button(f"Start sending ({to_send})", type="primary", disabled=to_send == 0):
        write_json(TEMPLATE_FILE, {"subject": subject, "body": body})
        try:
            sender = sender_for(account, opts["reply_to"])
            sender.connect()
        except Exception as err:  # noqa: BLE001
            st.error(friendly_error(err))
            return
        send_batch(sender, queue[:to_send], email_col, opts["campaign"], subject, body, opts["delay"])


def history_section() -> None:
    st.divider()
    with st.expander("Emails sent so far"):
        log = read_log()
        log = log[log["status"] != "test"] if not log.empty else log
        if log.empty:
            st.write("Nothing sent yet.")
        else:
            st.dataframe(log.iloc[::-1].head(300), hide_index=True)
            st.download_button("Download the full list", LOG_FILE.read_bytes(), "sent_log.csv", "text/csv")
    with st.expander("Do-not-email list"):
        st.caption("One email address per line. Add @company.com to leave out a whole company. "
                   "Nobody on this list is ever emailed. Add people here when they reply \"no\".")
        current = DNC_FILE.read_text(encoding="utf-8") if DNC_FILE.exists() else ""
        text = st.text_area("Addresses", value=current, height=160, key="dnc_text", label_visibility="collapsed")
        if st.button("Save list"):
            DNC_FILE.write_text(text.strip() + "\n", encoding="utf-8")
            st.success("Saved.")


def init_state() -> None:
    if "subject" in st.session_state:
        return
    saved = read_json(TEMPLATE_FILE)
    st.session_state.update({
        "subject": saved.get("subject", DEFAULT_SUBJECT), "body": saved.get("body", DEFAULT_BODY),
        "campaign": DEFAULT_CAMPAIGN, "daily_limit": 40, "delay": (3, 6), "skip_any": True, "reply_to": "",
    })


def main() -> None:
    st.set_page_config(page_title="Batch Mailer", page_icon="✉️", layout="centered")
    init_state()
    st.title("Batch Mailer")
    st.caption("Send personal emails to a list, from your own mailbox. Replies come back to your normal inbox.")

    opts = sidebar_settings()
    if "flash" in st.session_state:
        st.info(st.session_state.pop("flash"))

    account = read_json(ACCOUNT_FILE)
    if not account:
        connect_screen()
    elif connected_bar(account):
        steps(account, opts)
    history_section()


if __name__ == "__main__":
    main()
