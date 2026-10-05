#!/bin/bash
# Batch Mailer installer and updater for Mac.
#
# Paste this into Terminal and press Enter:
#   curl -LsSf https://raw.githubusercontent.com/mkarasneh07/batch-mailer/main/install.sh | bash
#
# Installs into your home folder: no admin rights and no Python needed.
# Run the same line again to update. Your connected email, sent history and
# do-not-email list are kept.
set -euo pipefail

REPO="mkarasneh07/batch-mailer"
SOURCE="${BATCH_MAILER_SOURCE:-https://github.com/$REPO/archive/refs/heads/main.tar.gz}"
DIR="${BATCH_MAILER_DIR:-$HOME/BatchMailer}"
RUN_ARGS=(run --quiet --python 3.12 --with-requirements requirements.txt streamlit run app.py)
TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT

fail() {
  echo ""
  echo "Batch Mailer couldn't be installed: $1"
  echo "Check the internet connection, then run the same line again."
  exit 1
}

echo ""
echo "Installing Batch Mailer..."
mkdir -p "$DIR/tools"

echo "  1/4  Getting the app"
curl -LsSf "$SOURCE" | tar -xz -C "$TMP" || fail "couldn't download the app"
cp -R "$TMP"/*/. "$DIR"/

if [ -x "$DIR/tools/uv" ]; then
  echo "  2/4  Engine already installed"
else
  echo "  2/4  Getting the engine that runs it"
  case "$(uname -s)-$(uname -m)" in
    Darwin-arm64)  target="aarch64-apple-darwin" ;;
    Darwin-*)      target="x86_64-apple-darwin" ;;
    Linux-x86_64)  target="x86_64-unknown-linux-gnu" ;;
    Linux-aarch64) target="aarch64-unknown-linux-gnu" ;;
    *) fail "this computer type isn't supported" ;;
  esac
  curl -LsSf "https://github.com/astral-sh/uv/releases/latest/download/uv-$target.tar.gz" \
    | tar -xz -C "$DIR/tools" --strip-components=1 || fail "couldn't download the engine"
fi

echo "  3/4  Downloading Python and components (1-2 minutes the first time)"
(cd "$DIR" && tools/uv run --quiet --python 3.12 --with-requirements requirements.txt python -c "print('ok')" >/dev/null) \
  || fail "downloading Python or its components failed"

# Skip the web framework's one-time "enter your email" question.
mkdir -p "$HOME/.streamlit"
[ -f "$HOME/.streamlit/credentials.toml" ] || printf '[general]\nemail = ""\n' > "$HOME/.streamlit/credentials.toml"

echo "  4/4  Adding Batch Mailer to the Desktop"
mkdir -p "$HOME/Desktop"
LAUNCHER="$HOME/Desktop/Batch Mailer.command"
cat > "$LAUNCHER" <<EOF
#!/bin/bash
cd "$DIR" && tools/uv ${RUN_ARGS[*]}
EOF
chmod +x "$LAUNCHER"

echo ""
echo "Done. Next time, double-click 'Batch Mailer' on your Desktop."
if [ "$(uname -s)" = "Darwin" ] && [ -z "${BATCH_MAILER_NO_START:-}" ]; then
  open "$LAUNCHER"
fi
