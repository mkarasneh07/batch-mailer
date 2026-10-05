#!/bin/bash
# Batch Mailer for Mac. Double-click to open. No Python needed.
# First run downloads a small engine (uv) that fetches Python and everything else (1-2 minutes, once).
cd "$(dirname "$0")"

if [ ! -x tools/uv ]; then
  echo "First-time setup. This takes a minute or two, only once..."
  if [ "$(uname -m)" = "arm64" ]; then target="aarch64-apple-darwin"; else target="x86_64-apple-darwin"; fi
  mkdir -p tools
  if ! curl -LsSf "https://github.com/astral-sh/uv/releases/latest/download/uv-$target.tar.gz" | tar -xz -C tools --strip-components=1; then
    echo "Setup couldn't download what it needs. Check the internet connection and try again."
    read -r -p "Press Enter to close."
    exit 1
  fi
fi

mkdir -p ~/.streamlit
[ -f ~/.streamlit/credentials.toml ] || printf '[general]\nemail = ""\n' > ~/.streamlit/credentials.toml

echo "Opening Batch Mailer in your browser..."
tools/uv run --quiet --python 3.12 --with-requirements requirements.txt streamlit run app.py
