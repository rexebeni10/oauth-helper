# Outlook Mail Reader

Personal tool: authorize **your** Outlook mailbox in the browser, save a local `token.json`, then print recent verification codes with Python + Microsoft Graph.

Live helper: [native/index.html](https://rexebeni10.github.io/oauth-helper/native/) (GitHub Pages).

## What it does

1. OAuth2 + PKCE in a static page (no password stored here).
2. Microsoft issues a refresh token (`Mail.Read`).
3. You download `token.json` to this folder (never commit it).
4. `outlook_otp.py` lists Inbox + Junk and extracts 4–8 digit codes.

## Setup

```powershell
cd oauth-helper
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Copy `token.example.json` to `token.json` and replace the refresh token after you complete the helper page, **or** use the page’s Download button.

```powershell
python outlook_otp.py --minutes 10
```

Revoke the app any time: [Microsoft app access](https://account.microsoft.com/privacy/app-access).

## Security

- `token.json` is mailbox access. Keep it on disk only.
- Do not paste it into chat, email, or GitHub.
- This is a personal/dev helper, not a multi-user product.
