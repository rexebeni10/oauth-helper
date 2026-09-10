"""Read recent Outlook messages via Microsoft Graph and print OTP-style codes.

Usage:
  python outlook_otp.py
  python outlook_otp.py --token token.json --minutes 10

Put token.json next to this script (download it from native/index.html).
Never commit token.json.
"""
from __future__ import annotations

import argparse
import html as htmlmod
import json
import re
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Optional

import requests

HERE = Path(__file__).resolve().parent
DEFAULT_TOKEN = HERE / "token.json"

CLIENT_ID = "9d540e09-3b89-4c2f-be10-2fed49864b53"
TOKEN_URL = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token"
SCOPE = "Mail.Read offline_access"
GRAPH = "https://graph.microsoft.com/v1.0"

CODE_RE = re.compile(r"\b(\d{4,8})\b")
SKIP_CODES = {"0000", "00000", "000000", "1234", "123456", "111111"}
AADSTS_RE = re.compile(r"AADSTS\d+")


def load_blob(path: Path) -> dict[str, Any]:
    raw = path.read_text(encoding="utf-8").strip()
    if not raw:
        raise ValueError(f"{path} is empty")
    if raw.startswith("{"):
        obj = json.loads(raw)
        if not isinstance(obj, dict):
            raise ValueError("token file must be a JSON object")
        return obj
    return {"refresh_token": raw, "origin": ""}


def redeem(refresh_token: str, *, origin: Optional[str]) -> tuple[Optional[str], str]:
    data = {
        "client_id": CLIENT_ID,
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "scope": SCOPE,
    }
    headers: dict[str, str] = {}
    if origin:
        headers["Origin"] = origin
    try:
        r = requests.post(TOKEN_URL, data=data, headers=headers, timeout=20)
    except requests.RequestException as e:
        return None, f"net:{type(e).__name__}"
    if r.status_code == 200:
        tok = (r.json() or {}).get("access_token")
        return (tok, "") if tok else (None, "no_access_token")
    desc = ""
    try:
        body = r.json()
        if isinstance(body, dict):
            desc = str(body.get("error_description") or body.get("error") or "")
    except ValueError:
        desc = r.text or ""
    m = AADSTS_RE.search(desc)
    return None, (m.group(0) if m else f"http:{r.status_code}")


def strip_html(content: str) -> str:
    text = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", content or "")
    text = re.sub(r"(?is)<br\s*/?>", "\n", text)
    text = re.sub(r"(?is)</p>", "\n", text)
    text = re.sub(r"(?is)<[^>]+>", " ", text)
    text = htmlmod.unescape(text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def extract_codes(subject: str, body: str) -> list[str]:
    found: list[str] = []
    seen: set[str] = set()
    for src in (subject or "", body or ""):
        for m in CODE_RE.finditer(src):
            code = m.group(1)
            if code in SKIP_CODES or code in seen:
                continue
            seen.add(code)
            found.append(code)
    found.sort(key=lambda c: (0 if len(c) == 6 else 1, -len(c)))
    return found


def fetch_recent(access_token: str, *, minutes: float, top: int) -> list[dict[str, Any]]:
    since = datetime.now(timezone.utc) - timedelta(minutes=minutes)
    since_iso = since.strftime("%Y-%m-%dT%H:%M:%SZ")
    headers = {"Authorization": f"Bearer {access_token}"}
    out: list[dict[str, Any]] = []
    for folder in ("inbox", "junkemail"):
        url = f"{GRAPH}/me/mailFolders/{folder}/messages"
        params = {
            "$top": str(top),
            "$orderby": "receivedDateTime desc",
            "$select": "subject,from,receivedDateTime,body,bodyPreview",
            "$filter": f"receivedDateTime ge {since_iso}",
        }
        try:
            r = requests.get(url, params=params, headers=headers, timeout=30)
        except requests.RequestException as e:
            print(f"  [warn] {folder}: {e}")
            continue
        if r.status_code != 200:
            print(f"  [warn] {folder} HTTP {r.status_code}")
            continue
        for msg in r.json().get("value") or []:
            sender = ((msg.get("from") or {}).get("emailAddress") or {}).get("address") or "?"
            subject = (msg.get("subject") or "").strip()
            body_obj = msg.get("body") or {}
            raw = body_obj.get("content") or ""
            if (body_obj.get("contentType") or "").lower() == "html":
                text = strip_html(raw)
            else:
                text = (raw or msg.get("bodyPreview") or "").strip()
            out.append({
                "folder": folder,
                "from": sender,
                "subject": subject,
                "received": msg.get("receivedDateTime"),
                "codes": extract_codes(subject, text),
                "preview": text[:400],
            })
    out.sort(key=lambda m: m.get("received") or "", reverse=True)
    return out


def print_messages(msgs: list[dict[str, Any]], *, minutes: float) -> None:
    if not msgs:
        print(f"\n  No messages in the last {minutes:g} minute(s).")
        return
    print(f"\n  {len(msgs)} message(s) in the last {minutes:g} minute(s):\n")
    all_codes: list[str] = []
    for i, m in enumerate(msgs, 1):
        codes = m.get("codes") or []
        all_codes.extend(codes)
        code_s = ", ".join(codes) if codes else "(none found)"
        print(f"  [{i}] {m.get('received')}")
        print(f"      folder : {m.get('folder')}")
        print(f"      from   : {m.get('from')}")
        print(f"      subject: {m.get('subject')}")
        print(f"      codes  : {code_s}")
        prev = (m.get("preview") or "").replace("\n", " / ")
        if prev:
            print(f"      preview: {prev[:220]}")
        print()
    uniq: list[str] = []
    for c in all_codes:
        if c not in uniq:
            uniq.append(c)
    six = [c for c in uniq if len(c) == 6]
    if six:
        print(f"  >>> Likely verification code(s): {', '.join(six)}")
    elif uniq:
        print(f"  >>> Code candidate(s): {', '.join(uniq)}")


def main() -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description="Print recent Outlook OTP codes via Graph.")
    ap.add_argument("--token", type=Path, default=DEFAULT_TOKEN,
                    help=f"token.json from the helper page (default: {DEFAULT_TOKEN})")
    ap.add_argument("--minutes", type=float, default=10.0,
                    help="look-back window in minutes (default: 10)")
    ap.add_argument("--top", type=int, default=25, help="max messages per folder")
    args = ap.parse_args()

    if not args.token.is_file():
        print(f"[fatal] missing {args.token}")
        print("Download token.json from native/index.html and save it here.")
        return 2

    blob = load_blob(args.token)
    rt = (blob.get("refresh_token") or "").strip()
    if not rt:
        print("[fatal] no refresh_token in token file")
        return 2
    origin = blob.get("origin")
    if origin is None:
        origin = None
    elif origin == "":
        origin = None
    else:
        origin = str(origin)

    email = blob.get("email") or "?"
    print(f"Redeeming Graph token for {email} ...")
    access, err = redeem(rt, origin=origin)
    if not access:
        print(f"[fail] token redeem failed: {err}")
        return 2
    print(f"OK — scanning Inbox + Junk (last {args.minutes:g} min) ...")
    msgs = fetch_recent(access, minutes=args.minutes, top=args.top)
    print_messages(msgs, minutes=args.minutes)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
