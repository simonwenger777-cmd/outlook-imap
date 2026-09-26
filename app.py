"""Outlook OAuth2 IMAP reader plus Firstmail recovery inbox."""

import email
import imaplib
import os
from email.header import decode_header
from email.utils import parsedate_to_datetime

import requests
from flask import Flask, render_template_string, request

app = Flask(__name__)

TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
OUTLOOK_IMAP = "outlook.office365.com"
FIRSTMAIL_IMAP = "imap.firstmail.ltd"
IMAP_SCOPES = (
    "https://outlook.office.com/IMAP.AccessAsUser.All offline_access",
    "https://outlook.office365.com/IMAP.AccessAsUser.All offline_access",
)
MAX_ACCOUNTS = 5
MAX_MESSAGES = 15

PAGE = """
<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Outlook IMAP</title>
  <style>
    :root { color-scheme: dark; }
    body { margin: 0; font: 16px/1.45 system-ui, sans-serif; background: #10141a; color: #e7edf5; }
    main { max-width: 880px; margin: 0 auto; padding: 28px 16px 64px; }
    h1 { font-size: 1.4rem; margin: 0 0 8px; }
    p.note { color: #9aabbd; margin: 0 0 18px; }
    label { display: block; margin: 12px 0 6px; color: #c5d2e0; }
    textarea, input[type=number] {
      width: 100%; box-sizing: border-box; border: 1px solid #2c3a4a; border-radius: 8px;
      background: #0c1016; color: inherit; padding: 10px 12px; font: inherit;
    }
    textarea { min-height: 140px; font-family: ui-monospace, Consolas, monospace; font-size: 13px; }
    .row { display: flex; gap: 16px; align-items: end; flex-wrap: wrap; }
    .row .grow { flex: 1; min-width: 180px; }
    button {
      background: #3d8bfd; color: #041018; border: 0; border-radius: 8px;
      padding: 10px 16px; font: inherit; font-weight: 650; cursor: pointer;
    }
    .card { margin-top: 22px; border: 1px solid #2c3a4a; border-radius: 12px; padding: 14px 16px; background: #161d27; }
    .err { color: #ffb4b4; }
    .ok { color: #9ddeb5; }
    article { border-top: 1px solid #2c3a4a; padding: 10px 0; }
    article:first-of-type { border-top: 0; }
    .meta { color: #9aabbd; font-size: 0.92rem; }
    pre { white-space: pre-wrap; word-break: break-word; margin: 8px 0 0; color: #d5deea; }
  </style>
</head>
<body>
<main>
  <h1>Outlook IMAP</h1>
  <p class="note">Строка аккаунта: email:password:recovery email:recovery password:refresh token:client id. Одна строка — один ящик. Данные нигде не сохраняются.</p>
  <form method="post">
    <label for="accounts">Аккаунты</label>
    <textarea id="accounts" name="accounts" required>{{ accounts }}</textarea>
    <div class="row">
      <div class="grow">
        <label for="limit">Писем из каждого ящика</label>
        <input id="limit" name="limit" type="number" min="1" max="{{ max_messages }}" value="{{ limit }}">
      </div>
      <label><input type="checkbox" name="recovery" value="1" {% if recovery %}checked{% endif %}> ещё recovery через imap.firstmail.ltd:993</label>
      <button type="submit">Открыть почту</button>
    </div>
  </form>
  {% for block in results %}
  <section class="card">
    <h2>{{ block.email }}</h2>
    {% if block.error %}<p class="err">{{ block.error }}</p>{% else %}<p class="ok">Outlook: {{ block.messages|length }} писем</p>{% endif %}
    {% for msg in block.messages %}
    <article>
      <div><strong>{{ msg.subject }}</strong></div>
      <div class="meta">{{ msg.date }} · {{ msg.from }}</div>
      <pre>{{ msg.preview }}</pre>
    </article>
    {% endfor %}
    {% if block.recovery_error %}<p class="err">Recovery: {{ block.recovery_error }}</p>{% endif %}
    {% if block.recovery is not none %}
      <p class="ok">Recovery {{ block.recovery_email }}: {{ block.recovery|length }} писем</p>
      {% for msg in block.recovery %}
      <article>
        <div><strong>{{ msg.subject }}</strong></div>
        <div class="meta">{{ msg.date }} · {{ msg.from }}</div>
        <pre>{{ msg.preview }}</pre>
      </article>
      {% endfor %}
    {% endif %}
  </section>
  {% endfor %}
</main>
</body>
</html>
"""


def parse_account(line):
    parts = [part.strip() for part in line.strip().split(":")]
    if len(parts) < 6 or not parts[0] or not parts[-1] or not parts[-2]:
        raise ValueError(
            "Нужен формат email:password:recovery email:recovery password:refresh token:client id"
        )
    email_addr = parts[0]
    client_id = parts[-1]
    refresh_token = parts[-2]
    middle = parts[1:-2]
    recovery_at = next((i for i, part in enumerate(middle) if "@" in part), None)
    if recovery_at is None:
        raise ValueError("В строке нет recovery email")
    password = ":".join(middle[:recovery_at])
    recovery_email = middle[recovery_at]
    recovery_password = ":".join(middle[recovery_at + 1 :])
    if not password or not recovery_password:
        raise ValueError("Пустой пароль ящика или recovery")
    return {
        "email": email_addr,
        "password": password,
        "recovery_email": recovery_email,
        "recovery_password": recovery_password,
        "refresh_token": refresh_token,
        "client_id": client_id,
    }


def decode_mime(value):
    if not value:
        return ""
    chunks = []
    for part, charset in decode_header(value):
        if isinstance(part, bytes):
            chunks.append(part.decode(charset or "utf-8", errors="replace"))
        else:
            chunks.append(part)
    return "".join(chunks)


def message_preview(msg, limit=700):
    if msg.is_multipart():
        plain = None
        html = None
        for part in msg.walk():
            disposition = str(part.get("Content-Disposition") or "")
            if "attachment" in disposition.lower():
                continue
            ctype = part.get_content_type()
            payload = part.get_payload(decode=True)
            if payload is None:
                continue
            text = payload.decode(part.get_content_charset() or "utf-8", errors="replace")
            if ctype == "text/plain" and plain is None:
                plain = text
            elif ctype == "text/html" and html is None:
                html = text
        text = plain if plain is not None else (html or "")
    else:
        payload = msg.get_payload(decode=True) or b""
        text = payload.decode(msg.get_content_charset() or "utf-8", errors="replace")
    text = " ".join(text.split())
    if len(text) > limit:
        return text[:limit] + "…"
    return text


def format_date(value):
    if not value:
        return ""
    try:
        return parsedate_to_datetime(value).strftime("%Y-%m-%d %H:%M")
    except (TypeError, ValueError, IndexError):
        return value


def read_inbox(imap, limit):
    status, _ = imap.select("INBOX", readonly=True)
    if status != "OK":
        raise RuntimeError("Не удалось открыть INBOX")
    status, data = imap.search(None, "ALL")
    if status != "OK":
        raise RuntimeError("Поиск писем не удался")
    ids = data[0].split() if data and data[0] else []
    messages = []
    for num in ids[-limit:]:
        status, msg_data = imap.fetch(num, "(RFC822)")
        if status != "OK" or not msg_data or not msg_data[0]:
            continue
        raw = msg_data[0][1]
        msg = email.message_from_bytes(raw)
        messages.append(
            {
                "from": decode_mime(msg.get("From")),
                "subject": decode_mime(msg.get("Subject")) or "(без темы)",
                "date": format_date(msg.get("Date")),
                "preview": message_preview(msg),
            }
        )
    messages.reverse()
    return messages


def access_token_for(account):
    last_detail = "пустой ответ"
    for scope in IMAP_SCOPES:
        response = requests.post(
            TOKEN_URL,
            data={
                "client_id": account["client_id"],
                "refresh_token": account["refresh_token"],
                "grant_type": "refresh_token",
                "scope": scope,
            },
            timeout=30,
        )
        try:
            payload = response.json()
        except ValueError:
            payload = {}
        access_token = payload.get("access_token")
        if access_token:
            return access_token
        last_detail = payload.get("error_description") or payload.get("error") or response.text[:300]
    raise RuntimeError(f"Нет access_token: {last_detail}")


def outlook_messages(account, limit):
    access_token = access_token_for(account)

    auth_string = f"user={account['email']}\x01auth=Bearer {access_token}\x01\x01"
    imap = imaplib.IMAP4_SSL(OUTLOOK_IMAP, 993, timeout=30)
    try:
        imap.authenticate("XOAUTH2", lambda _: auth_string.encode())
        return read_inbox(imap, limit)
    finally:
        try:
            imap.logout()
        except imaplib.IMAP4.error:
            imap.shutdown()


def recovery_messages(account, limit):
    imap = imaplib.IMAP4_SSL(FIRSTMAIL_IMAP, 993, timeout=30)
    try:
        imap.login(account["recovery_email"], account["recovery_password"])
        return read_inbox(imap, limit)
    finally:
        try:
            imap.logout()
        except imaplib.IMAP4.error:
            imap.shutdown()


def check_account(account, limit, with_recovery):
    block = {
        "email": account["email"],
        "error": None,
        "messages": [],
        "recovery_email": account["recovery_email"],
        "recovery": None,
        "recovery_error": None,
    }
    try:
        block["messages"] = outlook_messages(account, limit)
    except Exception as exc:
        block["error"] = str(exc)
    if with_recovery:
        try:
            block["recovery"] = recovery_messages(account, limit)
        except Exception as exc:
            block["recovery_error"] = str(exc)
    return block


@app.get("/health")
def health():
    return {"ok": True}


@app.route("/", methods=["GET", "POST"])
def index():
    accounts_text = ""
    limit = 10
    with_recovery = False
    results = []
    if request.method == "POST":
        accounts_text = request.form.get("accounts") or ""
        with_recovery = request.form.get("recovery") == "1"
        try:
            limit = int(request.form.get("limit") or 10)
        except ValueError:
            limit = 10
        limit = max(1, min(MAX_MESSAGES, limit))
        lines = [line for line in accounts_text.splitlines() if line.strip()]
        if len(lines) > MAX_ACCOUNTS:
            results.append(
                {
                    "email": "Лимит",
                    "error": f"За один запрос не больше {MAX_ACCOUNTS} строк",
                    "messages": [],
                    "recovery": None,
                    "recovery_error": None,
                    "recovery_email": "",
                }
            )
            lines = lines[:MAX_ACCOUNTS]
        for line in lines:
            try:
                account = parse_account(line)
            except ValueError as exc:
                results.append(
                    {
                        "email": line.split(":", 1)[0][:80],
                        "error": str(exc),
                        "messages": [],
                        "recovery": None,
                        "recovery_error": None,
                        "recovery_email": "",
                    }
                )
                continue
            results.append(check_account(account, limit, with_recovery))
    return render_template_string(
        PAGE,
        accounts=accounts_text,
        limit=limit,
        recovery=with_recovery,
        results=results,
        max_messages=MAX_MESSAGES,
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8000")))
