"""Outlook OAuth2 IMAP reader plus Firstmail recovery inbox."""

import base64
import email
import imaplib
import os
import re
from email.header import decode_header
from email.utils import parseaddr, parsedate_to_datetime
from html import unescape

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
    :root { color-scheme: light; }
    * { box-sizing: border-box; }
    body { margin: 0; font: 15px/1.45 "Segoe UI", system-ui, sans-serif; background: #e7edf4; color: #1c2430; }
    main { max-width: 920px; margin: 0 auto; padding: 28px 16px 72px; }
    h1 { font-size: 1.35rem; font-weight: 680; margin: 0 0 6px; }
    p.note { color: #5d6b7c; margin: 0 0 16px; }
    form.panel, section.box { background: #fff; border: 1px solid #d5deea; border-radius: 14px; padding: 16px; }
    section.box { margin-top: 18px; }
    label { display: block; margin: 10px 0 6px; color: #3d4b5c; }
    textarea, input[type=number] {
      width: 100%; border: 1px solid #c9d4e2; border-radius: 8px; background: #f8fafc;
      color: inherit; padding: 10px 12px; font: inherit;
    }
    textarea { min-height: 120px; font-family: ui-monospace, Consolas, monospace; font-size: 13px; }
    .row { display: flex; gap: 14px; align-items: end; flex-wrap: wrap; margin-top: 4px; }
    .row .grow { flex: 1; min-width: 160px; }
    button.submit {
      background: #1f6feb; color: #fff; border: 0; border-radius: 8px;
      padding: 10px 16px; font: inherit; font-weight: 650; cursor: pointer;
    }
    .err { color: #a12626; }
    .ok { color: #17693a; margin: 0 0 8px; }
    h2 { font-size: 1.05rem; margin: 0 0 12px; }
    article.mail { border: 1px solid #e1e8f0; border-radius: 12px; padding: 14px 14px 10px; margin: 0 0 12px; background: #fbfcfe; }
    .who { display: flex; gap: 8px 12px; align-items: baseline; flex-wrap: wrap; color: #5d6b7c; font-size: 0.92rem; }
    .from { color: #1c2430; font-size: 0.98rem; }
    .addr { color: #6b7c90; }
    time { margin-left: auto; white-space: nowrap; }
    h3 { font-size: 1.05rem; font-weight: 680; margin: 6px 0 10px; }
    .codes { display: flex; flex-wrap: wrap; gap: 8px; margin: 0 0 10px; }
    button.chip {
      display: inline-flex; align-items: center; gap: 10px; border: 1px solid #b9d4ff;
      background: #eef5ff; color: #12315c; border-radius: 999px; padding: 7px 8px 7px 14px;
      font: inherit; cursor: pointer;
    }
    button.chip .chip-value { font-family: ui-monospace, Consolas, monospace; font-size: 1.15rem; font-weight: 720; letter-spacing: 0.08em; }
    button.chip .chip-action { background: #1f6feb; color: #fff; border-radius: 999px; padding: 4px 10px; font-size: 0.78rem; font-weight: 650; }
    button.chip.copied { border-color: #8ed0aa; background: #eefaf3; }
    button.chip.copied .chip-action { background: #17693a; }
    details summary { cursor: pointer; color: #1f6feb; font-weight: 650; padding: 4px 0 8px; }
    iframe.letter { width: 100%; min-height: 280px; border: 1px solid #e1e8f0; border-radius: 10px; background: #fff; }
    pre.body { white-space: pre-wrap; word-break: break-word; margin: 0; padding: 14px; background: #fff; border: 1px solid #e1e8f0; border-radius: 10px; }
    @media (max-width: 640px) {
      time { margin-left: 0; }
      button.chip { width: 100%; justify-content: space-between; }
    }
  </style>
</head>
<body>
<main>
  <h1>Outlook IMAP</h1>
  <p class="note">Строка аккаунта: email:password:recovery email:recovery password:refresh token:client id. Одна строка — один ящик. Данные нигде не сохраняются. Код до 8 символов копируется кнопкой, письмо раскрывать не нужно.</p>
  <form class="panel" method="post">
    <label for="accounts">Аккаунты</label>
    <textarea id="accounts" name="accounts" required>{{ accounts }}</textarea>
    <div class="row">
      <div class="grow">
        <label for="limit">Писем из каждого ящика</label>
        <input id="limit" name="limit" type="number" min="1" max="{{ max_messages }}" value="{{ limit }}">
      </div>
      <label><input type="checkbox" name="recovery" value="1" {% if recovery %}checked{% endif %}> ещё recovery через imap.firstmail.ltd:993</label>
      <button class="submit" type="submit">Открыть почту</button>
    </div>
  </form>
  {% macro letters(items) %}
    {% for msg in items %}
    <article class="mail">
      <div class="who">
        <strong class="from">{{ msg.from_name or msg.from_addr or msg.from }}</strong>
        {% if msg.from_name and msg.from_addr %}<span class="addr">{{ msg.from_addr }}</span>{% endif %}
        <time>{{ msg.date }}</time>
      </div>
      <h3>{{ msg.subject }}</h3>
      {% if msg.codes %}
      <div class="codes">
        {% for code in msg.codes %}
        <button type="button" class="chip" data-code="{{ code }}" onclick="copyCode(this)">
          <span class="chip-value">{{ code }}</span>
          <span class="chip-action">копировать</span>
        </button>
        {% endfor %}
      </div>
      {% endif %}
      <details>
        <summary>Открыть письмо</summary>
        {% if msg.html %}
        <iframe class="letter" sandbox="allow-same-origin allow-popups allow-popups-to-escape-sandbox" srcdoc="{{ msg.html }}" onload="fitFrame(this)"></iframe>
        {% else %}
        <pre class="body">{{ msg.text }}</pre>
        {% endif %}
      </details>
    </article>
    {% endfor %}
  {% endmacro %}
  {% for block in results %}
  <section class="box">
    <h2>{{ block.email }}</h2>
    {% if block.error %}<p class="err">{{ block.error }}</p>{% else %}<p class="ok">Outlook: {{ block.messages|length }} писем</p>{% endif %}
    {{ letters(block.messages) }}
    {% if block.recovery_error %}<p class="err">Recovery: {{ block.recovery_error }}</p>{% endif %}
    {% if block.recovery is not none %}
      <p class="ok">Recovery {{ block.recovery_email }}: {{ block.recovery|length }} писем</p>
      {{ letters(block.recovery) }}
    {% endif %}
  </section>
  {% endfor %}
</main>
<script>
function fitFrame(frame) {
  try {
    const doc = frame.contentDocument;
    if (!doc) return;
    const height = Math.max(doc.documentElement.scrollHeight, doc.body ? doc.body.scrollHeight : 0, 240);
    frame.style.height = (height + 24) + "px";
  } catch (err) {}
}
async function copyCode(btn) {
  const code = btn.dataset.code || "";
  try {
    await navigator.clipboard.writeText(code);
  } catch (err) {
    const area = document.createElement("textarea");
    area.value = code;
    document.body.appendChild(area);
    area.select();
    document.execCommand("copy");
    area.remove();
  }
  const action = btn.querySelector(".chip-action");
  btn.classList.add("copied");
  action.textContent = "скопировано";
  setTimeout(() => {
    btn.classList.remove("copied");
    action.textContent = "копировать";
  }, 1400);
}
</script>
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


LABELED_CODE = re.compile(
    r"(?i)(?:verification code|verify code|one[- ]time code|one[- ]time|код|code|pin|otp|пароль|passcode|password|verification|verify|confirm(?:ation)?)"
    r"[^0-9]{0,40}(\d{4,8})"
)
DIGIT_CODE = re.compile(r"(?<!\d)(\d{4,8})(?!\d)")
MIXED_CODE = re.compile(
    r"(?<![A-Za-z0-9])(?=[A-Za-z0-9]*[A-Za-z])(?=[A-Za-z0-9]*\d)[A-Za-z0-9]{4,8}(?![A-Za-z0-9])"
)
SPACED_CODE = re.compile(r"(?<!\d)(\d{2,4}(?:[ \u00a0]\d{2,4}){1,3})(?!\d)")
PREFIX_CODE = re.compile(r"(?<![A-Za-z0-9])([A-Za-z]-?\d{3,6})(?!\d)")


def part_text(part):
    payload = part.get_payload(decode=True)
    if payload is None:
        raw = part.get_payload()
        return raw if isinstance(raw, str) else ""
    return payload.decode(part.get_content_charset() or "utf-8", errors="replace")


def message_bodies(msg):
    plain = None
    html_body = None
    if msg.is_multipart():
        for part in msg.walk():
            disposition = str(part.get("Content-Disposition") or "")
            if "attachment" in disposition.lower():
                continue
            ctype = part.get_content_type()
            if ctype == "text/plain" and plain is None:
                plain = part_text(part)
            elif ctype == "text/html" and html_body is None:
                html_body = part_text(part)
    else:
        body = part_text(msg)
        if msg.get_content_type() == "text/html":
            html_body = body
        else:
            plain = body
    return plain or "", html_body or ""


def html_to_text(html):
    cleaned = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", " ", html)
    cleaned = re.sub(r"(?i)<br\s*/?>", "\n", cleaned)
    cleaned = re.sub(r"(?i)</(p|div|tr|h\d|li|table)>", "\n", cleaned)
    cleaned = re.sub(r"(?i)</td>", " ", cleaned)
    cleaned = re.sub(r"<[^>]+>", " ", cleaned)
    return unescape(cleaned)


def add_code(found, seen, raw):
    code = re.sub(r"[ \u00a0]+", "", raw or "")
    if not (4 <= len(code) <= 8) or not re.fullmatch(r"[A-Za-z0-9]+(?:-[A-Za-z0-9]+)?", code):
        return
    digits = re.sub(r"\D", "", code)
    if not digits:
        return
    if digits.isdigit() and len(digits) == 4 and 1990 <= int(digits) <= 2039 and not re.search(r"[A-Za-z]", code):
        return
    key = code.lower()
    bare = re.sub(r"[^A-Za-z0-9]", "", code).lower()
    if key in seen or bare in seen:
        return
    for index, existing in enumerate(found):
        old = re.sub(r"[^A-Za-z0-9]", "", existing).lower()
        if bare == old or bare in old or old in bare:
            if len(bare) > len(old) or (len(bare) == len(old) and re.search(r"[A-Za-z]", code)):
                found[index] = code
                seen.discard(existing.lower())
                seen.add(key)
                seen.add(bare)
            return
    seen.add(key)
    seen.add(bare)
    found.append(code)


def extract_codes(*chunks):
    text = "\n".join(chunk for chunk in chunks if chunk)
    found = []
    seen = set()
    for pattern in (LABELED_CODE, SPACED_CODE, PREFIX_CODE, DIGIT_CODE, MIXED_CODE):
        for match in pattern.finditer(text):
            value = match.group(1) if match.lastindex else match.group(0)
            if pattern is MIXED_CODE and value.islower():
                continue
            add_code(found, seen, value)
            if len(found) >= 6:
                return found
    return found


def sanitize_html(html):
    cleaned = re.sub(r"(?is)<(script|iframe|object|embed|form|meta|link)\b[^>]*>.*?</\1>", "", html)
    cleaned = re.sub(r"(?is)<(script|iframe|object|embed|form|meta|link)\b[^>]*/?>", "", cleaned)
    cleaned = re.sub(r"(?is)\son[a-z]+\s*=\s*(\"[^\"]*\"|'[^']*'|[^\s>]+)", "", cleaned)
    cleaned = re.sub(r"(?is)(href|src)\s*=\s*(['\"])\s*javascript:[^'\"]*\2", r"\1=\2#\2", cleaned)
    return cleaned


def inline_related_images(msg, html):
    for part in msg.walk():
        cid = part.get("Content-ID")
        payload = part.get_payload(decode=True) if cid else None
        if not cid or not payload or len(payload) > 400_000:
            continue
        token = cid.strip().strip("<>")
        mime = part.get_content_type() or "application/octet-stream"
        uri = f"data:{mime};base64,{base64.b64encode(payload).decode()}"
        html = html.replace(f"cid:{token}", uri)
    return html


def present_message(msg):
    plain, html_body = message_bodies(msg)
    visible = html_to_text(html_body) if html_body else plain
    raw_from = decode_mime(msg.get("From"))
    name, addr = parseaddr(raw_from)
    wrapped = ""
    if html_body:
        safe = inline_related_images(msg, sanitize_html(html_body))
        wrapped = (
            "<!doctype html><html><head><meta charset='utf-8'>"
            "<base target='_blank'>"
            "<style>body{margin:0;padding:16px;font:15px/1.5 Segoe UI,system-ui,sans-serif;"
            "color:#1c2430;background:#fff;overflow-wrap:anywhere}"
            "img{max-width:100%;height:auto} a{color:#1f6feb}</style></head><body>"
            f"{safe}</body></html>"
        )
    return {
        "from": raw_from,
        "from_name": name,
        "from_addr": addr,
        "subject": decode_mime(msg.get("Subject")) or "(без темы)",
        "date": format_date(msg.get("Date")),
        "text": plain or visible,
        "html": wrapped,
        "codes": extract_codes(decode_mime(msg.get("Subject")), visible),
    }


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
        messages.append(present_message(msg))
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
