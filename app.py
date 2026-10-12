"""Outlook OAuth2 IMAP reader plus Firstmail recovery inbox."""

import base64
import email
import imaplib
import os
import re
import time
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
    "https://outlook.office.com/IMAP.AccessAsUser.All",
    "https://outlook.office365.com/IMAP.AccessAsUser.All",
)
OUTLOOK_HOSTS = ("outlook.office365.com", "imap-mail.outlook.com")
FATAL_TOKEN_ERRORS = {"invalid_grant", "interaction_required", "invalid_client"}
MAX_ACCOUNTS = 5
MAX_MESSAGES = 12
HTTP = requests.Session()

PAGE = """
<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta name="color-scheme" content="light">
  <title>Почта</title>
  <style>
    :root { color-scheme: light; }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      background: #d5d5d5;
      color: #1a1a1a;
      font: 15px/1.5 "Segoe UI", system-ui, sans-serif;
    }
    main { max-width: 760px; margin: 0 auto; padding: 40px 16px 88px; }
    h1 { margin: 0 0 8px; font-size: 1.7rem; font-weight: 600; letter-spacing: -0.04em; }
    p.note { margin: 0 0 18px; color: #5e5e5e; }
    form.panel, section.box {
      background: #f4f4f4;
      border-radius: 32px;
      padding: 18px;
      box-shadow: 0 16px 40px rgba(0, 0, 0, 0.06);
    }
    section.box { margin-top: 16px; }
    textarea {
      width: 100%;
      min-height: 140px;
      border: 0;
      border-radius: 24px;
      background: #e6e6e6;
      color: #1a1a1a;
      padding: 16px 18px;
      font: 13px/1.45 ui-monospace, Consolas, monospace;
      resize: vertical;
    }
    textarea:focus { outline: none; background: #efefef; }
    button.submit {
      margin-top: 14px;
      border: 0;
      border-radius: 999px;
      background: #2f80ed;
      color: #fff;
      padding: 12px 22px;
      font: inherit;
      font-weight: 650;
      cursor: pointer;
    }
    button.submit:disabled { background: #8a8a8a; cursor: default; }
    h2 { margin: 2px 4px 4px; font-size: 1rem; font-weight: 600; }
    p.empty { margin: 10px 6px 4px; color: #6d6d6d; }
    .tools { display: flex; gap: 10px; align-items: center; flex-wrap: wrap; margin-top: 14px; }
    input.find {
      flex: 1;
      min-width: 200px;
      margin-top: 0;
      border: 0;
      border-radius: 999px;
      background: #f4f4f4;
      color: #1a1a1a;
      padding: 12px 18px;
      font: inherit;
    }
    button.refresh { margin-top: 0; white-space: nowrap; }
    input.find:focus { outline: 2px solid #2f80ed; }
    article.mail { background: #fafafa; border-radius: 26px; padding: 16px 16px 12px; margin-top: 12px; }
    article.notion { box-shadow: inset 0 0 0 2px #2f80ed; }
    .badge {
      background: #2f80ed;
      color: #fff;
      border-radius: 999px;
      padding: 2px 10px;
      font-size: 0.75rem;
      font-weight: 700;
      letter-spacing: 0.04em;
    }
    .preview { margin: -4px 0 12px; color: #5e5e5e; }
    .notion-count { margin: 4px 4px 0; color: #1a56c4; font-weight: 650; }
    .who { display: flex; gap: 8px 12px; align-items: baseline; flex-wrap: wrap; color: #6a6a6a; font-size: 0.92rem; }
    .from { color: #161616; font-size: 1rem; }
    .source { letter-spacing: 0.08em; text-transform: uppercase; font-size: 0.72rem; color: #8a8a8a; }
    time { margin-left: auto; white-space: nowrap; }
    h3 { margin: 8px 0 12px; font-size: 1.08rem; font-weight: 600; letter-spacing: -0.02em; }
    .codes { display: flex; flex-wrap: wrap; gap: 8px; margin: 0 0 12px; }
    button.chip {
      display: inline-flex; align-items: center; gap: 10px;
      border: 0; background: #e7f1ff; color: #12315c;
      border-radius: 999px; padding: 7px 8px 7px 16px;
      font: inherit; cursor: pointer;
    }
    button.chip .chip-value { font-family: ui-monospace, Consolas, monospace; font-size: 1.15rem; font-weight: 700; letter-spacing: 0.08em; }
    button.chip .chip-action { background: #2f80ed; color: #fff; border-radius: 999px; padding: 5px 12px; font-size: 0.78rem; font-weight: 650; }
    button.chip.copied { background: #2f80ed; color: #fff; }
    button.chip.copied .chip-action { background: #fff; color: #1a56c4; }
    details summary {
      display: inline-flex;
      align-items: center;
      cursor: pointer;
      background: #2f80ed;
      color: #fff;
      font-weight: 650;
      border-radius: 999px;
      padding: 8px 16px;
      list-style: none;
    }
    details summary::-webkit-details-marker { display: none; }
    details .when-open { display: none; }
    details[open] .when-open { display: inline; }
    details[open] .when-closed { display: none; }
    details[open] summary { margin-bottom: 12px; }
    iframe.letter, pre.body {
      width: 100%; min-height: 240px; border: 0; border-radius: 22px; background: #fff;
      filter: grayscale(1);
    }
    pre.body { white-space: pre-wrap; word-break: break-word; margin: 0; padding: 16px; min-height: 0; }
    @media (max-width: 640px) {
      time { margin-left: 0; }
      button.chip { width: 100%; justify-content: space-between; }
    }
  </style>
</head>
<body>
<main>
  <h1>Почта</h1>
  <p class="note">Одна строка на ящик: email:password:recovery email:recovery password:refresh token:client id. Outlook и recovery открываются сами. Код копируется кнопкой.</p>
  <form class="panel" method="post" id="mail-form">
    <textarea id="accounts" name="accounts" placeholder="вставьте строки аккаунтов" required>{{ accounts }}</textarea>
    <button class="submit" type="submit">Открыть</button>
  </form>
  {% if results %}
  <div class="tools">
    <input id="find" class="find" type="search" placeholder="Найти письмо, например Notion">
    <button class="submit refresh" type="submit" form="mail-form">Обновить список</button>
  </div>
  {% endif %}
  {% macro letters(items) %}
    {% for msg in items %}
    <article class="mail{% if msg.notion %} notion{% endif %}" data-find="{{ (msg.from ~ ' ' ~ msg.subject ~ ' ' ~ msg.preview ~ ' ' ~ msg.source) | lower }}">
      <div class="who">
        {% if msg.notion %}<span class="badge">Notion</span>{% endif %}
        <strong class="from">{{ msg.from_name or msg.from_addr or msg.from }}</strong>
        {% if msg.from_name and msg.from_addr %}<span>{{ msg.from_addr }}</span>{% endif %}
        <span class="source">{{ msg.source }}</span>
        <time>{{ msg.date }}</time>
      </div>
      <h3>{{ msg.subject }}</h3>
      {% if msg.preview %}<p class="preview">{{ msg.preview }}</p>{% endif %}
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
        <summary>
          <span class="when-closed">Развернуть письмо</span>
          <span class="when-open">Свернуть письмо</span>
        </summary>
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
    {% set notion_count = block.messages | selectattr('notion') | list | length %}
    {% if notion_count %}<p class="notion-count">Notion: {{ notion_count }}</p>{% endif %}
    {% if block.messages %}
      {{ letters(block.messages) }}
    {% else %}
      <p class="empty">Писем нет</p>
    {% endif %}
  </section>
  {% endfor %}
</main>
<script>
function fitFrame(frame) {
  try {
    const doc = frame.contentDocument;
    if (!doc) return;
    const height = Math.max(doc.documentElement.scrollHeight, doc.body ? doc.body.scrollHeight : 0, 220);
    frame.style.height = (height + 28) + "px";
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
const find = document.getElementById("find");
if (find) {
  find.addEventListener("input", function () {
    const query = find.value.trim().toLowerCase();
    document.querySelectorAll("article.mail").forEach(function (card) {
      card.hidden = Boolean(query) && !(card.dataset.find || "").includes(query);
    });
  });
}
document.getElementById("mail-form").addEventListener("submit", function (event) {
  document.querySelectorAll("button.submit").forEach(function (button) {
    button.disabled = true;
  });
  const clicked = event.submitter;
  if (clicked) clicked.textContent = clicked.classList.contains("refresh") ? "Обновляем" : "Открываем";
});
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


CODE_LABEL = re.compile(
    r"(?i)(?<![A-Za-zА-Яа-яЁё])"
    r"(?<!zip )(?<!postal )(?<!error )(?<!status )(?<!country )(?<!area )(?<!bar )"
    r"(?:verification code|security code|login code|confirmation code|authentication code|"
    r"sign[- ]in code|one[- ]time password|one[- ]time code|one[- ]time pin|"
    r"passcode|otp|pin|код подтверждения|одноразовый код|код безопасности|код доступа|"
    r"проверочный код|пароль|password|код(?:а|у|ом|е)?|code)"
    r"(?![A-Za-zА-Яа-яЁё])"
)
CODE_NEAR = re.compile(
    r"(?<!\d)(\d{4,8})(?!\d)"
    r"|(?<![A-Za-z0-9])([A-Za-z]-\d{3,6})(?!\d)"
    r"|(?<![A-Za-z0-9])((?=[A-Za-z0-9]*[A-Za-z])(?=[A-Za-z0-9]*\d)[A-Za-z0-9]{5,8})(?![A-Za-z0-9])"
)
CODE_SPACED = re.compile(r"(?<!\d)(\d{2,4}(?:[ \u00a0]\d{2,4}){1,2})")


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
    cleaned = re.sub(r"(?i)</(?:p|div|tr|h\d|li|table|td|th|center|blockquote)>", "\n", cleaned)
    cleaned = re.sub(r"(?i)<(?:p|div|tr|h\d|li|td|th|center|blockquote)\b[^>]*>", "\n", cleaned)
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


def code_lines(text):
    values = []
    for raw_line in text.splitlines():
        line = raw_line.strip().strip("*_>#-–—[]()«»\"'")
        if not line:
            continue
        spaced = CODE_SPACED.fullmatch(line)
        if spaced:
            values.append(spaced.group(1))
            continue
        if " " in line:
            continue
        if re.fullmatch(r"[A-Za-z]-\d{3,6}|\d{4,8}", line):
            values.append(line)
        elif re.fullmatch(r"(?=[A-Za-z]*\d)(?=\d*[A-Za-z])[A-Za-z0-9]{5,8}", line):
            values.append(line)
    return values


def candidates_near(window):
    runs = [match.span() for match in re.finditer(r"(?<!\d)(?:\d[ \u00a0]?){9,}(?!\d)", window)]

    def buried(span):
        return any(span[0] >= start and span[1] <= end for start, end in runs)

    six = []
    other = []
    for match in CODE_NEAR.finditer(window):
        if buried(match.span()):
            continue
        value = next(group for group in match.groups() if group)
        if re.fullmatch(r"\d{6}", value):
            six.append(value)
        else:
            other.append(value)
    spaced = []
    for match in CODE_SPACED.finditer(window):
        if buried(match.span()) or re.match(r"[ \u00a0]\d", window[match.end() :]):
            continue
        spaced.append(match.group(1))
    return six + code_lines(window) + spaced + other


def take_code(found, seen, values):
    before = len(found)
    for value in values:
        add_code(found, seen, value)
        if len(found) > before:
            return True
    return False


def extract_codes(*chunks):
    text = "\n".join(chunk for chunk in chunks if chunk)
    text = re.sub(r"https?://\S+", " ", text)
    text = text.replace("\u00a0", " ")
    found = []
    seen = set()
    labeled = False
    for match in CODE_LABEL.finditer(text):
        labeled = True
        window = text[match.end() : match.end() + 400]
        before = text[max(0, match.start() - 48) : match.start()]
        if take_code(found, seen, candidates_near(window) or candidates_near(before)):
            if len(found) >= 3:
                break
    if not found:
        take_code(found, seen, code_lines(text))
    elif labeled:
        for value in code_lines(text):
            if re.fullmatch(r"\d{6}", re.sub(r"\D", "", value)):
                add_code(found, seen, value)
    return found[:4]


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
            "img{max-width:100%;height:auto} a{color:#1a1a1a}</style></head><body>"
            f"{safe}</body></html>"
        )
    subject = decode_mime(msg.get("Subject")) or "(без темы)"
    preview = " ".join((plain or visible).split())[:180]
    blob = " ".join([raw_from, name or "", addr or "", subject, preview]).lower()
    return {
        "from": raw_from,
        "from_name": name,
        "from_addr": addr,
        "subject": subject,
        "date": format_date(msg.get("Date")),
        "text": plain or visible,
        "html": wrapped,
        "preview": preview,
        "notion": "notion" in blob,
        "uid": (msg.get("Message-ID") or f"{addr}|{subject}|{msg.get('Date')}").strip(),
        "codes": extract_codes(subject, plain, visible),
    }


def format_date(value):
    if not value:
        return ""
    try:
        return parsedate_to_datetime(value).strftime("%Y-%m-%d %H:%M")
    except (TypeError, ValueError, IndexError):
        return value


def close_imap(imap):
    if imap is None:
        return
    try:
        imap.logout()
    except Exception:
        try:
            imap.shutdown()
        except Exception:
            return


def mailbox_name(item):
    text = item.decode("utf-8", "replace") if isinstance(item, bytes) else str(item)
    quoted = re.findall(r'"([^"]*)"', text)
    if quoted:
        return quoted[-1]
    parts = text.split()
    return parts[-1] if parts else ""


def interesting_mailboxes(imap):
    names = []
    try:
        status, data = imap.list()
        if status == "OK":
            names = [mailbox_name(item) for item in data or [] if item]
    except imaplib.IMAP4.error:
        names = []
    chosen = []
    seen = set()
    for name in names:
        key = name.lower()
        if key == "inbox" or "junk" in key or "spam" in key or key in {"bulk mail", "bulk"}:
            if key not in seen:
                seen.add(key)
                chosen.append(name)
    if "inbox" not in seen:
        chosen.insert(0, "INBOX")
    return chosen


def search_ids(imap, *criteria):
    try:
        status, data = imap.search(None, *criteria)
    except (imaplib.IMAP4.error, OSError):
        return []
    if status != "OK" or not data or not data[0]:
        return []
    return data[0].split()


def notion_header_ids(imap, ids):
    found = []
    for num in ids:
        try:
            status, data = imap.fetch(num, "(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT)])")
        except (imaplib.IMAP4.error, OSError):
            continue
        if status != "OK" or not data:
            continue
        raw = b"".join(part[1] for part in data if isinstance(part, tuple) and part[1])
        if b"notion" in raw.lower():
            found.append(num)
    return found


def fetch_ids(imap, ids):
    messages = []
    for num in ids:
        try:
            status, msg_data = imap.fetch(num, "(RFC822)")
        except imaplib.IMAP4.error:
            continue
        if status != "OK" or not msg_data or not isinstance(msg_data[0], tuple):
            continue
        messages.append(present_message(email.message_from_bytes(msg_data[0][1])))
    return messages


def read_folder(imap, mailbox, recent_limit):
    target = f'"{mailbox}"' if any(char in mailbox for char in ' "') else mailbox
    try:
        status, _ = imap.select(target, readonly=True)
    except imaplib.IMAP4.error:
        return []
    if status != "OK":
        return []
    all_ids = search_ids(imap, "ALL")
    recent = all_ids[-recent_limit:] if recent_limit else []
    notion = search_ids(imap, "FROM", "notion") + search_ids(imap, "SUBJECT", "notion")
    if not notion:
        notion = notion_header_ids(imap, all_ids[-80:])
    ordered = []
    seen = set()
    for num in list(notion[-30:]) + list(recent):
        if num not in seen:
            seen.add(num)
            ordered.append(num)
    return fetch_ids(imap, ordered)


def read_inbox(imap, limit):
    return read_folder(imap, "INBOX", limit)


def access_token_for(account):
    bodies = [{"scope": scope} for scope in IMAP_SCOPES]
    bodies.append({})
    for extra in bodies:
        data = {
            "client_id": account["client_id"],
            "refresh_token": account["refresh_token"],
            "grant_type": "refresh_token",
        }
        data.update(extra)
        for attempt in range(2):
            try:
                response = HTTP.post(TOKEN_URL, data=data, timeout=25)
            except requests.RequestException:
                time.sleep(0.4)
                continue
            try:
                payload = response.json()
            except ValueError:
                payload = {}
            token = payload.get("access_token")
            if token:
                return token
            if payload.get("error") in FATAL_TOKEN_ERRORS:
                return ""
            break
    return ""


def outlook_messages(account, limit):
    token = access_token_for(account)
    if not token:
        return []
    auth = f"user={account['email']}\x01auth=Bearer {token}\x01\x01".encode()
    for host in OUTLOOK_HOSTS:
        imap = None
        try:
            imap = imaplib.IMAP4_SSL(host, 993, timeout=25)
            imap.authenticate("XOAUTH2", lambda _: auth)
            found = []
            mailboxes = interesting_mailboxes(imap)
            mailboxes.sort(key=lambda name: 0 if name.lower() == "inbox" else 1)
            for mailbox in mailboxes[:4]:
                try:
                    recent = limit if mailbox.lower() == "inbox" else 0
                    found.extend(read_folder(imap, mailbox, recent))
                except Exception:
                    continue
            if found:
                return dedupe_messages(found)
        except Exception:
            continue
        finally:
            close_imap(imap)
    return []


def recovery_messages(account, limit):
    imap = None
    try:
        imap = imaplib.IMAP4_SSL(FIRSTMAIL_IMAP, 993, timeout=25)
        imap.login(account["recovery_email"], account["recovery_password"])
        return read_inbox(imap, limit)
    except Exception:
        return []
    finally:
        close_imap(imap)


def stamp(messages, source):
    for message in messages:
        message["source"] = source
    return messages


def dedupe_messages(messages):
    unique = []
    seen = set()
    for message in messages:
        key = message.get("uid") or f"{message.get('from')}|{message.get('date')}|{message.get('subject')}"
        if key in seen:
            continue
        seen.add(key)
        unique.append(message)
    return unique


def by_date(messages):
    return sorted(messages, key=lambda item: item.get("date") or "", reverse=True)


def check_account(account):
    outlook = []
    recovery = []
    try:
        outlook = stamp(outlook_messages(account, 25), "Outlook")
    except Exception:
        outlook = []
    try:
        recovery = stamp(recovery_messages(account, 8), "Recovery")
    except Exception:
        recovery = []
    notion = by_date([item for item in outlook + recovery if item.get("notion")])
    rest = by_date([item for item in outlook if not item.get("notion")])[:20]
    rest.extend(by_date([item for item in recovery if not item.get("notion")])[:8])
    return {"email": account["email"], "messages": dedupe_messages(notion + rest)}


@app.get("/health")
def health():
    return {"ok": True}


@app.route("/", methods=["GET", "POST"])
def index():
    accounts_text = ""
    results = []
    try:
        if request.method == "POST":
            accounts_text = request.form.get("accounts") or ""
            lines = [line for line in accounts_text.splitlines() if line.strip()][:MAX_ACCOUNTS]
            for line in lines:
                try:
                    account = parse_account(line)
                except ValueError:
                    continue
                try:
                    results.append(check_account(account))
                except Exception:
                    results.append({"email": account["email"], "messages": []})
    except Exception:
        results = []
    return render_template_string(PAGE, accounts=accounts_text, results=results)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8000")))
