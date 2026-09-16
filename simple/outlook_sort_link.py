"""Сортировка почты Outlook и связанные письма.

Группирует письма в треды (ConversationID, иначе тема + участники),
ставит категории и сохраняет черновик через Reply(), чтобы письмо
осталось в той же переписке. Send() в скрипте нет.

Windows, классический Outlook:

  python simple/outlook_sort_link.py
  python simple/outlook_sort_link.py --inbox
  python simple/outlook_sort_link.py --dry-run --thread-json simple/sample_inbox.json
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any

OL_MAIL = 43
OL_FOLDER_INBOX = 6
CATEGORY_REPLY = "Reply"
CATEGORY_FYI = "FYI"
CATEGORY_TRIAGE = "Triage"
LINK_PREFIX = "Связка"
MIN_SUBJECT_MERGE = 8

SUBJECT_PREFIXES = re.compile(
    r"^((re|fw|fwd|aw|sv|vs|отв|на|пересл)\s*[:：]\s*)+",
    re.IGNORECASE,
)
EMAIL_RE = re.compile(r"[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}", re.IGNORECASE)
FYI_RE = re.compile(
    r"no-?reply|noreply|mailer-daemon|notifications?@|unsubscribe|"
    r"отписаться|newsletter|рассылк",
    re.IGNORECASE,
)
REPLY_RE = re.compile(
    r"\b(прошу|нужно|необходим|подтверд|уточн|пожалуйста|please|asap|срочно)\b",
    re.IGNORECASE,
)
LLAMA_URL = "http://127.0.0.1:8081/v1/chat/completions"
SYSTEM_PROMPT = (
    "Ты помощник по почте. По треду верни JSON без markdown:\n"
    '{"needs_reply": true, "category": "Reply", "draft": "текст письма"}\n'
    "category: Reply | FYI | Triage\n"
    "needs_reply=false для рассылок — draft тогда пустая строка.\n"
    "Пиши черновик от имени получателя, на языке письма. "
    "Не выдумывай сроки и суммы. Никогда не вызывай отправку. Только JSON."
)


@dataclass
class Message:
    id: str
    conversation_id: str
    sender: str
    to: str
    subject: str
    body: str
    received: datetime
    unread: bool = True
    cc: str = ""
    raw: Any = field(default=None, repr=False)

    def participants(self) -> set[str]:
        found: set[str] = set()
        for blob in (self.sender, self.to, self.cc):
            found.update(addr.casefold() for addr in EMAIL_RE.findall(blob or ""))
        return found


@dataclass
class Thread:
    key: str
    messages: list[Message]
    category: str = CATEGORY_TRIAGE
    needs_reply: bool = False
    draft: str = ""

    @property
    def latest(self) -> Message:
        return self.messages[-1]

    @property
    def subject(self) -> str:
        return self.latest.subject if self.messages else ""

    @property
    def normalized_subject(self) -> str:
        return normalize_subject(self.subject)

    def participants(self) -> set[str]:
        people: set[str] = set()
        for msg in self.messages:
            people |= msg.participants()
        return people


def normalize_subject(subject: str) -> str:
    text = re.sub(r"\s+", " ", subject or "").strip()
    while True:
        stripped = SUBJECT_PREFIXES.sub("", text).strip()
        if stripped == text:
            break
        text = stripped
    return text.casefold()


def parse_received(value: str | datetime) -> datetime:
    if isinstance(value, datetime):
        return value
    return datetime.fromisoformat(value.replace("Z", "+00:00")).replace(tzinfo=None)


def load_messages(path: str) -> tuple[str, list[Message]]:
    with open(path, encoding="utf-8") as fh:
        data = json.load(fh)
    account = str(data.get("account", ""))
    messages = [message_from_dict(row) for row in data["messages"]]
    return account, messages


def message_from_dict(row: dict[str, Any]) -> Message:
    return Message(
        id=str(row.get("id") or ""),
        conversation_id=str(row.get("conversation_id") or ""),
        sender=str(row.get("from") or ""),
        to=str(row.get("to") or ""),
        cc=str(row.get("cc") or ""),
        subject=str(row.get("subject") or ""),
        body=str(row.get("body") or ""),
        received=parse_received(str(row["received"])),
        unread=bool(row.get("unread", True)),
    )


def group_threads(messages: list[Message]) -> list[Thread]:
    buckets: dict[str, list[Message]] = {}
    for msg in messages:
        cid = (msg.conversation_id or "").strip()
        key = f"cid:{cid}" if cid else f"subj:{normalize_subject(msg.subject)}"
        buckets.setdefault(key, []).append(msg)

    threads = [
        Thread(key=key, messages=sorted(items, key=lambda m: m.received))
        for key, items in buckets.items()
    ]
    merged = _merge_split_threads(threads)
    merged.sort(key=lambda t: t.latest.received, reverse=True)
    return merged


def _merge_split_threads(threads: list[Thread]) -> list[Thread]:
    result: list[Thread] = []
    used = [False] * len(threads)
    for i, left in enumerate(threads):
        if used[i]:
            continue
        bundle = list(left.messages)
        keys = [left.key]
        used[i] = True
        changed = True
        while changed:
            changed = False
            people = set()
            for msg in bundle:
                people |= msg.participants()
            subject = normalize_subject(bundle[-1].subject)
            if len(subject) < MIN_SUBJECT_MERGE:
                break
            for j, right in enumerate(threads):
                if used[j]:
                    continue
                if normalize_subject(right.subject) != subject:
                    continue
                if people & right.participants():
                    bundle.extend(right.messages)
                    bundle.sort(key=lambda m: m.received)
                    keys.append(right.key)
                    used[j] = True
                    changed = True
        result.append(
            Thread(key="+".join(keys), messages=sorted(bundle, key=lambda m: m.received))
        )
    return result


def classify_thread(thread: Thread, account: str = "") -> Thread:
    last_inbound = _last_inbound(thread, account)
    blob = "\n".join(
        [last_inbound.sender, last_inbound.subject, last_inbound.body[:4000]]
    )
    if FYI_RE.search(blob):
        thread.category = CATEGORY_FYI
        thread.needs_reply = False
        thread.draft = ""
        return thread
    if "?" in last_inbound.body or REPLY_RE.search(last_inbound.body):
        thread.category = CATEGORY_REPLY
        thread.needs_reply = True
        thread.draft = build_linked_draft(thread, account)
    else:
        thread.category = CATEGORY_TRIAGE
        thread.needs_reply = False
        thread.draft = ""
    return thread


def _last_inbound(thread: Thread, account: str) -> Message:
    me = (account or "").casefold()
    inbound = [m for m in thread.messages if me and me not in m.sender.casefold()]
    return inbound[-1] if inbound else thread.latest


def build_linked_draft(thread: Thread, account: str = "") -> str:
    lines = [
        f"Категория: {thread.category}",
        f"Связанные письма этого треда ({len(thread.messages)}):",
        "",
    ]
    me = (account or "").casefold()
    for index, msg in enumerate(thread.messages, start=1):
        mine = " (я)" if me and me in msg.sender.casefold() else ""
        stamp = msg.received.strftime("%Y-%m-%d %H:%M")
        lines.append(f"{index}. {stamp} {msg.sender}{mine}: {msg.subject}")
    lines += ["", "--- Черновик ответа ---", "", ""]
    return "\n".join(lines)


def format_thread_for_model(thread: Thread) -> str:
    parts = []
    for msg in thread.messages:
        parts.append(
            f"From: {msg.sender}\nTo: {msg.to}\nSubject: {msg.subject}\n"
            f"{msg.body[:8000]}\n---"
        )
    return "\n".join(parts)


def parse_model_json(text: str) -> dict[str, Any]:
    start, end = text.find("{"), text.rfind("}")
    if start < 0 or end <= start:
        raise ValueError(f"модель вернула не JSON: {text[:400]}")
    data = json.loads(text[start : end + 1])
    for key in ("needs_reply", "category", "draft"):
        if key not in data:
            raise ValueError(f"нет поля {key}")
    if data["category"] not in {CATEGORY_REPLY, CATEGORY_FYI, CATEGORY_TRIAGE}:
        data["category"] = CATEGORY_TRIAGE
    data["needs_reply"] = bool(data["needs_reply"])
    data["draft"] = str(data["draft"] or "")
    return data


def llama_complete(thread_text: str, timeout: int = 180) -> dict[str, Any]:
    import urllib.error
    import urllib.request

    payload = {
        "model": "qwen3-8b",
        "temperature": 0.2,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": thread_text},
        ],
    }
    req = urllib.request.Request(
        LLAMA_URL,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            body = json.loads(resp.read().decode("utf-8"))
    except urllib.error.URLError as exc:
        raise SystemExit(
            f"llama-server недоступен на {LLAMA_URL} ({exc}). "
            "Запустите его с --host 127.0.0.1 или уберите флаг --llm."
        ) from exc
    text = body["choices"][0]["message"]["content"]
    return parse_model_json(text)


def apply_llm(thread: Thread) -> Thread:
    result = llama_complete(format_thread_for_model(thread))
    thread.category = result["category"]
    thread.needs_reply = result["needs_reply"]
    model_draft = result["draft"].strip()
    digest = build_linked_draft(thread)
    thread.draft = f"{digest}{model_draft}\n" if model_draft else digest
    if thread.category == CATEGORY_FYI:
        thread.needs_reply = False
        thread.draft = ""
    return thread


def thread_category_label(thread: Thread) -> str:
    short = (thread.normalized_subject or "без темы")[:40]
    return f"{thread.category}, {LINK_PREFIX}:{short}"


def thread_to_dict(thread: Thread) -> dict[str, Any]:
    return {
        "key": thread.key,
        "category": thread.category,
        "needs_reply": thread.needs_reply,
        "subject": thread.subject,
        "messages": [
            {
                "id": m.id,
                "from": m.sender,
                "subject": m.subject,
                "received": m.received.isoformat(timespec="seconds"),
            }
            for m in thread.messages
        ],
        "draft_preview": (thread.draft[:240] + "…") if len(thread.draft) > 240 else thread.draft,
    }


def _outlook_app():
    if sys.platform != "win32":
        raise SystemExit(
            "Outlook COM работает только на Windows. "
            "Для проверки логики: --thread-json simple/sample_inbox.json --dry-run"
        )
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise SystemExit("Нужен пакет pywin32: pip install pywin32") from exc
    return win32com.client.Dispatch("Outlook.Application")


def _item_smtp(item: Any) -> str:
    addr = str(getattr(item, "SenderEmailAddress", "") or "")
    if "@" in addr:
        return addr
    try:
        sender = item.Sender
        exchange = sender.GetExchangeUser()
        smtp = str(exchange.PrimarySmtpAddress or "") if exchange else ""
        if smtp:
            return smtp
    except Exception:
        pass
    return addr


def _current_user_smtp(namespace: Any) -> str:
    try:
        accessor = namespace.CurrentUser.PropertyAccessor
        smtp = accessor.GetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
        )
        if smtp:
            return str(smtp)
    except Exception:
        pass
    try:
        return str(namespace.CurrentUser.Address or "")
    except Exception:
        return ""


def message_from_outlook(item: Any) -> Message:
    received = getattr(item, "ReceivedTime", None)
    if received is None:
        received = datetime.now()
    if hasattr(received, "strftime") and not isinstance(received, datetime):
        received = datetime(
            received.year,
            received.month,
            received.day,
            received.hour,
            received.minute,
            received.second,
        )
    to_line = str(getattr(item, "To", "") or "")
    cc_line = str(getattr(item, "CC", "") or "")
    return Message(
        id=str(getattr(item, "EntryID", "") or ""),
        conversation_id=str(getattr(item, "ConversationID", "") or ""),
        sender=_item_smtp(item),
        to=to_line,
        cc=cc_line,
        subject=str(getattr(item, "Subject", "") or ""),
        body=str(getattr(item, "Body", "") or "")[:20000],
        received=received if isinstance(received, datetime) else datetime.now(),
        unread=bool(getattr(item, "UnRead", False)),
        raw=item,
    )


def expand_conversation(namespace: Any, item: Any) -> list[Any]:
    found: list[Any] = []
    seen: set[str] = set()
    try:
        conv = item.GetConversation()
        if conv is None:
            return [item]
        table = conv.GetTable()
        while not table.EndOfTable:
            row = table.GetNextRow()
            entry_id = None
            try:
                entry_id = row("EntryID")
            except Exception:
                try:
                    entry_id = row.GetValue("EntryID")
                except Exception:
                    entry_id = None
            if not entry_id or entry_id in seen:
                continue
            seen.add(str(entry_id))
            try:
                found.append(namespace.GetItemFromID(entry_id))
            except Exception:
                continue
    except Exception:
        return [item]
    return found or [item]


def collect_outlook_messages(
    selected_only: bool,
    inbox: bool,
    unread_only: bool,
    limit: int,
) -> tuple[str, list[Message]]:
    app = _outlook_app()
    namespace = app.GetNamespace("MAPI")
    account = _current_user_smtp(namespace)
    raw_items: list[Any] = []

    if selected_only or not inbox:
        explorer = app.ActiveExplorer()
        selection = explorer.Selection
        if selection.Count < 1:
            raise SystemExit("В Outlook ничего не выделено. Выделите письмо или укажите --inbox.")
        for index in range(1, selection.Count + 1):
            item = selection.Item(index)
            if int(getattr(item, "Class", 0)) != OL_MAIL:
                continue
            raw_items.extend(expand_conversation(namespace, item))
    else:
        folder = namespace.GetDefaultFolder(OL_FOLDER_INBOX)
        items = folder.Items
        items.Sort("[ReceivedTime]", True)
        for item in items:
            if int(getattr(item, "Class", 0)) != OL_MAIL:
                continue
            if unread_only and not bool(item.UnRead):
                continue
            raw_items.append(item)
            if len(raw_items) >= limit:
                break
        expanded: list[Any] = []
        for item in raw_items:
            expanded.extend(expand_conversation(namespace, item))
        raw_items = expanded

    unique: dict[str, Message] = {}
    for item in raw_items:
        if int(getattr(item, "Class", 0)) != OL_MAIL:
            continue
        msg = message_from_outlook(item)
        unique[msg.id or f"anon-{len(unique)}"] = msg
    return account, list(unique.values())


def save_linked_draft(thread: Thread, account: str = "") -> None:
    """Создаёт связанное письмо через Reply(), не CreateItem и не Send."""
    source = _last_inbound(thread, account).raw
    if source is None:
        for msg in reversed(thread.messages):
            if msg.raw is not None:
                source = msg.raw
                break
    if source is None:
        raise SystemExit("Нет COM-объекта письма, черновик создать нельзя")
    reply = source.Reply()
    original = str(reply.Body or "")
    reply.Body = thread.draft + "\n" + original
    try:
        label = thread_category_label(thread)
        reply.Categories = label
        for msg in thread.messages:
            if msg.raw is None:
                continue
            msg.raw.Categories = label
            msg.raw.Save()
        reply.Save()
    except Exception as exc:
        raise SystemExit(f"Outlook не дал сохранить черновик: {exc}") from exc


def apply_categories_only(thread: Thread) -> None:
    label = thread_category_label(thread)
    for msg in thread.messages:
        if msg.raw is None:
            continue
        try:
            msg.raw.Categories = label
            msg.raw.Save()
        except Exception:
            continue


def process_messages(
    messages: list[Message],
    account: str,
    use_llm: bool,
    draft_triage: bool,
) -> list[Thread]:
    threads = group_threads(messages)
    processed: list[Thread] = []
    for thread in threads:
        classify_thread(thread, account)
        if use_llm:
            apply_llm(thread)
        if thread.category == CATEGORY_TRIAGE and draft_triage:
            thread.needs_reply = True
            if not thread.draft:
                thread.draft = build_linked_draft(thread, account)
        processed.append(thread)
    return processed


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Сортирует Outlook и создаёт связанные черновики в том же треде"
    )
    parser.add_argument("--inbox", action="store_true", help="Взять письма из Входящих")
    parser.add_argument("--all", action="store_true", help="С --inbox брать и прочитанные")
    parser.add_argument("--limit", type=int, default=200, help="Максимум писем из Inbox")
    parser.add_argument("--thread-json", help="JSON вместо Outlook (проверка без Windows)")
    parser.add_argument("--dry-run", action="store_true", help="Ничего не писать в Outlook")
    parser.add_argument("--llm", action="store_true", help="Классификация и текст через llama-server")
    parser.add_argument("--draft-triage", action="store_true", help="Черновик и для категории Triage")
    args = parser.parse_args(argv)

    if args.thread_json:
        account, messages = load_messages(args.thread_json)
    else:
        account, messages = collect_outlook_messages(
            selected_only=not args.inbox,
            inbox=args.inbox,
            unread_only=not args.all,
            limit=args.limit,
        )

    if not messages:
        print("Писем нет")
        return 0

    threads = process_messages(messages, account, args.llm, args.draft_triage)
    print(json.dumps([thread_to_dict(t) for t in threads], ensure_ascii=False, indent=2))

    if args.dry_run or args.thread_json:
        return 0

    drafts = 0
    for thread in threads:
        apply_categories_only(thread)
        if thread.needs_reply and thread.draft.strip():
            save_linked_draft(thread, account)
            drafts += 1
    print(f"Тредов: {len(threads)}. Связанных черновиков: {drafts}. Отправку делает человек.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
