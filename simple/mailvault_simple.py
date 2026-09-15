"""Minimal Outlook helper: selected mail (or JSON) -> local llama-server -> Drafts.

No send. Talks only to 127.0.0.1. Optional win32com on Windows.
"""

from __future__ import annotations

import argparse
import json
import sys
import urllib.error
import urllib.request
from typing import Any

LLAMA_URL = "http://127.0.0.1:8081/v1/chat/completions"
SYSTEM_PROMPT = """Ты помощник по почте. По треду верни JSON без markdown:
{"needs_reply": true, "category": "Reply", "draft": "текст письма"}
category: Reply | FYI | Triage
needs_reply=false для рассылок и уведомлений — draft тогда пустая строка.
Пиши черновик от имени получателя, на языке письма. Не выдумывай сроки и суммы.
Никогда не вызывай отправку. Только JSON."""


def llama_complete(thread_text: str, timeout: int = 180) -> dict[str, Any]:
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
            "Запустите его с --host 127.0.0.1"
        ) from exc
    text = body["choices"][0]["message"]["content"]
    return parse_model_json(text)


def parse_model_json(text: str) -> dict[str, Any]:
    start, end = text.find("{"), text.rfind("}")
    if start < 0 or end <= start:
        raise ValueError(f"модель вернула не JSON: {text[:400]}")
    data = json.loads(text[start : end + 1])
    for key in ("needs_reply", "category", "draft"):
        if key not in data:
            raise ValueError(f"нет поля {key}")
    if data["category"] not in {"Reply", "FYI", "Triage"}:
        data["category"] = "Triage"
    data["needs_reply"] = bool(data["needs_reply"])
    data["draft"] = str(data["draft"] or "")
    return data


def format_thread(messages: list[dict[str, str]]) -> str:
    parts = []
    for msg in messages:
        parts.append(
            f"From: {msg.get('from', '')}\n"
            f"Subject: {msg.get('subject', '')}\n"
            f"{msg.get('body', '')}\n---"
        )
    return "\n".join(parts)


def outlook_selected_thread() -> tuple[Any, list[dict[str, str]]]:
    import win32com.client  # type: ignore

    app = win32com.client.Dispatch("Outlook.Application")
    selection = app.ActiveExplorer().Selection
    if selection.Count < 1:
        raise SystemExit("В Outlook ничего не выделено")
    item = selection.Item(1)
    messages = [{
        "from": str(getattr(item, "SenderEmailAddress", "") or ""),
        "subject": str(item.Subject or ""),
        "body": str(item.Body or "")[:12000],
    }]
    try:
        conv = item.GetConversation()
        if conv is not None:
            table = conv.GetTable()
            rows = []
            while not table.EndOfTable:
                row = table.GetNextRow()
                rows.append(row)
            # Keep it simple: current item is enough if conversation API is picky.
            if rows:
                messages[0]["subject"] = f"{item.Subject} (тред {len(rows)} писем)"
    except Exception:
        pass
    return item, messages


def save_outlook_draft(source_item: Any, draft_text: str, category: str) -> None:
    import win32com.client  # type: ignore

    app = win32com.client.Dispatch("Outlook.Application")
    mail = app.CreateItem(0)
    reply_to = str(getattr(source_item, "SenderEmailAddress", "") or "")
    mail.To = reply_to
    mail.Subject = "Re: " + str(source_item.Subject or "")
    mail.Body = draft_text
    try:
        mail.Categories = category
        source_item.Categories = category
        source_item.Save()
    except Exception:
        pass
    mail.Save()


def load_thread_json(path: str) -> list[dict[str, str]]:
    with open(path, encoding="utf-8") as fh:
        data = json.load(fh)
    return data["messages"]


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Local Outlook draft helper")
    parser.add_argument("--thread-json", help="Файл треда вместо Outlook")
    parser.add_argument("--print-only", action="store_true", help="Не писать в Outlook")
    args = parser.parse_args(argv)

    source = None
    if args.thread_json:
        messages = load_thread_json(args.thread_json)
    else:
        if sys.platform != "win32":
            raise SystemExit("Без Outlook используйте --thread-json (пример: simple/sample_thread.json)")
        source, messages = outlook_selected_thread()

    result = llama_complete(format_thread(messages))
    print(json.dumps(result, ensure_ascii=False, indent=2))

    if args.print_only or source is None:
        return 0
    if result["needs_reply"] and result["draft"].strip():
        save_outlook_draft(source, result["draft"], result["category"])
        print("Черновик сохранён в Outlook. Отправку делает человек.")
    else:
        try:
            source.Categories = result["category"]
            source.Save()
        except Exception:
            pass
        print("Ответ не нужен, выставлена категория", result["category"])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
