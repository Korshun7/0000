from __future__ import annotations

import json
from pathlib import Path

import outlook_sort_link as osl

SIMPLE = Path(__file__).resolve().parents[1] / "simple"
SAMPLE = SIMPLE / "sample_inbox.json"
PY_SRC = SIMPLE / "outlook_sort_link.py"
PS_SRC = SIMPLE / "outlook_sort_link.ps1"


def test_normalize_subject_strips_prefixes() -> None:
    assert osl.normalize_subject("Re: Re: Счёт по договору 17") == "счёт по договору 17"
    assert osl.normalize_subject("Fwd: График поставки насосов") == "график поставки насосов"
    assert osl.normalize_subject("Отв: На: Тема") == "тема"


def test_group_keeps_native_conversation() -> None:
    _, messages = osl.load_messages(str(SAMPLE))
    invoice = [m for m in messages if m.conversation_id == "conv-invoice"]
    threads = osl.group_threads(invoice)
    assert len(threads) == 1
    assert [m.id for m in threads[0].messages] == ["a1", "a2", "a3"]


def test_group_merges_split_outlook_threads() -> None:
    _, messages = osl.load_messages(str(SAMPLE))
    threads = osl.group_threads(messages)
    pumps = [t for t in threads if "насосов" in t.normalized_subject]
    assert len(pumps) == 1
    assert {m.id for m in pumps[0].messages} == {"b1", "b2"}


def test_group_does_not_merge_short_generic_subjects() -> None:
    _, messages = osl.load_messages(str(SAMPLE))
    threads = osl.group_threads(messages)
    hellos = [t for t in threads if t.normalized_subject == "hi"]
    assert len(hellos) == 2


def test_classify_noreply_as_fyi() -> None:
    _, messages = osl.load_messages(str(SAMPLE))
    threads = osl.process_messages(messages, "me@company.com", use_llm=False, draft_triage=False)
    news = next(t for t in threads if "noreply" in t.latest.sender)
    assert news.category == osl.CATEGORY_FYI
    assert news.needs_reply is False
    assert news.draft == ""


def test_classify_question_as_reply_and_build_linked_digest() -> None:
    _, messages = osl.load_messages(str(SAMPLE))
    threads = osl.process_messages(messages, "me@company.com", use_llm=False, draft_triage=False)
    invoice = next(t for t in threads if "договору" in t.normalized_subject)
    assert invoice.category == osl.CATEGORY_REPLY
    assert invoice.needs_reply is True
    assert "Связанные письма этого треда (3)" in invoice.draft
    assert "alice@vendor.com" in invoice.draft
    assert "(я)" in invoice.draft


def test_threads_sorted_by_latest_first() -> None:
    _, messages = osl.load_messages(str(SAMPLE))
    threads = osl.process_messages(messages, "me@company.com", use_llm=False, draft_triage=False)
    received = [t.latest.received for t in threads]
    assert received == sorted(received, reverse=True)


def test_cli_dry_run_json(capsys) -> None:
    code = osl.main(["--thread-json", str(SAMPLE), "--dry-run"])
    assert code == 0
    payload = json.loads(capsys.readouterr().out)
    assert isinstance(payload, list)
    assert {row["category"] for row in payload} >= {"Reply", "FYI", "Triage"}
    pumps = next(row for row in payload if "насосов" in row["subject"].lower())
    assert len(pumps["messages"]) == 2
    assert pumps["needs_reply"] is True


def test_scripts_never_send_mail() -> None:
    py_text = PY_SRC.read_text(encoding="utf-8")
    ps_text = PS_SRC.read_text(encoding="utf-8")
    assert ".Send(" not in py_text
    assert "CreateItem(" not in py_text
    assert ".Reply()" in py_text
    assert ".Send(" not in ps_text
    assert "CreateItem(" not in ps_text
    assert ".Reply()" in ps_text
