# Модель данных

Файл БД: `%LOCALAPPDATA%\MailVault\mailvault.db` (SQLite 3, WAL, `synchronous=FULL` для надёжности на одном клиенте).

DDL: `schemas/mailvault.sql`.

## Сущности

**stores** — почтовое хранилище Outlook (OST/PST), чтобы EntryID не смешивались между профилями.

**messages** — нормализованное письмо.

- `entry_id`, `store_id`, `internet_message_id`
- `conversation_id`, `conversation_index` (hex)
- `sent_at`, `folder`, `is_sent`
- `from_addr`, `to_addrs`, `cc_addrs`
- `subject`, `body_text`, `body_hash`
- `has_attachments`, `attachment_names`
- `status`: ingested → embedded → classified → drafted | skipped | quarantined

**threads** — агрегаты по conversation_id: число писем, last_activity, needs_reply.

**classifications** — выход 8B + слой правил.

- `priority` 0–100
- `needs_reply` bool
- `intent`: question | request | fyi | meeting | invoice | legal | newsletter | ndr | other
- `rule_hit` — если сработало правило, модель можно не звать
- `confidence`
- `model_id`, `prompt_hash`

**embeddings** — вектор на чанк (письмо режется, если длинное).

**contacts** — локальная память: язык, форма обращения, «не обещать скидки» и т.п. Только то, что пользователь подтвердил или что стабильно видно в исходящих.

**drafts** — варианты ответа.

- `variant` 1..3
- `body_text`, `subject`
- `citations_json` — список entry_id
- `risks_json`
- `accepted` / `edited_body` / `copied_to_outlook_at`

**jobs** — очередь. `lease_until`, `attempts`, `last_error`.

**audit_events** — кто утвердил черновик, хеш текста. Без тела письма.

## Индексы

- уникальность `(store_id, entry_id)`
- уникальность `internet_message_id` где не NULL
- `conversation_id`, `sent_at`
- FTS5: subject + body_text + from_addr

## Почему не PostgreSQL

Один пользователь, один ПК, 200 писем/день + годы истории — миллионы строк максимум. SQLite проще вывезти в бэкап одним файлом и не слушает порт.
