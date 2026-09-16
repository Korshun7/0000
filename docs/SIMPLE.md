# Как сделать простой скрипт: сортировка Outlook и связанные письма

Нужны три вещи, которые уже есть в Outlook, без надстроек и Graph:

1. Прочитать письма через COM (`Outlook.Application`).
2. Собрать переписку (`ConversationID` / `GetConversation()`).
3. Создать **связанное** письмо методом `Reply()`, а не новым `CreateItem`.

Скрипты в `simple/` делают это целиком. Ниже — зачем каждый шаг.

## 1. Подключиться к классическому Outlook

New Outlook — веб-клиент, COM там нет. Нужен классический Outlook с открытым профилем.

PowerShell:

```powershell
$outlook = New-Object -ComObject Outlook.Application
$inbox = $outlook.GetNamespace("MAPI").GetDefaultFolder(6)  # olFolderInbox
```

Python: пакет `pywin32`, тот же объект `Outlook.Application`.

Если Outlook покажет «программа пытается получить доступ к почте» — Allow. Это не баг скрипта.

## 2. Сортировка

Не путать с кнопкой Sort в ленте. Скрипт:

- берёт непрочитанные (или выделенные);
- раскладывает по категориям `Reply` / `FYI` / `Triage`;
- внутри каждой связки сортирует письма по `ReceivedTime`.

Рассылку можно отсечь правилом Outlook (`noreply@` → FYI) ещё до скрипта. Скрипт повторяет ту же эвристику, чтобы пилот работал сразу.

## 3. Связанные письма

Связка — это тред, а не «похожая тема».

| Источник | Зачем |
| --- | --- |
| `MailItem.ConversationID` | Нормальный Exchange/Microsoft 365 тред |
| `GetConversation().GetTable()` | Подтянуть Sent Items и уже прочитанные письма той же переписки |
| Тема без `Re:`/`Fwd:` + общие участники | Склеить то, что Outlook разорвал (часто после Fwd) |

Короткие темы (`Hi`, `Test`) не склеиваются — слишком много ложных пар.

На все письма связки ставится одна категория `Связка:<тема>`. В представлении «По категориям» они лежат вместе, даже если папки разные.

## 4. Создать связанное письмо, а не новое

Так **нельзя**, если нужна цепочка:

```powershell
$mail = $outlook.CreateItem(0)
$mail.Subject = "Re: " + $original.Subject
$mail.To = $original.SenderEmailAddress
$mail.Save()
```

Outlook нарисует `Re:`, но это другое письмо: другой `ConversationIndex`, нет In-Reply-To.

Так **нужно**:

```powershell
$reply = $lastInbound.Reply()
$reply.Body = $digest + $reply.Body
$reply.Save()   # Черновики. Человек жмёт Отправить.
# $reply.Send()  — никогда
```

`Reply()` копирует служебные поля исходного письма. Черновик появляется в папке Черновики и в режиме переписок виден в том же треде.

## 5. Что класть в тело черновика

Без модели — список связанных писем (кто, когда, тема) и пустой блок «Черновик ответа». Человек дописывает текст, видя контекст.

С локальным `llama-server` (`--llm`) модель возвращает JSON `{needs_reply, category, draft}`. Текст вставляется в тот же `Reply()`. Облачные API в скрипте не вызываются.

## 6. Проверка без Windows

```bash
python simple/outlook_sort_link.py --dry-run --thread-json simple/sample_inbox.json
python -m pytest tests/test_outlook_sort_link.py
```

В `simple/sample_inbox.json` есть: обычный тред из трёх писем, разорванный Forward, рассылка `noreply`, две коротких темы `Hi`, которые склеивать нельзя.

## Файлы

| Файл | Роль |
| --- | --- |
| `simple/outlook_sort_link.ps1` | Скрипт без Python, только COM |
| `simple/outlook_sort_link.py` | Тот же алгоритм + JSON + опционально LLM |
| `simple/sample_inbox.json` | Учебный ящик |
| `tests/test_outlook_sort_link.py` | Группировка, категории, запрет Send |
