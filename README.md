# Простой скрипт сортировки Outlook и связанных писем

Берёт письма классического Outlook, собирает **тред** (связанные письма одной переписки) и кладёт **черновик-ответ в ту же цепочку**. Ничего не отправляет.

## Что поставить

- Windows, **классический** Outlook (не «Новый Outlook»).
- Либо PowerShell 5.1 — без лишних пакетов.
- Либо Python 3.10+ и `pip install pywin32`.

Outlook должен быть открыт.

## Самый короткий путь — PowerShell

Выделите письмо во Входящих:

```powershell
powershell -File simple/outlook_sort_link.ps1
```

Или разложить непрочитанные (до 200 штук):

```powershell
powershell -File simple/outlook_sort_link.ps1 -Inbox
```

Посмотреть, что получится, без записи в Outlook:

```powershell
powershell -File simple/outlook_sort_link.ps1 -Inbox -DryRun
```

## Тот же алгоритм на Python

Выделенное письмо:

```bat
python simple/outlook_sort_link.py
```

Непрочитанные Входящие:

```bat
python simple/outlook_sort_link.py --inbox
```

Проверка без Outlook (этот репозиторий, любой ОС):

```bat
python simple/outlook_sort_link.py --dry-run --thread-json simple/sample_inbox.json
```

## Как устроена связка

1. Письма группируются по `ConversationID` Outlook.
2. Если Outlook разорвал переписку, скрипт склеивает её по нормализованной теме (`Re:` / `Fwd:` / `Отв:` снимаются) и пересечению отправителей/получателей. Короткие темы вроде `Hi` не склеиваются.
3. На всю связку ставится категория `Reply` / `FYI` / `Triage` и метка `Связка:<тема>`.
4. Если нужен ответ, вызывается **`Reply()` у последнего входящего письма**, черновик сохраняется. Так сохраняются `ConversationIndex` и заголовки In-Reply-To — письмо остаётся в треде.
5. Новое пустое письмо через `CreateItem` не создаётся: визуально будет `Re:`, но цепочка в Outlook порвётся. `Send()` в скрипте нет.

Правила без модели:

| Категория | Когда |
| --- | --- |
| FYI | `noreply`, рассылка, unsubscribe / «отписаться» |
| Reply | в тексте есть вопрос или «прошу» / «нужно» / please |
| Triage | остальное; категория есть, черновика нет, пока не указать `-DraftTriage` / `--draft-triage` |

## Если уже крутится локальная модель

Тот же Python-скрипт умеет спросить `llama-server` на `127.0.0.1:8081`:

```bat
python simple/outlook_sort_link.py --inbox --llm
```

Без модели скрипт всё равно соберёт связку и подставит в черновик список писем треда — текст ответа дописывает человек.

## Если Outlook спросит разрешение

Один раз Allow. Для постоянного запуска: Файл → параметры Outlook → центр управления безопасностью → программный доступ.

Планировщик заданий Windows: раз в час `outlook_sort_link.ps1 -Inbox`, Outlook при этом должен быть запущен.

Подробный разбор: [docs/SIMPLE.md](docs/SIMPLE.md).
