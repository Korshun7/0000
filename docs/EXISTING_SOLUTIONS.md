# Готовые решения (что есть на рынке)

Короткий ответ: **продукта, который одновременно сидит в классическом Outlook, сортирует почту, пишет ответы по треду, работает на ПК и не умеет выйти в интернет — нет.** Есть облачные надстройки Outlook, локальные клиенты «вместо Outlook» и открытые заготовки.

## Сводка

| Решение | Outlook | Сортировка | Черновики по треду | ИИ без интернета | Готовность |
| --- | --- | --- | --- | --- | --- |
| Microsoft 365 Copilot | да | да | да | нет, облако Microsoft | продукт |
| MailMaestro / Sally | да, add-in | частично | да | нет, OpenAI/Claude/Gemini | продукт |
| OutlookAI | да, add-in | нет | да | нет, ChatGPT | OSS, облако |
| Envoy | OAuth Graph/Gmail, не COM | заявлено | заявлено | ИИ локально (Phi-4), почта через OAuth = сеть | waitlist / early |
| Canary Mail | свой клиент, не Outlook UI | локально | часто облако | триаж on-device, генерация обычно нет | продукт, другой клиент |
| Pontius | IMAP/CLI, не панель Outlook | да | да | да, если Ollama локально | продукт/утилита |
| ThunderAI + Thunderbird | нет, это Thunderbird | теги | да | да, Ollama / llama.cpp | самый зрелый офлайн-стек |
| Email-Triage | IMAP / провайдеры | да | да | да, Ollama | молодой self-host |
| microsoft/local-email-agent | Graph MCP | поиск | агент + HITL | нет: Graph + Azure embeddings | sample |
| Outlook Classic MCP + локальная LLM | COM к Outlook | через агента | через агента | только если агент локальный, не Claude | прототип |
| MailVault (этот репозиторий) | COM/MAPI | да | да | да, by design | проект, не продукт |

## 1. Облачные надстройки Outlook — «готово», но не подходит

Это то, что люди обычно имеют в виду под «ИИ для Outlook»:

- **Microsoft 365 Copilot** — лучшая интеграция в Outlook, треды, черновики. Письма уходят в контур Microsoft.
- **MailMaestro** — панель в Outlook, 3 варианта ответа, саммари треда. Модели OpenAI / Anthropic / Google.
- **Sally for Outlook** — add-in, умные ответы, история. Облако.
- **OutlookAI** — открытый add-in, бьёт в ChatGPT аккаунт пользователя. Локального инференса нет.
- **SaneBox, Spark Premium, Superhuman, Shortwave, Lavender** — сортировка и/или стиль; серверная обработка почты.

Имеет смысл только если требование air-gap снять. Для исходного ТЗ — дисквалификация.

## 2. «Локальный ИИ», но не тот контур

**Envoy** ([getenvoy.me](https://getenvoy.me/)) — ближе всех по маркетингу: Windows, Phi-4 Mini на ПК, подсказки ответов, категории, человек подтверждает отправку, Enterprise «on-prem». На практике почта подключается **OAuth к Outlook/Gmail** (то есть Graph/Google API), продукт на waitlist. Это не Extended MAPI и не «нет возможности выйти в интернет».

**Canary Mail** — готовый почтовый клиент. Триаж/приоритет считают on-device. Генерация черновиков у таких клиентов часто всё равно облачная. Outlook как UI не сохраняется.

**Pontius** — локальные модели (Ollama и др.), офлайн после установки, база контактов/правил. Это менеджер почты/CLI, не надстройка Outlook.

**microsoft/local-email-agent** — официальный sample: Phi локально + HITL. Забор почты через Microsoft 365 MCP (Graph), эмбеддинги в примере — Azure OpenAI. Имя «local» не означает air-gap.

## 3. Что реально можно поставить офлайн сегодня

### Если можно уйти с Outlook → Thunderbird

Самый готовый локальный стек:

- [ThunderAI](https://github.com/micz/ThunderAI) — теги, черновики, Ollama или любой OpenAI-compatible (`127.0.0.1`).
- ThunderClaude, AI Mail Support — то же семейство, llama.cpp / LM Studio.

Почта: IMAP/Exchange в Thunderbird. ИИ: llama.cpp на loopback. Не Outlook, но это единственная связка «скачал — работает — можно закрыть интернет у модели».

### Если Outlook обязателен → только сборка

Готового add-in с локальной Qwen нет. Рабочие куски:

1. **Доступ к письмам без Graph:** [Outlook Classic MCP](https://github.com/JacobBowie/Outlook-Classic-MCP), [claude-outlook-bridge](https://github.com/ChiefStarKid/claude-outlook-bridge), [outlook-mcp-server](https://github.com/marlonluo2018/outlook-mcp-server) — COM к установленному Outlook, Drafts-by-default.
2. **Инференс:** llama-server, не Ollama (у Ollama есть выход наружу за моделями).
3. **Триаж self-host:** [Email-Triage](https://github.com/Unlimited-Data-Works-LLC/Email-Triage) — классификация и стиль через Ollama, SQLite; не панель Outlook.

Склеить MCP + Qwen в Claude Desktop / Cursor — облачный агент снова видит письма. Нужен свой оркестратор, как в `docs/ARCHITECTURE.md`.

## 4. Практическая рекомендация

- Нужен **готовый продукт в Outlook завтра** и сеть допустима → Copilot или MailMaestro.
- Нужен **офлайн ИИ и почта не обязана быть Outlook** → Thunderbird + ThunderAI + llama.cpp, модели с флешки.
- Нужен **Outlook + air-gap + 200 писем + ответы по переписке** → готового нет; ближайшие кирпичи — COM/MCP + Email-Triage + собственная панель. Это и есть MailVault.

Перед пилотом любого «100% local» продукта проверять три вещи: ходит ли бинарник в интернет (Ollama, OAuth, телеметрия), пишет ли в Graph, есть ли в коде Send без подтверждения.
