# PLAYBOOK: Claude Setup & Tools — настройка, возможности, инструменты

> Создан: 31 мая 2026 · Чат: "Shikardos Team Sonnet 4.6 (Quick)"
> Модель: Claude Sonnet 4.6

---

## ЗАДАЧА 1: Подключение новой папки проекта к Cowork

**СТАТУС:** Завершено

**ЧТО РАБОТАЕТ:**
1. Открыть Cowork на десктопе
2. Claude автоматически предлагает или используй `request_cowork_directory`
3. Указать путь `D:\CLAUDE Cowork\shikardosteam.github.io`
4. Папка монтируется — Claude получает доступ к файлам
5. Bash-путь внутри сессии: `/sessions/sweet-hopeful-cannon/mnt/shikardosteam.github.io/`
6. CLAUDE.md находится в `memory-base/CLAUDE.md` (НЕ в корне)

**ЧТО НЕ РАБОТАЕТ:**
- `Read` напрямую на `D:\CLAUDE Cowork\shikardosteam.github.io\CLAUDE.md` — файл не в корне
- Попытка читать папку через Read (только через Glob или bash `ls`)

**ОШИБКИ КОТОРЫЕ ДОПУСТИЛИ:**
- Сначала искал CLAUDE.md в корне — его там нет, он в `memory-base/CLAUDE.md`

**ФАЙЛЫ:**
- `D:\CLAUDE Cowork\shikardosteam.github.io\memory-base\CLAUDE.md` — точка входа базы знаний
- `D:\CLAUDE Cowork\shikardosteam.github.io\TASKS.md` — список задач

**ПРОМПТ ДЛЯ ПОВТОРЕНИЯ:**
```
Прочитай memory-base/CLAUDE.md и подтверди контекст проекта.
```

---

## ЗАДАЧА 2: Анализ очереди задач проекта

**СТАТУС:** Завершено (анализ сделан, задачи не выполнены)

**ЧТО РАБОТАЕТ:**
- Читать `TASKS.md` + `memory-base/06-current-status.md` параллельно
- Разделить задачи на 3 категории: "Claude делает сам", "Нужен Валерий", "Нужен CTO/команда"

**ЗАДАЧИ ВЫСОКОГО ПРИОРИТЕТА (из TASKS.md на 31 мая 2026):**

Claude может делать прямо сейчас:
1. FAQ → обновить категорию «Цифры» под финальную формулу КЭ [HIGH]
2. FAQ → добавить категорию «Как распределяется трафик» (Вариант Б, жеребьёвка) [MED]
3. Раздел Мейкер → блок «Критерии входа» (50+ видео, 5K просмотров, 5K подписчиков) [MED]
4. Раздел Мейкер → блок «Формула КЭ» финальная версия [MED]
5. Орг структура → «5 команд × 10 AI-офферщиков, 250 ниш/год» [MED]
6. Дашборд Раздел 5 «Офферы» → переделать под 250 ниш/год [HIGH]
7. Дашборд Раздел 19 «КЭ» → перевёрстка под финальную формулу [HIGH]
8. 9 Job Descriptions для найма топ-команды [HIGH]
9. Наполнение `tz-cto.md` и `tz-bot-dashboard.md` из .docx [HIGH]

Нужны действия Валерия:
- Найм EA #1 и EA #2 (критический приоритет)
- Видео-питч 15-30 мин для @shikardosteam
- Закрыть 25 мест Early Investor по $2 500
- LinkedIn профиль — регулярный постинг
- Postiz — подключить и настроить автопостинг

**ФАЙЛЫ:**
- `D:\CLAUDE Cowork\shikardosteam.github.io\TASKS.md`
- `D:\CLAUDE Cowork\shikardosteam.github.io\memory-base\06-current-status.md`

**ПРОМПТ ДЛЯ ПОВТОРЕНИЯ:**
```
Прочитай TASKS.md и 06-current-status.md. Выдай топ-5 задач которые я могу сделать 
прямо сейчас без найма команды, с HIGH приоритетом.
```

---

## ЗАДАЧА 3: Разбор версий Claude и их отличий

**СТАТУС:** Завершено (информация зафиксирована)

**ЧТО РАБОТАЕТ — итоговая шпаргалка:**

| Версия | Доступ к файлам ПК | Коннекторы | Dispatch | Голос |
|---|---|---|---|---|
| Веб (claude.ai) | ❌ | ✅ | ❌ | ❌ |
| Мобильное app | ❌ | ✅ | ✅ отправитель | ✅ |
| Cowork desktop | ✅ | ✅ | ✅ исполнитель | ❌ |

**Коннекторы у Валерия (на 31 мая 2026):**
- Gmail (6 инструментов)
- Google Calendar (8 инструментов)
- Google Drive (8 инструментов)
- iPhone Calendar — Read & Write (Ask each time)
- iPhone Reminders — Read & Write
- Location — Read only
- Notion (инструменты подключены в десктопе)
- Miro (инструменты подключены)

**DISPATCH — как работает:**
1. Телефон → меню ☰ → Dispatch → "Pair with your desktop"
2. На десктопе подтвердить сопряжение
3. После: задача с телефона → Cowork выполняет → ответ в тот же чат
4. Dispatch messages — уже включено ✅ (из скрина настроек)

**ПРОМПТ ДЛЯ ПОВТОРЕНИЯ:**
```
Объясни разницу между мобильной версией Claude, веб-версией и Cowork desktop.
Какие коннекторы у меня подключены?
```

---

## ЗАДАЧА 4: Обзор и цены инструментов (Apollo, Postiz, Brightdata)

**СТАТУС:** Завершено (исследование, не внедрение)

**APOLLO.IO:**
- База: 300M+ контактов, 30M+ компаний
- Фильтры: должность, страна, технологии, размер компании, стартап-опыт
- Планы 2026: Free ($0, 100 email-кредитов/мес) · Basic $49/мес · Pro $79/мес · Org $119/мес
- Реальная стоимость: $600–1 000/мес при активном использовании (кредиты)
- Для Валерия: **Free или Basic** для найма 9 топ-ролей — нужны только email
- Что умеет через Claude: найти кандидатов по фильтрам → список с email/LinkedIn → email-цепочки
- Ограничение: не ищет в Telegram/арбитражных сообществах

**POSTIZ:**
- 28+ платформ: Instagram, TikTok, YouTube, Facebook, Telegram, LinkedIn, Threads, VK, X, Reddit, Pinterest, Discord, Slack, Medium, WordPress и др.
- Планы 2026: Standard $29/мес (5 каналов) · Team $39/мес (10 каналов) · Pro $49/мес (30 каналов) · Ultimate $99/мес
- Open source — можно self-host бесплатно (хостинг ~$5-10/мес)
- Для Валерия: **Standard $29/мес** — закрывает задачу дублирования контента Founder без найма помощника
- Платформы нужные для проекта: YouTube ✅ TikTok ✅ Instagram ✅ Facebook ✅ Telegram ✅ LinkedIn ✅

**BRIGHTDATA:**
- Веб-скрапинг и прокси, enterprise-уровень
- Цена: Growth $499/мес, реальный средний чек $24 000/год
- Для Валерия: **НЕ нужен сейчас** — в стеке уже есть Apify
- Когда вернуться: когда CTO скажет что Apify не справляется с объёмами Года 2-3

**РЕКОМЕНДАЦИИ К ДЕЙСТВИЮ:**
1. Postiz Standard $29/мес — установить через Connectors, закрывает TASKS.md задачу про дублирование контента
2. Apollo Free — зарегистрировать для поиска кандидатов на 9 топ-ролей
3. Brightdata — пропустить

**ПРОМПТ ДЛЯ ПОВТОРЕНИЯ:**
```
Сравни Apollo.io, ZoomInfo и Hunter.io для поиска кандидатов на C-Level роли в AI-стартап.
Нужны: COO, CTO, CMO, CFO, CCO — удалённая работа за equity (акции).
```

---

## ЗАДАЧА 5: Настройка десктопных и мобильных настроек Claude

**СТАТУС:** Завершено (часть настроек включена в процессе чата)

**НАСТРОЙКИ КОТОРЫЕ ВКЛЮЧИЛИ/РЕКОМЕНДОВАНЫ:**

Мобильное приложение (Settings → Capabilities):
- ✅ Search and reference chats — ВКЛЮЧЕНО (было выключено, включили)
- ✅ Generate memory from chat history — уже было включено
- ✅ Dispatch messages notifications — уже было включено
- ⬜ Response completions — рекомендовано включить
- ⬜ Code permission requests — рекомендовано включить

Десктоп веб (claude.ai/settings):
- ✅ Search and reference chats — ВКЛЮЧЕНО
- ✅ Generate memory from chat history — включено
- ⬜ Connector discovery — рекомендовано включить
- Tool access mode: "Load tools when needed" — оставить как есть

**PROFILE — Instructions for Claude (рекомендованный текст):**
```
Я Валерий Шикардос, основатель AI-контент проекта Shikardos Team. 
Всегда отвечай на русском языке. Я арбитражник и маркетолог с 2009 года, 
фокус на AI-автоматизации с 2023. Работаю удалённо из Анталья, Турция (UTC+3). 
Веду закрытую команду 5 000 контент-мейкеров. 
Предпочитаю конкретику без воды, короткие ответы если задача простая, 
подробно только когда реально нужно.
```

**"What best describes your work?"** → выбрать **Marketing**

**ARTIFACTS/VISUALS:**
- Включить в Settings → General → Visuals → Artifacts
- Что это: отдельное окно для кода, документов, интерактивных виджетов
- Полезно для: калькулятор ROI инвестора, визуализация орг-структуры, интерактивные дашборды

**ПРОМПТ ДЛЯ ПОВТОРЕНИЯ:**
```
Проверь мои текущие настройки Claude и скажи что стоит включить/выключить 
для максимальной эффективности при работе с Cowork проектом.
```

---

## ЗАДАЧА 6: Понимание структуры проекта в новой папке

**СТАТУС:** Завершено

**СТРУКТУРА ПАПКИ (на 31 мая 2026):**
```
D:\CLAUDE Cowork\shikardosteam.github.io\
├── index.html              — главный сайт (~6 948 строк)
├── dashboard.html          — дашборд Command Center (22 раздела)
├── TASKS.md                — список задач
├── README.md
├── CNAME
├── robots.txt
├── sitemap.xml
├── favicon.png
├── founder.JPG
├── create_crm.py           — скрипт CRM
├── CRM_Shikardos_Team_2026-05-02.xlsx
├── Распределение акций управляющей команде Shikardos Team.xlsx
├── memory-base/            — база знаний (20+ файлов)
│   ├── CLAUDE.md           — точка входа (читать первым!)
│   ├── 01-idea.md
│   ├── 02-math.md
│   ├── 03-team.md
│   ├── 04-tech.md
│   ├── 05-design.md
│   ├── 06-current-status.md
│   ├── 07-open-questions.md
│   ├── 08-archive-discarded.md
│   ├── founder.md
│   ├── rules-of-work.md
│   ├── investor-economics.md
│   ├── efficiency-coefficient.md
│   ├── glossary.md (stub)
│   ├── legal.md (stub)
│   ├── motivation.md (stub)
│   ├── crisis.md (stub)
│   ├── tz-bot-dashboard.md (stub)
│   ├── tz-cto.md (stub)
│   ├── youtube-mechanics.md (stub)
│   ├── launch-roadmap.md (stub)
│   ├── communications.md
│   └── source-docs/        — бэкапы исходных .md
├── backups/2026-05-01/     — бэкап сайта
├── cabinet/                — личный кабинет
├── webapp/                 — веб-приложение
└── playbooks/              — этот файл здесь
```

**BASH ПУТЬ:** `/sessions/sweet-hopeful-cannon/mnt/shikardosteam.github.io/`
(Внимание: сессия-хэш меняется между сессиями! Всегда проверяй через `ls /sessions/`)

**ПРОМПТ ДЛЯ ПОВТОРЕНИЯ:**
```
Прочитай memory-base/CLAUDE.md, TASKS.md и 06-current-status.md. 
Выдай краткую сводку: что готово, что в работе, топ-3 задачи на сегодня.
```

---

## ВАЖНЫЕ ФАКТЫ О ПРОЕКТЕ (быстрая шпаргалка)

- **Старт:** 15 мая 2026, 00:00 UTC+3 (уже прошёл на момент создания плейбука)
- **Модель для рутины:** Sonnet 4.6
- **Модель для сложного:** Opus 4.6
- **Три источника правды:** index.html > Excel > dashboard.html
- **Главное правило:** НЕ использовать document.write в HTML
- **Founder 1%:** НЕ показывать отдельно на сайте — внутри 15% команды
- **Лазерная эпиляция:** ТОЛЬКО в разделе Фаундер
- **Временная зона:** UTC+3 везде (не МСК)
- **5 000 мейкеров — потолок,** удвоение через подрядчиков мейкеров с Года 2

---

*Playbook создан автоматически Claude Sonnet 4.6 · 31 мая 2026*
