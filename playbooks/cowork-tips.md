# Playbook: Советы и правила работы с Cowork

> Источник: чат "Shikardos Team Opus 4.6 (Daily)", июнь 2026

---

## Общие правила (добавлены в CLAUDE.md и project instructions)

### Excel — ТОЛЬКО через Python
**Правило:** Excel-таблицы создавать ТОЛЬКО через Python (openpyxl) на компе пользователя. Не через bash-песочницу, не через Apps Script, не через JS.
- Bash в Cowork нестабильный — часто не стартует ("Workspace still starting")
- Apps Script имеет 6-минутный лимит — большие таблицы не успевают
- Python на локальном компе — быстро, надёжно, без лимитов

### Сессии Cowork
- Каждый чат = отдельная сессия с изолированным контекстом
- Project instructions (все данные о Shikardos Team) переносятся автоматически между чатами
- История чата НЕ переносится
- Когда задача завершена — лучше начать новый чат (чистый контекст = быстрее работа)
- Старые чаты можно удалять — ничего не сломается
- Данные сессий: `C:\Users\Acer\AppData\Roaming\Claude\local-agent-mode-sessions\` (~103 МБ)

### Cowork на мобильном
- Cowork доступен ТОЛЬКО на десктопе (Research Preview)
- Мобильное приложение Claude — только обычные чаты, без файлов и браузера
- Для удалённого управления: AnyDesk (бесплатный, лёгкий) или TeamViewer

### Выбор модели
| Модель | Когда |
|---|---|
| Sonnet 4.6 | Рутина: текстовые правки, простые задачи, 80% работы |
| Opus 4.6 | Сложные задачи: Python-скрипты, CRM, парсинг, кросс-файловая работа |
| Opus 4.7 | Архитектурные решения: реорганизация базы, аудит проекта, юр/фин вопросы |

### Спящий режим компа
Для фоновых задач (Claude работает пока ты гуляешь):
Настройки → Система → Питание → "Переводить в спящий режим: Никогда" (при питании от сети)

---

## Файлы базы знаний (memory-base)

Добавлены в этом чате:
- `memory-base/trendhero-workflow.md` — процесс извлечения данных из trendHERO скриншотов
- `memory-base/CLAUDE.md` — добавлено правило №11 про Python для Excel, добавлена ссылка на trendhero-workflow.md в карту базы

---

## Подключённые сервисы

### Google Sheet CRM
- URL: https://docs.google.com/spreadsheets/d/1SfZjBiDFlfA6AiTIl-1BV8SFeFEGSI5Lp4tNyZJ5kt8/
- Apps Script: https://script.google.com/u/0/home/projects/1iNwXSYZ1Vpok1Uw5hAQWyOqUVWt3BEyEUex3aeoBFO_wMQEW1377YcH8/edit
- Аккаунт: shikardosteam@gmail.com
- Вкладки: Instagram, Аутрич, YouTube, Telegram, Приоритет A, Кросс-платформа, С Email, Статистика
