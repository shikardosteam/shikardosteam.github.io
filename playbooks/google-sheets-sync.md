# Playbook: Синхронизация CRM с Google Sheets через Apps Script

> Источник: чат "Shikardos Team Opus 4.6 (Daily)", июнь 2026

---

## ЗАДАЧА: Обновить Google Sheet — добавить вкладку Аутрич
**СТАТУС:** Завершено

**Google Sheet URL:** https://docs.google.com/spreadsheets/d/1SfZjBiDFlfA6AiTIl-1BV8SFeFEGSI5Lp4tNyZJ5kt8/

**Apps Script проект:** https://script.google.com/u/0/home/projects/1iNwXSYZ1Vpok1Uw5hAQWyOqUVWt3BEyEUex3aeoBFO_wMQEW1377YcH8/edit

**ЧТО РАБОТАЕТ:**
1. Открыть Apps Script через Google Sheet → Расширения → Apps Script
2. Вставить код через Monaco API (НЕ через typing):
   ```javascript
   monaco.editor.getModels()[0].setValue(code);
   ```
3. Сохранить через JS: `document.querySelector('[aria-label="Сохранить проект на Диск"]').click()`
4. Выбрать функцию через JS (dropdown не кликается нормально):
   ```javascript
   const listbox = document.querySelector('[aria-label="Выберите функцию"]');
   const options = listbox.querySelectorAll('[role="option"]');
   for (const opt of options) {
     if (opt.textContent.includes('createOutreach')) { opt.click(); break; }
   }
   ```
5. Нажать "Выполнить" — кнопка координаты ~(487, 86)

**Структура скрипта createOutreach():**
- Читает данные из существующих вкладок IG/YT/TG в Google Sheet
- Нормализация хендлов: `norm(h)` — toLowerCase, trim, strip @, remove _.–
- Проверка релевантности: `isRel(t)` — поиск ключевых слов (ai, нейро, контент, affiliate, арбитраж, smm, таргет и др.)
- Сопоставление людей по хендлам между IG/YT/TG
- Расчёт приоритета A/B/C
- Создание вкладки "Аутрич" с 22 колонками, секциями-разделителями, цветовым кодированием

**ЧТО НЕ РАБОТАЕТ:**
- `SpreadsheetApp.getUi().alert()` — вызывает ошибку "Cannot call SpreadsheetApp.getUi() from this context" при запуске из редактора. Заменять на `Logger.log()`
- Typing кода через browser automation — ломает синтаксис, спецсимволы, кавычки. ТОЛЬКО через Monaco API
- Клик по dropdown "Выберите функцию" через coordinate click — не регистрируется. ТОЛЬКО через JavaScript
- Функция `norm` маленькая и стоит первой в файле — dropdown по умолчанию выбирает её, а не createOutreach. ОБЯЗАТЕЛЬНО переключать через JS

**ОШИБКИ КОТОРЫЕ ДОПУСТИЛИ:**
1. Первый запуск выполнил функцию `norm` вместо `createOutreach` — потеряли время
2. `getUi().alert()` — упал на строке 138. Потратили 6 минут ожидания + ещё перезапуск
3. Скрипт работает 5-6 минут (лимит Apps Script = 6 мин для бесплатных аккаунтов). Едва укладывается
4. Bash-песочница не стартовала ~10 раз подряд — пришлось писать Apps Script вместо прямой загрузки данных

**ФАЙЛЫ:**
- Код Apps Script внутри проекта (Код.gs) — https://script.google.com/u/0/home/projects/1iNwXSYZ1Vpok1Uw5hAQWyOqUVWt3BEyEUex3aeoBFO_wMQEW1377YcH8/edit

**ПРОМПТ ДЛЯ ПОВТОРЕНИЯ:**
"Обнови вкладку Аутрич в Google Sheet CRM. Таблица: https://docs.google.com/spreadsheets/d/1SfZjBiDFlfA6AiTIl-1BV8SFeFEGSI5Lp4tNyZJ5kt8/. Apps Script проект уже настроен. Нужно: открыть Apps Script, выбрать createOutreach через JS (не через клик), запустить. ВАЖНО: НЕ использовать getUi().alert() — только Logger.log(). Код вставлять через monaco.editor.getModels()[0].setValue()."

---

## Обновление вкладки Telegram в Google Sheet
**СТАТУС:** Завершено

**ЧТО РАБОТАЕТ:**
Тот же подход: Apps Script + Monaco API. Функция для Telegram проще — просто setValues() с массивом данных + стилизация шапки (золотой цвет).

**Результат:** 59 строк, 8 колонок, золотой заголовок
