# Playbook: Очистка диска C: Windows + перенос на D:

> Источник: чат "Shikardos Team Opus 4.6 (Daily)", июнь 2026

---

## ЗАДАЧА: Освободить место на C: (было 350 МБ свободных из 99 ГБ)
**СТАТУС:** Завершено (~22 ГБ освобождено)

---

### Шаг 1: Отключить гибернацию → +12.7 ГБ
**ЧТО РАБОТАЕТ:**
CMD от администратора:
```
powercfg /h off
```
Удаляет hiberfil.sys. Гибернация не нужна если ноут в розетке дома. Сон (sleep) работает без неё.

Вернуть обратно: `powercfg /h on`

**ОШИБКИ:** Валерий сначала набрал `powercfg /schedule off` — такой команды нет. Также запустил CMD без прав админа — выдало ошибку "требуются права администратора".

---

### Шаг 2: Очистить pip кэш → +2-3 ГБ
**ЧТО РАБОТАЕТ:**
CMD (можно без админа):
```
pip cache purge
```
Удаляет 209 файлов кэша загрузок. Установленные пакеты продолжают работать.

---

### Шаг 3: Перенос CapCut на D: через junction → +8.9 ГБ
**ЧТО РАБОТАЕТ:**
1. Закрыть CapCut полностью
2. Вырезать `C:\Users\Acer\AppData\Local\CapCut` → вставить в `D:\CapCut`
3. CMD от администратора:
```
mklink /J "C:\Users\Acer\AppData\Local\CapCut" "D:\CapCut"
```
4. Запустить CapCut — работает

**ЧТО НЕ РАБОТАЕТ / НЮАНСЫ:**
- Старые проекты покажут "Медиаматериалы утеряны" — исходные видео были на C: по другому пути. Нужно перепривязать медиа или удалить старые проекты
- При обновлении CapCut junction может сломаться — придётся повторить

---

### НЕ ДЕЛАТЬ: Удаление Claude кэша
**ЧТО СЛУЧИЛОСЬ:**
Сначала я рекомендовал удалить папку Claude в AppData\Local (~12 ГБ). Но при проверке через WizTree оказалось что 11.5 ГБ из 12 — это `claudevm.bundle` (виртуальная машина для bash-песочницы Cowork), а НЕ кэш. Удаление сломает Cowork.

**Путь:** `C:\Users\Acer\AppData\Local\Packages\Claude_pzs8sxrjxfljc\LocalCache\Roaming\Claude\vm_bundles\claudevm.bundle`

**УРОК:** Всегда проверять через WizTree что именно лежит в папке перед советом удалять.

---

### Перенос сессий Cowork на D: (НЕ сделано, опционально)
**КАК СДЕЛАТЬ:**
1. Закрыть Claude Desktop
2. Вырезать `C:\Users\Acer\AppData\Roaming\Claude\local-agent-mode-sessions` → `D:\CLAUDE Cowork\local-agent-mode-sessions`
3. CMD от администратора:
```
mklink /J "C:\Users\Acer\AppData\Roaming\Claude\local-agent-mode-sessions" "D:\CLAUDE Cowork\local-agent-mode-sessions"
```
**Статус сессий:** ~103 МБ (не критично)

---

## Инструменты для диагностики диска
- **WizTree Portable** — https://wiztreefree.com — сканирует C: за 10 секунд, визуальная карта
- **TreeSize Free** — https://www.jam-software.com/treesize_free — альтернатива

---

## Распределение диска C: (99 ГБ, Acer, Windows 11)
| Папка | Размер | Можно чистить? |
|---|---|---|
| Users\Acer\AppData | 44.2 ГБ | Частично |
| Windows | 20.7 ГБ | Нет (WinSxS 7 ГБ — системное) |
| [Файлы в корне C:] | 14.7 ГБ → 2 ГБ (после удаления hiberfil) | hiberfil.sys удалён |
| Packages (Store apps) | 13.2 ГБ | Удалить неиспользуемые приложения |
| Claude VM | 12.1 ГБ | НЕТ |
| CapCut | 8.9 ГБ → 0 (перенесён) | Перенесён на D: |
| Google/Chrome | 5.2 ГБ | Кэш: Ctrl+Shift+Del |
| pip | 3 ГБ → 0 (очищен) | pip cache purge |
