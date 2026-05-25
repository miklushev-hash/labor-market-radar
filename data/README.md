# Хранение периодов

Здесь хранятся собранные периоды проекта:

- недели;
- месяцы;
- кварталы;
- полугодия;
- годы.

## Принцип

Один период = одна папка.

## Примеры структуры

### Неделя

```text
data/
  2026-W16/
    raw_items.jsonl
    normalized_items.jsonl
    weekly_metrics.json
    weekly_digest_multilayer.md
```

### Месяц

```text
data/
  2026-03/
    monthly_digest.md
```

### Квартал

```text
data/
  2026-Q1/
    quarterly_digest.md
```

### Полугодие

```text
data/
  2026-H1/
    halfyear_digest.md
```

### Год

```text
data/
  2025/
    annual_digest.md
```

## Что хранится внутри недельной папки

- `raw_items.jsonl` — сырые материалы после сбора;
- `normalized_items.jsonl` — очищенные и размеченные записи;
- `weekly_metrics.json` — агрегаты и сравнения;
- `weekly_digest_multilayer.md` — итоговый Markdown-выпуск недели.

## Что хранится внутри месячной папки

- `monthly_digest.md` — итоговый Markdown-выпуск месяца.

## Что хранится внутри квартальной папки

- `quarterly_digest.md` — итоговый Markdown-выпуск квартала.

## Что хранится внутри полугодовой папки

- `halfyear_digest.md` — итоговый Markdown-выпуск полугодия.

## Что хранится внутри годовой папки

- `annual_digest.md` — итоговый Markdown-выпуск года.

## Что лежит отдельно в `presentations`

Для каждого собранного периода рядом должны существовать итоговые файлы для чтения и печати.

Актуальные HTML- и PDF-версии обзоров лежат в корне `presentations`. Папка `presentations/Архив` хранит неактуальные или исторические версии и не является источником для заполнения `Хранилище.xlsx` без отдельного решения пользователя.

### Неделя

Если неделя собрана как короткий выпуск:

- `<week>_short_update_brief.html`
- `<week>_short_update_brief.pdf`

Если неделя собрана как полная директорская версия:

- `<week>_director_deck.html`
- `<week>_director_deck.pdf`

### Месяц

Если месяц собран как короткий выпуск:

- `<period>_monthly_brief.html`
- `<period>_monthly_brief.pdf`

Если месяц собран как полная директорская версия:

- `<period>_monthly_director_deck.html`
- `<period>_monthly_director_deck.pdf`

### Квартал

Если квартал собран как короткий выпуск:

- `<period>_quarterly_brief.html`
- `<period>_quarterly_brief.pdf`

Если квартал собран как полная директорская версия:

- `<period>_quarterly_director_deck.html`
- `<period>_quarterly_director_deck.pdf`

### Полугодие

Если полугодие собрано как короткий выпуск:

- `<period>_halfyear_brief.html`
- `<period>_halfyear_brief.pdf`

Если полугодие собрано как полная директорская версия:

- `<period>_halfyear_director_deck.html`
- `<period>_halfyear_director_deck.pdf`

### Год

Если год собран как короткий выпуск:

- `<period>_annual_brief.html`
- `<period>_annual_brief.pdf`

Если год собран как полная директорская версия:

- `<period>_annual_director_deck.html`
- `<period>_annual_director_deck.pdf`
