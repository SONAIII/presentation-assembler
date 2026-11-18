# Deck archetypes: what slides a typical deck consists of

This file is the **first step** for open-brief requests («собери обзор рынка», «нужна стратегия на 2026»). The AI should:

1. Identify the archetype from triggers in the request.
2. Take the archetype's skeleton as a starting outline.
3. **Show the user the plan before building anything** (шаг 2 SKILL.md → STOP).
4. After confirmation — for each row pick the exact family/variant via [families-eda.md](families-eda.md) and the shape rules in [content-assembly-rules.md](content-assembly-rules.md).

**Principle:** the skeleton is a starting point, not gospel. Adapting it to the actual content is normal (see the «Adaptation» section at the end of each archetype). Скелеты калиброваны под семейства шаблона Еды; кегли семейств в скобках (напр. `statement` 166pt, `bodytext` 110pt) — дизайнерские размеры Еды из fill-map семейства (`<family>.json`), при другом шаблоне отличаются.

---

## 0. Карта «контент → семейства» — используй ШИРИНУ шаблона

**User-правило (2026-07-03, ужесточено 2026-07-23): не собирай каждый дек из
одной и той же пятёрки** (cover / two_values / columns / timeline / closing) —
шаблон богаче, и у пользователя не должно складываться впечатление, что скилл
умеет 5 структур. Повторяемость видна и МЕЖДУ деками: постоянный пользователь
на втором-третьем деке получает те же слайды — типовая жалоба тестеров.
Правило выбора: пройди карту ЦЕЛИКОМ, для каждого содержательного блока выпиши
2–3 кандидата и **не бери первого по привычке** — предпочитай семейство,
которого в деке ещё нет; частая пятёрка — не больше половины содержательных
слайдов; одинаковый тип контента показывай разными семействами.

| Тип контента | Семейства-кандидаты (выбирай разные!) |
|---|---|
| 1-2 ключевые цифры | `two_values` · `hero` (цифра+фото+утверждение) |
| 3-5 метрик | `metrics` (стат-сетка) · `impact` (метрики+мини-чарты) · `grouped` (факты в группах) |
| Тезисы / причины / выводы | `takeaways` (3 развёрнутых) · `text_grid` (2×2) · `mixed` (сетка до 8) · `bodytext` (ОДИН тезис крупно) · custom-сетка карточек при 5+ (рецепт — custom-recipes-eda.md; плитки выпилены 0.21.0) |
| ОДНА ключевая мысль на весь слайд | `statement` (заявление 166pt на жёлтом) · `bodytext` (тезис + пояснение) · `hero` (если мысль — цифра) |
| Ссылка / приглашение попробовать | `qr` (QR + призыв; картинку QR даёт пользователь) · `closing` (CTA) |
| Колонки с описаниями | `columns` (4 варианта: фото/нумерованные/карточки) |
| Люди / команда | `team_10a`/`team_10b` (10 чел) · `orgchart` (структура) · `columns` био (2-3 с фото) |
| Процесс / этапы | `steps` · `timeline` (3-6 вех) · `funnel` (воронка) · `gantt` (план-график) |
| Сравнение | `proscons` (плюсы/минусы) · `compare` (таблица) · `charts_compare` (два чарта с дельтами) |
| Динамика / данные | `linechart` · `charts_duo` (BAR+COLUMN) · `datatable` (большая таблица) |
| Картинки | `images` (3 фото / мудборд / скриншоты) · `hero` |
| Сделка / условия | `contract` |
| Структурные | `cover` ×2 · `toc` ×2 · `divider` · `closing` · `contacts` ×2 |

Полный каталог с вариантами: [families-eda.md](families-eda.md).

## 1. Market overview / analytics

**Triggers:** «обзор рынка», «рыночный анализ», «landscape», «конкурентный анализ», «market overview», «исследование рынка», «состояние индустрии»
**Typical length:** 5–8 slides
**Closing:** default **no**. The final slide is «Выводы», not «Спасибо» / not «Вопросы?». Если нужен CTA — см. [editorial-principles.md → Rule 2a](editorial-principles.md).

### Skeleton

| # | Slide type | Семейство | What goes there |
|---|---|---|---|
| 1 | Cover | `cover` | Title of the overview + period («Рынок X, 2026») |
| 2 | Executive summary | `takeaways` (баннер-вывод + 3 колонки) или `bodytext` (один тезис) | The main takeaway of the overview |
| 3 | Market size | `metrics` (N=3 → FILL дизайнерской сетки) | TAM / SAM / growth (or 3 key numbers) |
| 4 | Market dynamics | `linechart` (данные через `charts`) | Динамика по датам, серии = сегменты |
| 5 | Key players | `grouped` (4 факта) или `text_grid` | 4 player categories with short theses |
| 6 | Trends / drivers | `takeaways` (3) · `grouped` (4) · custom-сетка (5) | 3–5 trends: name + 1-line desc |
| 7 | Conclusions | `takeaways` или `bodytext` | What this means for us / for the reader |

### Adaptation

- If **there are 5+ players** — split на 2 слайда (top-4 + остальные списком) или custom-сетка карточек (custom-recipes-eda.md).
- If **no numerical data** on market size — skip slide 3, expand the trends section.
- If the request is **only about competitors** (no market size / no dynamics) — trim to 4 slides: cover, executive summary, players, conclusions.

---

## 2. Strategy / vision

**Triggers:** «стратегия», «vision», «куда идём», «стратегия 2026», «направление развития», «pillars», «дорожная карта стратегии»
**Typical length:** 8–12 slides
**Closing:** default **no** (the final slide is the main thesis or roadmap). If the strategy is presented at a large event — optionally add one.

### Skeleton

| # | Slide type | Семейство | What goes there |
|---|---|---|---|
| 1 | Cover | `cover` | Strategy name + year |
| 2 | Раздел: «контекст» | `divider` | The setup — why now |
| 3 | Where we are now | `metrics` (N=3) | 3 key numbers describing the current state |
| 4 | Раздел: «куда идём» | `divider` | The setup — vision |
| 5 | Vision / manifesto | `bodytext` (тезис 110pt + абзац) | A single coherent manifesto |
| 6 | Pillars / directions | `columns` (нумерованные/карточки) или `grouped` | 3–4 directions with theses |
| 7 | Roadmap / stages | `timeline` (3–4 → кривая, 5–6 → прямая) | Stages with dates |
| 8 | Success metrics | `metrics` (N=4 → GENERATE-сетка) | KPIs we'll measure |
| 9 | Closing thesis | `bodytext` | The one thought to leave the audience with |

### Adaptation

- If pillars **number 5+** — custom-сетка карточек (custom-recipes-eda.md) или split на два слайда.
- If the strategy is **short, for 1 year** — drop the roadmap (slide 7) or merge it with success metrics.
- If **there is no «where we are now»** (young project) — drop slide 3, open with the vision.

---

## 3. Quarterly / periodic report

**Triggers:** «отчёт за квартал», «итоги Q1/Q2/Q3/Q4», «summary за месяц», «годовой отчёт», «итоги периода», «retrospective»
**Typical length:** 6–10 slides
**Closing:** optional (add one if the report is presented externally; for an internal review — no).

### Skeleton

| # | Slide type | Семейство | What goes there |
|---|---|---|---|
| 1 | Cover | `cover` | «Итоги Q1 2026» or «Отчёт за январь» |
| 2 | Executive summary | `takeaways` (баннер + 3 вывода) | The headline of the period |
| 3 | Key metrics | `metrics` (N=4) или `impact` | KPIs for the period (with delta) |
| 4 | What we did | `grouped` (до 4) или custom-сетка | 3–5 main achievements |
| 5 | Trend per direction | `linechart` или `charts_duo` | Chart + комментарий-вывод |
| 6 | Challenges / what we didn't do | `proscons` или `text_grid` | 2–4 problems with honest framing |
| 7 | Plans for the next period | `timeline` или `gantt` | 3–4 events with dates |

### Adaptation

- If **many achievements (>5)** — группировка `grouped` + split на два слайда, или custom-сетка карточек.
- If **no charts** (no time series) — replace slide 5 with `metrics`/`two_values`, or drop it.
- If the report is **only about metrics** (KPI review) — trim to 4 slides: cover, summary, metrics, next steps.
- If the report is **annual** — extend to 10–12 slides with a second pass of «achievements / challenges»; при ≥3 разделах используй `toc` с `active_index` перед каждым.

---

## 4. Pitch / investor deck

**Triggers:** «питч», «инвест-дек», «pitch deck», «продаём идею», «презентация для инвестора», «sales-дек»
**Typical length:** 8–12 slides
**Closing:** optional — CTA через `closing` (title = «Контакты» / «Готовы обсуждать» + контакт). **Никогда «Спасибо»** — см. [editorial-principles.md → Rule 2a](editorial-principles.md).

### Skeleton

| # | Slide type | Семейство | What goes there |
|---|---|---|---|
| 1 | Cover | `cover` | Product/company name + tagline |
| 2 | Problem | `divider` или `bodytext` | 1 sharp thesis about the problem |
| 3 | Solution | `bodytext` | Тезис + paragraph: what we do |
| 4 | How it works | `steps` (3 шага) или `timeline` | Process steps |
| 5 | Market (hero number) | `hero` | TAM hero-число + supporting text |
| 6 | Traction | `metrics` (N=4) | Key growth metrics |
| 7 | Team | `columns` three_photos (2–3 био) или `team_10b` | Photos with names and roles |
| 8 | Ask | `bodytext` | What we want (round, partnership) |

### Adaptation

- If **no team photos** — `columns` три текстовые карточки (имя + экспертиза), фото-варианты без своих фото не выдумываем.
- If **the product is complex** — add a second «how it works» slide (`mixed` — история/инфографика до 8 блоков).
- If **competitors must be mentioned** — add `proscons` или `compare` between «solution» and «market».
- If **the pitch is short (5–6 slides)** — keep only: cover, problem, solution, traction, ask.

---

## 5. Хурал / large brand event

**Triggers:** «хурал», «townhall», «большая презентация», «годовая презентация», «всеобщая встреча команды», «annual review»
**Typical length:** 15–25 slides
**Closing:** **recommended** — CTA через `closing`. Варианты для хурала: «Что запомнить» (повтор exec summary в 3–5 буллетах), «Что дальше команде» с конкретным next step, «Спасибо команде» с **конкретными именами** (не безличное «Спасибо»). **Никогда «Спасибо. Вопросы?»** — это муветон, см. [editorial-principles.md → Rule 2a](editorial-principles.md). Это единственный архетип, где closing default YES — но всегда CTA, не «Спасибо».
**Closing line доклада — отдельным `statement`** (хэндофф Ultima 2026-08-11, O7): если у доклада есть финальная фраза-выжимка («We are turning away four in ten people who want to pay us a premium…»), не хорони её строкой в плашке so-what рядом с таблицей — дай ей собственный слайд `statement` (режим good/neutral по смыслу) ПЕРЕД закрывающим CTA. Финальная мысль, произносимая вслух, заслуживает own real estate.

### Skeleton

| # | Slide type | Семейство | What goes there |
|---|---|---|---|
| 1 | Cover | `cover` | Event name + date |
| 2 | Агенда | `toc` (active_index=0) | Разделы хурала; копия `toc` перед КАЖДЫМ разделом со своим active_index |
| 3 | Hero metrics | `metrics` (N=6 → GENERATE) | 6 headline numbers of the year |
| 4 | Hero metric + chart | `linechart` | One big metric of the year with history |
| 5 | Раздел: «что сделали» | `toc` (active_index=1) | Setup |
| 6 | 4 key projects | `grouped` | By direction / by team |
| 7 | Year timeline | `timeline` (5–6 вех → прямая линия) | Main events of the year |
| 8 | Team — лидеры | `columns` two_photos | 2 people / new leaders |
| 9 | Team — вся команда | `team_10a` (с секциями) или `team_10b` | 10 человек с ролями |
| 10 | Раздел: «куда идём» | `toc` (active_index=2) | Setup for the future |
| 11 | Strategy on one slide | `columns` или `grouped` | 4 directions |
| 12 | Roadmap | `timeline` или `gantt` | Key dates of the next year |
| 13 | CTA («Что запомнить» / «Что дальше команде» / «Спасибо: {имена}») | `closing` | Final slide — конкретный CTA |

### Adaptation

- This is the **most extensible archetype** — between sections you can insert extra slides per sub-topic (`images`, `mixed`, `funnel`, `charts_compare`) as needed.
- Rhythm rule: секционный `toc` (с выделением СВОЕГО раздела) или `divider` every 3–5 content slides — правило выбора в [quality-gates.md](quality-gates.md) «Выбор разделителя».
- If **<10 slides** — this isn't a хурал, it's a strategy (see archetype 2); агенда-повторы не нужны.
- If **>25 slides** — это intent-вопрос про структуру дека (один материал vs N материалов): «Это один доклад на 25+ слайдов, или несколько спикеров — стоит ли разделить по speaker'ам?». Уместен (см. [editorial-principles.md → Rule 16](editorial-principles.md)).

---

## 6. Dashboard / one-pager

**Triggers:** «дашборд», «one-pager», «summary на одном слайде», «всё на 1 слайд», «плотная аналитика», «KPI snapshot»
**Typical length:** 1–3 slides
**Closing:** **no** (by definition — a compact format).

### Skeleton

| # | Slide type | Семейство | What goes there |
|---|---|---|---|
| 1 | (opt) Cover | `cover` | Title + period. If the one-pager stands alone — no cover |
| 2 | Main dashboard | `datatable` (таблица 9×4) или `charts_duo` | Everything key — table or two charts |
| 3 | (opt) Drill-down | `impact` / `grouped` / `two_values` | Drill-down on one direction |

### Adaptation

- If the user asks for **exactly 1 slide** — no cover, only slide 2.
- If **too many metrics don't fit** — собери два one-pagers silently (раздели по смысловым кластерам), покажи в плане. Не задавай limitation-вопрос про приоритеты (см. [editorial-principles.md → Rule 16](editorial-principles.md)). Если приоритет неочевиден из контента — intent-вопрос про **главную мысль** (что должно остаться у аудитории) — это OK.
- If the data is **tabular** — `datatable` / `compare`. If the data is **visual / chart-based** — `linechart` / `charts_duo`.

---

## 7. Internal project status report

**Triggers:** «статус проекта», «проектный отчёт», «status report», «как идёт проект», «отчёт по инициативе»
**Typical length:** 5–8 slides
**Closing:** default **no**.

### Skeleton

| # | Slide type | Семейство | What goes there |
|---|---|---|---|
| 1 | Cover | `cover` | Project name + reporting period |
| 2 | Where we are now | `bodytext` | Тезис + summary paragraph |
| 3 | Progress metrics | `metrics` (N=3) | 3 KPIs: where we are, where we're going, % done |
| 4 | What we did this period | `grouped` (до 4) или custom-сетка | 3–5 achievements: name + 1-line desc |
| 5 | Risks / blockers | `proscons` или `text_grid` | 2–4 risks with honest framing |
| 6 | Next steps | `timeline` или `steps` | 3–5 steps with dates |

### Adaptation

- If **the project just started** — drop slide 4 («what we did»), replace it with a «plan» slide (`steps`/`gantt`).
- If **many risks** (>4) — split across 2 slides (risks vs. mitigations) или custom-сетка карточек.
- If **there's a financial part** — add `metrics` (budget / actual / forecast / variance) или `datatable` after slide 3.

---

## What to do if no archetype fits

If the request doesn't fit any of the 7 archetypes:

1. **First** — check whether it's a one-slide request, not a deck: форма контента → семейство по [content-assembly-rules.md](content-assembly-rules.md) §1.
2. **If** the request really is a new deck type — **ask the user**:
   - Is it close to one of the 7 archetypes?
   - What 4–7 key slides does the user mentally see?
3. **Do not invent** a new archetype silently. Better to assemble slide by slide via the shape rules than to guess at the structure.

---

## Relationship with SKILL.md

After choosing an archetype the AI **must show the user the plan** (шаг 2 SKILL.md) — a numbered list «slide → семейство → черновой контент» — and wait for confirmation (STOP) before building the deckspec. This is the most common point of failure if skipped: N slides «not quite right».

After the plan is confirmed — for each row the AI picks the exact family/variant via [families-eda.md](families-eda.md) and the shape rules in [content-assembly-rules.md](content-assembly-rules.md), then writes the deckspec (шаг 3).
