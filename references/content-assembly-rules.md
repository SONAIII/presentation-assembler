# Content assembly rules: mapping raw user content to slides

This file is the **second entry point** after [deck-archetypes.md](deck-archetypes.md). Archetypes answer the question "what slides does a typical deck of type X have?". This file answers the question "I have a chunk of content — how does it fit on slides?".

When to use:

- The user gave **raw content** without naming an archetype («положи это на слайды», «оформи в презентацию», «вот текст и цифры»).
- An archetype is already chosen, but the content **doesn't fit** the skeleton (too long / too short / non-standard).
- At шаг 2 of SKILL.md the chosen family **doesn't cover the shape of the content** cleanly (5 метрик при 3-карточной сетке, 7 тезисов, 6 вех с датами, etc.).

---

## 1. Content shape → семейство

Полный каталог семейств и контракты слотов — [families-eda.md](families-eda.md); карта «тип контента → семейства-кандидаты» (и правило разнообразия) — [deck-archetypes.md](deck-archetypes.md), раздел 0.

### 1.0 Атомизация: инвентарь смысловых атомов ДО выбора семейства

«Форма контента» — это не «сколько верхнеуровневых блоков», а полный
инвентарь СМЫСЛОВЫХ АТОМОВ (user-правило 2026-07-16: агент разбирает контент
на атомы и собирает слайд из них, а не заливает текст в ближайший паттерн).
Алфавит:

| Атом | Что это | Примеры |
|---|---|---|
| `badge` | короткий ярлык-категория над заголовком | «Шаг 1 · оценка», «Пилот», «Q3» |
| `title` | имя карточки/блока | «Метрика ПБ — рост» |
| `sub` | уточнение при title | «новые партнёры · новые проекты» |
| `pill` | выделенная ключевая сущность (метрика, термин) | «FTE gain», «CR 4,2%» |
| `value` | число как главный элемент | «34%», «180+» |
| `caption` | примечание-мелочь: метод, база, оговорка | «опросно», «замер по итогам» |
| `lead` | жирный лид-ин внутри строки | «**Мониторинг:** бивикли-статусы…» |
| `bullets` | честное перечисление ОДНОРОДНЫХ элементов | список городов, каналов, критериев |
| `body` | связный абзац | пояснение, вывод |

Атомы вкладываются: группа/колонка → карточка → атомы. Инвентарь — внутренний,
пользователю не показывается (правило 8 SKILL.md): «2 группы; в первой 2
карточки {badge, title, pill, caption}; во второй 2 карточки {badge, bullets
с lead}». Семейство выбирается ПОСЛЕ и ПО инвентарю.

Два жёстких правила:

- **Буллет — только честному перечислению однородных элементов** (города,
  каналы, критерии). Шаги, роли, статусы, пары «название + суть + оговорка» —
  НЕ буллеты и не абзацы через точку с запятой: это типизированные атомы,
  каждому — своя визуальная роль на слайде.
- **Соответствие двустороннее.** (а) У каждого видного элемента дизайнерского
  слайда — источник в контенте, иначе слоты пустуют (fill-ratio,
  [quality-gates.md](quality-gates.md)); (б) у каждого ТИПА атома контента —
  своя ВИДИМАЯ роль на слайде: badge остаётся плашкой, pill — пилюлей,
  caption — подписью. Атом, влитый предложением в `desc`, — ПОТЕРЯННЫЙ атом,
  то есть потерянный визуальный якорь. Семейство не даёт ролей под все типы
  атомов → эскалация строго по веткам fill → custom (генерация — фолбэк): богатая
  структура — легитимное основание для custom ([custom-html.md](custom-html.md),
  рецепт «бейдж-карточки» в [custom-recipes-eda.md](custom-recipes-eda.md));
  утрамбовка атомов в текст запрещена так же, как фабрикация недостающих
  полей.

> **Before choosing a family — describe the content shape first** (шаг 2 SKILL.md). Write down: сколько групп и карточек и инвентарь атомов каждой (§1.0). A family fits ONLY if every visually prominent element of the designer slide (pill, photo, caption, metric value) has a source in the user's content — короткий контент в развесистой карточке роутер честно отклонит по fill-ratio (см. [quality-gates.md](quality-gates.md)) и уйдёт в GENERATE. Обратный перекос — структурно богатый контент в бедном слоте — роутер отклонить НЕ может: перелив отсекается атомизацией §1.0 до сборки и тестом B11 по рендеру. Не выдумывай недостающие поля (no-fabrication).

Быстрый справочник по форме контента (роутеры решают FILL↔GENERATE сами и печатают причину — не переспрашивай пользователя):

| Each item shape | Семейство / ветка |
|---|---|
| `{value, label}` × 2 | `two_values` |
| `{value, label}` × 3 | `metrics` → FILL дизайнерской сетки |
| `{value, label}` × 4–12 | `metrics` → GENERATE сетка на бренде |
| 1 hero-число + контекст + 2 заметки | `hero` |
| `{name, desc}` × 2–3 развёрнутые | `columns` (роутер примеряет дизайнерские карточки; не заполнил — реворд или custom) |
| `{name, desc}` × N короткие или N > 3 | реворд под донора (развернуть до 3 / сгруппировать до 4 в `grouped`/`text_grid`) или custom-сетка карточек (custom-recipes-eda.md); плитки выпилены 0.21.0 |
| `{date, desc}` × 3–4 | `timeline` → FILL дизайнерской кривой |
| `{date, desc}` × 5–6 | `timeline` → GENERATE прямая линия |
| `{date, desc}` × 7+ | честная ошибка — дели на два слайда |
| Шаги процесса × 3 (без дат) | `steps` |
| За/против | `proscons` |
| Сравнение по критериям | `compare` (таблица + hero-стата) |
| Табличные данные 4–9 строк × ≤4 колонок (и ≥ ⅔ сетки — полупустая таблица сиротит контент; user 2026-07-29: 3 строки в полноширинной таблице «одиноки» — для 2–3 записей бери `columns`/`metrics`/`takeaways`) | `datatable` |
| План-график по месяцам | `gantt` |
| Воронка 3 уровня | `funnel` |
| Факторы/импакт (3 подписи) | `impact` (⚠ данные мини-чартов passthrough) |
| Динамика по датам с данными | `linechart` (`charts` в deckspec, ISO-даты одного года) |
| Два чарта рядом | `charts_compare` (дельты) / `charts_duo` (BAR+COLUMN) |
| Люди × 10 / оргсхема × 11 | `team` / `orgchart` |
| Фото-подборка | `images` (three_pics / moodboard / screenshots) + `photos` |
| 4 нумерованных факта | `grouped` |
| История/инфографика 8 блоков | `mixed` |
| Крупный тезис + поясняющий абзац | `bodytext` |
| Одна ключевая мысль без пояснений | `statement` (166pt на жёлтом, до 3 строк) |
| QR-код + призыв к действию | `qr` (картинку QR даёт пользователь через `photos`) |
| Вводный абзац + сетка абзацев 2×2 | `text_grid` |

Таблица описывает ПЛОСКИЕ формы: строка `{name, desc}` подходит, только если
item реально состоит из имени и однородного описания. Item с бейджами,
пилюлями, подписями или лид-инами в `{name, desc}` НЕ сворачивается (§1.0,
правило б) — такому контенту нужно семейство с ролями под все атомы,
structure или custom.

Если форма не ложится ни в одно семейство и ни в одну генеративную ветку — честно скажи пользователю, что пока не соберёшь красиво (SKILL.md «что честно НЕ умеет»). НЕ выдумывай раскладку и НЕ задавай limitation-вопрос ([editorial-principles.md → Rule 16](editorial-principles.md)).

---

## 2. Text limits

Числовые лимиты живут НЕ в этом файле: ёмкость каждого слота объявлена в fill-map шаблона (`box_w_in` / `size_pt` / `max_lines`) и **меряется глифами при каждой сборке** — `⚠ overflow: [слот]` в выводе build_deck = сигнал REWORD (переформулируй и пересобери; не сжимай, не пропускай).

Ориентиры при написании контента (чтобы реворды не понадобились):

- Заголовок = вывод в одну видимую строку; двухстрочный допустим там, где спека объявила перенос (chart-слайд — тогда подзаголовок обязан быть пуст, машинный чек).
- Подписи метрик — 2–4 слова; значения — 3–4 знака (иначе единый кегль группы уедет вниз).
- Слова длиннее ~12 букв в узких боксах (описания прямого таймлайна, категории BAR) — реворди заранее.

**What to do when content exceeds a limit:**

1. **Switch family** — длинная цитата → не буллет, а `bodytext`; сложное число «1 234 567 ₽» → не метрика, а текст абзаца.
2. **Split** — тезисы на два слайда с ребалансом (7 → 4+3), метрики 3+2, длинный таймлайн на два. Покажи split в плане на шаге 2.
3. **Shorten only if пользователь дал избыточный текст** — intent-вопрос «какая фраза важнее» уместен. Limitation-вопрос «у нас не помещается» — нет ([editorial-principles.md → Rule 16](editorial-principles.md)).
4. **Speaker notes** — полный текст в `"notes"` слайда (SKILL.md «Заметки спикера»), на канве — краткая версия. Штатный выход для дословных абзацев, которые жалко терять: смысл уезжает спикеру, а не в корзину.

Never silently shrink кегль: структурный тир не мутируется вообще; в контентном сжимаются только value-слоты и только до флоров ([quality-gates.md](quality-gates.md), «Фит-стратегия по тирам»).

### 2.1 Layout integrity

Геометрию считает движок: генеративные ветки проходят машинные чекеры (`verify_tiles` / `verify_timeline` / стилевой чекер агенды) — кривой слайд физически не сохранится. Зона ответственности контента: не проси N, на котором семейство честно откажет (timeline 7+), и предпочитай сплит «впихиванию ещё одного».

### 2.2 Дословный текст пользователя — fit-отчёт одним батчем

Когда пользователь принёс **готовые формулировки** (файл со структурой и текстами, «просто положи на слайды», «причеши, текст не трогай»), правила меняются (user-фидбек 2026-07-07):

- **Молча не переписывай.** Silent-reword правила 16 — для контента, который пишем МЫ; чужие формулировки без согласия не меняются. Ожидания зафиксированы ещё в брифе (SKILL.md, шаг 1, вопрос 2): канва фиксирована, дословность не гарантируется, смысл — гарантируется.
- **Не веди пинг-понг «слайд за раз».** Прогони deckspec через сборку целиком, собери ВСЕ overflow-флаги и покажи ОДИН fit-отчёт: «слайд N, что не помещается» списком, для каждого места меню:
  1. **reword** — два варианта короче; **цифры, термины, названия, оговорки — неприкосновенны**;
  2. **сплит** — разнести на два слайда;
  3. **вместительнее форма** — та же мысль в семействе с большей текстовой ёмкостью (`bodytext`, `text_grid` вместо плиток/метрик);
  4. **в заметки** — полный текст в `"notes"` слайда, на канве краткая версия.
- **Смысл > дословность > краткость.** Reword-вариант, теряющий факт, число или оговорку, не предлагается вообще — лучше сплит или заметки.
- Согласованное внеси в deckspec одним заходом и пересобери один раз, не итерируй по слайду.

---

## 3. When to ask intent-вопрос vs silently adapt

**Asking the user is good ONLY for intent (что пользователь хочет), never for our limitations (наша инфра).** Канон и примеры — [editorial-principles.md → Rule 16](editorial-principles.md): нехватку контента, избыток и нестандартный N (5 метрик, 7 тезисов, 6 вех) движок разруливает сам (роутеры + сплит с ребалансом, выбор по action title) — это НЕ повод для вопроса, показывай решение в плане. Ниже — сигналы, специфичные для сборки контента, где вопрос про СМЫСЛ/ДАННЫЕ уместен:

| Signal | What to do |
|---|---|
| Пользователь принёс **готовый дословный текст**, местами не влезает | **Fit-отчёт одним батчем** (§2.2). Не переписывай молча (не наш текст) и не согласовывай по одному слайду (пинг-понг). |
| **Аудитория/тон неясны** (контент маркетинговый, а запрос «отчёт») | **Intent-вопрос ОК.** «Кому показываем и в каком тоне — питч или регулярный отчёт?» (влияет на плотность, лексику и архетип). |
| Content contains **numbers without context** («рост на 47%») | **Intent-вопрос ОК.** «47% — относительно какого периода / какой базы?» (это про данные у пользователя, не наша инфра). |
| **Charts / tables without data** («сделай график динамики») | **Intent-вопрос ОК.** «Какие точки данных? Дай 4–8 значений — или заменим на текстовый блок про тренд.» |

**Rule:** intent-вопрос upfront — хорошо. Limitation-вопрос («у нас не помещается, что выбираем») — никогда.

---

## 4. Выбор шаблона

- **Шаблон фиксируется в брифе** первым вопросом (жёсткое правило 2 SKILL.md); в deckspec — ключ `template`. Шаблонов **три**: `еда` (**Еда**), `gor-services` (**Городские сервисы**) и `ridetech-khural` (**Хурал Райдтеха**; алиасы: хурал, райдтех) — спрашиваем, какой.
- **Критерии выбора по типу контента:** **Еда** — бренд Еды (фуд/маркетинг Еды); **Городские сервисы** — внутренний корпоративный шаблон (стратегия/питч/отчёт/аналитика); **Хурал Райдтеха** — доклады на хурале Райдтеха: единственный хурал библиотеки, «хурал» ведёт сюда; внутри — доклады разных сервисов райдтех-вертикали (Go/Yango/Самокаты/Драйв/…), сервис выбирает обложку. В речи и брифе называть полностью — «Хурал Райдтеха». Не подходит ни один явно — уточни у пользователя, не додумывай.
- **Хурал Райдтеха: сервис обложки** — первый уточняющий вопрос шаблона (12 вариантов: go/yango/scooters/intercity/transport/pro/shift/drive/gas/powerbanks/ultima; нет ответа → neutral). Режимов black/white нет (обложки/statement/финал тёмные, контентные слайды светлые — фиксировано донорами). Custom-контур есть: рецептурник `custom-recipes-ridetech-khural.md`, дизайн-система `assets/ridetech-khural/deck.css` (светлый дефолт, тёмный `.slide.black`). Номер секции на разделителе — всегда двузначный (01, 02…).
- **Городские: режим `mode` (black/white)** — дефолт ПО АРХЕТИПУ: аналитика/отчёт/IR → `white`, питч/стратегия/хурал → `black`; цвета-акценты (red/blue/yellow) миксуются внутри дека. Многие семейства идут dark/white-парой (выбор варианта — `*_dark`/`*_white`).
- **Городские: орг-структура — всегда custom** (донор #30 бесполезен: отделы/составы уникальны). Идём в custom-HTML по `custom-recipes-gor-services.md`; фото людей — только из материалов пользователя, иначе плейсхолдеры.
- **Стиль кастома и фото-пул следуют из пака автоматически:** дизайн-система (`assets/<tdir>/deck.css`, палитра атомов), шрифты/лэйауты кастома (`layouts/<tdir>/custom-profile.json`) и одобренный фото-пул (`brand-tokens.photo_pool`) берутся из выбранного шаблона — руками не переключаются.
- Тем/режимов внутри Еды нет (одна ось); у Городских — ось black/white (см. правило `mode` выше).

---

## 5. Tone, language, format

| What | Rule |
|---|---|
| Numbers | Digits, not words. «24», not «двадцать четыре». «1,2 млн», not «один и две десятых миллиона». |
| Large numbers | Non-breaking spaces as thousands separators in Russian, not commas: «1 247», not «1,247». Decimals — comma: «12,4», not «12.4» (in RU). |
| Percentages | «47%», no space. If suffix is a separate font run (as in metrics) — «47» + «%» as `suffix`. |
| Dates | «18 мая 2026», «Q2 2026», «2026», «I квартал». **NOT** «18.05.26», **NOT** «May 18, 2026» when the deck is in Russian. |
| Emojis | **Forbidden** on final slides. Ever. |
| Жирный | Канон бренда: жирный — ФАЙЛОВОЕ семейство стиля («YS Text Bold»), не синтетика b=1. В deckspec — разметкой `**фрагмент**` (движок ставит файловый bold); жирнить точечно: цену, метрику, лид-ин — не целые абзацы. |
| Abbreviations | First use — spell out: «КАИ (Корпоративный аналитический инструмент)», then just «КАИ». Exceptions: widely known (КПИ, MAU, DAU, IPO, B2B) — no spell-out. |
| ALL CAPS | Do not use in body. In titles — only if дизайнерский слайд так свёрстан (у Еды заголовки не капсом). |
| Quotes | RU guillemets «…», not "…". Inside guillemets — „…". |
| Dashes | Em dash «—» for pauses, hyphen «-» only inside compound words. NBSP после коротких предлогов ставит движок автоматически (`_typo`). |
| Colon in title | Avoid. «Q2: главное» → «Главное в Q2» or «Q2 — главное». |
| Финальная точка | Самостоятельный абзац (подпись, плитка, факт) не заканчивается точкой — движок срежет автоматически (`_typo`). |
| Таблицы: выравнивание | Правило вёрстки (user 2026-07-20): текстовая колонка — влево, колонка из чисел — вправо, заголовок — как его колонка. В `datatable` движок ставит сам, в остальных ветках проверяет линт (`table-align`). Числа со словами-юнитами («18 тыс») считаются текстом. |
| Таблицы: порядок строк | Строки — осмысленным порядком (группировка по типу, убывание значения), не в порядке поступления данных (user-фидбек 2026-07-16). |

---

## 6. Slide-to-slide connections

| Situation | What to do |
|---|---|
| Нужен ли разделитель (`divider`) перед каждым разделом? | Only before major shifts in meaning. Default: 1 разделитель per 3–5 content slides. For decks ≤ 5 slides — not needed. |
| Агенда с выделением раздела (`toc` + `active_index`)? | Только длинным декам с ≥3 разделами — по копии перед каждым разделом (см. [quality-gates.md](quality-gates.md) «Выбор разделителя»). Короткому деку — простой `divider` или ничего. |
| How many slides between two dividers? | 3–5. Fewer — rhythm breaks. More — the reader forgets which chapter they're in. |
| CTA-слайд («Контакты», «Что дальше») | **Do not add by default.** Only if there's a concrete next step for the audience. **Никогда «Спасибо» / «Вопросы?»** — см. [editorial-principles.md → Rule 2a](editorial-principles.md). |
| Cover at the start | Default **yes** for decks ≥ 3 slides. For a one-pager (1–2 slides) — no. Обложка — ровно одна на дек. |
| Subtitle on cover | Default — period / context («Стратегия 2026», «Итоги Q2», «Май 2026»). If no context — no subtitle, do not invent one. |
| Page numbers | Живут в мастере как авто-поле (`a:fld`) — не трогать ни движком, ни руками: python-pptx показывает кэш, PowerPoint рендерит верно (см. [common-mistakes.md](common-mistakes.md)). |

---

## 7. Things NOT to assemble without an explicit ask

Things the AI **must not add on its own**, even if they would "logically fit":

1. **CTA-слайд** («Контакты», «Что дальше», «Зарегистрируйтесь», «Подробнее на сайте»). Only if there's a concrete next step for the audience. **Никогда «Спасибо» / «Вопросы?»** — см. [editorial-principles.md → Rule 2a](editorial-principles.md).
2. **Placeholder slides with fabricated numbers** («примерно 100 млн», «X% роста»). Better an empty slide with a TODO comment than a made-up fact.
3. **Empty dividers** between blocks (разделитель «just for rhythm»). Every divider must carry meaning (the section's name); otherwise skip it.
4. **Footnote / source** when the user gave no source. Do not invent «Данные: внутренняя аналитика, 2026» if no source was provided.
5. **«Команда» / «О нас» / «Контакты»** when the user gave no list of people / contacts. Do not invent names or roles.
6. **Roadmap / future timeline** when the user only gave a retrospective. First check whether a forward plan exists.
7. **Competitor comparisons** when the user gave no competitors and no theses about them.
8. **Фото со стоков или из интернета.** Только одобренный пул шаблона («Нейро-фото Еда») или фото от пользователя.

**Rule:** a 5-slide deck of honest content beats an 8-slide deck where 3 slides are filler.

---

## Relationship with SKILL.md and archetypes

- [deck-archetypes.md](deck-archetypes.md) — provides the **skeleton** of a deck by request type + карта «контент → семейства» (раздел 0). Used first on open briefs.
- This file — provides the **content-fitting rules** for the skeleton (or in place of the skeleton if no archetype matches).
- [families-eda.md](families-eda.md) — provides the **exact family and variant** for each row of the skeleton. Used last.

Chain for an open brief: archetype → план → confirm (STOP) → assembly rules (validate content against slides) → families-eda (pick family/variant) → deckspec.

Chain for "положи это на слайды": assembly rules (content shape → slide set) → план → confirm (STOP) → families-eda (pick family/variant) → deckspec.
