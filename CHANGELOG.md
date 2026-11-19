# Changelog

## 3.3.1 — 2026-11-19

- Добавлен шаблон «Комитет» по исходному комитету и финальному слайду дебиторки: восемь композиций, нативные графики и таблицы.
- В основной процесс включена проверка аргументов с конкретными предложениями недостающих метрик, сравнений и доказательств.
- Увеличены внутренние поля, сокращён текст примеров, добавлены функциональные разделители и разделы презентации.
- Палитра приведена к слайду по дебиторке; крупные синие акценты удалены.
- Исправлено обнаружение системных шрифтов на macOS; добавлены проверки данных и переполнения таблиц.
- Обновлены README, руководство пользователя и примеры. Для текущего набора пройдены 13 проверок; визуальная проверка выполнена в LibreOffice, автоматический экспорт PowerPoint не подтверждён из-за тайм-аута.

## 2026-10-06 — описание и руководство пользователя

- Описание построено вокруг задач и результата: отчёты, защита решений,
  доклады, отдельные слайды, перевёрстка и исправление дефектов.
- Добавлено руководство: целевая аудитория, помощь в подготовке, подключение и повторное
  использование источников, готовые запросы, практики и FAQ.
- Разделены профиль источников, доступ к сервису и обновление данных;
  исправлен относительный путь в YAML-примере.
- Техническая основа перенесена в справочную часть. Код сборки не изменён.
- Номер версии публикации присваивает SkillStore; предыдущие номера ниже
  относятся к истории разработки пакета.

## 3.1.0 — 2026-08-31

- Made source discovery universal through workspace-local, credential-free
  `.presentation-assembler/sources.yaml` profiles.
- Removed booking-team defaults, embedded dashboard IDs, Wiki slugs and
  provider-specific collectors.
- Clarified the boundary: vendored `presentations-viscomms` provides template
  mechanics and visual QA; assembler owns brief, evidence, storyline, editorial
  approval, objections and delivery.
- Added provider-agnostic evidence handoff and kept layout `family/variant` out
  of the content contract.
- Hardened custom HTML, Chromium isolation, AppleScript path passing, temporary
  files, and template/font SHA-256 verification.
- Replaced global `pip --user` setup with a managed isolated venv and pinned
  dependencies.

## 3.0.3 — 2026-08-27

- Corrected installation instructions: the skill is not yet available in SkillStore, so the supported channel is the versioned `.skill` attachment on the Wiki release page.
- Added clean replacement guidance and an explicit condition for switching to `store install` after the catalog version reaches `ai_checked`.

## 3.0.2 — 2026-08-27

- Made the content approval gate explicit: review must contain all visible slide text and separate speaker theses; an outline alone no longer unlocks deckspec or PPTX generation.
- Moved optional source-skill declarations into standards-compatible `metadata` and added a cross-platform `compatibility` declaration.
- Added production README with installation, preconditions, operation, security, testing and support instructions.
- Added a pinned runtime dependency manifest.

## 3.0.1

- Introduced the standalone evidence-to-PPTX workflow and owned editorial overlays on top of the Visual Communications rendering foundation.
