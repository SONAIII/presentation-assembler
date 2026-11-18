# Production release checklist

## Package integrity

- [ ] `SKILL.md` frontmatter passes `skill-builder/scripts/quick_validate.py`.
- [ ] `name` equals the top-level directory name.
- [ ] No `.env`, credentials, private keys, tokens, caches, `__pycache__` or `.pyc` files are packaged.
- [ ] The archive contains one top-level `presentation-assembler/` directory.
- [ ] `version.json`, `SKILL.md` metadata and `CHANGELOG.md` agree on the release version.

## Runtime acceptance

- [ ] `scripts/preflight.py` completes without blockers.
- [ ] Unit tests pass.
- [ ] `examples/mini_deck.json` builds successfully.
- [ ] Crash-check, lint and composition gates pass.
- [ ] The built deck renders and is visually inspected in Microsoft PowerPoint on macOS.
- [ ] At least one custom HTML smoke test passes when Chromium is part of the supported environment.

## Product behaviour

- [ ] Brief, content-review and final-render STOP points are preserved.
- [ ] Content-review shows full canvas text and separate speaker theses.
- [ ] No deckspec or PPTX is generated before explicit content approval.
- [ ] Numerical claims are evidence-backed; missing data is never fabricated.
- [ ] Restyle mode preserves approved text, notes and images.

## Release communication

- [ ] Installation, update and rollback instructions are published.
- [ ] Supported platforms and known limitations are explicit.
- [ ] Security/data-flow statement is current.
- [ ] Examples cover a new deck, a data-backed update and restyling.
- [ ] Support queue and reproducible bug-report format are listed.
