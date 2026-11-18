# Source profiles

Store reusable source declarations in `.presentation-assembler/sources.yaml` in
the workspace. The file may be committed to a team repository, but it must not
contain tokens, passwords, cookies, OAuth data or other credentials.

Each profile has a unique name, a short `description`, optional `use_when`, and
one or more sources with `kind`, free-form `provider`, `locator` and `purpose`.
Relative file locators are resolved from the profile file directory.

Flow:

1. Use sources explicitly provided in the request first.
2. Otherwise inspect the nearest profile file and propose the matching profile.
3. If several profiles match, ask the user; never choose silently.
4. If no profile exists, ask which files, links or connected services contain
   the facts. If a connector is unavailable, request a file or export.
5. After successful use, offer to save a reusable profile and write it only
   after explicit confirmation.

Validate with:

```bash
sh scripts/run.sh scripts/assemble.py sources validate
```

See `examples/sources.example.yaml`.
