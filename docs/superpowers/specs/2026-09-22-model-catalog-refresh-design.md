# Model catalog refresh ("Update Models") — design

**Date:** 2026-09-22 · **Target release:** v2.2.0 · **Branch:** `feature/model-catalog-refresh`

## Problem

The model dropdown and the per-model reasoning-effort filter are hardcoded
(`Config.AvailableModels`, `Config.ReasoningEffortsForModel`). OpenAI ships new
models and effort levels faster than we can release, and the hardcoded list has
already drifted in both directions:

- Missing: `gpt-6-astra`, `gpt-6-sol`, `gpt-6-luna`, `gpt-5.6-sol/-terra/-luna`,
  and the `max` effort. `Config.LoadFromFile` drops any `<Model>` not in the
  hardcoded array, so an admin cannot even opt in via config.xml.
- Stale: `gpt-5.5-pro`, `gpt-5.4*`, `gpt-4.1-*`, `gpt-5.3-codex` are rejected
  for ChatGPT-account tokens.
- `gpt-5.5` (the current default and the configured model on our RDS servers)
  **retires 2026-10-14**.

## Ground truth (live probes, 2026-09-22, Plus/Pro-lite ChatGPT token)

- `GET https://chatgpt.com/backend-api/codex/models?client_version=X.Y.Z`
  (Bearer + optional `ChatGPT-Account-ID`) returns `{"models":[…]}` with, per
  model: `slug`, `display_name`, `description`, `visibility` (`list`/`hide`),
  `priority`, `default_reasoning_level`, `supported_reasoning_levels[{effort,
  description}]`, `upgrade{model, migration_markdown, retirement_at}` and many
  Codex-only fields (~360 KB). This is what Codex CLI caches in
  `~/.codex/models_cache.json`.
- `client_version` is **required** (400 without; must be `X.Y.Z`) and **gates
  the list**: `0.1.0` → empty, `0.100.0` → 1 model, `0.154.0` → 7,
  `0.155.1` (latest released CLI) → 9 incl. `gpt-6-sol`/`gpt-6-luna`. A
  hardcoded version would silently freeze the list. Inference itself is not
  gated (the app sends no client version on `/responses`).
- The catalog advertises effort `ultra`, but `/responses` rejects it:
  `Invalid value: 'ultra'. Supported values are: 'none', 'minimal', 'low',
  'medium', 'high', 'xhigh', and 'max'.` `ultra` is a Codex-client mode
  (max reasoning + task delegation). The same message comes back for any
  invalid effort on any valid model → a zero-inference way to learn the
  server's global effort set.
- `max` works on `gpt-6-astra`; per-model subsets match the catalog
  (`gpt-5.5` rejects `max`, `gpt-5.6-sol` rejects `minimal`).
- Function tools work on `gpt-6-astra`/`gpt-5.6-sol` with the app's exact request
  shape (catalog `tool_mode`/`use_responses_lite` are Codex-client concepts).
- Omitting `reasoning` (the app's `None`) is served as `medium`.
- `upgrade` means either a hard retirement (`gpt-5.5` → `gpt-5.6-sol`,
  `retirement_at` set) or a soft nudge (`gpt-5.6-sol` → `gpt-6-sol`, no date).
- Latest Codex CLI version: `api.github.com/repos/openai/codex/releases/latest`
  → `tag_name: "rust-v0.155.1"`.

## Design

### Catalog (`Services/Models/`, pure, immutable)

`ModelCatalog` holds `ModelCatalogEntry` items (slug, display name, description,
listed, priority, reasoning levels, upgrade `{model, retirement_at}`), the
probed server effort set (nullable), `fetched_at`, `client_version`.

- `ListedSlugs` — `visibility == list`, priority order. Hidden models are still
  valid in config.xml.
- `EffortsFor(slug)` — `"None"` first, then the model's catalog levels that are in
  the server effort set (probed, else the built-in known set
  `none/minimal/low/medium/high/xhigh/max`), display-cased (`XHigh`, `Max`, else
  capitalized). Catalog `none` is folded into the app's `None`. Unknown model →
  `["None"]`.
- `ResolveEffectiveModel(configured, now)` — configured model if in the catalog
  and not past `retirement_at`; else follow `upgrade.model` (≤ 5 hops, cycle-safe);
  else the default (first listed, non-retired model).
- `ResolveWireEffort(model, effort)` — `None`/empty → omit; effort offered for the
  model → the catalog's exact wire string; not offered → omit (server default) so
  a stale combination never 400s; model unknown → lowercase pass-through.

`BuiltInModelCatalog` — the 7 listed models from the 0.155.1 probe; used until the
first refresh and whenever no cache is readable.

### Refresh (`ModelCatalogUpdater`, admin-initiated)

1. Access token + account id from `CodexAuthService` (normal lazy refresh).
2. `client_version` = `<ModelCatalogClientVersion>` from Program Files config.xml
   if set and valid; else the max of {built-in floor `0.155.1`, last cached
   version, latest `openai/codex` GitHub release}.
3. Fetch + parse + validate (slug/effort charset + length caps, dedupe, ≤ 200
   models, response ≤ 8 MB). Zero usable listed models = failure; the existing
   list is kept.
4. Probe the server effort set with one deliberately invalid effort against the
   first listed model; parse the 400. Unparseable → `null` → built-in known set.
5. Save a trimmed cache (`models.json`, schema 1) atomically to
   `C:\ProgramData\OutlookAI\models.json` (shared; the installer grants
   Authenticated Users Modify) and `%LOCALAPPDATA%\OutlookAI\models.json`
   (fallback). Swap `Config.ModelCatalog` in-process.

### Load / config

- `Config.LoadConfig()` loads the newest readable cache (shared vs per-user by
  `fetched_at`), else built-in, **before** the config layers so validation uses it.
- `<Model>` accepted if in the catalog (case-insensitive, stored canonical);
  `<ReasoningEffort>` accepted if `None` or offered by any catalog model.
- `Config.DefaultModel` = catalog default; `Config.EffectiveModel` =
  `ResolveEffectiveModel(Model, Clock())` (`Clock` is a test seam).
- Wire: `CodexChatService` sends `EffectiveModel` and `ResolveWireEffort(...)`;
  the four reasoning dropdowns (chat, copilot, reports, variants) use
  `EffectiveModel`.

### Settings UI (AI Behavior group)

- `[Update Models]` button beside the model dropdown (explicit colors — Server
  2025 rule).
- Info line under the model: display name + description; retirement warning
  (`Retires 2026-10-14 — switch to gpt-5.6-sol.` / `Retired … requests use …`) or
  a newer-model nudge.
- Status line: `Model list: built-in (as of 2026-09-22)` / `Model list: 7 models
  from ChatGPT, updated <local time>`; after refresh: added/removed slugs,
  shared-save warning, "saved model no longer offered" hint.
- Refresh does not change the saved model/effort; **Save AI Settings** still does.

## Review hardening (code review, 2026-09-22)

A max-effort review (15 verified findings, several reproduced against the
built DLL) changed:

- **Retirements survive the model leaving the list.** A refresh carries a
  vanished model forward (hidden) when it has an announced retirement (date and
  replacement kept, for up to 365 days past the date) or config names it.
  Absence alone never reroutes traffic (Codex round, below).
- **Newest data wins.** The built-in list is dated (`FetchedAt = AsOf`); an older
  models.json loses to a newer build. Copies stamped > 1 day in the future are
  ignored. Dates outside 2000–2200 are rejected.
- **client_version floor** uses the newest cache on disk too (another admin's
  refresh), never a config-pinned version (`client_version_source`).
- **No silent regressions.** A catalog with no parseable reasoning levels is
  rejected; a failed or implausible probe reuses the last probed set.
- **Bounded network steps.** GitHub lookup 10 s (WhenAny, so a stalled proxy
  can't hold the refresh), fetch 60 s, probe 15 s; cancellation disposes the
  response because .NET Framework's stream reads ignore their token.
- **Settings.** Opens on the effective model (a retired saved model shows its
  replacement plus a notice), keeps the intended effort across model switches,
  re-applies config.xml after a refresh, formats dates with the invariant
  culture, and flags a list fetched for a different ChatGPT account
  (`account_id` in the cache). Save-copy failures show in a dialog.
- **Open panes refresh.** `Config.AiSettingsChanged` fires after Save AI
  Settings and after a refresh; chat, copilot, reports and variants re-push
  their effort lists (the WebUI keeps the user's pick when still offered).
- **Installer** no longer writes `<Model>gpt-5.5</Model>` and carries
  `<ModelCatalogClientVersion>` over from the previous config.xml.

## Codex adversarial review (2026-09-22)

Verdict needs-attention, 6 findings (3 high), all accepted:

- **Upgrades keep the installed model.** The installer carries `<Model>` over
  (slug-validated); only a fresh install relies on the catalog default.
  Otherwise an install that never saved AI Settings would jump from gpt-5.5 to
  gpt-6-astra on upgrade.
- **Absence ≠ retirement.** The listing is filtered by client_version and plan,
  inference isn't. Soft upgrades are never turned into retirements; announced
  dates are kept as announced (a model missing early still routes to itself
  until its date). Models named in any config layer (`Config.ModelsNamedInConfig`)
  plus the effective model are kept hidden; Settings lists them as "not in
  ChatGPT's list … still used as configured".
- **No stale overwrite.** Saves take a delete-on-close lock file
  (`models.json.lock`, 5 s), re-read the copy under the lock and keep it if
  newer; the updater then adopts that newer list.
- **Newest metadata wins** when carrying models (known catalogs ordered by
  `FetchedAt`).
- **Atomic reload.** `LoadConfigFromPaths` merges the layers into a scratch copy
  and publishes it once.
- **Hidden replacements are selectable** in the Settings picker.

Second Codex pass: #1, #5 and #6 confirmed resolved. The remaining concurrency
gaps in #2–#4 were closed by **merging under the write lock**:
`ModelCatalogStore.Commit(build)` reads each copy under its lock and lets the
updater build the catalog to write from what is on disk at that moment. If
that copy was fetched after ours, it is adopted and this session's configured
models are added. Otherwise ours is written, with carried metadata taken
newest-first from that copy, the session's list and the built-in list.
`Retain` picks the newest entry per slug *before* judging eligibility (so a
withdrawn retirement stays withdrawn). It also keeps the replacements that
announced retirements route to, transitively (so a route can't dead-end when
the listing omits both models).

Third Codex pass: A and C confirmed resolved. Closed the rest:
- **Withdrawals persist.** A newest entry without a retirement is kept (hidden)
  while an older source still announces one, so the withdrawal survives
  later refreshes and cache round trips.
- **This fetch counts as a metadata source**, so adopting another session's
  newer list keeps a configured model that only this fetch knows about.
- `Commit` reports **the newest committed copy**, the same choice `Load` makes.

## Out of scope / follow-ups

- Background auto-refresh. `None` label semantics (it means "server default",
  served as `medium`). Voice model (not in this catalog).

## Security

Token only sent to the two hardcoded `chatgpt.com` endpoints; never logged. The
GitHub call carries no credentials. Cache content is re-validated on load (a
local user who can write ProgramData can already modify `auth.json`; same
accepted-risk model). Refresh is behind the Settings admin password.
