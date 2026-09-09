# `sliger`

<p align="center">
  <img src="https://raw.githubusercontent.com/slidoapp/sliger/main/static/images/sliger-black.png" /> <br />
  <strong>Slide of the tiger</strong> <br />
  <em>Slide the power of Python (and Jinja2) into Google Slides</em>
</p>

<p align="center">

  ![PyPI](https://img.shields.io/pypi/v/sliger)
  ![License: Apache 2.0](https://img.shields.io/badge/License-Apache_2.0-green.svg)
  ![Python](https://img.shields.io/badge/python-3.10%2B-blue.svg)

</p>

## PyCon Italy 2023

[![Saving time, money and sanity by Pythonizing Google Slides](https://img.youtube.com/vi/uCkyusWpMbs/0.jpg)](https://www.youtube.com/watch?v=uCkyusWpMbs)

`sliger` is a small CLI and Python library that talks to the Google Slides and Drive APIs so you can:

- duplicate a presentation or an individual slide
- delete a slide by index
- render [Jinja2](https://jinja.palletsprojects.com/) templates that live in text boxes (`jinjify`)
- replace `![image](path-or-url)` placeholders with real images (`imagify`)

## Install

The CLI, as a uv tool:

```bash
uv tool install sliger
```

As a library in another project:

```bash
uv add sliger
```

From a checkout:

```bash
uv sync
```

Requires Python 3.10+ and [uv](https://docs.astral.sh/uv/). To publish a tagged release, see [Releasing](#releasing).

## Prerequisites

Enable the **Google Slides API** and **Google Drive API** on a GCP project, then authenticate in one of two ways:

**Service account** — a JSON key. Share the presentation with the `client_email` in that file.

```bash
sliger --creds-file service-account.json --presentation-id PRESENTATION_ID jinjify
```

**Your Google user (OAuth)** — a **Desktop app** OAuth client (GCP → APIs & Services → Credentials → OAuth client ID → Desktop app). A Web application client will not work: the CLI listens on a random `http://localhost` port, which only Desktop clients allow. Put the OAuth consent screen in Testing and add your Google account as a test user. The first run opens a browser; the token is cached at `~/.config/sliger/token.json`.

```bash
sliger --client-secrets client_secret.json --presentation-id PRESENTATION_ID jinjify
```

User OAuth requests the Slides scope plus `drive.file` (files this app creates). Copying an existing presentation needs broader Drive access:

```bash
sliger --client-secrets client_secret.json --full-drive \
  --presentation-id PRESENTATION_ID duplicate-presentation --copy-title "Copy"
```

`--creds-file` also accepts client secrets or a saved token. Environment variables: `SLIGER_CREDS_FILE`, `SLIGER_CLIENT_SECRETS`, `SLIGER_TOKEN_FILE`, `SLIGER_PRESENTATION_ID`, `SLIGER_CONFIG_PATH`, `SLIGER_FULL_DRIVE`.

## Usage

Every command needs credentials and a presentation ID. If a presentation lives at
`https://docs.google.com/presentation/d/1ijjVtlf9Jq1Rr0xTOMZWcSAUbfl6oA1aaBickwpUdGQ/edit`
the presentation ID is `1ijjVtlf9Jq1Rr0xTOMZWcSAUbfl6oA1aaBickwpUdGQ`.

Turn off **Tools → Preferences → Use smart quotes** in Google Slides so apostrophes in templates stay as plain `'` / `"` characters Jinja can parse. See the [community thread](https://support.google.com/docs/thread/82024200/the-formatting-on-apostrophes-changes-everytime-i-use-the-grammar-spell-check?hl=en).

### `duplicate-presentation`

```bash
sliger --creds-file creds.json --presentation-id PRESENTATION_ID \
  duplicate-presentation --copy-title "A new presentation"
```

Prints the new presentation ID on stdout. Sharing is **off by default** (the service account already owns the copy). Opt in only if you need a public link:

```bash
sliger --creds-file creds.json --presentation-id PRESENTATION_ID \
  duplicate-presentation --copy-title "A new presentation" --anyone-can-view
```

`--anyone-can-edit` still exists and is intentionally opt-in: it makes the copy writable by anyone with the link.

### `delete-slide` / `duplicate-slide`

Slide numbers are 1-based.

```bash
sliger --creds-file creds.json --presentation-id PRESENTATION_ID delete-slide --id 3
sliger --creds-file creds.json --presentation-id PRESENTATION_ID duplicate-slide --id 3
```

`--slide-number` is an alias of `--id`.

### `jinjify`

Renders Jinja in every text box. Built-in context:

| Name | Meaning |
| --- | --- |
| `now` | `time.localtime()` at render time |
| `strftime` | `time.strftime` plus `%O` for an ordinal day (`1st`, `2nd`, …) |

```
Hi! Today is {{ strftime("%A, %O %B", now) }}
```

renders to `Hi! Today is Friday, 2nd September`.

Pass extra variables as JSON (or a Python dict literal):

```bash
sliger --creds-file creds.json --presentation-id PRESENTATION_ID \
  jinjify --data '{"company_name": "Slido"}'
```

`--dry-run` prints each `original → rendered` change and does not call the Slides API.

```bash
sliger --creds-file creds.json --presentation-id PRESENTATION_ID \
  jinjify --data '{"company_name": "Slido"}' --dry-run
```

Or `--data-file vars.json`. Custom Python functions come from a TOML config:

```toml
[function_map]
greet_pycon = "custom_functions.greet_pycon"
```

```bash
sliger --creds-file creds.json --presentation-id PRESENTATION_ID \
  --config-path config.toml jinjify
```

Each value must be a `module.function` dotted path. The directory that contains the TOML file is added to `sys.path`, so a sibling `custom_functions.py` works. Those functions run with your process privileges — only load config you trust.

```
{{ greet_pycon() }}
```

Jinja on a shape is stored in that shape's alt-text (`sliger:…`), so you can **re-jinjify the same deck** when the warehouse moves. Functions may return a string, a table (`list[dict]` / `TableResult`), an image (`ImageResult` / path / bytes), or `sliger_repeat("items")` to duplicate the slide per row.

Built-in `sql()` and named connections. sqlite ships in core; warehouses are extras (`pip install 'sliger[bigquery]'`, `'sliger[snowflake]'`, `'sliger[databricks]'`, or `'sliger[warehouses]'` for all three). Every connector uses the same `:named` parameters and the same `TableResult` path:

```toml
[connections.db]
type = "sqlite"
url = "file:./metrics.db"

[connections.bq]
type = "bigquery"
project = "env:GCP_PROJECT"
dataset = "metrics"

[connections.snow]
type = "snowflake"
account = "env:SNOWFLAKE_ACCOUNT"
user = "env:SNOWFLAKE_USER"
password = "env:SNOWFLAKE_PASSWORD"
warehouse = "COMPUTE_WH"
database = "ANALYTICS"
schema = "PUBLIC"

[connections.dbx]
type = "databricks"
host = "env:DATABRICKS_HOST"
http_path = "env:DATABRICKS_HTTP_PATH"
token = "env:DATABRICKS_TOKEN"

[function_map]
events_count = "custom_functions.events_count"
```

```
{{ sql("select sum(n) as events from metrics where account = :account_uuid") }}
{{ sliger_repeat("deals") }}
{{ item.name }}
```

URLs work too: `bigquery://PROJECT/DATASET`, `snowflake://USER:PASSWORD@ACCOUNT/DB/SCHEMA?warehouse=WH`, `databricks://token:TOKEN@HOST/sql/1.0/warehouses/ID`. `sql()` and mapped functions pull missing arguments from `--data`. One bad box becomes `[sliger error: …]` and does not abort the rest of the deck. A missing extra fails with `pip install 'sliger[snowflake]'` (or the matching extra).

To add a warehouse: implement `sliger.connectors.base.Connector`, list it in `sliger.connectors.registry.BUILTIN`, and (if the SDK is heavy) add a matching `[project.optional-dependencies]` extra. Third-party packages can expose a `sliger.connectors` entry point or call `sliger.connectors.register()`.

### `render` / `inspect` / `repl`

```bash
sliger --creds-file creds.json --presentation-id TEMPLATE_ID --config-path config.toml \
  render --copy-title "Acme Q3" --data-file acme.json

sliger --creds-file creds.json --presentation-id TEMPLATE_ID inspect --data-file acme.json

sliger --creds-file creds.json --config-path config.toml \
  repl "{{ sql('select 1 as n') }}"
```

`render` copies the template, expands repeaters, jinjifies, imagifies, and writes a provenance line into speaker notes.

### `imagify`

Looks for a text box whose entire content is:

```
![image](<IMAGE_PATH_OR_URL>)
```

The inside is Jinja-rendered first. A local path is uploaded to Drive and inserted at the text box's size/position; an `http(s)://` URL is used directly.

```
![image]({{ revenue_chart(account_uuid) }})
```

```bash
sliger --creds-file creds.json --presentation-id PRESENTATION_ID \
  --config-path config.toml imagify --data '{"account_uuid": "abc"}'
```

`--dry-run` lists placeholders and resolved paths without uploading.

Local files are uploaded only long enough for Slides to fetch them, then deleted from Drive. During that window they are link-readable — do not put secrets in those files. HTTP(S) URLs skip Drive entirely.

## Python API

```python
from sliger import Sliger

client = Sliger("creds.json", "TEMPLATE_ID", config_path="config.toml")
report = client.inspect({"company_name": "Example"})
result = client.render("Example Q3", {"company_name": "Example"})
print(result.url)
```

## Demo

`demo/` is a Streamlit UI that used this library at PyCon Italy 2023. It is not installed with the package.

```bash
uv sync --group demo
uv run streamlit run demo/automater_frontend.py
```

## Development

```bash
uv sync
uv run ruff check sliger tests
uv run ruff format sliger tests
uv run pytest
```

`uv.lock` is the source of truth for dependency versions; CI installs with `uv sync --locked`. Optional: `uv run pre-commit install` to run Ruff on each commit.

## Testing against GCP

`gcloud auth login` authenticates the `gcloud` CLI only. The Python client uses
[Application Default Credentials](https://cloud.google.com/docs/authentication/application-default-credentials)
from `gcloud auth application-default login`, and that token must include Slides
and Drive scopes (the default ADC login does not).

```bash
gcloud auth application-default login \
  --scopes=https://www.googleapis.com/auth/presentations,https://www.googleapis.com/auth/drive

gcloud auth application-default set-quota-project YOUR_PROJECT_ID
gcloud services enable slides.googleapis.com drive.googleapis.com --project YOUR_PROJECT_ID
```

Then, against a throwaway deck (the live test creates and deletes one). From the repo root, if `.secrets/user-token.json` exists you only need:

```bash
export GOOGLE_CLOUD_QUOTA_PROJECT=nnewtonians-questionario1-prd
SLIGER_LIVE=1 uv run pytest -m live -s
```

Or pass an explicit token path (must be a real file, not a `...` placeholder):

```bash
export SLIGER_CREDS_FILE="$PWD/.secrets/user-token.json"
export GOOGLE_CLOUD_QUOTA_PROJECT=nnewtonians-questionario1-prd
SLIGER_LIVE=1 uv run pytest -m live -s
```

## Releasing

1. Register this GitHub repository as a [PyPI Trusted Publisher](https://docs.pypi.org/trusted-publishers/) for project `sliger`, workflow `release.yml`, environment `pypi`.
2. Tag a version that matches `pyproject.toml` and `CHANGELOG.md`:

```bash
git tag v0.2.0
git push origin v0.2.0
```

The release workflow builds the wheel, runs `tests/smoke_test.py` against it, and publishes with `uv publish`.

## Security notes

- Keep service-account JSON keys and OAuth client secrets out of git. The `.gitignore` already ignores `*service_account*.json`.
- Duplicated presentations are **not** world-writable. That used to be the default and is now `--anyone-can-edit`.
- `imagify` still has to make a local upload briefly link-readable so Slides can fetch it; the Drive file is deleted immediately after insert.
