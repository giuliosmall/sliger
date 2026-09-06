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

Requires Python 3.10+ and [uv](https://docs.astral.sh/uv/).

## Prerequisites

1. A GCP service account JSON key with the **Google Slides API** and **Google Drive API** enabled.
2. Share the presentation with the service account email (the `client_email` field in the JSON key).

The CLI also reads `SLIGER_CREDS_FILE`, `SLIGER_PRESENTATION_ID`, and `SLIGER_CONFIG_PATH` if you prefer not to pass those flags every time.

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

Uploaded images are given **anyone-with-the-link reader** access because the Slides API fetches them over HTTP. Do not put secrets in those files.

## Python API

```python
from sliger import Sliger

client = Sliger("creds.json", "PRESENTATION_ID", config_path="config.toml")
new_id = client.duplicate_presentation("Slido @ Example")
client.presentation_id = new_id
client.jinjify({"company_name": "Example", "account_uuid": "abc"})
client.imagify({"account_uuid": "abc"})
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

## Security notes

- Keep service-account JSON keys out of git. The `.gitignore` already ignores `*service_account*.json`.
- Duplicated presentations are **not** world-writable. That used to be the default and is now `--anyone-can-edit`.
- `imagify` still has to publish uploaded images as link-readable; that is a Google Slides API constraint.
