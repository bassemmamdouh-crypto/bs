# bs

## Cursor Cloud specific instructions

### What this repo is
This repo is a single-file Python helper library, `metabase_client.py`, exposing
`ret_metabase(question, use_query=False, filters=None, warehouse=None)` — a client that
authenticates to a remote Metabase BI instance and returns query results as a pandas
`DataFrame`. There is **no server, web frontend, build step, or test suite**; it is a
library imported by other code, not a standalone runnable app.

### Dependencies
Runtime deps are `pandas` and `requests` (see `requirements.txt`). The startup update
script installs these; no other setup is required.

### Running / exercising the code
- It has no CLI/`__main__`. Use it via `from metabase_client import ret_metabase`.
- Import from the repo root (or set `PYTHONPATH` to the repo root) since it is a
  top-level module, not an installed package.

### Non-obvious caveats
- **Real end-to-end runs need credentials + network.** `IRAQ_METABASE_USERNAME` /
  `IRAQ_METABASE_PASSWORD` in `metabase_client.py` are hardcoded placeholders, and the
  default path calls `https://bi.marbah.info/api`. Without valid credentials + network
  access to that host, the default path fails at `POST /api/session` (401).
- **To exercise the core code path locally without real credentials**, run a small mock
  HTTP server implementing `POST /api/session` (return `{"id": "<token>"}`) and
  `POST /api/card/{id}/query/csv` (return CSV), then override the module constants at
  runtime (`metabase_client.IRAQ_METABASE_BASE_URL`, `..._USERNAME`, `..._PASSWORD`)
  and call `ret_metabase(...)`. This runs session auth, filter-param building, the CSV
  query, and pandas parsing end-to-end.
- **`use_query=True` is non-functional in this repo**: it calls `snowflake_query(...)`,
  which is **not defined anywhere here** and must be provided by the importing
  environment. It will raise `NameError` if invoked as-is.
