# Excel → Datasphere Upload Wizard

[![tests](https://github.com/jhanseo/datasphere-excel-loader/actions/workflows/tests.yml/badge.svg)](https://github.com/jhanseo/datasphere-excel-loader/actions/workflows/tests.yml)
[![license](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![python](https://img.shields.io/badge/python-3.9%2B-blue.svg)](https://www.python.org/downloads/)

**English** · [한국어](README.md)

A tool that loads selected sheets from Excel files into SAP Datasphere (HANA)
Open SQL Schema tables. It ships as a GUI wizard so business users can run it
themselves, plus a CLI for batch automation.

It picks up only the sheets whose names start with a configured keyword, and
creates the target table automatically — inferring column data types — when it
does not exist yet. Before uploading you can **Preview** each column's inferred
type and description (COMMENT), change them by hand, and then **Check** whether
the types you chose actually fit the data. Excel files protected by corporate
DRM are read through the installed copy of Excel.

> **Note**  
> This tool assumes an SAP Datasphere environment. You need a database user
> with access to an Open SQL Schema. DRM handling requires Microsoft Excel on
> Windows; everything else works on macOS and Linux too. The machine running
> the tool must have its public IP registered in the Datasphere IP allowlist —
> see [Prerequisites](#prerequisites).

## Features

- Upload several Excel files at once (drag and drop supported)
- Automatic header row detection; the row below the header becomes the column COMMENT
- Automatic data type inference, editable in the preview screen
- Per-column `PRIMARY KEY` and `NOT NULL` selection
- Three load modes: `TRUNCATE` / `INSERT` / `UPSERT`
- Choose how to handle values that don't fit a type (null them, skip the row, revert the type, or abort)
- DRM (DRMONE) detection and Excel-mediated reading
- Existing data is preserved when an upload fails (TRUNCATE rollback)
- Every `config.ini` setting is editable from the GUI tabs
- Bundled PyInstaller build script for a standalone executable

## Prerequisites

You need a **database user** for an Open SQL Schema in your Datasphere tenant.
See [Create a Database User](https://help.sap.com/docs/SAP_DATASPHERE/be5967d099974c69b77f4549425ca4c0/798e3fd6707940c3bd2219b2d1ebaac2.html?locale=en-US#create-a-database-user-with-password-based-authentication).

**The public IP address of the machine running this tool must be added to the
Datasphere IP allowlist.** Datasphere blocks direct database connections from
any address that is not on the list, so without this step the connection fails
even when the host, user and password are all correct. A tenant administrator
adds it under *System → Configuration → IP Allowlist → Trusted IPs*, using the
machine's **external (public) IPv4 address** — not the private address shown by
`ipconfig`. Managing the list requires a global role with Data Warehouse General
and System Information privileges, such as DW Administrator. Note that many
corporate networks hand out dynamic public IPs, and connecting over VPN changes
the address you present, so an IP that worked yesterday may stop working. See
[Manage IP Allowlist](https://help.sap.com/docs/SAP_DATASPHERE/9f804b8efa8043539289f42f372c4862/a3c214514ef94e899459f68f4c1e2a23.html).

## Repository layout

| File | Description |
| --- | --- |
| `wizard.py` | PyQt5 GUI wizard (entry point) |
| `excel_loader.py` | Excel reading, type inference, validation and HANA upload engine (also a CLI) |
| `config.ini.example` | Settings template — copy it to `config.ini` |
| `build_exe.bat` | (Windows) installs dependencies and builds the executable in one step |
| `wizard.spec` | PyInstaller build spec |
| `requirements.txt` | Dependencies |
| `tests/` | pytest suite (runs without a real database, Excel or PyQt5) |

## Install and run

```
git clone https://github.com/jhanseo/datasphere-excel-loader.git
cd datasphere-excel-loader
pip install -r requirements.txt
cp config.ini.example config.ini      # Windows: copy config.ini.example config.ini
python wizard.py
```

`config.ini` stores the database password in plain text, so it is listed in
`.gitignore`. Take care not to commit it.

### Building a distributable executable

Run `build_exe.bat` on Windows. It produces a `dist\ExcelUploadWizard\` folder
containing `ExcelUploadWizard.exe` and `config.ini`. **Zip the whole folder** to
distribute it; users unzip it and run the exe. No Python is required on the
target machine, but Microsoft Excel must be installed to handle DRM files.

Setting `ONEFILE = False` to `True` at the top of `wizard.spec` produces a single
`dist\ExcelUploadWizard.exe` instead. That build can fail when antivirus
real-time scanning locks the freshly written executable, and it starts more
slowly. In environments with DRM and corporate antivirus, the default folder
build is recommended.

If `build_exe.bat` won't run at all, execute these three lines in a command
prompt instead:

```
cd /d "path\to\this\folder"
python -m pip install -r requirements.txt
python -m PyInstaller --clean --noconfirm wizard.spec
```

## Screen layout

The wizard has three workflow tabs and five settings tabs.

**1. Select files** — pick the Excel files to upload, through the file dialog or
by dragging them onto the window. DRM-protected files appear with a 🔒 marker in
orange.

**Settings tabs** mirror the sections of `config.ini`: HANA connection (host,
port, user, password, schema, proxy), General (sheet keyword, load mode,
filename prefix, Excel window visibility, default folder), Header (auto
detection, header row number, description row, preview row count), Logging
(folder, level) and Archive (move processed files, archive folder). Press
**[Save settings]** at the bottom to write them to `config.ini`.

**2. Preview** — the left tree lists files and their target sheets; the right
pane shows the selected sheet's columns and a data sample. The column list shows
the column name, original name, inferred type, target type, KEY, NOT NULL,
Comment, null count and sample values. Target type and Comment are editable, and
clearing the "use" checkbox excludes that column from the upload. Clearing a
sheet's checkbox in the tree excludes the whole sheet.

Columns you tick under **[KEY]** become the `PRIMARY KEY`, in the order you tick
them for a composite key. A PRIMARY KEY is implicitly `NOT NULL` in HANA, so
ticking KEY also turns NOT NULL on and it cannot be cleared. **[NOT NULL]** can
be set on its own. Both constraints **only apply when the table is created**.

**3. Execution log** — shows progress and results live; the same content is
written to a dated file in the log folder.

## What the buttons do

**[Preview]** opens files through the installed Excel, which is what allows DRM
files to be read. Excel starts up and you complete any security login or
approval prompt it shows. The data read this way is cached in memory and reused
by [Run], so DRM files upload without a separate decryption step.

**[Check]** validates the columns whose data type you changed. When a value
cannot be converted, a dialog lists which sheet and column is affected, how many
rows, and example values — then asks how to handle them. There are four options.
`Convert to NULL` writes NULL into the failing cells only and loads the rest.
`Skip the row` drops entire rows containing a failing cell. `Revert to the
inferred type` rolls back to the safe inferred type so no data is lost.
`Abort the upload` treats the sheet as an error if there is any mismatch. A bulk
selector at the top applies one option to every column at once.

If you ticked [KEY] or [NOT NULL], **[Check]** also verifies two more things.
First, whether the data satisfies the constraints: empty values in a NOT NULL
column, or empty/duplicate values in KEY columns, are listed with the offending
row numbers and the check does not pass — loading would fail anyway, so untick
the box or fix the spreadsheet and check again. Second, whether the target table
already exists. If it does, the constraints cannot be applied, so the wizard
shows your selection next to the table's actual PRIMARY KEY and warns you. The
UPSERT-without-primary-key case is caught here too. If the HANA connection
details or schema are empty this lookup is skipped and a note is shown instead.

**[Run]** uploads the cached preview data when it exists. Without a cache, if
even one DRM-protected file is selected, the wizard shows a message pointing you
to Preview and **aborts the whole run**. When only non-DRM files are selected it
uploads directly without a preview. If you changed types but never ran a check,
it asks once more right before starting.

**[Dry run]** reports which sheet maps to which table and how many rows would be
loaded, without connecting to the database. Useful while setting things up.

## Load modes

Pick one under the **Load mode** setting in the *General* tab.

| Mode | Behaviour | SQL used |
| --- | --- | --- |
| `truncate` | Clear the table, then load (default) | `TRUNCATE` → `INSERT` |
| `insert` | Append to existing data | `INSERT` |
| `upsert` | Update on key match, insert otherwise | `UPSERT ... WITH PRIMARY KEY` |

With `insert`, uploading the same file twice loads it twice — be careful.
`upsert` requires the target table to have a `PRIMARY KEY` and fails without
one. For a table being created, the columns ticked as **[KEY]** in the preview
become the primary key; for an existing table, whatever key it already has is
used. **[Check]** verifies this up front.

## Upload safety

**Existing data is not lost.** In `truncate` mode, HANA's `TRUNCATE` is DDL
rather than DML. If the session has `AUTOCOMMIT DDL` enabled, the `TRUNCATE`
commits immediately, so a later INSERT failure plus rollback leaves the table
empty with no new data. `conn.setautocommit(False)` only covers DML and does not
prevent this.

The loader therefore issues `SET TRANSACTION AUTOCOMMIT DDL OFF` right after
connecting, bringing `TRUNCATE` into the rollback scope. A mid-upload failure
then leaves the previous data intact, and the log records that. If the statement
fails for lack of privilege, the load still proceeds but the log warns that
recovery would not be possible. The statements actually executed appear on the
`[DML]` lines of the log.

**Type, KEY and NOT NULL changes do not apply to existing tables.** The loader
does not re-create a table that already exists, so the type you picked in the
preview and the [KEY]/[NOT NULL] boxes you ticked are ignored in favour of
whatever the table already defines. Both **[Check]** and the confirmation step
just before uploading query the target table's real column types and primary key
and, when they differ from your selection, show a comparison table asking
whether to continue. If the change must be applied, cancel, then alter the
column types or constraints in Datasphere — or drop the table — and run again.

## Preview colour coding

The background colour of a column row shows its state. White means the inferred
type is unchanged, yellow means the type was changed but not yet checked, green
means it passed the check, and red means the check found mismatches and a
handling option has been assigned.

## CLI usage

For batch runs without the GUI, use `excel_loader.py`.

```
python excel_loader.py --config config.ini
python excel_loader.py --config config.ini --dry-run
python excel_loader.py --config config.ini --file A.xlsx --file B.xlsx
python excel_loader.py --config config.ini --only-table T_SALES,T_CUSTOMER
python excel_loader.py --config config.ini --allow-drm
python excel_loader.py --config config.ini --mode insert
```

`--mode` overrides `load_mode` from the config (`truncate` / `insert` /
`upsert`). The CLI has no preview screen, so [KEY]/[NOT NULL] can only be set in
the GUI.

Without `--file`, every Excel file in the `[GENERAL] excel_dir` folder is
processed. The CLI rejects DRM files by default; `--allow-drm` uses the same
Excel-mediated path as the GUI preview.

## Troubleshooting

If DRM preview fails, check that Excel is installed on the machine and that
`xlwings` and `pywin32` are present. If the HANA connection fails, work through
these in order: whether your firewall or proxy allows outbound 443 to the host,
whether a proxy needs to be configured, and **whether your current public IP is
on the Datasphere IP allowlist** — this is the most common cause, and it can
break without warning when your public IP changes or you connect through a VPN.
If the wizard reports that no target sheets were found, check that the sheet
name keyword in the *General* tab matches the actual sheet name prefix.

## Development

The tests run without a real HANA database, Excel or PyQt5. Database access is
replaced with a fake connection, the GUI with the Qt stand-in under
`tests/mockqt`, and the sample workbook is generated by the tests themselves.

```
pip install pandas openpyxl pytest
python -m pytest tests/ -v
```

They cover type inference and notation validation, the four mismatch handling
policies, KEY/NOT NULL constraint checks, the SQL and DDL produced by each load
mode, DRM detection and run blocking, and the GUI's signal and dialog flows.
Real Qt rendering and the xlwings DRM path still need manual verification on
Windows.

Contributions are welcome — please open an issue or a pull request. When
changing code, check that `pytest` passes and that `pyflakes` is clean.

## License

[MIT](LICENSE)
