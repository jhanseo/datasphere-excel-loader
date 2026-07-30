# Excel → Datasphere Upload Wizard

[![tests](https://github.com/jhanseo/datasphere-excel-loader/actions/workflows/tests.yml/badge.svg)](https://github.com/jhanseo/datasphere-excel-loader/actions/workflows/tests.yml)
[![license](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![python](https://img.shields.io/badge/python-3.9%2B-blue.svg)](https://www.python.org/downloads/)

[한국어](README.md) · **English**

A tool that uploads specific sheets from Excel files into SAP Datasphere (HANA)
Open SQL Schema tables. It ships as a GUI wizard so business users can run it
themselves, plus a CLI for batch automation.

Only sheets whose names start with a configured keyword are uploaded, and if the
target table does not exist it is created automatically with inferred data
types. Before uploading you can use **Preview** to inspect and edit each
column's data type and description (COMMENT), and **Check** to verify that the
types you changed actually fit the data.

## Prerequisites

You need a **database user**. Create a user in your Datasphere tenant with
access to an Open SQL Schema.
([Create a Database User](https://help.sap.com/docs/SAP_DATASPHERE/be5967d099974c69b77f4549425ca4c0/798e3fd6707940c3bd2219b2d1ebaac2.html?locale=en-US#create-a-database-user-with-password-based-authentication))

**The public IP of the machine running this tool must be added to the
Datasphere IP allowlist.**
([Manage IP Allowlist](https://help.sap.com/docs/SAP_DATASPHERE/9f804b8efa8043539289f42f372c4862/a3c214514ef94e899459f68f4c1e2a23.html))

## Features

- Select and upload multiple Excel files at once (drag and drop supported)
- Automatic header row detection; the row below the header becomes the column COMMENT
- Automatic data type inference, editable in the preview screen
- Per-column `PRIMARY KEY` / `NOT NULL` selection
- Three load modes: `TRUNCATE` / `INSERT` / `UPSERT`
- Choose how to handle type mismatches (convert to NULL · skip the row · revert the type · abort)
- Every `config.ini` setting is editable from the GUI tabs

## Repository layout

| File | Description |
| --- | --- |
| `wizard.py` | PyQt5 GUI wizard (entry point) |
| `excel_loader.py` | Excel reading, type inference, validation and HANA upload engine (also a CLI) |
| `config.ini.example` | Settings template — copy it to `config.ini` |
| `build_exe.bat` | (Windows) installs dependencies and builds the exe in one step |
| `wizard.spec` | PyInstaller build spec |
| `requirements.txt` | Required packages |
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
`.gitignore`. Take care not to commit it to the repository.

### Building a distributable executable
Run `build_exe.bat` on Windows.
When the build finishes, a `dist\ExcelUploadWizard\` folder is created
containing `ExcelUploadWizard.exe` and `config.ini`. **Zip the whole folder** to
distribute it; users unzip it and run the exe. No Python is required on the
user's PC, but Microsoft Excel must be installed to handle DRM-protected files.

Changing `ONEFILE = False` to `True` at the top of `wizard.spec` produces a
single `dist\ExcelUploadWizard.exe` instead.
With that option, the build may fail due to real-time antivirus scanning.
In DRM environments, the default folder build is recommended to reduce
antivirus false positives and speed up startup.

### When build_exe.bat won't run at all

Run these three lines directly in a command prompt.

```
cd /d "path\to\this\folder"
python -m pip install -r requirements.txt
python -m PyInstaller --clean --noconfirm wizard.spec
```

## Screen layout

The wizard is divided into three workflow tabs and five settings tabs.

**1. Select files** — pick the Excel files to upload. Add them through the file
dialog or drag them onto the window. DRM-protected files appear in the list
with a 🔒 marker in orange.

**Settings tabs** mirror the sections of `config.ini`. Every setting is
editable from the GUI: HANA connection (host, port, user, password, schema,
proxy), General (sheet keyword, load mode, filename prefix, Excel window
visibility, default folder), Header (auto detection, header row number,
description row, preview row count), Logging (folder, level) and Archive (move
processed files, archive folder). Press **[Save settings]** at the bottom to
write them to `config.ini`.

**2. Preview** — the left tree lists files and their target sheets; the right
pane shows the selected sheet's columns and a data sample. The column list
shows the column name, original name, inferred data type, target type, KEY,
NOT NULL, Comment, null count and sample values. Target type and Comment are
editable, and clearing the "use" checkbox excludes that column from the upload.
Clearing a sheet's checkbox in the tree excludes the whole sheet.

Columns ticked under **[KEY]** become the `PRIMARY KEY`. Ticking several
columns builds a composite key in the order you tick them. A PRIMARY KEY is
implicitly `NOT NULL` in HANA, so ticking KEY also turns NOT NULL on and it
cannot be cleared. **[NOT NULL]** can be set on its own. Both constraints
**only apply when the table is newly created (CREATE TABLE).**

**3. Execution log** — shows progress and results in real time; the same
content is also written to dated files in the log folder.

## What the buttons do

**[Preview]** uses a DRM bypass path that opens files through the installed
Excel. If DRM files are included, Excel starts up, and if a security login or
open-approval prompt appears, the user simply completes it. The data read this
way is cached in memory and reused as-is by the subsequent [Run], so DRM files
can be uploaded without a separate decryption step.

**[Check]** validates the columns whose data type you changed in the preview,
verifying that the conversion is actually possible. When unconvertible values
are found, a dialog opens showing which sheet and column are affected, how many
rows, and example values — then asks you to choose a handling option per
column. There are four options.
`Convert to NULL and continue` writes empty values into the failing cells only and loads the rest normally.
`Skip the row` excludes entire rows containing a failing cell from the upload.
`Revert to the original (inferred) type` rolls back to the safe inferred type so data loads without loss.
`Abort the upload` treats the sheet as an error if there is any mismatch.
A bulk selector at the top of the dialog applies one option to every column at
once.

If you ticked [KEY] or [NOT NULL], **[Check]** additionally verifies two
things. First, whether the data satisfies the constraints. If a NOT NULL column
contains empty values, or a KEY column contains empty or duplicate values, the
violations and offending row numbers are listed and the check does not pass —
uploading as-is would fail, so untick the box or fix the Excel file and check
again. Second, it queries whether the target table already exists. If it does,
the constraints cannot be applied, so the wizard shows your KEY/NOT NULL
selection side by side with the table's actual PRIMARY KEY and warns you. The
UPSERT-without-PRIMARY-KEY case is also caught at this step. If the HANA
connection details or schema are empty, this lookup is skipped and a note is
shown instead.

**[Run]** uploads the cached preview data when it exists. Without a cache, if
even one DRM-protected file is included, the wizard shows a message pointing
you to Preview and **aborts the whole run**. When only non-DRM files are
selected, running directly without a preview uploads normally. If you changed
types but never ran a check on some columns, it asks once more right before
starting.

**[Dry-run]** reports which sheet maps to which table and how many rows would
be loaded, without connecting to the database. Useful when first setting things
up.

## Load modes

Pick one of three under **Load mode** in the `[⚙ General]` tab.

| Mode | Behaviour | SQL executed |
| --- | --- | --- |
| `truncate` | Clear existing data, then load (default) | `TRUNCATE` → `INSERT` |
| `insert` | Keep existing data and append | `INSERT` |
| `upsert` | Update on key match, insert otherwise | `UPSERT ... WITH PRIMARY KEY` |

With `insert`, uploading the same file twice loads it twice — be careful.
`upsert` requires the target table to have a `PRIMARY KEY` and fails without
one. For a table being newly created, the columns ticked as **[KEY]** in the
preview become the primary key; for an existing table, the key already defined
on that table is used. **[Check]** verifies this condition up front.

**Type, KEY and NOT NULL changes do not apply to existing tables.** The loader
does not re-run `CREATE` on an existing table, so the data types you changed in
the preview and the [KEY]/[NOT NULL] boxes you ticked are ignored in favour of
whatever the table already defines. Both **[Check]** and the confirmation step
just before uploading query the target table's actual column types and PRIMARY
KEY and, when they differ from your selection, show a comparison table asking
whether to continue or cancel. If the change must be applied, cancel, then
alter the column types or constraints in Datasphere — or drop the table — and
run again.

## Preview colour coding

The background colour of a column row shows its state. White means the
inferred type is unchanged, yellow means the type was changed but not yet
checked, green means it passed the check, and red means the check found
mismatches and a handling option has been assigned.

## CLI usage

For batch runs without the GUI, use `excel_loader.py` as before.

```
python excel_loader.py --config config.ini
python excel_loader.py --config config.ini --dry-run
python excel_loader.py --config config.ini --file A.xlsx --file B.xlsx
python excel_loader.py --config config.ini --only-table T_SALES,T_CUSTOMER
python excel_loader.py --config config.ini --allow-drm
python excel_loader.py --config config.ini --mode insert
```

`--mode` overrides `load_mode` from the config (`truncate` / `insert` /
`upsert`). The CLI has no preview screen, so [KEY]/[NOT NULL] can only be set
in the GUI.

Without `--file`, every Excel file in the `[GENERAL] excel_dir` folder is
processed. The CLI rejects DRM files by default; `--allow-drm` uses the same
Excel bypass path as the GUI preview.

## Troubleshooting

If DRM file preview fails, check that Excel is installed on the machine and
that `xlwings` and `pywin32` are present. If the HANA connection fails, check
that your corporate firewall allows outbound 443 to the host, whether a proxy
is required in your environment, and **whether your current public IP is
registered in the Datasphere IP allowlist**. If the wizard reports that no
target sheets were found, check that the sheet name keyword in the
`[⚙ General]` tab matches the actual sheet name prefix.

## Development

The tests run without a real HANA database, Excel or PyQt5. Database access is
replaced with a fake connection, the GUI with the Qt stand-in under
`tests/mockqt`, and the sample workbooks are generated by the tests themselves.

```
pip install pandas openpyxl pytest
python -m pytest tests/ -v
```

They cover data type inference and notation validation, the four mismatch
handling policies, KEY/NOT NULL constraint checks, the SQL and DDL produced by
each load mode, DRM detection and run blocking, and the GUI's signal and dialog
flows. Real Qt rendering and the xlwings DRM bypass still need manual
verification on Windows.

## References

- *Create a HANA Database User* :
    https://help.sap.com/docs/SAP_DATASPHERE/be5967d099974c69b77f4549425ca4c0/798e3fd6707940c3bd2219b2d1ebaac2.html?locale=en-US#create-a-database-user-with-password-based-authentication

- *Manage IP Allowlist* — the public IP of the machine must be registered to connect :
    https://help.sap.com/docs/SAP_DATASPHERE/9f804b8efa8043539289f42f372c4862/a3c214514ef94e899459f68f4c1e2a23.html

- *FAQ: IP addresses used by Datasphere* :
    https://userapps.support.sap.com/sap/support/knowledge/en/3456052

- *How to determine whether an issue is caused by the IP allowlist* :
    https://userapps.support.sap.com/sap/support/knowledge/en/3535314

- *Installing Python* :

    (Windows) https://www.python.org/downloads/windows/
    (MacOS) https://www.python.org/downloads/macos/

---

The code and documentation in this project were written with the assistance of AI (Claude, Codex).
