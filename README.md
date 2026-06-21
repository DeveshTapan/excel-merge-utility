# Excel Merge Utility

A Windows-friendly Python utility for safely combining multiple Excel
workbooks into one output file.

## What it does

- Reads the first worksheet from each selected Excel file.
- Appends files in the order selected by the user.
- Requires identical column names, column count, and column order.
- Reads values as text by default to help preserve identifiers.
- Optionally verifies the merged output against the source data.
- Can move successfully processed source files to an archive folder.
- Provides both a graphical interface and command-line interface.
- Uses generic file and folder selection—there are no machine-specific paths.
- Writes through a temporary file so an interrupted operation does not leave a
  half-written final workbook.

## Requirements

- Python 3.9 or newer
- Windows is recommended for the graphical interface

Install the dependencies:

```powershell
py -3 -m pip install -r requirements.txt
```

## Run the graphical interface

Windows users can double-click:

```text
run_excel_merge.bat
```

The launcher uses its own current folder automatically and contains no
machine-specific path.

Alternatively, run:

```powershell
py -3 .\append_merge_excel.py --gui
```

The GUI asks you to:

1. Select the Excel files.
2. Arrange them in the required append order.
3. Choose an output folder.
4. Optionally choose an archive folder.
5. Click **Merge Now**.

The merged workbook is written to:

```text
<selected-output-folder>\MergedFiles\
```

GUI logs are written to:

```text
<selected-output-folder>\Log\
```

## Command-line examples

Merge two files:

```powershell
py -3 .\append_merge_excel.py file1.xlsx file2.xlsx
```

Choose the output directory:

```powershell
py -3 .\append_merge_excel.py file1.xlsx file2.xlsx --outdir "C:\Output"
```

Use an exact output filename and path:

```powershell
py -3 .\append_merge_excel.py file1.xlsx file2.xlsx --output "C:\Output\my_merged.xlsx"
```

Disable verification:

```powershell
py -3 .\append_merge_excel.py file1.xlsx file2.xlsx --no-verify
```

## Important data rule

Only use workbooks that you are authorized to process. Excel input files,
merged outputs, verification differences, logs, and archived source files are
excluded from this repository by `.gitignore`.

## Validation and safety

- Only the first worksheet from each source workbook is processed.
- Column names, count, and order must match exactly.
- Blank and duplicate column headings are rejected.
- Duplicate file selections are ignored.
- The output cannot overwrite a source workbook.
- Source files are moved only after a successful merge and verification.
- Excel's maximum worksheet row count is checked before writing.
- `defusedxml` is installed to harden XML parsing used by Excel libraries.

## Run the automated tests

```powershell
py -3 -m unittest discover -s tests -v
```
