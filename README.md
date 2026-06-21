# Excel Merge Utility

A Windows-friendly Python utility for combining multiple Excel workbooks into
one output file.

## What it does

- Reads the first worksheet from each selected Excel file.
- Appends files in the order selected by the user.
- Requires identical column names, column count, and column order.
- Reads values as text by default to help preserve identifiers.
- Optionally verifies the merged output against the source data.
- Can move successfully processed source files to an archive folder.
- Provides both a graphical interface and command-line interface.

## Requirements

- Python 3
- Windows is recommended for the graphical interface

Install the dependencies:

```powershell
py -3 -m pip install -r requirements.txt
```

## Run the graphical interface

```powershell
py -3 .\append_merge_excel.py --gui
```

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

Use a particular output filename inside the `MergedFiles` directory:

```powershell
py -3 .\append_merge_excel.py file1.xlsx file2.xlsx --outdir "C:\Output" --output "my_merged.xlsx"
```

Disable verification:

```powershell
py -3 .\append_merge_excel.py file1.xlsx file2.xlsx --no-verify
```

## Important data rule

Only use workbooks that you are authorized to process. Excel input files,
merged outputs, verification differences, logs, and archived source files are
excluded from this repository by `.gitignore`.

