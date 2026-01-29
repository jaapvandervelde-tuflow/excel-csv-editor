# Excel CSV editor

The Excel sheet [[CsvEditor.xlsm]] serves as a `.csv` file editor.

The sheet can load and save data from an underlying `.csv`, and stores additional data like column sizes, splits, and table style in a `.json` sidecar configuration (in the `meta` folder in the same location as the `.xlsm`).

Using a renamed `.cmd` file all the complexity can be hidden from the user, provided the location of the `.xlsm` is configured correctly in the `.cmd`.

## Simple Use

1. Clone the project into a subfolder in your project
```
git clone git@bmt-gitlab.bmt-wbm.local:jaap.vandervelde/excel-csv-editor.git
```

2. (Optional) turn your `.xlsx` into a `.csv`
3. Copy `open-csveditor.cmd` next to your `.csv` file
4. Ensure that `EDITOR_DIR` points to the where CsvEditor.xlsm is (see below)
5. Rename the `.cmd` to `<your .csv name>.cmd` or `edit-<your .csv name>.cmd`

You can now use the `.cmd` to edit the `.csv` with the editor. The editor will remember basic settings like column width, hidden columns, word wrapping, etc. and store the raw data in the `.csv`. If you hit 'Save' from Excel, it will automatically use the export function instaed.

If you want to store your `.csv` files separate from the `.cmd` files, you can update `DATA_DIR` to point to the correct folder with the matching `.csv` file(s). 

## Technical Usage

Set environment variables, e.g.:
```none
set "EXCEL_CSV_PATH=test.csv"
set "EXCEL_CSV_CWD=C:\working\dir"
```
Then open [[CsvEditor.xlsm]] and it will pick up the file and allow the user to edit it.

It is recommended to use a script like [[open-csveditor.cmd]] which automates all this. In fact, if you create a copy of `open-csveditor.cmd` and name it `<name of your file>.csv.cmd` (e.g. `example.csv.cmd`) or `edit-<name of your file>.csv.cmd` (e.g. `edit-example.csv.cmd`) then it will correctly infer the name of `example.csv` from the script name, and try to open that `.csv`. 

Note: ensure that the location of the `.xlsm` file is correctly adjusted in the `.cmd` files you create by updating the line:
```
set "EDITOR_DIR=%SCRIPT_DIR%"
```

By default, it assumes the `.xlsm` is in the same location as the command script, but if you clone this project inside your own, it could be something like:
```
set "EDITOR_DIR=%SCRIPT_DIR%..\excel-csv-editor\"
```

Similarly, the [[open-csveditor.cmd]] script assumes it is in the same folder as the `.csv` when you rename it. If you want to store `.csv` in a different location, update the `DATA_DIR` like:
```
set "DATA_DIR=%SCRIPT_DIR%..\csvs\"
```

The script will autocreate a `metadata/` and `temp/` folder in the same location as the `.xlsm` by default. The `metadata/` folder is used to store generated `.json` sidecar configuration files that store values like column width, hidden columns, word wrapping, etc.