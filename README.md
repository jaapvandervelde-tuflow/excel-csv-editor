# Excel CSV editor

The Excel sheet [[CsvEditor.xlsm]] serves as a `.csv` file editor.

The sheet can load and save data from an underlying `.csv`, and stores additional data like column sizes, splits, and table style in a `.json` sidecar configuration (in the `meta` folder in the same location as the `.xlsm`).

Using a renamed `.cmd` file all the complexity can be hidden from the user, provided the location of the `.xlsm` is configured correctly in the `.cmd`.

## Usage

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
