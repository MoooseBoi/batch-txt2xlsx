## description
An easy to use python tool for turning text files to excel files in batches.

---

## install
first, clone the project to your computer
```sh
git clone https://github.com/MoooseBoi/batch-txt2xlsx.git
cd batch-txt2xlsx
```
alternativly, download and extract the release .zip file.

second, install the required dependencies with
```sh
pip install -r requirements.txt
```

---

## how to use
move your .txt files to `in/` and execute `python main.py` (or double click `run.bat`)
the output of each file in `in/` will be saved as an .xlsx file with the same filename in `out/`
the files that were successfully extract to .xlsx will be moved to carry so you dont lose them :)

---

## config
- `row_offset`: rows to ignore in text files.
- `col_offset`: columns to ignore in text files.
- `row_buffer`: empty rows before starting table in excel sheet
- `col_buffer`: empty columns before starting table in excel sheet
