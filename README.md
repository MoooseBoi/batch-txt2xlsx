### description
An easy to use python tool for turning text files to excel files in batches.

---

### install
first, clone the project to your computer
```sh
git clone https://github.com/MoooseBoi/batch-txt2xlsx.git
cd batch-txt2xlsx
```

second, install the required dependencies by either running the install script or running the following command in your terminal
```sh
pip install -r requirements.txt
```

### how to use
move your .txt files to `in/` and execute `python main.py` 
the output of each file in `in/` will be saved as an .xlsx file with the same filename in `out/`.
the successfully extract .txt files will be moved to `carry/`

### config
- `row_offset`: rows to ignore in text files.
- `col_offset`: columns to ignore in text files.
- `row_buffer`: empty rows before starting table in excel sheet
- `col_buffer`: empty columns before starting table in excel sheet

---

### for none tech-savy users
make sure you install python from [here](https://www.python.org/downloads/), then double click `install.bat` to install the required python dependencies.
to run the program, follow the instructions in the "how to use" section, and then double click the `run.bat`.
if your computer warns you about a security issue window, click "more info" and "run anyways". dont worry, the .bat files are not viruses, they are meant to make your life easier.
