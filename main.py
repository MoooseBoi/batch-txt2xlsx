import os
import shutil
import traceback
import json
from openpyxl import Workbook
from concurrent.futures import ThreadPoolExecutor


def column_index(decimal):
    letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    result = ""

    while decimal > 0:
        digit = (decimal % 22) - 1
        result = letters[digit] + result
        decimal //= 22

    return result


def worker(filename, config):
    print(f"extracting {filename}")

    workbook = Workbook()
    sheet = workbook.active

    with open(f"in/{filename}.txt", "r") as file:
        data = file.read()
        data = data.replace("\r\n", "\n")

    rows = data.split("\n")[config.get("row_offset", 0):]

    row_start = config.get("row_buffer", 0) + 1
    col_start = config.get("col_buffer", 0) + 1

    for row, row_data in enumerate(rows, start=row_start):

        cols = row_data.split("\t")[config.get("col_offset", 0):]

        for col, col_data in enumerate(cols, start=col_start):
            col = column_index(col + config.get("col_buffer", 0))
            index = col + str(row)
            sheet[index] = col_data

    workbook.save(filename=f"out/{filename}.xlsx")
    shutil.move(f"in/{filename}.txt", f"carry/{filename}.txt")

    print(f"finished {filename}")


def main():
    with open("config.json", "r") as file:
        config = json.load(file)

    with ThreadPoolExecutor(max_workers=8) as executor:
        futures = [executor.submit(worker, file[:-4], config) for file in os.listdir("in/")]
        for future in futures:
            try:
                future.result()
            except Exception as e:
                traceback.print_exception(e)


if __name__ == "__main__":
    main()
