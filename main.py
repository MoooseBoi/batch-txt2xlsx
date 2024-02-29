import os
import shutil
import traceback
import yaml
from openpyxl import Workbook
from concurrent.futures import ThreadPoolExecutor


def column_index(decimal):
    letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    result = ""

    while decimal > 0:
        carry = decimal % 22
        result = letters[carry] + result
        decimal //= 22

    return result or "A"


def worker(filename):
    print(f"extracting {filename}")

    workbook = Workbook()
    sheet = workbook.active

    with open(f"in/{filename}.txt", "r") as file:
        data = file.read().replace("\r\n", "\n")

    for i, row_data in enumerate(data.split("\n")):
        if i == 0:
            continue

        for j, col_data in enumerate(row_data.split("\t")):
            col = column_index(j)
            sheet[col + str(i)] = col_data

    workbook.save(filename=f"out/{filename}.xlsx")
    shutil.move(f"in/{filename}.txt", f"carry/{filename}.txt")

    print(f"finished {filename}")


def main():
    with ThreadPoolExecutor(max_workers=8) as executor:
        futures = [executor.submit(worker, file[:-4]) for file in os.listdir("in/")]
        for future in futures:
            try:
                future.result()
            except Exception as e:
                traceback.print_exception(e)


if __name__ == "__main__":
    main()
