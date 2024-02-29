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
        carry = (decimal % 22) - 1
        result = letters[carry] + result
        decimal //= 22

    return result


def worker(filename, config: dict):
    print(f"extracting {filename}")

    workbook = Workbook()
    sheet = workbook.active

    with open(f"in/{filename}.txt", "r") as file:
        data = file.read().replace("\r\n", "\n")

    # TODO: the for lines are too long and complex, seperate config handle from for loops
    for i, row_data in enumerate(data.split("\n")[config.get("row_offset", 0):], start=1):
        i += config.get("row_buffer", 0)

        for j, col_data in enumerate(row_data.split("\t")[config.get("col_offset", 0):], start=1):
            col = column_index(j + config.get("col_buffer", 0))
            index = col + str(i)
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
