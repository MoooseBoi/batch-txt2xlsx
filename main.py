import os
import shutil
import traceback
import json
from openpyxl import Workbook
from multiprocessing import Pool


def column_index(decimal):
    letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    result = ""

    while decimal > 0:
        digit = (decimal % 22) - 1
        result = letters[digit] + result
        decimal //= 22

    return result


def worker(filename, config):
    print(f"Processing {filename}")

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


def worker_process(args):
    filename, config = args
    try:
        worker(filename, config)
    except Exception as e:
        print(f"Error while processing {filename}: {e}")


def main():
    with open("config.json", "r") as file:
        config = json.load(file)

    input_files = os.listdir("in/")

    with Pool() as pool:
        pool.map(worker_process, [(file[:-4], config) for file in input_files])


if __name__ == "__main__":
    main()
