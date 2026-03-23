from openreward.environments import Server

from spreadsheetbench import SpreadsheetBench

if __name__ == "__main__":
    Server([SpreadsheetBench]).run()
