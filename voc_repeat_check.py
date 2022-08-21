import openpyxl
from openpyxl import Workbook, load_workbook 
import math
import colorama
import pathlib
path = str(pathlib.Path(__file__).parent.resolve())

wb = load_workbook(path+"/ENG_vocab.xlsx")

ws = wb["Main"]
max = ws.max_row 
print(f"There are {max} words in your excel sheet.")
print("Checking for words that are listed more than once…")

def progress_bar(progress, total, color = colorama.Fore.YELLOW):
    percent = 100* (progress / float(total))
    #used symbols: - █ (alt + 219)
    bar = "█" * int(percent) + "-" * (100 - int(percent))
    print(color + f"\r|{bar}| {percent:.2f}%", end = "\r")
    if progress == total:
        print(colorama.Fore.GREEN + f"\r|{bar}| {percent:.2f}%", end = "\r")  


repeat = []
progress_bar(0, max)
for i in range(1,ws.max_row+1):
    for j in range(1,ws.max_row+1):
        f = ws["A" + str(i)].value
        s = ws["A" + str(j)].value

        if f == s and i != j:
            repeat.append(f)


    progress_bar(i+1, max)

print()
print("The words are: ")
for i in repeat:
    print(i)
input()