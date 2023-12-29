import csv
import os.path
import random
import win32print
import win32api

names = []
numbers = []

if os.path.isfile('names.csv') == True:
    with open('names.csv', 'r') as csvfile:
        csvreader = csv.reader(csvfile)
        names = next(csvreader)
        numberOfPeople = len(names)
else:
    numberOfPeople = int(input("Enter number of people: "))
    for i in range(0, numberOfPeople):
        ele = str(input())
        names.append(ele)
    with open('names.csv', 'w') as f:
        csvwriter = csv.writer(f)
        csvwriter.writerow(names)

if os.path.isfile('numbers.csv') == True:
    with open('numbers.csv', 'r') as csvfile:
        csvreader = csv.reader(csvfile)
        numbers = next(csvreader)
    if len(numbers) != numberOfPeople:
        numbers = random.sample(range(numberOfPeople), numberOfPeople)
        numbers = [x + 1 for x in numbers]
        with open('numbers.csv', 'w') as f:
            csvwriter = csv.writer(f)
            csvwriter.writerow(numbers)
else:
    numbers = random.sample(range(numberOfPeople), numberOfPeople)
    numbers = [x+1 for x in numbers]
    with open('numbers.csv', 'w') as f:
        csvwriter = csv.writer(f)
        csvwriter.writerow(numbers)

dict = {names[i]: numbers[i] for i in range(len(names))}
sortedDict = sorted(dict.items(), key=lambda item: item[1])
with open("output.md", "w") as f:
    print("# White Elephant Number Distribution", file=f)
    for key, value in sortedDict:
        print(value, ': ', key, file=f)
        print(value, ': ', key)

printResults = input('Do you want to print these results? (y/n): ').lower().strip() == 'y'

if printResults is True:
    printer_name = win32print.GetDefaultPrinter()
    hPrinter = win32print.OpenPrinter(printer_name)
    filename = "output.md"
    try:
        hJob = win32print.StartDocPrinter(hPrinter, 1, ('PrintJobName', None, 'RAW'))
        try:
           win32api.ShellExecute(0, "print", filename, None, ".", 0)
        finally:
            win32print.EndDocPrinter(hPrinter)
    finally:
        win32print.ClosePrinter(hPrinter)
else:
    print()