# Import packages
import csv
import os.path
import random
import win32print
import win32api

# Declare lists
names = []
numbers = []

file_exists = os.path.exists('./white-elephant-number-distributor')
if not file_exists:
    path = "./white-elephant-number-distributor"
    os.mkdir(path)

outputFileName = "./white-elephant-number-distributor/output.txt"

version = "v0.3.0"
print("You are using version", version, "\n")

# Checks for if names.csv exists, and uses the values in it if it does. Creates a new file and fills it if not.
if os.path.isfile('./white-elephant-number-distributor/names.csv') == True:
    with open('./white-elephant-number-distributor/names.csv', 'r') as csvfile:
        csvreader = csv.reader(csvfile)
        names = next(csvreader)
        numberOfPeople = len(names)
        print("Using values from ./white-elephant-number-distributor/names.csv.\n")
else:
    numberOfPeople = int(input("Enter number of people: "))
    print("Please enter the names of the people that are participating, pressing enter after each name:")
    for i in range(0, numberOfPeople):
        ele = str(input())
        names.append(ele)
    with open('./white-elephant-number-distributor/names.csv', 'w') as f:
        csvwriter = csv.writer(f)
        csvwriter.writerow(names)

# Checks for if numbers.csv exists, and uses the values in it if it does. Creates a new file and fills it if not.
if os.path.isfile('./white-elephant-number-distributor/numbers.csv') == True:
    with open('./white-elephant-number-distributor/numbers.csv', 'r') as csvfile:
        csvreader = csv.reader(csvfile)
        numbers = next(csvreader)
    if len(numbers) != numberOfPeople:
        numbers = random.sample(range(numberOfPeople), numberOfPeople)
        numbers = [x + 1 for x in numbers]
        with open('./white-elephant-number-distributor/numbers.csv', 'w') as f:
            csvwriter = csv.writer(f)
            csvwriter.writerow(numbers)
else:
    numbers = random.sample(range(numberOfPeople), numberOfPeople)
    numbers = [x+1 for x in numbers]
    with open('./white-elephant-number-distributor/numbers.csv', 'w') as f:
        csvwriter = csv.writer(f)
        csvwriter.writerow(numbers)
    print("New numbers were generated.\n")

# Combine both lists into a dictionary
dict = {names[i]: numbers[i] for i in range(len(names))}
sortedDict = sorted(dict.items(), key=lambda item: item[1])
with open(outputFileName, "w") as f:
    print("# White Elephant Number Distribution", file=f)
    print("These are the generated values.")
    for key, value in sortedDict:
        print(value, ': ', key, file=f)
        print(value, ': ', key)

# Asks the user if they would like to print these results to the default printer.
printResults = input('\nDo you want to print these results? (y/n): ').lower().strip() == 'y'

# Prints the results to the system default printer if requested.
if printResults is True:
    printer_name = win32print.GetDefaultPrinter()
    hPrinter = win32print.OpenPrinter(printer_name)
    filename = outputFileName
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