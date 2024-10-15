import csv
import re

FILE_NAME = "sheet3.csv"
days = {"29_9": [], "30_9": [], "01_10": [], "02_10": [], "other": []}


# this function extracts the name of the file from the filepath
def getFileName(filePath):
    # turning the file path into a raw string
    raw_filePath = repr(filePath)[1:-1]
    # instead of if elsing to select the the of slash in the text, just seperate the text by 2 delimiters
    pattern = "|".join(map(re.escape, "\/"))
    # patern = r'\\|/'
    # pattern = r'[\\/]'
    extracted_filepath = re.split(pattern, raw_filePath)[-1]
    return extracted_filepath.split(".")[0]


def parse_time_range(time_range):
    [start, end] = time_range.split("-")
    start_time = start.replace("h", ".")
    end_time = end.replace("h", ".")

    return (float(start_time), float(end_time))


print(parse_time_range("15h45-18h50"))


# list/arrays are stored as reference pointers so when you use it in an array they are not duplicated like primitive datatypes, but rather if we are changing it, we are directly manipulating it
def insertionSort(inputList, value):
    # if inputList is empty, we populate it with the first value
    if not inputList:
        inputList.append(value)
    else:
        # traversing through every values to check if the value us smaller than it, and will insert it as soon as its smaller than a value
        for i in range(len(inputList)):
            if parse_time_range(value[4]) < parse_time_range(inputList[i][4]):
                inputList.insert(i, value)
                return
                # done

        # if the value is large than all, we add it to the end
        inputList.append(value)


# Open the file with utf-8 encoding
with open(FILE_NAME, mode="r", encoding="utf-8") as file:
    csv_reader = csv.reader(file)

    for row in csv_reader:
        date = row[3]
        if date == "9/29/2024":
            days["29_9"].append(row)
        elif date == "9/30/2024":
            days["30_9"].append(row)
        elif date == "10/1/2024":
            days["01_10"].append(row)
        elif date == "10/2/2024":
            days["02_10"].append(row)

schedule = {"29_9": [], "30_9": [], "01_10": [], "02_10": [], "other": []}

for date in days:
    for items in days[date]:  # Iterating over the list of rows for each date
        insertionSort(schedule[date], items)

with open("data.csv", mode="w", newline="", encoding="utf-8") as file:
    writer = csv.writer(file)

    for date in schedule:
        writer.writerow([f"Entries for {date}:"])
        for items in schedule[date]:
            writer.writerow(items)
