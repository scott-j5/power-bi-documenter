import sys
import os
import zipfile
import json
import csv
import datetime
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox


# Interprets a .PBIX Layout file and processes the reports and visuals within it
def processJson(fileName, jsonOb):
    visuals = []
    # loop over sections (Tabs) in the Power BI Layout file
    for section in jsonOb["sections"]:
        tabName = section["displayName"]

        # loop over visuals within each section
        for visual in section["visualContainers"]:
            visualDict = {}
            dataFields = []

            # Setters for each of the required visual elements
            if "id" in visual:
                visId = visual["id"]
            else:
                print("Something weird here... Visual with no ID.")
                visId = ''
            if "dataTransforms" in visual:
                dataTransforms = visual["dataTransforms"]
                dataFields = extractDataFields(dataTransforms)
            else:
                dataTransforms = ""
            if "config" in visual:
                visualConfig = visual["config"]
                configFields = extractConfigFields(visualConfig)
            else:
                print("Something weird here... Visual with no Config.")
                visualConfig = ""
            visualDict["report"] = fileName
            visualDict["tab"] = tabName
            visualDict["visId"] = visId
            visualDict["name"] = configFields["name"]
            visualDict["type"] = configFields["type"]
            visualDict["dataFields"] = dataFields
            visuals.append(visualDict.copy())
    return visuals


# Extracts data from a visuals DataTransforms list
def extractDataFields(dataTransformsStr):
    dataFields = []
    field = {}
    dataTransformsOb = json.loads(dataTransformsStr)
    for s in dataTransformsOb["selects"]:
        field["field"] = s["queryName"]
        field["type"] = s["type"]["underlyingType"]
        dataFields.append(field.copy())
    return dataFields


# Extracts data from a visuals config list
def extractConfigFields(configStr):
    configFields = {}
    configOb = json.loads(configStr)
    singleVisual = configOb["singleVisual"]
    if "title" in singleVisual:
        title = singleVisual["title"]
        if "text" in title:
            configFields["name"] = title["text"]
        else:
            configFields["name"] = ""
    else:
        configFields["name"] = ""
    configFields["type"] = singleVisual["visualType"]
    return configFields


# Takes a list of visuals and writes them to a csv file
def writeToCsv(visualList, sourceDirectory):
    rows = []
    for visual in visualList:
        dataRow = {}
        dataRow["report"] = visual["report"]
        dataRow["tab"] = visual["tab"]
        dataRow["visId"] = visual["visId"]
        dataRow["visName"] = visual["name"]
        dataRow["type"] = visual["type"]
        if len(visual["dataFields"]) >= 1:
            for dataField in visual["dataFields"]:
                if dataField["type"] == 1:
                    dataSourceSplit = dataField["field"].split(".")
                    dataRow["queryType"] = dataField["type"]
                    dataRow["fromTable"] = dataSourceSplit[0]
                    dataRow["fromField"] = dataSourceSplit[-1]
                else:
                    dataRow["queryType"] = dataField["type"]
                    dataRow["fromTable"] = "Calc"
                    dataRow["fromField"] = "Calc"

                dataRow["completeFieldName"] = dataField["field"]
                rows.append(dataRow.copy())
        else:
            dataRow["queryType"] = ""
            dataRow["fromTable"] = ""
            dataRow["fromField"] = ""
            dataRow["completeFieldName"] = ""
            rows.append(dataRow.copy())

    # CSV output
    keys = rows[0].keys()
    with open(sourceDirectory + "PBIDOutput-" + datetime.datetime.now().strftime("%Y%m%d-%I%M") + ".csv", "w", newline='') as outputFile:
        dictWriter = csv.DictWriter(outputFile, keys)
        dictWriter.writeheader()
        dictWriter.writerows(rows)
    print("\nFile output to: " + sourceDirectory + "PBIDOutput-" + datetime.datetime.now().strftime("%d%m%Y-%I%M") + ".csv")
    messagebox.showinfo("PBID", "Power BI Documentation Scan Complete")
    # input("\n Scan complete. Press Enter to exit...")
    sys.exit("************** Process Complete **************")


def main(sourceDirectory):
    visualsComplete = []
    # Checks and alters the formatting of the passed path string if required
    if sourceDirectory.endswith('"') and sourceDirectory.startsWith('"'):
        sourceDirectory = sourceDirectory[1:-1]
    if not sourceDirectory.endswith("/"):
        sourceDirectory = sourceDirectory + "/"

    # Iterate through the pbix files within the folder path specified
    for fileName in os.listdir(sourceDirectory):
        # Check that the file is a pbix file
        if fileName.endswith(".pbix"):
            # Interpret and set the Layout JSON file to a variable
            archive = zipfile.ZipFile(os.path.join(sourceDirectory, fileName), 'r')
            jsonFile = archive.read("Report/Layout")
            jsonContent = json.loads(jsonFile)
            print("Scanning: " + fileName)
            visualsComplete.extend(processJson(fileName, jsonContent))
            print("Finished: " + fileName + "\n")
        else:
            print("Info: " + fileName + " Is not a valid .PBIX file. Ignored.")
    writeToCsv(visualsComplete, sourceDirectory)


# Check current python version and exit if
if sys.version_info[0] < 3:
    print(sys.version_info)
    input("\n **You must be using python version >3 to use PBID. Press Enter to exit...")
    sys.exit("************** Process Complete **************")

root = tk.Tk()
root.withdraw()

# Prompt for folder path and execute main function
print("Please wait for folder select prompt...")
main(filedialog.askdirectory())

