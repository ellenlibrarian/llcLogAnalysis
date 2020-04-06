# This script takes the output from the usage log script on the Learning Commons computers (text files with dates and times) and combines them for sorting to show the usage of various programs.
import pandas as pd
import openpyxl
# Open the file with the list of log file names and write to a list. Clean the list up by removing the \n.
fapp = open("applications.txt", "r")
applications = []
for line in fapp:
    applications.append(line)
applications[:] = [s.replace("\n", "") for s in applications]
# Open the spreadsheet and set up the worksheet.
wb=openpyxl.Workbook()
ws=wb.active
# Set up the master list of lists and add the column headers as the first entry
masterlist = [["Date", "Hour", "Minute", "Second", "Program", "Count"]]
# Iterate over the list, add the application name at the end of each line, split lines and write the data to the mastr list.
for i in range(len(applications)):
    fname = applications[i]
    fh = open(fname)
    for line in fh:
        log = line.rstrip() + ":" + applications[i].replace(".txt", "") + "\n"
        log = log.split(":")
        list = [log[0], log[1], log[2], log[3], log[4].rstrip(), "1"]
        masterlist.append(list)
for row in masterlist:
    ws.append(row)
wb.save("stats.xlsx")
# Sort data by number of uses, write to spreadsheet.
data = pd.read_excel('stats.xlsx')
byapplication = data.groupby('Program').Count.count().sort_values(ascending=False)
cleandata = pd.DataFrame(data, columns=["Date", "Hour", "Minute", "Second", "Program"])
with pd.ExcelWriter("stats.xlsx") as writer:
    cleandata.to_excel(writer, sheet_name="All usage",index=False)
    byapplication.to_excel(writer, sheet_name="Usage by count")
