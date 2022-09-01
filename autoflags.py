#   autoflags algorithm
#
#
#   "autoflags.py" python file is the algorithm that takes in an excel file name and uses reccursion to scan a parks calc map and returns a matrix
#   with each field name mapped to a statistic as well as a log with any errors encountered while traversing the calculation map. "flagScript.py"
#   is the script that traverses the DOI server folder achitecture and finds calculation maps. Once the map is found, it uses the "autoflags.py"
#   algorithm and stiches the error logs and flags matrices into one large batch file.
#
#
#   In order to successfully run this file, you need python3 installed, as well as the following python packages... (tqdm, openpyxl, pandas, numpy)
#   You may need to install those packages as well.
#
#   To install those packages, enter the following command into the Command Prompt/Terminal
#
#           "pip3 install --trusted-host pypi.org --trusted-host files.pythonhosted.org <PACKAGE_NAME>"
#
#
#   At the bottom of "autoflags.py" is a debug mode which allows you to perform the flags algorithm on a specific park wihtout compliling a large
#   batch file. To start, make sure you follow the following steps.
# 
#   1. Change the variable "parkFilesFolder" to the correct path to the DOI Park Files path. Your username should be embedded within the path if
#   you did this correctly.
#   2. Change the "debugState" value to "True"
#   3. Change the "debugPark" value to the appropriate 4-letter Park Unit Code ex:"XXXX"
#   4. Open up Terminal(MAC) or Command Prompt(WINDOWS) and navigate to the folder/path that the python file resides in.
#       - You can do this by entering the command 'cd <folder name>' repeatily until you are in the same directory as the python file.
#   5. Enter the command 'python3 autoflags.py' and press <Enter>.
#       - If you did the above steps correctly and setup the proper DOI Server Park Files Folder path, the program should return the flags matrix
#       for that given park, as well as an error log(if errors were present).
#
#   Tip: If you struggle with Command Prompt/Terminal abreviating long results and hiding the middle rows with ".....", uncomment the the 4
#   "pd.set_option()" lines. (The second pd.set_option line has two comment symbols because it isn't normally needed unless you struggle with
#   column width formatting problems)
#
#
#
#
#
#   ***For information on how to decode error log statements, see below.
# 
#   The Error log is formatted as follows...
#   PARK: The 4-letter Park Unit Code
#   SHEET: The sheet number in which the error resides (if it's 0, there is a file error)
#   TYPE: The error type, see the explanations below.
#   PROBLEM: Specific Problem and location information, see the explanations below.
#
#
#   ERROR TYPE                                               DESCRIPTION
#   "Formatting Error" ....................................  (There is a problem with the Calc Map's format)
#
#       - "Unknown Stat - <fieldcode>(<excel location>)"     (There is a calculation field that isn't being referenced anywhere else on the sheet
#                                                            and its field name is not an official statistic)
#                                                            This means that either an offical statistic isn't named correctly within the Calc
#                                                            Map, or there is a cell that should have been referenced in a formula, but there was
#                                                            a typo and its not being referenced. Hint: The error returns the excel location in
#                                                            which there is a problem, as well as the sheet number
#
#       - "Blank Field Name Located at <excel location>"     (There is a field name that isn't formatted correctly (It must be seperated by a
#                                                            dash surrounded by spaces ' - '))
#
#   "Formula Error" .......................................  (There is a problem with a cell's formula)
#
#       - "Pointing to empty Cell(<excel location>)"         (There is a problem with a formula as it is referencing a cell with no field name)
#
#   "Calc Map Error" ......................................  (There is a problem with the Calc Map specifically)
#
#       - "Skipped Park"                                     (The program skipped over this park because it was referenced to not be included in
#                                                            the batch file, probably due to a calc map that needs fixing.
#
#   "File Error" ..........................................  (There is a problem with the file in the folder)
#
#       - "No file named (XXXX Calc Map.xlsx)"               (There is no calc map in the parks folder, or the calc map isn't named correctly. It
#                                                            should be named "XXXX Calc Map.xlsx")
#   








import openpyxl as op
import pandas as pd
import numpy as np
import re
import os
from itertools import groupby
pd.set_option('mode.chained_assignment', None)


def runFile(fileName): #Function to run each Calc Map File
    path = os.getcwd() #Getting and setting the current Working Directory for future ref.
    file = fileName # Setting file as the filename parameter.
    name = file # Setting the file as the name
    wb = op.load_workbook(os.path.join(path, file), read_only=True) # Opening a read-only copy of the excel file workbook
    sheets = len(wb.worksheets) # Storing the number of sheets in the worbook

    #Declaring the official stats that will be looked for and used.
    stats = np.array(['REC', 'RECH', 'NREC', 'NRECH', 'CL', 'CCG', 'BC', 'TT', 'TRVS', 'MISC', 'NROS'])
    totalFieldCodes = [] #Creating a DataFrame for all of the field names
    totalStatCodes = [] #Creating a DataFrame for all of the stats that are in each workbook.
    table = [] #Final table that will be appended to the batch file containing the flags.
    allErrors = pd.DataFrame(columns=["SHEET", "TYPE", "PROBLEM"]) #DataFrame for the total errors for each sheet in the workbok

    


    #This loops over each sheet in the workbook
    for sheetNumber in np.arange(0, sheets):
        #Extracting the data
        currWB = wb.worksheets[sheetNumber] #Setting the current Workbook
        values = pd.DataFrame(currWB.values) #Values DataFrame
        formulas = values[values.apply(lambda string: string.astype(str).str.startswith('='))] #Formulas DataFrame

        #Finding and deleting any unwanted formulas that will mess with our algorithm.
        for col in np.arange(0, formulas.shape[1]):
            for row in np.arange(0, len(formulas[col])):
                if '$' in str(formulas[col][row]):
                    formulas[col][row] = str(formulas[col][row]).replace('$', '')
                if (str(formulas[col][row]).startswith('=INDEX')):
                    formulas[col][row] = ''
                if (str(formulas[col][row]).startswith('=\'')):
                    formulas[col][row] = ''
                if (str(formulas[col][row]).startswith('=LOOKUP')):
                    formulas[col][row] = ''
        
        #Regex for replacing negative constants in calc maps so they won't be confused with the field name statisic identification algorithm
        regexp = re.compile(r'^-\d+\.?\d*')
        for col in np.arange(0, values.shape[1]):
            for row in np.arange(0, len(values[col])):
                if (regexp.match(str(values[col][row]))):
                    values[col][row] = '6969696969'
        
        #Resetting the Index and Column Names to start from column "A" and row "1"
        values.columns = [chr(i) for i in range(ord('A'),65 + len(values.columns))]
        values.index += 1
        formulas.columns = [chr(i) for i in range(ord('A'),65 + len(formulas.columns))]
        formulas.index += 1



        #Finding all of the "=sum()" formulas and appending them to the above DataFrame
        sumObjects = [] #Dataframe for formulas that use the ex."=sum(B7:B15)" formula
        for col in formulas:
            for row in np.arange(1, len(formulas[col])):
                if ':' in str(formulas[col][row]):
                    list1 = []
                    list2 = re.findall(r'[A-Z][\d]+', str(formulas[col][row]))
                    list1.append(str(col + str(row)))
                    list1.append(list2)
                    sumObjects.append(list1)
        sumObjectsToAdd = []    #Dataframe containing each cell within the "=sum()" formula range
        for item in sumObjects: #Walking each range and appending to the above DataFrame
            newPlaces = []
            item1 = item[1]
            start = re.split('(\d+)', item1[0])[:-1]
            stop = re.split('(\d+)', item1[1])[:-1]
            for indexLook in np.arange(int(start[1]), int(stop[1])+1):
                stringValue = values[chr(ord(start[0]) - 1)][indexLook]
                if not(stringValue == None):
                    newPlaces.append(str(start[0] + str(indexLook)))
            listTemp = []
            listTemp.append(item[0])
            listTemp.append(newPlaces)
            sumObjectsToAdd.append(listTemp)
        for cell in sumObjectsToAdd:    #Rewriting te formula dataframe cells that have the "=sum()" ranges
            cellIndex = re.split('(\d+)', cell[0])[:-1]
            formulas[cellIndex[0]][int(cellIndex[1])] = str(formulas[cellIndex[0]][int(cellIndex[1])] + str(cell[1]))

        for g in np.arange(65, (values.shape[1] + 65 - 1)): #For each column
            for h in np.arange(1, values.shape[0] + 1):
                if not(values[chr(g)][h] is None):
                    if '-' in str(values[chr(g)][h]):
                        if not(str(values[chr(g)][h]).startswith("=")):
                            knownAlternatives = ['TRV', 'TRVH', 'TNRV', 'TNRVH']
                            if (values[chr(g)][h].split('-')[-1].replace(' ', '') in knownAlternatives):
                                if not(str(values[chr(g)][h]).startswith("Double -")):
                                    if not(str(values[chr(g)][h]).startswith("Alternative -")):
                                        origStat = values[chr(g)][h].split('-')[-1].replace(' ', '')
                                        if origStat == 'TRV':
                                            values[chr(g)][h] = 'Alternative - REC'
                                        if origStat == 'TRVH':
                                            values[chr(g)][h] = 'Alternative - RECH'
                                        if origStat == 'TNRV':
                                            values[chr(g)][h] = 'Alternative - NREC'
                                        if origStat == 'TNRVH':
                                            values[chr(g)][h] = 'Alternative - NRECH'
                                        tempRow = h
                                        currCell = values[chr(g)][tempRow]
                                        while (currCell != None):
                                            tempRow -= 1
                                            currCell = values[chr(g)][tempRow]
                                            if (currCell == None):
                                                values[chr(g)][tempRow] = str('Alternative - ' + origStat)
                                                # formulas[chr(g+1)][h] += str('+' + str(chr(g + 1)) + str(tempRow))
                                                if (pd.isna(formulas[chr(g+1)][h])):
                                                    formulas[chr(g+1)][h] = str('+' + str(chr(g + 1)) + str(tempRow))
                                                else:
                                                    formulas[chr(g+1)][h] += str('+' + str(chr(g + 1)) + str(tempRow))
                            if (',' in values[chr(g)][h].split('-')[-1].replace(' ', '')):
                                currentDouble = values[chr(g)][h].split('-')[-1].replace(' ', '')
                                splittedDouble = currentDouble.split(',')
                                values[chr(g)][h] = str('Double - ' + splittedDouble[0])
                                isStatistic = False
                                for item in splittedDouble:
                                    if item in stats:
                                        isStatistic = True
                                    elif item in knownAlternatives:
                                        isStatistic = True
                                if not(isStatistic):
                                    for numIndex in np.arange(1, len(splittedDouble)):
                                        tempRow = h
                                        currCell = values[chr(g)][tempRow]
                                        while (currCell != None):
                                            tempRow -= 1
                                            currCell = values[chr(g)][tempRow]
                                            if (currCell == None):
                                                values[chr(g)][tempRow] = str('Double - ' + splittedDouble[numIndex])
                                                if (pd.isna(formulas[chr(g+1)][h])):
                                                    formulas[chr(g+1)][h] = str('+' + str(chr(g + 1)) + str(tempRow))
                                                else:
                                                    formulas[chr(g+1)][h] += str('+' + str(chr(g + 1)) + str(tempRow))
                                else:
                                    actualStatistic = ''
                                    for item in splittedDouble:
                                        if item in stats:
                                            actualStatistic = item
                                        elif item in knownAlternatives:
                                            if item == 'TRV':
                                                actualStatistic = 'REC'
                                            if item == 'TRVH':
                                                actualStatistic = 'RECH'
                                            if item == 'TNRV':
                                                actualStatistic = 'NREC'
                                            if item == 'TNRVH':
                                                actualStatistic = 'NRECH'
                                    values[chr(g)][h] = str('Double - ' + actualStatistic)
                                    for item in splittedDouble:
                                        if not(item in stats):
                                            tempRow = h
                                            currCell = values[chr(g)][tempRow]
                                            while (currCell != None):
                                                tempRow -= 1
                                                currCell = values[chr(g)][tempRow]
                                                if (currCell == None):
                                                    values[chr(g)][tempRow] = str('Double - ' + item)
                                                    formulas[chr(g + 1)][h] += str('+' + str(chr(g + 1)) + str(tempRow))



                    
        #Extracting the Formula Cell Locations
        for col in formulas:    #Using regex to extract digits followed by numbers
            formulas[col] = formulas[col].apply(lambda x: np.unique(re.findall(r'[A-Z][\d]+', str(x))))
        formulasOriginal = formulas.copy() #Preserving the original formulas DataFrame

        #Seperating the column character from the row digits
        for col in formulas:
            for row in np.arange(1, formulas.shape[0] + 1):
                formulas[col][row] = np.array(pd.DataFrame(formulas[col][row])[0].apply(lambda x: re.split('(\d+)', x)[:-1]))
        
        #Creating a list of all of the unique formula cell locations
        uniqueFormulas = formulasOriginal.values.flatten()
        uniqueFormulas = [ele for ele in uniqueFormulas if ele.size > 0]
        uniqueFormulas = np.unique([x for xs in uniqueFormulas for x in xs])





        statIndices = []    #DataFrame for the indices for all of the Official Stat Excel Locations. ie. "K27"
        statNames = []      #DataFrame for the names of the Stats ie. "REC" 
        errorStats = pd.DataFrame(columns=["SHEET", "TYPE", "PROBLEM"]) #Dataframe for logging errors within each sheet.

        for g in np.arange(65, (values.shape[1] + 65 - 1)): #For each column
            for h in np.arange(1, values.shape[0] + 1): #For each row
                if (formulas[chr(g + 1)][h].size != 0): #If the value cell to the right has a formula
                    if not(values[chr(g)][h] is None):  #And If the cell isnt empty
                        if (not(str(chr(g + 1) + str(h)) in uniqueFormulas)): #And If the cell to the right is not mentioned in any of the other formulas (i.e. beggining of the tree/Stat)
                            
                            if not('-' in str(values[chr(g)][h])): #If there is no '-' in the field name
                                if values[chr(g)][h].replace(' ', '') in stats: #And If the field name is considered a statistic
                                    statNames.append(values[chr(g)][h].replace(' ', '')) #Add Field to the Stat List
                                    statIndices.append(str(chr(g + 1) + str(h))) #Add the Index to the Stat List
                                else:   #Otherwise this means that there is a stat in this worksheet that isnt named correctly, so we add it to the error log to be updated later
                                    errorStats.loc[len(errorStats)] = [sheetNumber + 1, "Formatting Error", ("Unknown Stat - " + str(values[chr(g)][h].replace(' ', '')) + '(' + str(chr(g + 1)) + str(h) + ')')] #If not a statistic, then add to error array
                                if (values[chr(g)][h].split('-')[-1].replace(' ', '') in stats): #If the field name is considered a statistic, append to the stat names
                                    statNames.append(values[chr(g)][h].split('-')[-1].replace(' ', ''))                    
                                    statIndices.append(str(chr(g + 1) + str(h)))
                                else: #Otherwise append the field name to the error log
                                    errorStats.loc[len(errorStats)] = [sheetNumber + 1, "Formatting Error", ("Unknown Stat - " + str(values[chr(g)][h].split('-')[-1].replace(' ', '')) + '(' + str(chr(g + 1)) + str(h) + ')')]
                    else: #If the cell is empty then add this to the error log
                        errorStats.loc[len(errorStats)] = [sheetNumber + 1, "Formatting Error", (str('Blank Field Name Located at ' + str(chr(g)) + str(h)))]
                
                if ('-' in str(values[chr(g)][h])): #There may be multiple stats within a sheet ie REC aand RECH
                    if ((values[chr(g)][h].split('-')[-1].replace(' ', '') in stats)):
                        if (not(values[chr(g)][h].split('-')[-1].replace(' ', '') in statNames)): #If its not already a stat, then append it.
                            statNames.append(values[chr(g)][h].split('-')[-1].replace(' ', ''))
                            statIndices.append(str(chr(g + 1) + str(h)))
            
            
        #Reccurisive Search Algorithm for traversing the excel formula trees within each sheet.
        def reccursiveSearch(arr):
            if arr.size == 0: #i.e you are at the end of the tree, just return and stop the algorithim traversal
                return
            for item in arr: #For each cell within the formula
                currName = values[chr(ord(item[0])-1)][int(item[1])] #Get the name
                if not(currName is None):
                    if currName == "PPV":
                        fieldNames.append(currName)
                    if '-' in currName:
                        if not(currName.split('-')[-1] in fieldNames):
                            fieldNames.append(currName.split('-')[-1].replace(' ', ''))
                    reccursiveSearch(formulas[str(item[0])][int(item[1])]) #Traverse down the next branch
                else:
                    errorStats.loc[len(errorStats)] = [sheetNumber + 1, "Formula Error", (str("Pointing To empty Cell" + '(' + str(item[0]) + str(item[1]) + ')'))]


        #For each stat that is within each sheet, we want to perform a reccursive search and add to the total field codes DataFrame
        for statLoc in statIndices:
            group = []
            fieldNames = []
            loc = re.split('(\d+)', statLoc)[:-1]
            if not('-' in values[chr(ord(loc[0])-1)][int(loc[1])]):
                totalStatCodes.append(values[chr(ord(loc[0])-1)][int(loc[1])].replace(' ', ''))
                group.append(values[chr(ord(loc[0])-1)][int(loc[1])].replace(' ', ''))
            else:
                totalStatCodes.append(values[chr(ord(loc[0])-1)][int(loc[1])].split('-')[-1].replace(' ', ''))
                group.append(values[chr(ord(loc[0])-1)][int(loc[1])].split('-')[-1].replace(' ', ''))
            reccursiveSearch(formulas[loc[0]][int(loc[1])])
            totalFieldCodes += fieldNames
            group.append(fieldNames)
            table.append(group)
        if (errorStats.shape[0] != 0):
            allErrors = pd.concat([allErrors, errorStats], axis=0)


    #Limit the list to only the unique field codes.
    totalFieldCodes = np.unique(totalFieldCodes)
    totalStatCodes = np.unique(totalStatCodes)
    table = pd.DataFrame(table)
    table2 = pd.DataFrame()
    table2['CODE'] = totalFieldCodes
    table2['PARK'] = name.split(' ')[0]
    for statcode in totalStatCodes:
        table2[statcode] = table2["CODE"].apply(lambda item: np.isin(item, np.array(table[table[0]==statcode][1])[0]))
    colArr = ['PARK', 'CODE', 'REC', 'RECH', 'NREC', 'NRECH', 'CL', 'CCG', 'BC', 'TT', 'TRVS', 'MISC', 'NROS']
    for item in totalStatCodes:
        if not(np.isin(item, colArr)):
            colArr.append(item)
    table2 = table2.reindex(columns=colArr)
    table2 = table2.fillna(0)
    table2.iloc[0:, 2:] = table2.iloc[0: , 2:].astype(bool)

    #Special cases with certian Calc Maps
    if(name.split(' ')[0] == "CUVA"): #CUVA SPECIAL CASE
        table2 = table2[table2.CODE != "RECH0"]
        table2 = table2.reset_index(drop=True)
    
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', 9999)
    #pd.set_option('display.max_colwidth', -1)
    pd.set_option('display.width', 9999)
    
    return(table2, allErrors)






# #FOR DEBUGGING ONLY

# pd.set_option('display.max_rows', None)
# pd.set_option('display.max_columns', None)
# #pd.set_option('display.max_colwidth', -1)
# pd.set_option('display.width', None)

parkFilesFolder = r'C:\Users\alackey\DOI\NPS-NRSS-EQD VUStats Internal - General\PARK FILES'
debugState = False
debugPark = "CUVA"

if debugState:
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    #pd.set_option('display.max_colwidth', -1)
    pd.set_option('display.width', None)

    print("DEBUGGING")
    originalPath = os.getcwd()
    print(originalPath)
    parkFilesPath = parkFilesFolder

    newDir = str(parkFilesPath + '/'.strip() + debugPark.strip())
    os.chdir(newDir)
    output = runFile(str(debugPark + str(' Calc Map.xlsx')))
    table = output[0]
    errors = output[1]
    print(table)
    print(errors)
