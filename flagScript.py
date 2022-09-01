#   flagScript script
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
#
#   This file is used for creating a batch excel file that contains the flags matrix as well as the error log. In order for this script to work,
#   the "autoflags.py" python file must be in the same folder as this script.
#
#
#   To start, Make sure the "debugSate" variable is set to "False". Then enter any Park Unit Codes that you want to exclude from the batch file
#   in the "skipParks" list. Change the variable "parkFilesFolder" to the correct path to the DOI Park Files path. Your username should be embedded
#   within the path if you did this correctly.


#   1. Make sure the "debugSate" variable is set to "False"
#   2. Change the variable "parkFilesFolder" to the correct path to the DOI Park Files path. Your username should be embedded within the path if
#   you did this correctly.
#   3. Open up Terminal(MAC) or Command Prompt(WINDOWS) and navigate to the folder/path that the python file resides in.
#       - You can do this by entering the command 'cd <folder name>' repeatily until you are in the same directory as the python file.
#   4. Enter the command 'python3 flagScript.py' and press <Enter>.
#       - If you did the above steps correctly and setup the proper DOI Server Park Files Folder path, the program should create an excel file
#       named "Batch.xlsx" in the same directory that the python script file resides in. The first sheet contains the flags matrix and the second sheet contains the error logs for all of the parks.
#   FYI: Each time you run the file, it will overwrite the old "Batch.xlsx" file, so beware.



from autoflags import *
import os
import sys
from tqdm import tqdm

# pd.set_option('display.max_rows', None)
# pd.set_option('display.max_columns', None)
# #pd.set_option('display.max_colwidth', -1)
# pd.set_option('display.width', None)



parkFilesFolder = r'C:\Users\alackey\DOI\NPS-NRSS-EQD VUStats Internal - General\PARK FILES'
skipParks = ['AMIS', 'FRSP']
debugState = False
debugPark = "MAPR"

if debugState:
    print("DEBUGGING")
    originalPath = os.getcwd()
    print(originalPath)
    parkFilesPath = parkFilesFolder
    try:
        newDir = str(parkFilesPath + '/'.strip() + debugPark.strip())
        os.chdir(newDir)
        output = runFile(str(debugPark + str(' Calc Map.xlsx')))
        table = output[0]
        errors = output[1]
        os.chdir(originalPath)
        print(table)
        print(errors)
    except:
        print("ERROR")
else:
    originalPath = os.getcwd()
    print(originalPath)

    try:
        os.chdir(parkFilesFolder)
        parkFolders = os.listdir('.')
    except OSError:
        sys.exit("Incorrect Folder Structure")
    parkFilesPath = os.getcwd()

    for park in skipParks:
        parkFolders.remove(park)
    parkFolders = [park for park in parkFolders if len (park) == 4]
    print("Skipping over", len(skipParks), "parks... ", skipParks)
    pbar = tqdm(parkFolders, unit="Parks", desc="Finiding Calc Map Files... ")
    print()
    print(len(parkFolders), "Park Folders Found")
    print()
    print(parkFolders)
    allParks = pd.DataFrame()
    errors = pd.DataFrame(columns=["PARK", "SHEET", "TYPE", "PROBLEM"])
    for folder in pbar:
        if (folder.isupper() & (len(folder) == 4)):
            newDir = str(parkFilesPath + '/'.strip() + folder.strip())
            os.chdir(newDir)
            filesList = os.listdir('.')
            if not(str(folder + str(' Calc Map.xlsx')) in filesList):
                errors.loc[len(errors)] = [folder, 0, "File Error", "No file named (XXXX Calc Map.xlsx)"]
            for file in filesList:
                if (file == str(folder + str(' Calc Map.xlsx'))):
                    pbar.set_description(f'Proccessing %s' % folder)
                    try:
                        a = runFile(file)
                        a0 = a[0]
                        a1 = a[1]
                        allParks = pd.concat([allParks, a0])
                        if (len(a0['CODE']) != len(np.unique(a0['CODE']))):
                            errors.loc[len(errors)] = [folder, 0, "PythonCalcError", str(len(a0['CODE']), len(np.unique(a0['CODE'])))]
                        if (len(a1) != 0):
                            for rowE in np.arange(0, len(a1)):
                                errors.loc[len(errors)] = [folder] + list(a1.iloc[rowE])
                    except Exception as e:
                        errors.loc[len(errors)] = [folder, 0, "Python Error", str(e)]
            os.chdir(parkFilesPath)
    for park in skipParks:
        errors.loc[len(errors)] = [park, 0, "Calc Map Error", "Skipped Park"]
    os.chdir(originalPath)
    allParks = allParks.fillna(0)
    allParks.iloc[0:, 2:] = allParks.iloc[0: , 2:].astype(bool)
    fillTrue = np.ones(len(allParks), dtype=bool)
    allParks.insert(loc=2, column="Used in Official Stats", value=fillTrue)
    allParks.rename(columns = {'CODE':'Field Name'}, inplace = True)
    with pd.ExcelWriter('Batch.xlsx') as writer:  
        allParks.to_excel(writer, sheet_name='BATCH TABLE')
        errors.to_excel(writer, sheet_name='ERRORS')
    print()
    print("Batch is complete, look for the file \'Batch.xlsx\'")