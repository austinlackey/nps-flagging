# Importing Packages
from autoflags import *
import os
import sys
from tqdm import tqdm
from IPython.display import display
import time

originalPath = os.getcwd()

def color_boolean(val):
    color =''
    if val == True:
        color = 'green'
    elif val == False:
        color = 'red'
    return 'color: %s' % color

def highlight_change(x, cellsThatChanged):
    color = 'background-color: white; font-weight: bolder'
    print(x)
    df1 = pd.DataFrame('', index=x.index, columns=x.columns)
    for index in cellsThatChanged:
        df1.iloc[index[1], index[0]] = color
    return df1

def find_changed_cells(correctDF, irmaDF):
    correctDF = correctDF.reset_index(drop=True)
    irmaDF = irmaDF.reset_index(drop=True)
    cellsThatChanged = []
    for col in np.arange(0, correctDF.shape[1]):
        for row in np.arange(0, correctDF.shape[0]):
            if correctDF.iloc[row, col] != irmaDF.iloc[row, col]:
                cellsThatChanged.append([col, row])
    return cellsThatChanged

# Verification Park
def verifyPark(parkName, color=True):

    # Path Variables
    parkFilesFolder = r'C:\Users\alackey\DOI\NPS-NRSS-EQD VUStats Internal - General\PARK FILES'
    os.chdir(originalPath)
    # Retreving IRMA Data
    irmaDataFrame = pd.read_csv('FlagsStatusRetrevial.csv') # Import IRMA Data
    irmaDataFrame = irmaDataFrame.sort_values(by=['Expr1']) # Sort the dataframe by field code
    irmaDataFrame = irmaDataFrame[irmaDataFrame['UnitCode'] == parkName].reset_index(drop=True) # Extract only the parks IRMA data
    irmaDataFrame.rename(columns = {'Expr1':'Code'}, inplace=True) # Rename Column
    irmaDataFrame = irmaDataFrame.drop(['Name'], axis=1) # Drop the park name column (un-needed)
    irmaDataFrame['Field'] = ['Input' if x is np.nan else 'Formula' for x in irmaDataFrame['Formula']]
    irmaDataFrame.drop(['Formula'], inplace=True, axis=1)
    # Change Directory to Verification Park
    newDir = str(parkFilesFolder + '/'.strip() + parkName.strip())
    os.chdir(newDir)

    # Retrieving Calc Map Data
    calcMapData = runFile(str(parkName + str(' Calc Map.xlsx')))[0]
    os.chdir(originalPath)
    calcMapData.rename(columns = {'PARK':'UnitCode', 'CODE':'Code'}, inplace=True) # Rename Columns

    # Data Cleaning
    irmaDataFrameCols = ['IsInSTATS', 'IsREC', 'IsRECH', 
                    'IsNREC', 'IsNRECH', 'IsCL', 
                    'IsCCG', 'IsBC', 'IsTT', 
                    'IsTRVS', 'IsMISC', 'IsNROS'] 

    calcMapCols = ['REC', 'RECH', 'NREC', 
                'NRECH', 'CL', 'CCG', 
                'BC', 'TT', 'TRVS', 
                'MISC', 'NROS']

    # Convert Boolean columns to booleans
    irmaDataFrame[irmaDataFrameCols] = irmaDataFrame[irmaDataFrameCols].astype(bool)
    calcMapData[calcMapCols] = calcMapData[calcMapCols].astype(bool)
    # Create Empty Dataframe
    correctDataFrame = irmaDataFrame.copy()
    correctDataFrame[irmaDataFrameCols] = False # Initialize Flag Column to False

    # Map Calc Map rows to the new frame
    correctDataFrame = correctDataFrame.merge(calcMapData, on='Code', how='left').drop(irmaDataFrameCols, axis=1)
    correctDataFrame = correctDataFrame.fillna(value=False).drop('UnitCode_y', axis=1) #Fill Null columns and drop duplicate reference column
    correctDataFrame.rename(columns = {'UnitCode_x':'UnitCode'}, inplace=True) # Rename reference column
    correctDataFrame.insert(3, 'INSTATS', False) # Insert INSTATS Column
    correctDataFrame['INSTATS'] = correctDataFrame[calcMapCols].any(axis='columns') # Make True for Calc Map Rows
    correctDataFrame = correctDataFrame.reset_index(drop=True)
    correctDataFrame.loc[correctDataFrame['Code'].isin(calcMapCols), 'INSTATS'] = True # If there is a STAT that is used in Stats but not referenced in a formula, (i.e REC, RECH, etc) Set InStats to True.
    irmaDataFrame.insert(4, 'Field', irmaDataFrame.pop('Field')) # Move Column
    irmaDataFrame.columns = correctDataFrame.columns # Match Column Names
    
    # Reorder Columns
    irmaDataFrame = irmaDataFrame[['UnitCode', 'Field', 'Label', 'Code', 'INSTATS', 'REC', 'RECH', 'NREC', 'NRECH', 'CL', 'CCG', 'BC', 'TT', 'TRVS', 'MISC', 'NROS']]
    correctDataFrame = correctDataFrame[['UnitCode', 'Field', 'Label', 'Code', 'INSTATS', 'REC', 'RECH', 'NREC', 'NRECH', 'CL', 'CCG', 'BC', 'TT', 'TRVS', 'MISC', 'NROS']]
    # Change sort order
    correctDataFrame = correctDataFrame.sort_values(by=['Field', 'Code'], ascending=(False, True))
    irmaDataFrame = irmaDataFrame.sort_values(by=['Field', 'Code'], ascending=(False, True))
    
    # Find the rows that are different
    correctDataFrame = correctDataFrame.reset_index(drop=True)
    irmaDataFrame = irmaDataFrame.reset_index(drop=True)
    differenceIndexes = []
    for index, row in correctDataFrame.iterrows():
        if not (correctDataFrame.iloc[index].to_numpy()==irmaDataFrame.iloc[index].to_numpy()).all():
            differenceIndexes.append(index)
    
    # Dataframe only containing the rows that need changes
    correctDataFrameChanged = correctDataFrame.iloc[differenceIndexes,:].copy().reset_index(drop=True)
    irmaDataFrameChanged = irmaDataFrame.iloc[differenceIndexes,:].copy().reset_index(drop=True)
    
    cellsAll = find_changed_cells(correctDataFrame, irmaDataFrame)
    cellsChanged = find_changed_cells(correctDataFrameChanged, irmaDataFrameChanged)
    if color:
        return(correctDataFrameChanged.style.applymap(color_boolean).applymap(color_boolean).apply(highlight_change, cellsThatChanged=cellsChanged, axis=None), correctDataFrame.style.applymap(color_boolean).apply(highlight_change, cellsThatChanged=cellsAll, axis=None))
    return(correctDataFrameChanged, correctDataFrame)