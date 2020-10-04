import glob  # Used to combine .CSV files together
import os
import sqlite3  # Used to create a DB of all the data
import openpyxl  # Library for manipulating data in Excel
import pandas as pd  # Library for manipulating data in Excel
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill  # , Font, Border

TEAM = 2122  # Place your Team number here

os.chdir('M:/EagleScout')  # (' ./EagleScout) # Set for your system
path = '.'

# ------------------------------- Perform file Clean up -------------------------------

# Remove the combined .csv file and all spreadsheets from previous runs
# This will provide clean data without corruption from earlier runs

if os.path.exists('./Spreadsheets/combined.csv'):
    os.remove('./Spreadsheets/combined.csv')

# Remove previous copy of the Database version
if os.path.exists('./DataBases/Combined_Raw.db'):
    os.remove('./DataBases/Combined_Raw.db')

# Remove previous Excel spreadsheet
if os.path.exists('./Spreadsheets/Tournament.xlsx'):
    os.remove('./Spreadsheets/Tournament.xlsx')

# Remove sorted combined spreadsheet
if os.path.exists('./Spreadsheets/Combined.xlsx'):
    os.remove('./Spreadsheets/Combined.xlsx')

# Remove sorted combined spreadsheet
if os.path.exists('./Spreadsheets/ScoreContrib.xlsx'):
    os.remove('./Spreadsheets/ScoreContrib.xlsx')

# Remove previous Excel spreadsheet
if os.path.exists('./Spreadsheets/Master.xlsx'):
    os.remove('./Spreadsheets/Master.xlsx')

# Remove previous Excel spreadsheet
if os.path.exists('./Spreadsheets/Temp.xlsx'):
    os.remove('./Spreadsheets/Temp.xlsx')

# Remove old copy of Match Schedule .csv
if os.path.exists('./Spreadsheets/Match_Schdule.csv'):
    os.remove('./Spreadsheets/Match_Schdule.csv')
# ----------------------------- End of file Clean up ------------------------------------
# ---------------------------------------------------------------------------------------

# Create the Database file
con = sqlite3.connect('./DataBases/Combined_Raw.db')

# ------------------------------Conditional Formatting values----------------------------

# Create fill patterns for colored text
redFill = PatternFill(start_color='F95555', end_color='F95555', fill_type='solid')
yellowFill = PatternFill(start_color='FBFE46', end_color='FBFE46', fill_type='solid')
greenFill = PatternFill(start_color='ACF99D', end_color='ACF99D', fill_type='solid')
clearFill = PatternFill(start_color='000000', end_color='000000', fill_type='solid')

######################################################################################
# --------------------------------Start Building the desired files----------------------


# Read in and merge all .CSV file names
files_in_dir = [f for f in glob.glob('*.csv')]

# Edit these entries for each column of data collected. These are column labels used when placing the
# scouted data onto the individual Team worksheets. See line 127 . These must match the content of './Config_Files/Header.txt'
# from line 79.
StrToInt_dict = {'Team': int, 'Match_Num': int, 'Auto_Cross': int, 'Auto_Outter': int, 'Auto_Bottom': int,
                 'Tele Outter': int, 'Tele_Bottom': int, 'Rotation': int, 'Position': int, 'Climb': int, 'Level': int,
                 'Driver_Perf': int, 'Auto_Perf': int, 'Name': str, 'Comments': str, 'Point_Contrib': int}

# Create a single combined .csv file with all data from all matches completed so far,then
# and add column headers as labels. Don't forget to edit the Headers.txt file to match!  PANDAS IS USED HERE
d1 = pd.read_csv('./Config_Files/Header.txt')  # Read in the text file content
d1.to_csv('./Spreadsheets/combined.csv', header=True, index=False)  # Write the text file content as headers

fout = open('./Spreadsheets/combined.csv', "a")  # Open the combined.csv file

for filenames in files_in_dir:
    # df = pd.read_csv(filenames)
    fName, fExt = (os.path.splitext(filenames))
    sName = fName.split('-')
    N = (sName[1])  # Match #
    T = (sName[0])  # Team #
    # df.insert(0,N,N,True)
    # df.to_csv('./Spreadsheets/combined.csv', index_label = (sName[0]), mode = 'a')

    for line in open(filenames):
        fout.write(str(T) + ",")
        fout.write(str(N) + ",")
        fout.write(line)
fout.close()  # Close out the combined.csv file

# ------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------

# Convert combined Raw csv file into one Master Raw Excel Data file
# Then add score contribution to each entry. THIS USES PANDAS!!!
with pd.ExcelWriter('./Spreadsheets/Combined.xlsx', engine='xlsxwriter') as writer: # Configure Pandas writer and set new file name,
    dt = pd.read_csv('./Spreadsheets/combined.csv')  #Read into the data table the existing .csv file.
    rows = len(dt.index)  #Count how many rows exist

    dt.to_excel(writer, sheet_name='All data', index=False) #Write the data table to sheet with the name "All data"
    worksheet = writer.sheets['All data']
    writer.save() # Save the worksheet

#  CUSTOMIZE FOR EACH YEAR------------------------------------------------
# --------------------Add custom calculated data--------------------------
# --------------------Edit each year per Game Scoring---------------------
# In 2020 this is used for Score Contribution-----------------------------
#  USING OPENPYXL
ScrCont = openpyxl.load_workbook('./Spreadsheets/Combined.xlsx', read_only=False, keep_vba=True)  # Load in the Combined.xlsx workbook
SN = ScrCont.get_sheet_by_name('All data')
ROWS = SN.max_row #Establish row count

for i in range(2, ROWS + 1 ):  # Add Score contribution
    SC = SN.cell(row = i, column = 3).value * 5  # 5 points for moving
    SC += SN.cell(row = i, column = 4).value * 4 # 4 points per Upper goal (Auto)
    SC += SN.cell(row = i, column = 5).value * 2 # 2 points per Lower goal (Auto)
    SC += SN.cell(row = i, column = 6).value * 2 # 2 points per Upper goal (Tele)
    SC += SN.cell(row = i, column = 7).value * 1 # 1 point per Lower goal (Tele)
    SC += SN.cell(row = i, column = 8).value * 10 # 10 points for Rotation
    SC += SN.cell(row = i, column = 9).value * 20 # 20 points for Position
    if SN.cell(row = i, column = 10).value == 0:
        val = 0 # No score
    elif SN.cell(row = i, column = 10).value == 1:
        val = 5 # 5 points for Parked
    elif SN.cell(row = i, column = 10).value == 2:
        val = 25 # 25 points for Hang
    elif SN.cell(row = i, column = 10).value == 3:
        val = 50 # 50 points for double Hang
    elif SN.cell(row = i, column = 10).value == 4:
        val = 75 # 75 points for tripple Hang

    SC += val
    SC += SN.cell(row = i, column = 11).value * 15 # 15 points for hang in Balance
    SN.cell(row = i , column = 16).value = SC

ScrCont.save('./Spreadsheets/ScoreContrib.xlsx') #Save Score Contribution spreadsheet


# ------------------------------------Save Raw data to Database file-----------------------
# Save Combined.xlsx to a Database file
db = pd.read_excel('./Spreadsheets/ScoreContrib.xlsx')
db.to_sql("Raw_Data", con, if_exists='replace', index=False)
#---------End of Database operations--------------------------------------------------------


# Parse through Combined.xlsx files and append content to appropriate team worksheet.
# Read in the file and set the values to 'int'
with pd.ExcelWriter('./Spreadsheets/Master.xlsx') as writer:
    df2 = pd.read_excel('./Spreadsheets/ScoreContrib.xlsx', converters=StrToInt_dict)
    # df2 = pd.read_excel('./Spreadsheets/Combined.xlsx')
    group = df2.groupby('Team')
    for Team, Team_df in group:
        Team_df.to_excel(writer, sheet_name=("T" + str(Team)), index=False)

    writer.save()

# ----------------------------Add formulas and manipulations ------------------------------


# Add formulas to each sheet for calculating values
Data = pd.ExcelFile('./Spreadsheets/Master.xlsx')
Teams = Data.sheet_names

Tnmt = openpyxl.load_workbook('./Spreadsheets/Master.xlsx', read_only=False, keep_vba=True)

# Add new worksheet for easy access to pertinant info
WS1 = Tnmt.create_sheet("Important Stuff", 0)
WS2 = Tnmt.create_sheet("Predictions")

# Add the same formulas to each team's sheet. Edit this section each year to match the game
for sht in Teams:
    sn = Tnmt.get_sheet_by_name(sht)
    sn['B16'] = str('Average')
    sn['B17'] = str('Standard Deviation')
    sn['B18'] = str('STDev % of Average')
    sn['B19'] = str('Total')
    sn['E21'] = str('Upper Ave')
    sn['E22'] = str('Lower Ave')
    sn['E23'] = str('Rotation Ave')
    sn['E24'] = str('Position Ave')
    sn['E25'] = str('Climb Ave')
    sn['E26'] = str('Points Ave')

    sn.cell(row=16, column=3).value = "=AVERAGE(C2:C13)"
    sn.cell(row=16, column=4).value = "=AVERAGE(D2:D13)"
    sn.cell(row=16, column=5).value = "=AVERAGE(E2:E13)"
    sn.cell(row=16, column=6).value = "=AVERAGE(F2:F13)"
    sn.cell(row=16, column=7).value = "=AVERAGE(G2:G13)"
    sn.cell(row=16, column=8).value = "=AVERAGE(H2:H13)"
    sn.cell(row=16, column=9).value = "=AVERAGE(I2:I13)"
    sn.cell(row=16, column=10).value = "=AVERAGE(J2:J13)"
    sn.cell(row=16, column=11).value = "=AVERAGE(K2:K13)"
    sn.cell(row=16, column=16).value = "=AVERAGE(P2:P13)"

    sn.cell(row=17, column=3).value = "=STDEV(C2:C13)"
    sn.cell(row=17, column=4).value = "=STDEV(D2:D13)"
    sn.cell(row=17, column=5).value = "=STDEV(E2:E13)"
    sn.cell(row=17, column=6).value = "=STDEV(F2:F13)"
    sn.cell(row=17, column=7).value = "=STDEV(G2:G13)"
    sn.cell(row=17, column=8).value = "=STDEV(H2:H13)"
    sn.cell(row=17, column=9).value = "=STDEV(I2:I13)"
    sn.cell(row=17, column=10).value = "=STDEV(J2:J13)"
    sn.cell(row=17, column=11).value = "=STDEV(K2:K13)"
    sn.cell(row=17, column=16).value = "=STDEV(P2:P13)"

    sn.cell(row=18, column=3).value = "=SUM(C17/C16)"
    sn.cell(row=18, column=4).value = "=SUM(D17/D16)"
    sn.cell(row=18, column=5).value = "=SUM(E17/E16)"
    sn.cell(row=18, column=6).value = "=SUM(F17/F16)"
    sn.cell(row=18, column=7).value = "=SUM(G17/G16)"
    sn.cell(row=18, column=8).value = "=SUM(H17/H16)"
    sn.cell(row=18, column=9).value = "=SUM(I17/I16)"
    sn.cell(row=18, column=10).value = "=SUM(J17/J16)"
    sn.cell(row=18, column=11).value = "=SUM(K17/K16)"
    sn.cell(row=18, column=16).value = "=SUM(P17/P16)"

    sn.cell(row=19, column=3).value = "=SUM(C2:C13)"
    sn.cell(row=19, column=4).value = "=SUM(D2:D13)"
    sn.cell(row=19, column=5).value = "=SUM(E2:E13)"
    sn.cell(row=19, column=6).value = "=SUM(F2:F13)"
    sn.cell(row=19, column=7).value = "=SUM(G2:G13)"
    sn.cell(row=19, column=8).value = "=SUM(H2:H13)"
    sn.cell(row=19, column=9).value = "=SUM(I2:I13)"
    sn.cell(row=19, column=10).value = "=SUM(J2:J13)"
    sn.cell(row=19, column=11).value = "=SUM(K2:K13)"
    sn.cell(row=19, column=16).value = "=SUM(P2:P13)"

    sn.cell(row=21, column=6).value = "=SUM(D16+F16)"
    sn.cell(row=22, column=6).value = "=SUM(E16+G16)"
    sn.cell(row=23, column=6).value = "=SUM(H16*1)"
    sn.cell(row=24, column=6).value = "=SUM(I16*1)"
    sn.cell(row=25, column=6).value = "=SUM(J16*1)"
    sn.cell(row=26, column=6).value = "=SUM(P16*1)"

C = 0
for sht in Teams:
    WS1.cell(row=2 + C, column=1).value = (sht[1:])
    C += 1  # Add Header info to "Important Stuff" sheet. Edit each year to match the game.
Stuff = ['Team #', 'Average Upper',  'Average Lower','Rotation', 'Position', 'Climb','Av Point Contrib']
s = 0
for st in Stuff:
    WS1.cell(row=1, column=1 + s).value = str(Stuff[s])
    s += 1

Tnmt.save('./Spreadsheets/Temp.xlsx')

# Copy data from sheets and cells to where it's needed. Edit each year to match the game.
Trnmt = openpyxl.load_workbook('./Spreadsheets/Temp.xlsx', read_only=False, keep_vba=True, data_only=False)

SN = Trnmt.get_sheet_by_name('Important Stuff')

D = 0
for tn in Teams:
    TT = "=" + tn + "!" + "F21"
    TU = "=" + tn + "!" + "F22"
    TV = "=" + tn + "!" + "F23"
    TW = "=" + tn + "!" + "F24"
    TX = "=" + tn + "!" + "F25"
    TY = "=" + tn + "!" + "P16"
    SN.cell(row=2 + D, column=2).value = TT  # Average Upper 
    SN.cell(row=2 + D, column=3).value = TU  # Average Lower
    SN.cell(row=2 + D, column=4).value = TV  # Average Rotation
    SN.cell(row=2 + D, column=5).value = TW  # Average Position
    SN.cell(row=2 + D, column=6).value = TX  # Average Climb
    SN.cell(row=2 + D, column=7).value = TY  # Average Point Contribution
    SN.cell(row=2 + D, column=2).number_format = '0.00'
    SN.cell(row=2 + D, column=3).number_format = '0.00'
    SN.cell(row=2 + D, column=4).number_format = '0.00'
    SN.cell(row=2 + D, column=5).number_format = '0.00'
    SN.cell(row=2 + D, column=6).number_format = '0.00'
    SN.cell(row=2 + D, column=7).number_format = '0.00'

    D += 1

SN.conditional_formatting.add('B2:B75',
                              CellIsRule(operator='greaterThan', formula=[20.1], stopIfTrue=True, fill=greenFill))
SN.conditional_formatting.add('B2:B75',
                              CellIsRule(operator='between', formula=[10.1,20.0], stopIfTrue=True, fill=yellowFill))
SN.conditional_formatting.add('B2:B75',
                              CellIsRule(operator='between', formula=[.01,10], stopIfTrue=True, fill=redFill))
SN.conditional_formatting.add('C2:C75',
                              CellIsRule(operator='greaterThan', formula=[10.00], stopIfTrue=True, fill=greenFill))
SN.conditional_formatting.add('C2:C75',
                              CellIsRule(operator='between', formula=[5.1, 9.99], stopIfTrue=True, fill=yellowFill))
SN.conditional_formatting.add('C2:C75',
                              CellIsRule(operator='between', formula=[.01, 5], stopIfTrue=True, fill=redFill))
SN.conditional_formatting.add('D2:D75',
                              CellIsRule(operator='greaterThan', formula=[.75], stopIfTrue=True, fill=greenFill))
SN.conditional_formatting.add('D2:D75',
                              CellIsRule(operator='between', formula=[.5, .74], stopIfTrue=True, fill=yellowFill))
SN.conditional_formatting.add('D2:D75',
                              CellIsRule(operator='between', formula=[.01, .49], stopIfTrue=True, fill=redFill))
SN.conditional_formatting.add('E2:E75',
                              CellIsRule(operator='greaterThan', formula=[.75], stopIfTrue=True, fill=greenFill))
SN.conditional_formatting.add('E2:E75',
                              CellIsRule(operator='between', formula=[.5, .74], stopIfTrue=True, fill=yellowFill))
SN.conditional_formatting.add('E2:E75',
                              CellIsRule(operator='between', formula=[.01, .49], stopIfTrue=True, fill=redFill))
SN.conditional_formatting.add('F2:F75',
                              CellIsRule(operator='greaterThan', formula=[.75], stopIfTrue=True, fill=greenFill))
SN.conditional_formatting.add('F2:F75',
                              CellIsRule(operator='between', formula=[.5, .74], stopIfTrue=True, fill=yellowFill))
SN.conditional_formatting.add('F2:F75',
                              CellIsRule(operator='between', formula=[.01, .49], stopIfTrue=True, fill=redFill))
SN.conditional_formatting.add('G2:G75',
                              CellIsRule(operator='greaterThan', formula=[120], stopIfTrue=True, fill=greenFill))
SN.conditional_formatting.add('G2:G75',
                              CellIsRule(operator='between', formula=[75, 119.9], stopIfTrue=True, fill=yellowFill))
SN.conditional_formatting.add('G2:G75',
                              CellIsRule(operator='between', formula=[.01, 74.9], stopIfTrue=True, fill=redFill))
Trnmt.save('./Spreadsheets/Tournament.xlsx')

# -----------------------Get data from other Excel files and add to Tournament.xls------------

book = openpyxl.load_workbook('./Spreadsheets/Tournament.xlsx')
writer = pd.ExcelWriter('./Spreadsheets/Tournament.xlsx', engine='openpyxl')
writer.book = book

df1 = pd.read_excel('/EagleScout/Python Scripts/Match_Schedule/Match_Schdule.xlsx', sheet_name='Match Schedule',
                    header=0)
df2 = pd.read_excel('/EagleScout/Python Scripts/Match_Schedule/Match_Schdule.xlsx', sheet_name= str(TEAM) + ' Schedule',
                    header=0)
df3 = pd.read_excel('./Spreadsheets/Tournament.xlsx', sheet_name='Predictions')
df1.to_excel(writer, sheet_name='Match Schedule', index=False)
df2.to_excel(writer, sheet_name= str(TEAM) + ' Schedule', index=False)

prd = book.get_sheet_by_name('Predictions')
prd.cell(row=1, column=1).value = 'Match #'
prd.cell(row=1, column=2).value = 'Red'
prd.cell(row=1, column=3).value = 'Blue'

# ---------------------- Create Prdictions for "our" matches as defined in line 318 -------------------
rows = len(df2.index)
for r in range(0, rows):
    m = df2.at[r, 'Match #']
    prd.cell(row=r + 2, column=1).value = m

# -------- Add score contributions values for each team on an alliance ------
for m in range(0, rows):
    R1 = "=T" + str(df2.at[0 + int(m), 'Red 1']) + "!P16"
    R2 = "=T" + str(df2.at[0 + int(m), 'Red 2']) + "!P16"
    R3 = "=T" + str(df2.at[0 + int(m), 'Red 3']) + "!P16"
    prd.cell(row=100 + (m * 3), column=5).value = R1
    prd.cell(row=101 + (m * 3), column=5).value = R2
    prd.cell(row=102 + (m * 3), column=5).value = R3
    prd.cell(row=2 + m, column=2).value = "=SUM(E" + str(100 + (m * 3)) + ":E" + str(102 + (m * 3))
    B1 = "=T" + str(df2.at[0 + int(m), 'Blue 1']) + "!P16"
    B2 = "=T" + str(df2.at[0 + int(m), 'Blue 2']) + "!P16"
    B3 = "=T" + str(df2.at[0 + int(m), 'Blue 3']) + "!P16"
    prd.cell(row=100 + (m * 3), column=6).value = B1
    prd.cell(row=101 + (m * 3), column=6).value = B2
    prd.cell(row=102 + (m * 3), column=6).value = B3
    prd.cell(row=2 + m, column=3).value = "=SUM(F" + str(100 + (m * 3)) + ":F" + str(102 + (m * 3))

writer.save()
writer.close()

# ----------------------------Create a Match Schedule.csv for use by Scouting Tablets-------------------------
MS = pd.read_excel('/EagleScout/Python Scripts/Match_Schedule/Match_Schdule.xlsx', sheet_name='Match Schedule')
MS.to_csv('./Spreadsheets/Match_Schdule.csv',
          columns=('Match #', 'Time', 'Red 1', 'Red 2', 'Red 3', 'Blue 1', 'Blue 2', 'Blue 3'), index=True)
