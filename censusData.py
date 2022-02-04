# Read census file...

import openpyxl, pprint

print('Openin Workbook..........')
wb = openpyxl.load_workbook(r'C:\Users\Georgee\Documents\automate_online-materials\censuspopdata.xlsx')
sheet = wb['Population by Census Tract']
CountyData = {}

# Fill in ccontyData with each county's population and tracts
print('Reading Rows..........')
for row in range(2, sheet.max_row + 1):    

    # Each row in spreadsheet has data for one census tract 
    state = sheet['B' + str(row)].value
    county = sheet['C' + str(row)].value
    pop = sheet['D' + str(row)].value
    
    # Make sure the key for this state exist
    CountyData.setdefault(state, {})

    # Make sure the key for this county in this state exist
    CountyData[state].setdefault(county, {'tracts': 0, 'pop': 0})

    # Each row represents one census tracts, increments by one
    CountyData[state][county]['tracts'] += 1

    # Increase the county pop by the pop in this census tract
    CountyData[state][county]['pop'] += int(pop)

# Open a new text file and wirte the content of the countyData to it
print('Writing results..........')
resultFile = open('census2010.py', 'w')
resultFile.write('allData = ' + pprint.pformat(CountyData))
resultFile.close()
print('Done!')