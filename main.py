

try:
    import urllib.request as urllib2
except ImportError:
    import urllib2

import json
import xlrd
import xlwt 
import openpyxl
import os
import pyperclip, re


json_obj = urllib2.urlopen('https://data.kcmo.org/resource/5sem-frgw.json')
data = json.load(json_obj)
print('Crimes in zip code 64130')
for item in data:
    if item['zip_code_1'] == '64130':
        print('')
        print ('Crime type: '+item['description'])
        print('Date: ' + item['reported_date'] )
        if item['race_1'] == 'B':
            race='Black'
        elif item['race_1'] == 'W':
            race='White'
        else: race='Unknown'
        print('Criminal race: ' + race)
        
        
        
        
print(" \n" *5)
print("============================")
print("============================")

flag1=1                       #Loop for the columns
for item in data:
    x7 = item['zip_code_1']
    if x7 == 'email' or x7 == 'Email' or x7 == 'EMAIL':
        item['zip_code_1']
        bad_column=j
        print('Email column detected, located in the column #' + str(bad_column) + ' ==> ' + openpyxl.utils.get_column_letter(bad_column))
                for k in range(2, rn+1):                  #Loop for cleaning
                    sh_data.cell(row=k, column=bad_column).value=''
                    d1.save('pii1_email_clear.xlsx')
                    print('Cleaning the emails is done and saved !')
print(x7)
if flag1 == 1:
    print('No email title detected in this sheet ...')
    print('Starting deep searching for email PII depending on the pattern...')
    # email regex:
    emailRegex = re.compile(r'''(
        [a-zA-Z0-9._%+-]+ # username
        @ # @ symbol
        [a-zA-Z0-9.-]+ # domain name
        (\.[a-zA-Z]{2,4}) # dot-something
        )''', re.VERBOSE)
    matches = []
    suspecious_col=[]
    print ('\n'*5)
    d1 = openpyxl.load_workbook('pii1.xlsx')
    print(type(d1))
    #scanning the excel file contents
    shsarray=d1.get_sheet_names()
    shn=0;
    for sheet in d1.worksheets:
        shn = shn + 1
    print("This excel file contains " + str(shn) + " sheets")
    for i in range(0, len(shsarray)): print('Sheet #' + str(i+1) + ' is: ' + str(shsarray[i]))
    for i in range(0, len(shsarray)): print(shsarray[i])
    for i in range (0, shn):                              #Loop for the sheets
        print('For the sheet #' + str(i+1) + ':')
        sh_data=d1.get_sheet_by_name(shsarray[i])
        rn=sh_data.max_row
        cn=sh_data.max_column
        print('Number of rows is:' + str(rn))             
        print('Number of rows is:' + str(cn))
        noTestRecords = 10
        threshPIISus = 0.7
        noThreshPIISus = noTestRecords * threshPIISus
        for j in range (1, cn+1):                         #Loop for the columns
            for jj in range (1, 4):                       #Loop for the first four rows 
                x7=sh_data.cell(row=jj, column=j).value
                for groups in emailRegex.findall(str(x7)):         #The condition
                    matches.append(groups[0])
                    
                if len(matches) > 0:
                    bad_column=j
                    suspecious_col.append(bad_column)
                    #suspecious_col[j]=bad_column
                    bad_row=jj                              #is there a possibility for horizental data pattern? 
                    print('Email pattern found in column #' + str(bad_column) + ' ==> ' + openpyxl.utils.get_column_letter(bad_column) + ', and row # ' + str(bad_row))
                
            #testing if within the first four rows there is a suspicious column
            if len(suspecious_col) >= noThreshPIISus:  #75% case
                print(' More than len(suspecious_col) the cells of the column #' + str(bad_column) + 'contain email addresses, this column must be filtered ...' )
                print('Cleaning in progress ... ')
                for k in range(1, rn+1):                  #Loop for cleaning
                    sh_data.cell(row=k, column=bad_column).value=''
                    d1.save('pii1_email_clear.xlsx')
                print('Cleaning the emails is done and saved !')
                    
                    
            if len(suspecious_col) < 3: #75% case
                print(' 50% or less of the cells of the column #' + str(bad_column) + 'contain email addresses, operator must decide ...' )
                print('please hit c for cleaning, d for start deep search, or s to skip')
                controlv = input().lower()

