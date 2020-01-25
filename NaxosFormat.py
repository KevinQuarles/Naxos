#! python3
# NaxosFormat.py - downloads Naxos and formats the sheets of the file
import os, re, shutil, pandas as pd, datetime as dt, time, openpyxl as xl, xlrd, calendar as cal

fileDate = input('Please add the file date:')
year = fileDate[:4]
month = fileDate[5:7]
currFolder = year + '-' + month

filePath = os.path.abspath(r'\\Naxos\{year}\{folder}'.format(year = year, 
folder = currFolder))
fileList = os.listdir(filePath)


for file in fileList:
    if file.endswith('.xlsx'):
        xl = file
        
xlsx = pd.ExcelFile(os.path.join(filePath, xl))

fileName = xl.strip('xlsx')
formattedFile = pd.ExcelWriter(os.path.join(filePath, fileName + '--Formatted.xlsx'))


sheetList = xlsx.sheet_names
neededTabs = [
'AR Sales'
,'NOA-NOC Sales'
,'VS Shipments'
,'DTC Fulfillment'
,'Statement'
,'Chargebacks']
#,'C1'

for sheet in sheetList:
    if neededTabs[0] in sheet:
        arSales = pd.read_excel(xlsx, sheet)
        arSales.dropna(axis = 1, how = 'all', inplace = True )
        arSales.dropna(axis = 0, how = 'all', inplace = True )
        arSales.dropna(axis = 0, subset = ['Order'], inplace = True)
        arSales['UPC'] = arSales['UPC'].astype(str)
        arSales['UPC'] = arSales['UPC'].str.rstrip('.0')
        arSales['FileDate'] = fileDate
        arSales['Tab'] = neededTabs[0]
        arSum = arSales['Original Sales $'].sum()
        arCount = arSales['FileDate'].count()
    elif neededTabs[1] in sheet:
        noanocSales = pd.read_excel(xlsx, sheet)
        noanocSales.dropna(axis = 1, how = 'all', inplace = True )
        noanocSales.dropna(axis = 0, how = 'all', inplace = True )
        noanocSales.dropna(axis = 0, subset = ['Item'], inplace = True)
        noanocSales = noanocSales[noanocSales.UPC != 'UPC']
        noanocSales['Net $'] = noanocSales['Sales USD'] + noanocSales['Returns USD']
        noanocSales['UPC'] = noanocSales['UPC'].astype(str)
        noanocSales['UPC'] = noanocSales['UPC'].str.rstrip('.0')
        noanocSales['FileDate'] = fileDate
        noanocSales['Tab'] = 'NOANOC'
        noanocSum = noanocSales['Net $'].sum()
        noanocCount = noanocSales['FileDate'].count()
    elif neededTabs[2] in sheet:
        vsPPS = pd.read_excel(xlsx, sheet)
        vsPPS.dropna(axis = 1, how = 'all', inplace = True )
        vsPPS.dropna(axis = 0, how = 'all', inplace = True )
        vsPPS.dropna(axis = 0, subset = ['Order'], inplace = True)
        vsPPS['UPC'] = vsPPS['UPC'].astype(str)
        vsPPS['UPC'] = vsPPS['UPC'].str.rstrip('.0')
        vsPPS['FileDate'] = fileDate
        vsPPS['Tab'] = 'VS PPS'
        vsPPSSum = vsPPS['Handling Fee'].sum()
        vsPPSCount = vsPPS['FileDate'].count()
    elif neededTabs[3] in sheet:
        d2c = pd.read_excel(xlsx, sheet)
        d2c.dropna(axis = 1, how = 'all', inplace = True )
        d2c.dropna(axis = 0, how = 'all', inplace = True )
        d2c.dropna(axis = 0, subset = ['Order'], inplace = True)
        d2c['UPC'] = d2c['UPC'].astype(str)
        d2c['UPC'] = d2c['UPC'].str.rstrip('.0')
        d2c['FileDate'] = fileDate
        d2c['Tab'] = neededTabs[3]
        d2cSum = d2c['Original Sales $'].sum()
        d2cCount = d2c['FileDate'].count()
    elif neededTabs[4] in sheet:
        stmt = pd.read_excel(xlsx, sheet)
        stmt.dropna(axis = 1, how = 'all', inplace = True )
        stmt.dropna(axis = 0, how = 'all', inplace = True )
        stmt.columns = ['LineDescription', 'LineTotal']
        stmt.dropna(axis = 0, subset = ['LineDescription'], inplace = True)
        stmt.fillna(0, inplace = True)
    elif neededTabs[5] in sheet:
        chbk = pd.read_excel(xlsx, sheet)
        chbk.dropna(axis = 1, how = 'all', inplace = True )
        chbk.dropna(axis = 0, how = 'all', inplace = True )
        chbk.columns = ['LineDescription', 'LineTotal']
        chbk.dropna(axis = 0, subset = ['LineDescription'], inplace = True)
        chbk.fillna(0, inplace = True)
        

       
dateCheckAR = arSales['FileDate'].unique()
dateCheckNOANOC = noanocSales['FileDate'].unique()
dateCheckVS = vsPPS['FileDate'].unique()
dateCheckD2C = d2c['FileDate'].unique()

print('--------------------------------------------------------------------------')        
print(str(dateCheckAR) +' '+ neededTabs[0] +' has $' + str(arSum) + ' and ' + str(arCount) + ' rows!')
print(str(dateCheckNOANOC) +' '+ neededTabs[1] +' has $' + str(noanocSum) + ' and ' + str(noanocCount) + ' rows!')
print(str(dateCheckVS) +' '+ neededTabs[2] +' has $' + str(vsPPSSum) + ' and ' + str(vsPPSCount) + ' rows!')
print(str(dateCheckD2C) +' '+ neededTabs[3] +' has $' + str(d2cSum) + ' and ' + str(d2cCount) + ' rows!')
print('--------------------------------------------------------------------------')
print('--------------------------------SUMMARY-----------------------------------')
print(stmt)
print('--------------------------------CHARGEBACK--------------------------------')        
print(chbk)

print('Are the files ready for export?')
answer = input('Y or N:')

while answer != 'y' and answer != 'Y':
    print('is it ready now?')
    answer = input('Y or N:')
    
    
# put it to he formatted file
arSales.to_excel(formattedFile, sheet_name=neededTabs[0], index=False)
noanocSales.to_excel(formattedFile, sheet_name=neededTabs[1], index=False)
vsPPS.to_excel(formattedFile, sheet_name=neededTabs[2], index=False)
d2c.to_excel(formattedFile, sheet_name=neededTabs[3], index=False)
stmt.to_excel(formattedFile, sheet_name=neededTabs[4], index=False)
chbk.to_excel(formattedFile, sheet_name=neededTabs[5], index=False)

# this will save the file in the folder
formattedFile.save()

print('File has been formatted')
print('Check the totals')


#TODO
#get c1 transfer tab count and sum
#get this month and last month frt ship tabs and print status



