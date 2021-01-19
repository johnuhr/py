## =============================================================
##    2020-10-08 Export casedata from Visma to HTML document(s)
##    John Uhrskov
##
##  This script was run against production folders 21.dec.2020 14.58
##
##
##
##
##
## =============================================================

import datetime
import io
import os
from os import path

import shutil

now = datetime.datetime.now()
isoDate = now.strftime('%Y%m%d')

server= 'SVRPRDVISSQL01'
db = 'F2002'

writeHtmlOverview=True
writeToExcel=True

currentDirectory = os.getcwd()

destinationDirectory = r'M:\Documents' 
##destinationDirectory = r'C:\Temp\Visma\Documents'

csv_file = f'vismaCaseDocFiles_{db}.csv'
csv_files_not_found = f'vismaCaseDocFilesNotFound_{db}.csv'


# =============================================
#  Read CASE OVERVIEW data from the SQL Server
# =============================================
import pandas as pd
import pyodbc

conn = pyodbc.connect('Driver={SQL Server};Server='+server+';Database='+db+';Tructed_Connection=yes;')

sql = """
SELECT --TOP (1000)
    [Case No.], [Description], [Client Name], [Nick Name], [Employee], [Case Status], [Case Type], [Actor no 1 Name] AS "Actor1", [Actor no 2 Name] AS "Actor2"
FROM [dbo].[all_Cases]
ORDER BY [Case No.] ASC
"""

if writeHtmlOverview==True:
    print("read_sql (all_Cases)")
    df = pd.read_sql(sql, conn)

# ==========================
#  Write data to Excel file
# ==========================

if writeToExcel:
    
    sheetName = now.strftime(isoDate)
    fileName = now.strftime('Visma'+db+' (%Y%m%d).xlsx')

    if os.path.exists(fileName):
        print ('removing existing Excel file')
        os.remove(fileName)

    writer = pd.ExcelWriter(fileName, engine='xlsxwriter', date_format='yyyy-MM-dd', datetime_format='yyyy-MM-dd' )

    print("Creating Excel dataset")
    df.to_excel(writer  , sheet_name = sheetName, float_format = '%.2f' , index = False , engine='xlsxwriter' )

    worksheet = writer.sheets[sheetName]
    workbook = writer.book    

    print("col_names")
    col_names = [{'header': col_name} for col_name in df.columns]

    print("add_table")
    worksheet.add_table(0, 0, df.shape[0], df.shape[1]-1, {
        'columns': col_names,
        # 'style' = option Format as table value and is case sensitive 
        # (look at the exact name into Excel)
        'style': 'Table Style Light 1'
    })

    print ('Automatically adjust width of columns in Excel file')
    for i, col in enumerate(df.columns):
        # find length of column i
        column_len = df[col].astype(str).str.len().max()
        # Setting the length if the column header is larger
        # than the max column value length
        column_len = max(column_len, len(col)) + 2
        # set the column length
        worksheet.set_column(i, i, column_len)

    # save writer object and created Excel file with data from DataFrame
    print('Excel save begin')
    writer.save()
    print('Excel save done')

    print ("moving Excel file to destination directory")
    shutil.move(os.path.join(currentDirectory, fileName), os.path.join(destinationDirectory, fileName))

## =============================
##    Write data to HTML file
## =============================

htmlheader = """<html>
<head>
<style type="text/css">
body {
  color:#003d66;
}
table {
  width:100%;
  border-collapse:collapse;
  border-color:#5b9bd5;
  border-style:solid;
  border-width:1px;
}
table th {
  background:#5B9BD5;
  text-align: left;
    position: sticky;
    top:0;
}
table tr:nth-child(odd) {
  background:#deeaf6;
}
table tr:nth-child(even) {
}
table td {
  border-color:#5b9bd5;
  border-width:1px;
  border-style:solid;
  color:#003d66;
  padding-left:6px;
}
table tr:hover {
  background-color:#ffe699;
}

</style>
</head>
<body>
"""

bodyheader = '<h2> Company '+db+' <br><font size="2">Data exported from Visma Finance '+isoDate+'</font> </h2>'
htmlfooter='<body></thml>'
htmlfilename=destinationDirectory+'\Visma'+db+'.html'

if writeHtmlOverview==True:
    with open(htmlfilename, "w") as file:
        file.write(htmlheader)

    with open(htmlfilename, "a") as file:
        file.write(bodyheader)

    with open(htmlfilename, 'a') as table:
        df.to_html(table)
        
    with open(htmlfilename, "a") as file:
        file.write(htmlfooter)



## =====================================================
##  Read data for each case and write data to html file
## =====================================================

#Sticky table header removed!
htmlheader = """<html>
<head>
<style type="text/css">
body {
  color:#003d66;
}
table {
  width:100%;
  border-collapse:collapse;
  border-color:#5b9bd5;
  border-style:solid;
  border-width:1px;
  empty-cells:show;
}
table th {
  background:#5B9BD5;
  text-align: left;
}
table tr:nth-child(odd) {
  background:#deeaf6;
}
table tr:nth-child(even) {
}
table td {
  border-color:#5b9bd5;
  border-width:1px;
  border-style:solid;
  color: #003d66;
  padding-left:6px;
  empty-cells:show;
}
table tr:hover {
  background-color: #ffe699;
}

</style>
</head>
<body>
"""

sql = """
SELECT 
[Case no.]
,[Case Status]
,[Case Type]
,[Created Date]
,[Actual End Date]
,[Description]
,[Nick Name]
,[Employee]
,[Assistant]
,[Client No]
,[Client Name]
,[Client Address line 1]
,[Client Address line 2]
,[Client Post Code]
,[Client Post Area]
,[Client Country]
,[Client Country Name]
,[Clients ref no]
,[Conflict checked by/date]
,[Old case no]
,[First Estimate]
,[Actual Estimate]

,[Actor no 1 Type]
,[Actor no 1]
,[Actor no 1 Name]
,[Actor no 1 ref]
,[Actor no 1 Description]

,[Actor no 2 Type]
,[Actor no 2]
,[Actor no 2 Name]
,[Actor no 2 ref]
,[Actor no 2 Description]

,[Actor no 3 Type]
,[Actor no 3]
,[Actor no 3 Name]
,[Actor no 3 ref]
,[Actor no 3 Description]

,[Actor no 4 Type]
,[Actor no 4]
,[Actor no 4 Name]
,[Actor no 4 ref]
,[Actor no 4 Description]
FROM [dbo].[all_Cases]
--WHERE [Case no.] = 2020144
"""

print("read_sql (All_Cases)")
vismaCaseData = pd.read_sql(sql, conn)

sql = """SELECT 
[Case no.],[Type],[Actor no.],[Actor name],[Actors ref.],[Description]
FROM [dbo].[all_Cases_addActors]
--WHERE [Case no.] = 2020144
"""
print("read_sql (all_Cases_addActors)")
vismaCaseAddActors = pd.read_sql(sql, conn)

sql = """SELECT 
[Case no.],[User name],[Due date],[Type],[Responsible],[Description],[Created date],[Changed date],[Resolved date]
FROM [dbo].[all_Cases_DueDates]
--WHERE [Case no.] = 2020144
ORDER BY [Case no.],[Due date] DESC,[Created date]
"""
print("read_sql (all_Cases_DueDates)")
vismaCaseDuedates = pd.read_sql(sql, conn)


sql = """SELECT 
[Doc no.],[Case no.],[Version no.],[Associate no.],[Actor],[Version Create date],[Original Create date],[Doc.gr],[Incomming],[Description],[Contact Name],[Created by Usr],[File path]
--,CAST('No' AS varchar(3)) AS "fileFound"
FROM [dbo].[all_Cases_Documents]
--WHERE [Case no.] = 2020144
ORDER BY [Doc no.] DESC ,[Version no.] DESC
"""
print("read_sql (all_Cases_Documents)")
vismaCaseDocuments = pd.read_sql(sql, conn)


if os.path.isfile(csv_file)==True:
    print("csv file found. file search result will be read from csv file.")
    df = pd.read_csv(csv_file)
    vismaCaseDocuments=vismaCaseDocuments.merge(df, how='inner', on='Doc no.')
    
if os.path.isfile(csv_file)==False:
    print("csv file not found. search for all documents started...")
    doc_count = len(vismaCaseDocuments)
    vismaCaseDocuments['fileFound']='No'
    col_no = vismaCaseDocuments.columns.get_loc('fileFound')
    for index, row in vismaCaseDocuments.iterrows():
        filePath = str(row.loc['File path'])
        print('Check if File Exists: ', str(index) + '/' + str(doc_count))
        if os.path.isfile(filePath)==True:
            vismaCaseDocuments.iat[index,col_no]='Yes'
    #save data to csv file
    vismaCaseDocuments.to_csv(csv_file, columns=(['Doc no.','fileFound']), index=False)

    vismaCaseDocuments[vismaCaseDocuments["fileFound"]=='No'].to_csv(csv_files_not_found, columns=(['Doc no.','fileFound','Case no.','File path']), index=False, encoding='mbcs')


sql="""
SELECT [Case no.],[Memo file]
FROM [dbo].[all_Cases]
ORDER BY [Case no.]
"""

print("read_sql (memoFile all_Cases)")
vismaCaseMemos = pd.read_sql(sql, conn)

RowCnt = len(vismaCaseData)
    
for index, row in vismaCaseData.iterrows():
    caseNo=str(row.loc['Case no.'])
    createdYear=r'20' + caseNo[2:4]
    
    print(f'CaseNo: {caseNo}   ({index}/{RowCnt})')
    
    htmlfilename = destinationDirectory +r'\\'+ db +r'\\Case_documents\\' + createdYear + r'\\'+ caseNo +r'\\Visma'+ caseNo +r'.html'

    os.makedirs(os.path.dirname(htmlfilename), exist_ok=True)
    
    with open(htmlfilename, "w") as file:
        file.write(htmlheader)

    with open(htmlfilename, "a") as file:
        file.write(bodyheader)

    with open(htmlfilename, "a") as file:
        tableHeader=r'<br><h3>Case data</h3>'
        file.write(tableHeader)
    
    # Remove 'Actor no 4' columns if [Actor no 4]==0
    if str(row.loc['Actor no 4'])=="0":
        del row['Actor no 4 Type']
        del row['Actor no 4']
        del row['Actor no 4 Name']
        del row['Actor no 4 ref']
        del row['Actor no 4 Description']

    # Remove 'Actor no 3' columns if [Actor no 3]==0
    if str(row.loc['Actor no 3'])=="0":
        del row['Actor no 3 Type']
        del row['Actor no 3']
        del row['Actor no 3 Name']
        del row['Actor no 3 ref']
        del row['Actor no 3 Description']

    # Remove 'Actor no 2' columns if [Actor no 2]==0
    if str(row.loc['Actor no 2'])=="0":
        del row['Actor no 2 Type']
        del row['Actor no 2']
        del row['Actor no 2 Name']
        del row['Actor no 2 ref']
        del row['Actor no 2 Description']

    # Remove 'Actor no 1' columns if [Actor no 1]==0
    if str(row.loc['Actor no 1'])=="0":
        del row['Actor no 1 Type']
        del row['Actor no 1']
        del row['Actor no 1 Name']
        del row['Actor no 1 ref']
        del row['Actor no 1 Description']

    row = row.reset_index()
    with open(htmlfilename, 'a') as table:
        row.to_html(table,header=False)


    ## ===================
    ##  Additional Actors
    ## ===================

    if not len(vismaCaseAddActors)==0:
        with open(htmlfilename, "a") as file:
            tableHeader=r'<br><h3>Add. actors</h3>'
            file.write(tableHeader)
        with open(htmlfilename, 'a') as table:
            df=vismaCaseAddActors.loc[vismaCaseAddActors['Case no.']==int(caseNo)]
            df=df.reset_index()
            df.to_html(table,header=True,index=True)

    ## ===========
    ##  Due dates
    ## ===========
    
    with open(htmlfilename, "a") as file:
        tableHeader=r'<br><h3>Due dates</h3>'
        file.write(tableHeader)
    with open(htmlfilename, 'a') as table:
        df=vismaCaseDuedates.loc[vismaCaseDuedates['Case no.']==int(caseNo)]
        df.to_html(table,header=True,index=False,columns=["User name","Due date","Type","Responsible","Description","Created date","Changed date","Resolved date"])


    ## ===========
    ##  Documents
    ## ===========
        
    df=vismaCaseDocuments.loc[vismaCaseDocuments['Case no.']==int(caseNo)]
   
    with open(htmlfilename, "a") as file:
        tableHeader=r'<br><h3>Documents</h3>'
        file.write(tableHeader)
    with open(htmlfilename, 'a') as table:
        df.to_html(table,header=True,index=False,columns=["Doc no.","Version no.","Associate no.","Actor","Version Create date","Original Create date","Doc.gr","Incomming","Description","Contact Name","Created by Usr"])

    with open(htmlfilename, "a") as file:
        tableHeader=r'<br><h3>Document file locations</h3>'
        file.write(tableHeader)
    with open(htmlfilename, 'a') as table:
        df.to_html(table,header=True,index=False,columns=["Doc no.","fileFound","File path"])

    ## ============
    ##  Memos file 
    ## ============

    df=vismaCaseMemos.loc[vismaCaseMemos['Case no.']==int(caseNo)].reindex()
    memoFile=str(df.iat[0,1])

    if path.exists(memoFile)==False and memoFile!="":
        with open(f"vismaMemoFileErrorlog_{db}.txt", 'a') as f1:
            f1.write(f"{now.strftime('%Y.%m.%d %H:%M:%S')} Case: {caseNo} Error: MemoFileNotFound File: [{memoFile}]" + os.linesep)

    if path.exists(memoFile)==True:
        with open(htmlfilename, "a") as file:
            tableHeader=r'<br><h3>Memo</h3>'
            file.write(tableHeader)

        f=open(memoFile, 'r') #open file for r=read (default character encoding is unicode utf-8) 
        mData=f.read()
        mData=mData.replace('\r\n','<br>')
        mData=mData.replace('\n','<br>')
        mData=mData.replace('\r','')

        mData = r'<table border="1" class="dataframe"><thead><th>File:</th></thead>  <tbody><tr><td>' + mData + r'</tbody></tr></td></table>'
        with open(htmlfilename, 'a') as htmlFile:
            htmlFile.write(mData)
                
    with open(htmlfilename, "a") as file:
        tableHeader=r'<br><h3>Memo file location</h3>'
        file.write(tableHeader)
    with open(htmlfilename, 'a') as table:
        df.to_html(table,header=True,index=False,columns=["Memo file"])


    ## =============
    ##  HTML footer
    ## =============

    with open(htmlfilename, "a") as file:
        file.write(htmlfooter)

