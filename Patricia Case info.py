# ======================================================
#    Define Excel filename with date-time in filename
#    "tabName" is used for naming the tab in Excel
# ======================================================

import datetime
import os
import shutil

now = datetime.datetime.now()
sheetName = now.strftime('%Y%m%d')

fileName = now.strftime('Patricia Case Info (%Y%m%d) Rds03.xlsx')

if os.path.exists(fileName):
    print ('existing file found')
    RevNo = 1
    while os.path.exists(fileName.replace(')', f' Rev{RevNo})')):
        RevNo += 1
    fileName = fileName.replace(')', f' Rev{RevNo})')

if os.path.exists(fileName):
    print ('removing existing file')
    os.remove(fileName)

currentDirectory = os.getcwd()

destinationDirectory = r'\\SVRCPHFILE01\Data$\Public\Zacco Excel Reports\Patricia\Rds03'

    
# ===================================
#    Read data from the SQL Server
# ===================================
import pandas as pd
import pyodbc

conn = pyodbc.connect('Driver={SQL Server};Server=SVRCMSTESTDB01;Database=Patricia_Rds03;Tructed_Connection=yes;')

sql = """

-- USE [Patricia_rds03]; 
-- Patricia Case Info Natalia 2021_01_12
-- Natalia 2021_01_18
-- IF OBJECT_ID('tempdb..##cases') IS NOT NULL DROP TABLE ##cases

SELECT --TOP (100)
	ISNULL(tl.[CASE_TYPE_TEXT] , SPACE(0)) AS [Case Type]
	, appt.[APPLICATION_TYPE_NAME] AS [Application Type]
	, cn.[CASE_NUMBER] AS [Case No.]
	, CASE WHEN CHARINDEX(' from case ', pc.[CASE_REMARK]) > 0 THEN SUBSTRING (pc.[CASE_REMARK], CHARINDEX(' from case ', pc.[CASE_REMARK]) +11, CASE WHEN CHARINDEX(',', pc.[CASE_REMARK], CHARINDEX(' from case ', pc.[CASE_REMARK])+11) = 0 THEN 200 ELSE (CHARINDEX(',', pc.[CASE_REMARK], CHARINDEX(' from case ', pc.[CASE_REMARK])+11) - CHARINDEX(' from case ', pc.[CASE_REMARK]) -11) END) ELSE cn.[CASE_NUMBER] END AS [from case]
	, ROW_NUMBER() OVER (PARTITION BY CASE WHEN CHARINDEX(' from case ', pc.[CASE_REMARK]) > 0 THEN SUBSTRING (pc.[CASE_REMARK], CHARINDEX(' from case ', pc.[CASE_REMARK]) +11, CASE WHEN CHARINDEX(',', pc.[CASE_REMARK], CHARINDEX(' from case ', pc.[CASE_REMARK])+11) = 0 THEN 200 ELSE (CHARINDEX(',', pc.[CASE_REMARK], CHARINDEX(' from case ', pc.[CASE_REMARK])+11) - CHARINDEX(' from case ', pc.[CASE_REMARK]) -11) END) ELSE cn.[CASE_NUMBER] END ORDER BY pc.[CASE_ID] ) AS [RowNo] 
	, ISNULL(pc.[OLD_CASE_ID], SPACE(0)) AS [Matter ID]
	, ISNULL(pc.[STATE_ID] , SPACE(0)) AS [Country]
	, ISNULL(pc.[CASE_CATCH_WORD] , SPACE(0)) AS [Catchword]
	, ISNULL(tmc.[TRADE_MARK_CATEGORY_LABEL] , SPACE(0)) AS [Appearance]
	, slt.[SERVICE_LEVEL_LABEL] AS [Service Level] 
	, ISNULL(pst.[STATUS_LABEL] , SPACE(0)) AS [Status]
	, ISNULL(te.[LOGIN_ID] , SPACE(0)) AS [Case Responsible]
	, ISNULL(ct.[LOGIN_ID] , SPACE(0)) AS [Case Administrator]
	, ISNULL(pc.[CASE_REMARK] , SPACE(0)) AS [Case Remark]

	, ISNULL(cl.[CLASS], SPACE(0)) AS  [Class]
	, ISNULL(cl.[GOODS_TEXT], SPACE(0)) AS  [Class Goods and Services]

	, ISNULL(ds.[STATE_ID], SPACE(0)) AS  [Designation Country]
	, ISNULL(et.[EVENT_SCHEME_ID], SPACE(0)) AS  [Term ID]
	, et.[TERM_DATE] AS  [Term Date]

	, dt20035.[DIARY_TEXT] AS [CPM Company] -- 20035  CPM COMPANY
	, dd20090.[DIARY_DATE] AS [Case Send to Dynamics] -- 20090  CASE SENT TO DYNAMICS
	, dd38.[DIARY_DATE] AS [Date of Order] -- 38  DATE OF ORDER
	, dd3.[DIARY_DATE] AS [Application Date] -- 3  BASIC APPLICATION DATE
	, dt4.[DIARY_TEXT] AS [Application No.] -- 4  BASIC APPLICATION NO
	, dd35.[DIARY_DATE] AS [Registration Date] -- 35  PATENT REGISTRATION DATE
	, dt37.[DIARY_TEXT] AS [Registration No.] -- 37  PATENT REGISTRATION NO
	, dd39.[DIARY_DATE] AS [Next Renewal] -- 39  NEXT ANNUITY/RENEWAL/MAINTENANCE FEE
	, dt10225.[DIARY_TEXT] AS [Design Type] -- 10225  DESIGN TYPE
	, dt337.[DIARY_TEXT] AS [Joint Registration] -- 337  JOINT REGISTRATION
	, dt10216.[DIARY_TEXT] AS [Number of Designs as Filed] -- 10216  NUMBER OF DESIGNS AS FILED
	, dt10217.[DIARY_TEXT] AS [Number of Designs as Renewed] -- 10217  NUMBER OF DESIGNS AS RENEWED
	, dt7.[DIARY_TEXT] AS [Priority Country] -- 7  PRIORITY COUNTRY
	, dd5.[DIARY_DATE] AS [Priority Date] -- 5  PRIORITY DATE
	, dt6.[DIARY_TEXT] AS [Priority No] -- 6  PRIORITY NO
	, dd11.[DIARY_DATE] AS [Publication Date] -- 11  PUBLICATION DATE
	, dt20105.[DIARY_TEXT] AS [Seniority Claimed] -- 20105  SENIORITY CLAIMED
	, dt20106.[DIARY_TEXT] AS [Seniority Claimed In] -- 20106  SENIORITY CLAIMED IN
	, dd20101.[DIARY_DATE] AS [To be Abandoned] -- 20101  TO BE ABANDONED
	, dd20167.[DIARY_DATE] AS [To be Abandoned - Designations] -- 20167  TO BE ABANDONED - DESIGNATIONS

	, GETDATE() AS [SqlDatetime]
	, pc.[CASE_ID]

--	, CAST(NULL AS varchar(3)) AS [Error(Class)]
--	, CAST(NULL AS varchar(3)) AS [Error(Class Goods and Services)]
--	, CAST(NULL AS varchar(3)) AS [Error(Designation Country)]
--	, CAST(NULL AS varchar(3)) AS [Error(Term ID)]
--	, CAST(NULL AS varchar(3)) AS [Error(Term Date)]

--INTO ##cases

FROM [dbo].[PAT_CASE] AS pc

	-- SELECT [CASE_ID] FROM [dbo].[VW_CASE_NUMBER] GROUP BY [CASE_ID] HAVING COUNT(*)>1
	LEFT JOIN [dbo].[VW_CASE_NUMBER] AS cn
		ON pc.[CASE_ID] = cn.[CASE_ID]
	
	-- SELECT [STATUS_ID] FROM [dbo].[PAT_STATUS_TEXT] WHERE [LANGUAGE_ID] = 3 GROUP BY [STATUS_ID] HAVING COUNT(*)>1
	LEFT JOIN [dbo].[PAT_STATUS_TEXT] AS pst
		ON pst.[STATUS_ID] = pc.[STATUS_ID]
		AND pst.[LANGUAGE_ID] = 3

	--Case Administrator (***** 1-MANY IS ALLOWED)
	-- SELECT * FROM [ROLES_LABEL] WHERE [LANGUAGE_ID] = 3 AND [ROLE_ID] = 37
	-- SELECT TOP 100 [CASE_ID] , [ROLE_ID] FROM [dbo].[CASE_TEAM] WHERE [ROLE_ID] = 37 GROUP BY [CASE_ID] , [ROLE_ID] HAVING COUNT(*)>1
	-- SELECT * FROM [dbo].[CASE_TEAM] WHERE [ROLE_ID] = 37 AND [CASE_ID] IN (SELECT TOP 100 [CASE_ID] FROM [dbo].[CASE_TEAM] WHERE [ROLE_ID] = 37 GROUP BY [CASE_ID] , [ROLE_ID] HAVING COUNT(*)>1) ORDER BY [CASE_ID] 
	LEFT JOIN (SELECT [CASE_ID], STRING_AGG([LOGIN_ID], N', ') AS [LOGIN_ID] FROM [dbo].[CASE_TEAM] WHERE [ROLE_ID] = 37 GROUP BY [CASE_ID]) AS ct
		ON pc.[CASE_ID] = ct.[CASE_ID]

	-- Case Responsible
	-- SELECT * FROM [ROLES_LABEL] WHERE [LANGUAGE_ID] = 3 AND [ROLE_ID] = 10
	-- SELECT [CASE_ID] FROM [dbo].[WORK_GROUP] GROUP BY [CASE_ID] HAVING COUNT(*)>1
	LEFT JOIN [dbo].[WORK_GROUP] AS wg
		ON pc.[CASE_ID] = wg.[CASE_ID]

	-- SELECT [TEAM_ID] FROM [dbo].[TEAM_EFFORT] WHERE [ROLE_ID] = 10 GROUP BY [TEAM_ID] HAVING COUNT(*)>1
	LEFT JOIN [dbo].[TEAM_EFFORT] AS te -- [Case Responsible]
		ON wg.[TEAM_ID] = te.[TEAM_ID]
		AND te.[ROLE_ID] = 10

	-- SELECT [CASE_TYPE_ID] FROM [dbo].[CASE_TYPE_LABEL] WHERE [CASE_TYPE_LANGUAGE] = 3 GROUP BY [CASE_TYPE_ID] HAVING COUNT(*)>1
	LEFT JOIN [dbo].[CASE_TYPE_LABEL] as tl
		ON pc.[CASE_TYPE_ID] = tl.[CASE_TYPE_ID]
		AND tl.[CASE_TYPE_LANGUAGE] = 3

	-- SELECT [APPLICATION_TYPE_ID] FROM [dbo].[APPLICATION_TYPE_TEXT] WHERE [LANGUAGE_ID] = 3 GROUP BY [APPLICATION_TYPE_ID] HAVING COUNT(*)>1
	LEFT JOIN [dbo].[APPLICATION_TYPE_TEXT] AS appt
		ON pc.[APPLICATION_TYPE_ID] = appt.[APPLICATION_TYPE_ID]
		AND appt.[LANGUAGE_ID] = 3

	-- SELECT [SERVICE_LEVEL_ID] FROM [dbo].[SERVICE_LEVEL_TEXT] WHERE [LANGUAGE_ID] = 3 GROUP BY [SERVICE_LEVEL_ID] HAVING COUNT(*)>1
	LEFT JOIN [dbo].[SERVICE_LEVEL_TEXT] AS slt
		ON pc.[SERVICE_LEVEL_ID] = slt.[SERVICE_LEVEL_ID]
		AND slt.[LANGUAGE_ID] = 3

	-- SELECT [TRADE_MARK_CATEGORY] FROM [dbo].[PAT_CASE_TM_CATEGORY_TEXT] WHERE [TRADE_MARK_CATEGORY_LANGUAGE] = 3 GROUP BY [TRADE_MARK_CATEGORY] HAVING COUNT(*)>1
	LEFT JOIN [dbo].[PAT_CASE_TM_CATEGORY_TEXT] AS tmc
		ON tmc.[TRADE_MARK_CATEGORY] = pc.[TRADE_MARK_CATEGORY] 
		AND tmc.[TRADE_MARK_CATEGORY_LANGUAGE] = 3


	-- EXEC SP_HELP N'DESIGNATED_STATES'
	-- SELECT COUNT(*) FROM [dbo].[DESIGNATED_STATES] WHERE [STATE_ID] IS NULL
	LEFT JOIN (SELECT [CASE_ID], STRING_AGG([STATE_ID], N', ') AS [STATE_ID]
			FROM [dbo].[DESIGNATED_STATES]
			WHERE [NOT_RENEWED] = 0 OR [NOT_RENEWED] IS NULL
			GROUP BY [CASE_ID] --(ORDER BY [CASE_ID], [STATE_ID])
	) AS ds
	ON pc.[CASE_ID] = ds.[CASE_ID]

	-- EXEC SP_HELP N'PATENT_CLASS'
	-- EXEC SP_HELP N'PATENT_CLASS_TYPES_LABEL'
	LEFT JOIN (SELECT pc.[CASE_ID] 
				, STRING_AGG(CAST(pct.[PATENT_CLASS_TYPE_ID] AS nvarchar(50)) +N': '+ ISNULL(pct.[PATENT_CLASS_TYPE_LABEL], SPACE(0)), N', ') AS [CLASS]
				, SPACE(0) AS [GOODS_TEXT]
				, 2 AS [Case Type] --2=Patent
			FROM [dbo].[PATENT_CLASS] AS pc
			JOIN [PATENT_CLASS_TYPES_LABEL] AS pct
			ON pc.[PATENT_CLASS_TYPE] = pct.[PATENT_CLASS_TYPE_ID]
			AND pct.[LANGUAGE_ID] = 3
			GROUP BY pc.[CASE_ID]
			UNION ALL

	-- EXEC SP_HELP N'DESIGN_CLASS'
			SELECT [CASE_ID] 
				, STRING_AGG(N'('+CAST([DESIGN_NUMBER] AS nvarchar(50)) +N') '+ [DESIGN_CLASS_ID], N', ') AS [CLASS]
				, STRING_AGG(ISNULL([DESIGN_GOODS_TEXT], SPACE(0)), N'|<->| ') AS [GOODS_TEXT]
				, 3 AS [Case Type] --Design
			FROM [dbo].[DESIGN_CLASS] 
			WHERE [LANGUAGE_ID] = 3
			GROUP BY [CASE_ID]
			UNION ALL
				-- EXEC SP_HELP N'TRADE_MARK_CLASS'
			SELECT [CASE_ID] 
				, STRING_AGG([TRADE_MARK_CLASS], N', ') AS [CLASS]
				, STRING_AGG(ISNULL([GOODS_TEXT], SPACE(0)), N'|<->| ') AS [GOODS_TEXT]
				, 4 AS [Case Type] --4=Trademark
			FROM [dbo].[TRADE_MARK_CLASS] 
			WHERE [LANGUAGE_ID] = 3
			GROUP BY [CASE_ID]
	) AS cl
	ON pc.[CASE_ID] = cl.[CASE_ID]
	AND tl.[CASE_TYPE_ID] = cl.[Case Type]

	--exec sp_help N'EVENT'
	LEFT JOIN (SELECT [CASE_ID] --NOT NULL
					, STRING_AGG([EVENT_SCHEME_ID], N', ') AS [EVENT_SCHEME_ID] --NOT NULL
					, STRING_AGG(ISNULL(CONVERT(nvarchar(10), [TERM_DATE], 121), SPACE(0)), N', ') AS [TERM_DATE] --NULLABLE
				FROM [dbo].[EVENT]
				-- SELECT TOP (10) * FROM [dbo].[EVENT]
				WHERE [DONE_DATE] IS NULL 
					AND [REOPENED_DATE] IS NOT NULL 
				GROUP BY [CASE_ID] 

	) AS et -- ACTIONS TERM ID & ACTIONS TERM DATE
	ON pc.[CASE_ID] = et.[CASE_ID]


	-- SELECT [CASE_ID] FROM [dbo].[DIARY_TEXT] WHERE [FIELD_NUMBER] = 20035 GROUP BY [CASE_ID] HAVING COUNT(*)>1
	LEFT JOIN [dbo].[DIARY_TEXT] AS dt20035
		ON pc.[CASE_ID] = dt20035.[CASE_ID]
		AND dt20035.[FIELD_NUMBER] = 20035

	-- SELECT [CASE_ID] FROM [dbo].[DIARY_TEXT] WHERE [FIELD_NUMBER] = 4 GROUP BY [CASE_ID] HAVING COUNT(*)>1
	LEFT JOIN [dbo].[DIARY_TEXT] AS dt4  -- [Application No.]
		ON pc.[CASE_ID] = dt4.[CASE_ID]
		AND dt4.[FIELD_NUMBER] = 4

	-- SELECT [CASE_ID] FROM [dbo].[DIARY_TEXT] WHERE [FIELD_NUMBER] = 10225 GROUP BY [CASE_ID] HAVING COUNT(*)>1
	LEFT JOIN [dbo].[DIARY_TEXT] AS dt10225  -- 10225  DESIGN TYPE
		ON pc.[CASE_ID] = dt10225.[CASE_ID]
		AND dt10225.[FIELD_NUMBER] = 10225

	-- SELECT [CASE_ID] FROM [dbo].[DIARY_DATE] WHERE [FIELD_NUMBER] = 3 GROUP BY [CASE_ID] HAVING COUNT(*)>1
	LEFT JOIN [dbo].[DIARY_DATE] AS dd3  -- [Application Date]
		ON pc.[CASE_ID] = dd3.[CASE_ID]
		AND dd3.[FIELD_NUMBER] = 3

	-- SELECT [CASE_ID] FROM [dbo].[DIARY_TEXT] WHERE [FIELD_NUMBER] = 37 GROUP BY [CASE_ID] HAVING COUNT(*)>1
	LEFT JOIN [dbo].[DIARY_TEXT] AS dt37  -- [Registration No.]
		ON pc.[CASE_ID] = dt37.[CASE_ID]
		AND dt37.[FIELD_NUMBER] = 37

	-- SELECT [CASE_ID] FROM [dbo].[DIARY_DATE] WHERE [FIELD_NUMBER] = 35 GROUP BY [CASE_ID] HAVING COUNT(*)>1
	LEFT JOIN [dbo].[DIARY_DATE] AS dd35  -- [Registration Date]
		ON pc.[CASE_ID] = dd35.[CASE_ID]
		AND dd35.[FIELD_NUMBER] = 35

	-- SELECT [CASE_ID] FROM [dbo].[DIARY_DATE] WHERE [FIELD_NUMBER] = 5 GROUP BY [CASE_ID] HAVING COUNT(*)>1
	LEFT JOIN [dbo].[DIARY_DATE] AS dd5  -- [PRIORITY DATE]
		ON pc.[CASE_ID] = dd5.[CASE_ID]
		AND dd5.[FIELD_NUMBER] = 5

	-- SELECT [CASE_ID] FROM [dbo].[DIARY_DATE] WHERE [FIELD_NUMBER] = 20090 GROUP BY [CASE_ID] HAVING COUNT(*)>1
	LEFT JOIN [dbo].[DIARY_DATE] AS dd20090  -- [Case Send to Dynamics]
		ON pc.[CASE_ID] = dd20090.[CASE_ID]
		AND dd20090.[FIELD_NUMBER] = 20090

	-- SELECT [CASE_ID] FROM [dbo].[DIARY_DATE] WHERE [FIELD_NUMBER] = 38 GROUP BY [CASE_ID] HAVING COUNT(*)>1
	LEFT JOIN [dbo].[DIARY_DATE] AS dd38  -- [Date Created]
		ON pc.[CASE_ID] = dd38.[CASE_ID]
		AND dd38.[FIELD_NUMBER] = 38

	-- SELECT [CASE_ID] FROM [dbo].[DIARY_TEXT] WHERE [FIELD_NUMBER] = 10216 GROUP BY [CASE_ID] HAVING COUNT(*)>1
	LEFT JOIN [dbo].[DIARY_TEXT] AS dt10216  -- 10216  NUMBER OF DESIGNS AS FILED
		ON pc.[CASE_ID] = dt10216.[CASE_ID]
		AND dt10216.[FIELD_NUMBER] = 10216

	-- SELECT [CASE_ID] FROM [dbo].[DIARY_TEXT] WHERE [FIELD_NUMBER] = 10217 GROUP BY [CASE_ID] HAVING COUNT(*)>1
	LEFT JOIN [dbo].[DIARY_TEXT] AS dt10217  -- 10217  NUMBER OF DESIGNS AS RENEWED
		ON pc.[CASE_ID] = dt10217.[CASE_ID]
		AND dt10217.[FIELD_NUMBER] = 10217

	-- SELECT [CASE_ID] FROM [dbo].[DIARY_TEXT] WHERE [FIELD_NUMBER] = 7 GROUP BY [CASE_ID] HAVING COUNT(*)>1
	LEFT JOIN [dbo].[DIARY_TEXT] AS dt7  -- 7  PRIORITY COUNTRY
		ON pc.[CASE_ID] = dt7.[CASE_ID]
		AND dt7.[FIELD_NUMBER] = 7

	-- SELECT [CASE_ID] FROM [dbo].[DIARY_TEXT] WHERE [FIELD_NUMBER] = 6 GROUP BY [CASE_ID] HAVING COUNT(*)>1
	LEFT JOIN [dbo].[DIARY_TEXT] AS dt6  -- 6  PRIORITY NO	
		ON pc.[CASE_ID] = dt6.[CASE_ID]
		AND dt6.[FIELD_NUMBER] = 6

	-- SELECT [CASE_ID] FROM [dbo].[DIARY_TEXT] WHERE [FIELD_NUMBER] = 20105 GROUP BY [CASE_ID] HAVING COUNT(*)>1
	LEFT JOIN [dbo].[DIARY_TEXT] AS dt20105  -- 20105  SENIORITY CLAIMED
		ON pc.[CASE_ID] = dt20105.[CASE_ID]
		AND dt20105.[FIELD_NUMBER] = 20105
		
	-- SELECT [CASE_ID] FROM [dbo].[DIARY_TEXT] WHERE [FIELD_NUMBER] = 20106 GROUP BY [CASE_ID] HAVING COUNT(*)>1
	LEFT JOIN [dbo].[DIARY_TEXT] AS dt20106  -- 20106  SENIORITY CLAIMED IN
		ON pc.[CASE_ID] = dt20106.[CASE_ID]
		AND dt20106.[FIELD_NUMBER] = 20106

	-- SELECT [CASE_ID] FROM [dbo].[DIARY_TEXT] WHERE [FIELD_NUMBER] = 337 GROUP BY [CASE_ID] HAVING COUNT(*)>1
	LEFT JOIN [dbo].[DIARY_TEXT] AS dt337  -- 337  JOINT REGISTRATION
		ON pc.[CASE_ID] = dt337.[CASE_ID]
		AND dt337.[FIELD_NUMBER] = 337

	-- SELECT [CASE_ID] FROM [dbo].[DIARY_DATE] WHERE [FIELD_NUMBER] = 39 GROUP BY [CASE_ID] HAVING COUNT(*)>1
	LEFT JOIN [dbo].[DIARY_DATE] AS dd39  -- [Next Renewal]
		ON pc.[CASE_ID] = dd39.[CASE_ID]
		AND dd39.[FIELD_NUMBER] = 39

	-- SELECT [CASE_ID] FROM [dbo].[DIARY_DATE] WHERE [FIELD_NUMBER] = 11 GROUP BY [CASE_ID] HAVING COUNT(*)>1
	LEFT JOIN [dbo].[DIARY_DATE] AS dd11  -- 11  PUBLICATION DATE
		ON pc.[CASE_ID] = dd11.[CASE_ID]
		AND dd11.[FIELD_NUMBER] = 11

	-- SELECT [CASE_ID] FROM [dbo].[DIARY_DATE] WHERE [FIELD_NUMBER] = 20101 GROUP BY [CASE_ID] HAVING COUNT(*)>1
	LEFT JOIN [dbo].[DIARY_DATE] AS dd20101  -- 20101  TO BE ABANDONED
		ON pc.[CASE_ID] = dd20101.[CASE_ID]
		AND dd20101.[FIELD_NUMBER] = 20101
		
	-- SELECT [CASE_ID] FROM [dbo].[DIARY_DATE] WHERE [FIELD_NUMBER] = 20167 GROUP BY [CASE_ID] HAVING COUNT(*)>1
	LEFT JOIN [dbo].[DIARY_DATE] AS dd20167  -- 20167  TO BE ABANDONED - DESIGNATIONS
		ON pc.[CASE_ID] = dd20167.[CASE_ID]
		AND dd20167.[FIELD_NUMBER] = 20167

WHERE 1=1
    AND tl.[CASE_TYPE_TEXT] in ('Design','Trademark')
    AND pc.[STATE_ID] IN ('EU','GB','IM','WO')
    AND pst.[STATUS_LABEL] IN ('Registered','To be abandoned','Opposition','Filed')

"""

print("read_sql/before")
df = pd.read_sql(sql, conn)
print("read_sql/after")

# ==============================
#    Write data to Excel file
# ==============================

writer = pd.ExcelWriter(fileName, engine='xlsxwriter', date_format='yyyy-MM-dd', datetime_format='yyyy-MM-dd hh:mm:ss')

print("to_excel")
##df.to_excel(fileName, sheet_name = sheetName, float_format = '%.2f' , index = False , engine='xlsxwriter' )
df.to_excel(writer  , sheet_name = sheetName, float_format = '%.2f' , index = False , engine='xlsxwriter' )

print("worksheet")
worksheet = writer.sheets[sheetName]
print("workbook")
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
if 1==1:
    for i, col in enumerate(df.columns):
        # find length of column i
        column_len = df[col].astype(str).str.len().max()
        # Setting the length if the column header is larger
        # than the max column value length
        column_len = max(column_len, len(col)) + 2
        # set the column length
        worksheet.set_column(i, i, column_len)

# save writer object and created Excel file with data from DataFrame
print('before save')
writer.save()
print('after save')

# ==========================================
#    Move Excel file to Network direcroty
# ==========================================

print ("moving file to network directory")
##shutil.move(os.path.join(currentDirectory, fileName), os.path.join(destinationDirectory, fileName))
