"""
Client Service Category Vol/KPI Monitor



"""

import numpy as np
import pyodbc
import pandas as pd
from pandas import TimeGrouper
from pandas import ExcelWriter
import xlsxwriter
import matplotlib.pyplot as plt
from functools import lru_cache
from datetime import datetime as dt
import calendar
import pyarrow as pa

# Declarations
# List of clients and categories to be included
client = 'Cigna East'
catList = ['Anesthesia', 'Ambulatory Surgical Care']
numCats = len(catList)


# Connect to and query SQL server for given Client & Service Category
def readSQL(client, category):

    # SQl context
    conn = pyodbc.connect(r'DRIVER={ODBC Driver 13 for SQL Server};'
                          r'SERVER=businteldw.stratose.com,1565;'
                          r'DATABASE=CAIDataWarehouse;'
                          r'Trusted_Connection=yes')
    sql = '''
        select --dc.ClientParentNameShort Client
            --, dst.CategoryDesc Category
            --, dd.DateMonthSSRS Month
            --, 
            dd.DateDay 
            , count(f.CMID) Claims
            , sum(f.CMAllowed) Allowed
            , sum(f.CMAllowedHit) Hit
            , sum(f.LineSavings) Savings
            , sum(f.CMAllowedhit)/sum(f.CMAllowed) HitRate
            , sum(f.LineSavings)/sum(f.CMAllowedHit) SaveRate
            , sum(f.LineSavings)/sum(f.CMAllowed) SaveRateEff
        
        
        from v_FactClaimLine f 
            join DimDate dd on f.dimdatereceivedkey = dd.dimdatekey
            join dimclient dc on f.dimclientkey = dc.dimclientkey
            join dimclaimeligible dce on f.dimclaimeligiblekey = dce.dimclaimeligiblekey
            join dimservicetype dst on f.dimservicetypekey = dst.dimservicetypekey
            join dimproduct dpr on f.dimproductkey = dpr.dimproductkey
            join dimclaimtype dct on f.dimclaimtypekey = dct.dimclaimtypekey
            join DimProvider prov on f.DimProviderKey = prov.DimProviderKey
        
        where dce.ClaimEligible = 'Eligible'
            and dd.DateDay between (convert(date, getdate() - 30)) and (convert(date, getdate()))
            and dc.ClientParentNameShort = '%s'
            and dst.CategoryDesc = '%s'
        
        group by --dc.ClientParentNameShort
            --, dst.CategoryDesc
            --, dd.DateMonthSSRS
            --, 
            dd.DateDay
        
        order by --dc.ClientParentNameShort
            --, dst.CategoryDesc
            --, dd.DateMonthSSRS
            --, 
            dd.DateDay
    ''' % (client, catList[0])

    # Read SQL data, set date as index, close connection to server
    df = pd.read_sql(sql, conn)
    df = df.set_index('DateDay')
    conn.close()

    return df


# Write results to Excel, add charts, format
def toExcel(df, sheet):

    #
    dfKPI = df.loc[:, 'HitRate':]
    rowCount = len(df.index)

    writer = pd.ExcelWriter(r'C:\Users\pallen\Documents\VolumeCheck.xlsx', engine='xlsxwriter')

    dfKPI.to_excel(writer, sheet_name=sheet, startrow=1, startcol=1)

    '''
    workbook = writer.book
    worksheet = writer.sheets[sheet]
    chart = workbook.add_chart({'type': 'line'})
    chart.add_series({'values': '=%s!$B$2:$D$%d'}) % (sheet, rowCount)
    worksheet.insert_chart('F2', chart)
    '''

    format1 = workbook.add_format({'num_format': '$#,##0'})
    format2 = workbook.add_format({'num_format': '0%'})
    format3 = workbook.add_format({'num_format': '#,##0'})

    worksheet.set_column('C:E', 15, format2)

    writer.save()

    return


for i in range(0, numCats):
    df = readSQL(client, catList[i])
    toExcel(df, catList[i])


cat = catList[0]
readSQL(client, cat)
