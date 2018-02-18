"""
Client Volume Monitor

"""

import numpy as np
import pandas as pd
from pandas import TimeGrouper
import pyodbc
from pandas import ExcelWriter
from datetime import datetime as dt
import calendar
from matplotlib import pyplot as plt
from functools import lru_cache

y = dt.now().year
m = dt.now().month
d = dt.now().day
currMonthID = y*100 + m
currMonthDays = calendar.monthrange(y, m)
mtd = (currMonthDays[1]/d)

# SQL
conn = pyodbc.connect(r'DRIVER={ODBC Driver 13 for SQL Server};'
                      r'SERVER=businteldw.stratose.com,1565;'
                      r'DATABASE=CAIDataWarehouse;'
                      r'Trusted_Connection=yes')
maxID = "13006445"
sql = '''
    select dc.ClientParentNameShort
        , dpr.Product
        , Grouped = 
        case dpr.Product
            when 'Dental' then 'Dental'
            when 'Workers Comp' then 'Workers Comp'
            else 'Medical'
        end
        , dd.DateMonthID 
        , dd.DateWeekID
        , dd.DateDay
        , dstc.CategoryDesc
        , sum(fc.ClaimCount) Claims
        , sum(fc.CMAllowed) Charges
        , lag(c.Claims, 1) over(partition by c.ClientParentNameShort, c.Product order by c.DateMonthID) claimsLag
        , lag(c.Charges, 1) over(partition by c.ClientParentNameShort, c.Product order by c.DateMonthID) chargesLag
       
    from FactClaim fc
        join DimDate dd on fc.dimdatereceivedkey = dd.dimdatekey
        join dimclient dc on fc.dimclientkey = dc.dimclientkey
        join dimclaimeligible dce on fc.dimclaimeligiblekey = dce.dimclaimeligiblekey
        join dimservicetypecategory dstc on fc.dimservicetypecategorykey = dstc.dimservicetypecategorykey
        join dimproduct dpr on fc.dimproductkey = dpr.dimproductkey
        join dimclaimtype dct on fc.dimclaimtypekey = dct.dimclaimtypekey
        
    where 
        fc.CMID in 
        (
            select fc1.CMID
            from FactClaim fc1
            where fc1.DimClaimEligibleKey = 1
                and fc1.CMID > %s 
            order by fc1.CMID desc offset 0 rows
        )
    
    group by 
        dc.ClientParentNameShort
        , dpr.Product
        , dd.DateMonthID
        , dd.DateWeekID  
        , dd.DateDay
        , dstc.CategoryDesc
;
''' % (maxID)

df = pd.read_sql(sql, conn)
conn.close()

# Calc and split
df['claimsMTD'] = df['Claims']*mtd
df['chargesMTD'] = df['Charges']*mtd
df['claims%Lag'] = df['claimsMTD'] / df['claimsLag'] - 1
df['charges%Lag'] = df['chargesMTD'] / df['chargesLag'] - 1
# df = df[df['DateMonthID'] == currMonthID]

dfFlag = df[(df.['charges%Lag'] >= .25) | (df.['charges%Lag'] <= -.25)]

dfMed = dfFlag[dfFlag['Grouped'] == "Medical"]
dfDent = dfFlag[dfFlag['Grouped'].Product == "Dental"]
dfWc = dfFlag[dfFlag['Grouped'] == "Workers Comp"]

# Write results to excel
writer = pd.ExcelWriter(r'C:\Users\pallen\Documents\Volume_Check_test.xlsx')

dfFlag[dfFlag['Grouped'] == "Medical"].to_excel(writer, 'Sheet1')
dfFlag[dfFlag['Grouped'] == "Medical"].to_excel(writer, 'Medical')
dfFlag[dfFlag['Grouped'].Product == "Dental"].to_excel(writer, 'Dental')
dfFlag[dfFlag['Grouped'] == "Workers Comp"].to_excel(writer, 'WC')

writer.save()

'''
check
total = len(df_flag)
total == len(df_med) + len(df_dent) + len(df_wc)


@lru_cache(maxsize=3)
def pull(connection, query, toMTD):

    dfPull = pd.read_sql(connection, query)

    dfPull.groupby('DateMonthID')

    dfPull['claimsMTD'] = dfPull['Claims'] * toMTD
    dfPull['chargesMTD'] = dfPull['Charges'] * toMTD
    dfPull['claims%Lag'] = dfPull['claimsMTD'] / dfPull['claimsLag'] - 1
    dfPull['charges%Lag'] = dfPull['chargesMTD'] / dfPull['chargesLag'] - 1

    return dfPull
'''