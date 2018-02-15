"""
Client Volume Monitor

"""

import numpy as np
import pandas as pd
import pyodbc
from pandas import ExcelWriter
from datetime import datetime as dt
import calendar


y = dt.now().year
m = dt.now().month
d = dt.now().day
currentMonthID = y*100 + m
currMonthDays = calendar.monthrange(dt.now().year, dt.now().month)
mtd = (currMonthDays[1]/d)

# SQL
conn = pyodbc.connect(r'DRIVER={ODBC Driver 13 for SQL Server};'
                      r'SERVER=businteldw.stratose.com,1565;'
                      r'DATABASE=CAIDataWarehouse;'
                      r'Trusted_Connection=yes')

sql = '''
    with cte_daily as
    (
        select dc.ClientParentNameShort
            , dpr.Product
            , dd.DateMonthID
            , sum(fc.ClaimCount) Claims
            , sum(fc.CMAllowed) Charges
            
        from FactClaim fc
            join DimDate dd on fc.dimdatereceivedkey = dd.dimdatekey
            join dimclient dc on fc.dimclientkey = dc.dimclientkey
            join dimclaimeligible dce on fc.dimclaimeligiblekey = dce.dimclaimeligiblekey
            join dimdiscountmethod ddm on fc.dimdiscountmethodkey = ddm.dimdiscountmethodkey
            join dimprovider dp on fc.dimproviderkey = dp.dimproviderkey
            join dimservicetypecategory dstc on fc.dimservicetypecategorykey = dstc.dimservicetypecategorykey
            join dimnetwork dn on fc.dimnetworkkey = dn.dimnetworkkey
            join dimproduct dpr on fc.dimproductkey = dpr.dimproductkey
            join dimclaimtype dct on fc.dimclaimtypekey = dct.dimclaimtypekey
            join DimClaimStatus dcs on fc.DimClaimStatusKey = dcs.DimClaimStatusKey
    
        where 
            dce.ClaimEligible = 'Eligible'
                and dd.DateDay between (convert(date, getdate() - 120)) and (convert(date, getdate()))
              
        group by 
            dc.ClientParentNameShort
            , dpr.Product
            , dd.DateMonthID
        
    )
    
    select c.ClientParentNameShort
        , c.Product
        , c.DateMonthID
        , c.Claims
        , c.Charges
        , lag(c.Claims, 1) over(partition by c.ClientParentNameShort, c.Product order by c.DateMonthID) Claims_prev
        , lag(c.Charges, 1) over(partition by c.ClientParentNameShort, c.Product order by c.DateMonthID) Charges_prev
        , c.Claims - (lag(c.Claims, 1) over(partition by c.ClientParentNameShort, c.Product order by c.DateMonthID)) as diff_claims
        , c.Charges - (lag(c.Charges, 1) over(partition by c.ClientParentNameShort, c.Product order by c.DateMonthID)) as diff_charges
        --, cast(c.Claims / (lag(c.Claims, 1) over(partition by c.ClientParentNameShort, c.Product order by c.DateMonthID)) as decimal(20,3)) as pdiff_claims
        --, cast(c.Charges / (lag(c.Charges, 1) over(partition by c.ClientParentNameShort, c.Product order by c.DateMonthID)) -1 as decimal(20,3)) as pdiff_charges
        
    from cte_daily c 
    
    order by c.ClientParentNameShort, c.Product, c.DateMonthID

'''

df = pd.read_sql(sql, conn)
conn.close()

# Calc and split
df['Claims_MTD'] = df['Claims']*mtd
df['Charges_MTD'] = df['Charges']*mtd
df['pDiff_Claims'] = df['Claims_MTD'] / df['Claims_prev'] - 1
df['pDiff_Charges'] = df['Charges_MTD'] / df['Charges_prev'] - 1
df = df[df['DateMonthID'] == currMonthID]

df_flag = df[(df.pDiff_Charges >= .25) | (df.pDiff_Charges <= -.25)]
df_grp = df_flag.groupby('Product')

df_med = df_grp.get_group('Group Health')
    df_med = df_med.append(df_grp.get_group('Claim Settlement (PPN)'), ignore_index=True)
        df_med = df_med.append(df_grp.get_group('Medicare Pricing Solutions'), ignore_index=True)

df_dent = df_grp.get_group('Dental')
df_wc = df_grp.get_group('Workers Comp')

# Write results to excel
writer = pd.ExcelWriter(r'C:\Users\pallen\Documents\Volume_Check_test.xlsx')

df_flag.to_excel(writer, 'Sheet1')
df_med.to_excel(writer, 'Medical')
df_dent.to_excel(writer, 'Dental')
df_wc.to_excel(writer, 'WC')

writer.save()

# check
total = len(df_flag)
total == len(df_med) + len(df_dent) + len(df_wc)