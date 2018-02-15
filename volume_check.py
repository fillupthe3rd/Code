"""
Client Volume Monitor

"""

import numpy as np
import pandas as pd
import pyodbc
from pandas import ExcelWriter

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
            , dpr.DimProductKey
            --, dd.DateDay
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
            , dpr.DimProductKey
            --, dd.DateDay
        
    )
    
    select c.ClientParentNameShort
        , c.Product
        , c.DimProductKey
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
df['pDiff_Claims'] = df['Claims'] / df['Claims_prev'] - 1
df['pDiff_Charges'] = df['Charges'] / df['Charges_prev'] - 1

df_flag = df[(df.pDiff_Charges >= .25) | (df.pDiff_Charges <= -.25)]
df_grp = df_flag.groupby('Product')

df_med = df_grp.get_group('Group Health')
    df_med = df_med.append(df_grp.get_group('Claim Settlement (PPN)'))
        df_med = df_med.append(df_grp.get_group('Medicare Pricing Solutions'))

df_dent = df_grp.get_group('Dental')
df_wc = df_grp.get_group('Workers Comp')

# Write results to excel
writer = pd.ExcelWriter(r'C:\Users\pallen\Documents\Volume_Check_test.xlsx')
df.to_excel(writer, 'Sheet1')

writer.save()


# check

total = len(df_flag)
total == len(df_med) + len(df_dent) + len(df_wc)