import numpy as np
import pandas as pd
from pandas import TimeGrouper
from pandas import ExcelWriter
import pyodbc
import calendar
from matplotlib import pyplot as plt
from datetime import datetime as dt

# VarDec
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
sql = '''
    with cte_daily as
    (
        select dc.ClientParentNameShort
            , dpr.Product
            , grp = 
                case dpr.Product
                    when 'Dental' then 'Dental'
                    when 'Workers Comp' then 'WC'
                    else 'Group Health'
                end
            , dd.DateMonthID 
            , dd.DateWeekID
            , dd.DateDay
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
            , dd.DateWeekID  
            , dd.DateDay
    )
    
select c.ClientParentNameShort
    , c.Product
    , c.grp
    , c.DateMonthID
    , c.DateWeekID
    , c.DateDay
    , c.Claims
    , c.Charges
    , lag(c.Claims, 1) over(partition by c.ClientParentNameShort, c.Product order by c.DateMonthID) Claims_prev
    , lag(c.Charges, 1) over(partition by c.ClientParentNameShort, c.Product order by c.DateMonthID) Charges_prev
    , c.Claims - (lag(c.Claims, 1) over(partition by c.ClientParentNameShort, c.Product order by c.DateMonthID)) as diff_claims
    , c.Charges - (lag(c.Charges, 1) over(partition by c.ClientParentNameShort, c.Product order by c.DateMonthID)) as diff_charges
    
from cte_daily c 
        
;
'''

df = pd.read_sql(sql, conn)
conn.close()

# Calc and split
df['Claims_MTD'] = df['Claims']*mtd
df['Charges_MTD'] = df['Charges']*mtd
df['pDiff_Claims'] = df['Claims_MTD'] / df['Claims_prev'] - 1
df['pDiff_Charges'] = df['Charges_MTD'] / df['Charges_prev'] - 1
df_curr = df[df['DateMonthID'] == currMonthID]

df_flag = df[(df.pDiff_Charges >= .25) | (df.pDiff_Charges <= -.25)]

df_med = df_flag[(df_flag.Product == "Group Health") | (df_flag.Product == "Claim Settlement (PPN)")
                 | (df_flag.Product == "Medicare Pricing Solutions")]
df_dent = df_flag[df_flag.Product == "Dental"]
df_wc = df_flag[df_flag.Product == "Workers Comp"]

# Write results to excel
writer = pd.ExcelWriter(r'C:\Users\pallen\Documents\Volume_Check_test.xlsx')

df_flag.to_excel(writer, 'Sheet1')
df_med.to_excel(writer, 'Medical')
df_dent.to_excel(writer, 'Dental')
df_wc.to_excel(writer, 'WC')

writer.save()

# check
# total = len(df_flag)
# total == len(df_med) + len(df_dent) + len(df_wc)


dfw = df.groupby('DateWeekID').agg(np.sum)
dfg.set_index('DateWeekID')


dfg.plot(style='k', legend=False)
plt.show()

