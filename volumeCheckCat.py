"""
Client Service Category Vol/KPI Monitor

"""

import numpy as np
import pandas as pd
from pandas import TimeGrouper
import pyodbc
from pandas import ExcelWriter
from datetime import datetime as dt
import calendar
import matplotlib.pyplot as plt
from functools import lru_cache
import xlsxwriter
from xlsxwriter import workbook, worksheet

# import openpyxl
# from openpyxl import workbook, worksheet, load_workbook
# from openpyxl.styles import NamedStyle, Border, Side, PatternFill, Font, GradientFill, Alignment
# from openpyxl.compat import range

# Declarations
y = dt.now().year
m = dt.now().month
d = dt.now().day
currMonthID = y * 100 + m
currMonthDays = calendar.monthrange(y, m)
mtd = (currMonthDays[1] / d)


def SQLpull():
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
            and dc.ClientParentNameShort = 'Cigna East'
            and dst.CategoryDesc = 'Anesthesia'
        
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
    '''

    df = pd.read_sql(sql, conn)
    df = df.set_index('DateDay')
    conn.close()
    return df



def munge(df):
    df = df.set_index('DateDay')
    return df


def toExcel():

    dfKPI = df.loc[:, 'HitRate':]
    writer = pd.ExcelWriter(r'C:\Users\pallen\Documents\VolumeCheck.xlsx', engine='xlsxwriter')

    dfKPI.to_excel(writer, sheet_name='Anesthesia')

    workbook = writer.book
    worksheet = writer.sheets['Anesthesia']

    chart = workbook.add_chart({'type': 'line'})

    #     [sheetname, first_row, first_col, last_row, last_col]
    chart.add_series({'values': '=Anesthesia!$B$2:$D$31'})

    worksheet.insert_chart('F2', chart)

    writer.save()

    return

def toExcelopen():
    writer = pd.ExcelWriter(r'C:\Users\pallen\Documents\VolumeCheck.xlsx', engine='openpyxl')
    wb = load_workbook(writer)
    dfTest.to_excel(writer, 'Anesthesia', startrow=0, startcol=0)
    ws = wb.active

    c1 = LineChart()
    data = Reference(ws, min_col=2, min_row=1, max_col=2, max_row=rowCount)
    c1.add_data(data, titles_from_data=True)

    ws.add_chart(c1, "D2")
    writer.save()
    workbook.close()

    return

df1.pivot(index='DateDay', values='SaveP').plot(kind='bar')

def style():
    wb = load_workbook(r'C:\Users\pallen\Documents\VolumeCheck.xlsx')

    head = NamedStyle(name="head")
    thin = Side(border_style="thin", color="000000")
    head.border = Border(top=thin, left=thin, right=thin, bottom=thin)
    head.fill = PatternFill(fill_type="solid", start_color="662766", end_color="662766")
    head.font = Font(color="FFFFFF")
    head.al = Alignment(horizontal="center", vertical="center")

    ws = wb.get_sheet_by_name('Medical')

    for column in range(2, 15):
        ws.cell(row=2, column=column).style = head
        ws.column_dimensions.height = 20

    ws = wb.get_sheet_by_name('Dental')

    for column in range(2, 15):
        ws.cell(row=2, column=column).style = head

    ws = wb.get_sheet_by_name('WC')

    for column in range(2, 15):
        ws.cell(row=2, column=column).style = head

    wb.save(r'C:\Users\pallen\Documents\VolumeCheck.xlsx')

    return 0


def plot(client, category):


