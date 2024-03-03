import pandas as pd
import streamlit as st
import plotly.express as px
from PIL import Image

st.set_page_config(page_title='Excel Pivot & Dashboarding')
st.header('Excel Pivot & Dashboarding')
st.subheader('Data')

#load data
excel_file = 'Sales Data 2023 ver.1.xlsx'
sheet_name = 'Data'

df_data = pd.read_excel(excel_file,
                   sheet_name=sheet_name,
                   usecols='A:K',
                   header=0)
st.dataframe(df_data)
#load Pivot
st.subheader('Pivot')
excel_file = 'Sales Data 2023 ver.1.xlsx'
sheet_name1 = 'Pivot-2'

df_tSalesCusProfit = pd.read_excel(excel_file,
                   sheet_name=sheet_name1,
                   usecols='A:F',
                   header=0)
st.dataframe(df_tSalesCusProfit)


df_tmonthSales = pd.read_excel(excel_file,
                   sheet_name=sheet_name1,
                   usecols='H:J',
                   header=1)
st.dataframe(df_tmonthSales)


df_tmonthProfit = pd.read_excel(excel_file,
                   sheet_name=sheet_name1,
                   usecols='L:M',
                   header=1)
st.dataframe(df_tmonthProfit)


df_tmonthSalesVar = pd.read_excel(excel_file,
                   sheet_name=sheet_name1,
                   usecols='O:Q',
                   header=1)
st.dataframe(df_tmonthSalesVar)


df_tregionCusProfit = pd.read_excel(excel_file,
                   sheet_name=sheet_name1,
                   usecols='S:U',
                   header=1)
st.dataframe(df_tregionCusProfit)

df_tSalesCompl = pd.read_excel(excel_file,
                   sheet_name=sheet_name1,
                   usecols='W:X',
                   header=0)
st.dataframe(df_tSalesCompl)

df_tProfitCompl = pd.read_excel(excel_file,
                   sheet_name=sheet_name1,
                   usecols='Y:Z',
                   header=0)
st.dataframe(df_tProfitCompl)

df_tCustCompl = pd.read_excel(excel_file,
                   sheet_name=sheet_name1,
                   usecols='AA:AB',
                   header=0)
st.dataframe(df_tCustCompl)

# --- PLOT PIE CHART
st.subheader('Chart')
pie_salesCompl = px.pie(df_tSalesCompl,
                title='Sales Completion Rate',
                values= 'Percentage',
                names='SALES COMPLETION')

st.plotly_chart(pie_salesCompl)

# --- PLOT BAR CHART
bar_chart = px.bar(df_tmonthSales,
                   x='Months',
                   #y='Sum of Target Sales',
                   y='Sum of Sales (Units).1',
                   #text='Votes',
                   color_discrete_sequence = ['#F63366']*len(df_tmonthSales),
                   template= 'plotly_white')
st.plotly_chart(bar_chart)
