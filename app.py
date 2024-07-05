import streamlit as st
import pandas as pd
import openpyxl
import io
from io import BytesIO
import xlsxwriter

# Title of the app
st.title("Site Visit Sheduled Cleaning App")

# File uploader widget
uploaded_file = st.file_uploader("Choose Site Visit Scheduled Base File", type="xlsx")
uploaded_file_1 = st.file_uploader("Choose Active Employee File", type="xlsx")

if uploaded_file is not None:
    # Read the uploaded file into a DataFrame
    df = pd.read_excel(uploaded_file)
    df1 = pd.read_excel(uploaded_file_1)
    df = pd.merge(df,df1,on='Followed By',how='left')
    df = df.dropna(subset='City')
    df = df[~df['City'].isin(['Lucknow','Pune East','Pune West','Mumbai','lucknow'])]
    df = df.drop_duplicates(subset=['Mobile'])
    df['Visit Date'] = df['Visit Date'].str[:11]
    df['Visit Date'] = pd.to_datetime(df['Visit Date'])

    df['Sub-Project'] = 'Other'

    df.loc[df['Project'] == 'Gaur Aero Mall','Sub-Project'] = 'Gaur Hindon'
    df.loc[df['Project'] == 'Gaur Airocity','Sub-Project'] = 'Gaur Hindon'
    df.loc[df['Project'] == 'Gaur Aero Suites','Sub-Project'] = 'Gaur Hindon'
    df.loc[df['Project'] == 'Gaur Aero Heights','Sub-Project'] = 'Gaur Hindon'
    df.loc[df['Project'] == 'Gaur Aero Mall,Gaur Aero Suites','Sub-Project'] = 'Gaur Hindon'



    df.loc[df['Project'] == 'Sikka Kaamna Greens','Sub-Project'] = 'Sikka Projects'

    df.loc[df['Project'] == 'MIGSUN DELTA 2','Sub-Project'] = 'Migsun Delta'
    df.loc[df['Project'] == 'MIGSUN DELTA 1','Sub-Project'] = 'Migsun Delta'


    df.loc[df['Project'] == 'GYC Galleria','Sub-Project'] = 'Gaur Yamuna City'
    df.loc[df['Project'] == 'Gaur Runway Hub','Sub-Project'] = 'Gaur Yamuna City'
    df.loc[df['Project'] == 'Gaur Runway Suites','Sub-Project'] = 'Gaur Yamuna City'
    df.loc[df['Project'] == 'Gaur Runway Hub and Food Court,Gaur Runway Suites','Sub-Project'] = 'Gaur Yamuna City'
    df.loc[df['Project'] == 'Gaur Runway Hub and Food Court','Sub-Project'] = 'Gaur Yamuna City'


    df.loc[df['Project'] == 'Bhutani Avenue 133','Sub-Project'] = 'Bhutani Avenue 133'

    df.loc[df['Project'] == 'Migsun JetSuites','Sub-Project'] = 'Migsun Jetsuites'




    df = df.filter(['City','HOD','Project','Sub-Project','Visit Date'])
    df = df.pivot_table(index='HOD',columns='Sub-Project',aggfunc='count',margins=True)
    df = df.fillna(0)
    df = df.rename(index={'All':'Grand Total'})

    # Display the cleaned DataFrame
    st.subheader("Cleaned Data")
    st.write(df)

    @st.cache_data
    def convert_df_to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=True, sheet_name='Sheet1')
            writer._save()
        processed_data = output.getvalue()
        return processed_data

    excel_data = convert_df_to_excel(df)

    st.download_button(
        label="Download Cleaned Data as Excel",
        data=excel_data,
        file_name='cleaned_data.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
