import streamlit as st
import pandas as pd
from collections import defaultdict
import os
import tempfile

def convert_excel_to_dict(sheet_path, index):
    data = pd.read_excel(sheet_path, sheet_name=index, skiprows=6)
    data = data.drop(columns=['Unnamed: 0', 'Base Year', 'Scale']).set_index('Country')
    data_dict = data.to_dict(orient='index')
    return data_dict

def merge_dict(gdp_dict, fx_dict, euro_countries):
    output_res = defaultdict(dict)
    err_lst = []
    for country in gdp_dict:
        for year in gdp_dict[country]:
            try:
                fx_rate = fx_dict["Euro Area"][year] if country in euro_countries else fx_dict[country][year]
                output_res[country][year] = float(gdp_dict[country][year]) / float(fx_rate)
            except:
                message = ''
                if country not in fx_dict or year not in fx_dict[country] or fx_dict[country][year] == '...':
                    message = "Country currency information missing"
                if year not in gdp_dict[country] or gdp_dict[country][year] == '...':
                    message = 'Invalid GDP data'
                err_lst.append((country, year, message))
                output_res[country][year] = None
    output_df = pd.DataFrame.from_dict(output_res, orient='index')
    output_df.index.name = 'Country'  # Set the index name to "Country"
    return output_df, err_lst

def process_excel_file(excel_path, euro_countries):
    gdp_data = convert_excel_to_dict(excel_path, 0)
    fx_data = convert_excel_to_dict(excel_path, 1)
    output_df, err = merge_dict(gdp_data, fx_data, euro_countries)
    output_df.reset_index(inplace=True)  # Reset the index to include the country names as a column
    base_name = os.path.basename(excel_path)
    output_file = os.path.splitext(base_name)[0] + "_output.xlsx"
    
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        output_df.to_excel(writer, index=False)  # Save the DataFrame without the index
        worksheet = writer.sheets['Sheet1']
        worksheet.set_column(0, 0, 40)  # Set the width of the first column
    return output_df, output_file, err

# Streamlit App
st.title("Excel Processor")

# File Upload
uploaded_file = st.file_uploader("Excel File", type=['xlsx'])
if uploaded_file is not None:
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
        tmp_file.write(uploaded_file.read())
        tmp_file_path = tmp_file.name

    gdp_data = convert_excel_to_dict(tmp_file_path, 0)
    countries = list(gdp_data.keys())
    DEFAULT_EURO = ['France', 'Ukraine', 'Sweden', 'Germany', 'Finland', 'Poland, Rep. of', 'Belgium', 'Greece', 'Italy', 'Ireland', 'Netherlands, The', 'Ireland']
    euro_countries = st.multiselect("Select countries using Euro Union currency", options=countries, default=DEFAULT_EURO)
    
    if st.button("Process"):
        output_df, output_file, err = process_excel_file(tmp_file_path, euro_countries)
        st.dataframe(output_df)
        st.download_button(label="Download Output CSV", data=output_df.to_csv(index=False), file_name=output_file, mime='text/csv')
        if err:
            st.dataframe(pd.DataFrame(err, columns=['Country', 'Year', 'Error']))

