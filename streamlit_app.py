import streamlit as st
import pandas as pd
import os
import io
import func

with st.sidebar:
    st.header('Cara penggunaan:')
    st.write('1. Upload file')
    st.write('2. Upload file')

st.header('GSE Inventory Issued Script')

if "reset" not in st.session_state:
    st.session_state.reset = False

def reset_page():
    # Reset all session state variables
    st.session_state.clear()

# Create a button to trigger the reset
if st.button('RESET'):
    reset_page()  # Call the reset function when the button is clicked
    st.session_state.reset = True  # Mark that the page has been reset
    st.rerun()

dataOrder = st.file_uploader('Data Order Detail', type=['csv'])

st.write('---')

dataRaw_1 = st.file_uploader('Data Issued 1', type='xls')
dataRaw_2 = st.file_uploader('Data Issued 2', type='xls')
dataRaw_3 = st.file_uploader('Data Issued 3', type='xls')
dataRaw_4 = st.file_uploader('Data Issued 4', type='xls')


if dataOrder and dataRaw_1 and dataRaw_2 and dataRaw_3 and dataRaw_4:
    if st.button("Olah Data!", type="primary"):     
        st.session_state.reset = False   
        try:
            dataRaw, dataProcessed, oldestDate, newestDate = func.readProcess_Order(dataOrder, dataRaw_1, dataRaw_2, dataRaw_3, dataRaw_4)

            output = io.BytesIO()

            with pd.ExcelWriter(output, date_format='m/d/yyyy', datetime_format='m/d/yyyy HH:MM:SS', engine='xlsxwriter') as writer:
                dataRaw.to_excel(writer, sheet_name='DATA_RAW', index=False)
                dataProcessed.to_excel(writer, sheet_name='DATA_PROCESSED', index=False)

            output.seek(0)

            st.session_state['oldestDate'] = oldestDate
            st.session_state['newestDate'] = newestDate
            st.session_state['processed_file'] = output
            st.session_state['dataRaw'] = dataRaw
            st.session_state['dataProcessed'] = dataProcessed

        except Exception as e:
            st.error(f"bruh {e}")

    if 'processed_file' in st.session_state:
        st.download_button(
            label="Download Data Raw & Processed",
            data=st.session_state['processed_file'],
            file_name="DATA_%s_%s.xlsx" %(st.session_state['oldestDate'],st.session_state['newestDate']),
            mime="text/csv"
        )

    
if st.session_state.reset:
    st.write("The page has been reset. Remove old files and upload a new file to begin.")
else:
    st.write("Upload a file and interact with the app.")