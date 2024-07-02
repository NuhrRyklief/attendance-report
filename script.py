import pandas as pd
import streamlit as st

# Streamlit widgets for user input
st.title('Attendance Report Processor')

file_path = st.text_input('Enter the file path of the Excel file:', 'attendance_test.xlsx')
sheet_name = st.text_input('Enter the sheet name:', 'in')

def process_attendance(file_path, sheet_name):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Find the index of the first row mentioning "Attendee Details"
        attendee_details_row = df[df.iloc[:, 0] == 'Attendee Details'].index[0]
        
        # Create the "out" sheet by deleting rows up to and including the "Attendee Details" row
        df_out = df.iloc[attendee_details_row + 1:].reset_index(drop=True)
        df_out.columns = df_out.iloc[0]
        df_out = df_out[1:]
        df_out.columns.name = None
        
        # Save the "out" sheet
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
            df_out.to_excel(writer, sheet_name='out', index=False)
        
        # Create the "yes" and "no" sheets
        df_yes = df_out[df_out['Attended'] == 'Yes']
        df_no = df_out[df_out['Attended'] == 'No']
        
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
            df_yes.to_excel(writer, sheet_name='yes', index=False)
            df_no.to_excel(writer, sheet_name='no', index=False)
        
        # Create the "yes_tt" sheet with total_time calculation
        df_yes['Join Time'] = pd.to_datetime(df_yes['Join Time'])
        df_yes['Leave Time'] = pd.to_datetime(df_yes['Leave Time'])
        
        # Group by Email and calculate the total time attended for each attendee
        total_time_df = df_yes.groupby('Email').apply(
            lambda x: (x['Leave Time'].max() - x['Join Time'].min())
        ).reset_index(name='Total Time')
        
        # Merge the total time back with the original df_yes
        df_yes_tt = pd.merge(df_yes, total_time_df, on='Email')
        
        # Convert 'Total Time' to string format
        df_yes_tt['Total Time'] = df_yes_tt['Total Time'].apply(lambda x: str(x))
        
        # Reorder columns to place 'Total Time' immediately after 'Leave Time'
        leave_time_index = df_yes_tt.columns.get_loc('Leave Time')
        cols = list(df_yes_tt.columns)
        cols.insert(leave_time_index + 1, cols.pop(cols.index('Total Time')))
        df_yes_tt = df_yes_tt[cols]
        
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
            df_yes_tt.to_excel(writer, sheet_name='yes_tt', index=False)
        
        # Create the "yes_tt_cleaned" sheet by removing duplicates and retaining total time
        df_yes_tt_cleaned = df_yes_tt.drop_duplicates(subset=['Email'], keep='first')
        
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
            df_yes_tt_cleaned.to_excel(writer, sheet_name='yes_tt_cleaned', index=False)
        
        st.success("Processing complete. Sheets 'out', 'yes', 'no', 'yes_tt', and 'yes_tt_cleaned' created.")
        
    except Exception as e:
        st.error(f"An error occurred: {e}")

if st.button('Process Attendance'):
    process_attendance(file_path, sheet_name)
