import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder
import pandas as pd
import io

def create_grid(df):
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(
        resizable=True,
        filterable=True,
        sorteable=True,
        editable=False
    )
    
    grid_options = gb.build()
    
    return AgGrid(
        df,
        gridOptions=grid_options,
        update_mode=GridUpdateMode.MODEL_CHANGED,
        allow_unsafe_jscode=True,
        theme='streamlit'
    )

def export_comparison(comparison_result):
    output = io.BytesIO()
    
    # Create Excel writer object
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        comparison_result['df1_styled'].to_excel(writer, sheet_name='File1', index=False)
        comparison_result['df2_styled'].to_excel(writer, sheet_name='File2', index=False)
        comparison_result['diff_summary'].to_excel(writer, sheet_name='Summary', index=False)
    
    # Prepare the file for download
    output.seek(0)
    
    # Create download button
    st.download_button(
        label="Download Comparison Report",
        data=output,
        file_name="comparison_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
