import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode
import pandas as pd
import io
from datetime import datetime

def create_grid(df, cell_styles=None):
    gb = GridOptionsBuilder.from_dataframe(df)
    
    # Configure default column behavior
    gb.configure_default_column(
        resizable=True,
        filterable=True,
        sorteable=True,
        editable=False
    )
    
    # Add cell styling if provided
    if cell_styles:
        # Create JavaScript function for cell styling
        cell_style_jscode = JsCode("""
        function(params) {
            return {
                'backgroundColor': params.data._cellStyles ? params.data._cellStyles[params.column.colId] : null
            };
        }
        """)
        
        gb.configure_grid_options(
            getRowStyle=None,
            getCellStyle=cell_style_jscode
        )
    
    grid_options = gb.build()
    
    # Add custom cell styling configuration
    if cell_styles:
        grid_options['context'] = {'cell_styles': cell_styles}
    
    return AgGrid(
        df,
        gridOptions=grid_options,
        update_mode='MODEL_CHANGED',
        allow_unsafe_jscode=True,
        theme='streamlit',
        custom_css={
            ".ag-cell-added": {"backgroundColor": "#D4EDDA !important"},
            ".ag-cell-deleted": {"backgroundColor": "#F8D7DA !important"},
            ".ag-cell-modified": {"backgroundColor": "#FFF3CD !important"}
        }
    )

def display_shape_differences(shape_differences):
    """
    Display shape differences in a formatted way
    """
    for diff in shape_differences:
        if diff['type'] == 'added':
            st.markdown(f"🟢 **Added Shape:**")
            st.json(diff['shape'])
        elif diff['type'] == 'deleted':
            st.markdown(f"🔴 **Deleted Shape:**")
            st.json(diff['shape'])
        else:  # modified
            st.markdown(f"🟡 **Modified Shape:**")
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("Original:")
                st.json(diff['old_shape'])
            with col2:
                st.markdown("Modified:")
                st.json(diff['new_shape'])

def export_comparison(comparison_result):
    """
    Export comparison results including shape differences
    """
    output = io.BytesIO()
    
    # Create Excel writer object
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Write data differences
        comparison_result['df1'].to_excel(writer, sheet_name='File1', index=False)
        comparison_result['df2'].to_excel(writer, sheet_name='File2', index=False)
        comparison_result['diff_summary'].to_excel(writer, sheet_name='Data_Summary', index=False)
        
        # Write shape differences
        if 'shape_differences' in comparison_result:
            shape_diff_df = pd.DataFrame(comparison_result['shape_differences'])
            if not shape_diff_df.empty:
                shape_diff_df.to_excel(writer, sheet_name='Shape_Differences', index=False)
    
    # Prepare the file for download
    output.seek(0)
    
    # Create download button
    st.download_button(
        label="Download Comparison Report",
        data=output,
        file_name=f"comparison_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
