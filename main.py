import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder
import comparison
import utils
import styles
from openpyxl import load_workbook

st.set_page_config(
    page_title="Excel Comparison Tool",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Apply custom CSS
styles.apply_custom_css()

def main():
    st.title("Excel File Comparison Tool")
    
    # File upload section
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("File 1")
        file1 = st.file_uploader("Upload first Excel file", type=['xlsx', 'xls'])
    
    with col2:
        st.subheader("File 2")
        file2 = st.file_uploader("Upload second Excel file", type=['xlsx', 'xls'])

    if file1 and file2:
        try:
            # Load and process files
            df1 = pd.read_excel(file1)
            df2 = pd.read_excel(file2)
            
            # Reset file pointers for shape comparison
            file1.seek(0)
            file2.seek(0)
            
            # Load workbooks for shape comparison
            wb1 = load_workbook(file1)
            wb2 = load_workbook(file2)
            
            # Get active sheet names
            sheet1_name = wb1.active.title
            sheet2_name = wb2.active.title
            
            # Extract and compare shapes
            shapes1 = comparison.extract_shape_info(wb1, sheet1_name)
            shapes2 = comparison.extract_shape_info(wb2, sheet2_name)
            shape_differences = comparison.compare_shapes(shapes1, shapes2)
            
            # Compare dataframes
            comparison_result = comparison.compare_dataframes(df1, df2)
            
            # Add shape differences to the comparison result
            comparison_result['shape_differences'] = shape_differences
            
            # Display comparison results
            st.subheader("Data Comparison")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("### File 1")
                grid1 = utils.create_grid(comparison_result['df1'], comparison_result['df1_styles'])
                
            with col2:
                st.markdown("### File 2")
                grid2 = utils.create_grid(comparison_result['df2'], comparison_result['df2_styles'])
            
            # Display shape differences
            if shape_differences:
                st.subheader("Shape Differences")
                utils.display_shape_differences(shape_differences)
            
            # Export options
            st.markdown("---")
            st.subheader("Export Results")
            utils.export_comparison(comparison_result)
            
        except Exception as e:
            st.error(f"Error processing files: {str(e)}")
    
    else:
        st.info("Please upload both Excel files to start comparison")

    # Add legend
    st.sidebar.markdown("### Legend")
    st.sidebar.markdown("""
    - 🟢 Added cells/shapes (Green)
    - 🔴 Deleted cells/shapes (Red)
    - 🟡 Modified cells/shapes (Yellow)
    """)

if __name__ == "__main__":
    main()
