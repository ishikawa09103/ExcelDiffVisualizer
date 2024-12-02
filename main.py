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
    st.title("Excel ファイル比較ツール")
    
    # File upload section
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("ファイル 1")
        file1 = st.file_uploader("1つ目のExcelファイルをアップロード", type=['xlsx', 'xls'])
    
    with col2:
        st.subheader("ファイル 2")
        file2 = st.file_uploader("2つ目のExcelファイルをアップロード", type=['xlsx', 'xls'])

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
            
            # Display comparison summary
            st.subheader("比較結果サマリー")

            # Create summary DataFrame
            summary_data = []
            for diff in comparison_result['diff_summary'].to_dict('records'):
                if diff['type'] == 'modified':
                    summary_data.append({
                        '変更タイプ': '変更',
                        '列': diff['column'],
                        '行 (変更前)': diff['row_index_old'],
                        '行 (変更後)': diff['row_index_new'],
                        '変更前の値': diff['value_old'],
                        '変更後の値': diff['value_new']
                    })
                else:
                    values = diff['values']
                    for col, val in values.items():
                        summary_data.append({
                            '変更タイプ': '追加' if diff['type'] == 'added' else '削除',
                            '列': col,
                            '行': diff['row_index'],
                            '値': val
                        })

            if summary_data:
                summary_df = pd.DataFrame(summary_data)
                st.dataframe(
                    summary_df.style.apply(lambda x: ['background-color: #FFF3CD' if v == '変更'
                                                    else 'background-color: #D4EDDA' if v == '追加'
                                                    else 'background-color: #F8D7DA' if v == '削除'
                                                    else '' for v in x],
                                         subset=['変更タイプ'])
                )
            else:
                st.info("差分は検出されませんでした")

            # Export options
            st.markdown("---")
            st.subheader("エクスポート")
            utils.export_comparison(comparison_result)
            
        except Exception as e:
            st.error(f"ファイルの処理中にエラーが発生しました: {str(e)}")
    
    else:
        st.info("比較を開始するには、両方のExcelファイルをアップロードしてください")

    # Add legend
    st.sidebar.markdown("### 凡例")
    st.sidebar.markdown("""
    - 🟢 追加されたセル/図形 (緑色)
    - 🔴 削除されたセル/図形 (赤色)
    - 🟡 変更されたセル/図形 (黄色)
    """)

if __name__ == "__main__":
    main()
