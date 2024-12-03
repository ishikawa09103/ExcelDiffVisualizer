import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder
import comparison
import utils
import styles
from openpyxl import load_workbook
import tempfile
import os

st.set_page_config(
    page_title="Excel Comparison Tool",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Apply custom CSS
styles.apply_custom_css()

def main():
    try:
        st.title("Excel ファイル比較ツール")
        
        # Initialize state management
        if 'upload_error' not in st.session_state:
            st.session_state.upload_error = None
        if 'comparison_error' not in st.session_state:
            st.session_state.comparison_error = None
        
        # File upload section with error handling
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("ファイル 1")
            try:
                file1 = st.file_uploader(
                    "1つ目のExcelファイルをアップロード",
                    type=['xlsx', 'xls'],
                    key="file1_uploader",
                    on_change=lambda: setattr(st.session_state, 'upload_error', None)
                )
            except Exception as e:
                st.error(f"ファイル1のアップロード中にエラーが発生しました: {str(e)}")
                file1 = None
        
        with col2:
            st.subheader("ファイル 2")
            try:
                file2 = st.file_uploader(
                    "2つ目のExcelファイルをアップロード",
                    type=['xlsx', 'xls'],
                    key="file2_uploader",
                    on_change=lambda: setattr(st.session_state, 'upload_error', None)
                )
            except Exception as e:
                st.error(f"ファイル2のアップロード中にエラーが発生しました: {str(e)}")
                file2 = None

        # Display any previous upload errors
        if st.session_state.upload_error:
            st.error(st.session_state.upload_error)
            st.session_state.upload_error = None

        if file1 and file2:
            try:
                # Save uploaded files to temporary locations
                try:
                    temp_dir = tempfile.mkdtemp()
                    file1_path = os.path.join(temp_dir, "file1.xlsx")
                    file2_path = os.path.join(temp_dir, "file2.xlsx")
                    
                    with open(file1_path, 'wb') as f:
                        f.write(file1.getvalue())
                    with open(file2_path, 'wb') as f:
                        f.write(file2.getvalue())
                    
                    # Get workbook information using openpyxl
                    try:
                        wb1 = load_workbook(file1_path)
                        wb2 = load_workbook(file2_path)
                        
                        # シート名の取得
                        sheets1 = wb1.sheetnames
                        sheets2 = wb2.sheetnames
                        
                        # クリーンアップ
                        wb1.close()
                        wb2.close()
                    except Exception as e:
                        st.error(f"シート情報の取得中にエラーが発生しました: {str(e)}")
                        return
                    
                    if not sheets1 or not sheets2:
                        st.error("有効なシートが見つかりません。")
                        return

                    # Debug output for sheet information
                    st.write(f"ファイル1のシート数: {len(sheets1)}")
                    st.write(f"ファイル1のシート名: {', '.join(sheets1)}")
                    st.write(f"ファイル2のシート数: {len(sheets2)}")
                    st.write(f"ファイル2のシート名: {', '.join(sheets2)}")
                    
                    # Sheet selection
                    col1, col2 = st.columns(2)
                    with col1:
                        sheet1 = st.selectbox("ファイル1のシートを選択", sheets1)
                    with col2:
                        sheet2 = st.selectbox("ファイル2のシートを選択", sheets2)
                    
                    # Reset file pointers for reading Excel
                    file1.seek(0)
                    file2.seek(0)
                    
                    # Load and compare sheets
                    try:
                        df1 = pd.read_excel(file1, sheet_name=sheet1)
                        df2 = pd.read_excel(file2, sheet_name=sheet2)
                        
                        if df1.empty or df2.empty:
                            st.warning("選択されたシートにデータが存在しません。")
                            return
                            
                        st.info("画像の比較を開始...")
                        shapes1 = comparison.extract_shape_info(file1_path, sheet1)
                        st.write(f"ファイル1の画像数: {len(shapes1)}")
                        shapes2 = comparison.extract_shape_info(file2_path, sheet2)
                        st.write(f"ファイル2の画像数: {len(shapes2)}")
                        shape_differences = comparison.compare_shapes(shapes1, shapes2)
                        
                        # Compare dataframes and process results
                        try:
                            comparison_result = comparison.compare_dataframes(df1, df2)
                            if not isinstance(comparison_result, dict):
                                st.error("比較結果の形式が不正です。")
                                return
                            
                            # Add shape differences to results
                            comparison_result['shape_differences'] = shape_differences
                        except Exception as e:
                            st.error(f"データの比較中にエラーが発生しました: {str(e)}")
                            return
                        
                        # Display comparison results with error handling
                        st.subheader("データ比較")
                        
                        col1, col2 = st.columns(2)
                        
                        try:
                            with col1:
                                st.markdown("### ファイル 1")
                                if 'df1' in comparison_result and not comparison_result['df1'].empty:
                                    grid1 = utils.create_grid(comparison_result['df1'], comparison_result.get('df1_styles', None))
                                else:
                                    st.warning("ファイル1のデータを表示できません。")
                            
                            with col2:
                                st.markdown("### ファイル 2")
                                if 'df2' in comparison_result and not comparison_result['df2'].empty:
                                    grid2 = utils.create_grid(comparison_result['df2'], comparison_result.get('df2_styles', None))
                                else:
                                    st.warning("ファイル2のデータを表示できません。")
                        except Exception as e:
                            st.error(f"データの表示中にエラーが発生しました: {str(e)}")
                            return
                        
                        # Display shape differences
                        if shape_differences:
                            st.subheader("図形の差分")
                            utils.display_shape_differences(shape_differences)
                        
                        # Display comparison summary
                        st.subheader("比較結果サマリー")
                        
                        # Create summary DataFrame
                        summary_data = []
                        
                        # Add data differences
                        if 'diff_summary' in comparison_result:
                            for diff in comparison_result['diff_summary'].to_dict('records'):
                                if diff['type'] == 'modified':
                                    col_idx = df1.columns.get_loc(diff['column'])
                                    cell_ref_old = utils.get_excel_cell_reference(col_idx, diff['row_index_old'])
                                    cell_ref_new = utils.get_excel_cell_reference(col_idx, diff['row_index_new'])
                                    summary_data.append({
                                        'ブック': 'ファイル1 → ファイル2',
                                        '変更タイプ': 'データ変更',
                                        'セル位置 (変更前)': cell_ref_old,
                                        'セル位置 (変更後)': cell_ref_new,
                                        '変更前の値': diff['value_old'],
                                        '変更後の値': diff['value_new']
                                    })
                                else:
                                    row_idx = diff['row_index']
                                    range_ref = utils.get_excel_range_reference(row_idx, 0, len(df1.columns) - 1)
                                    row_values = []
                                    for col in df1.columns:
                                        val = diff['values'].get(col, '')
                                        if pd.notna(val):
                                            row_values.append(f"{col}: {val}")
                                    
                                    current_book = 'ファイル2' if diff['type'] == 'added' else 'ファイル1'
                                    summary_data.append({
                                        'ブック': current_book,
                                        '変更タイプ': 'データ追加' if diff['type'] == 'added' else 'データ削除',
                                        'セル位置': f"{row_idx + 1}行目 ({range_ref})",
                                        '値': ' | '.join(row_values)
                                    })
                        
                        if summary_data:
                            summary_df = pd.DataFrame(summary_data)
                            st.dataframe(summary_df)
                        else:
                            st.info("差分は検出されませんでした")
                        
                        # Export options
                        st.markdown("---")
                        st.subheader("エクスポート")
                        utils.export_comparison(comparison_result, sheet1, sheet2)
                        
                    except Exception as e:
                        st.error(f"シートの処理中にエラーが発生しました: {str(e)}")
                        return
                        
                except Exception as e:
                    st.error(f"シート情報の取得中にエラーが発生しました: {str(e)}")
                    return
                    
            except Exception as e:
                st.error(f"ファイルの処理中にエラーが発生しました: {str(e)}")
                return
        else:
            st.info("比較を開始するには、両方のExcelファイルをアップロードしてください")
    
    except Exception as e:
        st.error(f"予期せぬエラーが発生しました: {str(e)}")

if __name__ == "__main__":
    main()