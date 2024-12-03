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
                    st.subheader("シートの選択")
                    selected_sheets1 = st.multiselect("ファイル1のシートを選択", sheets1, default=[sheets1[0]])
                    selected_sheets2 = st.multiselect("ファイル2のシートを選択", sheets2, default=[sheets2[0]])
                    
                    if len(selected_sheets1) != len(selected_sheets2):
                        st.error("選択されたシートの数が一致しません。")
                        return
                    
                    if not selected_sheets1 or not selected_sheets2:
                        st.warning("比較するシートを選択してください。")
                        return

                    # 全シートの比較結果を格納する配列
                    all_comparison_results = []
                    
                    # 各シートペアを比較
                    for sheet1, sheet2 in zip(selected_sheets1, selected_sheets2):
                        st.subheader(f"シートの比較: {sheet1} vs {sheet2}")
                        
                        try:
                            # Reset file pointers for reading Excel
                            file1.seek(0)
                            file2.seek(0)
                            
                            # Load sheet data
                            df1 = pd.read_excel(file1, sheet_name=sheet1)
                            df2 = pd.read_excel(file2, sheet_name=sheet2)
                            
                            if df1.empty or df2.empty:
                                st.warning(f"シート '{sheet1}' または '{sheet2}' にデータが存在しません。")
                                continue
                            
                            # Compare shapes
                            st.info(f"シート '{sheet1}' と '{sheet2}' の画像比較を開始...")
                            shapes1 = comparison.extract_shape_info(file1_path, sheet1)
                            st.write(f"ファイル1の画像数: {len(shapes1)}")
                            shapes2 = comparison.extract_shape_info(file2_path, sheet2)
                            st.write(f"ファイル2の画像数: {len(shapes2)}")
                            shape_differences = comparison.compare_shapes(shapes1, shapes2)
                            
                            # Compare data
                            comparison_result = comparison.compare_dataframes(df1, df2)
                            if not isinstance(comparison_result, dict):
                                st.error("比較結果の形式が不正です。")
                                continue
                            
                            # Add sheet names and shape differences to results
                            comparison_result['sheet1_name'] = sheet1
                            comparison_result['sheet2_name'] = sheet2
                            comparison_result['shape_differences'] = shape_differences
                            
                            all_comparison_results.append(comparison_result)
                            
                            # Display individual sheet comparison results
                            col1, col2 = st.columns(2)
                            
                            with col1:
                                st.markdown(f"### ファイル 1 - {sheet1}")
                                if 'df1' in comparison_result and not comparison_result['df1'].empty:
                                    grid1 = utils.create_grid(comparison_result['df1'], comparison_result.get('df1_styles', None))
                                else:
                                    st.warning("データを表示できません。")
                            
                            with col2:
                                st.markdown(f"### ファイル 2 - {sheet2}")
                                if 'df2' in comparison_result and not comparison_result['df2'].empty:
                                    grid2 = utils.create_grid(comparison_result['df2'], comparison_result.get('df2_styles', None))
                                else:
                                    st.warning("データを表示できません。")
                            
                            # Display shape differences for this sheet pair
                            if shape_differences:
                                st.subheader(f"図形の差分 ({sheet1} vs {sheet2})")
                                utils.display_shape_differences(shape_differences)
                            
                        except Exception as e:
                            st.error(f"シート '{sheet1}' と '{sheet2}' の比較中にエラーが発生しました: {str(e)}")
                            continue
                    
                    # 全シートの比較が完了した後、サマリーを表示
                    if all_comparison_results:
                        st.markdown("---")
                        st.subheader("全体の比較結果サマリー")
                        
                        # Create combined summary DataFrame
                        all_summary_data = []
                        
                        for result in all_comparison_results:
                            sheet1_name = result['sheet1_name']
                            sheet2_name = result['sheet2_name']
                            
                            if 'diff_summary' in result:
                                for diff in result['diff_summary'].to_dict('records'):
                                    if diff['type'] == 'modified':
                                        df1 = result['df1']
                                        col_idx = df1.columns.get_loc(diff['column'])
                                        cell_ref_old = utils.get_excel_cell_reference(col_idx, diff['row_index_old'])
                                        cell_ref_new = utils.get_excel_cell_reference(col_idx, diff['row_index_new'])
                                        
                                        all_summary_data.append({
                                            'シート名': f"{sheet1_name} → {sheet2_name}",
                                            '変更タイプ': 'データ変更',
                                            'セル位置 (変更前)': cell_ref_old,
                                            'セル位置 (変更後)': cell_ref_new,
                                            '変更前の値': diff['value_old'],
                                            '変更後の値': diff['value_new']
                                        })
                                    else:
                                        df = result['df1']
                                        row_idx = diff['row_index']
                                        range_ref = utils.get_excel_range_reference(row_idx, 0, len(df.columns) - 1)
                                        
                                        current_sheet = sheet2_name if diff['type'] == 'added' else sheet1_name
                                        all_summary_data.append({
                                            'シート名': current_sheet,
                                            '変更タイプ': 'データ追加' if diff['type'] == 'added' else 'データ削除',
                                            'セル位置': f"{row_idx + 1}行目 ({range_ref})",
                                            '値': ' | '.join([f"{k}: {v}" for k, v in diff['values'].items() if pd.notna(v)])
                                        })
                        
                        if all_summary_data:
                            summary_df = pd.DataFrame(all_summary_data)
                            st.dataframe(summary_df)
                        else:
                            st.info("全シートで差分は検出されませんでした")
                        
                        # Export options for all sheets
                        st.markdown("---")
                        st.subheader("エクスポート")
                        for result in all_comparison_results:
                            utils.export_comparison(result, result['sheet1_name'], result['sheet2_name'])
                    
                except Exception as e:
                    st.error(f"比較処理中にエラーが発生しました: {str(e)}")
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
