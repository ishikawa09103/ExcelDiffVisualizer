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
            # Load workbooks
            wb1 = load_workbook(file1)
            wb2 = load_workbook(file2)
            
            # Get sheet names
            sheets1 = wb1.sheetnames
            sheets2 = wb2.sheetnames

            # シート情報のデバッグ出力を追加
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
            
            # Read selected sheets
            df1 = pd.read_excel(file1, sheet_name=sheet1)
            df2 = pd.read_excel(file2, sheet_name=sheet2)
            
            # Extract and compare shapes for selected sheets
            st.write("画像の比較を開始...")
            shapes1 = comparison.extract_shape_info(wb1, sheet1)
            st.write(f"ファイル1の画像数: {len(shapes1)}")
            shapes2 = comparison.extract_shape_info(wb2, sheet2)
            st.write(f"ファイル2の画像数: {len(shapes2)}")
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
            
            # データの差分を追加
            for diff in comparison_result['diff_summary'].to_dict('records'):
                if diff['type'] == 'modified':
                    col_idx = df1.columns.get_loc(diff['column'])
                    cell_ref_old = utils.get_excel_cell_reference(col_idx, diff['row_index_old'])
                    cell_ref_new = utils.get_excel_cell_reference(col_idx, diff['row_index_new'])
                    summary_data.append({
                        'シート名': f"{sheet1} → {sheet2}",
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
                    
                    current_sheet = sheet2 if diff['type'] == 'added' else sheet1
                    summary_data.append({
                        'シート名': current_sheet,
                        '変更タイプ': 'データ追加' if diff['type'] == 'added' else 'データ削除',
                        'セル位置': f"{row_idx + 1}行目 ({range_ref})",
                        '値': ' | '.join(row_values)
                    })
            
            def get_size_info(width, height):
                try:
                    if width is None or height is None:
                        return "サイズ情報なし"
                    w = float(width) if width is not None else 0
                    h = float(height) if height is not None else 0
                    return f"サイズ: 幅 {w:.1f}px, 高さ {h:.1f}px"
                except (TypeError, ValueError):
                    return "サイズ情報なし"

            # 画像の差分を追加
            for diff in shape_differences:
                if diff['type'] == 'added':
                    shape = diff.get('shape', {})
                    if shape.get('type') == 'image':
                        cell_ref = utils.get_excel_cell_reference(
                            shape.get('x', 0),
                            shape.get('y', 0)
                        )
                        size_info = get_size_info(
                            shape.get('width'),
                            shape.get('height')
                        )
                        summary_data.append({
                            'シート名': sheet2,
                            '変更タイプ': '画像追加',
                            'セル位置': cell_ref,
                            '値': size_info
                        })
                elif diff['type'] == 'deleted':
                    shape = diff.get('shape', {})
                    if shape.get('type') == 'image':
                        cell_ref = utils.get_excel_cell_reference(
                            shape.get('x', 0),
                            shape.get('y', 0)
                        )
                        size_info = get_size_info(
                            shape.get('width'),
                            shape.get('height')
                        )
                        summary_data.append({
                            'シート名': sheet1,
                            '変更タイプ': '画像削除',
                            'セル位置': cell_ref,
                            '値': size_info
                        })
                else:  # modified
                    old_shape = diff.get('old_shape', {})
                    new_shape = diff.get('new_shape', {})
                    if old_shape.get('type') == 'image' and new_shape.get('type') == 'image':
                        cell_ref_old = utils.get_excel_cell_reference(
                            old_shape.get('x', 0),
                            old_shape.get('y', 0)
                        )
                        cell_ref_new = utils.get_excel_cell_reference(
                            new_shape.get('x', 0),
                            new_shape.get('y', 0)
                        )
                        
                        old_size_info = get_size_info(
                            old_shape.get('width'),
                            old_shape.get('height')
                        )
                        new_size_info = get_size_info(
                            new_shape.get('width'),
                            new_shape.get('height')
                        )
                        
                        summary_data.append({
                            'シート名': f"{sheet1} → {sheet2}",
                            '変更タイプ': '画像変更',
                            'セル位置 (変更前)': cell_ref_old,
                            'セル位置 (変更後)': cell_ref_new,
                            '変更前の値': old_size_info,
                            '変更後の値': new_size_info
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
            utils.export_comparison(comparison_result, sheet1, sheet2)
            
        except Exception as e:
            st.error(f"ファイルの処理中にエラーが発生しました: {str(e)}")
    
    else:
        st.info("比較を開始するには、両方のExcelファイルをアップロードしてください")

if __name__ == "__main__":
    main()
