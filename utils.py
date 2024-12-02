import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode
import pandas as pd
import io
from datetime import datetime

def get_excel_cell_reference(column_index, row_index):
    """
    Convert 0-based column and row indices to Excel cell reference (e.g., A1, B2)
    """
    def get_column_letter(col_idx):
        result = ""
        while col_idx >= 0:
            result = chr(65 + (col_idx % 26)) + result
            col_idx = col_idx // 26 - 1
        return result
    
    return f"{get_column_letter(column_index)}{row_index + 1}"

def get_excel_range_reference(row_index, start_col_index, end_col_index):
    """
    Get Excel range reference for a row (e.g., A5:E5)
    """
    start_ref = get_excel_cell_reference(start_col_index, row_index)
    end_ref = get_excel_cell_reference(end_col_index, row_index)
    return f"{start_ref}:{end_ref}"

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
    Display shape differences in a formatted way with improved image information
    """
    st.write("画像の差分処理を開始...")
    
    for diff in shape_differences:
        st.write(f"処理中の差分タイプ: {diff['type']}")
        st.write(f"差分の内容: {diff}")
        
        if diff['type'] == 'added':
            shape = diff.get('shape', {})
            st.write(f"追加された形状の情報: {shape}")
            if shape.get('type') == 'image':
                try:
                    st.markdown(f"🟢 **追加された画像:**")
                    cell_ref = get_excel_cell_reference(shape.get('x', 0), shape.get('y', 0))
                    st.write(f"- 位置: セル {cell_ref}")
                    if shape.get('width') is not None and shape.get('height') is not None:
                        st.write(f"- サイズ: 幅 {shape['width']:.1f}px, 高さ {shape['height']:.1f}px")
                    else:
                        st.write("- サイズ情報なし")
                except Exception as e:
                    st.error(f"画像情報の表示中にエラー: {str(e)}")
        elif diff['type'] == 'deleted':
            shape = diff.get('shape', {})
            st.write(f"削除された形状の情報: {shape}")
            if shape.get('type') == 'image':
                try:
                    st.markdown(f"🔴 **削除された画像:**")
                    cell_ref = get_excel_cell_reference(shape.get('x', 0), shape.get('y', 0))
                    st.write(f"- 位置: セル {cell_ref}")
                    if shape.get('width') is not None and shape.get('height') is not None:
                        st.write(f"- サイズ: 幅 {shape['width']:.1f}px, 高さ {shape['height']:.1f}px")
                    else:
                        st.write("- サイズ情報なし")
                except Exception as e:
                    st.error(f"画像情報の表示中にエラー: {str(e)}")
            else:
                st.markdown(f"""
                - 種類: {shape.get('type', 'unknown')}
                - 位置: セル {get_excel_cell_reference(shape.get('x', 0), shape.get('y', 0))}
                - テキスト: {shape.get('text', '') or 'なし'}
                """)
        else:  # modified
            st.write("変更された形状の情報:")
            old_shape = diff.get('old_shape', {})
            new_shape = diff.get('new_shape', {})
            st.write(f"変更前: {old_shape}")
            st.write(f"変更後: {new_shape}")
            
            st.markdown(f"🟡 **変更された要素:**")
            col1, col2 = st.columns(2)
            with col1:
                try:
                    st.markdown("**変更前:**")
                    if old_shape.get('type') == 'image':
                        cell_ref = get_excel_cell_reference(old_shape.get('x', 0), old_shape.get('y', 0))
                        st.write(f"- 位置: セル {cell_ref}")
                        if old_shape.get('width') is not None and old_shape.get('height') is not None:
                            st.write(f"- サイズ: 幅 {old_shape['width']:.1f}px, 高さ {old_shape['height']:.1f}px")
                        else:
                            st.write("- サイズ情報なし")
                    else:
                        st.markdown(f"""
                        - 種類: {old_shape.get('type', 'unknown')}
                        - 位置: セル {get_excel_cell_reference(old_shape.get('x', 0), old_shape.get('y', 0))}
                        - テキスト: {old_shape.get('text', '') or 'なし'}
                        """)
                except Exception as e:
                    st.error(f"変更前の情報表示中にエラー: {str(e)}")
            
            with col2:
                try:
                    st.markdown("**変更後:**")
                    if new_shape.get('type') == 'image':
                        cell_ref = get_excel_cell_reference(new_shape.get('x', 0), new_shape.get('y', 0))
                        st.write(f"- 位置: セル {cell_ref}")
                        if new_shape.get('width') is not None and new_shape.get('height') is not None:
                            st.write(f"- サイズ: 幅 {new_shape['width']:.1f}px, 高さ {new_shape['height']:.1f}px")
                        else:
                            st.write("- サイズ情報なし")
                    else:
                        st.markdown(f"""
                        - 種類: {new_shape.get('type', 'unknown')}
                        - 位置: セル {get_excel_cell_reference(new_shape.get('x', 0), new_shape.get('y', 0))}
                        - テキスト: {new_shape.get('text', '') or 'なし'}
                        """)
                except Exception as e:
                    st.error(f"変更後の情報表示中にエラー: {str(e)}")

def export_comparison(comparison_result, sheet1_name=None, sheet2_name=None):
    """
    Export comparison results including shape differences and sheet names
    """
    output = io.BytesIO()
    
    # Create Excel writer object
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Write data differences with sheet names
        sheet1_label = f'File1_{sheet1_name}' if sheet1_name else 'File1'
        sheet2_label = f'File2_{sheet2_name}' if sheet2_name else 'File2'
        comparison_result['df1'].to_excel(writer, sheet_name=sheet1_label, index=False)
        comparison_result['df2'].to_excel(writer, sheet_name=sheet2_label, index=False)
        # Create a more detailed summary DataFrame with Excel-style cell references
        summary_data = []
        for diff in comparison_result['diff_summary'].to_dict('records'):
            if diff['type'] == 'modified':
                col_idx = comparison_result['df1'].columns.get_loc(diff['column'])
                cell_ref_old = get_excel_cell_reference(col_idx, diff['row_index_old'])
                cell_ref_new = get_excel_cell_reference(col_idx, diff['row_index_new'])
                summary_data.append({
                    '変更タイプ': '変更',
                    'セル位置 (変更前)': cell_ref_old,
                    'セル位置 (変更後)': cell_ref_new,
                    '変更前の値': diff['value_old'],
                    '変更後の値': diff['value_new'],
                    '類似度': f"{diff.get('similarity', 1.0):.2%}"
                })
            else:
                row_idx = diff['row_index']
                df = comparison_result['df1']
                range_ref = get_excel_range_reference(row_idx, 0, len(df.columns) - 1)
                row_values = []
                for col in df.columns:
                    val = diff['values'].get(col, '')
                    if pd.notna(val):
                        row_values.append(f"{col}: {val}")
                
                summary_data.append({
                    '変更タイプ': '行追加' if diff['type'] == 'added' else '削除',
                    'セル位置': f"{row_idx + 1}行目 ({range_ref})",
                    '値': ' | '.join(row_values),
                    '類似度': 'N/A'
                })
        
        summary_df = pd.DataFrame(summary_data)
        if not summary_df.empty:
            # シート名でソート可能にするために列の順序を調整
            columns_order = ['シート名', '変更タイプ', 'セル位置', 'セル位置 (変更前)', 'セル位置 (変更後)', '値', '変更前の値', '変更後の値']
            existing_columns = [col for col in columns_order if col in summary_df.columns]
            other_columns = [col for col in summary_df.columns if col not in columns_order]
            summary_df = summary_df[existing_columns + other_columns]
            summary_df.to_excel(writer, sheet_name='Data_Summary', index=False)
        
        # Write shape differences
        if 'shape_differences' in comparison_result:
            shape_diff_df = pd.DataFrame(comparison_result['shape_differences'])
            if not shape_diff_df.empty:
                shape_diff_df.to_excel(writer, sheet_name='Shape_Differences', index=False)
    
    # Prepare the file for download
    output.seek(0)
    
    # Create download button
    st.download_button(
        label="比較レポートをダウンロード",
        data=output,
        file_name=f"comparison_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
