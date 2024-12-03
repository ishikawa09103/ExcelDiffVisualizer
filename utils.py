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
    try:
        gb = GridOptionsBuilder.from_dataframe(df)
        
        # Configure default column behavior
        gb.configure_default_column(
            resizable=True,
            filterable=True,
            sorteable=True,
            editable=False,
            suppressMovable=True
        )
        
        # Add cell styling if provided
        if cell_styles:
            cell_style_jscode = JsCode("""
            function(params) {
                try {
                    return {
                        'backgroundColor': params.data && params.data._cellStyles ? params.data._cellStyles[params.column.colId] : null
                    };
                } catch (e) {
                    console.warn('Cell style error:', e);
                    return null;
                }
            }
            """)
            
            gb.configure_grid_options(
                getRowStyle=None,
                getCellStyle=cell_style_jscode,
                onGridReady=JsCode("""
                function(params) {
                    params.api.sizeColumnsToFit();
                }
                """)
            )
        
        grid_options = gb.build()
        
        if cell_styles:
            grid_options['context'] = {'cell_styles': cell_styles}
            
        grid_options['onGridReady'] = JsCode("""
        function(params) {
            try {
                params.api.sizeColumnsToFit();
            } catch (e) {
                console.warn('Grid ready error:', e);
            }
        }
        """)
        
        return AgGrid(
            df,
            gridOptions=grid_options,
            update_mode='VALUE_CHANGED',
            allow_unsafe_jscode=True,
            theme='streamlit',
            custom_css={
                ".ag-cell-added": {"backgroundColor": "#D4EDDA !important"},
                ".ag-cell-deleted": {"backgroundColor": "#F8D7DA !important"},
                ".ag-cell-modified": {"backgroundColor": "#FFF3CD !important"}
            },
            key=f"grid_{id(df)}"
        )
    except Exception as e:
        st.error(f"グリッドの作成中にエラーが発生しました: {str(e)}")
        return st.dataframe(df)

def display_shape_differences(shape_differences):
    """
    Display shape differences in a formatted way with improved image information
    """
    st.write("画像の差分処理を開始...")
    
    for diff in shape_differences:
        if diff['type'] == 'added':
            shape = diff.get('shape', {})
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
            old_shape = diff.get('old_shape', {})
            new_shape = diff.get('new_shape', {})
            
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

def export_comparison(comparison_results, sheets1, sheets2, sheet_pairs=None):
    """
    Export comparison results for all sheets in a single Excel file
    comparison_results: List of comparison result dictionaries, each containing df1, df2, diff_summary, etc.
    sheets1, sheets2: Lists of sheet names from both files for tracking added/deleted sheets
    """
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # サマリーデータの作成
        all_summary_data = []
        
        # シート名の変更情報を追加
        if sheet_pairs:
            sheet_changes = []
            for old_name, new_name, similarity in sheet_pairs:
                sheet_changes.append({
                    'シート名': f"{old_name} → {new_name}",
                    '変更タイプ': 'シート名変更',
                    '値': f"シート名が変更されました (類似度: {similarity:.1%})"
                })
            all_summary_data.extend(sheet_changes)

        # シートの追加/削除情報を追加
        added_sheets = set(sheets2) - set(sheets1) - set(s[1] for s in sheet_pairs) if sheet_pairs else set(sheets2) - set(sheets1)
        deleted_sheets = set(sheets1) - set(sheets2) - set(s[0] for s in sheet_pairs) if sheet_pairs else set(sheets1) - set(sheets2)
        
        if added_sheets or deleted_sheets:
            sheet_changes = []
            for sheet in added_sheets:
                sheet_changes.append({
                    'シート名': sheet,
                    '変更タイプ': 'シート追加',
                    '値': f"新しいシート '{sheet}' が追加されました"
                })
            for sheet in deleted_sheets:
                sheet_changes.append({
                    'シート名': sheet,
                    '変更タイプ': 'シート削除',
                    '値': f"シート '{sheet}' が削除されました"
                })
            all_summary_data.extend(sheet_changes)
        
        # 各シートペアの比較結果をまとめて処理
        for i, result in enumerate(comparison_results):
            sheet1_name = result.get('sheet1_name', f'Sheet1_{i+1}')
            sheet2_name = result.get('sheet2_name', f'Sheet2_{i+1}')
            sheet_pair = f"{sheet1_name} → {sheet2_name}"
            
            # データの変更を処理
            data_changes = []
            if 'diff_summary' in result:
                for diff in result['diff_summary'].to_dict('records'):
                    change_info = {
                        'シート名': sheet_pair,
                        '変更タイプ': 'データ変更'
                    }
                    
                    if diff['type'] == 'modified':
                        change_info.update({
                            'セル位置 (変更前)': get_excel_cell_reference(result['df1'].columns.get_loc(diff['column']), diff['row_index_old']),
                            'セル位置 (変更後)': get_excel_cell_reference(result['df1'].columns.get_loc(diff['column']), diff['row_index_new']),
                            '変更前の値': diff['value_old'],
                            '変更後の値': diff['value_new']
                        })
                    else:
                        df = result['df1'] if diff['type'] == 'deleted' else result['df2']
                        range_ref = get_excel_range_reference(diff['row_index'], 0, len(df.columns) - 1)
                        change_info.update({
                            '変更タイプ': '行追加' if diff['type'] == 'added' else '行削除',
                            'セル位置': f"{diff['row_index'] + 1}行目 ({range_ref})",
                            '値': ' | '.join([f"{k}: {v}" for k, v in diff['values'].items() if pd.notna(v)])
                        })
                    
                    data_changes.append(change_info)
            
            # 図形の変更を処理
            shape_changes = []
            if 'shape_differences' in result:
                for shape_diff in result['shape_differences']:
                    shape_info = {
                        'シート名': sheet_pair,
                        '変更タイプ': f'図形{shape_diff["type"]}'
                    }
                    
                    if shape_diff['type'] == 'modified':
                        old_shape = shape_diff['old_shape']
                        new_shape = shape_diff['new_shape']
                        shape_info.update({
                            'セル位置 (変更前)': get_excel_cell_reference(old_shape['x'], old_shape['y']),
                            'セル位置 (変更後)': get_excel_cell_reference(new_shape['x'], new_shape['y']),
                            '変更前の値': f"Type: {old_shape['type']}, Text: {old_shape.get('text', '')}",
                            '変更後の値': f"Type: {new_shape['type']}, Text: {new_shape.get('text', '')}"
                        })
                    else:
                        shape = shape_diff.get('shape', {})
                        shape_info.update({
                            'セル位置': get_excel_cell_reference(shape['x'], shape['y']),
                            '値': f"Type: {shape['type']}, Text: {shape.get('text', '')}"
                        })
                    
                    shape_changes.append(shape_info)
            
            # シートごとの変更をまとめて追加
            all_summary_data.extend(data_changes + shape_changes)
            
            # 元のデータを保存
            result['df1'].to_excel(writer, sheet_name=f'F1_{sheet1_name[:26]}', index=False)
            result['df2'].to_excel(writer, sheet_name=f'F2_{sheet2_name[:26]}', index=False)
        
        # サマリーシートの作成
        if all_summary_data:
            summary_df = pd.DataFrame(all_summary_data)
            # 列の順序を整理
            columns_order = ['シート名', '変更タイプ', 'セル位置', 'セル位置 (変更前)', 
                           'セル位置 (変更後)', '値', '変更前の値', '変更後の値']
            existing_columns = [col for col in columns_order if col in summary_df.columns]
            summary_df = summary_df[existing_columns]
            
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
    
    # ファイルをダウンロード可能な状態にする
    output.seek(0)
    
    # ダウンロードボタンを作成
    st.download_button(
        label="比較レポートをダウンロード",
        data=output,
        file_name=f"comparison_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
