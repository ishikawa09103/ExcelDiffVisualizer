import pandas as pd
import streamlit as st
import numpy as np
from openpyxl import load_workbook
from openpyxl.drawing.spreadsheet_drawing import SpreadsheetDrawing
from openpyxl.drawing.image import Image
from openpyxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D

def extract_shape_info(wb, sheet_name):
    st.write("図形情報の抽出を開始...")
    st.write(f"ワークシート名: {sheet_name}")
    
    shapes_info = []
    ws = wb[sheet_name]
    
    try:
        # Method 1: _drawing.drawingsから画像を取得
        st.write("描画オブジェクトの検索...")
        if hasattr(ws, '_drawing') and ws._drawing:
            for shape in ws._drawing.drawings:
                try:
                    if hasattr(shape, '_rel') and shape._rel.target:
                        # 図形の情報をデバッグ出力
                        st.write(f"検出された図形の種類: {shape._type if hasattr(shape, '_type') else 'unknown'}")
                        st.write(f"図形の位置: セル {getattr(shape, 'col', 0)}, {getattr(shape, 'row', 0)}")
                        st.write(f"図形のサイズ: 幅 {getattr(shape, 'width', 'N/A')}, 高さ {getattr(shape, 'height', 'N/A')}")
                        
                        # 画像情報を取得
                        x = getattr(shape, 'col', 0) or 0
                        y = getattr(shape, 'row', 0) or 0
                        width = getattr(shape, 'width', None)
                        height = getattr(shape, 'height', None)
                        
                        shapes_info.append({
                            'type': 'image',
                            'x': x,
                            'y': y,
                            'width': width,
                            'height': height,
                            'text': ''
                        })
                except Exception as e:
                    st.warning(f"画像の処理中にエラー: {str(e)}")
                    continue
        
        # Method 2: _imagesから画像を取得
        st.write("埋め込み画像オブジェクトの検索...")
        if hasattr(ws, '_images'):
            for img in ws._images:
                try:
                    anchor = getattr(img, 'anchor', None)
                    if anchor:
                        # 図形の情報をデバッグ出力
                        st.write(f"検出された図形の種類: image")
                        st.write(f"図形の位置: セル {getattr(anchor, 'col', 0)}, {getattr(anchor, 'row', 0)}")
                        st.write(f"図形のサイズ: 幅 {getattr(anchor, 'width', 'N/A')}, 高さ {getattr(anchor, 'height', 'N/A')}")
                        
                        x = getattr(anchor, 'col', 0) or 0
                        y = getattr(anchor, 'row', 0) or 0
                        width = getattr(anchor, 'width', None)
                        height = getattr(anchor, 'height', None)
                        
                        shapes_info.append({
                            'type': 'image',
                            'x': x,
                            'y': y,
                            'width': width,
                            'height': height,
                            'text': ''
                        })
                except Exception as e:
                    st.warning(f"画像の処理中にエラー: {str(e)}")
                    continue

        # Method 3: 図形（シェイプ）の検出
        st.write("図形（シェイプ）オブジェクトの検索...")

        # 全ての可能な図形コンテナを確認
        shape_containers = []

        # 1. _drawingから図形を検出
        if hasattr(ws, '_drawing'):
            st.write("_drawingから図形を検索中...")
            # _shapesプロパティを確認
            if hasattr(ws._drawing, '_shapes'):
                st.write("_shapes検出")
                shape_containers.extend(ws._drawing._shapes)
            # shapesプロパティを確認
            if hasattr(ws._drawing, 'shapes'):
                st.write("shapes検出")
                shape_containers.extend(ws._drawing.shapes)

        # 2. 直接のshapesプロパティを確認
        if hasattr(ws, 'shapes'):
            st.write("ワークシートの直接のshapesを検索中...")
            shape_containers.extend(ws.shapes)

        # 3. _chartsプロパティを確認（図形が含まれる可能性がある）
        if hasattr(ws, '_charts'):
            st.write("_chartsから図形を検索中...")
            shape_containers.extend(ws._charts)

        st.write(f"検出された図形コンテナの数: {len(shape_containers)}")

        for shape in shape_containers:
            try:
                st.write(f"図形オブジェクトを処理中: {shape}")
                st.write(f"図形のタイプ: {type(shape)}")
                st.write(f"図形の属性: {dir(shape)}")

                # 図形の種類を決定
                shape_type = None
                if hasattr(shape, 'type'):
                    shape_type = shape.type
                elif hasattr(shape, '_type'):
                    shape_type = shape._type
                elif hasattr(shape, 'shape_type'):
                    shape_type = shape.shape_type
                else:
                    shape_type = str(type(shape).__name__)

                st.write(f"検出された図形の種類: {shape_type}")

                # 位置情報の取得
                x, y = 0, 0
                if hasattr(shape, 'anchor'):
                    anchor = shape.anchor
                    x = getattr(anchor, 'col', 0)
                    y = getattr(anchor, 'row', 0)
                    st.write(f"アンカー位置: ({x}, {y})")
                elif hasattr(shape, '_anchor'):
                    anchor = shape._anchor
                    x = getattr(anchor, 'col', 0)
                    y = getattr(anchor, 'row', 0)
                    st.write(f"_アンカー位置: ({x}, {y})")
                elif hasattr(shape, 'coordinate'):
                    x, y = shape.coordinate
                    st.write(f"座標位置: ({x}, {y})")

                # サイズ情報の取得
                width = getattr(shape, 'width', None)
                height = getattr(shape, 'height', None)
                if width is not None and height is not None:
                    st.write(f"図形のサイズ: 幅 {width}, 高さ {height}")

                shapes_info.append({
                    'type': 'shape',
                    'shape_type': shape_type,
                    'x': x,
                    'y': y,
                    'width': width,
                    'height': height,
                    'text': getattr(shape, 'text', '')
                })
                st.write("図形情報を追加しました")

            except Exception as e:
                st.warning(f"図形の処理中にエラー: {str(e)}")
                st.write(f"エラーの詳細: {type(e).__name__}")
                continue
                    
    except Exception as e:
        st.warning(f"ワークシートの処理中にエラー: {str(e)}")
    
    # 図形の合計数と種類別集計を表示
    st.write(f"検出された図形の合計数: {len(shapes_info)}")
    st.write("図形の種類別集計:")
    shape_types = {}
    for shape in shapes_info:
        shape_type = shape.get('type', 'unknown')
        shape_types[shape_type] = shape_types.get(shape_type, 0) + 1
    for shape_type, count in shape_types.items():
        st.write(f"- {shape_type}: {count}個")
    
    return shapes_info

def compare_shapes(shapes1, shapes2):
    """
    Compare shapes between two Excel files
    """
    differences = []
    
    # Find added and modified shapes
    for idx2, shape2 in enumerate(shapes2):
        found_match = False
        for idx1, shape1 in enumerate(shapes1):
            if (shape1['x'] == shape2['x'] and 
                shape1['y'] == shape2['y'] and 
                shape1['type'] == shape2['type']):
                found_match = True
                # Check for modifications
                if (shape1['width'] != shape2['width'] or 
                    shape1['height'] != shape2['height'] or 
                    shape1['text'] != shape2['text']):
                    differences.append({
                        'type': 'modified',
                        'shape_index': idx2,
                        'old_shape': shape1,
                        'new_shape': shape2
                    })
                break
        
        if not found_match:
            differences.append({
                'type': 'added',
                'shape_index': idx2,
                'shape': shape2
            })
    
    # Find deleted shapes
    for idx1, shape1 in enumerate(shapes1):
        found_match = False
        for shape2 in shapes2:
            if (shape1['x'] == shape2['x'] and 
                shape1['y'] == shape2['y'] and 
                shape1['type'] == shape2['type']):
                found_match = True
                break
        
        if not found_match:
            differences.append({
                'type': 'deleted',
                'shape_index': idx1,
                'shape': shape1
            })
    
    return differences

def compare_dataframes(df1, df2):
    """
    Compare two dataframes and return differences with improved row matching
    """
    # Create copies for styling
    df1_result = df1.copy()
    df2_result = df2.copy()
    
    # Initialize style information
    df1_styles = []
    df2_styles = []
    differences = []
    
    # Get common columns
    common_cols = list(set(df1.columns) & set(df2.columns))
    
    # Identify potential key columns
    key_columns = [col for col in common_cols if any(key in col.lower() 
                  for key in ['id', 'code', 'key', 'name', 'no', '番号'])]
    
    if not key_columns:
        # If no key columns found, use the first column and additional columns for better matching
        key_columns = common_cols[:min(3, len(common_cols))]
    
    def create_row_hash(row):
        """Create a hash value for row matching based on key columns with improved type handling"""
        values = []
        
        def is_no_column(col_name):
            """No.列かどうかを判定"""
            return col_name.lower().strip() in ['no', 'no.', '番号']
        
        def normalize_numeric(val):
            """数値を正規化して文字列に変換"""
            if pd.isna(val) or val is None:
                return ''
            try:
                if isinstance(val, (int, np.integer)):
                    return str(int(val))
                elif isinstance(val, float):
                    # 整数の場合は整数として扱う
                    if val.is_integer():
                        return str(int(val))
                    # 小数の場合は固定精度で表現
                    return f"{val:.6f}".rstrip('0').rstrip('.')
                else:
                    return str(val)
            except (AttributeError, ValueError, TypeError):
                return str(val)

        def normalize_string(val):
            """文字列を正規化"""
            if pd.isna(val) or val is None:
                return ''
            return str(val).strip().lower()  # 大文字小文字を区別しない

        # No.列以外のキー列を優先して処理
        no_columns = []
        other_key_columns = []
        
        for col in key_columns:
            if is_no_column(col):
                no_columns.append(col)
            else:
                other_key_columns.append(col)
        
        # No.列以外のキー列を処理
        for col in other_key_columns:
            try:
                val = row[col]
                if pd.api.types.is_numeric_dtype(type(val)):
                    values.append(normalize_numeric(val))
                else:
                    values.append(normalize_string(val))
            except Exception:
                values.append('')
        
        # No.列を処理（行番号の自動更新を考慮）
        for col in no_columns:
            try:
                val = row[col]
                # No.列は参考情報として扱い、重みを小さくする
                values.append(f"no_ref:{normalize_numeric(val)}")
            except Exception:
                values.append('')
        
        # キー列の重み付けを反映したハッシュ値を生成
        weighted_values = []
        for i, val in enumerate(values):
            if val.startswith('no_ref:'):
                # No.列は最も低い重みを設定
                weight = 0.1
            else:
                # その他のキー列は順序に基づいて重み付け
                weight = len(other_key_columns) - i if i < len(other_key_columns) else 0.5
            weighted_values.append(f"{weight}:{val}")
        
        return '||'.join(weighted_values)
    
    # Create hash values for both dataframes with error handling
    try:
        df1['_row_hash'] = df1.apply(create_row_hash, axis=1)
        df2['_row_hash'] = df2.apply(create_row_hash, axis=1)
    except Exception as e:
        # エラーが発生した場合は、インデックスをハッシュとして使用
        df1['_row_hash'] = df1.index.astype(str)
        df2['_row_hash'] = df2.index.astype(str)
    
    # Initialize tracking sets
    matched_df1_indices = set()
    matched_df2_indices = set()
    
    # First pass: Find exact matches using hash values
    hash_map_df2 = {hash_val: idx for idx, hash_val in enumerate(df2['_row_hash'])}
    
    for idx1, hash_val in enumerate(df1['_row_hash']):
        if hash_val in hash_map_df2:
            idx2 = hash_map_df2[hash_val]
            if idx2 not in matched_df2_indices:
                matched_df1_indices.add(idx1)
                matched_df2_indices.add(idx2)
                
                # Check for modifications in matched rows
                row1, row2 = df1.iloc[idx1], df2.iloc[idx2]
                for col in common_cols:
                    val1 = str(row1[col]).strip() if pd.notna(row1[col]) else ''
                    val2 = str(row2[col]).strip() if pd.notna(row2[col]) else ''
                    
                    # No.列の場合は特別な処理
                    if col.lower().strip() in ['no', 'no.', '番号']:
                        # No.列の変更は、他の列に変更がある場合のみ記録
                        continue
                    
                    if val1 != val2:
                        # 実際の変更として記録
                        df1_styles.append({
                            'field': col,
                            'rowIndex': idx1,
                            'cellClass': 'ag-cell-modified'
                        })
                        df2_styles.append({
                            'field': col,
                            'rowIndex': idx2,
                            'cellClass': 'ag-cell-modified'
                        })
                        differences.append({
                            'type': 'modified',
                            'column': col,
                            'row_index_old': idx1,
                            'row_index_new': idx2,
                            'value_old': val1,
                            'value_new': val2
                        })
    
    # Second pass: Handle remaining rows using similarity matching
    def calculate_row_similarity(row1, row2):
        """Calculate similarity between two rows with improved matching logic"""
        matches = 0
        total_weight = 0
        no_column_differences = []  # No.列の差分を追跡
        
        def calculate_string_similarity(s1, s2):
            """文字列の類似度を計算（レーベンシュタイン距離ベース）"""
            if not s1 and not s2:  # 両方空の場合
                return 1.0
            if not s1 or not s2:  # どちらかが空の場合
                return 0.0
                
            # 文字列を正規化
            s1 = str(s1).strip().lower()
            s2 = str(s2).strip().lower()
            
            if s1 == s2:
                return 1.0
                
            # 簡易的なレーベンシュタイン距離の計算
            len_s1, len_s2 = len(s1), len(s2)
            if len_s1 < len_s2:
                s1, s2 = s2, s1
                len_s1, len_s2 = len_s2, len_s1
            
            # 文字の一致度を計算
            matches = sum(1 for i in range(min(len_s1, len_s2)) if s1[i] == s2[i])
            return matches / max(len_s1, len_s2)
        
        def calculate_numeric_similarity(v1, v2):
            """数値の類似度を計算"""
            try:
                if pd.isna(v1) and pd.isna(v2):
                    return 1.0
                if pd.isna(v1) or pd.isna(v2):
                    return 0.0
                
                n1 = float(v1)
                n2 = float(v2)
                
                if n1 == n2:
                    return 1.0
                    
                # 数値の差に基づく類似度
                max_val = max(abs(n1), abs(n2))
                if max_val == 0:
                    return 1.0
                    
                diff_ratio = abs(n1 - n2) / max_val
                return max(0, 1 - diff_ratio)
            except (ValueError, TypeError):
                return 0.0
        
        for col in common_cols:
            # No.列の特別処理
            is_no_col = col.lower().strip() in ['no', 'no.', '番号']
            
            # キー列により高い重みを設定（No.列は低い重みに）
            if is_no_col:
                weight = 0.1
            else:
                weight = 3.0 if col in key_columns[:2] else 2.0 if col in key_columns else 1.0
            
            total_weight += weight
            val1 = row1[col]
            val2 = row2[col]
            
            # No.列の差分を追跡
            if is_no_col and val1 != val2:
                no_column_differences.append((col, val1, val2))
            
            # 数値型の場合
            if pd.api.types.is_numeric_dtype(type(val1)) or pd.api.types.is_numeric_dtype(type(val2)):
                similarity = calculate_numeric_similarity(val1, val2)
            else:
                # 文字列型の場合
                similarity = calculate_string_similarity(val1, val2)
            
            matches += weight * similarity
        
        return matches / total_weight if total_weight > 0 else 0
    
    # Process unmatched rows with similarity matching
    unmatched_df1 = [i for i in range(len(df1)) if i not in matched_df1_indices]
    unmatched_df2 = [i for i in range(len(df2)) if i not in matched_df2_indices]
    
    similarity_threshold = 0.8
    
    for idx1 in unmatched_df1:
        best_match = None
        best_similarity = similarity_threshold
        row1 = df1.iloc[idx1]
        
        for idx2 in unmatched_df2:
            row2 = df2.iloc[idx2]
            similarity = calculate_row_similarity(row1, row2)
            
            if similarity > best_similarity:
                best_similarity = similarity
                best_match = idx2
        
        if best_match is not None:
            # Found a similar row
            matched_df1_indices.add(idx1)
            matched_df2_indices.add(best_match)
            
            # Mark modified cells
            row2 = df2.iloc[best_match]
            for col in common_cols:
                val1 = str(row1[col]).strip() if pd.notna(row1[col]) else ''
                val2 = str(row2[col]).strip() if pd.notna(row2[col]) else ''
                if val1 != val2:
                    df1_styles.append({
                        'field': col,
                        'rowIndex': idx1,
                        'cellClass': 'ag-cell-modified'
                    })
                    df2_styles.append({
                        'field': col,
                        'rowIndex': best_match,
                        'cellClass': 'ag-cell-modified'
                    })
                    differences.append({
                        'type': 'modified',
                        'column': col,
                        'row_index_old': idx1,
                        'row_index_new': best_match,
                        'value_old': val1,
                        'value_new': val2
                    })
        else:
            # No similar row found - this row was deleted
            row = df1.iloc[idx1]
            for col in common_cols:
                if pd.notna(row[col]):
                    df1_styles.append({
                        'field': col,
                        'rowIndex': idx1,
                        'cellClass': 'ag-cell-deleted'
                    })
            differences.append({
                'type': 'deleted',
                'row_index': idx1,
                'values': row[common_cols].to_dict()
            })
    
    # Mark remaining unmatched rows in df2 as added
    for idx2 in range(len(df2)):
        if idx2 not in matched_df2_indices:
            row = df2.iloc[idx2]
            for col in common_cols:
                if pd.notna(row[col]):
                    df2_styles.append({
                        'field': col,
                        'rowIndex': idx2,
                        'cellClass': 'ag-cell-added'
                    })
            differences.append({
                'type': 'added',
                'row_index': idx2,
                'values': row[common_cols].to_dict()
            })
    
    # Remove temporary hash columns
    if '_row_hash' in df1_result.columns:
        df1_result.drop('_row_hash', axis=1, inplace=True)
    if '_row_hash' in df2_result.columns:
        df2_result.drop('_row_hash', axis=1, inplace=True)
    
    # Create difference summary
    diff_summary = pd.DataFrame(differences)
    
    return {
        'df1': df1_result,
        'df2': df2_result,
        'df1_styles': df1_styles,
        'df2_styles': df2_styles,
        'diff_summary': diff_summary
    }
