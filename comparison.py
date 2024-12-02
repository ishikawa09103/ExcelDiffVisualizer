import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.drawing.spreadsheet_drawing import SpreadsheetDrawing
from openpyxl.drawing.image import Image
from openpyxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D

def extract_shape_info(wb, sheet_name):
    """
    Extract shape information from an Excel worksheet using the latest openpyxl API
    """
    shapes_info = []
    ws = wb[sheet_name]
    
    try:
        # Get drawings from the worksheet
        drawings = ws._drawing if hasattr(ws, '_drawing') else []
        if not drawings and hasattr(ws, 'drawings'):
            drawings = ws.drawings
        
        # Process each drawing
        for drawing in drawings if drawings else []:
            shape_info = {
                'type': 'unknown',
                'x': 0,
                'y': 0,
                'width': None,
                'height': None,
                'text': '',
                'description': ''
            }
            
            try:
                # Get anchor information
                if hasattr(drawing, 'anchor'):
                    anchor = drawing.anchor
                    shape_info.update({
                        'x': getattr(anchor, 'col', 0),
                        'y': getattr(anchor, 'row', 0),
                        'width': getattr(anchor, 'width', None),
                        'height': getattr(anchor, 'height', None)
                    })
                
                # Determine shape type and extract specific information
                if hasattr(drawing, '_shape_type'):
                    shape_info['type'] = drawing._shape_type
                elif isinstance(drawing, Image):
                    shape_info['type'] = 'image'
                else:
                    shape_info['type'] = type(drawing).__name__
                
                # Extract text if available
                if hasattr(drawing, 'text'):
                    shape_info['text'] = drawing.text
                elif hasattr(drawing, 'title'):
                    shape_info['text'] = drawing.title
                
                # Get additional description if available
                if hasattr(drawing, 'description'):
                    shape_info['description'] = drawing.description
                
                shapes_info.append(shape_info)
                
            except Exception as shape_error:
                st.warning(f"図形の処理中にエラーが発生しました: {str(shape_error)}")
                continue
                
    except Exception as ws_error:
        st.warning(f"ワークシートの描画情報へのアクセス中にエラーが発生しました: {str(ws_error)}")
        return []
    
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

        for col in key_columns:
            try:
                val = row[col]
                if pd.api.types.is_numeric_dtype(type(val)):
                    values.append(normalize_numeric(val))
                else:
                    values.append(normalize_string(val))
            except Exception:
                values.append('')
        
        # キー列の重み付けを反映したハッシュ値を生成
        weighted_values = []
        for i, val in enumerate(values):
            # キー列の順序に基づいて重み付け
            weight = len(key_columns) - i
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
                    if val1 != val2:
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
            # キー列により高い重みを設定
            weight = 3.0 if col in key_columns[:2] else 2.0 if col in key_columns else 1.0
            total_weight += weight
            
            val1 = row1[col]
            val2 = row2[col]
            
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
