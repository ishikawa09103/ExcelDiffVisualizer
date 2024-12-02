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
    Compare two dataframes and return DataFrames with style information for AgGrid
    using an improved row matching algorithm that correctly handles added rows
    """
    # Create copies for styling
    df1_result = df1.copy()
    df2_result = df2.copy()
    
    # Initialize style information
    df1_styles = []
    df2_styles = []
    differences = []
    
    # Compare common columns
    common_cols = list(set(df1.columns) & set(df2.columns))
    
    # Create row signature for better matching
    def create_row_signature(row):
        # Use key columns if they exist (e.g., ID, Name, etc.)
        key_columns = [col for col in common_cols if any(key in col.lower() 
                      for key in ['id', 'code', 'key', 'name', 'no'])]
        
        if key_columns:
            values = [str(row[col]).strip() if pd.notna(row[col]) else '' 
                     for col in key_columns]
        else:
            # If no key columns, use all columns but give more weight to the first few
            values = [str(row[col]).strip() if pd.notna(row[col]) else '' 
                     for col in common_cols[:3]]
            
        return '||'.join(values)
    
    # Create signatures for both dataframes
    df1_signatures = df1.apply(create_row_signature, axis=1)
    df2_signatures = df2.apply(create_row_signature, axis=1)
    
    # Track matched rows
    matched_rows_df1 = set()
    matched_rows_df2 = set()
    
    # First pass: Find exact matches using signatures
    signature_matches = {}
    for idx1, sig1 in enumerate(df1_signatures):
        if sig1 in df2_signatures.values:
            idx2 = df2_signatures[df2_signatures == sig1].index[0]
            signature_matches[idx1] = idx2
            matched_rows_df1.add(idx1)
            matched_rows_df2.add(idx2)
            
            # Check for modifications in exactly matched rows
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
                        'value_new': val2,
                        'similarity': 1.0
                    })
    
    # Second pass: Handle remaining rows using similarity matching
    def calculate_similarity(row1, row2):
        matches = 0
        total = len(common_cols)
        for col in common_cols:
            val1 = str(row1[col]).strip() if pd.notna(row1[col]) else ''
            val2 = str(row2[col]).strip() if pd.notna(row2[col]) else ''
            if val1 == val2:
                matches += 1
        return matches / total if total > 0 else 0
    
    # Process remaining unmatched rows
    for idx1 in range(len(df1)):
        if idx1 in matched_rows_df1:
            continue
            
        best_match = None
        best_similarity = 0.7  # Minimum similarity threshold
        
        for idx2 in range(len(df2)):
            if idx2 in matched_rows_df2:
                continue
                
            similarity = calculate_similarity(df1.iloc[idx1], df2.iloc[idx2])
            if similarity > best_similarity:
                best_similarity = similarity
                best_match = idx2
        
        if best_match is not None:
            # Found a similar row
            matched_rows_df1.add(idx1)
            matched_rows_df2.add(best_match)
            
            # Mark modified cells
            row1, row2 = df1.iloc[idx1], df2.iloc[best_match]
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
                        'value_new': val2,
                        'similarity': best_similarity
                    })
        else:
            # No match found - row was deleted
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
        if idx2 not in matched_rows_df2:
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
    
    # Create difference summary
    diff_summary = pd.DataFrame(differences)
    
    return {
        'df1': df1_result,
        'df2': df2_result,
        'df1_styles': df1_styles,
        'df2_styles': df2_styles,
        'diff_summary': diff_summary
    }
