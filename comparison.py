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
                st.warning(f"Error processing shape: {str(shape_error)}")
                continue
                
    except Exception as ws_error:
        st.warning(f"Error accessing worksheet drawings: {str(ws_error)}")
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
    using an improved cell change detection algorithm
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
    
    # Create key for row matching (concatenate values of all columns)
    def create_row_key(row):
        return '||'.join(str(row[col]) if pd.notna(row[col]) else '' for col in common_cols)
    
    # Create dictionaries for quick lookup
    df1_dict = {create_row_key(row): (idx, row) for idx, row in df1.iterrows()}
    df2_dict = {create_row_key(row): (idx, row) for idx, row in df2.iterrows()}
    
    # Track processed rows
    processed_keys_df1 = set()
    processed_keys_df2 = set()
    
    # Function to find best matching row
    def find_best_match(row_key, source_dict, target_dict):
        if row_key in target_dict:
            return row_key, 1.0
        
        max_similarity = 0
        best_key = None
        source_parts = set(row_key.split('||'))
        
        for target_key in target_dict.keys():
            target_parts = set(target_key.split('||'))
            similarity = len(source_parts & target_parts) / len(source_parts | target_parts)
            if similarity > max_similarity and similarity >= 0.7:  # 70% similarity threshold
                max_similarity = similarity
                best_key = target_key
        
        return best_key, max_similarity

    # First pass: Find and mark modified cells
    for key1, (idx1, row1) in df1_dict.items():
        if key1 in processed_keys_df1:
            continue
            
        best_key2, similarity = find_best_match(key1, df1_dict, df2_dict)
        
        if best_key2:
            idx2, row2 = df2_dict[best_key2]
            processed_keys_df1.add(key1)
            processed_keys_df2.add(best_key2)
            
            # Check for modifications in matched rows
            for col in common_cols:
                val1, val2 = row1[col], row2[col]
                if pd.notna(val1) and pd.notna(val2) and val1 != val2:
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
                        'similarity': similarity
                    })
    
    # Second pass: Mark remaining rows as added/deleted
    for key1, (idx1, row1) in df1_dict.items():
        if key1 not in processed_keys_df1:
            for col in common_cols:
                if pd.notna(row1[col]):
                    df1_styles.append({
                        'field': col,
                        'rowIndex': idx1,
                        'cellClass': 'ag-cell-deleted'
                    })
            differences.append({
                'type': 'deleted',
                'row_index': idx1,
                'values': row1[common_cols].to_dict()
            })
    
    for key2, (idx2, row2) in df2_dict.items():
        if key2 not in processed_keys_df2:
            for col in common_cols:
                if pd.notna(row2[col]):
                    df2_styles.append({
                        'field': col,
                        'rowIndex': idx2,
                        'cellClass': 'ag-cell-added'
                    })
            differences.append({
                'type': 'added',
                'row_index': idx2,
                'values': row2[common_cols].to_dict()
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
