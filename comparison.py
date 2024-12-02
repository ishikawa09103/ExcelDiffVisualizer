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
    using an intelligent row matching algorithm to detect actual changes
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
    
    # Create a similarity matrix for row matching
    def calculate_row_similarity(row1, row2):
        matches = sum(row1[col] == row2[col] for col in common_cols if pd.notna(row1[col]) and pd.notna(row2[col]))
        total = sum(1 for col in common_cols if pd.notna(row1[col]) and pd.notna(row2[col]))
        return matches / total if total > 0 else 0

    # Track matched rows to avoid duplicate matches
    matched_rows_df1 = set()
    matched_rows_df2 = set()
    
    # Find matches using similarity threshold
    SIMILARITY_THRESHOLD = 0.8
    
    # First pass: Find exact matches
    for idx1 in range(len(df1)):
        if idx1 in matched_rows_df1:
            continue
            
        row1 = df1.iloc[idx1]
        best_match_idx = None
        best_match_score = 0
        
        for idx2 in range(len(df2)):
            if idx2 in matched_rows_df2:
                continue
                
            row2 = df2.iloc[idx2]
            similarity = calculate_row_similarity(row1, row2)
            
            if similarity == 1.0:  # Exact match
                matched_rows_df1.add(idx1)
                matched_rows_df2.add(idx2)
                break
            elif similarity > best_match_score:
                best_match_score = similarity
                best_match_idx = idx2
        
        # If no exact match but good similarity, mark as modified
        if best_match_score >= SIMILARITY_THRESHOLD and best_match_idx is not None:
            matched_rows_df1.add(idx1)
            matched_rows_df2.add(best_match_idx)
            
            # Check for modifications in matched rows
            row2 = df2.iloc[best_match_idx]
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
                        'rowIndex': best_match_idx,
                        'cellClass': 'ag-cell-modified'
                    })
                    differences.append({
                        'type': 'modified',
                        'column': col,
                        'row': idx1,
                        'value_old': val1,
                        'value_new': val2
                    })
    
    # Mark unmatched rows as added/deleted
    for idx1 in range(len(df1)):
        if idx1 not in matched_rows_df1:
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
                'row': idx1,
                'values': row[common_cols].to_dict()
            })
    
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
                'row': idx2,
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
