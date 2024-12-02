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
    """
    # Create copies for styling
    df1_result = df1.copy()
    df2_result = df2.copy()
    
    # Initialize style information
    df1_styles = []
    df2_styles = []
    
    # Compare common columns
    common_cols = list(set(df1.columns) & set(df2.columns))
    
    # Track differences
    differences = []
    
    # Compare values in common columns
    for col in common_cols:
        # Get maximum length
        max_len = max(len(df1), len(df2))
        
        # Pad shorter dataframe with NaN
        s1 = df1[col].reindex(range(max_len))
        s2 = df2[col].reindex(range(max_len))
        
        # Compare values
        for idx in range(max_len):
            val1 = s1.iloc[idx] if idx < len(df1) else np.nan
            val2 = s2.iloc[idx] if idx < len(df2) else np.nan
            
            if pd.isna(val1) and not pd.isna(val2):
                # Added in df2
                if idx < len(df2):
                    df2_styles.append({
                        'field': col,
                        'rowIndex': idx,
                        'cellClass': 'ag-cell-added'
                    })
                    differences.append({
                        'type': 'added',
                        'column': col,
                        'row': idx,
                        'value': val2
                    })
            elif not pd.isna(val1) and pd.isna(val2):
                # Deleted in df2
                if idx < len(df1):
                    df1_styles.append({
                        'field': col,
                        'rowIndex': idx,
                        'cellClass': 'ag-cell-deleted'
                    })
                    differences.append({
                        'type': 'deleted',
                        'column': col,
                        'row': idx,
                        'value': val1
                    })
            elif not pd.isna(val1) and not pd.isna(val2) and val1 != val2:
                # Modified
                if idx < len(df1):
                    df1_styles.append({
                        'field': col,
                        'rowIndex': idx,
                        'cellClass': 'ag-cell-modified'
                    })
                if idx < len(df2):
                    df2_styles.append({
                        'field': col,
                        'rowIndex': idx,
                        'cellClass': 'ag-cell-modified'
                    })
                differences.append({
                    'type': 'modified',
                    'column': col,
                    'row': idx,
                    'value_old': val1,
                    'value_new': val2
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
