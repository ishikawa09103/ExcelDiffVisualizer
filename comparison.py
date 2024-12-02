import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor

def extract_shape_info(workbook, sheet_name):
    sheet = workbook[sheet_name]
    shapes = []
    
    try:
        # Access drawings directly using the drawings property
        for drawing in sheet._drawings:
            shape_info = {
                'type': type(drawing).__name__,
                'name': drawing.name if hasattr(drawing, 'name') else '',
                'coordinates': {
                    'x': drawing.left if hasattr(drawing, 'left') else 0,
                    'y': drawing.top if hasattr(drawing, 'top') else 0,
                    'width': drawing.width if hasattr(drawing, 'width') else 0,
                    'height': drawing.height if hasattr(drawing, 'height') else 0
                }
            }
            
            # Additional properties for specific shape types
            if hasattr(drawing, 'text'):
                shape_info['text'] = drawing.text
            if hasattr(drawing, 'description'):
                shape_info['description'] = drawing.description
            
            shapes.append(shape_info)
            print(f"Debug: Found shape: {shape_info}")
    except Exception as e:
        print(f"Debug: Error in shape extraction: {str(e)}")
        print(f"Debug: Sheet properties: {dir(sheet)}")
    
    return shapes

def compare_shapes(shapes1, shapes2):
    """Compare shapes between two Excel files"""
    shape_differences = []
    
    # Create dictionaries for quick lookup
    shapes1_dict = {shape['name']: shape for shape in shapes1}
    shapes2_dict = {shape['name']: shape for shape in shapes2}
    
    print(f"Debug: Comparing shapes - File 1: {len(shapes1)} shapes, File 2: {len(shapes2)} shapes")
    
    # Find added and modified shapes
    for name, shape2 in shapes2_dict.items():
        if name not in shapes1_dict:
            print(f"Debug: Found added shape: {name}")
            shape_differences.append({
                'type': 'added',
                'shape_name': name,
                'details': shape2
            })
        else:
            shape1 = shapes1_dict[name]
            differences = {}
            
            # Compare coordinates with tolerance
            for coord in ['x', 'y', 'width', 'height']:
                val1 = shape1['coordinates'][coord]
                val2 = shape2['coordinates'][coord]
                if abs(val1 - val2) > 0.1:  # Using small tolerance for floating point comparison
                    differences[coord] = {
                        'old': val1,
                        'new': val2
                    }
                    print(f"Debug: Shape {name} - {coord} changed from {val1} to {val2}")
            
            # Compare other attributes
            for attr in ['type', 'text', 'description']:
                if attr in shape1 and attr in shape2:
                    if shape1[attr] != shape2[attr]:
                        differences[attr] = {
                            'old': shape1[attr],
                            'new': shape2[attr]
                        }
                        print(f"Debug: Shape {name} - {attr} changed from {shape1[attr]} to {shape2[attr]}")
            
            if differences:
                shape_differences.append({
                    'type': 'modified',
                    'shape_name': name,
                    'differences': differences
                })
    
    # Find deleted shapes
    for name in shapes1_dict:
        if name not in shapes2_dict:
            print(f"Debug: Found deleted shape: {name}")
            shape_differences.append({
                'type': 'deleted',
                'shape_name': name,
                'details': shapes1_dict[name]
            })
    
    return shape_differences

def compare_dataframes(df1, df2, file1=None, file2=None):
    """
    Compare two dataframes and their shapes, and return DataFrames with style information for AgGrid
    """
    # Initialize shape differences
    shape_differences = None
    
    # Extract and compare shapes if files are provided
    if file1 and file2:
        try:
            wb1 = load_workbook(file1)
            wb2 = load_workbook(file2)
            
            shapes1 = extract_shape_info(wb1, wb1.sheetnames[0])
            shapes2 = extract_shape_info(wb2, wb2.sheetnames[0])
            
            shape_differences = compare_shapes(shapes1, shapes2)
            print(f"Debug: Found {len(shapes1)} shapes in file1 and {len(shapes2)} shapes in file2")
        except Exception as e:
            print(f"Error comparing shapes: {str(e)}")
    
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
        'diff_summary': diff_summary,
        'shape_differences': shape_differences
    }
