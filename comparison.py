import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor
from openpyxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D

def extract_shape_info(workbook, sheet_name):
    """Extract shape information from an Excel worksheet"""
    sheet = workbook[sheet_name]
    shapes = []
    
    try:
        # Use _drawings instead of _shapes
        for drawing in sheet._drawings:
            for shape in drawing.shapes:
                try:
                    shape_info = {
                        'type': shape.shape_type if hasattr(shape, 'shape_type') else type(shape).__name__,
                        'name': shape.name if hasattr(shape, 'name') else '',
                        'description': shape.description if hasattr(shape, 'description') else '',
                        'coordinates': {
                            'x': shape.left if hasattr(shape, 'left') else 0,
                            'y': shape.top if hasattr(shape, 'top') else 0,
                            'width': shape.width if hasattr(shape, 'width') else 0,
                            'height': shape.height if hasattr(shape, 'height') else 0
                        }
                    }
                    shapes.append(shape_info)
                except Exception as e:
                    print(f"Warning: Error processing shape: {str(e)}")
                    continue
    except AttributeError:
        print(f"Info: No shapes found in sheet '{sheet_name}'")
    except Exception as e:
        print(f"Error: Failed to process shapes in sheet '{sheet_name}': {str(e)}")
    
    return shapes

def compare_shapes(shapes1, shapes2):
    """Compare shapes between two Excel files"""
    shape_differences = []
    
    # Create dictionaries for quick lookup
    shapes1_dict = {shape['name']: shape for shape in shapes1}
    shapes2_dict = {shape['name']: shape for shape in shapes2}
    
    # Find added and modified shapes
    for name, shape2 in shapes2_dict.items():
        if name not in shapes1_dict:
            shape_differences.append({
                'type': 'added',
                'shape_name': name,
                'details': shape2
            })
        else:
            shape1 = shapes1_dict[name]
            differences = {}
            
            # Compare coordinates
            for coord in ['x', 'y', 'width', 'height']:
                if shape1['coordinates'][coord] != shape2['coordinates'][coord]:
                    differences[coord] = {
                        'old': shape1['coordinates'][coord],
                        'new': shape2['coordinates'][coord]
                    }
            
            # Compare other attributes
            for attr in ['type', 'description']:
                if shape1[attr] != shape2[attr]:
                    differences[attr] = {
                        'old': shape1[attr],
                        'new': shape2[attr]
                    }
            
            if differences:
                shape_differences.append({
                    'type': 'modified',
                    'shape_name': name,
                    'differences': differences
                })
    
    # Find deleted shapes
    for name in shapes1_dict:
        if name not in shapes2_dict:
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
