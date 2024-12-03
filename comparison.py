import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.drawing.image import Image
import io
import zipfile
import os
from lxml import etree
import tempfile
import numpy as np

def _get_anchor_coordinates(anchor):
    """アンカー情報から座標を取得する共通関数"""
    try:
        if hasattr(anchor, 'to'):
            # Two cell anchor type
            from_marker = anchor._from if hasattr(anchor, '_from') else anchor.from_marker
            to_marker = anchor.to if hasattr(anchor, 'to') else anchor.to_marker
            
            col = getattr(from_marker, 'col', 0) or 0
            row = getattr(from_marker, 'row', 0) or 0
            
            # サイズの計算
            width = (getattr(to_marker, 'col', col) or col) - col
            height = (getattr(to_marker, 'row', row) or row) - row
            
            return col, row, width, height
        else:
            # One cell anchor type
            col = getattr(anchor, 'col', 0) or 0
            row = getattr(anchor, 'row', 0) or 0
            width = getattr(anchor, 'width', None)
            height = getattr(anchor, 'height', None)
            
            return col, row, width, height
    except Exception as e:
        st.warning(f"アンカー座標の取得中にエラー: {str(e)}")
        return 0, 0, None, None

def _process_drawing(drawing, shape_type='unknown'):
    """描画オブジェクトの処理"""
    try:
        anchor = getattr(drawing, '_anchor', None) or getattr(drawing, 'anchor', None)
        if not anchor:
            return None

        col, row, width, height = _get_anchor_coordinates(anchor)
        
        shape_info = {
            'type': shape_type,
            'x': col,
            'y': row,
            'width': width,
            'height': height,
            'text': getattr(drawing, 'text', '')
        }
        
        # 図形タイプの詳細情報を取得
        if hasattr(drawing, 'style'):
            shape_info['style'] = drawing.style
        if hasattr(drawing, '_type'):
            shape_info['shape_type'] = drawing._type
        elif hasattr(drawing, 'type'):
            shape_info['shape_type'] = drawing.type
            
        return shape_info
    except Exception as e:
        st.warning(f"描画オブジェクトの処理中にエラー: {str(e)}")
        return None

def extract_shape_info(file_path, sheet_name):
    """
    Extract shape information from an Excel sheet
    """
    try:
        wb = openpyxl.load_workbook(file_path)
        sheet = wb[sheet_name]
        shapes = []
        
        for shape in sheet._images:
            if isinstance(shape, Image):
                shape_info = {
                    'type': 'image',
                    'x': shape.anchor._from.col,
                    'y': shape.anchor._from.row,
                    'width': shape.width,
                    'height': shape.height
                }
                shapes.append(shape_info)
        
        wb.close()
        return shapes
    except Exception as e:
        st.error(f"画像情報の抽出中にエラー: {str(e)}")
        return []

def compare_shapes(shapes1, shapes2):
    """
    Compare shapes between two Excel sheets
    """
    differences = []
    
    # Track matched shapes to avoid duplicate comparisons
    matched_shapes2 = set()
    
    # Compare shapes from file1 with shapes from file2
    for shape1 in shapes1:
        match_found = False
        for i, shape2 in enumerate(shapes2):
            if i in matched_shapes2:
                continue
                
            if (shape1['x'] == shape2['x'] and 
                shape1['y'] == shape2['y']):
                # Shapes are in the same position, check for modifications
                if (shape1.get('width') != shape2.get('width') or 
                    shape1.get('height') != shape2.get('height')):
                    differences.append({
                        'type': 'modified',
                        'old_shape': shape1,
                        'new_shape': shape2
                    })
                matched_shapes2.add(i)
                match_found = True
                break
        
        if not match_found:
            differences.append({
                'type': 'deleted',
                'shape': shape1
            })
    
    # Find added shapes (those in file2 that weren't matched)
    for i, shape2 in enumerate(shapes2):
        if i not in matched_shapes2:
            differences.append({
                'type': 'added',
                'shape': shape2
            })
    
    return differences

def calculate_sheet_similarity(df1, df2):
    """
    Calculate similarity between two dataframes
    """
    try:
        # Get common columns
        common_cols = list(set(df1.columns) & set(df2.columns))
        if not common_cols:
            return 0.0
            
        # Calculate column similarity
        column_similarity = len(common_cols) / max(len(df1.columns), len(df2.columns))
        
        # Calculate data similarity
        matching_cells = 0
        total_cells = 0
        
        for col in common_cols:
            df1_col = df1[col].astype(str)
            df2_col = df2[col].astype(str)
            
            # Compare each cell
            matches = (df1_col == df2_col).sum()
            matching_cells += matches
            total_cells += max(len(df1_col), len(df2_col))
        
        data_similarity = matching_cells / total_cells if total_cells > 0 else 0
        
        # Calculate overall similarity
        return (column_similarity + data_similarity) / 2
    except Exception as e:
        st.error(f"類似度計算中にエラー: {str(e)}")
        return 0.0

def find_similar_sheets(sheets1, sheets2, file1_path, file2_path, similarity_threshold=0.95):
    """
    Find similar sheets with different names
    """
    sheet_pairs = []
    renamed_sheets = []
    
    # Find unpaired sheets
    unpaired_sheets1 = set(sheets1) - set(sheets2)
    unpaired_sheets2 = set(sheets2) - set(sheets1)
    
    for sheet1 in unpaired_sheets1:
        df1 = pd.read_excel(file1_path, sheet_name=sheet1)
        
        for sheet2 in unpaired_sheets2:
            df2 = pd.read_excel(file2_path, sheet_name=sheet2)
            
            similarity = calculate_sheet_similarity(df1, df2)
            if similarity >= similarity_threshold:
                sheet_pairs.append((sheet1, sheet2, similarity))
                renamed_sheets.extend([sheet1, sheet2])
                break
    
    return sheet_pairs, renamed_sheets

def compare_dataframes(df1, df2):
    """
    Compare two dataframes and return differences with styling information
    """
    # Convert all columns to string type for comparison
    df1 = df1.astype(str)
    df2 = df2.astype(str)
    
    # Get common columns
    common_cols = list(set(df1.columns) & set(df2.columns))
    
    # Create results dataframes
    df1_result = df1.copy()
    df2_result = df2.copy()
    
    # Initialize style dictionaries
    df1_styles = {col: [''] * len(df1) for col in df1.columns}
    df2_styles = {col: [''] * len(df2) for col in df2.columns}
    
    # Track differences
    differences = []
    
    # Compare rows
    df1['_row_hash'] = df1[common_cols].apply(lambda x: hash(tuple(x)), axis=1)
    df2['_row_hash'] = df2[common_cols].apply(lambda x: hash(tuple(x)), axis=1)
    
    # Find modified rows
    for idx1, row1 in df1.iterrows():
        hash1 = row1['_row_hash']
        matching_rows = df2[df2['_row_hash'] == hash1]
        
        if len(matching_rows) == 0:
            # Row was deleted
            differences.append({
                'type': 'deleted',
                'row_index': idx1,
                'values': row1[common_cols].to_dict()
            })
            # Style deleted row
            for col in df1.columns:
                if col != '_row_hash':
                    df1_styles[col][idx1] = '#F8D7DA'
        else:
            # Row exists in both, check for modifications
            idx2 = matching_rows.index[0]
            for col in common_cols:
                if row1[col] != matching_rows.iloc[0][col]:
                    differences.append({
                        'type': 'modified',
                        'column': col,
                        'row_index_old': idx1,
                        'row_index_new': idx2,
                        'value_old': row1[col],
                        'value_new': matching_rows.iloc[0][col]
                    })
                    df1_styles[col][idx1] = '#FFF3CD'
                    df2_styles[col][idx2] = '#FFF3CD'
    
    # Find added rows
    for idx2, row2 in df2.iterrows():
        hash2 = row2['_row_hash']
        if len(df1[df1['_row_hash'] == hash2]) == 0:
            # Row was added
            differences.append({
                'type': 'added',
                'row_index': idx2,
                'values': row2[common_cols].to_dict()
            })
            # Style added row
            for col in df2.columns:
                if col != '_row_hash':
                    df2_styles[col][idx2] = '#D4EDDA'
    
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