import pandas as pd
import numpy as np

def compare_dataframes(df1, df2):
    """
    Compare two dataframes and return styled versions with differences highlighted
    """
    # Create copies for styling
    df1_styled = df1.copy()
    df2_styled = df2.copy()
    
    # Initialize style dataframes
    df1_style = pd.DataFrame('', index=df1.index, columns=df1.columns)
    df2_style = pd.DataFrame('', index=df2.index, columns=df2.columns)
    
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
                    df2_style.loc[idx, col] = 'added'
                    differences.append({
                        'type': 'added',
                        'column': col,
                        'row': idx,
                        'value': val2
                    })
            elif not pd.isna(val1) and pd.isna(val2):
                # Deleted in df2
                if idx < len(df1):
                    df1_style.loc[idx, col] = 'deleted'
                    differences.append({
                        'type': 'deleted',
                        'column': col,
                        'row': idx,
                        'value': val1
                    })
            elif not pd.isna(val1) and not pd.isna(val2) and val1 != val2:
                # Modified
                if idx < len(df1):
                    df1_style.loc[idx, col] = 'modified'
                if idx < len(df2):
                    df2_style.loc[idx, col] = 'modified'
                differences.append({
                    'type': 'modified',
                    'column': col,
                    'row': idx,
                    'value_old': val1,
                    'value_new': val2
                })
    
    # Apply styles
    def apply_styles(val, style):
        if style == 'modified':
            return 'background-color: #FFF3CD'
        elif style == 'added':
            return 'background-color: #D4EDDA'
        elif style == 'deleted':
            return 'background-color: #F8D7DA'
        return ''
    
    df1_styled = df1_styled.style.apply(lambda x: df1_style.applymap(apply_styles))
    df2_styled = df2_styled.style.apply(lambda x: df2_style.applymap(apply_styles))
    
    # Create difference summary
    diff_summary = pd.DataFrame(differences)
    
    return {
        'df1_styled': df1_styled,
        'df2_styled': df2_styled,
        'diff_summary': diff_summary
    }
