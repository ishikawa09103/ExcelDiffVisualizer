import streamlit as st

def apply_custom_css():
    st.markdown("""
        <style>
            .stApp {
                max-width: 100%;
                padding: 1rem;
            }
            
            .ag-cell-modified {
                background-color: #FFF3CD !important;
            }
            
            .ag-cell-added {
                background-color: #D4EDDA !important;
            }
            
            .ag-cell-deleted {
                background-color: #F8D7DA !important;
            }
            
            .ag-theme-streamlit {
                --ag-header-background-color: #f1f3f4;
                --ag-header-foreground-color: #333;
                --ag-header-cell-hover-background-color: #dddfe1;
                --ag-row-hover-color: #f8f9fa;
            }
            
            .streamlit-expanderHeader {
                background-color: #f8f9fa;
                border-radius: 4px;
            }
        </style>
    """, unsafe_allow_html=True)
