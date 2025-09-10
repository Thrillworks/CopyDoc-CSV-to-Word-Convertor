"""Streamlit UI for Figma Copy Workflow"""

import streamlit as st
import tempfile
import os
import sys
from pathlib import Path
from io import StringIO
import pandas as pd

# Import the figma workflow modules
sys.path.append(str(Path(__file__).parent / "src"))

try:
    from figma_copy_workflow.parser import csv_to_word, word_to_csv, word_to_csv_new
except ImportError as e:
    st.error(f"Error importing figma_copy_workflow modules: {e}")
    st.stop()

def main():
    """Main Streamlit application"""
    st.set_page_config(
        page_title="CopyDoc CSV to Word Convertor",
        page_icon="🎨",
        layout="wide"
    )
    
    st.title("🎨 CopyDoc CSV to Word Convertor")
    st.markdown("Convert between CSV and Word documents for copy management.")
    
    # Mode selection
    st.sidebar.title("🛠️ Conversion Mode")
    mode = st.sidebar.radio(
        "Select conversion mode:",
        ["CSV to Word", "Word to CSV", "Word to New CSV"],
        help="Choose the direction of conversion"
    )
    
    
    # Formatting toggle
    st.sidebar.title("✨ Text Formatting")
    preserve_formatting = st.sidebar.toggle(
        "Preserve formatting",
        value=True,
        help="Enable to preserve bold, italic, and list formatting as Markdown. Disable for plain text."
    )
    
    if preserve_formatting:
        st.sidebar.success("📝 Formatting: **Markdown** (bold, *italic*, lists)")
    else:
        st.sidebar.info("📄 Formatting: Plain text only")
    
    if mode == "CSV to Word":
        csv_to_word_ui(preserve_formatting)
    elif mode == "Word to CSV":
        word_to_csv_ui(preserve_formatting)
    else:
        word_to_new_csv_ui(preserve_formatting)

def csv_to_word_ui(preserve_formatting: bool):
    """UI for CSV to Word conversion"""
    st.header("📊 CSV to Word Document")
    st.markdown("Upload a CSV file to convert it into a formatted Word document.")
    
    # Show formatting info
    if preserve_formatting:
        st.info("ℹ️ Text formatting will be preserved as Markdown when extracting from Word documents.")
    else:
        st.info("ℹ️ Only plain text will be extracted from Word documents (no formatting).")
    
    uploaded_csv = st.file_uploader(
        "Upload CSV file",
        type=['csv'],
        help="Select a CSV file to convert to Word format"
    )
    
    if uploaded_csv is not None:
        # Preview CSV
        try:
            # Read CSV for preview
            csv_content = uploaded_csv.read().decode('utf-8')
            uploaded_csv.seek(0)  # Reset file pointer
            
            df = pd.read_csv(StringIO(csv_content))
            st.subheader("📋 CSV Preview")
            st.dataframe(df.head(), use_container_width=True)
            
            if st.button("🔄 Convert to Word", type="primary"):
                try:
                    # Save CSV to temporary file
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.csv', mode='w', encoding='utf-8') as tmp_csv:
                        tmp_csv.write(csv_content)
                        tmp_csv_path = tmp_csv.name
                    
                    # Create temporary Word file
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_word:
                        tmp_word_path = tmp_word.name
                    
                    # Convert CSV to Word
                    with st.spinner("🔄 Converting CSV to Word..."):
                        csv_to_word(tmp_csv_path, tmp_word_path)
                    
                    st.success("✅ Conversion completed successfully!")
                    
                    # Provide download
                    with open(tmp_word_path, 'rb') as f:
                        st.download_button(
                            label="💾 Download Word Document",
                            data=f.read(),
                            file_name=f"{uploaded_csv.name.rsplit('.', 1)[0]}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    
                    # Cleanup
                    os.unlink(tmp_csv_path)
                    os.unlink(tmp_word_path)
                    
                except Exception as e:
                    st.error(f"❌ Error converting CSV to Word: {str(e)}")
                    # Cleanup on error
                    for path in [tmp_csv_path, tmp_word_path]:
                        try:
                            if 'path' in locals():
                                os.unlink(path)
                        except:
                            pass
                        
        except Exception as e:
            st.error(f"❌ Error reading CSV file: {str(e)}")

def word_to_csv_ui(preserve_formatting: bool):
    """UI for Word to CSV conversion"""
    st.header("📄 Word Document to CSV")
    st.markdown("Upload the original CSV and modified Word document to extract changes back to CSV format.")
    
    # Show formatting info
    if preserve_formatting:
        st.info("ℹ️ Text formatting will be preserved as Markdown (**bold**, *italic*, - lists) when extracting from Word.")
    else:
        st.info("ℹ️ Only plain text will be extracted from Word documents (no formatting).")
    
    col1, col2 = st.columns(2)
    
    with col1:
        original_csv = st.file_uploader(
            "Upload Original CSV",
            type=['csv'],
            help="The original CSV file used to create the Word document",
            key="original_csv"
        )
    
    with col2:
        modified_word = st.file_uploader(
            "Upload Modified Word Document",
            type=['docx'],
            help="The Word document with modifications",
            key="modified_word"
        )
    
    if original_csv is not None and modified_word is not None:
        # Preview original CSV
        try:
            original_csv_content = original_csv.read().decode('utf-8')
            original_csv.seek(0)  # Reset file pointer
            
            df = pd.read_csv(StringIO(original_csv_content))
            st.subheader("📋 Original CSV Preview")
            st.dataframe(df.head(), use_container_width=True)
            
            if st.button("🔄 Extract Changes to CSV", type="primary"):
                try:
                    # Save files to temporary locations
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.csv', mode='w', encoding='utf-8') as tmp_orig_csv:
                        tmp_orig_csv.write(original_csv_content)
                        tmp_orig_csv_path = tmp_orig_csv.name
                    
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_word:
                        tmp_word.write(modified_word.read())
                        tmp_word_path = tmp_word.name
                    
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.csv') as tmp_output_csv:
                        tmp_output_csv_path = tmp_output_csv.name
                    
                    # Convert Word back to CSV
                    with st.spinner("🔄 Extracting changes from Word to CSV..."):
                        word_to_csv(tmp_orig_csv_path, tmp_word_path, tmp_output_csv_path, preserve_formatting)
                    
                    st.success("✅ Changes extracted successfully!")
                    
                    # Show preview of updated CSV
                    updated_df = pd.read_csv(tmp_output_csv_path)
                    st.subheader("📊 Updated CSV Preview")
                    st.dataframe(updated_df.head(), use_container_width=True)
                    
                    # Provide download
                    with open(tmp_output_csv_path, 'rb') as f:
                        st.download_button(
                            label="💾 Download Updated CSV",
                            data=f.read(),
                            file_name=f"{original_csv.name.rsplit('.', 1)[0]}_updated.csv",
                            mime="text/csv"
                        )
                    
                    # Cleanup
                    for path in [tmp_orig_csv_path, tmp_word_path, tmp_output_csv_path]:
                        os.unlink(path)
                        
                except Exception as e:
                    st.error(f"❌ Error extracting changes: {str(e)}")
                    # Cleanup on error
                    for path in [tmp_orig_csv_path, tmp_word_path, tmp_output_csv_path]:
                        try:
                            if 'path' in locals():
                                os.unlink(path)
                        except:
                            pass
                        
        except Exception as e:
            st.error(f"❌ Error reading original CSV: {str(e)}")

def word_to_new_csv_ui(preserve_formatting: bool):
    """UI for Word to New CSV conversion"""
    st.header("📄 Word Document to New CSV")
    st.markdown("Upload a Word document to convert its content into a new CSV file with the standard format.")
    
    # Show formatting info
    if preserve_formatting:
        st.info("ℹ️ Text formatting will be preserved as Markdown (**bold**, *italic*, - lists) when extracting from Word.")
    else:
        st.info("ℹ️ Only plain text will be extracted from Word documents (no formatting).")
    
    uploaded_word = st.file_uploader(
        "Upload Word Document",
        type=['docx'],
        help="Select a Word document to convert to CSV format",
        key="new_csv_word"
    )
    
    if uploaded_word is not None:
        st.subheader("📋 Document Information")
        st.write(f"**Filename:** {uploaded_word.name}")
        st.write(f"**File size:** {len(uploaded_word.read())} bytes")
        uploaded_word.seek(0)  # Reset file pointer
        
        if st.button("🔄 Convert to CSV", type="primary"):
            try:
                # Save Word document to temporary file
                with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_word:
                    tmp_word.write(uploaded_word.read())
                    tmp_word_path = tmp_word.name
                
                with tempfile.NamedTemporaryFile(delete=False, suffix='.csv') as tmp_csv:
                    tmp_csv_path = tmp_csv.name
                
                # Convert Word to CSV
                with st.spinner("🔄 Converting Word document to CSV..."):
                    word_to_csv_new(tmp_word_path, tmp_csv_path, preserve_formatting)
                
                st.success("✅ Conversion completed successfully!")
                
                # Show preview of generated CSV
                csv_df = pd.read_csv(tmp_csv_path)
                st.subheader("📊 Generated CSV Preview")
                st.dataframe(csv_df, use_container_width=True)
                
                # Show some statistics
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total Rows", len(csv_df))
                with col2:
                    unique_groups = csv_df['group'].nunique()
                    st.metric("Unique Groups", unique_groups)
                with col3:
                    non_empty_text = (csv_df['figma_text'].str.strip() != '').sum()
                    st.metric("Content Rows", non_empty_text)
                
                # Provide download
                with open(tmp_csv_path, 'rb') as f:
                    st.download_button(
                        label="💾 Download CSV File",
                        data=f.read(),
                        file_name=f"{uploaded_word.name.rsplit('.', 1)[0]}_export.csv",
                        mime="text/csv"
                    )
                
                # Cleanup
                for path in [tmp_word_path, tmp_csv_path]:
                    os.unlink(path)
                    
            except Exception as e:
                st.error(f"❌ Error converting Word to CSV: {str(e)}")
                # Cleanup on error
                for path in [tmp_word_path, tmp_csv_path]:
                    try:
                        if 'path' in locals():
                            os.unlink(path)
                    except:
                        pass

if __name__ == "__main__":
    main() 