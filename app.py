"""Streamlit UI for Figma Copy Workflow"""

import streamlit as st
import tempfile
import os
import pandas as pd
import sys

# Set upload limit to 1GB
st.set_option('server.maxUploadSize', 1024)

# Inject Open Graph meta tags for social sharing
st.markdown(
    '''
    <meta property="og:title" content="CopyDoc CSV to Word Convertor" />
    <meta property="og:description" content="Convert between CSV and Word documents for copy management. Streamlined workflow for copy teams." />
    ''',
    unsafe_allow_html=True
)

# Add src directory to Python path for imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

# Import the figma workflow modules
try:
    from figma_copy_workflow.parser import csv_to_word, word_to_csv
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
        ["CSV to Word", "Word to CSV"],
        help="Choose the direction of conversion"
    )
    
    if mode == "CSV to Word":
        csv_to_word_interface()
    else:
        word_to_csv_interface()

def csv_to_word_interface():
    """Interface for CSV to Word conversion"""
    st.header("📄 CSV to Word Document")
    st.markdown("Upload a CSV file to convert it to a Word document.")
    
    # File upload
    uploaded_file = st.file_uploader(
        "Choose a CSV file",
        type=['csv'],
        help="Upload a CSV file with your copy data"
    )
    
    if uploaded_file is not None:
        try:
            # Read the CSV data
            df = pd.read_csv(uploaded_file)
            
            # Display preview
            st.subheader("📊 Data Preview")
            st.dataframe(df.head(10), use_container_width=True)
            st.info(f"Total rows: {len(df)}")
            
            # Convert button
            if st.button("🔄 Convert to Word", type="primary"):
                with st.spinner("Converting CSV to Word document..."):
                    try:
                        # Create temporary files
                        with tempfile.NamedTemporaryFile(suffix='.csv', delete=False, mode='w', encoding='utf-8') as tmp_csv:
                            # Write CSV content to temporary file
                            uploaded_file.seek(0)
                            csv_content = uploaded_file.read().decode('utf-8')
                            tmp_csv.write(csv_content)
                            tmp_csv_path = tmp_csv.name
                        
                        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp_file:
                            output_path = tmp_file.name
                        
                        # Perform conversion
                        csv_to_word(tmp_csv_path, output_path)
                        
                        # Read the generated file
                        with open(output_path, 'rb') as f:
                            word_data = f.read()
                        
                        # Clean up temporary files
                        os.unlink(tmp_csv_path)
                        os.unlink(output_path)
                        
                        # Provide download button
                        st.success("✅ Conversion completed successfully!")
                        st.download_button(
                            label="📥 Download Word Document",
                            data=word_data,
                            file_name=f"{uploaded_file.name.rsplit('.', 1)[0]}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                        
                    except Exception as e:
                        st.error(f"❌ Error during conversion: {str(e)}")
                        
        except Exception as e:
            st.error(f"❌ Error reading CSV file: {str(e)}")

def word_to_csv_interface():
    """Interface for Word to CSV conversion"""
    st.header("📊 Word Document to CSV")
    st.markdown("Upload the original CSV and modified Word document to extract changes back to CSV format.")
    
    col1, col2 = st.columns(2)
    
    with col1:
        original_csv = st.file_uploader(
            "Upload Original CSV",
            type=['csv'],
            help="The original CSV file used to create the Word document",
            key="original_csv"
        )
    
    with col2:
        uploaded_file = st.file_uploader(
            "Upload Modified Word Document",
            type=['docx'],
            help="The Word document with modifications",
            key="modified_word"
        )
    
    if original_csv is not None and uploaded_file is not None:
        try:
            # Preview original CSV
            original_csv_content = original_csv.read().decode('utf-8')
            original_csv.seek(0)  # Reset file pointer
            
            df = pd.read_csv(original_csv)
            st.subheader("📋 Original CSV Preview")
            st.dataframe(df.head(), use_container_width=True)
            
            # Convert button
            if st.button("🔄 Extract Changes to CSV", type="primary"):
                with st.spinner("Extracting changes from Word to CSV..."):
                    try:
                        # Create temporary files
                        with tempfile.NamedTemporaryFile(suffix='.csv', delete=False, mode='w', encoding='utf-8') as tmp_orig_csv:
                            tmp_orig_csv.write(original_csv_content)
                            tmp_orig_csv_path = tmp_orig_csv.name
                        
                        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp_word:
                            tmp_word.write(uploaded_file.read())
                            tmp_word_path = tmp_word.name
                        
                        with tempfile.NamedTemporaryFile(suffix='.csv', delete=False) as tmp_output:
                            output_path = tmp_output.name
                        
                        # Perform conversion
                        word_to_csv(tmp_orig_csv_path, tmp_word_path, output_path)
                        
                        # Read the generated CSV
                        df = pd.read_csv(output_path)
                        
                        # Display results
                        st.success("✅ Changes extracted successfully!")
                        st.subheader("📊 Updated CSV Preview")
                        st.dataframe(df, use_container_width=True)
                        st.info(f"Total rows: {len(df)}")
                        
                        # Provide download button
                        csv_data = df.to_csv(index=False)
                        st.download_button(
                            label="📥 Download Updated CSV",
                            data=csv_data,
                            file_name=f"{original_csv.name.rsplit('.', 1)[0]}_updated.csv",
                            mime="text/csv"
                        )
                        
                        # Clean up temporary files
                        os.unlink(tmp_orig_csv_path)
                        os.unlink(tmp_word_path)
                        os.unlink(output_path)
                        
                    except Exception as e:
                        st.error(f"❌ Error extracting changes: {str(e)}")
                        # Cleanup on error
                        for path in [tmp_orig_csv_path, tmp_word_path, output_path]:
                            try:
                                if 'path' in locals():
                                    os.unlink(path)
                            except Exception:
                                pass
                        
        except Exception as e:
            st.error(f"❌ Error reading original CSV: {str(e)}")

# Add footer
def add_footer():
    """Add footer with additional information"""
    st.markdown("---")
    st.markdown(
        """
        <div style='text-align: center; color: #666;'>
            <p>🎨 CopyDoc CSV to Word Convertor</p>
            <p>Built with Streamlit for efficient copy management</p>
        </div>
        """,
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
    add_footer()