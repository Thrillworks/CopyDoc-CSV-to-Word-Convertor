from .helpers import read_csv_data, group_data_by_section, create_word_document, read_word_document_data, update_csv_with_word_changes, write_csv_data, extract_word_document_to_csv_format


def csv_to_word(csv_file_path: str, word_file_path: str):
    """Convert CSV data to a formatted Word document.
    
    Args:
        csv_file_path: Path to the input CSV file
        word_file_path: Path where the output Word document should be saved
    """
    # Read and process CSV data
    data = read_csv_data(csv_file_path)
    
    # Group data by section
    grouped_data = group_data_by_section(data)
    
    # Create Word document
    create_word_document(grouped_data, word_file_path)


def word_to_csv(origin_csv_file_path: str, word_file_path: str, csv_file_path: str, preserve_formatting: bool = True):
    """Convert updated Word document back to CSV format.
    
    This function takes the original CSV file and a Word document (created by csv_to_word)
    that may have been edited, and creates a new CSV file with the updated text content
    while preserving the original CSV structure.
    
    Args:
        origin_csv_file_path: Path to the original CSV file
        word_file_path: Path to the (potentially edited) Word document
        csv_file_path: Path where the updated CSV file should be saved
        preserve_formatting: Whether to preserve formatting as Markdown (True) or extract plain text (False)
    """
    # Read the original CSV data
    original_csv_data = read_csv_data(origin_csv_file_path)
    
    # Extract updated text content from Word document, mapped by ID
    word_updates = read_word_document_data(word_file_path, preserve_formatting)
    
    # Update the original CSV data with changes from Word document
    updated_csv_data = update_csv_with_word_changes(original_csv_data, word_updates)
    
    # Write the updated data to new CSV file
    write_csv_data(updated_csv_data, csv_file_path)


def word_to_csv_new(word_file_path: str, csv_file_path: str, preserve_formatting: bool = True):
    """Convert Word document directly to CSV format.
    
    This function takes a Word document and creates a new CSV file with the expected
    structure (id, frame, group, layer_name, figma_text, round_2, round_3) by extracting
    content from the Word document.
    
    Args:
        word_file_path: Path to the Word document to convert
        csv_file_path: Path where the output CSV file should be saved
        preserve_formatting: Whether to preserve formatting as Markdown (True) or extract plain text (False)
    """
    # Extract content from Word document in CSV format
    csv_data = extract_word_document_to_csv_format(word_file_path, preserve_formatting)
    
    # Write the data to CSV file
    write_csv_data(csv_data, csv_file_path)