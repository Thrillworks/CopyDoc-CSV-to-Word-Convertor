# CopyDoc-CSV-to-Word-Convertor

A Streamlit web application for converting between CSV and Word documents for Figma copy management workflows.

## Features

- **CSV to Word**: Convert CSV data to formatted Word documents
- **Word to CSV**: Extract content from Word documents to CSV format
- **Interactive UI**: User-friendly web interface built with Streamlit
- **File Upload/Download**: Easy file handling with drag-and-drop support

## Project Structure

```
figma_streamlit_app/
├── app.py                    # Main Streamlit application
├── requirements.txt          # Python dependencies
├── README.md                # This file
├── .streamlit/              # Streamlit configuration
│   └── config.toml          # App configuration
├── src/                     # Source code
│   └── figma_copy_workflow/ # Core functionality modules
│       ├── __init__.py
│       ├── parser.py        # CSV/Word conversion logic
│       ├── helpers.py       # Utility functions
│       └── cli.py           # Command-line interface
├── pages/                   # Additional Streamlit pages (if needed)
├── data/                    # Sample data files
├── assets/                  # Static assets (images, CSS, etc.)
└── tests/                   # Test files
```

## Installation

1. **Clone or download this project**
   ```bash
   git clone <repository-url>
   cd figma_streamlit_app
   ```

2. **Create a virtual environment (recommended)**
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

## Running the Application

### Local Development
```bash
streamlit run app.py
```

The app will open in your browser at `http://localhost:8501`

### Production Deployment

#### Streamlit Community Cloud
1. Push your code to a GitHub repository
2. Go to [Streamlit Community Cloud](https://streamlit.io/cloud)
3. Connect your GitHub account and deploy the repository
4. Set the main file path to `app.py`

#### Docker Deployment
```dockerfile
FROM python:3.9-slim

WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt

COPY . .

EXPOSE 8501

CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
```

#### Other Cloud Platforms
- **Heroku**: Add a `Procfile` with `web: streamlit run app.py --server.port=$PORT --server.address=0.0.0.0`
- **Azure**: Use Azure Container Apps or App Service
- **AWS**: Deploy using ECS, Lambda, or EC2

## Usage

1. **CSV to Word Conversion**:
   - Select "CSV to Word" mode
   - Upload a CSV file
   - Choose title and content columns
   - Download the generated Word document

2. **Word to CSV Conversion**:
   - Select "Word to CSV" mode
   - Upload a Word document (.docx)
   - Download the extracted CSV data

## Dependencies

- `streamlit>=1.47.1` - Web app framework
- `pandas>=2.3.1` - Data manipulation
- `python-docx>=1.1.0` - Word document processing

## Development

To extend the application:

1. **Add new pages**: Create Python files in the `pages/` directory
2. **Modify core logic**: Edit files in `src/figma_copy_workflow/`
3. **Add assets**: Place images, CSS files in `assets/`
4. **Add tests**: Create test files in `tests/`

## Configuration

Customize the app appearance and behavior by editing `.streamlit/config.toml`:

- Theme colors
- Server settings
- Browser behavior

## Troubleshooting

**Import Errors**: Make sure all dependencies are installed with `pip install -r requirements.txt`

**File Upload Issues**: Check file formats (CSV files should be UTF-8 encoded, Word files should be .docx format)

**Performance**: For large files, consider processing in chunks or adding progress indicators

## License

This project is open source and available under the MIT License.