# 📄 PDF Text Comparator with Interactive Selection

A Streamlit web application that displays PDF files with **interactive text selection**. Draw rectangles on PDF pages to select text regions, then compare with highlighted differences. Additions shown in green, deletions in red with strikethrough.

## Features

- **🖱️ Interactive Selection**: Draw rectangles directly on PDF images to select text
- **📄 PDF Viewer**: View actual PDF pages rendered as high-quality images
- **📕📘 Two PDF Upload**: Upload and compare two different PDF documents
- **📄 Page-by-Page Selection**: Choose specific pages and select text regions
- **🎯 Precise Text Extraction**: Extract text from user-drawn rectangular regions
- **📊 Side-by-Side Layout**: View both PDFs simultaneously for easy selection
- **✏️ Editable Text**: Review and edit extracted text before comparison
- **🔍 Complete Comparison**: Shows ALL differences (additions AND deletions)
- **🎨 Color Highlighting**: 
  - Green = New/added text
  - Red with strikethrough = Deleted/removed text
- **📐 Format Preservation**: Maintains original line breaks and paragraph structure
- **📥 PDF Export**: Download comparison report as a formatted PDF document

## Installation

1. **Clone or download this project**

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

   **Note**: 
   - PyMuPDF (fitz) may require additional system dependencies on some platforms
   - streamlit-drawable-canvas requires a modern web browser with Canvas API support

## Usage

1. **Start the application**:
   ```bash
   streamlit run app.py
   ```

2. **Upload PDFs**:
   - In the sidebar, upload **PDF 1** and **PDF 2**
   - Total page count will be displayed

3. **Select and Load Pages**:
   - For each PDF:
     - Choose which page to view from the dropdown
     - Click "Load Page" button
     - PDF page will render as an image

4. **Draw Selection Rectangles**:
   - Use your mouse to draw rectangles on the PDF images
   - Click and drag to select the text area you want to compare
   - The selected area will be highlighted with a colored rectangle
   - You can draw multiple rectangles (the last one will be used)

5. **Extract Text**:
   - Click "Extract Selected Text" button below each PDF
   - Text from the selected region will be extracted
   - Extracted text appears in an editable text area

6. **Review and Edit** (Optional):
   - Check the extracted text for accuracy
   - Edit if needed before comparison

7. **Compare**:
   - Once both texts are extracted, click "Compare Selected Texts"
   - View comparison with **all differences** highlighted:
     - Green = Added text
     - Red/strikethrough = Deleted text
   - Original formatting is preserved

8. **Download**:
   - Click "Download as PDF" button
   - Get a formatted PDF file with color-coded highlighting
   - Open with any PDF reader

## Use Cases

- **Contract Clause Comparison**: Select and compare specific clauses between contract versions
- **Legal Document Review**: Highlight and compare critical sections in legal documents
- **Partial Document Comparison**: Compare only relevant sections, not entire documents
- **Paragraph-Level Analysis**: Focus on specific paragraphs or sections
- **Selective Review**: Choose exactly which text regions to compare
- **Multi-page Comparison**: Select text from different pages of the same document

## Requirements

- Python 3.7+
- streamlit
- PyPDF2
- PyMuPDF (fitz)
- pdf2image
- Pillow
- reportlab
- streamlit-drawable-canvas

## Technical Details

- **Interactive Canvas**: Uses `streamlit-drawable-canvas` for drawing selections on PDF images
- **PDF Rendering**: Uses `PyMuPDF (fitz)` for high-quality PDF to image conversion (2x zoom)
- **Region-Based Extraction**: Coordinate-based text extraction using PyMuPDF's clip rectangles
- **Text Comparison**: Python's `difflib.SequenceMatcher` for intelligent difference detection
- **PDF Generation**: `reportlab` library for creating formatted PDF documents
- **Selection Method**: Rectangle drawing with left-click-drag interaction
- **Comparison Level**: Word-level comparison for precise detection of all differences
- **Format Preservation**: Special tokenization preserves line breaks and paragraph structure
- **Highlighting Strategy**: Shows both additions (green) and deletions (red/strikethrough)
- **Output Formats**: 
  - Real-time HTML preview in browser
  - Downloadable PDF document with color-coded highlighting

## Notes

- **Drawing Rectangles**: Click and drag on the PDF image to draw selection rectangles
- **Selection Tips**: Try to draw rectangles that closely fit the text you want to extract
- **Text Accuracy**: Text extraction works best with clear, selectable text (not scanned images)
- **Multiple Rectangles**: If you draw multiple rectangles, only the last one will be used
- **Page Selection**: You can load different pages to select text from various parts of the document
- **Editing**: Extracted text can be edited in the text area before comparison
- **Performance**: PDF rendering may take a moment for large files
- **Browser Compatibility**: Requires a modern browser with HTML5 Canvas support
- **Coordinate Precision**: The 2x zoom ensures accurate text extraction from selected regions
- **Format Preservation**: Line breaks and structure are maintained in the comparison

## License

MIT License - Feel free to use and modify as needed.
