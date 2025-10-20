import streamlit as st
import PyPDF2
import fitz  # PyMuPDF
import difflib
from io import BytesIO
from PIL import Image, ImageDraw
import numpy as np
from streamlit_drawable_canvas import st_canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_LEFT

st.set_page_config(page_title="PDF Text Comparator", page_icon="📄", layout="wide")

def pdf_to_image(pdf_file, page_num, zoom=2):
    """Convert a single PDF page to image"""
    try:
        pdf_file.seek(0)
        pdf_document = fitz.open(stream=pdf_file.read(), filetype="pdf")
        
        if page_num < 0 or page_num >= len(pdf_document):
            return None, 0, 0
        
        page = pdf_document[page_num]
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        # Ensure image is in RGB mode for compatibility
        if img.mode != 'RGB':
            img = img.convert('RGB')
        
        pdf_document.close()
        return img, pix.width, pix.height
    except Exception as e:
        st.error(f"Error converting PDF: {str(e)}")
        return None, 0, 0

def extract_text_from_region(pdf_file, page_num, x0, y0, x1, y1, zoom=2):
    """Extract text from a specific region of a PDF page"""
    try:
        pdf_file.seek(0)
        pdf_document = fitz.open(stream=pdf_file.read(), filetype="pdf")
        
        if page_num < 0 or page_num >= len(pdf_document):
            return ""
        
        page = pdf_document[page_num]
        
        # Adjust coordinates for zoom level
        rect = fitz.Rect(x0/zoom, y0/zoom, x1/zoom, y1/zoom)
        
        # Extract text from the rectangle
        text = page.get_text("text", clip=rect)
        
        pdf_document.close()
        return text.strip()
    except Exception as e:
        st.error(f"Error extracting text: {str(e)}")
        return ""

def get_pdf_page_count(pdf_file):
    """Get total number of pages in PDF"""
    try:
        pdf_file.seek(0)
        pdf_document = fitz.open(stream=pdf_file.read(), filetype="pdf")
        count = len(pdf_document)
        pdf_document.close()
        return count
    except:
        return 0

def highlight_all_differences(text1, text2):
    """Compare two texts and return formatted output with ALL differences highlighted"""
    def tokenize_with_newlines(text):
        """Split text into words and newline tokens"""
        tokens = []
        for line in text.split('\n'):
            if line.strip():
                tokens.extend(line.split())
                tokens.append('\n')
            else:
                tokens.append('\n')
        return tokens
    
    tokens1 = tokenize_with_newlines(text1)
    tokens2 = tokenize_with_newlines(text2)
    
    diff = difflib.SequenceMatcher(None, tokens1, tokens2)
    
    html_output = []
    diff_data = []
    
    for tag, i1, i2, j1, j2 in diff.get_opcodes():
        if tag == 'equal':
            for token in tokens2[j1:j2]:
                if token == '\n':
                    html_output.append('<br>')
                    diff_data.append(('newline', '\n'))
                else:
                    html_output.append(token)
                    diff_data.append(('normal', token))
        elif tag == 'insert':
            for token in tokens2[j1:j2]:
                if token == '\n':
                    html_output.append('<br>')
                    diff_data.append(('newline', '\n'))
                else:
                    html_output.append(f'<span style="background-color: #90EE90; font-weight: bold; padding: 2px;">{token}</span>')
                    diff_data.append(('insert', token))
        elif tag == 'delete':
            for token in tokens1[i1:i2]:
                if token == '\n':
                    html_output.append('<br>')
                    diff_data.append(('newline', '\n'))
                else:
                    html_output.append(f'<span style="background-color: #FFB6C1; text-decoration: line-through; padding: 2px;">{token}</span>')
                    diff_data.append(('delete', token))
        elif tag == 'replace':
            for token in tokens1[i1:i2]:
                if token == '\n':
                    html_output.append('<br>')
                    diff_data.append(('newline', '\n'))
                else:
                    html_output.append(f'<span style="background-color: #FFB6C1; text-decoration: line-through; padding: 2px;">{token}</span>')
                    diff_data.append(('delete', token))
            for token in tokens2[j1:j2]:
                if token == '\n':
                    html_output.append('<br>')
                    diff_data.append(('newline', '\n'))
                else:
                    html_output.append(f'<span style="background-color: #90EE90; font-weight: bold; padding: 2px;">{token}</span>')
                    diff_data.append(('insert', token))
    
    return ' '.join(html_output), diff_data

def create_pdf_with_highlights(diff_data, pdf1_name="PDF 1", pdf2_name="PDF 2"):
    """Create a PDF document with highlighted differences"""
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, topMargin=0.5*inch, bottomMargin=0.5*inch)
    
    elements = []
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=18,
        textColor=colors.HexColor('#333333'),
        spaceAfter=20,
        alignment=TA_LEFT
    )
    
    legend_style = ParagraphStyle(
        'Legend',
        parent=styles['Normal'],
        fontSize=11,
        spaceAfter=20,
        alignment=TA_LEFT
    )
    
    content_style = ParagraphStyle(
        'Content',
        parent=styles['Normal'],
        fontSize=11,
        leading=18,
        alignment=TA_LEFT,
        wordWrap='CJK'
    )
    
    elements.append(Paragraph("PDF Text Comparison Report", title_style))
    elements.append(Spacer(1, 0.2*inch))
    
    legend_text = '''
    <font color="black"><b>Legend:</b></font><br/>
    <font color="green"><b>Green/Bold = Added Text</b></font><br/>
    <font color="red"><strike>Red/Strikethrough = Deleted Text</strike></font>
    '''
    elements.append(Paragraph(legend_text, legend_style))
    elements.append(Spacer(1, 0.3*inch))
    
    elements.append(Paragraph(f"<b>Comparison Result</b>", styles['Heading2']))
    elements.append(Spacer(1, 0.2*inch))
    
    current_paragraph = []
    
    for tag, token in diff_data:
        if tag == 'newline':
            if current_paragraph:
                para_text = ' '.join(current_paragraph)
                elements.append(Paragraph(para_text, content_style))
                current_paragraph = []
        elif tag == 'insert':
            current_paragraph.append(f'<font color="green"><b>{token}</b></font>')
        elif tag == 'delete':
            current_paragraph.append(f'<font color="red"><strike>{token}</strike></font>')
        else:
            current_paragraph.append(token)
    
    if current_paragraph:
        para_text = ' '.join(current_paragraph)
        elements.append(Paragraph(para_text, content_style))
    
    doc.build(elements)
    buffer.seek(0)
    return buffer

# Main app
st.title("📄 PDF Text Comparator with Interactive Selection")
st.markdown("Upload PDFs, draw rectangles to select text regions, and compare with highlighted differences.")

# Initialize session state
for key in ['pdf1_file', 'pdf2_file', 'pdf1_name', 'pdf2_name', 'pdf1_page', 'pdf2_page',
            'pdf1_img', 'pdf2_img', 'pdf1_width', 'pdf2_width', 'pdf1_height', 'pdf2_height',
            'selected_text1', 'selected_text2', 'pdf1_total_pages', 'pdf2_total_pages']:
    if key not in st.session_state:
        if 'name' in key:
            st.session_state[key] = "PDF 1" if '1' in key else "PDF 2"
        elif 'page' in key or 'width' in key or 'height' in key or 'total' in key:
            st.session_state[key] = 0
        else:
            st.session_state[key] = None

# Sidebar
with st.sidebar:
    st.header("📤 Upload & Select")
    
    # PDF 1
    st.subheader("📕 PDF 1")
    uploaded_file1 = st.file_uploader("Upload first PDF", type="pdf", key="pdf1")
    
    if uploaded_file1:
        st.session_state.pdf1_file = uploaded_file1
        st.session_state.pdf1_name = uploaded_file1.name
        
        if st.session_state.pdf1_total_pages == 0:
            st.session_state.pdf1_total_pages = get_pdf_page_count(uploaded_file1)
        
        st.info(f"📄 {st.session_state.pdf1_total_pages} pages")
        
        page_num1 = st.selectbox(
            "Select page",
            range(1, st.session_state.pdf1_total_pages + 1),
            key="page_select1"
        )
        
        if st.button("Load Page", key="load1", type="primary"):
            img, w, h = pdf_to_image(uploaded_file1, page_num1 - 1)
            if img:
                st.session_state.pdf1_img = img
                st.session_state.pdf1_width = w
                st.session_state.pdf1_height = h
                st.session_state.pdf1_page = page_num1 - 1
                st.success("✓ Page loaded")
    
    st.markdown("---")
    
    # PDF 2
    st.subheader("📘 PDF 2")
    uploaded_file2 = st.file_uploader("Upload second PDF", type="pdf", key="pdf2")
    
    if uploaded_file2:
        st.session_state.pdf2_file = uploaded_file2
        st.session_state.pdf2_name = uploaded_file2.name
        
        if st.session_state.pdf2_total_pages == 0:
            st.session_state.pdf2_total_pages = get_pdf_page_count(uploaded_file2)
        
        st.info(f"📄 {st.session_state.pdf2_total_pages} pages")
        
        page_num2 = st.selectbox(
            "Select page",
            range(1, st.session_state.pdf2_total_pages + 1),
            key="page_select2"
        )
        
        if st.button("Load Page", key="load2", type="primary"):
            img, w, h = pdf_to_image(uploaded_file2, page_num2 - 1)
            if img:
                st.session_state.pdf2_img = img
                st.session_state.pdf2_width = w
                st.session_state.pdf2_height = h
                st.session_state.pdf2_page = page_num2 - 1
                st.success("✓ Page loaded")
    
    st.markdown("---")
    st.info("💡 Draw rectangles on the PDF to select text areas for comparison")

# Main content
col1, col2 = st.columns(2)

with col1:
    st.subheader(f"📕 {st.session_state.pdf1_name}")
    
    if st.session_state.pdf1_img:
        # Convert PIL Image to numpy array for better compatibility with Posit Connect
        img_array1 = np.array(st.session_state.pdf1_img)
        
        canvas_result1 = st_canvas(
            fill_color="rgba(255, 165, 0, 0.3)",
            stroke_width=2,
            stroke_color="#FF0000",
            background_image=Image.fromarray(img_array1),
            height=st.session_state.pdf1_height,
            width=st.session_state.pdf1_width,
            drawing_mode="rect",
            update_streamlit=True,
            key="canvas1",
        )
        
        if st.button("Extract Selected Text", key="extract1", type="primary"):
            if canvas_result1.json_data is not None:
                objects = canvas_result1.json_data.get("objects", [])
                if objects:
                    # Get the last drawn rectangle
                    rect = objects[-1]
                    x0 = rect['left']
                    y0 = rect['top']
                    x1 = x0 + rect['width']
                    y1 = y0 + rect['height']
                    
                    text = extract_text_from_region(
                        st.session_state.pdf1_file,
                        st.session_state.pdf1_page,
                        x0, y0, x1, y1
                    )
                    
                    if text:
                        st.session_state.selected_text1 = text
                        st.success(f"✓ Extracted {len(text)} characters")
                    else:
                        st.warning("No text found in selection")
                else:
                    st.warning("Please draw a rectangle first")
        
        if st.session_state.selected_text1:
            st.text_area(
                "Selected Text (editable)",
                value=st.session_state.selected_text1,
                height=200,
                key="text1_display"
            )
            st.session_state.selected_text1 = st.session_state.text1_display
    else:
        st.info("👆 Upload and load a PDF page above")

with col2:
    st.subheader(f"📘 {st.session_state.pdf2_name}")
    
    if st.session_state.pdf2_img:
        # Convert PIL Image to numpy array for better compatibility with Posit Connect
        img_array2 = np.array(st.session_state.pdf2_img)
        
        canvas_result2 = st_canvas(
            fill_color="rgba(255, 165, 0, 0.3)",
            stroke_width=2,
            stroke_color="#0000FF",
            background_image=Image.fromarray(img_array2),
            height=st.session_state.pdf2_height,
            width=st.session_state.pdf2_width,
            drawing_mode="rect",
            update_streamlit=True,
            key="canvas2",
        )
        
        if st.button("Extract Selected Text", key="extract2", type="primary"):
            if canvas_result2.json_data is not None:
                objects = canvas_result2.json_data.get("objects", [])
                if objects:
                    rect = objects[-1]
                    x0 = rect['left']
                    y0 = rect['top']
                    x1 = x0 + rect['width']
                    y1 = y0 + rect['height']
                    
                    text = extract_text_from_region(
                        st.session_state.pdf2_file,
                        st.session_state.pdf2_page,
                        x0, y0, x1, y1
                    )
                    
                    if text:
                        st.session_state.selected_text2 = text
                        st.success(f"✓ Extracted {len(text)} characters")
                    else:
                        st.warning("No text found in selection")
                else:
                    st.warning("Please draw a rectangle first")
        
        if st.session_state.selected_text2:
            st.text_area(
                "Selected Text (editable)",
                value=st.session_state.selected_text2,
                height=200,
                key="text2_display"
            )
            st.session_state.selected_text2 = st.session_state.text2_display
    else:
        st.info("👆 Upload and load a PDF page above")

# Comparison section
st.markdown("---")

if st.session_state.selected_text1 and st.session_state.selected_text2:
    if st.button("🔍 Compare Selected Texts", type="primary", use_container_width=True):
        with st.spinner("Comparing texts..."):
            html_output, diff_data = highlight_all_differences(
                st.session_state.selected_text1,
                st.session_state.selected_text2
            )
            
            st.header("📊 Comparison Result")
            
            st.markdown("""
            <div style='padding: 15px; background-color: #f0f0f0; border-radius: 5px; margin-bottom: 20px;'>
                <strong style='font-size: 16px;'>Legend:</strong><br/>
                <span style='background-color: #90EE90; padding: 3px 8px; margin: 5px 5px 5px 10px; font-weight: bold; border-radius: 3px;'>Green = Added Text</span>
                <span style='background-color: #FFB6C1; padding: 3px 8px; margin: 5px; text-decoration: line-through; border-radius: 3px;'>Red/Strikethrough = Deleted Text</span>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown(f"""
            <div style='padding: 25px; border: 2px solid #4CAF50; border-radius: 8px; background-color: white; 
                        max-height: 600px; overflow-y: auto; line-height: 2.0; font-size: 15px;'>
                {html_output}
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown("---")
            st.subheader("⬇️ Download Report")
            
            pdf_buffer = create_pdf_with_highlights(diff_data, st.session_state.pdf1_name, st.session_state.pdf2_name)
            
            st.download_button(
                label="📥 Download as PDF",
                data=pdf_buffer,
                file_name="comparison_report.pdf",
                mime="application/pdf",
                type="primary",
                use_container_width=True
            )
else:
    st.info("👆 Select and extract text from both PDFs to enable comparison")

# Instructions
with st.expander("ℹ️ How to Use"):
    st.markdown("""
    ### Step-by-Step Guide:
    
    1. **Upload PDFs**: Upload two PDF files in the sidebar
    2. **Select Pages**: Choose which page to view from each PDF
    3. **Load Pages**: Click "Load Page" to display the PDF pages
    4. **Draw Selection**:
       - Use your mouse to draw rectangles on the PDF images
       - Select the text areas you want to compare
       - You can draw multiple rectangles (last one will be used)
    5. **Extract Text**: Click "Extract Selected Text" button
    6. **Review**: Check the extracted text in the text area (editable)
    7. **Compare**: Once both texts are extracted, click "Compare Selected Texts"
    8. **Download**: Save the comparison report as PDF
    
    ### Tips:
    - 🖱️ Draw rectangles by clicking and dragging on the PDF
    - ✏️ You can edit the extracted text before comparison
    - 🔄 Load different pages to select text from multiple areas
    - 🎨 Green = additions, Red/strikethrough = deletions
    - 📏 Try to draw accurate rectangles around the text you want
    """)
