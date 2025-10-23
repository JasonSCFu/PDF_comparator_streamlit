import streamlit as st
from docx import Document
from docx.shared import Inches, RGBColor
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn
import tempfile
import base64
import os
import re
from typing import List, Tuple

# -------------- Utilities --------------

def html_escape(s: str) -> str:
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )

def to_preserved_html(text: str) -> str:
    """Convert plain text to HTML preserving spaces and line breaks."""
    escaped = html_escape(text)
    # Replace spaces and tabs with non-breaking spaces to preserve alignment
    escaped = escaped.replace("  ", " &nbsp;")
    escaped = escaped.replace("\t", "&nbsp;&nbsp;&nbsp;&nbsp;")
    # Convert line breaks to <br>
    escaped = escaped.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "<br>")
    return f"<div class='preserve'>{escaped}</div>"

def tokenize_words_keep_ws(text: str) -> List[str]:
    """Tokenize into words but keep whitespace tokens to allow format-preserving diff highlighting."""
    # Split on word boundaries, keep whitespace and punctuation
    tokens = re.findall(r"\w+|\s+|[^\w\s]", text, flags=re.UNICODE)
    return tokens

def diff_to_html(a: str, b: str) -> Tuple[str, str]:
    """Return (html_a, html_b) with span-based highlights for insert/delete/replace while preserving formatting."""
    import difflib

    a_tokens = tokenize_words_keep_ws(a)
    b_tokens = tokenize_words_keep_ws(b)

    sm = difflib.SequenceMatcher(a=a_tokens, b=b_tokens, autojunk=False)

    a_out: List[str] = []
    b_out: List[str] = []

    def wrap(token: str, cls: str) -> str:
        # Preserve whitespace faithfully
        if token.isspace():
            token_html = token.replace(" ", "&nbsp;").replace("\t", "&nbsp;&nbsp;&nbsp;&nbsp;").replace("\n", "<br>")
        else:
            token_html = html_escape(token)
        return f"<span class='{cls}'>{token_html}</span>"

    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'equal':
            seg_a = ''.join(a_tokens[i1:i2])
            seg_b = ''.join(b_tokens[j1:j2])
            # Equal segments shown normally
            a_out.append(html_escape(seg_a).replace(" ", "&nbsp;").replace("\t", "&nbsp;&nbsp;&nbsp;&nbsp;").replace("\n", "<br>"))
            b_out.append(html_escape(seg_b).replace(" ", "&nbsp;").replace("\t", "&nbsp;&nbsp;&nbsp;&nbsp;").replace("\n", "<br>"))
        elif tag == 'replace':
            for tok in a_tokens[i1:i2]:
                a_out.append(wrap(tok, 'replace-a'))
            for tok in b_tokens[j1:j2]:
                b_out.append(wrap(tok, 'replace-b'))
        elif tag == 'delete':
            for tok in a_tokens[i1:i2]:
                a_out.append(wrap(tok, 'del'))
        elif tag == 'insert':
            for tok in b_tokens[j1:j2]:
                b_out.append(wrap(tok, 'ins'))

    a_html = ''.join(a_out)
    b_html = ''.join(b_out)

    # Wrap in containers and ensure line breaks are preserved
    def wrap_container(content: str) -> str:
        return f"<div class='preserve'>{content}</div>"

    return wrap_container(a_html), wrap_container(b_html)

def add_highlight(run, highlight_type):
    """Add highlight color to a run in the Word document."""
    if highlight_type == 'del':
        run.font.strike = True
        run.font.color.rgb = RGBColor(255, 0, 0)  # Red for deletions
    elif highlight_type == 'ins':
        run.font.highlight_color = 4  # Green highlight for insertions
    elif highlight_type == 'replace-a':
        run.font.strike = True
        run.font.color.rgb = RGBColor(255, 165, 0)  # Orange for replaced text
    elif highlight_type == 'replace-b':
        run.font.highlight_color = 5  # Yellow highlight for replacement

def make_download_button_docx(a: str, b: str, filename: str) -> None:
    """Generate a Word document with highlighted differences."""
    try:
        doc = Document()
        doc.add_heading('Text Comparison with Highlights', 0)

        # Add Text A
        doc.add_heading('Text A', level=1)
        a_tokens = tokenize_words_keep_ws(a)
        b_tokens = tokenize_words_keep_ws(b)
        import difflib
        sm = difflib.SequenceMatcher(a=a_tokens, b=b_tokens, autojunk=False)

        current_paragraph = doc.add_paragraph()
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == 'equal':
                text = ''.join(a_tokens[i1:i2])
                # Split by newlines and add to paragraphs
                parts = text.split('\n')
                for i, part in enumerate(parts):
                    if i > 0:
                        current_paragraph = doc.add_paragraph()
                    run = current_paragraph.add_run(part)
            elif tag == 'replace':
                for tok in a_tokens[i1:i2]:
                    if tok == '\n':
                        current_paragraph = doc.add_paragraph()
                    else:
                        run = current_paragraph.add_run(tok)
                        add_highlight(run, 'replace-a')
            elif tag == 'delete':
                for tok in a_tokens[i1:i2]:
                    if tok == '\n':
                        current_paragraph = doc.add_paragraph()
                    else:
                        run = current_paragraph.add_run(tok)
                        add_highlight(run, 'del')

        # Add Text B
        doc.add_heading('Text B', level=1)
        current_paragraph = doc.add_paragraph()
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == 'equal':
                text = ''.join(b_tokens[j1:j2])
                # Split by newlines and add to paragraphs
                parts = text.split('\n')
                for i, part in enumerate(parts):
                    if i > 0:
                        current_paragraph = doc.add_paragraph()
                    run = current_paragraph.add_run(part)
            elif tag == 'replace':
                for tok in b_tokens[j1:j2]:
                    if tok == '\n':
                        current_paragraph = doc.add_paragraph()
                    else:
                        run = current_paragraph.add_run(tok)
                        add_highlight(run, 'replace-b')
            elif tag == 'insert':
                for tok in b_tokens[j1:j2]:
                    if tok == '\n':
                        current_paragraph = doc.add_paragraph()
                    else:
                        run = current_paragraph.add_run(tok)
                        add_highlight(run, 'ins')

        # Save to a temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            doc.save(tmp.name)
            with open(tmp.name, 'rb') as f:
                data = f.read()
        st.download_button("Download highlighted Word Document", data=data, file_name=filename, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        os.unlink(tmp.name)
    except Exception as e:
        st.error(f"Error generating Word document: {str(e)}")

# -------------- Streamlit App --------------

st.set_page_config(page_title="Text Diff Highlighter", layout="centered")

st.markdown(
    """
    <style>
    .preserve { white-space: pre-wrap; font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace; line-height: 1.5; }
    .del { background: #ffe0e0; text-decoration: line-through; }
    .ins { background: #e0ffe0; }
    .replace-a { background: #fff0cc; text-decoration: line-through; }
    .replace-b { background: #e6f0ff; }
    .panel { border: 1px solid #ddd; border-radius: 6px; padding: 12px; background: #fafafa; }
    .heading { font-weight: 600; margin-bottom: 8px; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Compare Text with Highlighted Differences")

st.markdown("Enter text in the areas below to compare and highlight differences while preserving original formatting.")

col1, col2 = st.columns(2)
with col1:
    st.markdown("**Text A**")
    text_a = st.text_area("", height=200, key="ta", label_visibility="collapsed", placeholder="Enter first text here...")

with col2:
    st.markdown("**Text B**")
    text_b = st.text_area("", height=200, key="tb", label_visibility="collapsed", placeholder="Enter second text here...")

st.caption("Tip: The text areas will preserve line breaks and spacing in the comparison.")

run = st.button("Compare Texts")

if run:
    if not text_a and not text_b:
        st.warning("Please enter text in both areas before comparing.")
    else:
        a_html, b_html = diff_to_html(text_a, text_b)

        st.subheader("Differences")

        st.markdown("**Text A (with highlights)**")
        st.markdown(f"<div class='panel'>{a_html}</div>", unsafe_allow_html=True)

        st.markdown("**Text B (with highlights)**")
        st.markdown(f"<div class='panel'>{b_html}</div>", unsafe_allow_html=True)

        # Generate and provide download
        make_download_button_docx(text_a, text_b, "highlighted_diff.docx")
