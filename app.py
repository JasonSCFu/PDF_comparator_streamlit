import streamlit as st
from docx import Document
from docx.shared import Inches, RGBColor
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn
import tempfile
import base64
import os
import re
import sqlite3
import csv
import io
from typing import List, Tuple, Dict, Optional
from datetime import datetime

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

# -------------- Database Functions --------------

DB_NAME = "text_records.db"

def init_db():
    """Initialize the SQLite database and create tables if they don't exist."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS text_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            column_a TEXT,
            column_b TEXT,
            column_c TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    
    # Create settings table for column names
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        )
    """)
    
    # Initialize default column names if not exists
    cursor.execute("SELECT COUNT(*) FROM settings WHERE key LIKE 'column_%_name'")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO settings (key, value) VALUES ('column_a_name', 'Column A')")
        cursor.execute("INSERT INTO settings (key, value) VALUES ('column_b_name', 'Column B')")
        cursor.execute("INSERT INTO settings (key, value) VALUES ('column_c_name', 'Column C')")
    
    conn.commit()
    conn.close()

def save_text_record(name: str, text_content: str, column: str = 'column_a') -> bool:
    """Save a new text record or update if name already exists."""
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        
        # Check if record exists
        cursor.execute("SELECT id FROM text_records WHERE name = ?", (name,))
        exists = cursor.fetchone()
        
        if exists:
            # Update existing record
            cursor.execute(f"""
                UPDATE text_records 
                SET {column} = ?, updated_at = ?
                WHERE name = ?
            """, (text_content, datetime.now(), name))
        else:
            # Insert new record
            cursor.execute(f"""
                INSERT INTO text_records (name, {column}, updated_at)
                VALUES (?, ?, ?)
            """, (name, text_content, datetime.now()))
        
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        st.error(f"Error saving record: {str(e)}")
        return False

def get_all_record_names() -> List[str]:
    """Get all record names from the database."""
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM text_records ORDER BY updated_at DESC")
        names = [row[0] for row in cursor.fetchall()]
        conn.close()
        return names
    except Exception as e:
        st.error(f"Error fetching records: {str(e)}")
        return []

def get_text_record(name: str, column: str = 'column_a') -> Optional[str]:
    """Get text content by record name and column."""
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute(f"SELECT {column} FROM text_records WHERE name = ?", (name,))
        result = cursor.fetchone()
        conn.close()
        return result[0] if result else None
    except Exception as e:
        st.error(f"Error fetching record: {str(e)}")
        return None

def get_full_record(name: str) -> Optional[Dict]:
    """Get full record with all columns by name."""
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT name, column_a, column_b, column_c FROM text_records WHERE name = ?", (name,))
        result = cursor.fetchone()
        conn.close()
        if result:
            return {
                'name': result[0],
                'column_a': result[1],
                'column_b': result[2],
                'column_c': result[3]
            }
        return None
    except Exception as e:
        st.error(f"Error fetching record: {str(e)}")
        return None

def delete_text_record(name: str) -> bool:
    """Delete a text record by name."""
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM text_records WHERE name = ?", (name,))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        st.error(f"Error deleting record: {str(e)}")
        return False

def get_all_records() -> List[Dict]:
    """Get all records with metadata."""
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT id, name, column_a, column_b, column_c, created_at, updated_at FROM text_records ORDER BY updated_at DESC")
        records = []
        for row in cursor.fetchall():
            records.append({
                'id': row[0],
                'name': row[1],
                'column_a': row[2],
                'column_b': row[3],
                'column_c': row[4],
                'created_at': row[5],
                'updated_at': row[6]
            })
        conn.close()
        return records
    except Exception as e:
        st.error(f"Error fetching records: {str(e)}")
        return []

def get_column_name(column_key: str) -> str:
    """Get custom column name from settings."""
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT value FROM settings WHERE key = ?", (f"{column_key}_name",))
        result = cursor.fetchone()
        conn.close()
        return result[0] if result else column_key.replace('_', ' ').title()
    except Exception as e:
        return column_key.replace('_', ' ').title()

def get_all_column_names() -> Dict[str, str]:
    """Get all custom column names."""
    return {
        'column_a': get_column_name('column_a'),
        'column_b': get_column_name('column_b'),
        'column_c': get_column_name('column_c')
    }

def set_column_name(column_key: str, name: str) -> bool:
    """Set custom column name in settings."""
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO settings (key, value)
            VALUES (?, ?)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value
        """, (f"{column_key}_name", name))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        st.error(f"Error setting column name: {str(e)}")
        return False

def export_to_csv() -> str:
    """Export all records to CSV format."""
    try:
        records = get_all_records()
        if not records:
            return None
        
        output = io.StringIO()
        fieldnames = ['id', 'name', 'column_a', 'column_b', 'column_c', 'created_at', 'updated_at']
        writer = csv.DictWriter(output, fieldnames=fieldnames)
        
        writer.writeheader()
        for record in records:
            writer.writerow(record)
        
        return output.getvalue()
    except Exception as e:
        st.error(f"Error exporting to CSV: {str(e)}")
        return None

# -------------- Streamlit App --------------

# Initialize database
init_db()

st.set_page_config(page_title="Text Diff Highlighter", layout="wide")

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

# Sidebar for database management
with st.sidebar:
    # Column Settings Section
    with st.expander("⚙️ Column Settings"):
        st.markdown("**Customize column names:**")
        column_names = get_all_column_names()
        
        new_col_a = st.text_input("Column A Name:", value=column_names['column_a'], key="col_a_input")
        new_col_b = st.text_input("Column B Name:", value=column_names['column_b'], key="col_b_input")
        new_col_c = st.text_input("Column C Name:", value=column_names['column_c'], key="col_c_input")
        
        if st.button("Save Column Names", key="save_col_names"):
            if new_col_a and new_col_b and new_col_c:
                set_column_name('column_a', new_col_a)
                set_column_name('column_b', new_col_b)
                set_column_name('column_c', new_col_c)
                st.success("Column names updated!")
                st.rerun()
            else:
                st.warning("All column names must be filled")
    
    st.divider()
    st.header("📚 Saved Text Records")
    
    # Get all saved records and current column names
    saved_names = get_all_record_names()
    column_names = get_all_column_names()
    
    if saved_names:
        st.markdown("**Load a saved record:**")
        selected_record = st.selectbox("Select a record", [""] + saved_names, key="select_record")
        
        if selected_record:
            # Show record details
            record_data = get_full_record(selected_record)
            if record_data:
                st.markdown("**Record contents:**")
                for col_name in ['column_a', 'column_b', 'column_c']:
                    if record_data[col_name]:
                        preview = record_data[col_name][:50] + "..." if len(record_data[col_name]) > 50 else record_data[col_name]
                        st.caption(f"**{column_names[col_name]}:** {preview}")
                
                st.markdown("**Select column to load:**")
                load_column = st.radio("Column", ['column_a', 'column_b', 'column_c'], 
                                      format_func=lambda x: column_names[x],
                                      key="load_column",
                                      horizontal=True)
                
                col_load1, col_load2 = st.columns(2)
                with col_load1:
                    if st.button("Load to Text A", key="load_a"):
                        text_content = get_text_record(selected_record, load_column)
                        if text_content is not None:
                            st.session_state["ta"] = text_content
                            st.success(f"Loaded to Text A")
                            st.rerun()
                        else:
                            st.warning(f"{column_names[load_column]} is empty")
                
                with col_load2:
                    if st.button("Load to Text B", key="load_b"):
                        text_content = get_text_record(selected_record, load_column)
                        if text_content is not None:
                            st.session_state["tb"] = text_content
                            st.success(f"Loaded to Text B")
                            st.rerun()
                        else:
                            st.warning(f"{column_names[load_column]} is empty")
                
                st.divider()
                if st.button("🗑️ Delete Record", key="delete_btn", type="secondary", use_container_width=True):
                    if delete_text_record(selected_record):
                        st.success(f"Deleted '{selected_record}'")
                        st.rerun()
        
        st.divider()
        st.markdown(f"**Total records:** {len(saved_names)}")
        
        # CSV Export
        st.divider()
        if st.button("📥 Export Database to CSV", key="export_csv", use_container_width=True):
            csv_data = export_to_csv()
            if csv_data:
                st.download_button(
                    label="Download CSV",
                    data=csv_data,
                    file_name="text_records_export.csv",
                    mime="text/csv",
                    use_container_width=True
                )
    else:
        st.info("No saved records yet. Save your text below!")

st.markdown("Enter text in the areas below to compare and highlight differences while preserving original formatting.")

# Get current column names
column_names = get_all_column_names()

col1, col2 = st.columns(2)

with col1:
    st.markdown("**Text A**")
    text_a = st.text_area("", height=200, key="ta", label_visibility="collapsed", placeholder="Enter first text here...")
    
    # Save Text A
    with st.expander("💾 Save Text A"):
        save_name_a = st.text_input("Enter a name for this text:", key="save_name_a")
        save_column_a = st.selectbox("Select column:", ['column_a', 'column_b', 'column_c'],
                                     format_func=lambda x: column_names[x],
                                     key="save_column_a")
        if st.button("Save Text A", key="save_btn_a"):
            if save_name_a and text_a:
                if save_text_record(save_name_a, text_a, save_column_a):
                    st.success(f"Saved as '{save_name_a}' in {column_names[save_column_a]}")
                    st.rerun()
            elif not save_name_a:
                st.warning("Please enter a name")
            else:
                st.warning("Text A is empty")

with col2:
    st.markdown("**Text B**")
    text_b = st.text_area("", height=200, key="tb", label_visibility="collapsed", placeholder="Enter second text here...")
    
    # Save Text B
    with st.expander("💾 Save Text B"):
        save_name_b = st.text_input("Enter a name for this text:", key="save_name_b")
        save_column_b = st.selectbox("Select column:", ['column_a', 'column_b', 'column_c'],
                                     format_func=lambda x: column_names[x],
                                     key="save_column_b")
        if st.button("Save Text B", key="save_btn_b"):
            if save_name_b and text_b:
                if save_text_record(save_name_b, text_b, save_column_b):
                    st.success(f"Saved as '{save_name_b}' in {column_names[save_column_b]}")
                    st.rerun()
            elif not save_name_b:
                st.warning("Please enter a name")
            else:
                st.warning("Text B is empty")

st.caption("Tip: The text areas will preserve line breaks and spacing in the comparison.")

run = st.button("Compare Texts", type="primary")

if run:
    if not text_a and not text_b:
        st.warning("Please enter text in both areas before comparing.")
    else:
        a_html, b_html = diff_to_html(text_a, text_b)

        st.subheader("Differences")

        col_a, col_b = st.columns(2)
        
        with col_a:
            st.markdown("**Text A (with highlights)**")
            st.markdown(f"<div class='panel'>{a_html}</div>", unsafe_allow_html=True)
        
        with col_b:
            st.markdown("**Text B (with highlights)**")
            st.markdown(f"<div class='panel'>{b_html}</div>", unsafe_allow_html=True)

        # Generate and provide download
        st.divider()
        make_download_button_docx(text_a, text_b, "highlighted_diff.docx")
