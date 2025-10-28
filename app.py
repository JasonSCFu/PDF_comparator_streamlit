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
import json
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

def normalize_text_for_comparison(text: str) -> str:
    """Normalize text for comparison by removing backspaces, slash content, and normalizing whitespace."""
    # Remove backspace characters
    text_no_backspace = text.replace('\b', '').replace('\x08', '')
    # Remove text surrounded by slashes before normalization (global replacement)
    text_no_slash = re.sub(r'/[^/]*/', '', text_no_backspace, flags=re.DOTALL)
    # Replace all whitespace sequences (spaces, tabs, newlines, etc.) with a single space
    # and strip leading/trailing whitespace
    normalized = re.sub(r'\s+', ' ', text_no_slash).strip()
    return normalized

def highlight_text_diff(a: str, b: str, highlight_for: str) -> str:
    """Highlight differences in text while preserving original formatting.
    
    Args:
        a: First text
        b: Second text
        highlight_for: 'a' to highlight what's unique to a (deletions), 'b' for what's unique to b (insertions)
    
    Returns:
        HTML string with highlighted differences
    """
    import difflib
    
    # Normalize text for comparison
    a_normalized = normalize_text_for_comparison(a)
    b_normalized = normalize_text_for_comparison(b)

    # Split into words for comparison
    a_words = a_normalized.split()
    b_words = b_normalized.split()

    # Create a matcher that ignores whitespace
    matcher = difflib.SequenceMatcher(lambda x: x == ' ', a_words, b_words, autojunk=False)

    # For display, preserve original formatting but remove slash content
    text_to_display = a if highlight_for == 'a' else b
    display_text = re.sub(r'/[^/]*/', '', text_to_display.replace('\b', '').replace('\x08', ''), flags=re.DOTALL)

    # If texts are identical after normalization, return original formatting
    if a_normalized == b_normalized:
        # Convert to HTML-safe format
        escaped = html_escape(display_text)
        escaped = escaped.replace(" ", "&nbsp;").replace("\t", "&nbsp;&nbsp;&nbsp;&nbsp;").replace("\n", "<br>")
        return escaped

    # Split display text into words while preserving positions
    display_words = []
    word_positions = []

    # Find word positions in the display text
    for word in re.finditer(r'\S+', display_text):
        display_words.append(word.group())
        word_positions.append((word.start(), word.end()))

    # Create result by reconstructing text with highlights
    result = []
    last_end = 0
    word_index = 0

    # Process each opcode from the matcher
    for opcode, i1, i2, j1, j2 in matcher.get_opcodes():
        if opcode == 'equal':
            # Add words without highlighting, preserving original spacing
            target_words = i2 - i1 if highlight_for == 'a' else j2 - j1
            for _ in range(target_words):
                if word_index < len(word_positions):
                    start, end = word_positions[word_index]
                    # Add any whitespace/formatting before this word (HTML-escaped)
                    spacing = display_text[last_end:start]
                    escaped_spacing = html_escape(spacing).replace(" ", "&nbsp;").replace("\t", "&nbsp;&nbsp;&nbsp;&nbsp;").replace("\n", "<br>")
                    result.append(escaped_spacing)
                    # Add the word itself (HTML-escaped)
                    result.append(html_escape(display_words[word_index]))
                    last_end = end
                    word_index += 1

        elif (opcode == 'delete' and highlight_for == 'a') or (opcode == 'insert' and highlight_for == 'b'):
            # Highlight differences
            target_words = i2 - i1 if highlight_for == 'a' else j2 - j1
            css_class = 'del' if highlight_for == 'a' else 'ins'

            for _ in range(target_words):
                if word_index < len(word_positions):
                    start, end = word_positions[word_index]
                    # Add any whitespace/formatting before this word
                    spacing = display_text[last_end:start]
                    escaped_spacing = html_escape(spacing).replace(" ", "&nbsp;").replace("\t", "&nbsp;&nbsp;&nbsp;&nbsp;").replace("\n", "<br>")
                    result.append(escaped_spacing)
                    # Add the highlighted word
                    result.append(f'<span class="{css_class}">{html_escape(display_words[word_index])}</span>')
                    last_end = end
                    word_index += 1

        elif opcode == 'replace':
            if highlight_for == 'a':
                # Highlight deleted words in text A
                for _ in range(i2 - i1):
                    if word_index < len(word_positions):
                        start, end = word_positions[word_index]
                        spacing = display_text[last_end:start]
                        escaped_spacing = html_escape(spacing).replace(" ", "&nbsp;").replace("\t", "&nbsp;&nbsp;&nbsp;&nbsp;").replace("\n", "<br>")
                        result.append(escaped_spacing)
                        result.append(f'<span class="replace-a">{html_escape(display_words[word_index])}</span>')
                        last_end = end
                        word_index += 1
            elif highlight_for == 'b':
                # Highlight inserted words in text B
                for _ in range(j2 - j1):
                    if word_index < len(word_positions):
                        start, end = word_positions[word_index]
                        spacing = display_text[last_end:start]
                        escaped_spacing = html_escape(spacing).replace(" ", "&nbsp;").replace("\t", "&nbsp;&nbsp;&nbsp;&nbsp;").replace("\n", "<br>")
                        result.append(escaped_spacing)
                        result.append(f'<span class="replace-b">{html_escape(display_words[word_index])}</span>')
                        last_end = end
                        word_index += 1

    # Add any remaining text after the last word
    if last_end < len(display_text):
        remaining = display_text[last_end:]
        escaped_remaining = html_escape(remaining).replace(" ", "&nbsp;").replace("\t", "&nbsp;&nbsp;&nbsp;&nbsp;").replace("\n", "<br>")
        result.append(escaped_remaining)

    return ''.join(result)

def diff_to_html(a: str, b: str) -> Tuple[str, str]:
    """Return (html_a, html_b) with span-based highlights for insert/delete/replace while preserving formatting."""
    # Use the complex comparison logic from compare_logic.py
    a_html = highlight_text_diff(a, b, 'a')
    b_html = highlight_text_diff(a, b, 'b')
    
    # Wrap in containers
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
    """Generate a Word document with highlighted differences using the complex comparison logic."""
    try:
        import difflib
        doc = Document()
        doc.add_heading('Text Comparison with Highlights', 0)

        # Normalize text for comparison
        a_normalized = normalize_text_for_comparison(a)
        b_normalized = normalize_text_for_comparison(b)
        
        # Split into words for comparison
        a_words = a_normalized.split()
        b_words = b_normalized.split()
        
        # Create matcher
        matcher = difflib.SequenceMatcher(lambda x: x == ' ', a_words, b_words, autojunk=False)

        # Helper function to process text for document
        def add_text_with_highlights(text: str, highlight_for: str):
            """Add text to document with highlights based on comparison."""
            # Prepare display text (remove backspaces and slash content)
            display_text = re.sub(r'/[^/]*/', '', text.replace('\b', '').replace('\x08', ''), flags=re.DOTALL)
            
            # Split display text into words while preserving positions
            display_words = []
            word_positions = []
            
            for word in re.finditer(r'\S+', display_text):
                display_words.append(word.group())
                word_positions.append((word.start(), word.end()))
            
            current_paragraph = doc.add_paragraph()
            last_end = 0
            word_index = 0
            
            # Process each opcode from the matcher
            for opcode, i1, i2, j1, j2 in matcher.get_opcodes():
                if opcode == 'equal':
                    target_words = i2 - i1 if highlight_for == 'a' else j2 - j1
                    for _ in range(target_words):
                        if word_index < len(word_positions):
                            start, end = word_positions[word_index]
                            # Add spacing before word
                            spacing = display_text[last_end:start]
                            if spacing:
                                parts = spacing.split('\n')
                                for i, part in enumerate(parts):
                                    if i > 0:
                                        current_paragraph = doc.add_paragraph()
                                    if part:
                                        current_paragraph.add_run(part)
                            # Add the word
                            current_paragraph.add_run(display_words[word_index])
                            last_end = end
                            word_index += 1
                
                elif (opcode == 'delete' and highlight_for == 'a') or (opcode == 'insert' and highlight_for == 'b'):
                    target_words = i2 - i1 if highlight_for == 'a' else j2 - j1
                    highlight_type = 'del' if highlight_for == 'a' else 'ins'
                    
                    for _ in range(target_words):
                        if word_index < len(word_positions):
                            start, end = word_positions[word_index]
                            # Add spacing before word
                            spacing = display_text[last_end:start]
                            if spacing:
                                parts = spacing.split('\n')
                                for i, part in enumerate(parts):
                                    if i > 0:
                                        current_paragraph = doc.add_paragraph()
                                    if part:
                                        current_paragraph.add_run(part)
                            # Add highlighted word
                            run = current_paragraph.add_run(display_words[word_index])
                            add_highlight(run, highlight_type)
                            last_end = end
                            word_index += 1
                
                elif opcode == 'replace':
                    if highlight_for == 'a':
                        # Highlight deleted words
                        for _ in range(i2 - i1):
                            if word_index < len(word_positions):
                                start, end = word_positions[word_index]
                                spacing = display_text[last_end:start]
                                if spacing:
                                    parts = spacing.split('\n')
                                    for i, part in enumerate(parts):
                                        if i > 0:
                                            current_paragraph = doc.add_paragraph()
                                        if part:
                                            current_paragraph.add_run(part)
                                run = current_paragraph.add_run(display_words[word_index])
                                add_highlight(run, 'replace-a')
                                last_end = end
                                word_index += 1
                    elif highlight_for == 'b':
                        # Highlight inserted words
                        for _ in range(j2 - j1):
                            if word_index < len(word_positions):
                                start, end = word_positions[word_index]
                                spacing = display_text[last_end:start]
                                if spacing:
                                    parts = spacing.split('\n')
                                    for i, part in enumerate(parts):
                                        if i > 0:
                                            current_paragraph = doc.add_paragraph()
                                        if part:
                                            current_paragraph.add_run(part)
                                run = current_paragraph.add_run(display_words[word_index])
                                add_highlight(run, 'replace-b')
                                last_end = end
                                word_index += 1
            
            # Add any remaining text
            if last_end < len(display_text):
                remaining = display_text[last_end:]
                parts = remaining.split('\n')
                for i, part in enumerate(parts):
                    if i > 0:
                        current_paragraph = doc.add_paragraph()
                    if part:
                        current_paragraph.add_run(part)

        # Add Text A with highlights
        doc.add_heading('Text A', level=1)
        add_text_with_highlights(a, 'a')
        
        # Add Text B with highlights
        doc.add_heading('Text B', level=1)
        add_text_with_highlights(b, 'b')

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
    
    # New flexible schema with JSON storage
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS text_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            columns_data TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    
    # Create settings table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        )
    """)
    
    # Initialize default column configuration if not exists
    cursor.execute("SELECT value FROM settings WHERE key = 'columns_config'")
    if not cursor.fetchone():
        default_columns = [
            {"id": "col_1", "name": "Column 1"},
            {"id": "col_2", "name": "Column 2"},
            {"id": "col_3", "name": "Column 3"}
        ]
        cursor.execute("INSERT INTO settings (key, value) VALUES ('columns_config', ?)", 
                      (json.dumps(default_columns),))
    
    conn.commit()
    conn.close()

def get_columns_config() -> List[Dict]:
    """Get the current column configuration."""
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT value FROM settings WHERE key = 'columns_config'")
        result = cursor.fetchone()
        conn.close()
        if result:
            return json.loads(result[0])
        return [{"id": "col_1", "name": "Column 1"}]
    except Exception as e:
        st.error(f"Error getting columns config: {str(e)}")
        return [{"id": "col_1", "name": "Column 1"}]

def set_columns_config(columns: List[Dict]) -> bool:
    """Set the column configuration."""
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO settings (key, value)
            VALUES ('columns_config', ?)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value
        """, (json.dumps(columns),))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        st.error(f"Error setting columns config: {str(e)}")
        return False

def save_text_record(name: str, text_content: str, column_id: str) -> bool:
    """Save a new text record or update if name already exists."""
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        
        # Check if record exists
        cursor.execute("SELECT columns_data FROM text_records WHERE name = ?", (name,))
        result = cursor.fetchone()
        
        if result:
            # Update existing record
            columns_data = json.loads(result[0]) if result[0] else {}
            columns_data[column_id] = text_content
            cursor.execute("""
                UPDATE text_records 
                SET columns_data = ?, updated_at = ?
                WHERE name = ?
            """, (json.dumps(columns_data), datetime.now(), name))
        else:
            # Insert new record
            columns_data = {column_id: text_content}
            cursor.execute("""
                INSERT INTO text_records (name, columns_data, updated_at)
                VALUES (?, ?, ?)
            """, (name, json.dumps(columns_data), datetime.now()))
        
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

def get_text_record(name: str, column_id: str) -> Optional[str]:
    """Get text content by record name and column ID."""
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT columns_data FROM text_records WHERE name = ?", (name,))
        result = cursor.fetchone()
        conn.close()
        if result and result[0]:
            columns_data = json.loads(result[0])
            return columns_data.get(column_id)
        return None
    except Exception as e:
        st.error(f"Error fetching record: {str(e)}")
        return None

def get_full_record(name: str) -> Optional[Dict]:
    """Get full record with all columns by name."""
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT name, columns_data FROM text_records WHERE name = ?", (name,))
        result = cursor.fetchone()
        conn.close()
        if result:
            columns_data = json.loads(result[1]) if result[1] else {}
            return {
                'name': result[0],
                'columns': columns_data
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
        cursor.execute("SELECT id, name, columns_data, created_at, updated_at FROM text_records ORDER BY updated_at DESC")
        records = []
        for row in cursor.fetchall():
            columns_data = json.loads(row[2]) if row[2] else {}
            record = {
                'id': row[0],
                'name': row[1],
                'created_at': row[3],
                'updated_at': row[4]
            }
            # Add each column as a separate field for CSV export
            record.update(columns_data)
            records.append(record)
        conn.close()
        return records
    except Exception as e:
        st.error(f"Error fetching records: {str(e)}")
        return []


def export_to_csv() -> str:
    """Export all records to CSV format."""
    try:
        records = get_all_records()
        if not records:
            return None
        
        # Get all unique fieldnames from all records
        fieldnames = ['id', 'name']
        columns_config = get_columns_config()
        for col in columns_config:
            fieldnames.append(col['id'])
        fieldnames.extend(['created_at', 'updated_at'])
        
        output = io.StringIO()
        writer = csv.DictWriter(output, fieldnames=fieldnames, extrasaction='ignore')
        
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
    with st.expander("⚙️ Manage Columns", expanded=False):
        st.markdown("**Configure your columns:**")
        columns_config = get_columns_config()
        
        # Initialize session state for columns if not exists
        if 'temp_columns' not in st.session_state:
            st.session_state.temp_columns = columns_config.copy()
        
        # Display current columns with edit/delete options
        for i, col in enumerate(st.session_state.temp_columns):
            col1, col2 = st.columns([4, 1])
            with col1:
                new_name = st.text_input(f"Column {i+1}", value=col['name'], key=f"col_name_{i}")
                st.session_state.temp_columns[i]['name'] = new_name
            with col2:
                if len(st.session_state.temp_columns) > 1:
                    if st.button("❌", key=f"del_col_{i}"):
                        st.session_state.temp_columns.pop(i)
                        st.rerun()
        
        # Add new column button
        col_add, col_save = st.columns(2)
        with col_add:
            if st.button("➕ Add Column", key="add_col", use_container_width=True):
                new_id = f"col_{len(st.session_state.temp_columns) + 1}"
                st.session_state.temp_columns.append({"id": new_id, "name": f"Column {len(st.session_state.temp_columns) + 1}"})
                st.rerun()
        
        with col_save:
            if st.button("💾 Save", key="save_cols", type="primary", use_container_width=True):
                # Validate that all columns have names
                if all(col['name'].strip() for col in st.session_state.temp_columns):
                    # Update IDs for consistency
                    for i, col in enumerate(st.session_state.temp_columns):
                        col['id'] = f"col_{i+1}"
                    if set_columns_config(st.session_state.temp_columns):
                        st.success("Columns updated!")
                        st.rerun()
                else:
                    st.warning("All columns must have names")
    
    st.divider()
    st.header("📚 Saved Text Records")
    
    # Get all saved records and current column config
    saved_names = get_all_record_names()
    columns_config = get_columns_config()
    
    if saved_names:
        st.markdown("**Load a saved record:**")
        selected_record = st.selectbox("Select a record", [""] + saved_names, key="select_record")
        
        if selected_record:
            # Show record details
            record_data = get_full_record(selected_record)
            if record_data:
                st.markdown("**Record contents:**")
                for col in columns_config:
                    col_id = col['id']
                    if col_id in record_data['columns'] and record_data['columns'][col_id]:
                        preview = record_data['columns'][col_id][:50] + "..." if len(record_data['columns'][col_id]) > 50 else record_data['columns'][col_id]
                        st.caption(f"**{col['name']}:** {preview}")
                
                st.markdown("**Select column to load:**")
                load_column = st.radio("Column", [col['id'] for col in columns_config], 
                                      format_func=lambda x: next((col['name'] for col in columns_config if col['id'] == x), x),
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
                            col_name = next((col['name'] for col in columns_config if col['id'] == load_column), load_column)
                            st.warning(f"{col_name} is empty")
                
                with col_load2:
                    if st.button("Load to Text B", key="load_b"):
                        text_content = get_text_record(selected_record, load_column)
                        if text_content is not None:
                            st.session_state["tb"] = text_content
                            st.success(f"Loaded to Text B")
                            st.rerun()
                        else:
                            col_name = next((col['name'] for col in columns_config if col['id'] == load_column), load_column)
                            st.warning(f"{col_name} is empty")
                
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

# Get current column configuration
columns_config = get_columns_config()

col1, col2 = st.columns(2)

with col1:
    st.markdown("**Text A**")
    text_a = st.text_area("", height=200, key="ta", label_visibility="collapsed", placeholder="Enter first text here...")
    
    # Save Text A
    with st.expander("💾 Save Text A"):
        save_name_a = st.text_input("Enter a name for this text:", key="save_name_a")
        save_column_a = st.selectbox("Select column:", [col['id'] for col in columns_config],
                                     format_func=lambda x: next((col['name'] for col in columns_config if col['id'] == x), x),
                                     key="save_column_a")
        if st.button("Save Text A", key="save_btn_a"):
            if save_name_a and text_a:
                if save_text_record(save_name_a, text_a, save_column_a):
                    col_name = next((col['name'] for col in columns_config if col['id'] == save_column_a), save_column_a)
                    st.success(f"Saved as '{save_name_a}' in {col_name}")
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
        save_column_b = st.selectbox("Select column:", [col['id'] for col in columns_config],
                                     format_func=lambda x: next((col['name'] for col in columns_config if col['id'] == x), x),
                                     key="save_column_b")
        if st.button("Save Text B", key="save_btn_b"):
            if save_name_b and text_b:
                if save_text_record(save_name_b, text_b, save_column_b):
                    col_name = next((col['name'] for col in columns_config if col['id'] == save_column_b), save_column_b)
                    st.success(f"Saved as '{save_name_b}' in {col_name}")
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
