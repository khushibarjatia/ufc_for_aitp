import streamlit as st
import os
from markitdown import MarkItDown
from io import BytesIO

# --- Configuration & Styling ---
st.set_page_config(page_title="Universal Doc Converter", page_icon="📄", layout="wide")
st.markdown("""
    <style>
    .stTextArea textarea { font-family: 'Courier New', Courier, monospace; font-size: 14px; }
    .main { background-color: #f8f9fa; }
    </style>
    """, unsafe_allow_html=True)

def format_size(size_bytes):
    """Helper to convert bytes to a human-readable string."""
    if size_bytes < 1024:
        return f"{size_bytes} Bytes"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.2f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.2f} MB"

def main():
    st.title("🚀 Universal File-to-Text Converter")
    st.info("Upload Word, Excel, PPT, PDF, HTML, or ZIP files to extract clean Markdown.")

    # [1] Initialize the Engine
    # MarkItDown handles the heavy lifting of parsing binary Office/PDF formats.
    md_engine = MarkItDown()

    # [2] Interface: Upload Area
    uploaded_files = st.file_uploader(
        "Drag and drop multiple files here", 
        type=["docx", "xlsx", "pptx", "pdf", "html", "zip"],
        accept_multiple_files=True
    )

    if uploaded_files:
        for uploaded_file in uploaded_files:
            base_filename = os.path.splitext(uploaded_file.name)[0]
            
            with st.expander(f"📄 Result: {uploaded_file.name}", expanded=True):
                # [3] Resilience: Error Handling
                try:
                    # MarkItDown requires a filename/extension to know which parser to use
                    # We pass the stream (BytesIO) and the filename together.
                    file_bytes = BytesIO(uploaded_file.getvalue())
                    
                    # Core Processing
                    result = md_engine.convert(uploaded_file.name, data=file_bytes)
                    converted_text = result.text_content

                    # Create Tabs for cleaner UX
                    tab1, tab2 = st.tabs(["📝 Preview & Download", "📊 File Size Comparison"])

                    with tab1:
                        # [2] Interface: Instant Preview
                        st.text_area(
                            label="Extracted Text (Markdown)",
                            value=converted_text,
                            height=350,
                            key=f"preview_{uploaded_file.name}"
                        )

                        # [2] Interface: Download Options
                        col1, col2 = st.columns(2)
                        new_filename_md = f"{base_filename}_converted.md"
                        new_filename_txt = f"{base_filename}_converted.txt"

                        with col1:
                            st.download_button(
                                label="📥 Download as .md",
                                data=converted_text,
                                file_name=new_filename_md,
                                mime="text/markdown",
                                key=f"dl_md_{uploaded_file.name}",
                                use_container_width=True
                            )
                        with col2:
                            st.download_button(
                                label="📄 Download as .txt",
                                data=converted_text,
                                file_name=new_filename_txt,
                                mime="text/plain",
                                key=f"dl_txt_{uploaded_file.name}",
                                use_container_width=True
                            )

                    with tab2:
                        # [4] File Size Comparison Logic
                        original_size = uploaded_file.size
                        # Measure converted text size in bytes using UTF-8 encoding
                        converted_size = len(converted_text.encode('utf-8'))
                        
                        # Calculate percentage reduction
                        if original_size > 0:
                            reduction = ((original_size - converted_size) / original_size) * 100
                        else:
                            reduction = 0

                        # Display Comparison Table
                        st.table({
                            "Version": ["Original File", "Converted Text"],
                            "Size": [format_size(original_size), format_size(converted_size)]
                        })

                        if reduction > 0:
                            st.success(f"✨ Success! The text version is **{reduction:.1f}% smaller** than the original.")
                        else:
                            st.warning("Note: The text version is larger or equal to the original (common for very small text-only files).")

                except Exception as e:
                    # [3] Resilience: Show polite error instead of crashing
                    st.error(f"⚠️ Could not read **{uploaded_file.name}**. Please check if the file is corrupted or password-protected.")
                    # Log the specific error for developer debugging in console
                    print(f"Error processing {uploaded_file.name}: {e}")

if __name__ == "__main__":
    main()
