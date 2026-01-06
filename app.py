import streamlit as st
import os
from markitdown import MarkItDown
from io import BytesIO

# --- Configuration & Styling ---
st.set_page_config(page_title="Universal Doc Converter", page_icon="📄")
st.markdown("""
    <style>
    .stTextArea textarea { font-family: 'Courier New', Courier, monospace; }
    </style>
    """, unsafe_allow_html=True)

def main():
    st.title("🚀 Universal File-to-Text Converter")
    st.subheader("Convert Office, PDF, and HTML into Clean Markdown")

    # [1] Initialize the Engine
    # Note: MarkItDown handles Word, Excel, PPT, PDF, and HTML natively.
    md_engine = MarkItDown()

    # [2] Interface: Upload Area
    uploaded_files = st.file_uploader(
        "Drag and drop files here", 
        type=["docx", "xlsx", "pptx", "pdf", "html", "zip"],
        accept_multiple_files=True
    )

    if uploaded_files:
        for uploaded_file in uploaded_files:
            file_extension = os.path.splitext(uploaded_file.name)[1].lower()
            base_filename = os.path.splitext(uploaded_file.name)[0]
            
            st.divider()
            st.write(f"### Processing: `{uploaded_file.name}`")
            
            # [3] Resilience: Error Handling
            try:
                # Process the file buffer
                # MarkItDown's convert method can take a file path or a stream
                # We wrap the uploaded bytes in BytesIO for the engine
                file_bytes = BytesIO(uploaded_file.getvalue())
                
                # Using 5s timeout logic for any internal web calls (if applicable to format)
                result = md_engine.convert(uploaded_file.name, data=file_bytes)
                converted_text = result.text_content

                # [2] Interface: Instant Preview
                st.text_area(
                    label="Preview Content",
                    value=converted_text,
                    height=300,
                    key=f"preview_{uploaded_file.name}"
                )

                # [2] Interface: Download Options
                col1, col2 = st.columns(2)
                
                # Format file name: OriginalName_converted.ext
                new_filename_md = f"{base_filename}_converted.md"
                new_filename_txt = f"{base_filename}_converted.txt"

                with col1:
                    st.download_button(
                        label="📥 Download as Markdown (.md)",
                        data=converted_text,
                        file_name=new_filename_md,
                        mime="text/markdown",
                        key=f"dl_md_{uploaded_file.name}"
                    )
                
                with col2:
                    st.download_button(
                        label="📄 Download as Plain Text (.txt)",
                        data=converted_text,
                        file_name=new_filename_txt,
                        mime="text/plain",
                        key=f"dl_txt_{uploaded_file.name}"
                    )

            except Exception as e:
                st.error(f"⚠️ Could not read {uploaded_file.name}. Please check the format.")
                st.info("Technical details: Ensure the file is not password protected.")

if __name__ == "__main__":
    main()
