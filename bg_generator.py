"""
Bulk BG Document Generator
Converts multiple BGs from a text file into individual Word documents
"""

import os
import re
from pathlib import Path
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH


def parse_bgs_from_file(file_path):
    """
    Parse BGs from a text file. Handles multiple formats:
    - Separated by dashes: "---" or "======"
    - Numbered clauses: "1. ", "2. ", etc.
    - Separated by blank lines
    """
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    bgs = []
    
    # First priority: Try dash/line separator format (---)
    if re.search(r'---+', content):
        bgs = re.split(r'---+', content)
        bgs = [bg.strip() for bg in bgs if bg.strip()]
        if len(bgs) > 1:
            print(f"✓ Detected '---' separator format: Found {len(bgs)} complete BGs")
            return bgs
    
    # Second priority: Extract numbered clauses (1. 2. 3. etc.) from content
    clause_match = re.findall(r'^\d+\.\s+(.+?)(?=^\d+\.|$)', content, re.MULTILINE | re.DOTALL)
    if clause_match and len(clause_match) > 1:
        bgs = [clause.strip() for clause in clause_match]
        print(f"✓ Detected numbered clauses format: Found {len(bgs)} clauses")
        return bgs
    
    # Fallback: treat multiple blank lines as separators
    bgs = re.split(r'\n\s*\n+', content)
    bgs = [bg.strip() for bg in bgs if bg.strip()]
    
    if len(bgs) > 1:
        print(f"✓ Detected blank line format: Found {len(bgs)} BGs")
    else:
        print("⚠ Warning: Could not detect multiple BGs. Make sure they are properly separated.")
    
    return bgs


def create_word_document(bg_text, output_file):
    """
    Create a Word document with the BG text
    """
    doc = Document()
    
    # Add title
    title = doc.add_heading('Background', level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Add content
    paragraph = doc.add_paragraph(bg_text)
    paragraph_format = paragraph.paragraph_format
    paragraph_format.line_spacing = 1.5
    paragraph_format.space_after = Pt(12)
    
    # Set font
    for run in paragraph.runs:
        run.font.size = Pt(12)
        run.font.name = 'Calibri'
    
    # Save document
    doc.save(output_file)


def generate_bg_documents(input_file, output_dir='BG_Documents'):
    """
    Main function to generate Word documents from BGs text file
    """
    # Create output directory
    Path(output_dir).mkdir(exist_ok=True)
    
    # Parse BGs
    bgs = parse_bgs_from_file(input_file)
    
    if not bgs:
        print("❌ No BGs found in the file!")
        return
    
    print(f"\n📄 Generating {len(bgs)} Word documents...\n")
    
    # Generate documents
    for idx, bg in enumerate(bgs, 1):
        filename = f"{output_dir}/BG_{idx:03d}.docx"
        try:
            create_word_document(bg, filename)
            print(f"✓ Created: {filename}")
        except Exception as e:
            print(f"❌ Error creating {filename}: {str(e)}")
    
    print(f"\n✅ Done! Generated {len(bgs)} documents in '{output_dir}' folder")


if __name__ == "__main__":
    # Configuration
    INPUT_FILE = "bgs_input.txt"  # Your text file with all BGs
    OUTPUT_DIR = "BG_Documents"
    
    if not os.path.exists(INPUT_FILE):
        print(f"❌ Error: '{INPUT_FILE}' not found!")
        print("Please create a text file with your BGs and name it 'bgs_input.txt'")
    else:
        generate_bg_documents(INPUT_FILE, OUTPUT_DIR)
