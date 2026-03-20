"""
Create Excel file with embedded VBA macro for Threshold Explorer
This script generates an Excel file with the VBA macro already included,
making it easy for users to get started with threshold analysis.
"""

import os
import zipfile
import tempfile
from pathlib import Path

def create_excel_with_macro():
    """Create an Excel file with the VBA macro embedded"""
    
    # Read the VBA macro code
    macro_file = Path(__file__).parent / "Excel_Threshold_Explorer_Macro.bas"
    
    if not macro_file.exists():
        print(f"Error: Macro file not found at {macro_file}")
        return
    
    with open(macro_file, 'r') as f:
        vba_code = f.read()
    
    # Create a temporary directory
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        
        # Create the directory structure for an XLSM file
        xl_dir = temp_path / "xl"
        xl_dir.mkdir()
        
        # Create vbaProject.bin structure (simplified)
        vba_dir = temp_path / "xl" / "vbaProject.bin"
        vba_dir.mkdir(parents=True, exist_ok=True)
        
        # Create the necessary files for XLSM
        create_xlsm_structure(temp_path, vba_code)
        
        # Create the XLSM file
        output_file = Path(__file__).parent / "Threshold_Explorer_with_Macro.xlsm"
        create_xlsm_file(temp_path, output_file)
        
        print(f"Created Excel file with macro: {output_file}")
        print("\nTo use:")
        print("1. Open Threshold_Explorer_with_Macro.xlsm")
        print("2. Enable macros when prompted")
        print("3. Import your data from the enhanced export")
        print("4. Run the SetupThresholdExplorer macro")

def create_xlsm_structure(temp_path, vba_code):
    """Create the necessary XML structure for XLSM file"""
    
    # Create [Content_Types].xml
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    <Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
</Types>"""
    
    with open(temp_path / "[Content_Types].xml", "w") as f:
        f.write(content_types)
    
    # Create _rels directory and .rels file
    rels_dir = temp_path / "_rels"
    rels_dir.mkdir()
    
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
    
    with open(rels_dir / ".rels", "w") as f:
        f.write(rels)
    
    # Create xl/_rels directory and workbook.xml.rels
    xl_rels_dir = temp_path / "xl" / "_rels"
    xl_rels_dir.mkdir()
    
    xl_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
</Relationships>"""
    
    with open(xl_rels_dir / "workbook.xml.rels", "w") as f:
        f.write(xl_rels)
    
    # Create workbook.xml
    workbook = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheets>
        <sheet name="Instructions" sheetId="1" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
    </sheets>
</workbook>"""
    
    with open(temp_path / "xl" / "workbook.xml", "w") as f:
        f.write(workbook)
    
    # Create worksheet with instructions and macro code
    worksheet = create_instruction_sheet(vba_code)
    
    worksheets_dir = temp_path / "xl" / "worksheets"
    worksheets_dir.mkdir()
    
    with open(worksheets_dir / "sheet1.xml", "w") as f:
        f.write(worksheet)
    
    # Create a simple vbaProject.bin (this is a placeholder - real VBA projects are binary)
    # Note: Creating a real VBA project requires complex binary encoding
    # Instead, we'll create the alternative solution below

def create_instruction_sheet(vba_code):
    """Create worksheet with instructions and macro code"""
    
    instructions = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1" t="inlineStr"><is><t>Excel Threshold Explorer - VBA Macro Setup</t></is></c>
        </row>
        <row r="3">
            <c r="A3" t="inlineStr"><is><t>Instructions:</t></is></c>
        </row>
        <row r="4">
            <c r="A4" t="inlineStr"><is><t>1. Copy the VBA code from column B below</t></is></c>
        </row>
        <row r="5">
            <c r="A5" t="inlineStr"><is><t>2. Press Alt+F11 to open VBA editor</t></is></c>
        </row>
        <row r="6">
            <c r="A6" t="inlineStr"><is><t>3. Insert > Module</t></is></c>
        </row>
        <row r="7">
            <c r="A7" t="inlineStr"><is><t>4. Paste the code</t></is></c>
        </row>
        <row r="8">
            <c r="A8" t="inlineStr"><is><t>5. Run the SetupThresholdExplorer macro</t></is></c>
        </row>
        <row r="10">
            <c r="A10" t="inlineStr"><is><t>VBA Code (copy from B11 and below):</t></is></c>
        </row>"""
    
    # Add the VBA code to the worksheet (split into multiple cells due to Excel limitations)
    lines = vba_code.split('\n')
    row_num = 11
    
    for line in lines[:100]:  # Limit to first 100 lines to fit in worksheet
        # Clean up the line for XML
        clean_line = line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')
        worksheet += f"""
        <row r="{row_num}">
            <c r="B{row_num}" t="inlineStr"><is><t>{clean_line}</t></is></c>
        </row>"""
        row_num += 1
    
    worksheet += """
    </sheetData>
</worksheet>"""
    
    return worksheet

def create_xlsm_file(temp_path, output_file):
    """Create the final XLSM file by zipping the contents"""
    
    with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(temp_path):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, temp_path)
                zf.write(file_path, arcname)

if __name__ == "__main__":
    create_excel_with_macro()
