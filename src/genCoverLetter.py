from docx import Document
import sys
import os

doc = Document(r"outlines\coverLetterOutlineGeneric.docx")

def genLetter(company,title)->None:
    # Replace placeholders in all paragraphs
    for para in doc.paragraphs:
        if "[COMPANY]" in para.text or "[POSITION]" in para.text:
            inline = para.runs
            for i in range(len(inline)):
                text = inline[i].text
                text = text.replace("[COMPANY]", company)
                text = text.replace("[POSITION]", title)
                inline[i].text = text
    

     # Create output directory
    output_dir = os.path.join("coverLetters", company)
    os.makedirs(output_dir, exist_ok=True)

    # Save DOCX
    docx_path = os.path.join(output_dir, f"{company}_cover_letter.docx")
    doc.save(docx_path)
    print(f"Saved DOCX: {docx_path}")

    # Save PDF (using docx2pdf — requires Microsoft Word on Windows or LibreOffice on Mac)
    try:
        from docx2pdf import convert
        pdf_path = os.path.join(output_dir, f"{company}_cover_letter.pdf")
        convert(docx_path, pdf_path)
        print(f"Saved PDF: {pdf_path}")
    except ImportError:
        print("⚠️ PDF export requires 'docx2pdf'. Install with: pip install docx2pdf")
    except Exception as e:
        print(f"⚠️ Could not generate PDF: {e}")




    
if __name__ == "__main__":
    if(len(sys.argv) != 3):
        print(f"Usage: coverLetterGen <company> <title>")
    else:
        genLetter(sys.argv[1],sys.argv[2])