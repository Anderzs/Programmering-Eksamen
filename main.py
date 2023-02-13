from json import load, dump
from enum import Enum
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH 
from datetime import datetime
import argparse


class Fag(Enum):
    FYSIK = "fysik.json",
    KEMI = "kemi.json",
    TEKNOLOGI = "teknologi.json"


class Journal:
    def __init__(self, _fag: str) -> None:
        match _fag.upper():
            case "FYSIK":
                self.fag = Fag.FYSIK
            case "KEMI":
                self.fag = Fag.KEMI
            case "TEKNOLOGI":
                self.fag = Fag.TEKNOLOGI
            case _:
                raise ValueError("Kunne ikke finde et eksisterende fag (fysik, kemi, teknologi)")
            
    def get_content(self, path: str) -> dict:
        with open(path, 'r') as f:
            return load(f)

    def create(self) -> bool:
        self.document = Document()
        self.data = self.get_content(path=self.fag.value[0])

        section = self.document.sections[0]
        header = section.header
        paragraph = header.paragraphs[0]
        paragraph.text = f"Anders Balleby Pedersen \
            \t{self.fag.name.capitalize()} \
            \t{datetime.today().strftime('%d/%m/%Y')} \
            \nUddannelsescenter Holstebro \
            \tTitel på dokument her \
            \t2.U HTX"
        
        paragraph.style = self.document.styles["Header"]

        if self.data['Front_page']['Create'] == True:
            heading = self.document.add_heading(f"{self.data['Front_page']['Title']}", 0)
            heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

            self.document.add_page_break()

        self.content = self.data['Content']
        for i in self.content:
            match i:
                case "Heading":
                    self.document.add_heading(f"{self.content[i]['Text']}", level=self.content[i]['Level'])
                case "Paragraph":
                    self.document.add_paragraph(f"{self.content[i]['Text']}")

        self.document.save('fysik_journal.docx')
        return True


        """
        document.add_heading('Document Title', 0)

        p = document.add_paragraph('A plain paragraph having some ')
        p.add_run('bold').bold = True
        p.add_run(' and some ')
        p.add_run('italic.').italic = True

        document.add_heading('Heading, level 1', level=1)
        document.add_paragraph('Intense quote', style='Intense Quote')

        document.add_paragraph(
            'first item in unordered list', style='List Bullet'
        )
        document.add_paragraph(
            'first item in ordered list', style='List Number'
        )

        document.add_picture('monty_truth.png', width=Inches(1.25))

        records = (
            (3, '101', 'Spam'),
            (7, '422', 'Eggs'),
            (4, '631', 'Spam, spam, eggs, and spam')
        )

        table = document.add_table(rows=1, cols=3)
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Qty'
        hdr_cells[1].text = 'Id'
        hdr_cells[2].text = 'Desc'
        for qty, id, desc in records:
            row_cells = table.add_row().cells
            row_cells[0].text = str(qty)
            row_cells[1].text = id
            row_cells[2].text = desc

        document.add_page_break()

        document.save('demo.docx')
        
        return True
        """
        return True

def load_parser() -> dict:
    parser = argparse.ArgumentParser(description="Program til oprettelse af journaler udfra skabeloner i .JSON filer",
                                    formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    #parser.add_argument("-a", "--archive", action="store_true", help="archive mode")
    #parser.add_argument("-v", "--verbose", action="store_true", help="increase verbosity")
    #parser.add_argument("-B", "--block-size", help="checksum blocksize")
    #parser.add_argument("--ignore-existing", action="store_true", help="skip files that exist")
    #parser.add_argument("--exclude", help="files to exclude")
    parser.add_argument("fag", help="faget der skal oprettes journal til")
    args = parser.parse_args()
    return vars(args)


if __name__ == "__main__":
    config = load_parser() # Indlæs argumenter fra CLI

    journal = Journal(config['fag'])
    if journal.create():
        print("Sucessfully created journal")
    else:
        print("Failed")
