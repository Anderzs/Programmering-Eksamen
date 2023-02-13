from json import load, dump
from enum import Enum
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from datetime import datetime
from sys import platform
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
        with open(path, 'r', encoding='utf-8') as f:
            return load(f)

    def load_front_page(self) -> None:
        if self.data['Front_page']['Create']:
            self.heading = self.document.add_heading(f"{self.data['Front_page']['Title']}", 0)
            self.heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Forklar dette i dokumentationen
            # https://stackoverflow.com/questions/60921603/how-do-i-change-heading-font-face-and-size-in-python-docx
            title_style = self.heading.style
            rFonts = title_style.element.rPr.rFonts
            rFonts.set(qn("w:asciiTheme"), "Times New Roman")

            test = self.document.add_paragraph("")
            test.add_run("Anders Balleby\n2.U\nFysik A", style='Front_page').bold = True
            test.alignment = WD_ALIGN_PARAGRAPH.CENTER

            if (image := self.data['Front_page']['Image'])['Exist']:
                url = image['URL']
                self.document.add_picture(url, width=Inches(5.9), height=Inches(2.36))

            self.document.add_page_break()

    def load_styles(self) -> None:
        # Font indstillinger

        obj_styles = self.document.styles

        # Forside
        obj_charstyle = obj_styles.add_style('Front_page', WD_STYLE_TYPE.CHARACTER)
        obj_font = obj_charstyle.font
        obj_font.name = f'{self.data["General"]["Font"]}'
        obj_font.size = Pt(22)

        # Overskrifter
        obj_charstyle = obj_styles.add_style('Overskrift', WD_STYLE_TYPE.CHARACTER)
        obj_font = obj_charstyle.font
        obj_font.name = f'{self.data["General"]["Font"]}'
        obj_font.size = Pt(self.data["General"]["Headings"]["Size"])

        # Afsnit
        obj_charstyle = obj_styles.add_style('Afsnit', WD_STYLE_TYPE.CHARACTER)
        obj_font = obj_charstyle.font
        obj_font.name = f'{self.data["General"]["Font"]}'
        obj_font.size = Pt(self.data["General"]["Paragraphs"]["Size"])

    def load_header(self) -> None:
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

    def create(self) -> bool:
        self.document = Document()
        self.data = self.get_content(path=self.fag.value[0])

        # Styles
        self.load_styles()

        # Header
        self.load_header()
        
        

        # Forside-Tjek
        self.load_front_page()

        self.content = self.data['Content']
        for i in self.content:
            match self.content[i]['Type']:
                case "Heading":
                    tmp_heading = self.document.add_heading("", 1)
                    tmp_heading.add_run(f"{i}", style='Overskrift')
                case "Paragraph":
                    tmp_paragraph = self.document.add_paragraph("")
                    tmp_paragraph.add_run(f"{i}", style='Afsnit')

        self.document.save(f'output/{self.heading.text}.docx')
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
        print(f"Opening {(name := journal.heading.text)}")
        
        if platform == "win32":
            from os import startfile
            startfile(f"C:/Users/ander/Documents/Programmering-Eksamen/output/{name}.docx")
        elif platform == "darwin":
            #  TODO: implementer understøttelse til MacOS / Darwin
            raise Error("Darwin & MacOS not supported yet.")
            pass

    else:
        raise Error("Failed to save journal")
