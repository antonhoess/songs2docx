# pip install python-docx

from typing import Optional
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, RGBColor, Cm, Inches
import configparser


str_text = \
"""1. Jesus lebt! Er ist der Retter, durch sein Kreuz schenkt er Erlösung.  Dm F 
Gottes Sohn, auf ihn sind wir getauft. Er ist Sieger.   Gm Bb
Wir warn tot in unsern Sünden, durch sein Blut ist uns vergeben.    Dm F
Ja, er lebt! Wir glauben: Er ist der Herr!  Gm C4 C
"""


class Txt2Docx:
    def __init__(self, filename: str) -> None:
        self._filename = filename

        # Document
        self._document = Document()

        # Styles
        self._styles = None
        self._define_styles()

        # Read the file
        self._title = "???"
        self._authors = "???"
        self._copyright = "???"

        # Create the document
        self._read_file()
        self._build_document()
    # end def

    def _read_file(self):
        config = configparser.ConfigParser()
        config.read(self._filename, "utf-8")

        default_section = config[config.default_section]
        self._title = default_section['TITLE']
        self._authors = default_section['AUTHORS']
        self._copyright = default_section['COPYRIGHT']
    # end def

    def _build_document(self):
        # Create the document
        #####################

        # Title
        p = self._document.add_paragraph(self._title, style="title")
        self._set_paragraph_format(p)

        # Authors
        p = self._document.add_paragraph(self._authors, style="authors")
        self._set_paragraph_format(p)
        p = self._document.add_paragraph("", style="authors_after")
        self._set_paragraph_format(p)

        ## Text
        self._read_file()  # XXX

        p = self._document.add_paragraph(str_text, style="text_normal")
        self._set_paragraph_format(p)

        # Copyright
        p = self._document.add_paragraph("", style="copyright")
        self._set_paragraph_format(p)
        p = self._document.add_paragraph(self._copyright, style="copyright")
        self._set_paragraph_format(p)

        # XXX Test
        p = self._document.add_paragraph("")
        p.add_run('And this is text ')
        p.add_run('some bold text ').bold = True
        p.add_run('and italic text.')
    # end def

    def save(self, filename: Optional[str] = None) -> None:
        # Determine filename
        if not filename:
            filename = self._filename

            if filename.endswith(".txt"):
                filename = filename[:-4]

            filename += ".docx"
        # end if

        # Set page settings
        self._set_page_settings()

        # Save the document
        self._document.save(filename)
    # end def

    def _set_page_settings(self):
        # Changing page settings
        sections = self._document.sections

        for section in sections:
            # Page size
            section.page_width = Inches(8.5)
            section.page_height = Inches(11.)

            # Margin
            margin = Cm(2.5)
            section.top_margin = margin
            section.bottom_margin = margin
            section.left_margin = margin
            section.right_margin = margin
        # end for
    # end def

    @staticmethod
    def _set_paragraph_format(paragraph) -> None:
        paragraph_format = paragraph.paragraph_format
        paragraph_format.space_before = Pt(0)
        paragraph_format.space_after = Pt(0)
        paragraph_format.line_spacing = 1
    # end def

    def _define_styles(self):
        # Define styles
        ###############
        self._styles = self._document.styles

        # Title
        style = self._styles.add_style("title", WD_STYLE_TYPE.PARAGRAPH)
        style.font.name = 'Arial'
        style.font.size = Pt(12)
        style.font.bold = True
        style.font.color.rgb = RGBColor(51, 51, 51)

        # Authors
        style = self._styles.add_style("authors", WD_STYLE_TYPE.PARAGRAPH)
        style.font.name = 'Arial'
        style.font.size = Pt(8)
        style = self._styles.add_style("authors_after", WD_STYLE_TYPE.PARAGRAPH)
        style.font.name = 'Arial'
        style.font.size = Pt(6)

        # Text normal
        style = self._styles.add_style("text_normal", WD_STYLE_TYPE.PARAGRAPH)
        style.font.name = 'Arial'
        style.font.size = Pt(11)

        # Text bold
        style = self._styles.add_style("text_bold", WD_STYLE_TYPE.PARAGRAPH)
        style.font.name = 'Arial'
        style.font.size = Pt(11)
        style.font.bold = True

        # Copyright
        style = self._styles.add_style("copyright", WD_STYLE_TYPE.PARAGRAPH)
        style.font.name = 'Arial'
        style.font.size = Pt(8)
    # end def
# end class


def main():
    doc = Txt2Docx(filename=r"Dir, Herr Jesus Christ, Ehre, Macht und Herrlichkeit_AP_201013.txt")
    doc.save()
# end def


if __name__ == "__main__":
    main()
# end if
