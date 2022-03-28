from docx import Document
from docx.shared import Cm
from datetime import datetime


def create_worddoc(var_dict, baseline_dict, project_cat):
    """
    This module creates a spatial report in a MSWord File
    """
    def par_formatter(paragraphs):
        """
        This function within a function formats paragraphs using normal, bold or italic.
        """
        paragraph = document.add_paragraph()
        for key, value in paragraphs.items():
            if " n" in key:  # Normal formatting
                paragraph.add_run(value)
            elif " b" in key:  # Bold
                paragraph.add_run(value).bold = True
            elif " i" in key:  # Italic
                paragraph.add_run(value).italic = True
            elif "em" in key:  # Style: Emphasis
                paragraph.add_run(value).style = 'Emphasis'
            elif "tl" in key:  # Style: Title
                document.add_paragraph(value, style="Title")
            elif "st" in key:  # Style: SubTitle
                document.add_paragraph(value, style="Subtitle")
            elif "lb" in key:  # Style: List Bullet
                document.add_paragraph(value, style="List Bullet")


    def head_formatter(headings):
        """
         This function within a function formats headings using the heading levels as input.
        """
        for key, value in headings.items():
            match key:
                case "0":
                    document.add_heading(value, level=0)
                case "1":
                    document.add_heading(value, level=1)
                case "2":
                    document.add_heading(value, level=2)
                case "3":
                    document.add_heading(value, level=3)
                case "4":
                    document.add_heading(value, level=4)


    # Unpack the variables that were passed to the function in a dictionary called var_dict, set up other vars
    sys_username = var_dict['username']
    url_choice = var_dict['url_choice']
    entity_choice = var_dict['entity_choice']
    SpatialFeatureChoice = var_dict['SpatialFeatureChoice']
    timenow = datetime.today().strftime('%Y-%m-%d-%H_%M')  # Stamp the date and time
    urlname = url_choice.split(sep=".", maxsplit=10)[1].capitalize()
    word_file_name = f"{urlname}_SpatialRep_{timenow}.docx"
    document = Document()

    #headings0 = {}
    #headings0['0'] = f"Spatial Report:\n{SpatialFeatureChoice}."
    paragraphs00 = {}
    paragraphs00['1 tl'] = "Spatial Report"
    paragraphs00['2 st'] = f"{SpatialFeatureChoice}"

    # Put together the dictionary
    paragraphs0 = {}
    paragraphs0['1 n'] = "This document was created "
    paragraphs0['2 i'] = f"{timenow} "
    paragraphs0['3 n'] = "from the "
    paragraphs0['4 b'] = f"{entity_choice} "
    paragraphs0['5 n'] = "CP3 system by the following user: "
    paragraphs0['6 i'] = f"'{sys_username}'. "

    paragraphs1 = {}
    paragraphs1['1 n'] = "The baseline "
    paragraphs1['2 b'] = f"'{baseline_dict['Name']}'"
    paragraphs1['3 n'] = " with description: "
    paragraphs1['4 b'] = f"'{baseline_dict['Description']}'"
    paragraphs1['5 n'] = " was used. "
    paragraphs1['6 n'] = "The spatial feature that was selected for the purpose of this report was: "
    paragraphs1['7 b'] = f"'{SpatialFeatureChoice}'"
    paragraphs1['8 n'] = "."

    paragraphs2 = {}
    paragraphs2['1 n'] = "We hope you find this useful! Sincerely, "
    paragraphs2['1 i'] = "The Novus3 Team."

    # Document Main Heading
    #head_formatter(headings0)
    par_formatter(paragraphs00)
    # Pass the dictionaries to the interpreter
    par_formatter(paragraphs0)
    par_formatter(paragraphs1)
    document.add_picture(f"./static/images/CP3logo.png", width=Cm(3))
    par_formatter(paragraphs2)

    # Add a page break
    document.add_page_break()

    # heading_1 = "1. Introduction"
    # heading1_1 = "a. Addresses, Contact Details & Website"

    headings1 = {}
    headings1['1'] = "1. Introduction"
    headings1['2'] = "a. Addresses, Contact Details & Website"

    # Put together the dictionary
    paragraphs3 = {}
    paragraphs3['1 tl'] = "This is a title."
    paragraphs3['2 st'] = "This is a sub title."
    paragraphs3['3 lb'] = "This is a list bullet."
    #paragraphs0['4 b'] = f"{entity_choice} "
    #paragraphs0['5 n'] = "CP3 system by the following user: "
    #paragraphs0['6 i'] = f"'{sys_username}'. "

    head_formatter(headings1)
    par_formatter(paragraphs3)



    # Add a Page with General Information regarding the Municipality
    # document.add_heading(heading_1, level=1)
    # document.add_heading(heading1_1, level=2)

    #table = document.add_table(rows=1, cols=2, style='Light Grid Accent 1')
    #heading_cells = table.rows[0].cells
    #heading_cells[0].text = 'Description'
    #heading_cells[1].text = 'Info'


    full_path = f"./DOWNLOAD_FOLDER/{word_file_name}"
    document.save(full_path)

    return full_path

