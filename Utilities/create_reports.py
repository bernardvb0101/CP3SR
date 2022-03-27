from docx import Document
from docx.shared import Inches
from datetime import datetime


def create_worddoc(username, full_url, baseline_dict, project_cat):
    """
    This module creates a spatial report in a MSWord File
    """

    document = Document()
    timenow = datetime.today().strftime('%Y-%m-%d-%H_%M_%S')  # Stamp the date and time


    urlname = full_url.split(sep=".", maxsplit=10)[1].capitalize()
    word_file_name = f"{urlname}_{username}_{timenow}.docx"

    doc_heading = f"Spatial Report from {full_url} CP3 System."
    paragraph0_0 = f"This document was created {timenow} from the {urlname} CP3 system. The baseline" \
                   f" {baseline_dict['Name']} with description {baseline_dict['Description']} was used."

    paragraph0_1 = f"This document contains information pulled from the following CP3 site: {full_url}. " \
                 f"CP3 system, based on the choices you have made in the CP3C web-application. Section 0 of this " \
                 f"document provides you with some background information about {municname}. You will find the address" \
                 f" and contact details that you would require to draft a letter as well as context to appraise the " \
                   f"quotation you are about issue to this client. All of the content that you might need for the " \
                   f"drafting of your quotation document is contained in the subsequent numbered sections of this " \
                   f"document, in the order that it is most likely to be presented and used."
    paragraph0_2 = "We hope you find this useful! Sincerely, The Novus3 Team."

    heading_1 = "1. Introduction"
    heading0_1 = "a. Addresses, Contact Details & Website"

    # Document Main Heading
    document.add_heading(doc_heading, 0)
    document.add_paragraph(paragraph0_0)
    document.add_paragraph(paragraph0_1)

    document.add_picture(f"./static/images/CP3logo.png", width=Inches(1.25))
    document.add_paragraph(paragraph0_2)

    # Add a page break
    document.add_page_break()

    # Add a Page with General Information regarding the Municipality
    document.add_heading(heading_1, level=1)

    document.add_heading(heading0_1, level=2)

    table = document.add_table(rows=1, cols=2, style='Light Grid Accent 1')
    heading_cells = table.rows[0].cells
    heading_cells[0].text = 'Description'
    heading_cells[1].text = 'Info'


    full_path = f"./DOWNLOAD_FOLDER/{word_file_name}"
    document.save(full_path)

    return full_path

