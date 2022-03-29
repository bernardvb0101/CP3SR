# import os
import plotly.express as px
from docx import Document
from docx.shared import Cm
# from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from datetime import datetime


def create_worddoc(var_dict, baseline_dict, df_project_cat, df_intersects2):
    """
    This module creates a spatial report in a MSWord File
    """
    global teller
    def par_formatter(paragraphs):
        """
        This function within a function formats paragraphs using normal, bold or italic.
        """
        paragraph = document.add_paragraph()
        for key, value in paragraphs.items():
            if " n" in key:  # Previous formatting
                paragraph.add_run(value)
            elif "al" in key: # Normal
                document.add_paragraph(value, style="Normal")
            elif " b" in key:  # Bold
                paragraph.add_run(value).bold = True
            elif " i" in key:  # Italic
                paragraph.add_run(value).italic = True
            elif " u" in key:  # Underline
                paragraph.add_run(value).underline = True
            elif "em" in key:  # Style: Emphasis
                paragraph.add_run(value).style = 'Emphasis'
            elif "tl" in key:  # Style: Title
                document.add_paragraph(value, style="Title")
            elif "st" in key:  # Style: SubTitle
                document.add_paragraph(value, style="Subtitle")
            elif "lb" in key:  # Style: List Bullet
                document.add_paragraph(value, style="List Bullet")
            elif "cp" in key:  # Style: Caption
                document.add_paragraph(value, style="Caption")


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


    # Unpack the variables that were passed to the function in a dictionary called var_dict
    sys_username = var_dict['username']
    url_choice = var_dict['url_choice']
    entity_choice = var_dict['entity_choice']
    SpatialFeatureChoice = var_dict['SpatialFeatureChoice']
    Layer_List = var_dict['Layer_List']
    total_datapoints = var_dict['total_datapoints']
    intersecting = var_dict['intersecting']
    no_intersects = var_dict['no_intersects']
    perc_no_intersects = (no_intersects / total_datapoints * 100)
    chosen_feature_qty = var_dict['chosen_feature_qty']

    # Set up other vars
    timenow = datetime.today().strftime('%Y-%m-%d-%H_%M')  # Stamp the date and time
    datenow = datetime.today().strftime('%Y-%m-%d')  # Stamp the date
    urlname = url_choice.split(sep=".", maxsplit=10)[1].capitalize()
    word_file_name = f"{urlname}_SpatialRep_{timenow}.docx"
    document = Document()

    #headings0 = {}
    #headings0['0'] = f"Spatial Report:\n{SpatialFeatureChoice}."
    paragraphs00 = {}
    paragraphs00['1 tl'] = "Spatial Query Report"
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
    paragraphs1['6 n'] = "The spatial feature that was selected for the purpose of this report was:\n"
    paragraphs1['7 b'] = f"'{SpatialFeatureChoice}'"
    paragraphs1['8 n'] = "."

    paragraphs2 = {}
    paragraphs2['1 n'] = f"This is a spatial feature report for the "
    paragraphs2['2 i'] = f"{entity_choice}."
    paragraphs2['3 n'] = f" The data contained in this report was sourced directly from the municipality's live CP3 " \
                         f"system. The information in this report therefore reflects the data as it was on the "
    paragraphs2['4 i'] = f"{datenow}"
    paragraphs2['5 n'] = f" when this report was requested. Any subsequent updates to the data contained the CP3 System" \
                         f" related to the applicable baseline from which this reports was drawn, would therefore not " \
                         f"reflect in this report. \n\nFor the "
    paragraphs2['6 i'] = f"{entity_choice}"
    paragraphs2['7 i'] = f", the following spatial features are available for the purpose of developing "
    paragraphs2['8 u'] = f"'Spatial Query Reports'"
    paragraphs2['9 n'] = f" (similar to this report):"
    teller = 10
    for feature in Layer_List:
        paragraphs2[f"{teller} lb"] = feature
        teller += 1

    paragraphs3 = {}
    paragraphs3['1 n'] = "We hope you find this useful! Sincerely, "
    paragraphs3['1 i'] = "The Novus3 Team."

    # Document Main Heading
    #head_formatter(headings0)
    par_formatter(paragraphs00)
    # Pass the dictionaries to the interpreter
    par_formatter(paragraphs0)
    par_formatter(paragraphs1)
    par_formatter(paragraphs2)
    document.add_picture(f"./static/images/CP3logo.png", width=Cm(3))
    par_formatter(paragraphs3)

    # Add a page break
    document.add_page_break()

    # heading_1 = "1. Introduction"
    # heading1_1 = "a. Addresses, Contact Details & Website"

    headings1 = {}
    headings1['1'] = "1. Introduction"

    # Put together the dictionary
    paragraphs4 = {}

    paragraphs4['1 n'] = f"The baseline that was queried for this report contains "
    paragraphs4['2 i'] = f"{total_datapoints}"
    paragraphs4['3 n'] = f" projects of which a total of "
    paragraphs4['4 i'] = f"{no_intersects} ({'{0:.3g}'.format(perc_no_intersects)}%)"
    paragraphs4['5 n'] = f" do not have any recorded intersection (overlap) with any of the following spatial features:"
    teller = 6
    for feature in Layer_List:
        paragraphs4[f"{teller} lb"] = f"{feature}"
        teller += 1
    paragraphs4[f"{teller} al"] = f"The probable reasons for the {no_intersects} non intersecting projects reported " \
                                  f"are that:"
    teller += 1
    paragraphs4[f"{teller} lb"] = f"these projects simply do not have a location associated with them yet (most " \
                                  f"frequently the case with 'no-intersect' projects) or;"
    teller += 1
    paragraphs4[f"{teller} lb"] = f"the location that was captured for some of these projects, may have been in " \
                                  f"the wrong place (e.g. outside the boundaries of {urlname} - this is very " \
                                  f"rarely the reason)."
    teller += 1
    paragraphs4[f"{teller} al"] = f"It is important to take note of the {'{0:.3g}'.format(perc_no_intersects)}%" \
                                  f" projects that do not intersect with any spatial feature when " \
                                  f"appraising this report because a similar proportion of non-intersecting " \
                                  f"projects may be present within the specific geographic feature queried."
    if perc_no_intersects <= 5:
        teller += 1
        paragraphs4[f"{teller} al"] = f"The {'{0:.3g}'.format(perc_no_intersects)}% of non-intersecting projects is a " \
                                      f"relative low percentage and therefore not much cause of concern for this " \
                                      f"particular report."

    teller += 1
    paragraphs4[f"{teller} cp"] = "Figure 1: Intersecting vs Non-Intersecting Projects"

    head_formatter(headings1)
    par_formatter(paragraphs4)
    # fig = px.pie(df_intersects2, values='Projects', names='Intersects', title='Intersecting vs Non Intersecting Projects')
    # fig = px.pie(df_intersects2, values='Projects', names='Intersects', color_discrete_sequence=px.colors.sequential.BuGn)
    colors = ['red', 'green']
    fig = px.pie(df_intersects2, values='Projects', names='Intersects')
    fig.update_traces(marker=dict(colors=colors, line=dict(color='#000000', width=2)))
    fig.write_image("./static/images/fig1.png")
    document.add_picture(f"./static/images/fig1.png", width=Cm(15))

    paragraphs5 = {}
    paragraphs5["1 al"] = "The above picture..."

    par_formatter(paragraphs5)



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

