# import os
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from docx import Document
from docx.shared import Cm
# from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from datetime import datetime


def create_worddoc(var_dict, baseline_dict, df_project_cat, df_intersects2, df_subset1, df_subset2, df_subset3,
                   df_subset4, number_of_plots, df_EntireSet):
    """
    This module creates a spatial report in a MSWord File
    """
    global teller, fig_nr, fig, tbl_nr, tbl
    fig_nr = 1
    fig = {}
    tbl_nr = 1
    tbl ={}

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
                try:
                    document.add_paragraph(value, style="Subtitle")
                except KeyError:
                    document.add_paragraph(value, style="Strong")
            elif 'se' in key:
                document.add_paragraph(value, style="Section")
            elif "lb" in key:  # Style: List Bullet
                try:
                    document.add_paragraph(value, style="List Bullet")
                except KeyError:
                    document.add_paragraph(value, style="Bulleted list")
            elif "xx" in key:
                document.add_paragraph(value, style="Bulleted list")
            elif "cp" in key:  # Style: Caption
                document.add_paragraph(value, style="Caption")


    def head_formatter(headings):
        """
         This function within a function formats headings using the heading levels as input.
        """
        for key, value in headings.items():
            if key == "0":
                    document.add_heading(value, level=0)
            elif key == "1":
                    document.add_heading(value, level=1)
            elif key == "2":
                    document.add_heading(value, level=2)
            elif key == "3":
                    document.add_heading(value, level=3)
            elif key == "4":
                    document.add_heading(value, level=4)
            """
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
            """


    # Put an R in front of the currency and format the amount
    def format_budget(n):
        """
        This function formats a number entry into a currency by putting an 'R' in front and delimiting with commas
        """
        currency = "R{:0,.0f}".format(n).replace('R-', '-R')
        # ${:0,.0f}
        return currency

    # Put an % in and * 100
    def format_percent(n):
        """
        This function formats a number entry into a percentage by putting an '%' in the back and delimiting with 2
        decimals
        """
        n = float(n)
        n = n * 100
        percentage = "{:.2f}%".format(n)
        return percentage


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

    df_EntireSet_Temp = df_EntireSet

    # Create variables about the projects
    maximum_projects = df_EntireSet[f'Projects per {SpatialFeatureChoice}'].max()
    maximum_projects_feature = df_EntireSet.loc[df_EntireSet[f'Projects per {SpatialFeatureChoice}'] == df_EntireSet[
        f'Projects per {SpatialFeatureChoice}'].max(), f'{SpatialFeatureChoice}'].iloc[0]
    minimum_projects = df_EntireSet[f'Projects per {SpatialFeatureChoice}'].min()
    minimum_projects_feature = df_EntireSet.loc[df_EntireSet[f'Projects per {SpatialFeatureChoice}'] == df_EntireSet[
        f'Projects per {SpatialFeatureChoice}'].min(), f'{SpatialFeatureChoice}'].iloc[0]
    average_projects = df_EntireSet[f'Projects per {SpatialFeatureChoice}'].mean()
    seventy_fifth_projects = df_EntireSet[f'Projects per {SpatialFeatureChoice}'].quantile(q=0.75)
    sum_projects = df_EntireSet[f'Projects per {SpatialFeatureChoice}'].sum()
    min_projects_perc_of_total = format_percent(minimum_projects/sum_projects)
    max_projects_perc_of_total = format_percent(maximum_projects/sum_projects)
    # Re-sort the df to get the top 5 wards with the most projects
    df_EntireSet_Temp.sort_values(f'Projects per {SpatialFeatureChoice}', inplace=True, ascending=True)
    sum_top_five_projects = df_EntireSet_Temp[f'Projects per {SpatialFeatureChoice}'].tail().sum()
    sum_top_five_projects_perc_of_total = format_percent(sum_top_five_projects/sum_projects)
    top_five_projfeat_text = " and".join(", ".join(df_EntireSet_Temp[f'{SpatialFeatureChoice}'].tail().tolist()).rsplit(',',1))


    # Create variables about the costs
    maximum_cost_float = df_EntireSet['Capital Demand'].max()
    maximum_cost = format_budget(maximum_cost_float)
    maximum_cost_feature = \
    df_EntireSet.loc[df_EntireSet['Capital Demand'] == df_EntireSet['Capital Demand'].max(), f'{SpatialFeatureChoice}'].iloc[0]
    minimum_cost_float = df_EntireSet['Capital Demand'].min()
    minimum_cost = format_budget(minimum_cost_float)
    minimum_cost_feature = \
        df_EntireSet.loc[df_EntireSet['Capital Demand'] == df_EntireSet['Capital Demand'].min(), f'{SpatialFeatureChoice}'].iloc[
            0]
    average_cost_float = df_EntireSet['Capital Demand'].mean()
    average_cost = format_budget(average_cost_float)
    seventy_fifth_cost_float = df_EntireSet['Capital Demand'].quantile(q=0.75)
    seventy_fifth_cost = format_budget(seventy_fifth_cost_float)
    sum_cost_float = df_EntireSet['Capital Demand'].sum()
    sum_cost = format_budget(sum_cost_float)
    max_cost_perc_of_total_float = df_EntireSet['Capital Demand'].max()/df_EntireSet['Capital Demand'].sum()
    min_cost_perc_of_total_float = df_EntireSet['Capital Demand'].min()/df_EntireSet['Capital Demand'].sum()
    max_cost_perc_of_total = format_percent(max_cost_perc_of_total_float)
    min_cost_perc_of_total = format_percent(min_cost_perc_of_total_float)
    # Re-sort the df to get the top 5 capital demand
    df_EntireSet_Temp.sort_values('Capital Demand', inplace=True, ascending=True)
    sum_top_five_cost_float = df_EntireSet_Temp['Capital Demand'].tail().sum()
    sum_top_five_cost = format_budget(sum_top_five_cost_float)
    sum_top_five_cost_perc_of_total_float = sum_top_five_cost_float/sum_cost_float
    sum_top_five_cost_perc_of_total = format_percent(sum_top_five_cost_perc_of_total_float)
    top_five_capfeat_text = " and".join(", ".join(df_EntireSet_Temp[f'{SpatialFeatureChoice}'].tail().tolist()).rsplit(',',1))


    # This provides the location of the template word file
    document = Document(docx=f"./ReportFactory/Master.docx")

    paragraphs00 = {}
    paragraphs00['1 tl'] = f"Spatial Query Report: {SpatialFeatureChoice}"

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
        paragraphs2[f"{teller} xx"] = feature
        teller += 1

    paragraphs3 = {}
    paragraphs3['1 n'] = "We hope you find this useful! Sincerely, "
    paragraphs3['2 i'] = "The Novus3 Team."

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

    headings1 = {}
    headings1['1'] = "Introduction"

    # Put together the dictionary
    paragraphs4 = {}

    paragraphs4['1 n'] = f"The baseline that was queried for this report contains "
    paragraphs4['2 i'] = f"{total_datapoints}"
    paragraphs4['3 n'] = f" projects of which a total of "
    paragraphs4['4 i'] = f"{no_intersects} ({'{0:.3g}'.format(perc_no_intersects)}%)"
    paragraphs4['5 n'] = f" projects do not have any recorded intersection (overlap) with {SpatialFeatureChoice}." \
                         f"The following spatial features are available for queries using this API profile:"
    teller = 6
    for feature in Layer_List:
        paragraphs4[f"{teller} xx"] = f"{feature}"
        teller += 1
    paragraphs4[f"{teller} al"] = f"\nThe probable reasons for the {no_intersects} non intersecting projects reported" \
                                  f" are that:"
    teller += 1
    paragraphs4[f"{teller} xx"] = f"these projects simply do not have a location associated with them yet (most " \
                                  f"frequently the case with 'no-intersect' projects) or;"
    teller += 1
    paragraphs4[f"{teller} xx"] = f"the location that was captured for some of these projects, may have been in " \
                                  f"the wrong place (e.g. outside the boundaries of {urlname} - this is very " \
                                  f"rarely the reason)."
    teller += 1
    paragraphs4[f"{teller} al"] = f"\nIt is important to take note of the {'{0:.3g}'.format(perc_no_intersects)}%" \
                                  f" projects that do not intersect with {SpatialFeatureChoice} feature when " \
                                  f"appraising this report."
    if perc_no_intersects <= 5:
        teller += 1
        paragraphs4[f"{teller} al"] = f"The {'{0:.3g}'.format(perc_no_intersects)}% of non-intersecting projects is a " \
                                      f"relative low percentage and therefore not much cause of concern for this " \
                                      f"particular report."

    teller += 1
    paragraphs4[f"{teller} cp"] = f"Figure {fig_nr}: Intersecting vs Non-Intersecting Projects"

    head_formatter(headings1)
    par_formatter(paragraphs4)

    colors = ['red', 'green']
    fig[fig_nr] = px.pie(df_intersects2, values='Projects', names='Intersects')
    fig[fig_nr].update_traces(marker=dict(colors=colors, line=dict(color='#000000', width=2)))
    fig[fig_nr].write_image(f"./static/images/fig{fig_nr}.png")
    document.add_picture(f"./static/images/fig{fig_nr}.png", width=Cm(15))
    fig_nr += 1

    # Add a page break
    # document.add_page_break()

    headings2 = {}
    headings2['1'] = f"{SpatialFeatureChoice} Analysis"
    paragraphs5 = {}
    paragraphs5['1 cp'] = f"Table {tbl_nr}: {SpatialFeatureChoice} Analysis"

    head_formatter(headings2)
    par_formatter(paragraphs5)

    table = document.add_table(rows=1, cols=4, style='Grid Table 4')
    heading_cells = table.rows[0].cells
    heading_cells[0].text = 'Description'
    heading_cells[1].text = 'Value'
    heading_cells[2].text = 'Description'
    heading_cells[3].text = 'Value'

    cells = table.add_row().cells
    cells[0].text = f"The total number of projects in all {SpatialFeatureChoice}:"
    cells[1].text = f"{sum_projects} projects"
    cells[2].text = f"The total capital demand in all {SpatialFeatureChoice}:"
    cells[3].text = f"{sum_cost}"

    cells = table.add_row().cells
    cells[0].text = f"The highest number of projects:"
    cells[1].text = f"{SpatialFeatureChoice}: {maximum_projects_feature}\n\nNumber of projects: {maximum_projects}\n\nPercentage of total: " \
                    f"{max_projects_perc_of_total}"
    cells[2].text = f"The highest capital demand:"
    cells[3].text = f"{SpatialFeatureChoice}: {maximum_cost_feature}\n\nCapital Demand: {maximum_cost}\n\nPercentage of total: " \
                    f"{max_cost_perc_of_total}"

    cells = table.add_row().cells
    cells[0].text = f"The lowest number of projects:"
    cells[1].text = f"{SpatialFeatureChoice}: {minimum_projects_feature}\n\nNumber of projects: {minimum_projects}\n\nPercentage of total: " \
                    f"{min_projects_perc_of_total}"
    cells[2].text = f"The lowest capital demand:"
    cells[3].text = f"{SpatialFeatureChoice}: {minimum_cost_feature}\n\nCapital Demand: {minimum_cost}\n\nPercentage of total: " \
                    f"{min_cost_perc_of_total}"

    if chosen_feature_qty > 10:
        cells = table.add_row().cells
        cells[0].text = f"The 5 {SpatialFeatureChoice} that collectively have with the highest number of projects:"
        cells[1].text = f"{SpatialFeatureChoice}: {top_five_projfeat_text}\n\nNumber of projects: {sum_top_five_projects}\n\nPercentage of total: " \
                        f"{sum_top_five_projects_perc_of_total}"
        cells[2].text = f"The 5 {SpatialFeatureChoice} that collectively have collective the highest capital demand:"
        cells[3].text = f"{SpatialFeatureChoice}: {top_five_capfeat_text}\n\nCapital Demand: {sum_top_five_cost}\n\nPercentage of total: " \
                        f"{sum_top_five_cost_perc_of_total}"

    cells = table.add_row().cells
    cells[0].text = f"The average number of projects per {SpatialFeatureChoice}:"
    cells[1].text = f"{'{0:.3g}'.format(average_projects)} projects"
    cells[2].text = f"The average capital demand per {SpatialFeatureChoice}:"
    cells[3].text = f"{average_cost}"

    cells = table.add_row().cells
    cells[0].text = f"The 75th percentile of projects per {SpatialFeatureChoice}:"
    cells[1].text = f"{seventy_fifth_projects} projects"
    cells[2].text = f"The 75th percentile of capital demand per {SpatialFeatureChoice}:"
    cells[3].text = f"{seventy_fifth_cost}"

    tbl_nr += 1

    paragraphs5a = {}
    if chosen_feature_qty > 10:
        paragraphs5a['1 cp'] = f"Figure {fig_nr}.1: The single highest {SpatialFeatureChoice} (Nr of Projects" \
                           f" & Capital Demand) vs the rest"
    else:
        paragraphs5a['1 cp'] = f"Figure {fig_nr}: The single highest {SpatialFeatureChoice} (Nr of Projects" \
                           f" & Capital Demand) vs the rest"
    par_formatter(paragraphs5a)
    colors = ['red', 'green']
    labels = ["Highest", "The Rest"]
    # Create subplots: use 'domain' type for Pie subplot
    fig[fig_nr] = make_subplots(rows=1, cols=2, specs=[[{'type': 'domain'}, {'type': 'domain'}]])
    fig[fig_nr].add_trace(go.Pie(labels=labels, values=[maximum_projects, (sum_projects - maximum_projects)],
                                 name="Nr of Projs"), 1, 1)
    fig[fig_nr].add_trace(
        go.Pie(labels=labels, values=[maximum_cost_float, (sum_cost_float - maximum_cost_float)],
               name="Cap Dem (R)"), 1, 2)
    # Use `hole` to create a donut-like pie chart
    fig[fig_nr].update_traces(hole=.4)
    fig[fig_nr].update_traces(marker=dict(colors=colors, line=dict(color='#000000', width=2)))
    fig[fig_nr].update_layout(annotations=[dict(text='Nr of Projs', x=0.16, y=0.5, font_size=10, showarrow=False),
                                           dict(text='Cap Dem (R)', x=0.85, y=0.5, font_size=10, showarrow=False)])
    fig[fig_nr].update_layout(legend=dict(orientation="h", y=0.99, x=0.35))

    fig[fig_nr].write_image(f"./static/images/fig{fig_nr}.png")
    document.add_picture(f"./static/images/fig{fig_nr}.png", width=Cm(17))

    if chosen_feature_qty < 11:  # If >10, it will be incremented in the next step, not now
        fig_nr += 1


    if chosen_feature_qty > 10:
        paragraphs5b = {}
        paragraphs5b['1 cp'] = f"Figure {fig_nr}.2: The 5 {SpatialFeatureChoice} with the highest values (Nr of Projects" \
                               f" & Capital Demand) vs the rest"
        par_formatter(paragraphs5b)

    if chosen_feature_qty > 10:
        colors = ['red', 'green']
        labels = ["5 Highest", "The Rest"]
        # Create subplots: use 'domain' type for Pie subplot
        fig[fig_nr] = make_subplots(rows=1, cols=2, specs=[[{'type':'domain'}, {'type':'domain'}]])
        fig[fig_nr].add_trace(go.Pie(labels=labels, values=[sum_top_five_projects, (sum_projects - sum_top_five_projects)],
                                     name="Nr of Projs"), 1, 1)
        fig[fig_nr].add_trace(
            go.Pie(labels=labels, values=[sum_top_five_cost_float, (sum_cost_float - sum_top_five_cost_float)],
                   name="Cap Dem (R)"), 1, 2)
        # Use `hole` to create a donut-like pie chart
        fig[fig_nr].update_traces(hole=.4)
        fig[fig_nr].update_traces(marker=dict(colors=colors, line=dict(color='#000000', width=2)))
        fig[fig_nr].update_layout(annotations=[dict(text='Nr of Projs', x=0.16, y=0.5, font_size=10, showarrow=False),
                                               dict(text='Cap Dem (R)', x=0.85, y=0.5, font_size=10, showarrow=False)])
        fig[fig_nr].update_layout(legend=dict(orientation="h", y=0.99, x=0.35))

        fig[fig_nr].write_image(f"./static/images/fig{fig_nr}.png")
        document.add_picture(f"./static/images/fig{fig_nr}.png", width=Cm(17))
        fig_nr += 1


    paragraphs6 ={}
    paragraphs7 = {}
    paragraphs8 = {}
    paragraphs9 = {}


    if number_of_plots == 1:
                fig[fig_nr] = px.bar(df_subset1, x=f'Projects per {SpatialFeatureChoice}', y= SpatialFeatureChoice,
                            color='Capital Demand',
                            height=900, orientation='h')
                # Sort images to follow each other
                fig[fig_nr].update_yaxes(categoryorder='total descending')
                # Write images to png
                fig[fig_nr].write_image(f"./static/images/fig{fig_nr}.png")
                # Add a page break
                document.add_page_break()
                paragraphs6['1 cp'] = f"Figure {fig_nr}: Projects and Capital Demand per {SpatialFeatureChoice}"
                par_formatter(paragraphs6)
                document.add_picture(f"./static/images/fig{fig_nr}.png", width=Cm(15))
    elif number_of_plots == 2:
                temp_fig_nr1 = f"{fig_nr}.1"
                temp_fig_nr2 = f"{fig_nr}.2"
                fig[temp_fig_nr1] = px.bar(df_subset1, x=f'Projects per {SpatialFeatureChoice}', y=SpatialFeatureChoice,
                            color='Capital Demand',
                            height=900, orientation='h')
                fig[temp_fig_nr2] = px.bar(df_subset2, x=f'Projects per {SpatialFeatureChoice}', y=SpatialFeatureChoice,
                                color='Capital Demand',
                                height=900, orientation='h')
                # Sort images to follow each other
                fig[temp_fig_nr1].update_yaxes(categoryorder='total descending')
                fig[temp_fig_nr2].update_yaxes(categoryorder='total descending')
                # Write images to png
                fig[temp_fig_nr1].write_image(f"./static/images/fig{temp_fig_nr1}.png")
                fig[temp_fig_nr2].write_image(f"./static/images/fig{temp_fig_nr2}.png")
                # Add a page break
                document.add_page_break()
                paragraphs6['1 cp'] = f"Figure {temp_fig_nr1}: Projects and Capital Demand per {SpatialFeatureChoice} (1/2)"
                par_formatter(paragraphs6)
                document.add_picture(f"./static/images/fig{temp_fig_nr1}.png", width=Cm(15))
                # Add a page break
                document.add_page_break()
                paragraphs7['1 cp'] = f"Figure {temp_fig_nr2}: Projects and Capital Demand per {SpatialFeatureChoice} (2/2)"
                par_formatter(paragraphs7)
                document.add_picture(f"./static/images/fig{temp_fig_nr2}.png", width=Cm(15))
    elif number_of_plots == 3:
                temp_fig_nr1 = f"{fig_nr}.1"
                temp_fig_nr2 = f"{fig_nr}.2"
                temp_fig_nr3 = f"{fig_nr}.3"
                fig[temp_fig_nr1] = px.bar(df_subset1, x=f'Projects per {SpatialFeatureChoice}', y=SpatialFeatureChoice,
                            color='Capital Demand',
                            height=900, orientation='h')
                fig[temp_fig_nr2] = px.bar(df_subset2, x=f'Projects per {SpatialFeatureChoice}', y=SpatialFeatureChoice,
                                color='Capital Demand',
                                height=900, orientation='h')
                fig[temp_fig_nr3] = px.bar(df_subset3, x=f'Projects per {SpatialFeatureChoice}', y=SpatialFeatureChoice,
                                color='Capital Demand',
                                height=900, orientation='h')
                # Sort images to follow each other
                fig[temp_fig_nr1].update_yaxes(categoryorder='total descending')
                fig[temp_fig_nr2].update_yaxes(categoryorder='total descending')
                fig[temp_fig_nr3].update_yaxes(categoryorder='total descending')
                # Write images to png
                fig[temp_fig_nr1].write_image(f"./static/images/fig{temp_fig_nr1}.png")
                fig[temp_fig_nr2].write_image(f"./static/images/fig{temp_fig_nr2}.png")
                fig[temp_fig_nr3].write_image(f"./static/images/fig{temp_fig_nr3}.png")
                # Add a page break
                document.add_page_break()
                paragraphs6['1 cp'] = f"Figure {temp_fig_nr1}: Projects and Capital Demand per {SpatialFeatureChoice} (1/3)"
                par_formatter(paragraphs6)
                document.add_picture(f"./static/images/fig{temp_fig_nr1}.png", width=Cm(15))
                # Add a page break
                document.add_page_break()
                paragraphs7['1 cp'] = f"Figure {temp_fig_nr2}: Projects and Capital Demand per {SpatialFeatureChoice} (2/3)"
                par_formatter(paragraphs7)
                document.add_picture(f"./static/images/fig{temp_fig_nr2}.png", width=Cm(15))
                # Add a page break
                document.add_page_break()
                paragraphs8['1 cp'] = f"Figure {temp_fig_nr3}: Projects and Capital Demand per {SpatialFeatureChoice} (3/3)"
                par_formatter(paragraphs8)
                document.add_picture(f"./static/images/fig{temp_fig_nr3}.png", width=Cm(15))
    elif number_of_plots == 4:
                temp_fig_nr1 = f"{fig_nr}.1"
                temp_fig_nr2 = f"{fig_nr}.2"
                temp_fig_nr3 = f"{fig_nr}.3"
                temp_fig_nr4 = f"{fig_nr}.4"
                fig[temp_fig_nr1] = px.bar(df_subset1, x=f'Projects per {SpatialFeatureChoice}', y=SpatialFeatureChoice,
                            color='Capital Demand',
                            height=900, orientation='h')
                fig[temp_fig_nr2] = px.bar(df_subset2, x=f'Projects per {SpatialFeatureChoice}', y=SpatialFeatureChoice,
                                color='Capital Demand',
                                height=900, orientation='h')
                fig[temp_fig_nr3] = px.bar(df_subset3, x=f'Projects per {SpatialFeatureChoice}', y=SpatialFeatureChoice,
                                color='Capital Demand',
                                height=900, orientation='h')
                fig[temp_fig_nr4] = px.bar(df_subset4, x=f'Projects per {SpatialFeatureChoice}', y=SpatialFeatureChoice,
                                color='Capital Demand',
                                height=900)
                # Sort images to follow each other
                fig[temp_fig_nr1].update_yaxes(categoryorder='total descending')
                fig[temp_fig_nr2].update_yaxes(categoryorder='total descending')
                fig[temp_fig_nr3].update_yaxes(categoryorder='total descending')
                fig[temp_fig_nr4].update_yaxes(categoryorder='total descending')
                # Write images to png
                fig[temp_fig_nr1].write_image(f"./static/images/fig{temp_fig_nr1}.png")
                fig[temp_fig_nr2].write_image(f"./static/images/fig{temp_fig_nr2}.png")
                fig[temp_fig_nr3].write_image(f"./static/images/fig{temp_fig_nr3}.png")
                fig[temp_fig_nr4].write_image(f"./static/images/fig{temp_fig_nr4}.png")
                # Add a page break
                document.add_page_break()
                paragraphs6['1 cp'] = f"Figure {temp_fig_nr1}: Projects and Capital Demand per {SpatialFeatureChoice} (1/4)"
                par_formatter(paragraphs6)
                document.add_picture(f"./static/images/fig{temp_fig_nr1}.png", width=Cm(15))
                # Add a page break
                document.add_page_break()
                paragraphs7['1 cp'] = f"Figure {temp_fig_nr2}: Projects and Capital Demand per {SpatialFeatureChoice} (2/4)"
                par_formatter(paragraphs7)
                document.add_picture(f"./static/images/fig{temp_fig_nr2}.png", width=Cm(15))
                # Add a page break
                document.add_page_break()
                paragraphs8['1 cp'] = f"Figure {temp_fig_nr3}: Projects and Capital Demand per {SpatialFeatureChoice} (3/4)"
                par_formatter(paragraphs8)
                document.add_picture(f"./static/images/fig{temp_fig_nr3}.png", width=Cm(15))
                # Add a page break
                document.add_page_break()
                paragraphs9['1 cp'] = f"Figure {temp_fig_nr4}: Projects and Capital Demand per {SpatialFeatureChoice} (4/4)"
                par_formatter(paragraphs9)
                document.add_picture(f"./static/images/fig{temp_fig_nr4}.png", width=Cm(15))

    fig_nr += 1

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

