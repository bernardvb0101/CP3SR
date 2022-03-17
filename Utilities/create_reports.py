from docx import Document
from docx.shared import Inches
from datetime import datetime


def create_worddoc(username, municname, cost_dict, IFMS_dict, address_dict, fin_dict):
    """
    This module creates a MSWord File with all information required by the user to build a quote for a client.
    the dictionaries and variables passed to this document are required for the building of such quote and to
    provide background knowledge of the municipality to the user.
    """

    global heading3, paragraph3_
    document = Document()
    timenow = datetime.today().strftime('%Y-%m-%d-%H_%M_%S')  # Stamp the date and time
    word_file_name = f"{municname}_{username}_{timenow}.docx"

    doc_heading = f"Quotation: {municname} CP3 System Deployment"
    paragraph0_0 = f"This document is a MSWord copy of a quote that was pulled from the CP3C app at " \
                   f"{timenow} by {username} for {municname}."
    heading_0 = "0. General Information about the Municipality"
    paragraph0_1 = f"This document contains all the information that you would require to build a a quotation for the " \
                 f"CP3 system, based on the choices you have made in the CP3C web-application. Section 0 of this " \
                 f"document provides you with some background information about {municname}. You will find the address" \
                 f" and contact details that you would require to draft a letter as well as context to appraise the " \
                   f"quotation you are about issue to this client. All of the content that you might need for the " \
                   f"drafting of your quotation document is contained in the subsequent numbered sections of this " \
                   f"document, in the order that it is most likely to be presented and used."
    paragraph0_2 = "We hope you find this useful! Sincerely, The Novus3 Team."
    heading0_1 = "a. Addresses, Contact Details & Website"
    heading0_2 = "b. Insights into the Municipality's CapEx Situation and IFMS"
    heading0_3 = "c. Insights into the Municipality's Financial Situation"

    heading_1 = "1. Introduction"
    paragraph1_1 = f"Thank you for the opportunity that you have afforded us to provide you with this quotation for " \
                    f"the implementation of the CP3 solution at {municname}. The deployment of the CP System can be" \
                    f" done rapidly - we will agree with you on the most suitable timelines. Depending on the " \
                    f"circumstances, most CP3 clients can do their first user login after as little as 3 weeks after " \
                    f"contract initiation. This is normally followed up by a period of approximately two months " \
                    f"wherein our staff will assist with the data on-boarding and data clean-up process, working with " \
                    f"the appropriate officials in this process."

    heading_2 = "2. Purpose"
    paragraph2_1 = f"The purpose of this quotation is to provide {municname} with a formal quotation" \
                   f" for an enterprise license, maintenance and support as well as training support for the " \
                   f"Collaboration, Planning, Prioritisation, Performance (CP3) system, which supports critical " \
                   f"processes pertaining to the IDP budgeting cycle and informs business and financial processes in the" \
                   f" municipality."

    heading_3 = "3. CP3 Core Functionality"
    paragraph3_1 = f"The CP3 system is deployed standard with the core functionality that every municipality needs. " \
                   f"In addition, there are a number op optional modules, depenmdding on functional requirements. " \
                   f"The CP3 Solution for {municname} will be deployed with the following modules: "
    heading3_1 = f"3.1. Capture Module (Standard Module)"
    paragraph3_1_1 = f"This module is used by all officials in the municipality who are involved in the process of" \
                     f"capital project identification and capital requests. Capital projects may stem from a variety" \
                     f" of sources. Most municipalities ussually have a backlog list of capital projects. These usually " \
                     f"exist on a spreadsheet or on paper in a filing cabinet. Capital projects stem from the IDP process, " \
                     f"infrastructure masterplans, asset management plans, from development that takes place in the" \
                     f" municipality and a host of other possible sources. All relevant project details are captured " \
                     f"including spatial data. Data fields for capturing will be configured for {municname}'s unique " \
                     f"requirements. This caputre module is further used by heads of department and by the executive, " \
                     f"to rapidly appraised the state and quantum of project required within the municipality. It is " \
                     f"a powerful and visually appealing tool to rpaidly get a sense of what the requirements across " \
                     f"the various sectors (departments) of the municipality are. The Capture Module facilitates the " \
                     f"project preparation process - the Prioritise Module require that the data assocaited with each " \
                     f"project is complete to ascertain, amongst other elements, its readiness to proceed to the next " \
                     f"step. It is for this reason that external funding institutions derive significant comfort from " \
                     f"the fact that the municipality is using CP3."
    heading3_2 = f"3.2. Prioritise Module (Standard Module)"
    paragraph3_2_1 = f"The powerful and flexible CP3 prioritisation module will be used to set up a prioritisation " \
                   f"model that reflects the needs of your municipality. There is no 'one-size-fits-all' solution " \
                   f"although our experienced staff will assist you by providing some best-practice guidelines gleaned " \
                   f"from working with a diversity of other municipalities. The prioritisation module helps you to " \
                   f"distill structure and order from a myriad of capital demands, taking into account a range of " \
                   f"considerations unique and specific to your municipality. The is a vital and very important step " \
                   f"in the journey towards the development of a defendable and outcomes-based budget for your " \
                   f"municipality. CP3 has often saved the day at council meetings, having the ability to provide rapid" \
                   f" evidence on how the priorities of the municipality was determined. Furthermore, you can rest " \
                   f"assured that the CP3 prioritisation module is simply a tool, it is not a rule. It is your " \
                   f"municipality's prorogative to override certain of the priorities and ommitt some of the projects " \
                   f"that may have scored well and introduce others that may not have. There are many reassons where " \
                   f"this may be required and this process is therefore elegantly engineered into the CP3 workflow " \
                   f"process."
    heading3_3 = f"3.3. Budget Fit Module (Standard Module)"
    paragraph3_3_1 = f"The Budget Fit Module is a dynamic budget development tool that performs the last vital steps " \
                     f"towards the creation of a defendable and affordable budget. The finalised budget that is " \
                     f"developed by the CP3 Budget Fit Module is exportable in a number of pre-formatted reports " \
                     f"that complies fully with National Treasury's specifications (e.g. SA6-, SA36- and MSCOA reports. " \
                     f"All budgets developed by CP3 already carriers the correct funding sources as " \
                     f"officials can only select funding sources that are applicable to their specific projects. " \
                     f"Furthermore, there is seamless integration with your municipality's Financial Management " \
                     f"System (FMS) through API protocols that seamlessly passed information from system to system " \
                     f"without any human intervention required once it has been set up. The budget fit allows for " \
                     f"an incredible amount of flexibility and testing numerous 'What if?' scenarios. The CP3 system " \
                     f"can be asked to perform the allocations per department in accordance with the prioritisation " \
                     f"model or alternatively, certain minimum pre-allocations can be made to each department or " \
                     f"division whilst leaving the powerful CP3 engine to fit project within these allocations and " \
                     f"subsequently augment and/or supplement these allocations based on the model outcomes. Budget " \
                     f"portions can be set aside thematically e.g. in terms of a policy or resolution that may have " \
                     f"been passed by the Executive of the municipality. An example can be to set aside 5% of the " \
                     f"budget for carbon-offset or climate change and resilience projects."
    heading3_4 = f"3.4. Tracking Module (Standard Module)"
    paragraph3_4_1 = "The tracking module is the municipality's trusted tool for the tracking of the implementation " \
                     "of capital projects. Project managers are required to plan their cash-flows and project " \
                     "implementation steps and milestones on the tracking module. This enables the municipality to " \
                     "perform in-year reporting with efficiency. In addition, the progress of project programmes as " \
                     "well as the progress of individual projects can be tracked and monitored on the system. " \
                     "Comparisons can be made between the progress of budget expenditure and physical progress - flags " \
                     "can be raised if serious descrepancies start to emerge. The system facilitates the ability to " \
                     "appraise the overall status of the municipality of being able to meat its expenditure targets " \
                     "and plays a vital role in the budget adjustment process half-way through the financial year. " \
                     "Project views and reports are in familiar formats that conforms with PMBOK and government's " \
                     "SIPDM principles. Gantt-charts and cash-flows can easily be exported and used in presentations and" \
                     " other reporting documentation."
    # Build a modules-chosen list
    # Enter the standard modules 1st
    modules_list = ['Capture Module', 'Prioritise Module', 'Budget Fit Module', 'Tracking Module']
    # Find the additional modules in the dictionary and aad them
    for key in cost_dict:
        if 'EIM' in key and not any('EIM' in string for string in modules_list):
            modules_list.append('Economic Impact Module (EIM)')
        if 'SIM' in key and not any('SIM' in string for string in modules_list):
            modules_list.append('Socio-Economic Impact Module (SIM)')
        if 'BAM' in key and not any('BAM' in string for string in modules_list):
            modules_list.append('Budget Adjustment Module (BAM)')
        if 'SGM' in key and not any('SGM' in string for string in modules_list):
            modules_list.append('Stage Gate Module (SGM)')
        if 'IGPM' in key and not any('IGPM' in string for string in modules_list):
            modules_list.append('Inter-Govermental Project Module (IGPM)')

    # Now create the headings and content for the additional modules.
    if len(modules_list) > 4:  # additional modules were chosen
        heading3 = {}
        paragraph3_ = {}
        for counter in range(5, len(modules_list) + 1):
            heading3[counter] = f"3.{counter} {modules_list[counter - 1]} (Optional Module)"
            if 'EIM' in modules_list[counter - 1]:
                paragraph3_[
                    counter] = f"The Economic Impact Module (EIM) is built bespoke for {municname}. The Economic " \
                               f"model is not a strategic model but an reactive model. A number of 4 " \
                               f"different economic multipliers can be chosen from. Assistance will be rendered" \
                               f" in this regard. The Economic Impact Module plays a major role in the " \
                               f"enhancement of the system. The economic multipliers are used in the " \
                               f"prioritisation module which renders the process of prioritisation much more " \
                               f"circumspective on the long term effects of expenditure. Significant insights " \
                               f"are gleaned from the economic perspective provided on prioritised budget " \
                               f"scenarios. The Economic Impact Module furthermore plays a major role when it " \
                               f"comes to reporting on the impact of the proposed budget on the community. " \
                               f"An economic report is available for downloading from the CP3 system that will " \
                               f"provide you with a host of useful metrics to report on."
            elif 'SIM' in modules_list[counter - 1]:
                paragraph3_[counter] = f"The Socio- Economic Impact Module is built bespoke for {municname}. This " \
                                       f"model is not a strategic model but an reactive model. It is a specifically " \
                                       f"focussed version of the normal Economic Impact Module. The multipliers that " \
                                       f"are chosen for the SIM will typically include elements such as job-creation, " \
                                       f"job-absorbtion, liveability indices, and so on. Assistance in making the most " \
                                       f"suitable choices will be rendered by the support team." \
                                       f" The Socio-Economic Impact Module plays a major role in the " \
                                       f"enhancement of the system. The SIM multipliers are used in the " \
                                       f"prioritisation module which renders the process of prioritisation much more " \
                                       f"circumspective on the long term effects of expenditure. Significant insights " \
                                       f"are gleaned from the socio-economic perspective provided on prioritised budget " \
                                       f"scenarios. The SIM furthermore plays a major role when it " \
                                       f"comes to reporting on the impact of the proposed budget on the community. " \
                                       f"An economic report is available for downloading from the CP3 system that will " \
                                       f"provide you with a host of useful metrics to report on."
            elif 'BAM' in modules_list[counter - 1]:
                paragraph3_[counter] = f"The Budget Adjustment Module is recommended for municipalities with numerous" \
                                       f" projects and a number of asssets-under construction on an ongoing basis. " \
                                       f"The budget adjustment process is an important process that seeks " \
                                       f"to optimise the expenditure of allocated budget for the financial year. The " \
                                       f"BAM facilitates an innovative process wherein budgets can be moved around " \
                                       f"between projects and the reasons for these actions can be recorded against " \
                                       f"each project. The application of the BAM in conjunction with the Budget Fit " \
                                       f"module that comes standard with the CP3 system, has proven itself to funtion " \
                                       f"as a vital and invaluable tool for a number of municipalities."
            elif 'SGM' in modules_list[counter - 1]:
                paragraph3_[counter] = f"The Stage Gate Module (SGM) facilitates the governance component of the CP3 " \
                                       f"system. The entire project life-cycle is facilitated on the system from the " \
                                       f"stage where the project is conceieved right through to implementation and " \
                                       f"sign-off on the last certificate of completion. Sign-off at each stage-gate " \
                                       f"is required by designated managers who are allocated with specific roles and " \
                                       f"user-rights on the CP3 syste. Projects are not allowed to advance to a new " \
                                       f"stage-gate unless evidence is provided of compliance with the current stage " \
                                       f"where-in the project finds itself. Budgets demands becomes stage specific when" \
                                       f"the SGM is activated. The project tree in the Capture Module carries the " \
                                       f"specific number of the stage where each project resides. This makes it easy " \
                                       f"to see, at a glance, how many projects resides in which stages. Altough all " \
                                       f"muniipalities should make use of the stage gate methodology, it is advisable " \
                                       f"to mature organisationally towards this objective."
            elif 'IGPM' in modules_list[counter - 1]:
                paragraph3_[counter] = f"The Inter-Governmental Project Module (SGM) allows {municname} to have sight" \
                                       f" on projects that are undertaken by provincial and national departments " \
                                       f"within your urban boundaries, or alternatively in proximity to it. This " \
                                       f"ability provides a vital insight and input into the prioritisation and " \
                                       f"allocation of capital. Form a legislative point of view, you are required to" \
                                       f" report on the activities of these entities in context if your own capital " \
                                       f"demand and subsequent allocations."

    heading_4 = "4. Background"
    paragraph4_1 = f"Novus3 was among the 1st practitioners to realise that the process of arriving at a defendable, " \
                   f"outcomes-based budget is more important than the budget itself. Because: “When developing a budget," \
                   f" you are prioritising, and by prioritising, you are creating a defendable budget. Novus3’s CP3 " \
                   f"software solution, managed to incorporate all of the elements required to strategically and " \
                   f"transparently incorporate all elements that are required for the sustainable Capital Investment " \
                   f"Planning process in the local government space. Managing capital and investment at a local " \
                   f"government level, is far more complex than most would appreciate. There is legislation, the IDP " \
                   f"process, technical inputs, strategic objectives, environmental and climate change objectives, " \
                   f"economic objectives, financial constraints and much more to take into account."
    if IFMS_dict['IFMS'] == 'Munsoft':
        paragraph4_2 = f"The CP3 system was developed by- and is supported and licensed by Novus3 (Pty) Ltd. Novus3 has " \
                       f"developed 14 CEFs so far and has successfully implemented and supported the CP3 system at a number " \
                       f"of municipalities ranging in size, including metros. On-boarding CP3 is an exciting and efficient " \
                       f"process, and it results in enduring, positive changes and improved financial outcomes at all " \
                       f"municiaplities that make use of the system.  Your current financial system is {IFMS_dict['IFMS']}." \
                       f"Novus3 and Munsoft are sibling companies and the implementation therefore comes with the " \
                       f"assurance that integration between the systems are seamless and unincumbered."
    else:
        paragraph4_2 = f"The CP3 system was developed by- and is supported and licensed by Novus3 (Pty) Ltd. Novus3 has " \
                       f"developed 14 CEFs so far and has successfully implemented and supported the CP3 system at a number " \
                       f"of municipalities ranging in size, including metros. On-boarding CP3 is an exciting and efficient " \
                       f"process, and it results in enduring, positive changes and improved financial outcomes at all " \
                       f"municiaplities that make use of the system.  Your current financial system is {IFMS_dict['IFMS']}." \
                       f" the CP3 system can integrate with any external system through API protocols that would not " \
                       f"require any human intervention (as per National Treasury's requirements). However, should " \
                       f"the financial service provider not be ready for integration, this can be done in a number of " \
                       f"tried and tested ways which will be discussed with your relevant official during on-boarding."

    heading_5 = "5. Deliverables"
    heading5_1 = "5.1. CP3 Deployment, Configuration and Cloud Hosting"
    paragraph5_1_1 = f"The CP3 system will be configured spatially and technically to {municname}'s specific needs." \
                   f"Technical fields will be set up after initial consultation with your officials regarding " \
                   f"the prioritisation model's elements. The technical data ultimately feeds the prioritisation " \
                   f"model with the data it needs to meaningfully prioritise.  Spatially, the system will centre " \
                   f"around your municipal area to easy the process of place finding. A list of core users will be " \
                   f"obtained from you and the respective roles and proposed user-rights will be discussed with you. " \
                   f"Thereafter, users will have to be invited from the project's designated owners within your " \
                   f"municipality."
    paragraph5_1_2 = f"On the server side, a specific cloud-hosted city is registred with you municipality's name in " \
                   f"the url. A host of security measures are implemented to render the site secure and impervious to " \
                   f"external logins or attacks. Regular server monitoring is undertaken and daily back-ups are performed." \
                   f" A full audit trail of all user activity is kept. A development team is on retainer with a specific " \
                   f"amount of designated support hours for your system specifically to ensure that any bugs are fixed " \
                   f"within the SLA and that any reasonable configuration requests that could not be configured " \
                   f"administratively, be incorporated with expediency."

    heading5_2 = "5.2. Specialist System and Capital Investment Planning Support as well as Training Support"
    paragraph5_2_1 = f"The following items relate to ad-hoc specialist system and municipal capital investment planning " \
                     f"support:"
    paragraph5_2_2 = "The core system module support and maintenance include, but are not limited to, helpdesk support, " \
                     "system administration and configuration requests (such as new MSCOA tables, new departmental " \
                     "structure etc.), continual basic system operation support, budget scenarios and template " \
                     "generation and uploading, prioritisation model development, grant allocation goal-seeking, " \
                     "specialised reporting requirements, spatial analysis support, IDP process advisory support, etc."
    paragraph5_2_3 = "It is important to note that we do understand that some months will be more active, and some " \
                     "less active.  We also understand that some requests / support needed require more specialised " \
                     "resources, or less specialised resources. The monthly fixed-fee, was therefore estimated from a " \
                     "blended rate which enables periods of inactivity to offset increased utilisation of more " \
                     "specialised resources when required."
    paragraph5_2_4 = "The following items relate to ad-hoc system training support:"
    paragraph5_2_5 = "System training is provided for General Users and Admin Users."
    paragraph5_2_6 = "Two separate training courses with modules are available for General and Admin users based on " \
                     "the roles and functions performed by each of the user types."
    paragraph5_2_7 = "A schedule of training modules is provided for General and Admin user training together with " \
                     "an estimated duration of time required to cover the content of each of the training modules."
    paragraph5_2_8 = "All training events will be hosted as online training events."
    paragraph5_2_9 = "A maximum of 10 participants will be allowed per training event to ensure that the trainer / " \
                     "moderator can provide focused individual attention and Q&A responses if required during the " \
                     "session."
    paragraph5_2_10 = "Training sessions will be set up in collaboration with the municipality based on specific " \
                      "focus areas and participant numbers per training event."
    paragraph5_2_11 = "Training is inclusive of contextual preparation, presentation, debriefing and follow-up."
    paragraph5_2_12 = "Training will be invoiced on a monthly basis at a fixed rate.  It is accepted that training may " \
                      "be very intensive over certain periods whilst other months may be very quiet in terms of " \
                      "training. The monthly fixed-fee, was therefore estimated from a blended rate which enables " \
                      "periods of inactivity to offset increased utilisation of more specialised resources when " \
                      "required."
    paragraph5_2_13 = "The bulk of the training support will be provided via online methods. In the event that on-site " \
                      "support is required, disbursement costs will be claimed against the support provision at client " \
                      "approved rates. Training venues will be provided by the client."
    paragraph5_2_14 = "Support and training will be invoiced based on monthly fixed fee as outlined in the " \
                      "previous point. Supportive evidence, and progress reports will be provided along with the " \
                      "monthly invoices for this purpose."

    heading_6 = "6. Costing"

    paragraph6_2 = f"The metrics used to determine the price point quoted for {municname} are listed in the costing " \
                   f"table below namely:"
    paragraph6_3 = "The annual approved capital expenditure of the municipality;"
    paragraph6_4 = "The approximate number of approved or funded projects;"
    paragraph6_5 = "The approximate number of projects seeking funding (wish-list)."
    paragraph6_6 = "These numbers are estimated from information that are published in the public domain. " \
                   "We often find that when we contract, the numbers are lower but as soon as we are onboarded, the " \
                   "numbers steadily rise to the levels that we have estimated due to organisation that is injected " \
                   "through the use of the CP3 product.  These numbers inform us on the elements that are the key cost " \
                   "drivers namely:"
    paragraph6_7 = "The levels of 1st, 2nd and 3rd tier support that would be required in terms of hours spent;"
    paragraph6_8 = "Maintenance and support hours required for ongoing system configuration, online help, bug fixing, " \
                   "etc."
    paragraph6_9 = "The size of the serve and the licensing costs of the supporting software to cater for the number " \
                   "of users and size of the data that has to be stored and processed."
    paragraph6_10 = "The estimated amount of effort to integrate with existing our legacy software."
    paragraph6_11 = "Our cost calculator reads these metrics from a database and determines the licensing cost as " \
                    "well as the cost for support and maintenance as explained above. "

    # Document Main Heading
    document.add_heading(doc_heading, 0)
    document.add_paragraph(paragraph0_0)
    document.add_paragraph(paragraph0_1)
    document.add_picture('CP3logo.png', width=Inches(1.25))
    document.add_paragraph(paragraph0_2)

    # Add a page break
    document.add_page_break()

    # Add a Page with General Information regarding the Municipality
    document.add_heading(heading_0, level=1)

    document.add_heading(heading0_1, level=2)
    table = document.add_table(rows=1, cols=2, style='Light Grid Accent 1')
    heading_cells = table.rows[0].cells
    heading_cells[0].text = 'Description'
    heading_cells[1].text = 'Info'

    for the_key, the_value in address_dict.items():
        cells = table.add_row().cells
        cells[0].text = the_key
        cells[1].text = the_value

    # Add a page break
    document.add_page_break()

    # Add a Page with  Information regarding the Municipality's CapEx and IFMS
    document.add_heading(heading0_2, level=2)
    table2 = document.add_table(rows=1, cols=2, style='Light Grid Accent 1')
    heading_cells2 = table2.rows[0].cells
    heading_cells2[0].text = 'Description'
    heading_cells2[1].text = 'Info'

    for key, value in IFMS_dict.items():
        cells = table2.add_row().cells
        cells[0].text = key
        cells[1].text = str(value)

    # Add a page break
    document.add_page_break()

    # Add a Page with  Information regarding the Municipality's Financial Situation
    document.add_heading(heading0_3, level=2)
    table3 = document.add_table(rows=1, cols=2, style='Light Grid Accent 1')
    heading_cells3 = table3.rows[0].cells
    heading_cells3[0].text = 'Description'
    heading_cells3[1].text = 'Info'

    for key, value in fin_dict.items():
        cells = table3.add_row().cells
        cells[0].text = key
        cells[1].text = str(value)

    # Add a page break
    document.add_page_break()

    # 1st Heading and paragraph
    document.add_heading(heading_1, level=1)
    document.add_paragraph(paragraph1_1)

    # 2nd Heading and paragraphs
    document.add_heading(heading_2, level=1)
    document.add_paragraph(paragraph2_1)

    # 3rd Heading, sub-headings and paragraphs
    # Provide a breakdown of the 4 default modeules 1st
    document.add_heading(heading_3, level=1)
    document.add_paragraph(paragraph3_1)
    # Enter the chosen modules as bullets:
    for item in modules_list:
        document.add_paragraph(item, style='List Number')
    document.add_heading(heading3_1, level=2)
    document.add_paragraph(paragraph3_1_1)
    document.add_heading(heading3_2, level=2)
    document.add_paragraph(paragraph3_2_1)
    document.add_heading(heading3_3, level=2)
    document.add_paragraph(paragraph3_3_1)
    document.add_heading(heading3_4, level=2)
    document.add_paragraph(paragraph3_4_1)

    # Enter additional module headings and content if it was selected.
    for counter in range(5, len(modules_list)+1):
        document.add_heading(heading3[counter], level=2)
        document.add_paragraph(paragraph3_[counter])

    # 4th Heading and paragraphs
    document.add_heading(heading_4, level=1)
    document.add_paragraph(paragraph4_1)
    document.add_paragraph(paragraph4_2)

    # 5th Heading and paragraphs
    document.add_heading(heading_5, level=1)
    document.add_heading(heading5_1, level=2)
    document.add_paragraph(paragraph5_1_1)
    document.add_paragraph(paragraph5_1_2)

    document.add_heading(heading5_2, level=2)
    document.add_paragraph(paragraph5_2_1)
    document.add_paragraph(paragraph5_2_2, style='List Bullet')
    document.add_paragraph(paragraph5_2_3, style='List Bullet')
    document.add_paragraph(paragraph5_2_4)
    document.add_paragraph(paragraph5_2_5, style='List Bullet')
    document.add_paragraph(paragraph5_2_6, style='List Bullet')
    document.add_paragraph(paragraph5_2_7, style='List Bullet')
    document.add_paragraph(paragraph5_2_8, style='List Bullet')
    document.add_paragraph(paragraph5_2_9, style='List Bullet')
    document.add_paragraph(paragraph5_2_10, style='List Bullet')
    document.add_paragraph(paragraph5_2_11, style='List Bullet')
    document.add_paragraph(paragraph5_2_12, style='List Bullet')
    document.add_paragraph(paragraph5_2_13, style='List Bullet')
    document.add_paragraph(paragraph5_2_14, style='List Bullet')

    # 6th Heading and paragraphs
    document.add_heading(heading_6, level=1)
    document.add_paragraph(paragraph6_2)
    document.add_paragraph(paragraph6_3, style='List Bullet')
    document.add_paragraph(paragraph6_4, style='List Bullet')
    document.add_paragraph(paragraph6_5, style='List Bullet')
    document.add_paragraph(paragraph6_6)
    document.add_paragraph(paragraph6_7, style='List Bullet')
    document.add_paragraph(paragraph6_8, style='List Bullet')
    document.add_paragraph(paragraph6_9, style='List Bullet')
    document.add_paragraph(paragraph6_10, style='List Bullet')
    document.add_paragraph(paragraph6_11)

    # Add a Page with  the quotation provided

    table4 = document.add_table(rows=1, cols=2, style='Light Grid Accent 1')
    heading_cells3 = table4.rows[0].cells
    heading_cells3[0].text = 'Description'
    heading_cells3[1].text = 'Info'

    for key, value in cost_dict.items():
        cells = table4.add_row().cells
        cells[0].text = key
        cells[1].text = str(value)

    full_path = f"./DOWNLOAD_FOLDER/{word_file_name}"
    document.save(full_path)

    return full_path

def make_msword_copy(username, municname, email_address, email_dict):
    """
    This function makes a word copy of a quote generated by email for record keeping and verification later.
    """
    document = Document()
    timenow = datetime.today().strftime('%Y-%m-%d-%H_%M_%S')  # Stamp the date and time
    word_file_name = f"{municname}_{username}_{timenow}.docx"

    doc_heading = f"Emailed Quotation: {municname} CP3 System Deployment"

    paragraph0_1 = f"This document is a MSWord copy of an email quote that was pulled from the CP3C app at " \
                   f"{timenow} by {username} for {municname}. The following table with costs was emailed to " \
                   f"{email_address}:"

    document.add_heading(doc_heading, 0)
    document.add_paragraph(paragraph0_1)

    # Add a page break
    document.add_page_break()

    table4 = document.add_table(rows=1, cols=2, style='Light Grid Accent 1')
    heading_cells3 = table4.rows[0].cells
    heading_cells3[0].text = 'Description'
    heading_cells3[1].text = 'Info'

    for key, value in email_dict.items():
        cells = table4.add_row().cells
        cells[0].text = key
        cells[1].text = str(value)

    full_path = f"./DOWNLOAD_FOLDER/{word_file_name}"
    document.save(full_path)

    return full_path