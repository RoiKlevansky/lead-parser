# index.py

import re
import pandas as pd
from tkinter.filedialog import askopenfilename, asksaveasfilename
import xlsxwriter


class Lead:
    """
    Class representing a single lead(a form that is received automatically and contains contact information of a
    potential client)

    Attributes
    ----------
    phone : str
        The phone number of the client
    email : str
        The email of the client
    name : str
        The name of the client
    destination : str
        The trip destination
    people_num : str
        The number of people in the trip
    kosher : str
        The kosher status of the trip
    details : str
        The added details of the lead

    Methods
    -------
    to_list()
        Returns a list of the attributes of the lead
    """

    def __init__(self, phone='', email='', name='', destination='', people_num='', kosher='', details=''):
        """
        Parameters are the attributes of the lead
        :param phone: str - The phone number of the client
        :param email: str - The email of the client
        :param name: str - The name of the client
        :param destination: str - The trip destination
        :param people_num: str - The number of people in the trip
        :param kosher: str - The kosher status of the trip
        :param details: str - The added details of the lead
        """
        self.phone = phone
        self.email = email
        self.name = name
        self.destination = destination
        self.people_num = people_num
        self.kosher = kosher
        self.details = details

    def to_list(self):
        """
        Returns a list of the attributes of the lead
        :return: list - The attributes of the lead
        """
        return [self.phone, self.email, self.name, self.destination, self.people_num, self.kosher, self.details]


def parse_telecall_lead(lead: str) -> Lead:
    """
    Parses a lead from a telecall lead string
    :param lead: str - The lead string
    :return: Lead - The parsed lead
    """
    phone_regex = r'טלפון לחזרה\s*(\d+)'
    name_regex = r'שם מלא:\s*([ \'א-ת]+)'
    destination_regex = r'יעד נסיעה:\s*([ \'א-ת]+)'
    people_num_regex = r'מספר נפשות:\s*([ \dא-ת]+)'
    kosher_regex = r'סוג טיול:\s*([ א-ת]+)'
    details_regex = r'ההודעה:\s*([^\n\r]+)'

    lead_obj = Lead()
    phone_match = re.search(phone_regex, lead)
    if phone_match:
        lead_obj.phone = phone_match.group(1)
    name_match = re.search(name_regex, lead)
    if name_match:
        lead_obj.name = name_match.group(1)
    destination_match = re.search(destination_regex, lead)
    if destination_match:
        lead_obj.destination = destination_match.group(1)
    people_num_match = re.search(people_num_regex, lead)
    if people_num_match:
        lead_obj.people_num = people_num_match.group(1)
    kosher_match = re.search(kosher_regex, lead)
    if kosher_match:
        lead_obj.kosher = kosher_match.group(1)
    details_match = re.search(details_regex, lead)
    if details_match:
        lead_obj.details = details_match.group(1)
    return lead_obj


def parse_telecall_interested(lead: str) -> Lead:
    """
    Parses a lead from a telecall interested message string
    :param lead: str - The lead string
    :return: Lead - The parsed lead
    """
    phone_regex = r'טלפון לחזרה\s*(\d+)'
    name_regex = r'שם הפונה\s:\s*([ \'א-ת]+)'
    details_regex = r'ההודעה:\s*([^\n\r]+)'

    lead_obj = Lead()
    phone_match = re.search(phone_regex, lead)
    if phone_match:
        lead_obj.phone = phone_match.group(1)
    name_match = re.search(name_regex, lead)
    if name_match:
        lead_obj.name = name_match.group(1)
    details_match = re.search(details_regex, lead)
    if details_match:
        lead_obj.details = details_match.group(1)
    return lead_obj


def parse_telecall_general_message(lead: str) -> Lead:
    """
    Parses a lead from a telecall general message string
    :param lead: str - The lead string
    :return: Lead - The parsed lead
    """
    phone_regex = r'טלפון לחזרה\s*(\d+)'
    name_regex = r'שם מלא:\s*([ \'א-ת]+)'
    details_regex = r'ההודעה:\s*([^\n\r]+)'

    lead_obj = Lead()
    phone_match = re.search(phone_regex, lead)
    if phone_match:
        lead_obj.phone = phone_match.group(1)
    name_match = re.search(name_regex, lead)
    if name_match:
        lead_obj.name = name_match.group(1)
    details_match = re.search(details_regex, lead)
    if details_match:
        lead_obj.details = details_match.group(1)
    return lead_obj


def parse_telecall_customer_support(lead: str) -> Lead:
    """
    Parses a lead from a telecall customer support string
    :param lead: str - The lead string
    :return: Lead - The parsed lead
    """
    phone_regex = r'טלפון לחזרה\s*(\d+)'
    name_regex = r'שם מלא:\s*([ \'א-ת]+)'
    destination_regex = r'יעד נסיעה:\s*([ \'א-ת]+)'
    details_regex = r'ההודעה:\s*([^\n\r]+)'

    lead_obj = Lead()
    phone_match = re.search(phone_regex, lead)
    if phone_match:
        lead_obj.phone = phone_match.group(1)
    name_match = re.search(name_regex, lead)
    if name_match:
        lead_obj.name = name_match.group(1)
    destination_match = re.search(destination_regex, lead)
    if destination_match:
        lead_obj.destination = destination_match.group(1)
    details_match = re.search(details_regex, lead)
    if details_match:
        lead_obj.details = details_match.group(1)
    return lead_obj


def parse_virtual_chat_lead(lead: str) -> Lead:
    """
    Parses a lead from a virtual chat log string
    :param lead: str - The lead string
    :return: Lead - The parsed lead
    """
    # TODO: This format saves some data on the client embedded to the message itself,
    # for now this data is unused. Fix this someday.
    phone_regex = r'מספר טלפון:\s*(\d+)'
    email_regex = r'כתובת מייל:\s*(\w+@\w+.[.\w]+)'
    name_regex = r'משתמש\s([ \'א-תa-zA-Z]+)'

    lead_obj = Lead()
    phone_match = re.search(phone_regex, lead)
    if phone_match:
        lead_obj.phone = phone_match.group(1)
    email_match = re.search(email_regex, lead)
    if email_match:
        lead_obj.email = email_match.group(1)
    name_match = re.search(name_regex, lead)
    if name_match:
        lead_obj.name = name_match.group(1)
    # Pass the raw data for now, as not doing so can make us lose data.
    # lead_obj.details = lead
    return lead_obj


def parse_trip_page_lead(lead: str) -> Lead:
    """
    Parses a lead from a trip page(Contact Us) string
    :param lead: str - The lead string
    :return: Lead - The parsed lead
    """
    phone_regex = r'טלפון:\s*(\d+)'
    name_regex = r'שם(?: מלא)?:\s*([ \'א-ת]+)'
    email_regex = r'דואר (?:האלקטרוני|אלקטרוני):\s*(\w+@\w+.[.\w]+)'
    details_regex = r'(?:הודעה|נושא):\s*([^\n\r]+)'

    lead_obj = Lead()
    phone_match = re.search(phone_regex, lead)
    if phone_match:
        lead_obj.phone = phone_match.group(1)
    name_match = re.search(name_regex, lead)
    if name_match:
        lead_obj.name = name_match.group(1)
    email_match = re.search(email_regex, lead)
    if email_match:
        lead_obj.email = email_match.group(1)
    details_match = re.search(details_regex, lead)
    if details_match:
        lead_obj.details = details_match.group(1)
    return lead_obj


def parse_contact_us_form(lead: str) -> Lead:
    """
    Parses a lead from a contact us form string
    :param lead: str - The lead string
    :return: Lead - The parsed lead
    """
    phone_regex = r'טלפון:\s*(\d+)'
    name_regex = r'שם(?: מלא)?:\s*([ \'א-ת]+)'
    email_regex = r'דואר (?:האלקטרוני|אלקטרוני):\s*(\w+@\w+.[.\w]+)'
    details_regex = r'(?:הודעה|נושא):\s*([^\n\r]+)'

    lead_obj = Lead()
    phone_match = re.search(phone_regex, lead)
    if phone_match:
        lead_obj.phone = phone_match.group(1)
    name_match = re.search(name_regex, lead)
    if name_match:
        lead_obj.name = name_match.group(1)
    email_match = re.search(email_regex, lead)
    if email_match:
        lead_obj.email = email_match.group(1)
    details_match = re.search(details_regex, lead)
    if details_match:
        lead_obj.details = details_match.group(1)
    return lead_obj


def parse_lead(lead: str) -> Lead | None:
    """
    Parses a lead from a string, returns None if the lead is not parsable or should be ignored
    :param lead: str - The lead string
    :return: Lead - The parsed lead
    """
    title = lead[0]
    telecall_interested_notice_regex = r'^טלקול,הודעות עבור תיירות-'
    telecall_lead_regex = r'^טלקול,לקוחות חדשים-'
    telecall_general_message_regex = r'^טלקול,הודעות כלליות-'
    telecall_customer_support_regex = r'^טלקול,שירות לקוחות-'
    telecall_grouped_regex = r'טלקול, ריכוז הודעות'
    virtual_chat_lead_regex = r'^ליד חדש מהאתר'
    trip_page_lead_regex = r'^דף טיול'
    contact_us_regex = r'טופס יצירת קשר'
    disarmed_regex = r'(מכירה|רכישה)\s\d+\s(בוצעה|בוצע)\sבהצלחה'

    if re.search(telecall_interested_notice_regex, title):
        return parse_telecall_interested(lead[1])
    elif re.search(telecall_lead_regex, title):
        return parse_telecall_lead(lead[1])
    elif re.search(telecall_general_message_regex, title):
        return parse_telecall_general_message(lead[1])
    elif re.search(telecall_customer_support_regex, title):
        return parse_telecall_customer_support(lead[1])
    elif re.search(telecall_grouped_regex, title):
        return None
    elif re.search(virtual_chat_lead_regex, title):
        return parse_virtual_chat_lead(lead[1])
    elif re.search(trip_page_lead_regex, title):
        return parse_trip_page_lead(lead[1])
    elif re.search(contact_us_regex, title):
        return parse_contact_us_form(lead[1])
    elif re.search(disarmed_regex, title):
        return None
    else:
        lead_obj = Lead()
        lead_obj.details = lead[0] + '\n------\n' + lead[1]
        return lead_obj


def parse_leads_file(file_name: str) -> list:
    """
    Parses a file of leads, returns a list of raw leads
    :param file_name: str - The file name
    :return: list - The list of leads in format (lead_title, lead_body)
    """
    df = pd.read_csv(file_name, encoding='windows-1255')
    leads = []
    titles_list = df['נושא'].tolist()
    contents_list = df['גוף'].tolist()
    for i in range(len(titles_list)):
        leads.append((titles_list[i], contents_list[i]))
    return leads


def export_leads_to_xlsx(leads: list, file_name: str):
    """
    Exports a list of parsed leads to a xlsx file
    :param leads: list - The list of parsed leads
    :param file_name: str - The file name
    :return: None
    """
    workbook = xlsxwriter.Workbook(file_name)
    worksheet = workbook.add_worksheet()
    worksheet.right_to_left()

    row = 0
    col = 0
    fields = ['מספר טלפון', 'מייל', 'שם הפונה', 'יעד נסיעה', 'מספר נפשות', 'סוג טיול', 'הערות']
    header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
    for field in fields:
        worksheet.write(row, col, field, header_format)
        col += 1
    col = 0
    row = 1
    columns_widths = [20, 20, 20, 15, 10, 10, 50]
    for lead in leads:
        for item in lead.to_list():
            worksheet.write(row, col, item)
            col += 1
        col = 0
        row += 1
    for i in range(len(columns_widths)):
        worksheet.set_column(i, i, columns_widths[i])
    columns = [{'header': field} for field in fields]
    worksheet.add_table(0, 0, row - 1, len(fields) - 1, {'columns': columns})
    workbook.close()


def main():
    """
    Main function
    :return: None
    """
    raw_leads = parse_leads_file(askopenfilename())
    parsed_leads = []
    for lead in raw_leads:
        parsed_lead = parse_lead(lead)
        if parsed_lead is not None:
            parsed_leads.append(parsed_lead)
    files = [('All Files', '*.*'), ('Excel File', '*.xlsx')]
    file = asksaveasfilename(filetypes=files, defaultextension=".xlsx")
    export_leads_to_xlsx(parsed_leads, file)


if __name__ == '__main__':
    main()
