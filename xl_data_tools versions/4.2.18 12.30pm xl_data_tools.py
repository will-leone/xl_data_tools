import os
import itertools
import xlwings as xw
import xlsxwriter
import datetime
import string
import subprocess

"""
This module builds on xlwings and xlsxwriter to provide convenient objects for pulling, cleaning, and writing data
between Excel and Python. It also includes functions for common tasks needed to locate and timestamp Excel file names.
"""


def mod_date(foo):
    """
    :param foo: file's name in the current directory; if given path, searches in the chosen directory
    :return: file modification date
    """
    if foo == os.path.split(foo)[0]:
        t = os.path.getmtime(foo)
        date = datetime.datetime.fromtimestamp(t)
    else:
        os.chdir(os.path.split(foo)[0])
        return mod_date(os.path.split(foo)[1])
    return date


def find_file(dir_path, name):
    """
    Searches for the newest version of a given file. Files are assumed to be formatted "d1.d2.d3 a1 a2 a3.E" ("D A.E")
    where D is the date with period separators; A is the core filename with spaces as separators; and E is the extension
    with a preceding period.
    :param dir_path: directory containing the desired file
    :param name: keywords from the name of the desired file
    :return: path of the desired file
    """

    os.chdir(dir_path)
    dir_list = os.listdir(dir_path)
    matches = list()
    for item in dir_list:
        if os.path.isfile(os.path.join(dir_path, item)):
            item_list = item.split()
            name_list = name.split()
            for lst in item_list, name_list:
                del lst[-1]
                for unit in lst:
                    if "." in unit:
                        lst.remove(unit)
            if all(component in item_list for component in name_list):
                matches.append(item)
    matches.sort(key=mod_date)

    return os.path.join(dir_path, matches[-1])


def empty_check(lst):
    """
    Determines whether the nested n-layer list contains only empty and/or None-type items.
    :param lst: any list
    :return: True if the nested list is (a) a list and (b) contains only empty lists, type objects,
    or None; otherwise False
    """
    try:
        print(lst)
        if not lst:
            return True
        if isinstance(lst, str):
            return False
        else:
            return all(map(empty_check, lst))
    except TypeError:
        return True


def hide_excel(boolean):
    """
    Hides Excel from the user interface and suppresses alerts if the input value is True. This script must be run again
    with False input to enable viewing for output Excel files, after which all Excel processes are exited.
    :param boolean: True or False boolean constant
    """
    for app in xw.apps:
        app.display_alerts = not boolean
        app.screen_updating = not boolean
    if boolean is False:
        try:
            while True:
                if subprocess.check_call("TASKKILL /F /IM excel.exe"):  # if call fails, raises CalledProcessError
                    pass
                else:
                    break
        finally:
            return


def terminate_excel():
    """
    Terminates all running Excel processes in Windows OS
    """
    while True:
        if subprocess.check_call("TASKKILL /F /IM excel.exe"):  # if cmd return code !=0, raises CalledProcessError
            pass
        else:
            break


def create_zip(directory, zip_name, files):
    """
    Removes all existing .zip files in the chosen directory with the given zip_name and creates a new .zip file with
    this name that contains the chosen files.
    :param directory: The directory where the zip file will be created
    :param zip_name: The name of the new zip file
    :param files: The files to be zipped
    """
    # Compile zip archive for reports if not comprised of a singled file
    os.chdir(directory)
    if len(reports) > 1:
        with os.scandir(os.getcwd()) as scan:
            for entry in scan:
                if zip_name in str(entry):
                    os.remove(entry)
        for foo in files:
            with zipfile.ZipFile(zip_name, "a") as my_zip:
                my_zip.write(foo)


def email_reports(sender, recipients, cc, bcc, subject, attachments, html, attachments_dir, hmtl_dir):
    """
    Sends out an SMTP email using SSL, HTML content, and up to one attachment (including .zip). Recipients' names must
    have the form "required_first_name optional_middle_name optional_last_name".
    :param sender: Sequence (a, b) where a is the sender's email and b is their email account password
    :param recipients: Sequence of pairs (a, b) where a is the recipient's name and b is their email
    :param cc: Sequence of pairs (a, b) where a is the cc recipient's name and b is their email
    :param bcc: Sequence of pairs (a, b) where a is the bcc recipient's name and b is their email
    :param subject: Subject title for the email
    :param attachments: File name of the attachment (including .zip) - no more than 1 per email
    :param html: File name of the html script defining the email body's content
    :param attachments_dir: Directory containing the attachments
    :param hmtl_dir: Directory containing the html script
    """

    # Construct formatted strings of names and emails for Message module
    recipient_names, cc_names, bcc_names = list(), list(), list()
    recipient_emails, cc_emails, bcc_emails = list(), list(), list()
    contact_names = {recipients: recipient_names, cc: cc_names, bcc: bcc_names}
    contact_emails = {recipients: recipient_emails, cc: cc_emails, bcc: bcc_emails}
    for group in contact_emails.keys():
        for pair in group:
            contact_names[group].append(pair[0].split()[0])
            contact_emails[group].append(pair[1])
        contact_names[group] = ", ".join(contact_names[group])
        contact_emails[group] = ", ".join(contact_emails[group])

    # Extract HTML content for email body
    os.chdir(html_dir)
    with open(html) as f:
        email_body = f.read()

    # Construct email
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = sender[0]
    msg['To'] = contact_emails[recipients]
    msg['Cc'] = contact_emails[cc]
    msg['Bcc'] = contact_emails[bcc]
    msg.set_content("""\
        <html>
          <head></head>
          <body>
          <body style="font-family:calibri; font-size: 16px" >
            <p> Hi, {}, </p>
            <p> {}
            </p>
          </body>
        </html>
        """.format(contact_names[recipients], email_body), subtype='html')

    # Prepare the attachment(s) for delivery
    os.chdir(attachments_dir)
    if len(reports) == 1:
        with open(reports[0], 'rb') as fp:
            msg.add_attachment(fp.read(), maintype="multipart", subtype="mixed", filename=reports[0])
    else:
        with open(rzip, 'rb') as myzip:
            msg.add_attachment(myzip.read(), maintype="multipart", subtype="mixed", filename=rzip)

    # Connect with the server and send the email with its attachment(s)
    with smtplib.SMTP(host='smtp.gmail.com', port=587) as s:
        context = ssl.create_default_context()
        s.starttls(context=context)
        s.login(sender[0], sender[1])
        s.send_message(msg)

    count -= 1

    return


class XlArray:
    """
    This class is meant for two-layer nested lists representing an Excel array: e.g., [[row_1], [row_2],...]
    """

    # Construct conversions between Excel array ranges and Pythonic indices
    alpha = string.ascii_uppercase
    extended_alpha = alpha + [x + y for x, y in zip(alpha, alpha)]
    convert = zip(range(1, len(extended_alpha) + 1), extended_alpha)
    convert_to_alpha = {x: y for x, y in convert}
    convert_to_num = {y: x for x, y in convert}

    def __init__(self, data, row, col):
        """
        :param data: Nested (or mono-layer) list representing an excel array (or row)
        :param row: Row location of the upper-left cell in the array (in Excel format, e.g., "2")
        :param col: Column location of the upper-left cell in the array (in Excel format - e.g., "B")
        """
        # If data is a mono-layer list (representing a row), convert it into a nested list (representing an array)
        if not all(itertools.starmap(isinstance, zip(data, [list] * len(data)))):
            data = [data]

        # Determine the finalized Excel array range for the row
        excel_range = col + str(row) + ":" + XlArray.convert_to_alpha[len(data) + XlArray.convert_to_num[col] - 1] + \
            str(row)

        # Set the instance variables
        self.data = data
        self.col = col
        self.row = row
        self.empty = empty_check(data)
        self.len = len(data)
        self.col_num = XlArray.convert_to_num[self.col]
        if not self.empty:
            self.last_col_num = self.col_num + self.len - 1
            self.last_col = XlArray.convert_to_alpha[self.last_col_num]
        else:
            self.last_col_num = self.col_num
            self.last_col = XlArray.convert_to_alpha[self.col_num]
        self.range = excel_range
        self.header = self.data[0]

    def empty(self, row_as_list):
        row_num = self.data.index(row_as_list)
        return empty_check(self.data[row_num])

    def remove(self, columns):
        """
        Removes the chosen columns in the instance's source array from the instance's own array with columns understood
        in Excel range terms.

        For instance, if the source array is [[a, b], [c,d]] with (row, col) = (2, "F"), the
        Excel interpretation is that the upper-left cell of the instance array is F2 while the range is F2:G3.
        Meanwhile, the instance's array's range is understood as [(i, j) for i, j in zip(range(2), range(2))].

        In the above case, self.remove("G") would reduce the source array to [[a], [c]] as "b" and "d" represent cells
        G2 and G3, respectively.

        :param columns: Columns in the source array in Excel's range interpretation - e.g., "A" for the 0th column
        """
        # Note that this section assumes no two rows/lists in the data array are identical due to list.index()
        for excluded_col in columns:
            excluded_col_num = XlArray.convert_to_num[excluded_col]     # e.g., column "B" becomes 2
            if excluded_col_num == self.col_num:                        # if the first column is to be excluded
                for record in self.data:
                    self.data[self.data.index(record)] = record[1:]     # remove the first column in all rows
                self.col = XlArray.convert_to_alpha[self.col_num + 1]   # adjust the Excel representation attributes
                self.col_num = XlArray.convert_to_num[self.col]
            elif excluded_col_num == self.last_col_num:                 # if the last column is to be excluded
                for record in self.data:
                    self.data[self.data.index(record)] = record[:len(self.data) - 1]
            elif self.col_num < excluded_col_num < self.last_col_num:   # if another column is to be excluded
                self.data = self.data[:excluded_col_num] + self.data[excluded_col_num + 1:]
            else:                                                       # if the column isn't in the instance array
                pass

    def filter(self, column, value, strict=True):
        """
        :param column: The column that will be searched in the array
        :param value: The cell content that will be searched for in the array
        :param strict: If true, the filter requires exact equivalence.
        :return: Filtered copy of the array with only those rows containing the desired entry in the desired column
        """
        filtered_array = list()
        filter_row = ""
        for record in self.data:                        # Here, rows are represented by lists
            if record[column] == value:                 # Strict equivalency required for a match
                if not filter_row:                      # Determine upper-left range value for the filtered array
                    filter_row = self.data.index(record) + self.row - 1
                filtered_array.append(record)
            elif not strict:
                if not filter_row:                      # Determine upper-left range value for the filtered array
                    filter_row = self.data.index(record) + self.row - 1
                try:
                    # if record[column] and value are splittable, see if all components of the former are in the latter
                    entry = record[column].split()
                    if all(entry[i] in value.split() for i in list(range(len(entry)))):
                        filtered_array.append(record)
                except TypeError:
                    pass

        return XlArray(filtered_array, filter_row, self.col)


class XlExtract:
    """
    Class Dependency: XlArray (for XlEdit.extract())

    Extract data from an existing Excel documents using the xlwings module.
    """

    def __init__(self, path):
        hide_excel(True)
        self.path = path
        self.name = os.path.split(path)[1]
        self.date = mod_date(path)
        self.wb = xw.books.open(self.path)
        self.sheets = self.wb.sheets

    def open(self):
        hide_excel(True)
        return self.wb

    def close(self):
        self.wb.close()
        hide_excel(False)
        return

    def init_sht(self):
        # create and name sheet
        pass

    def extract(self, exclude_sheets=None, exclude_cols=None):
        """
        Imports all data in the workbook with each sheet represented by a different XlArray object
        :param exclude_sheets: List of the names of the sheets from which data won't be collected
        :param exclude_cols: List of pairs (a,b) where a is the sheet name and b lists the columns to be excluded
        :return: Pairs consisting of each sheet number and the array in that sheet with all empty rows removed.
        """
        wb_data = list()
        for sht_name in [sheet for sheet in self.sheets if sheet not in exclude_sheets]:
            sht_xl = self.wb.sheets(sht_name)
            sht_range = sht_xl.cells
            sht_data = sht_xl.range(sht_range).value
            sht_array = XlArray(sht_data, 1, "A")
            for row in sht_array.data:
                if sht_array.data.empty(row):
                    sht_array.data.remove(row)
            for x_sheet, x_columns in exclude_cols:
                if x_sheet == sht_name:
                    sht_array.remove(x_columns)
            wb_data.append((sht_xl.index - 1, sht_array))  # sht.index is 1-based (as in Excel)
        return wb_data
        # create a range method here that opens a chosen sheet and scans it for the first completely empty row & column


class XlCreate:
    """
        Class Dependency: XlArray

        Write XlArray objects to an Excel file with easily-customized formatting.
    """
    def __init__(self, path):
        hide_excel(True)
        self.path = path
        self.name = os.path.split(path)[1]
        self.date = mod_date(path)
        self.wb = xw.books.open(self.path)
        self.sheets = self.wb.sheets

    def close(self):
        self.wb.close()
        hide_excel(False)
        return

    def init_sht(self):
        # create and name sheet
        pass

    def bind(self, sheet):
        pass

    """
        sht = report_writer.add_worksheet(wo_type)

        # Set the table creation parameters for each worksheet in the final reports
        # Determine table range
        convert = zip(range(1, len(string.ascii_uppercase) + 1), string.ascii_uppercase)
        convert_dict = {x: y for x, y in convert}
        header_end = convert_dict[len(header)]
        filtered_range = "A1:" + header_end + str(len(filtered_list) + 1)
        # Set table names
        table_name = "_".join(order_type_names[wo_type].split())
        # Word wrap each workbook, bold the header, and convert datetime
        wrap = report_writer.add_format({'text_wrap': 1})
        header_bold = report_writer.add_format({'bold': True, 'text_wrap': 1})
        date_format = report_writer.add_format({'num_format': 'm/d/yy'})
        # Create list of table headers
        header_lst = [{'header': col, 'header_format': header_bold} for col in header]

        # Insert the table and its data
        sht.add_table(filtered_range, {'columns': header_lst, 'name': table_name})
        for item in header:
            sht.write(0, header.index(item), item, header_bold)
        for m in range(len(filtered_list)):
            if filtered_list[m][0] == now:
                sht.write("A" + str(m + 2), "NO DATE", date_format)
            elif filtered_list[m][0] is None:
                print("None: ", filtered_list[m][0])
                sht.write("A" + str(m + 2), "NO DATE", date_format)
            else:
                sht.write_datetime("A" + str(m + 2), filtered_list[m][0], date_format)
            sht.write_row("B" + str(m + 2), filtered_list[m][1:], wrap)

        # Adjust the column widths
        for p in range(len(header)):
            len_lst = [len(str(cell[p])) for cell in filtered_list]
            if not len_lst:
                max_len = 16
            else:
                max_len = max(max(len_lst), 16)
            current_col = convert_dict[p + 1] + ":" + convert_dict[p + 1]
            sht.set_column(current_col, max_len)"""