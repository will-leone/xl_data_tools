import os
import csv
import itertools
import datetime
import string
import subprocess
from email.message import EmailMessage
import smtplib
import ssl
import zipfile
import xlwings as xw
import xlsxwriter

"""
This module provides convenient objects for pulling,
cleaning, and writing data between Excel and Python. 
It includes functions for common tasks needed to
locate and timestamp Excel file names.
"""


def remove_files(path, exclude=None):
    """
    :param path: Directory containing folders to be purged
    :param exclude: Folders to be left unmodified
    """
    with os.scandir(path) as iter_dir:
        for subdir in iter_dir:
            if os.DirEntry.is_dir(subdir) and (not exclude or all(exclude not in subdir.name)):
                with os.scandir(os.path.join(path, subdir)) as iter_subdir:
                    for item in iter_subdir:
                        os.remove(os.path.join(path, subdir, item))
    return


def mod_date(foo):
    """
    :param foo: path or path-like object representing a file
    :return: file modification date
    Requires Python version 3.6+ to accept path-like objects.
    """
    print(foo)
    if foo == os.path.split(foo)[1]:
        t = os.path.getmtime(foo)
        date = datetime.datetime.fromtimestamp(t)
    else:
        date = mod_date(os.path.split(foo)[1])
    return date


def find_file(dir_path, keywords):
    """
    Searches for the newest version of a given file.
    :param dir_path: directory containing the desired file
    :param keywords: string of keywords from the keywords of the desired file
    :return: path of the desired file
    """
    dir_list = os.listdir(dir_path)
    if isinstance(keywords, str):
        keywords = keywords.split()
    matches = list()
    initial_dir = os.getcwd()
    os.chdir(dir_path)
    for item in dir_list:
        while "." in item:
            loc = item.find(".")
            if loc == len(item) - 1:
                item = item[:-1]
            else:
                item = item[:loc] + item[loc + 1:]
        if os.path.isfile(os.path.join(dir_path, item)):
            item_list = item.split()
            if all(component in item_list for component in keywords):
                matches.append(item)
    if not matches:
        print(f"There is no file containing keywords '{keywords}' in"
               f"{dir_path}.")
    else:
        matches.sort(key=mod_date)
    os.chdir(initial_dir)

    return os.path.join(dir_path, matches[-1])


def empty_check(lst):
    """
    Determines whether the nested n-layer list contains only empty
    and/or None-type items.
    :param lst: any list, integer, float, or string
    :return: True if the nested list is (a) a list and (b) contains
    only empty lists, type objects, or None; otherwise, False
    """
    try:
        if not lst:
            return True
        if (isinstance(lst, str) or isinstance(lst, int) or
                isinstance(lst, float)):
            return False
        else:
            return all(map(empty_check, lst))
    except TypeError:
        # This indicates that lst contains None as an object
        return True


def terminate_excel():
    """
    Terminates all running Excel processes in Windows OS
    """
    while True:
        try:
            subprocess.check_call("TASKKILL /F /IM excel.exe")
        except subprocess.CalledProcessError:
            break
    return


def hide_excel(boolean):
    """
    Hides Excel from the user interface and suppresses alerts if the
    input value is True. This script must be run again with False
    input to enable viewing for output Excel files, after which all
    Excel processes are exited.
    :param boolean: True or False boolean constant
    """
    for app in xw.apps:
        app.display_alerts = not boolean
        app.screen_updating = not boolean
    if boolean is False:
        terminate_excel()
    return


def csv_extract(file, directory, header=None):
    """
    Converts a given CSV file into a pandas dataframe.
    :param file: Name of the CSV file
    :param directory: Name of the directory containing the CSV file
    :param header: Sequence containing all columns from the CSV to be
    included in the output. If None, the CSV's first line will be used.
    :return: pandas dataframe
    """
    initial_dir = os.getcwd()
    os.chdir(directory)
    with open(file, newline='') as csvfile:
        reader = csv.DictReader(csvfile, fieldnames=header)
        for row in reader:
            new_key = row[header[0]]
            if new_key is not None and new_key != "":
                csv_dict[new_key] = list()
                for column in header[1:]:
                    csv_dict[new_key].append(row[header[column]])
    os.chdir(initial_dir)
    return csv_dict


def create_zip(directory, zip_name, files):
    """
    Removes all existing .zip files in the chosen directory with the given
    zip_name and creates a new .zip file with
    this name that contains the chosen files.
    :param directory: The directory where the zip file will be created
    :param zip_name: The name of the new zip file
    :param files: List of the files to be zipped (as filenames)
    """
    # Compile zip archive for reports if not comprised of a singled file
    initial_dir = os.getcwd()
    os.chdir(directory)
    if len(files) > 1:
        with os.scandir(os.getcwd()) as scan:
            for entry in scan:
                if zip_name in str(entry):
                    os.remove(entry)
        for foo in files:
            with zipfile.ZipFile(zip_name, "a") as my_zip:
                my_zip.write(foo)
    os.chdir(initial_dir)


def send_email(sender, recipients, subject, html, html_dir, cc=None,
               bcc=None, attachments=None, attachments_dir=None):
    """
    Sends out an SMTP email using SSL, HTML content, and up to one
    attachment (including .zip). Recipients' names must have the form
    "required_first_name optional_middle_name optional_last_name". The
    sender's email is assumed to be Gmail/Google Inbox.
    :param sender: Sequence (a, b) where a is the sender's email and
    b is their email account password
    :param recipients: Sequence of pairs (a, b) where a is the
    recipient's name and b is their email
    :param cc: Sequence of pairs (a, b) where a is the cc
    recipient's name and b is their email
    :param bcc: Sequence of pairs (a, b) where a is the bcc
    recipient's name and b is their email
    :param subject: Subject title for the email
    :param attachments: File name of the attachment (including
    .zip) - no more than 1 per email
    :param html: File name of the html script defining the email
    body's content and signature
    :param attachments_dir: Directory containing the attachments
    :param html_dir: Directory containing the html script
    """

    # Construct formatted strings of names/emails for Message module
    recipient_names, cc_names, bcc_names = list(), list(), list()
    recipient_emails, cc_emails, bcc_emails = list(), list(), list()
    contact_lists = {'recipients': recipients, 'cc': cc, 'bcc': bcc}
    contact_names = {'recipients': recipient_names, 'cc': cc_names,
                     'bcc': bcc_names}
    contact_emails = {'recipients': recipient_emails, 'cc': cc_emails,
                      'bcc': bcc_emails}

    for group, contact_list in contact_lists.items():
        for contact in contact_list:
            contact_names[group].append(contact[0].split()[0])
            contact_emails[group].append(contact[1])
            contact_names[group] = ", ".join(contact_names[group])
            contact_emails[group] = "; ".join(contact_emails[group])

    # Extract HTML content for email body
    initial_dir = os.getcwd()
    os.chdir(html_dir)
    with open(html) as f:
        email_body = f.read()
    os.chdir(initial_dir)

    # Construct email
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = sender[0]
    msg['To'] = contact_emails['recipients']
    if not cc:
        msg['Cc'] = contact_emails['cc']
    if not bcc:
        msg['Bcc'] = contact_emails['bcc']
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
        """.format(contact_names[recipients], email_body),
                    subtype='html')
    if attachments is not None and attachments_dir is not None:
        # Prepare the attachment(s) for delivery
        initial_dir = os.getcwd()
        os.chdir(attachments_dir)
        if attachments[len(attachments) - 4:] == ".zip":
            with open(attachments, 'rb') as myzip:
                msg.add_attachment(myzip.read(), maintype="multipart",
                                   subtype="mixed", filename=attachments)
        else:
            with open(attachments, 'rb') as fp:
                msg.add_attachment(fp.read(), maintype="multipart",
                                   subtype="mixed", filename=attachments)
        os.chdir(initial_dir)

    # Connect with the server and send the email with its attachment(s)
    with smtplib.SMTP(host='smtp.gmail.com', port=587) as s:
        context = ssl.create_default_context()
        s.starttls(context=context)
        s.login(sender[0], sender[1])
        s.send_message(msg)

    return


def range_converter(xl_col_length=3):
    """
    Construct conversions between Excel array ranges and
    Pythonic indices (up to column ZZ in Excel)
    :param xl_col_length: Length of the longest desired
    Excel column (e.g., 2 for "A" to "ZZ", 3 for "A" to "ZZZ")
    """
    alpha_initial = string.ascii_uppercase
    alpha_extended = list(string.ascii_uppercase)

    if xl_col_length == 1:
        pass
    else:   # Expand list with same lexicographic ordering as
        # Excel (e.g. "Z" is followed by "AA", "AZ" by "BA")
        for k in range(2, xl_col_length + 1):
            new_sequences = list()
            for letter_sequence in alpha_extended:
                for new_letter in alpha_initial:
                    new_sequences.append("".join([letter_sequence,
                                                  new_letter]))
            alpha_extended.extend(new_sequences)
    convert = zip(range(1, len(alpha_extended) + 1), alpha_extended)
    convert_to_alpha = {x: y for x, y in convert}
    convert_to_num = {y: x for x, y in convert_to_alpha.items()}
    return convert_to_alpha, convert_to_num


class XlArray:
    """
    This class is meant for two-layer nested lists representing an
    Excel array: e.g., [[row_1], [row_2],...]
    """
    # Construct conversions between Excel array ranges and Pythonic indices
    converter = range_converter()
    convert_to_alpha = converter[0]
    convert_to_num = converter[1]

    def __init__(self, data, row, col):
        """
        :param data: Nested (or mono-layer) list representing an
        excel array (or row)
        :param row: Row location of the upper-left cell in the array
        (in Excel format, e.g., "2")
        :param col: Column location of the upper-left cell in the array
        (in Excel format - e.g., "B")
        """
        # If data is a mono-layer list (representing a row), convert it
        # into a nested list (representing an array)
        if not all(itertools.starmap(isinstance, zip(data,
                                                     [list] * len(data)))):
            data = [data]

        self.data = data
        self.col = col
        self.row = row
        self.len = len(data)  # Indicates the number of rows

        # Determine the finalized Excel array range
        self.empty = empty_check(data)
        if not self.empty:
            self.header = self.data[0]
            excel_range = (col + str(row) + ":" +
                           XlArray.convert_to_alpha[len(self.header) +
                           XlArray.convert_to_num[col] - 1] + str(self.len))
            # modified 5/24
            self.col_num = XlArray.convert_to_num[self.col]
            # XlArray.remove (below) may interfere with self.col_num
            self.last_col_num = self.col_num + len(self.header) - 1
            self.last_col = XlArray.convert_to_alpha[self.last_col_num]
            self.range = excel_range
            self.name = ""

    def empty(self, row_as_list):
        row_num = self.data.index(row_as_list)
        return empty_check(self.data[row_num])

    def remove(self, columns):
        """
        Removes the chosen columns in the instance's source array
        from the instance's own array with columns understood
        in Excel range terms.

        For instance, if the source array is [[a, b], [c,d]]
        with (row, col) = (2, "F"), the
        Excel interpretation is that the upper-left cell of the
        instance array is F2 while the range is F2:G3.
        Meanwhile, the instance's array's range is understood as
        [(i, j) for i, j in zip(range(2), range(2))].

        In the above case, self.remove(["G"]) would reduce the source
        array to [[a], [c]] as "b" and "d" represent cells
        G2 and G3, respectively.

        :param columns: Column (as string) or columns (as list of
        strings) in the source array in Excel's range
        interpretation - e.g., "A" for the 0th column
        """
        # Note that this section assumes no two rows/lists in the
        # data array are identical due to list.index()
        for excluded_col in columns:
            excluded_col_num = XlArray.convert_to_num[excluded_col]     # e.g., column "B" becomes 2
            if not self.empty and excluded_col_num == self.col_num:     # if the first column is to be excluded
                for record in self.data:
                    index = self.data.index(record)
                    self.data[index] = record[1:]                       # remove the first column in all rows
                self.col = XlArray.convert_to_alpha[self.col_num + 1]   # adjust the Excel representation attributes
                self.col_num = XlArray.convert_to_num[self.col]
            elif not self.empty and excluded_col_num == \
                    self.last_col_num:                                  # if the last column is to be excluded
                for record in self.data:
                    index = self.data.index(record)
                    self.data[index] = record[:-1]
            elif not self.empty and self.col_num < excluded_col_num \
                    < self.last_col_num:                                # if another column is to be excluded
                for record in self.data:
                    index = self.data.index(record)
                    self.data[index] = record[:excluded_col_num - 1] \
                                       + record[excluded_col_num:]      # Pythonic indexes!
            else:                                                       # if the column isn't in the instance array
                pass
        return

    def filter(self, column, value, strict=True):
        """
        :param column: The column that will be searched in
        the array
        :param value: The cell content that will be searched
        for in the array
        :param strict: If true, the filter requires exact
        equivalence.
        :return: Filtered copy of the array with only those
        rows containing the desired entry in the desired column
        """
        filtered_array = list()
        filter_row = ""
        for record in self.data:                            # Here, rows are represented by lists
            if record[column] == value:                     # Strict equivalency required for a match
                if not filter_row:                          # Determine upper-left range value for the filtered array
                    filter_row = (self.data.index(record)
                                  + self.row - 1)
                filtered_array.append(record)
            elif not strict:
                if not filter_row:                          # Determine upper-left range value for the filtered array
                    filter_row = (self.data.index(record)
                                  + self.row - 1)
                try:
                    # if record[column] and value are splittable,
                    # see if all components of the former are in the latter
                    entry = record[column].split()
                    if all(entry[i] in value.split() for
                           i in list(range(len(entry)))):
                        filtered_array.append(record)
                except TypeError:
                    pass

        return XlArray(filtered_array, filter_row, self.col)


class XlExtract:
    """
    Class Dependency: XlArray (for XlEdit.extract())

    Extract data from an existing Excel documents using
    the xlwings module.
    """

    def __init__(self, dir_path):
        hide_excel(True)
        self.path = dir_path
        self.name = os.path.split(dir_path)[1]
        self.date = mod_date(dir_path)
        self.wb = xw.Book(self.path)                                        # xw.books.open(self.path) returns error
        self.sheets = self.wb.sheets

    def open(self):
        hide_excel(True)
        return self.wb

    def close(self):
        try:
            hide_excel(False)
            self.wb.close()
        finally:
            return

    def init_sht(self, sheet_name, prior_sheet=None):
        """
        Create a new sheet in the workbook
        :param sheet_name: Desired name for the new sheet
        :param prior_sheet: Optional - the new sheet will
        be inserted after this sheet in the workbook
        """
        if prior_sheet is None:
            self.wb.sheets.add(sheet_name)
        else:
            self.wb.sheets.add(sheet_name, after=self.sheets)
        # create and name sheet
        pass

    def extract(self, exclude_sheets=None, exclude_cols=None,
                max_row=50000, max_col=100):
        """
        Imports all data in the workbook with each sheet represented
        by a different XlArray object
        :param exclude_sheets: List of the names of the sheets from
        which data won't be collected
        :param exclude_cols: List of pairs (a,b) where a is the sheet
        name and b lists the columns to be excluded
        :param max_row: Rows beyond this point will not be extracted
        :param max_col: Columns beyond this point will not be extracted
        :return: Pairs consisting of each sheet number and the array in
        that sheet with all empty rows removed.
        """
        wb_data = list()
        if exclude_sheets:
            sht_list = [sheet.name for sheet in self.sheets if sheet
                        not in exclude_sheets]
        else:
            sht_list = [sheet.name for sheet in self.sheets]
        for sht_name in sht_list:
            sht_xl = self.wb.sheets(sht_name)

            # Determine endpoints of the range to be extracted
            raw_data = sht_xl.range((1, 1), (max_row, max_col)).value
            col_len, row_len = list(), -1
            for row in raw_data:
                if empty_check(row):
                    row_len += 1
                    break
                else:
                    row_len += 1
                    j = -1
                    while j in list(range(-1, len(row))):
                        j += 1
                        if empty_check(row[j:]):
                            col_len.append(j)
                            break
            col_len = max(col_len)

            if col_len < max_col and row_len < max_row:
                last_cell_location = (XlArray.convert_to_alpha[col_len]
                                      + str(row_len))
            else:
                last_cell_location = (XlArray.convert_to_alpha[max_col]
                                      + str(max_row))
            sht_range = "A1:" + last_cell_location
            sht_data = sht_xl.range(sht_range).value
            sht_array = XlArray(sht_data, 1, "A")
            for row in sht_array.data:
                if empty_check(row):
                    sht_array.data.remove(row)
            try:
                for x_sheet, x_columns in exclude_cols:
                    if x_sheet == sht_name:
                        sht_array.remove(x_columns)
            except TypeError:                                                       # raised if no columns excluded
                pass
            wb_data.append((sht_xl.index - 1, sht_array))                           # sht.index is 1-based (as in Excel)

        self.close()
        return wb_data
        # create a range method here that opens a chosen sheet and
        # scans it for the first completely empty row & column


class XlCreate:
    """
        Class Dependency: XlArray

        Write XlArray objects to an Excel file with easily-customized
        formatting. Instantiating immediately opens a new
        Excel workbook, so consider instantiating within a "with" statement.
        (Otherwise, use XlCreate.close()) No extension is to be included
        in the filename.
    """
    def __init__(self, filename, dir_path):
        self.initial_dir = os.getcwd()
        os.chdir(dir_path)
        self.path = dir_path
        self.name = os.path.split(dir_path)[1]
        hide_excel(True)
        self.wb = xlsxwriter.Workbook(filename + ".xlsx")
        self.arrays = dict()
        self.header_bold = self.wb.add_format({'bold': True,
                                               'text_wrap': 1})             # Format object: Bold/wrap the header
        self.wrap = self.wb.add_format({'text_wrap': 1, 'align': 'top'})
        self.date_format = self.wb.add_format({'num_format': 'm/d/yy',
                                               'align': 'top'})             # Format object

    def close(self):
        self.wb.close()
        hide_excel(False)
        os.chdir(self.initial_dir)
        return

    def write(self, sheet_name, sheet_data, row=1, column="A",
              date_col=None, custom_width=None):
        """
        Adds a mapping between the new sheet name and its data
        to self.arrays. Writes the data to the new sheet.
        :param sheet_name: Name to be used for the new sheet
        :param sheet_data: Data to be mapped onto the new sheet
        starting with cell A1. Include the header.
        :param row: New sheet's row location of the upper-left
        cell in the array (in Excel format, e.g., "2")
        :param column: New sheet's column location of the
        upper-left cell in the array (in Excel format, e.g., "B")
        :param date_col: Columns (in Excel format) that are to be
        written as dates
        :param custom_width: Pairs (column, width) that determine
        column-specific width
        """

        # Construct conversions between Excel array ranges and Pythonic indices
        converter = range_converter()
        convert_to_alpha = converter[0]
        convert_to_num = converter[1]

        # Add mapping between new sheet name and its
        # data (translated into a XlArray object)
        self.arrays[sheet_name] = data = XlArray(sheet_data,
                                                 row, column)

        # Add a sheet with the chosen name and set the table name
        sht = self.wb.add_worksheet(sheet_name)
        table_name = "_".join(sheet_name.split())

        # Create list of table header formatting objects
        header_formatting = [{'header': col, 'header_format':
            self.header_bold} for col in data.header]
        # 5/23 This is running correctly

        # Insert the table and its data
        sht.add_table(data.range, {'columns': header_formatting,
                                   'name': table_name})
        for item in data.header:
            sht.write(0, data.col_num + data.header.index(item)
                      - 1, item, self.header_bold)
            # 5/24: added -1 above, verify that this works
        all_columns_xl = list()
        # represents the destination columns
        all_columns_py = dict()
        # represents the indexes of these columns in the source data.data
        for k in range(data.col_num, data.last_col_num + 1):                        # Added "+1" on 5/23 - check!
            all_columns_xl.append(convert_to_alpha[k])
        for col in all_columns_xl:
            all_columns_py[col] = convert_to_num[col] -\
                                  convert_to_num[all_columns_xl[0]]
        for row_py in range(1, data.len):
            for col in all_columns_xl:
                col_py = all_columns_py[col]
                if date_col and col in date_col:
                    if not isinstance(data.data[row_py][col_py],
                                      datetime.datetime):
                        sht.write(row_py, col_py, "NO DATE",
                                  self.date_format)                                 # sht.write() uses 0-base indexes
                    else:
                        sht.write_datetime(row_py, col_py,
                                           data.data[row_py][col_py],
                                           self.date_format)
                else:
                    sht.write(row_py, col_py,
                              data.data[row_py][col_py], self.wrap)

        # Adjust the column widths
        for col in all_columns_xl:
            if not custom_width or col not in custom_width:
                col_py = all_columns_py[col]
                len_lst = [len(str(record[col_py])) for
                           record in data.data[1:]]
                if not len_lst:
                    max_len = 16
                elif max(len_lst) > 50:
                    max_len = 50
                else:
                    max_len = max(max(len_lst), 16)
                sht.set_column(col + ":" + col, max_len)
            elif custom_width:
                custom_dict = {x: y for x, y in custom_width}
                sht.set_column(col + ":" + col, custom_dict[col])
            else:
                pass

        return
