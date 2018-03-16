#
import os
import shutil
import datetime
import re
import openpyxl
import logging
from validate_email import validate_email
from phonenumbers import parse, NumberParseException

CURRENT_YEAR = str(datetime.datetime.now().year)   # This year
PROJECT_ROOT = './Test'                        # Where we look (P drive)
CAD_SOURCE = os.path.join(PROJECT_ROOT,
                          'Project_Templates', 'CAD_Template')
REVIT_SOURCE = os.path.join(PROJECT_ROOT,
                            'Project_Templates', 'Revit_Template')
DEFAULT_SOURCE = os.path.join(PROJECT_ROOT,
                              'Project_Templates', 'Default_Template')
INFO_FILE = 'Project_Information.xlsx'
BAD_CHARS = ',;:"\'\\~!$%^=!#&{}[]|<>?*/\t\n'

logger = logging.getLogger('__name__')
stream_handler = logging.StreamHandler()
logger.addHandler(stream_handler)


class Person():
    """
    Information on a person.
    """

    def __init__(self, name=None, title=None,
                 address=None, csz=None, phone=None, email=None):
        self.name = name
        self.title = title
        self.address = address
        self.csz = csz
        self.phone = phone
        self.email = email


class Project():
    """
    Internal representation of a Studio G Project, before and after creation.
    """

    def __init__(self):
        """
        Create a blank project template.
        """
        self.year = None
        self.number = None
        self.name = None
        self.folder = None
        self.type = None
        self.scope = None
        self.exists = None
        self.manager = None
        self.billing = Person()
        self.contact = Person()
        self.is_live = False  # Set to true when matched to an actual project
        self.workbook = None  # Store the openpyxl workbook handle
        self.mode = None
        self.fileError = None  # Find out why a file can't be locked

    def from_folder(self, folder_name):
        """
        Set project number and name from a folder name.
        """
        self.folder = folder_name
        num, *name = folder_name.split(' ')
        self.year = num.split('.')[0]
        self.number = num.split('.')[1]
        self.name = ' '.join(name)
        return self.year, self.number, self.name

    def get_full_number(self):
        """Return the combination of year and number as a string"""
        if self.year and self.number:
            return f"{self.year}.{self.number:2}"
        else:
            logger.error("Project has no number")
            return ""

    def get_full_name(self):
        if self.folder:
            return self.folder
        if self.year and self.number and self.name:
            fn = self.get_full_number()
            self.folder = f"{fn} {self.name}"
            return self.folder
        else:
            logger.error("Project has no name")
            return ""

    def match_folder(self, search_str):
        """
        Return a list of folders that contain the search string.
        """
        pat = re.compile(r'\d\d\d\d\.\d\d')
        folder_list = os.listdir(PROJECT_ROOT)
        folder_list = [x for x in folder_list
                       if x.upper().count(search_str.upper()) > 0]
        folder_list = [x for x in folder_list if pat.match(x)]
        return folder_list

    def next_number(self):
        """
        Get the number of the most recent project,
        and return the next number up as a 2 or 3 digit string..
        """
        pnum = str(datetime.datetime.now().year)
        folder_list = os.listdir(PROJECT_ROOT)
        fileList = [f for f in folder_list if f[:4] == pnum]
        try:
            latestProject = sorted(fileList, reverse=True)[0]
            projectNumber = int(latestProject[5: 8])
            return "{:02}".format(projectNumber + 1)
        except IndexError:
            logger.error("No projects found.")
            return '000'

    def name_from_number(self, num_string=None):
        """
        Find project name from a provided number or number in the record.
        """
        if num_string:
            ns = num_string
        elif self.year and self.number:
            ns = self.get_full_number()
        else:
            logger.error("No number for name search")
            return ""
        length = len(ns)
        folder_list = os.listdir(PROJECT_ROOT)
        pfolder = [x for x in folder_list if x[: length] == ns]
        if (len(pfolder) == 1 and
                os.path.isdir(os.path.join(PROJECT_ROOT, pfolder[0]))):
                # We found the folder
            pname = pfolder[0][8:]
            pname = pname.lstrip(' -')
            return pname
        else:
            logger.info(f"Project {ns} not found")
            return False

    def path_to_project(self):
        """
        Develop full path to the current project.
        """
        if self.year and self.number:
            if self.name:
                name = self.get_full_name()
            else:
                name = self.name_from_number()
            if name:
                return os.path.join(PROJECT_ROOT, name)
            else:
                msg = "Check project number and name"
                logger.error(msg)
                self.fileError = msg
                return False

    def lockable(self, name):
        """
        Determine if the file is present and can be locked. No side effects.
        """
        logger.debug(f'Checking lock on file: {name}')
        if not os.path.isfile(name):
            self.fileError = "File does not exist."
            logger.error("lockable: File does not exist")
            return False
        # Check for an Excel lock file
        dir, file = os.path.split(name)
        lock_file = f"~${file}"
        lock_path = os.path.join(dir, lock_file)
        if os.path.isfile(lock_path):
            self.fileError = "File is locked by Excel"
            return False
        else:
            return True

    def start_new(self):
        """
        Initial setup for a new project. Return next project number.
        """
        self.mode = "Create"
        self.year = CURRENT_YEAR
        self.number = self.next_number()
        return self.get_full_number()

    def open(self):
        """
        Open and lock the Project Information spreadsheet.
        """
        filename = os.path.join(self.path_to_project(), INFO_FILE)
        logger.info(f"Opening file{filename}")
        if self.lockable(filename):
            # TODO: We should lock the file
            try:
                logger.debug(f"openpyxl is attempting to open {filename}")
                self.workbook = openpyxl.load_workbook(filename,
                                                       data_only=True)
            except Exception as e:
                logger.error(f"Project spreadsheet open failed: {filename}")
                logger.error(f"Error was: {e}")
                self.fileError = str(e)
                return False
            return True
        else:
            logger.error(f"Unable to lock information file: {self.fileError}.")
            return False

    def load(self):
        """
        Load project information from the folder and spreadsheet.
        File is assumed to be open already.
        """
        def get(sheet, col, row):
            """
            Gets value from sheet cell, replacing None with ""
            """
            v = sheet.cell(row=row, column=col).value
            return (v if v else "")

        try:
            logger.debug("Loading from spreadsheet")
            sheet = self.workbook.active
            self.manager = get(sheet, 3, 6)
            self.type = get(sheet, 3, 8)
            self.scope = get(sheet, 3, 10)
            self.contact.name = get(sheet, 3, 12)
            self.contact.title = get(sheet, 3, 13)
            self.contact.phone = get(sheet, 3, 14)
            self.contact.email = get(sheet, 3, 15)
            self.contact.address = get(sheet, 3, 17)
            self.contact.csz = get(sheet, 3, 18)
            self.billing.name = get(sheet, 3, 20)
            self.billing.title = get(sheet, 3, 21)
            self.billing.phone = get(sheet, 3, 22)
            self.billing.email = get(sheet, 3, 23)
            self.billing.address = get(sheet, 3, 25)
            self.billing.csz = get(sheet, 3, 26)
            return True
        except Exception as e:
            logger.error(f"Project spreadsheet load failed")
            logger.error(f"Error was: {e}")
            self.fileError = str(e)
            return False

    def clean(self, data):
        """
        Sanitizes strings before writing to the database.
        """
        cl = data.translate(None, BAD_CHARS)
        if cl == data:
            return True, cl
        else:
            return False, cl

    def validate(self):
        """
        Check the data in the record before writing changes.
        """
        if self.mode == 'Create':
            if self.year != CURRENT_YEAR:
                logger.error("Should reference current year")
                return False, "Should reference current year"
            path = self.path_to_project()
            if not path:
                return False, self.fileError
            if os.path.exists(self.path_to_project()):
                logger.error(f"Project exists {self.path_to_project()}")
                return False, "Project exists, cannot create"
            if self.name_from_number():
                logger.error("Project number is already in use")
                return False, "Project number is already in use"
        # These tests apply to all
        if len(self.name) < 3:
            logger.error("Project name too short")
            return False, "Project name too short"
        if len(self.manager) < 3:
            logger.error("Check Project Manager name")
            return False, "Check Project Manager name"
        if len(self.scope) < 10:
            logger.error("Provide scope description")
            return False, "Provide scope description"
        if self.contact.phone != "":
            try:
                parse(self.contact.phone)
            except NumberParseException:
                try:
                    parse(self.contact.phone, 'US')
                except NumberParseException:
                    msg = f"Invalid phone number: {self.contact.phone}"
                    logger.error(msg)
                    return False, msg
        if self.billing.phone != "":
            try:
                parse(self.billing.phone)
            except NumberParseException:
                try:
                    parse(self.billing.phone, 'US')
                except NumberParseException:
                    msg = f"Invalid phone number: {self.billing.phone}"
                    logger.error(msg)
                    return False, msg
        if self.contact.email != "":
            if validate_email(self.contact.email):
                pass
            else:
                logger.error("Invalid project contact email address")
                return False, "Invalid project contact email address"
        if self.billing.email != "":
            if validate_email(self.billing.email):
                pass
            else:
                logger.error("Invalid billing contact email address")
                return False, "Invalid billing contact email address"

        return True, ""

    def update(self):
        """
        Write out changes to the spreadsheet in an existing project.
        Assumes data is in the record and has been validated.
        """
        if self.mode == "Edit":
            logger.debug("Called update")
            sheet = self.workbook.active
            sheet['C6'].value = self.manager
            sheet['C8'].value = self.type
            sheet['C10'].value = self.scope
            sheet['C12'].value = self.contact.name
            sheet['C13'].value = self.contact.title
            sheet['C14'].value = self.contact.phone
            sheet['C15'].value = self.contact.email
            sheet['C17'].value = self.contact.address
            sheet['C18'].value = self.contact.csz
            sheet['C20'].value = self.billing.name
            sheet['C21'].value = self.billing.title
            sheet['C22'].value = self.billing.phone
            sheet['C23'].value = self.billing.email
            sheet['C25'].value = self.billing.address
            sheet['C26'].value = self.billing.csz
            filename = os.path.join(self.path_to_project(), INFO_FILE)
            self.workbook.save(filename)
            logger.debug("File written")
            return True
        else:
            logger.error("update called in wrong mode.")
            return False

    def close(self):
        """
        Close out the worksheet if it's open before exiting or creating new.
        """
        if self.workbook:
            self.workbook.close()
        # TODO: unlock the file

    def create(self):
        """
        Create a new project based on data provided in the project record.
        """
        if self.mode != "Create":
            logger.error("Should be called only in Create mode")
            return
        logger.info(f"Creating new project {self.number}")
        # Create the folder after checking it doesn't exist
        new_folder = self.path_to_project()
        if os.path.exists(new_folder):
            logger.error("Project folder exists - can't create.")
            return
        else:
            try:    # Copy from the relevant template to the new folder
                if self.type == 'CAD':
                    shutil.copytree(CAD_SOURCE, new_folder)
                elif self.type == 'Revit':
                    shutil.copytree(REVIT_SOURCE, new_folder)
                else:  # Default to default type
                    shutil.copytree(DEFAULT_SOURCE, new_folder)
            except OSError as e:
                logger.error(f"Failed to copy: {e}")
                return
        self.mode = "Edit"  # We have created the folder tree, now can edit.
        if self.open():
            if self.update():
                logger.info("Project information written successfully")
                return True
            else:
                logger.error("Unable to write project information")
                return False
        else:
            logger.error("Unable to open project information file")
