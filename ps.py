#
import os
import shutil
import datetime
import openpyxl
import logging
from validate_email import validate_email
from phonenumbers import parse, NumberParseException

CURRENT_YEAR = str(datetime.datetime.now().year)   # This year
PROJECT_ROOT = './Test'                        # Where we look (P drive)
CAD_SOURCE = os.path.join(PROJECT_ROOT, 'CAD_Template')
REVIT_SOURCE = os.path.join(PROJECT_ROOT, 'Revit_Template')
INFO_FILE = 'Project_Information.xlsx'

logger = logging.getLogger('__name__')
stream_handler = logging.StreamHandler()
logger.addHandler(stream_handler)


class Contact():
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
        self.type = None
        self.scope = None
        self.exists = None
        self.manager = None
        self.billing = Contact()
        self.contact = Contact()
        self.is_live = False  # Set to true when matched to an actual project
        self.workbook = None  # Store the openpyxl workbook handle
        self.mode = None
        self.fileError = None  # Find out why a file can't be locked
        # List of all folders in PROJECT_ROOT
        self.folder_list = os.listdir(PROJECT_ROOT)

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
        folder_list = os.listdir(PROJECT_ROOT)
        return [x for x in folder_list
                if x.upper().count(search_str.upper()) > 0]

    def next_project_number(self):
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
        pfolder = [x for x in FOLDER_LIST if x[: length] == ns]
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
                logger.error("Unable to create project path")
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
        self.number = self.next_project_number()
        return self.get_full_number()

    def load_project(self):
        """
        Load project information from the folder and spreadsheet.
        """
        filename = os.path.join(self.path_to_project(), INFO_FILE)
        logger.info(f"Opening file{filename}")
        if self.lockable(filename):
            # We should lock the file
            try:
                logger.debug(f"openpyxl is attempting to open {filename}")
                self.workbook = openpyxl.load_workbook(filename)
                sheet = self.workbook.active
                self.manager = sheet['C6'].value
                self.type = sheet['C8'].value
                self.scope = sheet['C10'].value
                self.contact.name = sheet['C12'].value
                self.contact.title = sheet['C13'].value
                self.contact.phone = sheet['C14'].value
                self.contact.email = sheet['C15'].value
                self.contact.address = sheet['C17'].value
                self.contact.csz = sheet['C18'].value
                self.billing.name = sheet['C20'].value
                self.billing.title = sheet['C21'].value
                self.billing.phone = sheet['C22'].value
                self.billing.email = sheet['C23'].value
                self.billing.address = sheet['C25'].value
                self.billing.csz = sheet['C26'].value
                # Keep it open for writing changes
                # pf.close()
                return True
            except Exception as e:
                logger.error(f"Project spreadsheet open failed: {filename}")
                logger.error(f"Error was: {e}")
                self.fileError = str(e)
                return False
        else:
            logger.error(f"Unable to lock information file: {self.fileError}.")
            return False

    def validate(self):
        """
        Check the data in the record before writing changes.
        """
        if self.mode == 'Create':
            if self.year != CURRENT_YEAR:
                logger.error("Should reference current year")
                return False, "Should reference current year"
            if os.exists(self.path_to_project()):
                logger.error("Project exists")
                return False, "Project exists, cannot create"
            if self.name_from_number():
                logger.error("Project number is already in use")
                return False, "Project number is already in use"
        # These tests apply to all
        if len(self.name) < 3:
            logger.error("Project name too short")
            return False, "Project name too short"
        if len(self.manager) < 3:
            logger.error("Project Manager name too short")
            return False, "Project Manager name too short"
        if len(self.scope) < 10:
            logger.error("Provide scope description")
            return False, "Provide scope description"
        if self.contact.phone != "":
            try:
                p = parse(self.contact.phone)
            except NumberParseException:
                p = parse(self.contact.phone, 'US')
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

    def update_project(self):
        """
        Write out changes to the spreadsheet in an existing project.
        Assumes data is in the record and has been validated.
        """
        logger.debug("Called update_project")
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

    def close_project(self):
        """
        Close out the worksheet before exiting.
        """
        self.workbook.close()


# def getProjectData(app):
#     """
#     Get information for a given project and instantiate in the GUI.
#     Return False if there's no such project.
#     """
#     pnum = app.project_number.get()
#     length = len(pnum)
#     pfolder = [x for x in FOLDER_LIST if x[: length] == pnum]
#     if (len(pfolder) == 1 and
#             os.path.isdir(os.path.join(PROJECT_ROOT, pfolder[0]))):
#             # We found the folder
#         pname = pfolder[0][9:]
#         pname = pname.lstrip(' -')
#         self.project_name = pname
#         self.load_project()
#         return (pfolder[0], pnum, pname)
#     else:
#         logger.error(f"getProjectData: No such project: {pnum}")
#         return False
#
#
# def makeProjectFolder(app):
#     """
#     Create the project name and create the folder for it.
#     Assumes data has been validated.
#     """
#     pass
#
#
# def setMode(app, mode):
#     """
#     Sets the mode to 'create' or 'modify' based on button press.
#     """
#     if mode == 'create':
#         app.next_pnum = f"{CURRENT_YEAR}.{getProjectNumber()}"
#         app.project_number.set(app.next_pnum)
#         app.project_number.state = 'disabled'
#         app.mode_label.set("// New project - enter name, type, and PM. //")
#         app.createButton.config(highlightbackground='#FF6600')
#         app.updateButton.config(highlightbackground='#AAAAAA')
#     elif mode == 'modify':
#         getProjectData(app)
#         app.mode_label.set("// Update project - revise information. //")
#         app.createButton.config(highlightbackground='#AAAAAA')
#         app.updateButton.config(highlightbackground='#FF6600')
#     else:
#         error('setMode: Invalid mode: {}'.format(mode))
#         return
#     app.mode = mode


def createProject(app):
    """
    Create a new project using pre-validated information.
    """
    global FOLDER_LIST  # We'll update it if successful with the new folder.

    if not app.mode == 'create':
        error('createProject: Mode not set to create')
        return  # We shouldn't have been called

    print(("Creating Project '{} {}'"
           " for {}.").format(app.project_number.get(),
                              app.project_name.get(),
                              app.project_pm.get()))
    new_folder = buildPath(app)

    project_type = app.project_type.get()
    print("Project type is: {}".format(project_type))
    print("Folder is: {}".format(new_folder))
    try:
        if project_type == 'CAD':
            shutil.copytree(CAD_SOURCE, new_folder)
        elif project_type == 'Revit':
            shutil.copytree(REVIT_SOURCE, new_folder)
        else:
            # 02 Folder only
            pass
    except OSError as e:
        error("copyFiles: {}".format(e))
    FOLDER_LIST = os.listdir(PROJECT_ROOT)


def modifyProject(app):
    """
    Alter project information for an existing project.
    """
    # check basic data is in place
    # write updated information to the spreadsheet
    print("Modify project {}".format(app.project_number.get()))
    fileName = buildPath(app, INFO_FILE)
    try:
        print('Opening file: {}'.format(fileName))
        pf = openpyxl.load_workbook(fileName)
        sheet = pf.active
        # app.project_name.set(sheet['B4'].value)
        sheet['B6'].value = app.project_pm.get()
        sheet['B7'].value = app.project_type.get()
        sheet['B9'].value = app.project_addr.get()
        sheet['B10'].value = app.project_csz.get()
        sheet['B12'].value = app.billing_name.get()
        sheet['B17'].value = app.billing_addr.get()
        sheet['B18'].value = app.billing_csz.get()
        pf.close()
    except FileNotFoundError:
        error("readProjectData: "
              "Project spreadsheet not found: {}".format(fileName))
    app.mode_label.set("// Project created. //")


def checkNewProject(app):
    """
    Check project name and PM data prior to creating a project.
    Project number was set when create mode was invoked.
    """
    pname = app.project_name.get()
    tr_table = str.maketrans('', '', ',;:"\'\\`~!%^#&{}|<>?*/')
    clean_name = pname.translate(tr_table)
    clean_name = clean_name.strip('_ .\t\n-')
    if clean_name != pname:
        error("Name cannot contain special characters")
        return False
    if len(clean_name) < 6:
        error("Project name too short.")
        return False
    app.project_pm.set(app.project_pm.get().strip(' \t\n-'))
    if len(app.project_pm.get()) < 3:
        error("Enter a valid project manager name")
        return False
    if not app.project_type.get() in ['Revit', 'CAD', 'Other']:
        error("Select a project type (CAD, Revit or Other)")
        return False
    app.mode_label.set("// Mode: CREATE - Ready to GO. //")
    return True


def write(app):
    """
    Writes to file system based on the current mode.
    """
    if app.mode == 'create':
        if checkNewProject(app):
            createProject(app)
    elif app.mode == 'modify':
        modifyProject(app)
    else:
        error("Invalid mode: {}".format(app.mode))
