#
import os
import shutil
import datetime
import openpyxl
import logging

CURRENT_YEAR = str(datetime.datetime.now().year)   # This year
PROJECT_ROOT = './Test'                        # Where we look (P drive)
CAD_SOURCE = os.path.join(PROJECT_ROOT, 'CAD_Template')
REVIT_SOURCE = os.path.join(PROJECT_ROOT, 'Revit_Template')
INFO_FILE = 'Project_Information.xlsx'

logger = logging.getLogger('__name__')
stream_handler = logging.StreamHandler()
logger.addHandler(stream_handler)


class Address():
    """
    Location information for a project or client.
    """

    def __init__(self, addr1=None, addr2=None, csz=None):
        self.addr1 = addr1
        self.addr2 = addr2
        self.csz = csz


class Contact():
    """
    Information on a person.
    """

    def __init__(self, name=None, title=None,
                 address=None, phone=None, email=None):
        self.name = name
        self.title = title
        self.address = address
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
        self.project_type = None
        self.exists = None
        self.project_manager = None
        self.billing_contact = Contact()
        self.project_contact = Contact()
        self.is_live = False  # Set to true when matched to an actual project
        # List of all folders in PROJECT_ROOT
        self.folder_list = os.listdir(PROJECT_ROOT)

    def get_full_number(self):
        """Return the combination of year and number as a string"""
        if self.year and self.number:
            return f"{self.year}.{self.number:2}"
        else:
            logger.error("Project has no number")
            return ""

    def get_full_name(self):
        if self.year and self.number and self.name:
            fn = self.get_full_number()
            return f"{fn} {self.name}"
        else:
            logger.error("Project has no name")
            return ""

    def next_project_number(self):
        """
        Get the number of the most recent project,
        and return the next number up as a 2 or 3 digit string..
        """
        pnum = str(datetime.datetime.now().year)
        fileList = [f for f in self.folder_list if f[:4] == pnum]
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
            logger.error(f"Project {ns} not found")
            return False

    def path_to_project(self):
        """
        Develop full path to the current project.
        """
        if self.year and self.number:
            if self.name:
                file_name = self.get_full_name()
            else:
                name = self.name_from_number()
            if name:
                p = os.path.join(PROJECT_ROOT,
                                 self.full_number + self.name)
                return p

    def load_project(self):
        """
        Load project information from the folder and spreadsheet.
        """
        filename = self.path_to_project() + '/' + INFO_FILE
        logger.info(f"Opening file{filename}")
        try:
            pf = openpyxl.load_workbook(filename)
            sheet = pf.active
            self.project_manager = sheet['B6'].value
            self.project_type = sheet['B7'].value
            self.project_contact.address.addr1 = sheet['B9'].value
            self.project_contact.address.csz = sheet['B10'].value
            self.billing_contact.name = sheet['B12'].value
            # app.billing_title.set(sheet['B13'].value)
            self.billing_contact.address.addr1 = sheet['B17'].value
            self.billing_contact.address.csz = sheet['B18'].value
            pf.close()
        except FileNotFoundError:
            logger.error(f"Project spreadsheet not found: {fileName}")
            return False


def getProjectData(app):
    """
    Get information for a given project and instantiate in the GUI.
    Return False if there's no such project.
    """
    pnum = app.project_number.get()
    length = len(pnum)
    pfolder = [x for x in FOLDER_LIST if x[: length] == pnum]
    if (len(pfolder) == 1 and
            os.path.isdir(os.path.join(PROJECT_ROOT, pfolder[0]))):
            # We found the folder
        pname = pfolder[0][9:]
        pname = pname.lstrip(' -')
        self.project_name = pname
        self.load_project()
        return (pfolder[0], pnum, pname)
    else:
        logger.error(f"getProjectData: No such project: {pnum}")
        return False


def makeProjectFolder(app):
    """
    Create the project name and create the folder for it.
    Assumes data has been validated.
    """
    pass


def setMode(app, mode):
    """
    Sets the mode to 'create' or 'modify' based on button press.
    """
    if mode == 'create':
        app.next_pnum = f"{CURRENT_YEAR}.{getProjectNumber()}"
        app.project_number.set(app.next_pnum)
        app.project_number.state = 'disabled'
        app.mode_label.set("// New project - enter name, type, and PM. //")
        app.createButton.config(highlightbackground='#FF6600')
        app.updateButton.config(highlightbackground='#AAAAAA')
    elif mode == 'modify':
        getProjectData(app)
        app.mode_label.set("// Update project - revise information. //")
        app.createButton.config(highlightbackground='#AAAAAA')
        app.updateButton.config(highlightbackground='#FF6600')
    else:
        error('setMode: Invalid mode: {}'.format(mode))
        return
    app.mode = mode


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
