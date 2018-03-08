# Read excel file and build project appropriately
import sys
import os


# import openpyxl

PROJECT_ROOT = '.'
CURRENT_YEAR = '2018'


def create(values):
    pass


def modify(values):
    pass


def check():
    if ps.project_data['year'] != int(CURRENT_year):
        return False


def getProjectNumber():
    """
    Get the number of the most recent project,
    and return the next number up as a 2 or 3 digit string..
    """
    fileList = filter(lambda x: x[:4] == CURRENT_YEAR,
                      os.listdir(PROJECT_ROOT))
    try:
        latestProject = sorted(fileList, reverse=True)[0]
        projectNumber = int(latestProject[5:8])
        return (CURRENT_YEAR, projectNumber + 1)
    except IndexError:
        print("No projects found.")
        return ('0000', '000')


def makeProjectFolder(name):
    """
    Create the project name and create the folder for it.
    """
    folderName = CURRENT_YEAR + '.' + getProjectNumber() + ' - ' + name
    os.mkdir(folderName)
