#!/usr/bin/python

import sys
import os
import openpyxl
import pprint

CURRENT_YEAR = '2018'
PROJECT_ROOT = '.'

# Read excel file and build project appropriately


def getProjectNumber():
    """
    Get the number of the most recent project,
    and return the next number up as a 2 or 3 digit string..
    """
    fileList = filter(lambda x: x[:4] == CURRENT_YEAR,
                      os.listdir(PROJECT_ROOT))
    latestProject = sorted(fileList, reverse=True)[0]
    projectNumber = int(latestProject[5:8])
    if projectNumber < 9:
        return '0{:d}'.format(projectNumber+1)
    else:
        return '{:d}'.format(projectNumber+1)


def makeProjectFolder(name):
    """
    Create the project name and create the folder for it.
    """
    folderName = CURRENT_YEAR+'.'+getProjectNumber()+' - '+name
    os.mkdir(folderName)


def getProjectData(fileName):
    """
    Read the data from Excel and return as a dictionary
    """
    print('Opening file: %s', fileName)
    pf = openpyxl.load_workbook(fileName)
    sheet = pf.active
    dict = {'ProjectName': sheet['B3'].value,
            'ProjectType': sheet['B5'].value,
            'ProjectStreet': sheet['B7'].value,
            'ProjectCSZ': sheet['B8'].value,
            'BillingName': sheet['B10'].value,
            'BillingTitle': sheet['B11'].value,
            'BillingPhone': sheet['B12'].value,
            'BillingEmail': sheet['B13'].value,
            'BillingStreet': sheet['B15'].value,
            'BillingCSZ': sheet['B16'].value}
    return dict


pp = pprint.PrettyPrinter(depth=4)

pp.pprint(getProjectData(sys.argv[2]))
