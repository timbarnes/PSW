Project Initialization and Update Application

Based on Python3 and tkinter

Two use modes:
    1. Project creation
    2. Project update

Project Creation

- Open the app, and select Create mode
- GUI updates to show the next available project number
- enter a project name and type, and project manager
- Optionally provide project address, billing contact and billing address
- Click "GO"
- Software checks the information for errors
- Software creates the project folder, copies in the files, makes a contact file, updates filenames, and emails Rick

Project Update

- Open the app, and select Modify modify mode
- Enter the project number of an existing project
- Tab to next field? Or hit Enter?
- Software loads known information from the project folder
- User edits the information
- Click "Update"
- Software writes the changes.

Data structures

Tkinter requires stringVars for all the GUI components, so they will be used
as the main storage mechanism.
- They are created by the Application __init__().
- As the Application is global, it's available to the helper functions.
- Create, Modify/Update, and GO operate on the stringVars, so the GUI
  is constantly kept up to date.

TASKS

* Set Create mode at startup and check for next project number
* Check / update next project number every time create mode is selected
* Reset / cleanup after GO action is completed
* Better visible indication of current mode
* Replace popups with a message area
* Integrate project data spreadsheet to create and update
