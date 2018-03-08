import sys
import logging
import wx
import ps

logger = ps.logger


class Application(wx.Frame):
    """
    Build the application window and initialize a project
    """
    # Initialize an empty project

    def __init__(self, master=None):
        # Create the main frame
        wx.Frame.__init__(self, None, wx.ID_ANY,
                          'Studio G Project Editor', size=(550, 665))
        # Create the communicating variables
        self.mode = 'Edit'
        self.project = ps.Project()

        self.build_GUI()

    def build_GUI(self):
        """
        Create all the elements of the UI and connect to variables.
        """

        def contact_line(row, label, callback):
            """
            Make a line with StaticText and TextCtrl and callback.
            """
            prompt = wx.StaticText(self.panel, label=label)
            self.sizer.Add(prompt, pos=(row, 0), flag=wx.ALL, border=4)
            t = wx.TextCtrl(self.panel)
            t.Bind(wx.EVT_KEY_DOWN, callback)
            self.sizer.Add(t, pos=(row, 1), span=(0, 3),
                           flag=wx.EXPAND | wx.ALL, border=4)
            return t

        # Create the main panel that goes in the frame
        self.panel = wx.Panel(self, wx.ID_ANY)
        self.panel.SetAutoLayout(True)
        # Create a statusbar
        self.sb = self.CreateStatusBar(2)
        self.sb.SetStatusText("Studio G Project Editor", 0)
        # # Create the sizer
        self.sizer = wx.GridBagSizer(4, 4)
        row = 0
        label = wx.StaticText(self.panel, label="Mode:")
        self.sizer.Add(label, pos=(row, 0), flag=wx.ALL, border=4)
        self.newButton = wx.Button(self.panel, label="Create New Project")
        self.newButton.Bind(wx.EVT_BUTTON, self.on_new)
        self.sizer.Add(self.newButton, pos=(row, 1), flag=wx.ALL, border=5)
        self.editButton = wx.Button(self.panel, label="Edit Project Info.")
        self.editButton.Bind(wx.EVT_BUTTON, self.on_edit)
        self.sizer.Add(self.editButton, pos=(row, 2), flag=wx.ALL, border=5)
        self.updateButton = wx.Button(self.panel, label="GO")
        self.updateButton.Bind(wx.EVT_BUTTON, self.on_go)
        self.sizer.Add(self.updateButton, pos=(row, 3), flag=wx.ALL, border=5)
        row += 1
        self.sizer.Add(wx.StaticLine(self.panel), pos=(row, 0), span=(1, 4))
        row += 1
        label = wx.StaticText(self.panel, label="Search:")
        self.sizer.Add(label, pos=(row, 0), flag=wx.ALL, border=4)
        self.searchString = wx.TextCtrl(self.panel)
        self.searchString.Bind(wx.EVT_KEY_DOWN, self.on_search)
        self.sizer.Add(self.searchString, pos=(row, 1),  span=(0, 2),
                       flag=wx.EXPAND | wx.ALL, border=5)
        self.searchButton = wx.Button(self.panel, label="Find Project")
        self.searchButton.Bind(wx.EVT_BUTTON, self.on_search)
        self.sizer.Add(self.searchButton, pos=(row, 3), flag=wx.ALL, border=5)
        row += 1
        label = wx.StaticText(self.panel, label="Project number && name:")
        self.sizer.Add(label, pos=(row, 0), flag=wx.ALL, border=4)
        self.projectNumber = wx.TextCtrl(self.panel)
        self.projectNumber.Bind(wx.EVT_KEY_DOWN, self.on_number)
        self.sizer.Add(self.projectNumber, pos=(row, 1),
                       flag=wx.ALL | wx.EXPAND, border=5)
        self.projectName = wx.TextCtrl(self.panel)
        self.projectName.Bind(wx.EVT_KEY_DOWN, self.on_project_name)
        self.sizer.Add(self.projectName, pos=(row, 2),  span=(0, 3),
                       flag=wx.EXPAND | wx.ALL, border=5)
        row += 1
        self.billingName = contact_line(row, "Project Manager:",
                                        self.on_project_manager)
        row += 1
        label = wx.StaticText(self.panel, label="Project Type:")
        self.sizer.Add(label, pos=(row, 0), flag=wx.ALL, border=4)
        type_list = ['None', 'Revit', 'CAD']
        self.projectType = wx.RadioBox(self.panel, choices=type_list,
                                       majorDimension=1,
                                       style=wx.RA_SPECIFY_ROWS)
        self.projectType.Bind(wx.EVT_RADIOBUTTON, self.on_project_type)
        self.sizer.Add(self.projectType, pos=(row, 1), span=(0, 4),
                       flag=wx.ALL, border=5)
        row += 1
        self.sizer.Add(wx.StaticLine(self.panel), pos=(row, 0), span=(1, 4))
        row += 1
        label = wx.StaticText(self.panel, label="PROJECT INFORMATION")
        self.sizer.Add(label, pos=(row, 1), span=(0, 3), flag=wx.EXPAND)
        row += 1
        self.contactName = contact_line(row, "Project Contact:",
                                             self.on_contact_name)
        row += 1
        self.contactTitle = contact_line(row, "  Contact Title:",
                                         self.on_contact_title)
        row += 1
        self.contactAddress = contact_line(row, "  Project Address:",
                                                self.on_contact_address)
        row += 1
        self.contactCSZ = contact_line(row, "  City / State / Zip:",
                                       self.on_contact_csz)
        row += 1
        self.contactPhone = contact_line(row, "  Contact phone",
                                         self.on_contact_phone)
        row += 1
        self.contactEmail = contact_line(row, "  Contact email",
                                         self.on_contact_email)
        row += 1
        self.sizer.Add(wx.StaticLine(self.panel), pos=(row, 0), span=(1, 4))
        row += 1
        label = wx.StaticText(self.panel, label="BILLING INFORMATION")
        self.sizer.Add(label, pos=(row, 1), span=(0, 3), flag=wx.EXPAND)
        row += 1
        self.billingName = contact_line(row, "Billing Name:",
                                        self.on_billing_name)
        row += 1
        self.billingTitle = contact_line(row, "  Billing Title:",
                                         self.on_billing_title)
        row += 1
        self.billingAddress = contact_line(row, "  Billing Address:",
                                                self.on_billing_address)
        row += 1
        self.billingCSZ = contact_line(row, "  City / State / Zip:",
                                       self.on_billing_csz)
        row += 1
        self.billingPhone = contact_line(row, "  Billing phone",
                                         self.on_billing_phone)
        row += 1
        self.billingEmail = contact_line(row, "  Billing email",
                                         self.on_billing_email)
        row += 1
        self.sizer.Add(wx.StaticLine(self.panel), pos=(row, 0), span=(1, 4))

        self.panel.SetSizerAndFit(self.sizer)

    def msg(self, message, field=0):
        """
        Display in Status area.
        """
        logger.info(message)
        self.sb.SetStatusText(message, field)

    def error(self, message, field=0):
        """
        Emit an error.
        """
        logger.error(message)
        self.sb.SetStatusText(message, field)
        wx.MessageBox(message, 'Error', wx.ICON_ERROR | wx.OK)

    def on_new(self, event):
        logger.debug("Setting mode to New")
        self.mode = 'Create'
        self.project.year = ps.CURRENT_YEAR
        self.project.number = self.project.next_project_number()
        full_number = self.project.get_full_number()
        self.projectNumber.SetValue(full_number)
        self.searchButton.Enable(False)
        self.searchString.Enable(False)
        self.msg(f"Creating new project #{full_number}")
        event.Skip()

    def on_edit(self, event):
        logger.debug("Setting mode to Edit")
        self.mode = 'Edit'
        self.searchButton.Enable(True)
        self.searchString.Enable(True)
        self.msg("Enter text to search")
        event.Skip()

    def on_go(self, event):
        logger.debug("Called GO")
        self.project.name = self.projectName.GetValue()
        full_name = self.project.get_full_name()
        if self.mode == 'Create':
            self.msg(f"Creating project {full_name}")
            self.create_project()
        else:
            self.msg(f"Updating project {full_name}")
            self.update_project()
        event.Skip()

    def on_search(self, event):
        logger.debug("Called Search")
        event.Skip()

    def on_number(self, event):
        logger.debug("on_number")
        event.Skip()

    def on_project_name(self, event):
        logger.debug("on_project_name")
        event.Skip()

    def on_project_manager(self, event):
        logger.debug("on_project_manager")
        event.Skip()

    def on_project_type(self, event):
        logger.debug("on_project_type")
        event.Skip()

    def on_contact_name(self, event):
        logger.debug("on_contact_name")
        event.Skip()

    def on_contact_title(self, event):
        logger.debug("on_contact_title")
        event.Skip()

    def on_contact_address(self, event):
        logger.debug("on_contact_address")
        event.Skip()

    def on_contact_csz(self, event):
        logger.debug("on_contact_csz")
        event.Skip()

    def on_contact_phone(self, event):
        logger.debug("on_contact_phone")
        event.Skip()

    def on_contact_email(self, event):
        logger.debug("on_contact_email")
        event.Skip()

    def on_billing_name(self, event):
        logger.debug("on_billing_name")
        event.Skip()

    def on_billing_title(self, event):
        logger.debug("on_billing_title")
        event.Skip()

    def on_billing_address(self, event):
        logger.debug("on_billing_address")
        event.Skip()

    def on_billing_csz(self, event):
        logger.debug("on_billing_csz")
        event.Skip()

    def on_billing_phone(self, event):
        logger.debug("on_billing_phone")
        event.Skip()

    def on_billing_email(self, event):
        logger.debug("on_billing_email")
        event.Skip()

    def create_project(self):
        """
        Create the project folder and populate based on type.
        Add available data to the project info spreadsheet.
        """
        # Populate data from GUI
        self.project.name = self.projectName.GetValue()
        if len(self.project.name) < 3:
            self.error("Check project name")
            return False
        self.project.project_manager = self.projectManager.GetValue()
        self.project.project_type = self.projectType.GetValue()
        self.project.project_contact.name = self.contactName.GetValue()
        self.project.project_contact.address = self.contactAddress.GetValue()
        self.project.project_contact.address.csz = self.contactCSZ.GetValue()
        self.project.project_contact.address.email =\
            self.contactEmail.GetValue()
        self.project.project_contact.address.phone =\
            self.contactPhone.GetValue()
        self.project.project_billing.name = self.billingName.GetValue()
        self.project.project_billing.address = self.billingAddress.GetValue()
        self.project.project_billing.address.csz = self.billingCSZ.GetValue()
        self.project.project_billing.address.email =\
            self.billingEmail.GetValue()
        self.project.project_billing.address.phone =\
            self.billingPhone.GetValue()
        # Now build the folder
        try:
            self.project.makeProjectFolder()
        except Exception:
            self.error("Unable to create: {self.project.get_full_name()}")
        # Copy the raw data from the template folder
        # Fill the spreadsheet that should be already in the folder
        # Name the billing spreadsheet
        # Name the proposal fileEdited


def main():
    """Top level function."""
    global logger
    logger.setLevel(logging.DEBUG)
    if len(sys.argv) == 2:
        flag = sys.argv[1]
        if flag == '-i':
            logger.debug('Setting logging level to INFO')
            logger.setLevel(logging.INFO)
        elif flag == '-w':
            logger.debug('Setting logging level to WARN')
            logger.setLevel(logging.WARN)
        else:
            logger.debug(f'Flag was <{flag}>')
    app = wx.App()
    Application().Show()
    app.MainLoop()


if __name__ == '__main__':
    main()
