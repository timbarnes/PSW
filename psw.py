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
                          'Studio G Project Editor', size=(550, 775))
        self.project = ps.Project()
        self.build_GUI()

    def build_GUI(self):
        """
        Create all the elements of the UI and connect to variables.
        """

        def single_line(row, label, callback=None):
            """
            Make a line with StaticText and TextCtrl and callback.
            """
            prompt = wx.StaticText(self.panel, label=label)
            self.sizer.Add(prompt, pos=(row, 0), flag=wx.ALL, border=4)
            t = wx.TextCtrl(self.panel)
            if callback:
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
        self.editButton = wx.Button(self.panel, label="Load Project Info.")
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
        self.searchString.Bind(wx.EVT_KEY_DOWN, self.on_check_search)
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
        self.projectManager = single_line(row, "Project Manager:", self.clean)
        row += 1
        label = wx.StaticText(self.panel, label="Project Type:")
        self.sizer.Add(label, pos=(row, 0), flag=wx.ALL, border=4)
        type_list = ['Empty', 'Revit', 'CAD']
        self.projectType = wx.RadioBox(self.panel, choices=type_list,
                                       majorDimension=1,
                                       style=wx.RA_SPECIFY_ROWS)
        # self.projectType.Bind(wx.EVT_RADIOBUTTON, self.on_project_type)
        self.sizer.Add(self.projectType, pos=(row, 1), span=(0, 4),
                       flag=wx.ALL, border=5)
        row += 1
        self.sizer.Add(wx.StaticLine(self.panel), pos=(row, 0), span=(1, 4))
        row += 1
        label = wx.StaticText(self.panel, label="PROJECT INFORMATION")
        self.sizer.Add(label, pos=(row, 1), span=(0, 3), flag=wx.EXPAND)
        row += 1
        self.scope = single_line(row, "Project Scope:", self.clean)
        # Needs to be set to be multiline
        self.scope.SetMinSize(wx.Size(70, 70))
        row += 1
        self.contactName = single_line(row, "Project Contact:", self.clean)
        row += 1
        self.contactTitle = single_line(row, "  Contact Title:", self.clean)
        row += 1
        self.contactPhone = single_line(row, "  Contact phone")
        row += 1
        self.contactEmail = single_line(row, "  Contact email", self.clean)
        row += 1
        self.contactAddress = single_line(row, "  Project Address:", self.clean)
        row += 1
        self.contactCSZ = single_line(row, "  City / State / Zip:", self.clean)
        row += 1
        self.sizer.Add(wx.StaticLine(self.panel), pos=(row, 0), span=(1, 4))
        row += 1
        label = wx.StaticText(self.panel, label="BILLING INFORMATION")
        self.sizer.Add(label, pos=(row, 1), span=(0, 3), flag=wx.EXPAND)
        row += 1
        self.billingName = single_line(row, "Billing Name:", self.clean)
        row += 1
        self.billingTitle = single_line(row, "  Billing Title:", self.clean)
        row += 1
        self.billingPhone = single_line(row, "  Billing phone", self.clean)
        row += 1
        self.billingEmail = single_line(row, "  Billing email", self.clean)
        row += 1
        self.billingAddress = single_line(row, "  Billing Address:", self.clean)
        row += 1
        self.billingCSZ = single_line(row, "  City / State / Zip:", self.clean)
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

    def popup_list(self, prompt, position, fields):
        """
        Popup a list of fields at position, and return the one selected.
        """
        # listbox = wx.ListBox(self.panel)
        self.menu = wx.Menu()
        for item in fields:
            print(item)
            m_item = self.menu.Append(-1, item)
            self.panel.Bind(wx.EVT_MENU, self.on_menu, m_item)
        self.msg(prompt)
        self.panel.PopupMenu(self.menu, position)

    def on_menu(self, event):
        logger.debug("Called on_menu")
        item = self.menu.FindItemById(event.GetId())
        year, num, name = self.project.from_folder(item.GetText())
        self.projectNumber.SetValue(f"{year}.{num}")
        self.projectName.SetValue(name)
        logger.debug(item.GetText())

    def on_new(self, event):
        logger.debug("Setting mode to New")
        full_number = self.project.start_new()
        # Prime the GUI appropriately
        self.projectNumber.SetValue(full_number)
        self.searchButton.Enable(False)
        self.searchString.Enable(False)
        self.msg(f"Creating new project #{full_number}")
        event.Skip()

    def on_edit(self, event):
        logger.debug("Setting mode to Edit")
        self.project.mode = 'Edit'
        self.searchButton.Enable(True)
        self.searchString.Enable(True)
        if self.project.number and self.project.name:
            # We have identified the project
            if self.project.open() and self.project.load():
                self.msg("Project loaded")
                # Populate the GUI from the record
                self.projectManager.SetValue(self.project.manager)
                logger.debug(self.project.type)
                t = ["Other", "Revit", "CAD"].index(self.project.type)
                self.projectType.SetSelection(t)
                self.scope.SetValue(self.project.scope)
                self.contactName.SetValue(self.project.contact.name)
                self.contactTitle.SetValue(self.project.contact.title)
                self.contactPhone.SetValue(self.project.contact.phone)
                self.contactEmail.SetValue(self.project.contact.email)
                self.contactAddress.SetValue(self.project.contact.address)
                self.contactCSZ.SetValue(self.project.contact.csz)
                self.billingName.SetValue(self.project.billing.name)
                self.billingTitle.SetValue(self.project.billing.title)
                self.billingPhone.SetValue(self.project.billing.phone)
                self.billingEmail.SetValue(self.project.billing.email)
                self.billingAddress.SetValue(self.project.billing.address)
                self.billingCSZ.SetValue(self.project.billing.csz)
                self.project.is_live = True
            else:
                self.error(self.project.fileError)
        else:
            self.error("Please select a project to load")

    def on_go(self, event):
        logger.debug("Called GO")
        self.project.name = self.projectName.GetValue()
        full_name = self.project.get_full_name()
        self.transfer_from_GUI()
        v, r = self.project.validate()
        if not v:
            self.error(r)
            return
        if self.project.mode == 'Create':
            self.msg(f"Creating project {full_name}")
            self.create()
        elif self.project.mode == "Edit":
            self.msg(f"Updating project {full_name}")
            self.project.update()
        else:
            self.error(f"Bad mode: {self.project.mode}")
        event.Skip()

    def on_check_search(self, event):
        if event.GetKeyCode() in [wx.WXK_RETURN, wx.WXK_TAB, wx.WXK_CONTROL_F]:
            self.on_search(event)
        else:
            event.Skip()

    def on_search(self, event):
        """
        Search for and select a folder.
        """
        logger.debug("Called Search")
        search_string = self.searchString.GetValue()
        results = self.project.match_folder(search_string)
        logger.debug(results)
        if len(results) > 20:
            self.error("Too many results. Try again")
        else:
            pos = self.ScreenToClient(wx.GetMousePosition())
            # Put up the popup menu. Results captured in callback.
            self.popup_list("Select a project", pos, results)

    def on_number(self, event):
        logger.debug("on_number")
        pass
        # We block any user activity in this field

    def on_project_name(self, event):
        logger.debug("on_project_name")
        if self.project.mode == "Edit":
            # Changes to the project name are disallowed
            pass
        else:
            self.project.name = self.projectName.GetValue()
            event.Skip()

    def clean(self, event):
        """
        Cleans text fields.
        """
        val = event.GetKeyCode()
        if val == wx.WXK_TAB:
            ctl = self.panel.Window_FindFocus()
            ctl.Navigate()
        elif chr(val) in ps.BAD_CHARS:
            pass
        else:
            event.Skip()

    def transfer_from_GUI(self):
        """
        Copy all the data from the GUI to the project record
        """
        self.project.name = self.projectName.GetValue()
        self.project.manager = self.projectManager.GetValue()
        options = ['Other', 'Revit', 'CAD']
        self.project.type = options[self.projectType.GetSelection()]
        self.project.scope = self.scope.GetValue()
        self.project.contact.name = self.contactName.GetValue()
        self.project.contact.title = self.contactTitle.GetValue()
        self.project.contact.address = self.contactAddress.GetValue()
        self.project.contact.csz = self.contactCSZ.GetValue()
        self.project.contact.email = self.contactEmail.GetValue()
        self.project.contact.phone = self.contactPhone.GetValue()
        self.project.billing.name = self.billingName.GetValue()
        self.project.billing.title = self.billingTitle.GetValue()
        self.project.billing.address = self.billingAddress.GetValue()
        self.project.billing.csz = self.billingCSZ.GetValue()
        self.project.billing.email = self.billingEmail.GetValue()
        self.project.billing.phone = self.billingPhone.GetValue()

    def create(self):
        """
        Create the project folder and populate based on type.
        Add available data to the project info spreadsheet.
        """
        if self.project.mode == 'Create':
            self.transfer_from_GUI()
            valid, response = self.project.validate()
            if valid:
                self.project.create()
            else:
                self.error(response)
        else:
            self.error("Mode error: Create called inappropriately")

    def update(self):
        """
        Validate entries and update the spreadsheet.
        """
        valid, response = self.project.validate_existing()
        if valid:
            self.project.modify()
        else:
            self.error(response)


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
