import traceback

import adsk.core
import adsk.fusion
import os
from ...lib import fusion360utils as futil
from ... import config

app = adsk.core.Application.get()
ui = app.userInterface

CMD_NAME = os.path.basename(os.path.dirname(__file__))
CMD_ID = f'{config.ADDIN_NAME}_{CMD_NAME}'
CMD_Description = 'Generates a bill of materials for the current design'
IS_PROMOTED = False

# Global variables by referencing values from /config.py
WORKSPACE_ID = config.design_workspace
TAB_ID = config.tools_tab_id
TAB_NAME = config.my_tab_name

PANEL_ID = config.my_panel_id
PANEL_NAME = config.my_panel_name
PANEL_AFTER = config.my_panel_after

# Resource location for command icons, here we assume a sub folder in this directory named "resources".
ICON_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources', '')

# radio button options
OPT_COLON = ", (colon)"
OPT_SEMICOLON = "; (semicolon)"

# Holds references to event handlers
local_handlers = []


# Executed when add-in is run.
def start():
    # ******************************** Create Command Definition ********************************
    cmd_def = ui.commandDefinitions.addButtonDefinition(CMD_ID, CMD_NAME, CMD_Description, ICON_FOLDER)

    # Add command created handler. The function passed here will be executed when the command is executed.
    futil.add_handler(cmd_def.commandCreated, command_created)

    # ******************************** Create Command Control ********************************
    # Get target workspace for the command.
    workspace = ui.workspaces.itemById(WORKSPACE_ID)

    # Get target toolbar tab for the command and create the tab if necessary.
    toolbar_tab = workspace.toolbarTabs.itemById(TAB_ID)
    if toolbar_tab is None:
        toolbar_tab = workspace.toolbarTabs.add(TAB_ID, TAB_NAME)

    # Get target panel for the command and and create the panel if necessary.
    panel = toolbar_tab.toolbarPanels.itemById(PANEL_ID)
    if panel is None:
        panel = toolbar_tab.toolbarPanels.add(PANEL_ID, PANEL_NAME, PANEL_AFTER, False)

    # Create the command control, i.e. a button in the UI.
    control = panel.controls.addCommand(cmd_def)

    # Now you can set various options on the control such as promoting it to always be shown.
    control.isPromoted = IS_PROMOTED


# Executed when add-in is stopped.
def stop():
    # Get the various UI elements for this command
    workspace = ui.workspaces.itemById(WORKSPACE_ID)
    panel = workspace.toolbarPanels.itemById(PANEL_ID)
    toolbar_tab = workspace.toolbarTabs.itemById(TAB_ID)
    command_control = panel.controls.itemById(CMD_ID)
    command_definition = ui.commandDefinitions.itemById(CMD_ID)

    # Delete the button command control
    if command_control:
        command_control.deleteMe()

    # Delete the command definition
    if command_definition:
        command_definition.deleteMe()

    # Delete the panel if it is empty
    if panel.controls.count == 0:
        panel.deleteMe()

    # Delete the tab if it is empty
    if toolbar_tab.toolbarPanels.count == 0:
        toolbar_tab.deleteMe()


# Function to be called when a user clicks the corresponding command button in the UI.
def command_created(args: adsk.core.CommandCreatedEventArgs):
    #futil.log(f'{CMD_NAME} Command Created Event')

    # Connect to the events that are needed by this command.
    futil.add_handler(args.command.execute, command_execute, local_handlers=local_handlers)
    futil.add_handler(args.command.inputChanged, command_input_changed, local_handlers=local_handlers)
    futil.add_handler(args.command.destroy, command_destroy, local_handlers=local_handlers)

    inputs = args.command.commandInputs

    text_box_message = "Generate Bill Of Materials as CSV file"
    text_box_input = inputs.addTextBoxCommandInput('text_box_input', 'Text Box', text_box_message, 1, True)
    text_box_input.isFullWidth = True

    radio_input = inputs.addRadioButtonGroupCommandInput('radio_input', 'Field separator')
    radio_items = radio_input.listItems
    radio_items.add(OPT_COLON, True)
    radio_items.add(OPT_SEMICOLON, False)
    radio_input.isFullWidth = True

    include_assemblies_input = inputs.addBoolValueInput('include_assemblies', 'Include assemblies', True, '', False)
    include_assemblies_input.tooltip = "If it does not include assemblies, the generated BOM will only have components that are not assemblies (i.e., that do not have sub-components)"

    include_hidden_input = inputs.addBoolValueInput('include_hidden', 'Include hidden', True, '', False)
    include_hidden_input.tooltip = ""

# This function will be called when the user clicks the OK button in the command dialog.
def command_execute(args: adsk.core.CommandEventArgs):
    #futil.log(f'{CMD_NAME} Command Execute Event')
    inputs = args.command.commandInputs
    
    # Display the palette that represents the TEXT COMMANDS palette
    text_palette = ui.palettes.itemById('TextCommands')
    if not text_palette.isVisible:
        text_palette.isVisible = True

    # Create BOM file
    create_bom(inputs)

# This function will be called when the user changes anything in the command dialog.
def command_input_changed(args: adsk.core.InputChangedEventArgs):
    pass

# This function will be called when the user completes the command.
def command_destroy(args: adsk.core.CommandEventArgs):
    global local_handlers
    local_handlers = []


def create_bom(inputs):
    global ui

    sep = ','
    separator_selected = inputs.itemById('radio_input').selectedItem.name
    if separator_selected == OPT_SEMICOLON:
        sep = ";"

    include_assemblies = inputs.itemById('include_assemblies').value
    include_hidden = inputs.itemById('include_hidden').value

    try:
        product = app.activeProduct
        design = adsk.fusion.Design.cast(product)
        title = 'Extract BOM'
        if not design:
            ui.messageBox('No active design', title)
            return

        # Gather information about each unique component
        dbom = {}
        for occ in design.rootComponent.allOccurrences:
            comp = occ.component
            
            # Skips the occurrence if it is an assembly and the user does not want to include them in the BOM
            is_assembly = occ.childOccurrences.count > 0
            if not include_assemblies and is_assembly:
                continue

            # Skips the occurrence if it is not visible and the user does not want to include hidden entities
            if not include_hidden and not occ.isVisible:
                continue

            if comp.partNumber not in dbom.keys():
                dbom[comp.partNumber] = {'Name': comp.name, 'Description': comp.description, 'Instances': 1, 'Assembly': 'True' if is_assembly else 'False'}
            else:
                dbom[comp.partNumber]['Instances'] += 1

        msg = f"Part number{sep}Name{sep}Description{sep}Instances{sep}Assembly\n"
        for pn, props in dbom.items():
            msg = msg + f"{str(pn).replace(sep, '-')}{sep}"
            for v in props.values():
                msg = msg + f"{str(v).replace(sep, '-')}{sep}"
            msg = msg + "\n"
      
        # Set styles of file dialog.
        fileDlg = ui.createFileDialog()
        fileDlg.isMultiSelectEnabled = False
        fileDlg.filter = '*.csv'
        fileDlg.title = 'Fusion Save File Dialog'
        
        # Show file save dialog
        dlgResult = fileDlg.showSave()
        if dlgResult == adsk.core.DialogResults.DialogOK:
            with open(fileDlg.filename, "w") as f:
                f.write(msg)
        else:
            return
        
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
