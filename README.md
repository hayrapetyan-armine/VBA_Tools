## Synopsis

VBA Tools is the Excel add-in for easy importing and exporting add-in source files (*.bas, etc).

## Motivation

If you write add-in using VBA you should manually import/export each file for committing, which is boring and time-consuming.
The aim of this add-in is solving it.

## Installation/Build

1. In the Excel create a new empty workbook
2. Open the Visual Basic Editor from Developer > Visual Basic Editor (Keys: ALT+F11). In the Project Explorer pane select VBAProject. This selects the empty workbook.
3. Right click on VBAProject then in the opened context menu choose "Import File..."
4. In the browse windows select necessary source files: ProjectList.frm and Main.bas. ProjectList.frx file should be in the same directory with ProjectList.frm
5. Go to the Tools > References. In the References dialog, scroll down to Microsoft Visual Basic for Applications Extensibility 5.3 and check that item in the list
6. The workbook containing your code module now has to be saved as an Excel Add-In (*.xlam) file
7. Add ribbon.xml into saved Add-In using Office Custom UI Editor as example.

## Add to Excel

1. Go to File > Options > Add-Ins. Select "Excel Add-ins" from Manage list and click on "Go..." button to open the Add-Ins dialog.
2. To install your Add-In place a tick in the check-box next to your Add-In's name and click [OK]. If it is absent from list Browse and select it before that.
3. New tab called "VBA Tools" will be appeared in Ribbon
4. Tick Options > Trust Center > Trust Center Settings > Macro settings > Trust Access to VBA Project

## Usage

Scenario 1:
1. Navigate to add-in tab named "VBA Tools"
2. Click Import button, prompt window should be appeared with list of currently installed add-ins
3. Choose an add-in where it is required to import files
4. Click Import on the prompt and file browsing dialog should be opened
5. Select necessary files and click Select
6. All existed files will be removed from add-in and ".frm", ".cls", ".bas" files will be imported

Scenario 2:
1. Navigate to add-in tab named "VBA Tools"
2. Click Export button, prompt window should be appeared with list of currently installed add-ins
3. Choose an add-in which sources you would like to export and click Export on the prompt
4. Source files of the add-in will be exported into the add-in's directory