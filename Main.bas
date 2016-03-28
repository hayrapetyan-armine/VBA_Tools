Attribute VB_Name = "Main"
Option Explicit
Public gSelectedOption As String
Public gStrTask As String
Public gStrAddins() As String

' Detects import button is clicked, then adjusts userform.
Public Sub Init(control As IRibbonControl)
    If IsHasTrustAccess = False Then
        Exit Sub
    End If
    CollectAddInList
    If Len(Join(gStrAddins)) = 0 Then
        Exit Sub
    End If
    gStrTask = control.ID
    ProjectList.Show
End Sub

Private Function IsHasTrustAccess() As Boolean
    On Error Resume Next
    If Len(ThisWorkbook.VBProject.Name) Then
    End If
    If Err.Number Then
        MsgBox "This add-in needs access to VBA Project, please tick -" & vbCr & _
        "Options, Trust Center, Trust Center Settings, Macro settings, Trust Access to VBA Project", vbInformation, "Trust Access to VBA Project"
        IsHasTrustAccess = False
        Exit Function
    End If
    IsHasTrustAccess = True
End Function

' Collect add-ins list and set it into global variable
' Message box will be appeared if no one add-in is installed (except this one)
Public Sub CollectAddInList()
    Dim ad As AddIn
    Dim strAddinName As String
    Dim strTemp() As String: strTemp = Split(ThisWorkbook.VBProject.Filename, "\")
    For Each ad In Application.AddIns
        If ad.Installed And ad.Name <> strTemp(UBound(strTemp)) Then
            strAddinName = strAddinName & ad.Name & ","
        End If
    Next ad
    If Len(strAddinName) = 0 Then
        MsgBox "There is no installed add-in except ""VBA Tools""", vbInformation, "Add-ins are not found"
        Exit Sub
    End If
    gStrAddins = Split(Left$(strAddinName, Len(strAddinName) - 1), ",")
End Sub

' Exports all files in add-in directory.
Public Function ExportModules() As String
    Dim blnExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim strSourceWorkbook As String
    Dim strExportPath As String
    Dim strFileName As String
    Dim cmpComponent As VBIDE.VBComponent
    ''' NOTE: This workbook must be open in Excel.
    strSourceWorkbook = ActiveWorkbook.Name
    On Error GoTo errmsg
    Set wkbSource = Application.Workbooks(Main.gSelectedOption)
    If wkbSource.VBProject.Protection = 1 Then
        MsgBox "The VBA in this workbook is protected," & _
            "not possible to export the code", vbInformation
        Exit Function
    End If
    strExportPath = Application.AddIns(Split(Main.gSelectedOption, ".")(0)).Path & "\"
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        blnExport = True
        strFileName = cmpComponent.Name
        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                strFileName = strFileName & ".cls"
            Case vbext_ct_MSForm
                strFileName = strFileName & ".frm"
            Case vbext_ct_StdModule
                strFileName = strFileName & ".bas"
            Case vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                blnExport = False
        End Select
        If blnExport Then
            ''' Export the component to a text file.
            cmpComponent.Export strExportPath & strFileName
        End If
    Next cmpComponent
    ExportModules = "Export is ready"
    Exit Function
errmsg:
    MsgBox Main.gSelectedOption & " has no valid path.", vbInformation
End Function

' Imports only .cls, .frm, .bas type files, but you can select any file.
Public Function ImportModules() As String
    Dim wkbTarget As Excel.Workbook
    Dim objFSO As Scripting.FileSystemObject
    Dim cmpComponents As VBIDE.VBComponents
    Dim varFilePaths As Variant
    Dim i As Integer
    ''' NOTE: This workbook must be open in Excel.
    On Error GoTo errmsg
    Set wkbTarget = Application.Workbooks(Main.gSelectedOption)
    If wkbTarget.VBProject.Protection = 1 Then
        MsgBox "The VBA in this workbook is protected," & _
            "not possible to Import the code"
        Exit Function
    End If
    varFilePaths = Application.GetOpenFilename _
      (Title:="Please choose files to import", _
       MultiSelect:=True)
    If IsArray(varFilePaths) Then
        Set objFSO = New Scripting.FileSystemObject
        ''' Delete all modules/Userforms from the ActiveWorkbook.
        Call DeleteVBAModulesAndUserForms(wkbTarget)
        Set cmpComponents = wkbTarget.VBProject.VBComponents
        For i = LBound(varFilePaths) To UBound(varFilePaths)
            If (objFSO.GetExtensionName(varFilePaths(i)) = "cls") Or _
                (objFSO.GetExtensionName(varFilePaths(i)) = "frm") Or _
                (objFSO.GetExtensionName(varFilePaths(i)) = "bas") Then
                cmpComponents.Import varFilePaths(i)
            End If
        Next i
    Else
        Exit Function
    End If
    ImportModules = "Import is ready"
    Exit Function
errmsg:
    MsgBox Main.gSelectedOption & " has no valid path.", vbInformation
End Function

' Utility function to delete all modules and userforms before import.
Function DeleteVBAModulesAndUserForms(ByRef wkbTarget As Workbook)
        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        Set VBProj = wkbTarget.VBProject
        For Each VBComp In VBProj.VBComponents
            ''' Delete all except Thisworkbook and worksheet module.
            If VBComp.Type <> vbext_ct_Document Then
                VBProj.VBComponents.Remove VBComp
            End If
        Next VBComp
End Function
