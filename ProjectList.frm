VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProjectList 
   Caption         =   "Project List Userform"
   ClientHeight    =   3216
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   4740
   OleObjectBlob   =   "ProjectList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProjectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Functionality of import/export button.
Private Sub ProcessButton_Click()
    Dim i As Integer
    Dim strResult As String
    Do While i < ProjectList.Controls.Count - 1 And Not ProjectList.Controls(i)
        i = i + 1
    Loop
    If ProjectList.Controls(i) Then
        Main.gSelectedOption = ProjectList.Controls(i).Caption
        If InStr(UCase$(Main.gSelectedOption), ".XLAM") = 0 Then
            MsgBox Main.gSelectedOption & " has no valid path.", vbInformation
            Unload ProjectList
            Exit Sub
        End If
        If Main.gStrTask = "Import" Then
            strResult = ImportModules
        Else
            strResult = ExportModules
        End If
        If strResult <> vbNullString Then
            MsgBox strResult, vbInformation, "Process is completed"
            Unload ProjectList
        End If
    End If
End Sub

' Adjusts userform size, width, etc, depending on count of installed add-ins.
Private Sub UserForm_Initialize()
    Dim btn As CommandButton
    Set btn = ProjectList.ProcessButton
    Dim opt As control
    Dim strAddin As Variant
    Dim i As Integer
    With ProjectList
        .Caption = "Projects List"
        For Each strAddin In gStrAddins
            Set opt = .Controls.Add("Forms.OptionButton.1", "radioBtn" & i, True)
            With opt
                .Caption = strAddin
                .Top = .Height * i
                .GroupName = "Options"
                .Width = 120
            End With
            .Height = opt.Height * (i + 2.5)
            i = i + 1
        Next
        btn.Caption = Main.gStrTask
        btn.Top = .Height - btn.Height + (0.5 * opt.Height) - 3
        btn.Left = (.Width * 0.5) - (btn.Width * 0.5)
        .Height = .Height + btn.Height + (0.5 * opt.Height)
    End With
End Sub
