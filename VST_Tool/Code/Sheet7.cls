VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub BrowseAddStates_Click()
    Dim AddStatesPath As Variant
    AddStatesPath = Application.GetOpenFilename("AddStates.exe, *.exe")
    If (AddStatesPath <> False) Then
        ThisWorkbook.Worksheets("File Paths").Range("B8").Value = AddStatesPath
    End If
End Sub

Private Sub BrowseAprButton_Click()
    Dim AprPath As Variant
    AprPath = Application.GetSaveAsFilename("apr_summary.csv", "CSV Files (*.csv), *.csv", 1, "Select APR save file location")
    If (AprPath <> False) Then
        ThisWorkbook.Worksheets("File Paths").Range("B16").Value = AprPath
    End If
End Sub
