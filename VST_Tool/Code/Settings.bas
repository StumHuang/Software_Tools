Attribute VB_Name = "Settings"
Option Explicit

Sub ClearSettings()
    Application.ScreenUpdating = False
    
    Worksheets("File Paths").Range("B2:B5").ClearContents
    Worksheets("File Paths").Range("B8").ClearContents
    Worksheets("File Paths").Range("B16").ClearContents
    
    Worksheets("Parameters").Range("A4:CB65536").ClearContents
    
    Worksheets("State Var Colors").Range("A2:D65536").ClearContents
    
    Worksheets("Device Settings").Range("A2:C65536").ClearContents
    
    Worksheets("Memory Regions").Range("A2:E65536").ClearContents
    
    Worksheets("Cal Changes").Range("A2:C65536").ClearContents
    
    Worksheets("Added Parameters").Range("A2:C65536").ClearContents
    
    Worksheets("Other Settings").Range("B1").ClearContents
    Worksheets("Other Settings").Range("B4").ClearContents
    Worksheets("Other Settings").Range("B7").ClearContents
    Worksheets("Other Settings").Range("B10").ClearContents

    Worksheets("Other Settings").Shapes("NoHooksCheckBox").ControlFormat.Value = -4146
    Worksheets("Other Settings").Shapes("KamRegionCheckBox").ControlFormat.Value = -4146
    Worksheets("Other Settings").Shapes("AddToTreeCheckBox").ControlFormat.Value = -4146
    Worksheets("Other Settings").Shapes("AprCheckBox").ControlFormat.Value = -4146
    
    A2lDefaults
    
    Application.ScreenUpdating = True
End Sub

Sub A2lDefaults()
    Worksheets("A2L Import Settings").Range("B2").Value = ""
    Worksheets("A2L Import Settings").Range("B3").Value = False
    Worksheets("A2L Import Settings").Range("B4").Value = False
    Worksheets("A2L Import Settings").Range("B5").Value = False
    Worksheets("A2L Import Settings").Range("B6").Value = False
    Worksheets("A2L Import Settings").Range("B7").Value = True
    Worksheets("A2L Import Settings").Range("B8").Value = True
    Worksheets("A2L Import Settings").Range("B9").Value = True
    Worksheets("A2L Import Settings").Range("B10").Value = 1
    Worksheets("A2L Import Settings").Range("B11").Value = "#"
    Worksheets("A2L Import Settings").Range("B12").Value = True
    Worksheets("A2L Import Settings").Range("B13").Value = True
    Worksheets("A2L Import Settings").Range("B14").Value = True
    Worksheets("A2L Import Settings").Range("B15").Value = True
End Sub

Sub CopySettings()
    Dim OldFilename As Variant
    Dim wkb As Object
    Dim OldBook As Workbook, CurrentBook As Workbook
    Dim oldWS As Worksheet, curWS As Worksheet
    Dim RowNo As Long
    
    Set CurrentBook = ThisWorkbook
    
    OldFilename = Application.GetOpenFilename("Excel Workbook (*.xls; *.xlsm),*.xls;*.xlsm")
    If (OldFilename = False) Then
        Exit Sub
    End If
    
    ' Code to check if the excel is already open(From where settings is getting copied to new VST tool)
    Dim Filename As String
    Filename = Right(OldFilename, Len(OldFilename) - InStrRev(OldFilename, "\"))
    If AlreadyOpen(Filename) Then
        If msgLogDisp(Filename & " is already open and needs to be closed. Do you want to close it to copy settings?", vbYesNo, "Confirm", vbNo) = vbYes Then
            Set wkb = GetObject(OldFilename)
            wkb.Close
        Else
            Exit Sub
        End If
    End If
    
    Application.EnableEvents = False
    Set OldBook = Workbooks.Open(Filename:=OldFilename, ReadOnly:=True)
    ActiveWindow.Visible = False
    
    Application.ScreenUpdating = False
    
    If (Not DoesWorkSheetExist("Parameters", OldBook.Name)) Then
        msgLogDisp "This does not appear to be a VST Tool spreadsheet"
        OldBook.Close
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ThisWorkbook.Activate
    ClearSettings
    
    ' File Paths
    If (DoesWorkSheetExist("File Paths", OldBook.Name)) Then
        CopyCell CurrentBook, OldBook, "File Paths", 2, 2
        CopyCell CurrentBook, OldBook, "File Paths", 3, 2
        CopyCell CurrentBook, OldBook, "File Paths", 4, 2
        CopyCell CurrentBook, OldBook, "File Paths", 5, 2
        CopyCell CurrentBook, OldBook, "File Paths", 8, 2
    ElseIf (DoesWorkSheetExist("Build VST", OldBook.Name)) Then
        Set curWS = CurrentBook.Sheets("File Paths")
        Set oldWS = OldBook.Sheets("Build VST")
        curWS.Range("B2").Value = oldWS.Range("B7").Value
        curWS.Range("B3").Value = oldWS.Range("B8").Value
        curWS.Range("B4").Value = oldWS.Range("B9").Value
        ' Don't even try MAP path because it didn't exit
        ' AddStates may or may not exist, so check first
        If (curWS.Range("A8").Value = oldWS.Range("A12").Value) Then
            curWS.Range("B8").Value = oldWS.Range("B12").Value
        End If
    End If
    
    ' Parameters
    Set curWS = CurrentBook.Sheets("Parameters")
    Set oldWS = OldBook.Sheets("Parameters")
    RowNo = oldWS.Range("A" & oldWS.Rows.Count).End(xlUp).Row
    ' Check if any parameter found (value greater than 3 as parameter starts from 3rd row)
    If RowNo > 3 Then
        oldWS.Range("A4:CB" & RowNo).Copy
        curWS.Activate
        curWS.Range("A4").Select
        curWS.Paste
        curWS.Range("A1").Select
    End If
    
    ' State Var Colors
    If (DoesWorkSheetExist("State Var Colors", OldBook.Name)) Then
        Set curWS = CurrentBook.Sheets("State Var Colors")
        Set oldWS = OldBook.Sheets("State Var Colors")
        RowNo = oldWS.Range("A" & oldWS.Rows.Count).End(xlUp).Row
        If RowNo > 1 Then
            oldWS.Range("A2:D" & RowNo).Copy
            curWS.Activate
            curWS.Range("A2").Select
            curWS.Paste
            curWS.Range("A1").Select
        End If
    End If
    
    ' A2L Import Settings
    If (DoesWorkSheetExist("A2L Import Settings", OldBook.Name)) Then
        Set curWS = CurrentBook.Sheets("A2L Import Settings")
        Set oldWS = OldBook.Sheets("A2L Import Settings")
        oldWS.Range("B2:B15").Copy
        curWS.Activate
        curWS.Range("B2").Select
        curWS.Paste
        curWS.Range("A1").Select
    End If
    
    ' Device Settings
    If (DoesWorkSheetExist("Device Settings", OldBook.Name)) Then
        Set curWS = CurrentBook.Sheets("Device Settings")
        Set oldWS = OldBook.Sheets("Device Settings")
        RowNo = oldWS.Range("A" & oldWS.Rows.Count).End(xlUp).Row
        ' Check if any parameter found (value greater than 1 as parameter starts from 3rd row)
        If RowNo > 1 Then
            'RowNo = oldWS.Range("A2" & oldWS.Rows.Count).End(xlUp).Row
            oldWS.Range("A2:C" & RowNo).Copy
            curWS.Activate
            curWS.Range("A2").Select
            curWS.Paste
            curWS.Range("A1").Select
        End If
    End If
    
    ' Memory Regions
    If (DoesWorkSheetExist("Memory Regions", OldBook.Name)) Then
        Set curWS = CurrentBook.Sheets("Memory Regions")
        Set oldWS = OldBook.Sheets("Memory Regions")
        RowNo = oldWS.Range("A" & oldWS.Rows.Count).End(xlUp).Row
        ' Check if any parameter found (value greater than 1 as parameter starts from 3rd row)
        If RowNo > 1 Then
            oldWS.Range("A2:E" & RowNo).Copy
            curWS.Activate
            curWS.Range("A2").Select
            curWS.Paste
            curWS.Range("A1").Select
        End If
    End If
    
    ' Cal Changes
    If (DoesWorkSheetExist("Cal Changes", OldBook.Name)) Then
        Set curWS = CurrentBook.Sheets("Cal Changes")
        Set oldWS = OldBook.Sheets("Cal Changes")
        RowNo = oldWS.Range("A" & oldWS.Rows.Count).End(xlUp).Row
        ' Check if any parameter found (value greater than 1 as parameter starts from 3rd row)
        If RowNo > 1 Then
            oldWS.Range("A2:C" & RowNo).Copy
            curWS.Activate
            curWS.Range("A2").Select
            curWS.Paste
            curWS.Range("A1").Select
        End If
    End If
    
    ' Added Parameters
    If (DoesWorkSheetExist("Added Parameters", OldBook.Name)) Then
        Set curWS = CurrentBook.Sheets("Added Parameters")
        Set oldWS = OldBook.Sheets("Added Parameters")
        RowNo = oldWS.Range("A" & oldWS.Rows.Count).End(xlUp).Row
        ' Check if any parameter found (value greater than 1 as parameter starts from 3rd row)
        If RowNo > 1 Then
            oldWS.Range("A2:C" & RowNo).Copy
            curWS.Activate
            curWS.Range("A2").Select
            curWS.Paste
            curWS.Range("A1").Select
        End If
    End If
    
    ' Other Settings
    If (DoesWorkSheetExist("Other Settings", OldBook.Name)) Then
        Set curWS = CurrentBook.Sheets("Other Settings")
        Set oldWS = OldBook.Sheets("Other Settings")
        ' Default decimals - been here for a long time, so no need to check the cell first
        CopyCell CurrentBook, OldBook, "Other Settings", 1, 2
        
        ' Y axis decimals
        If (curWS.Range("A4").Value = oldWS.Range("A4").Value) Then
            CopyCell CurrentBook, OldBook, "Other Settings", 4, 2
        End If
        
        ' X axis map decimals
        If (curWS.Range("A7").Value = oldWS.Range("A7").Value) Then
            CopyCell CurrentBook, OldBook, "Other Settings", 7, 2
        End If
        
        ' X axis curve decimals
        If (curWS.Range("A10").Value = oldWS.Range("A10").Value) Then
            CopyCell CurrentBook, OldBook, "Other Settings", 10, 2
        End If
        
        ' Check boxes
        On Error Resume Next ' In case control doesn't exist in old spreadsheet
        curWS.Shapes("NoHooksCheckBox").ControlFormat.Value = -4146
        curWS.Shapes("NoHooksCheckBox").ControlFormat.Value = oldWS.Shapes("NoHooksCheckBox").ControlFormat.Value
        curWS.Shapes("KamRegionCheckBox").ControlFormat.Value = -4146
        curWS.Shapes("KamRegionCheckBox").ControlFormat.Value = oldWS.Shapes("KamRegionCheckBox").ControlFormat.Value
        curWS.Shapes("AddToTreeCheckBox").ControlFormat.Value = -4146
        curWS.Shapes("AddToTreeCheckBox").ControlFormat.Value = oldWS.Shapes("AddToTreeCheckBox").ControlFormat.Value
        curWS.Shapes("AprCheckBox").ControlFormat.Value = -4146
        curWS.Shapes("AprCheckBox").ControlFormat.Value = oldWS.Shapes("AprCheckBox").ControlFormat.Value
        On Error GoTo 0
    End If
        
    Application.CutCopyMode = False
    OldBook.Close savechanges:=False
    Application.EnableEvents = True
    
    CurrentBook.Worksheets("Main").Activate
    Application.ScreenUpdating = True
End Sub

' Function to check if excel is already opened
Function AlreadyOpen(sFname As String) As Boolean
    Dim wkb As Workbook
    On Error Resume Next
    Set wkb = Workbooks(sFname)
    AlreadyOpen = Not wkb Is Nothing
    Set wkb = Nothing
    On Error GoTo 0
End Function

Sub CopyCell(Book1 As Workbook, Book2 As Workbook, SheetName As String, RowNum As Long, ColNum As Long)
    Book1.Worksheets(SheetName).Cells(RowNum, ColNum).Value = Book2.Worksheets(SheetName).Cells(RowNum, ColNum).Value
End Sub

Public Function DoesWorkSheetExist(WorkSheetName As String, Optional WorkBookName As String) As Boolean
    Dim WS As Worksheet
     
    On Error Resume Next
    If WorkBookName = vbNullString Then
        Set WS = Sheets(WorkSheetName)
    Else
        Set WS = Workbooks(WorkBookName).Sheets(WorkSheetName)
    End If
    On Error GoTo 0
    
    DoesWorkSheetExist = Not WS Is Nothing
End Function

