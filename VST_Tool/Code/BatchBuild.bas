Attribute VB_Name = "BatchBuild"
Option Explicit

' Sub routine to add A2L files using button
Sub AddA2Ls()
    Dim AtiFileCell As Range
    Dim FileFilter As String, FileType As String
    Dim FileColumn As Long
    ' Define Ranges
    Set AtiFileCell = Worksheets("File Paths").Range("B2")
    FileFilter = "(*.ati;*.a2l),*.ati;*.a2l"
    FileType = "A2L"
    FileColumn = 1
    GetBatchBuildFiles AtiFileCell, FileFilter, FileType, FileColumn
End Sub

' Sub routine to add MAP files using button
Sub AddMAPs()
    Dim MapFileCell As Range
    Dim FileFilter As String, FileType As String
    Dim FileColumn As Long
    ' Define Ranges
    Set MapFileCell = Worksheets("File Paths").Range("B5")
    FileFilter = "(*.map),*.map"
    FileType = "MAP"
    FileColumn = 3
    GetBatchBuildFiles MapFileCell, FileFilter, FileType, FileColumn
End Sub

' Sub routine to add H32 files using button
Sub AddH32s()
    Dim H32FileCell As Range
    Dim FileFilter As String, FileType As String
    Dim FileColumn As Long
    ' Define Ranges
    Set H32FileCell = Worksheets("File Paths").Range("B3")
    FileFilter = "(*.h32),*.h32"
    FileType = "H32"
    FileColumn = 5
    GetBatchBuildFiles H32FileCell, FileFilter, FileType, FileColumn
End Sub

' Function to add user selected files to the table
Sub GetBatchBuildFiles(AtiFileCell As Range, FileFilter As String, FileType As String, FileColumn As Long)
    Dim StartRow As Long
    Dim AtiFilePath As Variant
    Dim n As Long
    ' Get the data start row (append files if already existed)
    StartRow = Worksheets("Batch Build").Cells(Worksheets("Batch Build").Rows.Count, FileColumn).End(xlUp).Row
    
    ' Change to saved directory and open dialog
    If (FileOrDirExists(AtiFileCell.Value)) Then
        ChDir AtiFileCell.Value
    End If
    FileFilter = "Strategy Description File " + FileFilter
    FileType = "Select " + FileType + " files"
    ' Choose Files
    AtiFilePath = Application.GetOpenFilename(FileFilter, , FileType, , True)
    ' If user cancels file selection
    If Not (IsArray(AtiFilePath)) Then
        If AtiFilePath = False Then
            Exit Sub
        End If
    End If
    ' Add selected files to the table
    For n = LBound(AtiFilePath) To UBound(AtiFilePath)
        Worksheets("Batch Build").Cells(StartRow + n, FileColumn).Value = FileParts(AtiFilePath(n), "path")
        Worksheets("Batch Build").Cells(StartRow + n, FileColumn + 1).Value = FileParts(AtiFilePath(n), "filename")
    Next n
    
End Sub

' Function to clear the selected files data in table
Sub ClearAddedFiles()
    If Worksheets("Batch Build").Cells(Worksheets("Batch Build").Rows.Count, "A").End(xlUp).Row > 8 Then
        With Worksheets("Batch Build")
            .Range("A9:F" & .Cells(.Rows.Count, "A").End(xlUp).Row).ClearContents
        End With
    Else
        With Worksheets("Batch Build")
            .Range("A9:F10000").ClearContents
        End With
    End If
End Sub

'Function to run batch build mode with multiple A2L files
Sub RunBatchBuild()
    Dim nRows As Long
    Dim FileNames() As String
    Dim ShowMisMatchError As Boolean
    ShowMisMatchError = True
    ' Go from row 9 to end
    For nRows = 9 To Worksheets("Batch Build").Cells(Worksheets("Batch Build").Rows.Count, "A").End(xlUp).Row
        ' First empty cell in column 1 denotes end of table data
        If IsEmpty(Worksheets("Batch Build").Cells(nRows, 1).Value) Then
            Exit For
        End If
        ' clear the variable data
        ReDim FileNames(2) As String
        ' Check if user has given A2L file path and name
        If Not (IsEmpty(Worksheets("Batch Build").Cells(nRows, 1).Value)) And (Not (IsEmpty(Worksheets("Batch Build").Cells(nRows, 2).Value))) Then
            ' Get selected A2L file name
            FileNames(0) = Worksheets("Batch Build").Cells(nRows, 1).Value + Worksheets("Batch Build").Cells(nRows, 2).Value
            ' Check if the file exist
            If FileOrDirExists(FileNames(0)) = False Then
                msgLogDisp ("A2L file does not found in Row number: " + CStr(nRows) + ". So skiping this row.")
                Worksheets("Batch Build").Range("A" + CStr(nRows) + ":F" + CStr(nRows)).Font.Color = RGB(255, 0, 0)
                GoTo nextRow
            End If
        Else
            Worksheets("Batch Build").Range("A" + CStr(nRows) + ":F" + CStr(nRows)).Font.Color = RGB(255, 0, 0)
            GoTo nextRow
        End If
        ' Check if user has given MAP file path and name
        If Not (IsEmpty(Worksheets("Batch Build").Cells(nRows, 3).Value)) And (Not (IsEmpty(Worksheets("Batch Build").Cells(nRows, 4).Value))) Then
            ' Get selected MAP file name
            FileNames(1) = Worksheets("Batch Build").Cells(nRows, 3).Value + Worksheets("Batch Build").Cells(nRows, 4).Value
            ' Check if the file exist
            If FileOrDirExists(FileNames(1)) = False Then
                msgLogDisp ("MAP file does not found in Row number: " + CStr(nRows) + ". So skiping this row.")
                Worksheets("Batch Build").Range("A" + CStr(nRows) + ":F" + CStr(nRows)).Font.Color = RGB(255, 0, 0)
                'ActiveSheet.Cells(RowNum, ParamCol).Font.ColorIndex = 3
                GoTo nextRow
            End If
        Else
            Worksheets("Batch Build").Range("A" + CStr(nRows) + ":F" + CStr(nRows)).Font.Color = RGB(255, 0, 0)
        End If
        ' Check if user has given H32 file path and name
        If Not (IsEmpty(Worksheets("Batch Build").Cells(nRows, 5).Value)) And (Not (IsEmpty(Worksheets("Batch Build").Cells(nRows, 6).Value))) Then
            ' Get selected H32 file name
            FileNames(2) = Worksheets("Batch Build").Cells(nRows, 5).Value + Worksheets("Batch Build").Cells(nRows, 6).Value
            ' Check if the file exist
            If FileOrDirExists(FileNames(2)) = False Then
                msgLogDisp ("H32 file does not found in Row number: " + CStr(nRows) + ". So skiping this row.")
                GoTo nextRow
            End If
        Else
            msgLogDisp ("H32 file info incorrect in Row number: " + CStr(nRows) + ". So skiping this row.")
            Worksheets("Batch Build").Range("A" + CStr(nRows) + ":F" + CStr(nRows)).Font.Color = RGB(255, 0, 0)
            GoTo nextRow
        End If
        ' one time pop up message to the user
        If nRows = 9 Then
            msgLogDisp "In batch mode, each VST file will be named to match the H32 file.", vbInformation
        End If
        Worksheets("Batch Build").Range("A" + CStr(nRows) + ":F" + CStr(nRows)).Font.Color = RGB(0, 0, 0)
        'call the orignal batch build function
        BuildVSTFile True, ShowMisMatchError, FileNames
nextRow:
    Next nRows
End Sub
