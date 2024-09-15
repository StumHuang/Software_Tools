Attribute VB_Name = "FileFunctions"
Option Explicit

Function FileParts(ByVal EntirePath As String, Part As String) As String
    Dim DotPos As Long, SlashPos As Long
    ' Find "." from the end
    DotPos = InStrRev(EntirePath, ".")
    
    ' Find "/" from the end
    SlashPos = InStrRev(EntirePath, "\")
    
    Select Case Part
        Case "ext"
            FileParts = Mid(EntirePath, DotPos + 1)
        Case "path"
            If (SlashPos > 1) Then
                FileParts = Left(EntirePath, SlashPos)
            Else
                FileParts = ""
            End If
        Case "nameonly"
            FileParts = Mid(EntirePath, SlashPos + 1, DotPos - SlashPos - 1)
        Case "filename"
            FileParts = Mid(EntirePath, SlashPos + 1)
        Case Else
            FileParts = "Bad Part Argument"
    End Select
    
End Function

Function FileOrDirExists(PathName As String) As Boolean
     'Macro Purpose: Function returns TRUE if the specified file
     '               or folder exists, false if not.
     'PathName     : Supports Windows mapped drives or UNC
     '             : Supports Macintosh paths
     'File usage   : Provide full file path and extension
     'Folder usage : Provide full folder path
     '               Accepts with/without trailing "\" (Windows)
     '               Accepts with/without trailing ":" (Macintosh)
     
    Dim iTemp As Long
     
     'Ignore errors to allow for error evaluation
    On Error Resume Next
    iTemp = GetAttr(PathName)
     
     'Check if error exists and set response appropriately
    Select Case Err.Number
    Case Is = 0
        FileOrDirExists = True
    Case Else
        FileOrDirExists = False
    End Select
     
     'Resume error checking
    On Error GoTo 0
End Function
