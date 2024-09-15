Attribute VB_Name = "ShellFileCopyModule"
Option Explicit

Private Type SHFILEOPSTRUCT
        hwnd As Long
        wFunc As Long
        pFrom As String
        pTo As String
        fFlags As Integer
        fAnyOperationsAborted As Long
        hNameMappings As Long
        lpszProgressTitle As String
End Type

Private Declare PtrSafe Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_NOCONFIRMATION = &H10
Private Const FO_COPY = &H2

Public Function ShellFileCopy(src As String, dest As String, Optional NoConfirm As Boolean = False) As Boolean

'PURPOSE: COPY FILES VIA SHELL API
'THIS DISPLAYS THE COPY PROGRESS DIALOG BOX
'PARAMETERS: src: Source File (FullPath)
            'dest: Destination File (FullPath)
            'NoConfirm (Optional): If set to
            'true, no confirmation box
            'is displayed when overwriting
            'existing files, and no
            'copy progress dialog box is
            'displayed
            
            'Returns (True if Successful, false otherwise)
            
'EXAMPLE:
  'dim bSuccess as boolean
  'bSuccess = ShellFileCopy ("C:\MyFile.txt", "D:\MyFile.txt")

Dim WinType_SFO As SHFILEOPSTRUCT
Dim lRet As Long
Dim lflags As Long

lflags = FOF_ALLOWUNDO
If NoConfirm Then lflags = lflags & FOF_NOCONFIRMATION
With WinType_SFO
    .wFunc = FO_COPY
    .pFrom = src
    .pTo = dest
    .fFlags = lflags
End With

lRet = SHFileOperation(WinType_SFO)
ShellFileCopy = (lRet = 0)

End Function
