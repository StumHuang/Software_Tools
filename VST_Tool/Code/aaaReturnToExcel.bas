Attribute VB_Name = "aaaReturnToExcel"
'*******************************************************************************
' ReturnToExcel
'
' Subroutines
'   ActivateExcel
' Functions
'*******************************************************************************
' Required Modules
'   None
'*******************************************************************************
' Required References
'   None
'*******************************************************************************
' Revision History
'   2015-04-05: Added generic header information
'               Clean up code formatting
'   2016-05-31: Update Win32 API declarations to be 64-bit friendly
'*******************************************************************************

Option Explicit
'
' Required Win32 API Declarations
'
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, lpdwProcessId As Long) As Long
Private Declare PtrSafe Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
'
' Constants used with APIs
'
Private Const SW_SHOW = 5
Private Const SW_RESTORE = 9

Sub ActivateExcel()
    ForceForegroundWindow Application.hwnd
End Sub


Public Function ForceForegroundWindow(ByVal hwnd As Long) As Boolean
    Dim ThreadID1 As Long
    Dim ThreadID2 As Long
    Dim nRet As Long
    '
    ' Nothing to do if already in foreground.
    '
    If hwnd = GetForegroundWindow() Then
        ForceForegroundWindow = True
    Else
        ' First need to get the thread responsible for
        ' the foreground window, then the thread running
        ' the passed window.
        '
        ThreadID1 = GetWindowThreadProcessId(GetForegroundWindow, ByVal 0&)
        ThreadID2 = GetWindowThreadProcessId(hwnd, ByVal 0&)
        '
        ' By sharing input state, threads share their
        ' concept of the active window.
        '
        If ThreadID1 <> ThreadID2 Then
            AttachThreadInput ThreadID1, ThreadID2, True
            nRet = SetForegroundWindow(hwnd)
            AttachThreadInput ThreadID1, ThreadID2, False
        Else
            nRet = SetForegroundWindow(hwnd)
        End If
        '
        ' Restore and repaint
        '
        If IsIconic(hwnd) Then
            ShowWindow hwnd, SW_RESTORE
        Else
            ShowWindow hwnd, SW_SHOW
        End If
        '
        ' SetForegroundWindow return accurately reflects
        ' success.
        ForceForegroundWindow = CBool(nRet)
    End If
End Function
