Attribute VB_Name = "aaaWebUpdate"
'*******************************************************************************
' WebUpdate Module
'
' To use this module:
' - Copy and paste the module to your workbook
' - In "This Workbook", and a call to "CheckForUpdatesAtStart" in the
'   "Workbook_Open" event.  The code should look like this:
'       Private Sub Workbook_Open()
'           CheckForUpdatesAtStart
'       End Sub
' - Update the values of the constants below.  Each setting has notes
'   in the comments above to explain the setting.
' - In order to store the web version in document properties, you must first create
'   the document property setting. Just run the TestAddCustomProperties sub once.
' Subroutines
'   CheckForUpdatesAtStart
'   CheckWebVersion (private)
'   RunInstaller (private)
'   DoTracking (private)
'   CheckHangingExcel
' Functions
'   GetWebVersion
'   GetWebVersionVal
'   GetCurrentVersionVal
'   CheckProcessRunning
'*******************************************************************************
' Required Modules
'   None
'*******************************************************************************
' Required References
'   None
'*******************************************************************************
' Revision History
'   2010-11-16 - First release
'   2010-11-18 - Updated comments but no functional changes
'   2016-03-30 - Update to read from SharePoint rather than web server
'                WinHTTP service no longer required
'                Add GetCurrentVersionVal function to supply interface to get version number
'   2016-04-05 - Remove GetWebVersionOld as it's not needed
'                Add additional header details
'                Write web version to custom document property to avoid re-reading from
'                  web every time we want to know the value (esp. when being referenced
'                  in the ribbon.
'   2016-04-08 - Add StatusBar update to indicate when update check is happening
'   2016-04-20 - Update to revert to using WinHTTP to fetch version number.
'                Enabled "auto login" setting on WinHttpReq to work better with Sharepoint
'                Only update WebVer property if it has changed
'   2016-04-21 - Add functionality for using a checkbox in the ribbon rather than on a worksheet
'                  for setting whether to update on start
'                Add functions to get/set update at start settings
'                Add ribbon callbacks
'   2016-09-20 - Include MsgBoxTitle within message box to ensure user knows what application
'                  needs updating.
'   2018-08-24 - Include auto updater. Now uses LinkDownloader to download a new version after
'                  Excel exits.
'   2018-10-07 - Make sure LinkDownloader exists
'   2019-01-16 - Add usage tracking feature
'   2019-02-01 - Add configurable option for using installer
'   2019-02-14 - Add error handling in RunInstaller
'   2019-03-02 - Add CheckRunningProcess function
'                Explicitly check if LinkDownloader is running before running again
'                Add subroutine to check if multiple Excel instances exist
'                Fix bug for false error message about LinkDownloader running
'   2020-04-28 - Use PCCNDATA webservices for version number and usage tracking
'*******************************************************************************

Option Explicit

' VersionFileUrl (string):  The URL to a text file containing the version number of the version
' available for download from the web.  The text file should contain only a version number
' which is a valid number.  In other words, no letters or other symbols.  It will be
' compared numerically to the value of "CurrentVersion"
Const VersionFileUrl = "https://pccndata.ford.com/api/v1/appdetails/vsttool"

' UseCheckBox (boolean):  If true, a checkbox will be referenced to determine whether or not to run
' the web check.  This checkbox allows the user to disable the auto check at start-up.
' Only set one of UseCheckBox and UseRibbonCheckBox to true
Const UseCheckBox = True

' UseRibbonCheckBox (boolean):  If true, a checkbox in the ribbon will be referenced to determine whether or not to run
' the web check.  This checkbox allows the user to disable the auto check at start-up.
' Only set one of UseCheckBox and UseRibbonCheckBox to true
Const UseRibbonCheckBox = False

' UseInstaller (boolean):  If true, will use the LinkDownloader tool to automatically download and install an
' update after Excel closes. If false, will simply direct user to the webpage.
Const UseInstaller = False

' CheckBoxWorksheet (string): Name of the worksheet which contains the checkbox to enable/disable
' checking for updates.  The checkbox object should be named "UpdateCheckBox".
Const CheckBoxWorksheet = "Revision History"

' MsgBoxTitle (string):  The title bar used in the message box when a new version is available.
Const MsgBoxTitle = "VST Tool"

' CurrentVersion:  A numeric version number for this document.  This is the value which will be
' compared against the version number on the web.
Const CurrentVersion = 240301

' DownloadUrl (string):  This is the URL to which the user will be directed to download a new version
Private DownloadUrl As String

' CheckForUpdatesAtStart:  Call this subroutine from an "auto_open" macro or from the "Workbook_Open" event for "ThisWorkbook"
Public Sub CheckForUpdatesAtStart()
    If (GetUpdateAtStart) Then CheckWebVersion
End Sub

Private Sub CheckWebVersion()
    Dim MsgBoxButtons As VbMsgBoxStyle
    Dim AddMessage As String, msgTmp As String
    Dim WebVersion As Double
    Dim resp As VbMsgBoxResult
    ' Settings for MsgBox depending on whether a checkbox is used to enable automatic update checks
    If (UseCheckBox Or UseRibbonCheckBox) Then
        MsgBoxButtons = vbYesNoCancel
        AddMessage = Chr(10) & "To turn off this reminder, select 'Cancel'"
    Else
        MsgBoxButtons = vbYesNo
        AddMessage = ""
    End If
    
    ' Fetch the latest version number from the web
    Application.StatusBar = "Checking for updates to " & MsgBoxTitle & "..."
    WebVersion = GetWebVersion()
    Application.StatusBar = False
    
    If (WebVersion > CurrentVersion) Then
        If UseInstaller Then
            msgTmp = "Would you like to update now? The update will be installed once you close Excel."
        Else
            msgTmp = "Would you like to go to the website to download the new version now?"
        End If
        resp = msgLogDisp("A new version of " & MsgBoxTitle & " is available." & Chr(10) & Chr(10) & msgTmp & Chr(10) & AddMessage, MsgBoxButtons + vbQuestion, MsgBoxTitle, vbNo)
        If (resp = vbYes) Then
            If UseInstaller Then
                ' Run installation sub
                RunInstaller
            Else
                ' Open a browser
                Dim objshell As WshShell
                Set objshell = New WshShell
                objshell.Run (DownloadUrl)
            End If
        ElseIf (resp = vbCancel) Then
            ' Disable checkbox
            SetUpdateAtStart (False)
            If UseRibbonCheckBox Then
                'ResetRibbon
            End If
        End If
    End If

End Sub

Private Sub RunInstaller()
    Dim thisAddinPath As String, linkDownloaderFullPath As String, args As String
    ' Check if LinkDownloader is already running
    If CheckProcessRunning("LinkDownloader_tmp.exe") > 0 Then
        msgLogDisp "LinkDownloader is already running and waiting for Excel to exit. After closing Excel, check the task manager to make sure all instances of Excel have shutdown completely.", vbExclamation, MsgBoxTitle
        Exit Sub
    End If

    ' Get the path to this file
    thisAddinPath = Application.ThisWorkbook.Path
    'Added by TCS offshore (#54)
    'Check if the file is exist or not
    If Dir(thisAddinPath & "\LinkDownloader.exe") = "" Then
        msgLogDisp "LinkDownloader.exe does not exist in " & thisAddinPath & vbCrLf & "Automatic update cannot be performed. Please manually download the latest version of " & MsgBoxTitle & ", and make sure to put the LinkDownloader.exe file in the same directory.", vbExclamation, MsgBoxTitle
        End
    End If
        
    ' Copy the installer executable to a temp file; because the update will overwrite the installer
    ' Before copying, delete prior temp file if it exists
    Dim oFso As FileSystemObject
    Set oFso = New FileSystemObject
    On Error Resume Next
    oFso.DeleteFile thisAddinPath & "\LinkDownloader_tmp.exe", True
    On Error GoTo CopyError
    oFso.CopyFile thisAddinPath & "\LinkDownloader.exe", thisAddinPath & "\LinkDownloader_tmp.exe"
    On Error GoTo 0
    Set oFso = Nothing
    
    ' Set up command to perform download using installer
    linkDownloaderFullPath = """" & thisAddinPath & "\" & "LinkDownloader_tmp.exe"""
    args = " -z -u """ & DownloadUrl & """ -l """ & thisAddinPath & """ -p Excel"
    
    ' use COM object to run downloader
    On Error GoTo DownloadError:
    Dim objshell As WshShell
    Set objshell = New WshShell
    objshell.Run (linkDownloaderFullPath & args)
    Set objshell = Nothing
    On Error GoTo 0
    Exit Sub
    
CopyError:
    msgLogDisp "Error making copy of LinkDownloader. Please check if LinkDownloader is already running.", vbExclamation, MsgBoxTitle
    Exit Sub
    
DownloadError:
    msgLogDisp "Error running LinkDownloader. Try completely closing Excel and any copies of LinkDownloader. Otherwise, try manually downloading an update.", vbExclamation, MsgBoxTitle
    Exit Sub
    
End Sub


Function GetWebVersion() As Double
    Dim Options As String, UserName As String, VersionUrlFull As String
    Dim success As Boolean
    Dim result As Object
    GetWebVersion = 0
    
    Dim WinHttpReq As WinHttpRequest
    Set WinHttpReq = New WinHttpRequest

    ' Set request options for tracking
    Options = "?currentVersion=" & CurrentVersion
    UserName = (Environ$("Username"))
    If (UserName <> "") Then
        Options = Options & "&user=" & UserName
    End If
    VersionUrlFull = VersionFileUrl & Options

    ' Send the HTTP Request.
    On Error GoTo TheEnd
    WinHttpReq.Open "GET", VersionUrlFull, True
    WinHttpReq.SetAutoLogonPolicy (0)
    WinHttpReq.SetTimeouts 5000, 5000, 5000, 5000
    WinHttpReq.send
    success = WinHttpReq.WaitForResponse(5)
    If Not (success) Then
        Exit Function
    End If
    
    Set result = JsonConverter.ParseJson(WinHttpReq.ResponseText)
    GetWebVersion = Val(result("result")("version"))
    DownloadUrl = result("result")("downloadUrl")

    ' Write to document properties
    If (GetWebVersion > 0 And GetWebVersion <> ThisWorkbook.CustomDocumentProperties("WebVer").Value) Then
        ThisWorkbook.CustomDocumentProperties("WebVer").Value = GetWebVersion
    End If
    
TheEnd:

End Function

'------------------------------------
' External calls to get version information
'------------------------------------
Function GetCurrentVersionVal() As Double
    GetCurrentVersionVal = CurrentVersion
End Function

Function GetWebVersionVal() As Double
    On Error Resume Next
    GetWebVersionVal = 0
    GetWebVersionVal = ThisWorkbook.CustomDocumentProperties("WebVer").Value
    On Error GoTo 0
End Function

Private Sub TestVersions()
    Debug.Print "Current: " & GetCurrentVersionVal
    Debug.Print "Latest: " & GetWebVersionVal
    Debug.Print "Update: " & GetUpdateAtStart
End Sub

'------------------------------------
' External calls to get/set update at start status
'------------------------------------
Function GetUpdateAtStart() As Boolean
    If (UseCheckBox) Then
        GetUpdateAtStart = Worksheets(CheckBoxWorksheet).Shapes("UpdateCheckbox").ControlFormat.Value
    ElseIf (UseRibbonCheckBox) Then
        GetUpdateAtStart = ThisWorkbook.CustomDocumentProperties("UpdateCheck").Value
    Else
        GetUpdateAtStart = True
    End If
End Function

Sub SetUpdateAtStart(ByVal MyVal As Boolean)
    If (UseCheckBox) Then
        Worksheets(CheckBoxWorksheet).Shapes("UpdateCheckbox").ControlFormat.Value = MyVal
    ElseIf (UseRibbonCheckBox) Then
        ThisWorkbook.CustomDocumentProperties("UpdateCheck").Value = MyVal
        ThisWorkbook.Save
    End If
End Sub

'------------------------------------
' Add Custom Document Properties
'
' These only need to be run once in a new Workbook
'------------------------------------

Private Sub AddWebVerCustomProperty()
    ThisWorkbook.CustomDocumentProperties.Add Name:="WebVer", LinkToContent:=False, Value:=0, Type:=msoPropertyTypeNumber
End Sub

Private Sub AddUpdateCheckCustomProperty()
    ThisWorkbook.CustomDocumentProperties.Add Name:="UpdateCheck", LinkToContent:=False, Value:=0, Type:=msoPropertyTypeBoolean
End Sub

' Check for running process
Function CheckProcessRunning(process As String) As Long
    Dim objList As Object
    Set objList = GetObject("winmgmts:").ExecQuery("select * from win32_process where name='" & process & "'")
    CheckProcessRunning = objList.Count
End Function

' Check for multiple Excel instances which prevent LinkDownloader from running
Sub CheckHangingExcel()
    If CheckProcessRunning("excel.exe") > 1 Then
        msgLogDisp "There is currently more than one instance of Excel running. Excel may not be completely shutdown. Check the task manager after Excel finishes closing.", vbExclamation, MsgBoxTitle
    End If
End Sub
