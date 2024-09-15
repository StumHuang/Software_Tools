Attribute VB_Name = "ApiTest"
Option Explicit

'Valid root hives:
'HKEY_CLASSES_ROOT   (2147483648)
'HKEY_CURRENT_USER   (2147483649)
'HKEY_LOCAL_MACHINE  (2147483650)
'HKEY_USERS          (2147483651)
'HKEY_CURRENT_CONFIG (2147483653)
Function ReadRegStr(RootKey, Key, RegType) As String
    Dim oCtx, oLocator, oReg, oInParams, ReadReg
    Set oCtx = CreateObject("WbemScripting.SWbemNamedValueSet")
    oCtx.Add "__ProviderArchitecture", RegType
    Set oLocator = CreateObject("Wbemscripting.SWbemLocator")
    Set oReg = oLocator.ConnectServer("", "root\default", "", "", , , , oCtx).Get("StdRegProv")
    Set oInParams = oReg.Methods_("GetStringValue").InParameters
    oInParams.hDefKey = RootKey
    oInParams.sSubKeyName = Key
    ReadReg = oReg.ExecMethod_("GetStringValue", oInParams, , oCtx).sValue
    ReadRegStr = IIf(IsNull(ReadReg), "", ReadReg)
End Function

Function CheckRegKeys() As Boolean
    Const HKEY_CLASSES_ROOT = 2147483648#
    Dim Key1 As String, Key2 As String, Key3 As String, Key4 As String, Key5 As String
    Dim b
    'try 32 bit keys first, if not found try 64 bit
    For Each b In Array(32, 64)
        Key1 = ReadRegStr(HKEY_CLASSES_ROOT, "CLSID\{5A0C649C-7E51-4862-88F4-E4FD493EF2D8}\LocalServer32\", b)
        Key2 = ReadRegStr(HKEY_CLASSES_ROOT, "CLSID\{45C15B10-23E1-406E-B50E-11C5649739C0}\LocalServer32\", b)
        Key3 = ReadRegStr(HKEY_CLASSES_ROOT, "CLSID\{0E8CFE06-C139-4858-B816-F78FD36CB5B8}\LocalServer32\", b)
        Key4 = ReadRegStr(HKEY_CLASSES_ROOT, "CLSID\{EEB19044-ABF5-4231-A6A2-0B4062C99B89}\LocalServer32\", b)
        Key5 = ReadRegStr(HKEY_CLASSES_ROOT, "CLSID\{CDD3FF3C-D343-4C20-9C58-C67B0B2AA10D}\LocalServer32\", b)
        If Not ((InStr(LCase(Key1), "vision.exe") > 0) And (Key1 = Key2) And (Key1 = Key3) And (Key1 = Key4) And (Key1 = Key5)) Then
            CheckRegKeys = False
        Else
            CheckRegKeys = True
            Exit For
        End If
    Next b
End Function

Sub TestAndFixApi2()
    Dim PathToVision As Variant
    Dim CommandLine As String
    Dim resp As VbMsgBoxResult
    Dim HelpDocUrl As String
    Dim objshell As Object
    
    If Not CheckRegKeys Then
        msgLogDisp "There appears to be a problem with the COM interface to Vision." _
            & Chr(10) & "After clicking okay, please specify which copy of Vision you are using by locating the Vision.exe file.", vbExclamation
        If (FileOrDirExists("C:\Program Files (x86)\Accurate Technologies\")) Then
            ChDir ("C:\Program Files (x86)\Accurate Technologies\")
        ElseIf (FileOrDirExists("C:\Program Files\Accurate Technologies\")) Then
            ChDir ("C:\Program Files\Accurate Technologies\")
        Else
            ChDir ("C:\")
        End If
        If Not AutomatedMode Then PathToVision = Application.GetOpenFilename("Vision.exe (*.exe),*.exe") Else PathToVision = False
        If (PathToVision = False) Then
            msgLogDisp "It is unlikely that this macro will be able to successfully communitate with Vision.  You may see an error #429 when you try to run the script.", vbExclamation
        Else
            ChDir (FileParts(PathToVision, "path"))
            CommandLine = "RegisterCOMInterface.bat"
            Shell CommandLine
            Application.Wait (Now() + TimeValue("00:00:10"))
            ActivateExcel
            
            ' Check Again
            If Not CheckRegKeys Then
                resp = msgLogDisp("COM interface is still not working properly. Please review the API troubleshooting document available on the GCMT for further help. Would you like to open that document now?", vbExclamation + vbYesNo, "Test Vision API", vbNo)
                If (resp = vbYes) Then
                    ' Open a browser
                    HelpDocUrl = "https://azureford.sharepoint.com/sites/GlobalCalibrationMethodologyTeam/tools/Shared%20Documents/Excel%20Macros/Excel%20Vision%20Integration/Fixing%20ATI%20Vision%20API%20Macro%20Issues.doc"
                    Set objshell = CreateObject("Wscript.Shell")
                    objshell.Run (HelpDocUrl)
                End If
            Else
                msgLogDisp "COM interface is now working properly.", vbInformation
            End If
        End If
    End If
End Sub
