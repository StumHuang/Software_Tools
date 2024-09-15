Attribute VB_Name = "BuildVST"
Option Explicit

'Create the ActiveX object
Private Strategy As Object
Private Ret As VISION_STATUS_CODES
Private QuitNow As Boolean

Function BuildVSTFile(BatchProcess As Boolean, ByRef ShowMisMatchError As Boolean, Optional FileNames As Variant) As String
    Dim MultipleFilesFlag As Boolean
    Dim AtiFileCell As Range, H32FileCell As Range, VSTFileCell As Range, AddStatesCell As Range
    Dim AtiFilePath As String, H32FilePath As String, VSTFilePath As String
    Dim n As Long, FileNum As Long
    ReDim VSTFilePaths(0 To 0) As String
    Dim H32FilePaths As Variant
    
    MultipleFilesFlag = IsArray(FileNames)

    ' Global flag to quit
    QuitNow = False
    
    On Error Resume Next
    Set Strategy = CreateObject("Vision.StrategyFileInterface") 'Late binding
    If Err.Number <> 0 Then
        QuitNow = True
        GoTo CleanUp
    End If
    On Error GoTo 0
    
      
    ' Define Ranges
    Set AtiFileCell = Worksheets("File Paths").Range("B2")
    Set H32FileCell = Worksheets("File Paths").Range("B3")
    Set VSTFileCell = Worksheets("File Paths").Range("B4")
    Set AddStatesCell = Worksheets("File Paths").Range("B8")
    
    If MultipleFilesFlag Then
        AtiFilePath = FileNames(0)
    Else
        ' Choose Files
        AtiFilePath = GetFilePath(AtiFileCell, "Strategy Description File (*.ati;*.a2l),*.ati;*.a2l", "")
    End If
    If QuitNow Then GoTo CleanUp
    If Not BatchProcess Then
        ' One at a time
        'Check if commandline arguements are provided
        If MultipleFilesFlag Then 'vst build from external caller
            H32FilePath = FileNames(2)
            VSTFilePath = FileNames(3)
        Else 'normal vst build
            H32FilePath = GetFilePath(H32FileCell, "Memory Image File (*.h32;*.hex;*.s19;*.s37;*.mot),*.h32;*.hex;*.s19;*.s37;*.mot", AtiFilePath)
            If (QuitNow) Then GoTo CleanUp
            VSTFilePath = GetFilePath(VSTFileCell, "ATI Strategy File (*.vst),*.vst", H32FilePath, FileParts(H32FilePath, "nameonly"))
            If (QuitNow) Then GoTo CleanUp
        End If
        H32FilePaths = Array(H32FilePath)
        VSTFilePaths(0) = VSTFilePath
    Else
        If MultipleFilesFlag Then 'batch build from spreadsheet tab
            ReDim H32FilePaths(0) As String
            H32FilePaths(0) = FileNames(2)
            ReDim VSTFilePaths(0) As String
        Else 'manual batch build
            msgLogDisp "In batch mode, each VST file will be named to match the H32 file.", vbInformation
            H32FilePaths = GetFilePaths(H32FileCell, "Memory Image File (*.h32;*.hex;*.s19;*.s37;*.mot),*.h32;*.hex;*.s19;*.s37;*.mot", AtiFilePath)
            If (QuitNow) Then GoTo CleanUp
            ReDim VSTFilePaths(LBound(H32FilePaths) To UBound(H32FilePaths)) As String
        End If
        
        For n = LBound(H32FilePaths) To UBound(H32FilePaths)
            VSTFilePaths(n) = FileParts(H32FilePaths(n), "path") & FileParts(H32FilePaths(n), "nameonly") & ".vst"
        Next n
    End If
    
    ' Progress Box
    Progress.Label1.caption = "Initializing..."
    Progress.Label2.caption = ""
    Progress.Show
    Application.Wait (Now + TimeValue("0:00:02"))

    ' Run AddStates if desired
    If (AddStatesCell.Value <> "") Then
        RunAddStates AddStatesCell.Value, AtiFilePath
    End If
    
    ' Strategy Description
    ImportDesc AtiFilePath
    If (QuitNow) Then GoTo CleanUp
    AddMemoryRegions Strategy, ShowMisMatchError, BatchProcess
    
    ' Run Template
    Progress.Label1.caption = "Applying Template..."
    Progress.Repaint
    If MultipleFilesFlag Then
        VSTParameterTemplate ShowMisMatchError, Strategy, AtiFilePath, BatchProcess, FileNames
    Else
        VSTParameterTemplate ShowMisMatchError, Strategy, AtiFilePath, BatchProcess
    End If
    QuitNow = False ' In case it got set while loading a MAP file during VSTParameterTemplate
    
    ' Loop through all H32 files and create VST files
    For FileNum = LBound(H32FilePaths) To UBound(H32FilePaths)
        H32FilePath = H32FilePaths(FileNum)
        VSTFilePath = VSTFilePaths(FileNum)
    
        ' Memory Image
        ImportMemory H32FilePath
        If (QuitNow) Then GoTo CleanUp
        
        ' Add Metadata from H32
        AddH32Meta H32FilePath, Strategy
    
        ' Cal Changes (newer Vision versions)
        If (GetNumericVisionApiVersion >= 3008002) Then
            UpdateCalibrationValues Strategy, AtiFilePath
        End If
        If (QuitNow) Then Exit For
    
        'Update the state variable colour
        changecolour Strategy, BatchProcess, ShowMisMatchError
    
        ' Save
        SaveVST VSTFilePath
        If (QuitNow) Then GoTo CleanUp
        
        ' Make calibration changes after saving for older Vision versions
        If (GetNumericVisionApiVersion < 3008002) Then
            DoCalChanges VSTFilePath
        End If
        
        ' Add new VST to Vision if desired
        AddToTree VSTFilePath
    Next FileNum
    
'----------------------------------
' Clean Up
CleanUp:
    Set Strategy = Nothing
    If (QuitNow) Then
        Progress.Label1.caption = "Failed!"
        debugLog = debugLog & "Failed!" & vbCrLf & "WScriptOutputAsFailed"
    Else
        Progress.Label1.caption = "Done!"
        debugLog = debugLog & "Done!" & vbCrLf & "WScriptOutputAsSuccess"
    End If
    Progress.Label2.caption = ""
    Progress.Repaint
    Application.Wait (Now + TimeValue("0:00:02"))
    Progress.Hide
    BuildVSTFile = debugLog
End Function

Sub SaveVST(ByVal VSTFilePath As String)
    Progress.Label1.caption = "Saving New VST..."
    Progress.Label2.caption = FileParts(VSTFilePath, "filename")
    Progress.Repaint
    Ret = Strategy.SaveAs(VSTFilePath)
    If (Ret <> VISION__OK) Then
        Progress.Hide
        ActivateExcel
        msgLogDisp "Error saving new VST file."
        QuitNow = True
    End If
End Sub

Function GetFilePath(ByVal SaveCell As Range, ByVal FilterString As String, ByVal PreviousPath As String, Optional DefaultName As String = "") As String
    ' Change to saved directory and open dialog
    If (FileOrDirExists(SaveCell.Value)) Then
        ChDir SaveCell.Value
    End If
    If (DefaultName = "") Then
        GetFilePath = Application.GetOpenFilename(FilterString)
    Else
        GetFilePath = Application.GetSaveAsFilename(DefaultName, FilterString, 1, "Save VST File")
    End If
    
    ' If they cancelled then set global quit flag
    If GetFilePath = "False" Then
        QuitNow = True
        Exit Function
    End If
    
    ' Write path back to spreadsheet
    If (StrComp(FileParts(GetFilePath, "path"), FileParts(PreviousPath, "path")) = 0) Then
        SaveCell.Value = ""
    Else
        SaveCell.Value = FileParts(GetFilePath, "path")
    End If
End Function

Function GetFilePaths(ByVal SaveCell As Range, ByVal FilterString As String, ByVal PreviousPath As String) As Variant
    Dim AddMore As Boolean, OkToAdd As Boolean
    Dim ResultsTmp As Variant, Results As Variant
    Dim FirstPath As String
    Dim resp As VbMsgBoxResult
    AddMore = True
    Do While (AddMore)
        ' Change to saved directory and open dialog
        If (FileOrDirExists(SaveCell.Value)) Then
            ChDir SaveCell.Value
        End If
        ResultsTmp = Application.GetOpenFilename(FilterString, 1, "Select Multiple Files", "", True)
        
        ' If they cancelled then set global quit flag
        OkToAdd = True
        If Not IsArray(ResultsTmp) Then
            If (ResultsTmp = "False") Then
                ' Not an array and false: must have been cancel button
                OkToAdd = False
                If (IsEmpty(Results)) Then
                    ' Cancelled before selecting anything; then quit. Otherwise will continue to ask for more
                    QuitNow = True
                    Exit Function
                End If
            End If
        End If
        
        ' Concactenate results
        If OkToAdd Then
            If (IsEmpty(Results)) Then
                Results = ResultsTmp
            Else
                Results = Split(Join(Results, Chr(1)) & Chr(1) & Join(ResultsTmp, Chr(1)), Chr(1))
            End If
        End If
        
        ' Ask if they have more
        resp = msgLogDisp("Select more H32 Files?", vbQuestion + vbYesNoCancel, "VST Tool: Select H32 Files", vbNo)
        If resp = vbNo Then
            AddMore = False
        ElseIf resp = vbCancel Then
            QuitNow = True
            Exit Function
        End If
        
    Loop
        
    ' Write path back to spreadsheet
    FirstPath = Results(1)
    If (StrComp(FileParts(FirstPath, "path"), FileParts(PreviousPath, "path")) = 0) Then
        SaveCell.Value = ""
    Else
        SaveCell.Value = FileParts(FirstPath, "path")
    End If

    GetFilePaths = Results
End Function

Sub ImportDesc(ByVal AtiFilePath As String)
    Dim StrategyPreset As String, GroupSeparatorString As String
    Dim ImportFunctions As Boolean, SwapAxes As Boolean, IgnoreMemoryRegions As Boolean, UseExtendedLimits As Boolean, EnforceLimits As Boolean, DeleteExistingItems As Boolean, ReplaceExistingItems As Boolean, ClearDeviceSettings As Boolean, AllowBrackets As Boolean, OrganizeDataItemsInGroups As Boolean, UseDisplayIdentifiers As Boolean
    Dim StructureNameOption As Long
    
    Progress.Label1.caption = "Importing Strategy Description File..."
    Progress.Repaint
    
    ' Import Options
    AllowBrackets = Worksheets("A2L Import Settings").Cells(12, 2).Value
    ClearDeviceSettings = Worksheets("A2L Import Settings").Cells(9, 2).Value
    DeleteExistingItems = Worksheets("A2L Import Settings").Cells(7, 2).Value
    ImportFunctions = Worksheets("A2L Import Settings").Cells(3, 2).Value
    ReplaceExistingItems = Worksheets("A2L Import Settings").Cells(8, 2).Value
    SwapAxes = Worksheets("A2L Import Settings").Cells(4, 2).Value
    UseDisplayIdentifiers = Worksheets("A2L Import Settings").Cells(14, 2).Value
    StructureNameOption = Worksheets("A2L Import Settings").Cells(10, 2).Value
    StrategyPreset = Worksheets("A2L Import Settings").Cells(2, 2).Value
    IgnoreMemoryRegions = Worksheets("A2L Import Settings").Cells(5, 2).Value
    UseExtendedLimits = Worksheets("A2L Import Settings").Cells(6, 2).Value
    OrganizeDataItemsInGroups = Worksheets("A2L Import Settings").Cells(13, 2).Value
    GroupSeparatorString = Worksheets("A2L Import Settings").Cells(11, 2).Value
    EnforceLimits = Worksheets("A2L Import Settings").Cells(15, 2).Value
    
    Ret = -1
    On Error Resume Next
    Ret = Strategy.SetASAP2ImportProperties2(StrategyPreset, ImportFunctions, SwapAxes, IgnoreMemoryRegions, UseExtendedLimits, EnforceLimits, DeleteExistingItems, ReplaceExistingItems, ClearDeviceSettings, AllowBrackets, OrganizeDataItemsInGroups, UseDisplayIdentifiers, StructureNameOption, GroupSeparatorString)
    On Error GoTo 0
    If (Ret <> VISION__OK) Then
        Ret = Strategy.SetASAP2ImportProperties(AllowBrackets, ClearDeviceSettings, DeleteExistingItems, ImportFunctions, ReplaceExistingItems, SwapAxes, UseDisplayIdentifiers, StructureNameOption, StrategyPreset, IgnoreMemoryRegions, UseExtendedLimits, OrganizeDataItemsInGroups, GroupSeparatorString)
        If (Ret <> VISION__OK) Then
            Progress.Hide
            ActivateExcel
            msgLogDisp "Error setting A2L import properties."
            Exit Sub
        End If
    End If
    
    Ret = Strategy.SetASAP2ImportProperties2(StrategyPreset, ImportFunctions, SwapAxes, IgnoreMemoryRegions, UseExtendedLimits, EnforceLimits, DeleteExistingItems, ReplaceExistingItems, ClearDeviceSettings, AllowBrackets, OrganizeDataItemsInGroups, UseDisplayIdentifiers, StructureNameOption, GroupSeparatorString)
    
    Ret = Strategy.Import(AtiFilePath)
    If (Ret <> VISION__OK) Then
        Progress.Hide
        ActivateExcel
        msgLogDisp "Error importing strategy description file."
        Exit Sub
    End If
End Sub

Sub RunAddStates(ByVal AddStatesPath As String, ByRef AtiFilePath As String)
    Dim NewAtiFilePath As String, OrigDirectory As String, CommandLine As String
    
    Progress.Label1.caption = "Running AddStates Script..."
    Progress.Repaint
    NewAtiFilePath = FileParts(AtiFilePath, "path") & FileParts(AtiFilePath, "nameonly") & "_state." & FileParts(AtiFilePath, "ext")
    If (FileOrDirExists(NewAtiFilePath & "")) Then
        Progress.Label1.caption = "New A2L Exists...skipping..."
        Progress.Repaint
        AtiFilePath = NewAtiFilePath
    Else
        If (FileOrDirExists(AddStatesPath)) Then
            OrigDirectory = CurDir
            ChDir (FileParts(AddStatesPath, "path"))
            CommandLine = """" & AddStatesPath & """ """ & AtiFilePath & """"
            Dim WshShell As WshShell
            Set WshShell = New WshShell
            WshShell.Run CommandLine, 6, True
            If Not (FileOrDirExists(NewAtiFilePath & "")) Then
                    Progress.Hide
                    ActivateExcel
                    msgLogDisp "AddStates did not complete successfully. Continuing with original A2L file.", vbExclamation
                    Progress.Show
            Else
                AtiFilePath = NewAtiFilePath
            End If
            ChDir (OrigDirectory)
        Else
            Progress.Hide
            ActivateExcel
            msgLogDisp "Could not find AddStates script."
            Progress.Show
        End If
    End If
End Sub

Sub ImportMemory(ByVal H32FilePath As String)
    Progress.Label1.caption = "Setting memory import properties..."
    Progress.Repaint
    
    Dim RegionNames(0) As Variant
    RegionNames(0) = vbNull
    
    Ret = Strategy.SetHexImportProperties(1, 0, 0, &HFFFFFFFF, RegionNames)
    If (Ret <> VISION__OK) Then
        Progress.Hide
        ActivateExcel
        msgLogDisp "Error setting HEX import properties."
        Exit Sub
    End If
    
    Ret = Strategy.SetSRecordImportProperties(1, 0, 0, &HFFFFFFFF, RegionNames, 1)
    If (Ret <> VISION__OK) Then
        Progress.Hide
        ActivateExcel
        msgLogDisp "Error setting S-Record import properties."
        Exit Sub
    End If
    
    Progress.Label1.caption = "Importing H32 File..."
    Progress.Repaint
    Ret = Strategy.Import(H32FilePath)
    If (Ret <> VISION__OK) Then
        Progress.Hide
        ActivateExcel
        msgLogDisp "Error importing H32 file."
        Exit Sub
    End If
End Sub

Sub DoCalChanges(ByVal VSTFilePath As String)
    Dim msgret As VbMsgBoxResult
    Dim PCMDevice As Object
    Dim FailedToActivate As Boolean
    Dim RowNum As Long, i As Long
    Dim WorkSheetName As String, SettingName As String, missing_string As String
    Dim SettingValue As Double
    
    If (Worksheets("Cal Changes").Cells(2, 1).Value <> "") Then
        Progress.Label1.caption = "Modifying Calibration Values"
        Progress.Label2.caption = ""
        Progress.Repaint
            
        '' Data bins for missing parameters
        Dim misscount As Long
        Dim missingparams() As String
        misscount = 0
            
        '' Open Project
        Dim Project As Object
        Set Project = CreateObject("Vision.ProjectInterface")
        If (Project.IsOpen = False) Then
            Progress.Hide
            ActivateExcel
            msgLogDisp "No project appears to be open within Vision." & Chr(10) & "Calibration changes cannot be made.", vbOKOnly + vbCritical
            Progress.Show
        Else
            ' Warning
            If (Project.Online = True) Then
                Progress.Hide
                ActivateExcel
                msgret = msgLogDisp("Warning:  The VST Tool macro is about to take your current PCM device off line.  This will interrupt data acquistion if any is in progress.  Press cancel if you wish to skip calibration changes so that data acquistion may continue.", vbOKCancel + vbExclamation, , vbCancel)
                Progress.Show
                If (msgret = vbCancel) Then
                    Set Project = Nothing
                    Exit Sub
                End If
            End If
                        
            ' Open PCM
            Set PCMDevice = Project.FindDevice("PCM")
            If (PCMDevice Is Nothing) Then
                Progress.Hide
                ActivateExcel
                msgLogDisp "The device named ""PCM"" could not be found.  Please ensure your project has a PCM device and try again.", vbOKOnly + vbCritical
                Progress.Show
            Else
                PCMDevice.AddStrategy (VSTFilePath)
                PCMDevice.ActiveStrategy = VSTFilePath
                
                ' Check if it worked
                FailedToActivate = False
                If (PCMDevice.ActiveStrategy Is Nothing) Then
                    FailedToActivate = True
                ElseIf (PCMDevice.ActiveStrategy.Filename <> VSTFilePath) Then
                    FailedToActivate = True
                End If
                
                If (FailedToActivate) Then
                    Progress.Hide
                    ActivateExcel
                    msgLogDisp "Could not set new VST as the active strategy in the current project.  Calibration changes cannot be made.", vbCritical
                    Progress.Show
                    Set PCMDevice = Nothing
                    Set Project = Nothing
                    Exit Sub
                End If
                
                ' Make cal changes
                Dim DataItem As Object
                
                RowNum = 2
                WorkSheetName = "Cal Changes"
                Do While (Worksheets(WorkSheetName).Cells(RowNum, 1).Value <> "")
                    SettingName = Worksheets(WorkSheetName).Cells(RowNum, 1).Value
                    SettingValue = Worksheets(WorkSheetName).Cells(RowNum, 2).Value
            
                    Progress.Label2.caption = SettingName
                    Progress.Repaint
            
                    Set DataItem = PCMDevice.FindDataItem(SettingName)
                    If (Not (DataItem Is Nothing)) Then
                        DataItem.TargetValue = SettingValue
                    Else
                        ReDim Preserve missingparams(misscount)
                        missingparams(misscount) = SettingName
                        misscount = misscount + 1
                    End If
                    
                    RowNum = RowNum + 1
                Loop
                
                ' Message box for any missing parameters
                If (misscount > 0) Then
                    Progress.Label1.caption = "Please respond to open message box..."
                    Progress.Label2.caption = ""
                    Progress.Repaint
                    
                    missing_string = missingparams(0)
                    If (misscount > 1) Then
                        For i = 1 To (misscount - 1)
                            missing_string = missing_string & Chr(10) & missingparams(i)
                        Next
                    End If
                    ActivateExcel
                    msgLogDisp "The following parameters were not found:" & Chr(10) & Chr(10) & missing_string, vbExclamation
                End If
                
                Progress.Label1.caption = "Re-saving VST..."
                Progress.Label2.caption = FileParts(VSTFilePath, "filename")
                Progress.Repaint
                PCMDevice.ActiveStrategySaveAs (VSTFilePath)
                PCMDevice.RemoveStrategy (VSTFilePath)
                Set PCMDevice = Nothing
                Set Project = Nothing
                
            End If ' PCM is found
        End If ' project is open
    End If ' sheet is not blank
End Sub

Sub AddToTree(ByVal VSTFilePath As String)
    Dim PCMDevice As Object
    
    If (Worksheets("Other Settings").Shapes("AddToTreeCheckBox").ControlFormat.Value = 1) Then
        '' Open Project
        Dim Project As Object
        Set Project = CreateObject("Vision.ProjectInterface")
        If (Project.IsOpen = False) Then
            Exit Sub
        Else
            ' Open PCM
            Set PCMDevice = Project.FindDevice("PCM")
            If (PCMDevice Is Nothing) Then
                Progress.Hide
                ActivateExcel
                msgLogDisp "The device named ""PCM"" could not be found.  Please ensure your project has a PCM device and try again.", vbOKOnly + vbCritical
                Progress.Show
            Else
                PCMDevice.AddStrategy (VSTFilePath)
            End If ' PCM is found
        End If ' project is open
    End If ' sheet is not blank
End Sub
