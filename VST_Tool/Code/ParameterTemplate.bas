Attribute VB_Name = "ParameterTemplate"
Option Explicit

Public Enum VISION_STATUS_CODES
    VISION__OK = 0                  'Same as VISION_OK, but will sort to top of list in VB.
    VISION_DUPLICATE_NAME = 4
    VISION_ERROR = 1                'General error code.
    VISION_FILE_CREATE_FAIL = 1004  'Creation of new file failed
    VISION_FILE_OPEN_FAIL = 1000    'Failed to open file
    VISION_FILE_READ_FAIL = 1002    'Read from file failed
    VISION_FILE_SAVE_FAIL = 1001    'Failed to save file
    VISION_FILE_WRITE_FAIL = 1003   'Write to file failed
    VISION_INVALID_PARM = 2         'Invalid parameter.
    VISION_ITEM_NOT_FOUND = 3       'Item being searched for not found.
    VISION_LICENSING_ERROR = 5
    VISION_MEMORY_NOT_FOUND = 6
    VISION_NOT_INITIALIZED = 10000  'Vision system not initialized (no project open).
    VISION_OK = 0                   'Operation completed successfully.
End Enum

Sub VSTParameterTemplate(ByRef ShowMisMatchError As Boolean, Optional ExistingHandle As Object = Nothing, Optional AtiFilePath As String = "", Optional BatchMode As Boolean = False, Optional FileNames As Variant)
    Dim MultipleFilesFlag As Boolean
    Dim sFilePath As Variant
    Dim RowNum As Long, GroupNum As Long, ColNum As Long, i As Long, Processor As Long
    Dim WorkSheetName As String, ParamName As String, ParamEq As String, Comment As String, ReplacementParameter As String, Formula As String
    Dim temphandle As Object, tempGroupHandle As Object, temphandle2 As Object, matches As Object
    Dim splitedata() As String
    Dim DataA As Double, DataB As Double, DataC As Double
    Dim DefaultDecimalCell As Range, DefaultMapYDecimalCell As Range, DefaultMapXDecimalCell As Range, DefaultCurveXDecimalCell As Range, Rng1 As Range, AtiFileCell As Range, MapFileCell As Range
    Dim missing_string As String, MAPFilePath As String, NvramStart As String, NvramEnd As String, NvramLength As String, VSTFilePath As String
    
    'On Error GoTo CleanUp
    MultipleFilesFlag = Not IsMissing(FileNames)
    If (ExistingHandle Is Nothing) Then
    
        ' Choose which VST
        sFilePath = Application.GetOpenFilename("VISION Strategy files (*.vst),*.vst")
        If (sFilePath = "False") Then
            Exit Sub
        End If
        
        ' Progress Box
        Progress.Label1.caption = "Initializing..."
        Progress.Label2.caption = ""
        Progress.Show
        Application.Wait (Now + TimeValue("0:00:02"))
    
        'Create the ActiveX object
        Dim Strategy As Object
        Set Strategy = CreateObject("Vision.StrategyFileInterface")   'Late binding
        
        Dim Ret As VISION_STATUS_CODES
        Ret = Strategy.Open(sFilePath)
        If (Ret <> VISION__OK) Then
            ActivateExcel
            msgLogDisp "Error opening strategy file."
            Exit Sub
        End If
        
        ' Add memory regions if desired.  On new VSTs, this will get called within BuildVSTFile
        AddMemoryRegions Strategy, ShowMisMatchError, BatchMode

    Else
        ' We were passed an existing handle from another sub
        Set Strategy = ExistingHandle
    End If
    
    ' Add dummy parameters
    Progress.Label1.caption = "Add Parameters"
    Progress.Label2.caption = ""
    Progress.Repaint
    RowNum = 2
    WorkSheetName = "Added Parameters"
    Dim MeasurementGroup As Object
    
    For GroupNum = 1 To Strategy.GroupDataItem.Items.Count
        If (Strategy.GroupDataItem.Items(GroupNum).DataItemName = "Measurements") Then
            Set MeasurementGroup = Strategy.GroupDataItem.Items(GroupNum)
            Exit For
        End If
    Next
    
    Do While (Worksheets(WorkSheetName).Cells(RowNum, 1).Value <> "")
        ParamName = Trim(Worksheets(WorkSheetName).Cells(RowNum, 1).Value)
        ParamEq = Worksheets(WorkSheetName).Cells(RowNum, 2).Value
        Comment = Worksheets(WorkSheetName).Cells(RowNum, 3).Value
        Progress.Label2.caption = ParamName
        If (ParamEq = "DELETE") Then
            Set temphandle = Strategy.FindOrSearchDataItem2(ParamName, False)
            If Not (temphandle Is Nothing) Then
                ' remove it
                Set tempGroupHandle = FindItemsGroup(Strategy.GroupDataItem, ParamName)
                tempGroupHandle.RemoveItem (ParamName)
            End If
        Else
            If (MeasurementGroup Is Nothing) Then
                GoTo ContinueDo
            End If
            Set temphandle = MeasurementGroup.FindDataItem(ParamName)
            If (temphandle Is Nothing) Then
                ' Create new parameter if...
                '     equation starts with =
                '     or contains a comma. But there either must not be a ( or it must be after the comma.
                '     This check is to make sure that the comma found isn't part of a function in the equation
                If (Left(ParamEq, 1) = "=") Or ((InStr(ParamEq, ",")) And ((InStr(ParamEq, "(") = 0) Or (InStr(ParamEq, "(") > InStr(ParamEq, ",")))) Then
                    ' Create regular scalar using address of provided parameter
                    Set temphandle = MeasurementGroup.CreateItem(VISION_DATAITEM_SCALAR, ParamName)
                    
                    If (Left(ParamEq, 1) = "=") Then
                        ' Equal sign in equation. Direct copy of address and formula
                        ReplacementParameter = Right(ParamEq, Len(ParamEq) - 1)
                    
                        Set temphandle2 = MeasurementGroup.FindDataItem(ReplacementParameter)
                        If Not (temphandle2 Is Nothing) Then
                            If (temphandle2.Type = VISION_DATAITEM_SCALAR) Then
                                temphandle.BaseAddress = temphandle2.BaseAddress
                                temphandle.FormulaType = temphandle2.FormulaType
                                temphandle.FormulaAlgebraicEquation = temphandle2.FormulaAlgebraicEquation
                                temphandle.ByteOrder = temphandle2.ByteOrder
                                temphandle.DataSize = temphandle2.DataSize
                                temphandle.DataType = temphandle2.DataType
                                temphandle.MemoryType = temphandle2.MemoryType
                            End If
                            Set temphandle2 = Nothing
                        End If
                    ElseIf InStr(ParamEq, ",") Then
                        ' BB - comma in equation
                            
                        ' BB - left of comma is replacement parameter
                        ReplacementParameter = Left(ParamEq, InStr(ParamEq, ",") - 1)
                        ' BB - right of comma is formula
                        Formula = Right(ParamEq, Len(ParamEq) - InStr(ParamEq, ","))
                            
                        Set temphandle2 = MeasurementGroup.FindDataItem(ReplacementParameter)
                        If Not (temphandle2 Is Nothing) Then
                            temphandle.BaseAddress = temphandle2.BaseAddress
                            temphandle.ByteOrder = temphandle2.ByteOrder
                            temphandle.DataSize = temphandle2.DataSize
                            temphandle.DataType = temphandle2.DataType
                            temphandle.MemoryType = temphandle2.MemoryType
                            Set temphandle2 = Nothing
                        End If
                        ' RegExp to determine slope offset formula
                        ' Formula should always begin with x
                        ' followed slope and/or offset
                        ' Formula : x*(A/B) + C
                        Dim fStr As String
                        Dim RegEx As RegExp
                        Set RegEx = New RegExp
                        With RegEx
                            .Pattern = "x(\*([\d\/\.]+))?([\+\-]([\d\/\.]+))?"
                            .Global = True
                        End With
                        fStr = Formula
                        Set matches = RegEx.Execute(fStr)
                        If Not IsEmpty(matches) And (matches.Count > 0) Then
                            ' Get A/B
                            If Not IsEmpty(matches.Item(0).SubMatches(1)) Then
                                If InStr(matches.Item(0).SubMatches(1), "/") Then
                                    ' if slope is of the form A/B
                                    splitedata = Split(matches.Item(0).SubMatches(1), "/")
                                    DataA = splitedata(0)
                                    DataB = splitedata(1)
                                Else
                                    DataA = matches.Item(0).SubMatches(1)
                                    DataB = 1
                                End If
                            Else
                                DataA = 1
                                DataB = 1
                            End If
                            ' Get C
                            If Not IsEmpty(matches.Item(0).SubMatches(3)) Then
                                DataC = matches.Item(0).SubMatches(2)
                            Else
                                DataC = 0
                            End If
                            ' Update formula
                            temphandle.FormulaType = 5   ' VISION_FORMULA_ALGEBRAIC
                            temphandle.SetFormulaSlopeOffset DataA, DataB, DataC
                        Else
                            ' Assign algebraic formula if slope offset not detected
                            temphandle.FormulaType = 1      ' VISION_FORMULA_ALGEBRAIC
                            temphandle.FormulaAlgebraicEquation = Formula
                        End If  ' Not IsEmpty(matches)...Check Regexp match
                    End If 'Left(ParamEq, 1) = "="
                Else
                    ' Virtual scalar
                    Set temphandle = MeasurementGroup.CreateItem(VISION_DATAITEM_VIRTUALSCALAR, ParamName)
                    temphandle.FormulaEquation = ParamEq
                End If 'ParamEq check
                
                ' Set Comment
                If (Comment = "") Then
                    temphandle.Comments = "Dummy parameter added by VST Tool to avoid missing channel errors"
                Else
                    temphandle.Comments = Comment
                End If
            End If ' temphandle is nothing
        End If ' add vs. delete
ContinueDo:
        Set temphandle = Nothing
        RowNum = RowNum + 1
    Loop
    
    ' Default decimals
    WorkSheetName = "Other Settings"
    Set DefaultDecimalCell = Worksheets(WorkSheetName).Cells(1, 2)
    Set DefaultMapYDecimalCell = Worksheets(WorkSheetName).Cells(4, 2)
    Set DefaultMapXDecimalCell = Worksheets(WorkSheetName).Cells(7, 2)
    Set DefaultCurveXDecimalCell = Worksheets(WorkSheetName).Cells(10, 2)
    
    If ((DefaultDecimalCell.Value <> "") And IsNumeric(DefaultDecimalCell.Value)) Then
        ChangeDecimals Strategy.GroupDataItem, DefaultDecimalCell.Value
    End If
    If ((DefaultMapYDecimalCell.Value <> "") And IsNumeric(DefaultMapYDecimalCell.Value)) Then
        ChangeYDecimals Strategy.GroupDataItem, DefaultMapYDecimalCell.Value
    End If
    If ((DefaultMapXDecimalCell.Value <> "") And IsNumeric(DefaultMapXDecimalCell.Value)) Then
        ChangeMapXDecimals Strategy.GroupDataItem, DefaultMapXDecimalCell.Value
    End If
    If ((DefaultCurveXDecimalCell.Value <> "") And IsNumeric(DefaultCurveXDecimalCell.Value)) Then
        ChangeCurveXDecimals Strategy.GroupDataItem, DefaultCurveXDecimalCell.Value
    End If
        
    ' Generic defaults
    Set Rng1 = Worksheets("Parameters").Cells.Columns(1).Find(What:="_ALL_", After:=Worksheets("Parameters").Cells(4, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True)
    If Not (Rng1 Is Nothing) Then
        ChangeAllDefaults Strategy.GroupDataItem, Rng1.Row
    End If
    
    ' Data bins for missing parameters
    Dim misscount As Long
    Dim missingparams() As String
    misscount = 0
        
    ' Loop through parameters
    Progress.Label1.caption = "Modifying Parameter Settings"
    WorkSheetName = "Parameters"
    Dim DataItem As Object
    'Dim ScalarHandle As Object
    Dim CurrentParameterName As String
    RowNum = 4
    Do While (Worksheets(WorkSheetName).Cells(RowNum, 1).Value <> "")
        CurrentParameterName = Trim(Worksheets(WorkSheetName).Cells(RowNum, 1).Value)
        If (CurrentParameterName <> "_ALL_") Then
                
            Progress.Label2.caption = CurrentParameterName
            Progress.Repaint
            
            'Find the Data Item in the strategy, and set to DataItem
            Set DataItem = Strategy.FindOrSearchDataItem(CurrentParameterName)
            
            If (Not (DataItem Is Nothing)) Then
                Select Case DataItem.Type
                    Case VISION_DATAITEM_SCALAR
                        For ColNum = 2 To 20
                            UpdatePropertyPerColumn DataItem, RowNum, ColNum
                        Next
                        UpdatePropertyPerColumn DataItem, RowNum, 77
                    Case VISION_DATAITEM_STATEVARIABLE
                        For ColNum = 2 To 2
                            UpdatePropertyPerColumn DataItem, RowNum, ColNum
                        Next
                    Case VISION_DATAITEM_ARRAY1D
                        UpdatePropertyPerColumn DataItem, RowNum, 2
                        For ColNum = 39 To 56
                            UpdatePropertyPerColumn DataItem, RowNum, ColNum
                        Next
                        UpdatePropertyPerColumn DataItem, RowNum, 79
                    Case VISION_DATAITEM_ARRAY2D
                        UpdatePropertyPerColumn DataItem, RowNum, 2
                        For ColNum = 57 To 74
                            UpdatePropertyPerColumn DataItem, RowNum, ColNum
                        Next
                        UpdatePropertyPerColumn DataItem, RowNum, 80
                    Case VISION_DATAITEM_TABLE2D
                        UpdatePropertyPerColumn DataItem, RowNum, 2
                        For ColNum = 21 To 56
                            UpdatePropertyPerColumn DataItem, RowNum, ColNum
                        Next
                        UpdatePropertyPerColumn DataItem, RowNum, 75
                        UpdatePropertyPerColumn DataItem, RowNum, 78
                        UpdatePropertyPerColumn DataItem, RowNum, 79
                    Case VISION_DATAITEM_TABLE3D
                        UpdatePropertyPerColumn DataItem, RowNum, 2
                        For ColNum = 21 To 74
                            UpdatePropertyPerColumn DataItem, RowNum, ColNum
                        Next
                        UpdatePropertyPerColumn DataItem, RowNum, 75
                        UpdatePropertyPerColumn DataItem, RowNum, 76
                        UpdatePropertyPerColumn DataItem, RowNum, 78
                        UpdatePropertyPerColumn DataItem, RowNum, 79
                        UpdatePropertyPerColumn DataItem, RowNum, 80
                    Case Else
                        ActivateExcel
                        msgLogDisp "The parameter " & CurrentParameterName & " is not a scalar, state variable, table, function or array."
                End Select
            Else
                ReDim Preserve missingparams(misscount)
                missingparams(misscount) = CurrentParameterName
                misscount = misscount + 1
            End If
        
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
    
    ' Process APRs
    If (Worksheets("Other Settings").Shapes("AprCheckBox").ControlFormat.Value = 1) Then
        Progress.Label1.caption = "Setting APR Range Limits"
        Progress.Label2.caption = ""
        Progress.Repaint
    
        ' Do it
        If (AtiFilePath = "") Then
            Set AtiFileCell = Worksheets("File Paths").Range("B2")
            AtiFilePath = GetFilePath(AtiFileCell, "Strategy Description File (*.ati;*.a2l),*.ati;*.a2l", "")
        End If
        DoAprs Strategy, AtiFilePath
    End If
    
    ' Loop through device settings
    Progress.Label1.caption = "Modifying Device Settings"
    Progress.Label2.caption = ""
    Progress.Repaint
    RowNum = 2
    WorkSheetName = "Device Settings"
    Dim SettingName As String
    Dim SettingValue As String
    Do While (Worksheets(WorkSheetName).Cells(RowNum, 1).Value <> "")
        SettingName = Trim(Worksheets(WorkSheetName).Cells(RowNum, 1).Value)
        SettingValue = Trim(Worksheets(WorkSheetName).Cells(RowNum, 2).Value)

        Progress.Label2.caption = SettingName
        Progress.Repaint
        
        If (SettingValue = "DELETE") Then
            Ret = Strategy.RemoveStrategySetting(SettingName)
        Else
            Ret = Strategy.SetStrategySetting(SettingName, SettingValue)
        End If
        
        If (Ret <> VISION__OK) Then
            ActivateExcel
            msgLogDisp "Error setting " & SettingName
        End If
        
        RowNum = RowNum + 1
    Loop
   
    Progress.Label2.caption = ""
    Progress.Repaint
   
    ' Make any cal changes

    If (GetNumericVisionApiVersion >= 3008002) Then
        UpdateCalibrationValues Strategy, AtiFilePath
    End If
    
    ' Create KAM memory region
    ActivateExcel
    WorkSheetName = "Other Settings"
    Set MapFileCell = Worksheets("File Paths").Range("B5")
    MAPFilePath = "False"
    If (Worksheets("Other Settings").Shapes("KamRegionCheckBox").ControlFormat.Value = 1) Then
        Progress.Label1.caption = "Adding KAM Region"
        Progress.Label2.caption = ""
        If MultipleFilesFlag = True Then
            If Not (FileNames(1) = "") Then
                MAPFilePath = FileNames(1)
            Else
                MAPFilePath = "False"
            End If
        Else
            MAPFilePath = GetFilePath(MapFileCell, "Memory Map File (*.map),*.map", CurDir & "\")
        End If
        
        If Not (MAPFilePath = "False") Then
            Processor = DetermineProcessor(MAPFilePath)
            Select Case Processor
                Case 0 ' Not Determined
                    msgLogDisp "Can't determine type of processor.  Cannot create KAM region.", vbExclamation
                Case 1 ' Bosch TriCore
                    NvramStart = GetAddressFromMap(MAPFilePath, "__ENVRAM_FORD_START", 1)
                    NvramEnd = GetAddressFromMap(MAPFilePath, "__ENVRAM_END", 1)
                Case 2 ' PPC
                    msgLogDisp "Cannot create KAM region for PPC processors.", vbExclamation
                Case 3, 4 ' Conti TriCore
                    NvramStart = GetAddressFromMap(MAPFilePath, "KARAM_Start", 3)
                    NvramEnd = GetAddressFromMap(MAPFilePath, "KARAM_End", 3)
                Case 5
                    NvramStart = GetAddressFromMap(MAPFilePath, "__ENVRAM_START", 5)
                    NvramEnd = GetAddressFromMap(MAPFilePath, "__ENVRAM_END", 5)
                Case 6
                    NvramStart = GetAddressFromMap(MAPFilePath, "__ENVRAM_START", 6)
                    NvramEnd = GetAddressFromMap(MAPFilePath, "__ENVRAM_END", 6)
                Case Else
                    msgLogDisp "Processor Type Not Supported.  Cannot create KAM region.", vbExclamation
            End Select
            If ((NvramStart <> "0") And (NvramEnd <> "0") And (NvramStart <> NvramEnd)) Then
                NvramLength = HexAdd(HexSubtract(NvramStart, NvramEnd), "0x1")
                DoMemoryRegionWithWorkArounds Strategy, "KAM", NvramStart, NvramLength, "RK", 0, ShowMisMatchError, BatchMode
            End If
        Else
            msgLogDisp "No map file specified. KAM region not added."
        End If
    End If
    
    ' Enable No Hooks
    If (Worksheets("Other Settings").Shapes("NoHooksCheckBox").ControlFormat.Value = 1) Then
        Progress.Label1.caption = "Enabling No Hooks"
        Progress.Label2.caption = ""
        
        If MultipleFilesFlag = True Then
            If Not (FileNames(1) = "") Then
                MAPFilePath = FileNames(1)
            Else
                MAPFilePath = "False"
            End If
        Else
            MAPFilePath = GetFilePath(MapFileCell, "Memory Map File (*.map),*.map", CurDir & "\")
        End If
        
        If Not (MAPFilePath = "False") Then
            FindMemoryForNoHooks MAPFilePath, Strategy
        Else
            msgLogDisp "No map file specified. Unable to setup NoHooks."
        End If ' Found MAP file
        
    End If ' Requested no hooks
    
    ' Return here if called from another sub
    If (Not ExistingHandle Is Nothing) Then
        Exit Sub
    End If
    
    ' Save Strategy
    Progress.Label1.caption = "Saving changes to VST..."
    Progress.Label2.caption = ""
    Progress.Repaint
    Strategy.Save
    VSTFilePath = Strategy.Filename
    Set Strategy = Nothing
    
    ' Make calibration changes after saving
    If (GetNumericVisionApiVersion < 3008002) Then
        DoCalChanges VSTFilePath
    End If
    
    ' Clean up
CleanUp:
    Progress.Label1.caption = "Cleaning up..."
    Progress.Label2.caption = ""
    Progress.Repaint
    
    Set Strategy = Nothing
    
    ' Close
    Progress.Label1.caption = "Done!"
    Progress.Repaint
    Application.Wait (Now + TimeValue("0:00:02"))
    Progress.Hide
End Sub

Sub UpdatePropertyPerColumn(DataItem As Object, RowNum As Long, ColNum As Long)
    Dim WorkSheetName As String, CurrentCellValue As String
    
    WorkSheetName = "Parameters"
    CurrentCellValue = Worksheets(WorkSheetName).Cells(RowNum, ColNum).Value
    If (CurrentCellValue <> "") Then
        Select Case ColNum
            Case 2
                DataItem.Comments = CurrentCellValue
            Case 3
                DataItem.SmallStep = CurrentCellValue
            Case 4
                DataItem.LargeStep = CurrentCellValue
            Case 5
                DataItem.EnableRangeLimit = CurrentCellValue
            Case 6
                DataItem.MinimumLimit = CurrentCellValue
            Case 7
                DataItem.MaximumLimit = CurrentCellValue
            Case 8
                DataItem.DecimalPlaces = CurrentCellValue
            Case 9
                DataItem.DisplayColor = CurrentCellValue
            Case 10
                DataItem.DisplayBkgrdColor = CurrentCellValue
            Case 11
                DataItem.MinimumThreshold = CurrentCellValue
            Case 12
                DataItem.MaximumThreshold = CurrentCellValue
            Case 13
                DataItem.EnableRangeText = CurrentCellValue
            Case 14
                DataItem.MinimumText = CurrentCellValue
            Case 15
                DataItem.MaximumText = CurrentCellValue
            Case 16
                DataItem.EnableRangeColors = CurrentCellValue
            Case 17
                DataItem.MinimumColor = CurrentCellValue
            Case 18
                DataItem.MinimumBkgrdColor = CurrentCellValue
            Case 19
                DataItem.MaximumColor = CurrentCellValue
            Case 20
                DataItem.MaximumBkgrdColor = CurrentCellValue
            Case 21
                DataItem.XAxisSmallStep = CurrentCellValue
            Case 22
                DataItem.XAxisLargeStep = CurrentCellValue
            Case 23
                DataItem.XAxisEnableRangeLimit = CurrentCellValue
            Case 24
                DataItem.XAxisMinimumLimit = CurrentCellValue
            Case 25
                DataItem.XAxisMaximumLimit = CurrentCellValue
            Case 26
                DataItem.XAxisDecimalPlaces = CurrentCellValue
            Case 27
                DataItem.XAxisDisplayColor = CurrentCellValue
            Case 28
                DataItem.XAxisDisplayBkgrdColor = CurrentCellValue
            Case 29
                DataItem.XAxisMinimumThreshold = CurrentCellValue
            Case 30
                DataItem.XAxisMaximumThreshold = CurrentCellValue
            Case 31
                DataItem.XAxisEnableRangeText = CurrentCellValue
            Case 32
                DataItem.XAxisMinimumText = CurrentCellValue
            Case 33
                DataItem.XAxisMaximumText = CurrentCellValue
            Case 34
                DataItem.XAxisEnableRangeColors = CurrentCellValue
            Case 35
                DataItem.XAxisMinimumColor = CurrentCellValue
            Case 36
                DataItem.XAxisMinimumBkgrdColor = CurrentCellValue
            Case 37
                DataItem.XAxisMaximumColor = CurrentCellValue
            Case 38
                DataItem.XAxisMaximumBkgrdColor = CurrentCellValue
            Case 39
                DataItem.YAxisSmallStep = CurrentCellValue
            Case 40
                DataItem.YAxisLargeStep = CurrentCellValue
            Case 41
                DataItem.YAxisEnableRangeLimit = CurrentCellValue
            Case 42
                DataItem.YAxisMinimumLimit = CurrentCellValue
            Case 43
                DataItem.YAxisMaximumLimit = CurrentCellValue
            Case 44
                DataItem.YAxisDecimalPlaces = CurrentCellValue
            Case 45
                DataItem.YAxisDisplayColor = CurrentCellValue
            Case 46
                DataItem.YAxisDisplayBkgrdColor = CurrentCellValue
            Case 47
                DataItem.YAxisMinimumThreshold = CurrentCellValue
            Case 48
                DataItem.YAxisMaximumThreshold = CurrentCellValue
            Case 49
                DataItem.YAxisEnableRangeText = CurrentCellValue
            Case 50
                DataItem.YAxisMinimumText = CurrentCellValue
            Case 51
                DataItem.YAxisMaximumText = CurrentCellValue
            Case 52
                DataItem.YAxisEnableRangeColors = CurrentCellValue
            Case 53
                DataItem.YAxisMinimumColor = CurrentCellValue
            Case 54
                DataItem.YAxisMinimumBkgrdColor = CurrentCellValue
            Case 55
                DataItem.YAxisMaximumColor = CurrentCellValue
            Case 56
                DataItem.YAxisMaximumBkgrdColor = CurrentCellValue
            Case 57
                DataItem.ZAxisSmallStep = CurrentCellValue
            Case 58
                DataItem.ZAxisLargeStep = CurrentCellValue
            Case 59
                DataItem.ZAxisEnableRangeLimit = CurrentCellValue
            Case 60
                DataItem.ZAxisMinimumLimit = CurrentCellValue
            Case 61
                DataItem.ZAxisMaximumLimit = CurrentCellValue
            Case 62
                DataItem.ZAxisDecimalPlaces = CurrentCellValue
            Case 63
                DataItem.ZAxisDisplayColor = CurrentCellValue
            Case 64
                DataItem.ZAxisDisplayBkgrdColor = CurrentCellValue
            Case 65
                DataItem.ZAxisMinimumThreshold = CurrentCellValue
            Case 66
                DataItem.ZAxisMaximumThreshold = CurrentCellValue
            Case 67
                DataItem.ZAxisEnableRangeText = CurrentCellValue
            Case 68
                DataItem.ZAxisMinimumText = CurrentCellValue
            Case 69
                DataItem.ZAxisMaximumText = CurrentCellValue
            Case 70
                DataItem.ZAxisEnableRangeColors = CurrentCellValue
            Case 71
                DataItem.ZAxisMinimumColor = CurrentCellValue
            Case 72
                DataItem.ZAxisMinimumBkgrdColor = CurrentCellValue
            Case 73
                DataItem.ZAxisMaximumColor = CurrentCellValue
            Case 74
                DataItem.ZAxisMaximumBkgrdColor = CurrentCellValue
            Case 75
                On Error Resume Next
                DataItem.XAxisScalarDataItemForRunPt = CurrentCellValue
                If (DataItem.XAxisScalarDataItemForRunPt <> CurrentCellValue) Then
                    msgLogDisp "Error setting X axis running point for " & DataItem.DataItemName & vbCrLf & _
                        "Make sure running point name is valid and is of the form ""this->Measurements.PARAM_NAME"""
                End If
                On Error GoTo 0
            Case 76
                On Error Resume Next
                DataItem.YAxisScalarDataItemForRunPt = CurrentCellValue
                If (DataItem.YAxisScalarDataItemForRunPt <> CurrentCellValue) Then
                    msgLogDisp "Error setting Y axis running point for " & DataItem.DataItemName & vbCrLf & _
                        "Make sure running point name is valid and is of the form ""this->Measurements.PARAM_NAME"""
                End If
                On Error GoTo 0
            Case 77
                DataItem.EngineeringUnits = CurrentCellValue
            Case 78
                DataItem.XAxisEngineeringUnits = CurrentCellValue
            Case 79
                DataItem.YAxisEngineeringUnits = CurrentCellValue
            Case 80
                DataItem.ZAxisEngineeringUnits = CurrentCellValue
        End Select
    End If
End Sub

Sub ChangeDecimals(GroupItemHandle As Object, DecimalValue As Long)
    Dim DataItemHandle As Object
    Progress.Label1.caption = "Setting Default Decimal Places"
    Progress.Label2.caption = GroupItemHandle.FullDataItemName
    Progress.Repaint

    For Each DataItemHandle In GroupItemHandle.Items
        Select Case DataItemHandle.Type
            Case VISION_DATAITEM_GROUP
                If (DataItemHandle.DataItemName <> "Measurements") Then
                    ChangeDecimals DataItemHandle, DecimalValue
                End If
            Case VISION_DATAITEM_SCALAR
                DataItemHandle.DecimalPlaces = NotZero(DecimalValue, DataItemHandle.DecimalPlaces)
            Case VISION_DATAITEM_ARRAY1D
                DataItemHandle.YAxisDecimalPlaces = NotZero(DecimalValue, DataItemHandle.YAxisDecimalPlaces)
            Case VISION_DATAITEM_TABLE2D
                DataItemHandle.XAxisDecimalPlaces = NotZero(DecimalValue, DataItemHandle.XAxisDecimalPlaces)
                DataItemHandle.YAxisDecimalPlaces = NotZero(DecimalValue, DataItemHandle.YAxisDecimalPlaces)
            Case VISION_DATAITEM_ARRAY2D
                DataItemHandle.ZAxisDecimalPlaces = NotZero(DecimalValue, DataItemHandle.ZAxisDecimalPlaces)
            Case VISION_DATAITEM_TABLE3D
                DataItemHandle.XAxisDecimalPlaces = NotZero(DecimalValue, DataItemHandle.XAxisDecimalPlaces)
                DataItemHandle.YAxisDecimalPlaces = NotZero(DecimalValue, DataItemHandle.YAxisDecimalPlaces)
                DataItemHandle.ZAxisDecimalPlaces = NotZero(DecimalValue, DataItemHandle.ZAxisDecimalPlaces)
        End Select
    Next
End Sub

Sub ChangeYDecimals(GroupItemHandle As Object, DecimalValue As Long)
    Dim DataItemHandle As Object
    Progress.Label1.caption = "Setting Y-Axis Decimal Places"
    Progress.Label2.caption = GroupItemHandle.FullDataItemName
    Progress.Repaint

    For Each DataItemHandle In GroupItemHandle.Items
        Select Case DataItemHandle.Type
            Case VISION_DATAITEM_GROUP
                If (DataItemHandle.DataItemName = "Characteristics") Or (DataItemHandle.DataItemName = "Maps") Then
                    ChangeYDecimals DataItemHandle, DecimalValue
                End If
            Case VISION_DATAITEM_TABLE3D
                DataItemHandle.YAxisDecimalPlaces = DecimalValue
        End Select
    Next
End Sub

Sub ChangeMapXDecimals(GroupItemHandle As Object, DecimalValue As Long)
    Dim DataItemHandle As Object
    Progress.Label1.caption = "Setting X-Axis Decimal Places"
    Progress.Label2.caption = GroupItemHandle.FullDataItemName
    Progress.Repaint

    For Each DataItemHandle In GroupItemHandle.Items
        Select Case DataItemHandle.Type
            Case VISION_DATAITEM_GROUP
                If (DataItemHandle.DataItemName = "Characteristics") Or (DataItemHandle.DataItemName = "Maps") Then
                    ChangeMapXDecimals DataItemHandle, DecimalValue
                End If
            Case VISION_DATAITEM_TABLE3D
                DataItemHandle.XAxisDecimalPlaces = DecimalValue
        End Select
    Next
End Sub

Sub ChangeCurveXDecimals(GroupItemHandle As Object, DecimalValue As Long)
    Dim DataItemHandle As Object
    Progress.Label1.caption = "Setting X-Axis Decimal Places"
    Progress.Label2.caption = GroupItemHandle.FullDataItemName
    Progress.Repaint

    For Each DataItemHandle In GroupItemHandle.Items
        Select Case DataItemHandle.Type
            Case VISION_DATAITEM_GROUP
                If (DataItemHandle.DataItemName = "Characteristics") Or (DataItemHandle.DataItemName = "Curves") Then
                    ChangeCurveXDecimals DataItemHandle, DecimalValue
                End If
            Case VISION_DATAITEM_TABLE2D
                DataItemHandle.XAxisDecimalPlaces = DecimalValue
        End Select
    Next
End Sub

Sub ChangeAllDefaults(GroupItemHandle As Object, RowNum As Long)
    Dim DataItemHandle As Object
    Dim ColNum As Long
    
    Progress.Label1.caption = "Setting Default Options for All Parameters"
    Progress.Label2.caption = GroupItemHandle.FullDataItemName
    Progress.Repaint
    
    For Each DataItemHandle In GroupItemHandle.Items
        Select Case DataItemHandle.Type
            Case VISION_DATAITEM_GROUP
                ChangeAllDefaults DataItemHandle, RowNum
            Case Else
                For ColNum = 2 To 74
                    On Error Resume Next
                    UpdatePropertyPerColumn DataItemHandle, RowNum, ColNum
                    On Error GoTo 0
                Next
        End Select
    Next
End Sub

Function NotZero(In1 As Long, In2 As Long) As Long
    If In2 = 0 Then
        NotZero = 0
    Else
        NotZero = In1
    End If
End Function

' Based on FileToSHA1Hex() and GetFileBytes() from http://stackoverflow.com/questions/2826302/how-to-get-the-md5-hex-hash-for-a-file-using-vba

Private Function GetFileBytes(ByVal Path As String) As Byte()
    Dim lngFileNum As Long
    Dim bytRtnVal() As Byte
    lngFileNum = FreeFile
    If LenB(Dir(Path)) Then ''// Does file exist?
        Open Path For Binary Access Read As lngFileNum
        ReDim bytRtnVal(LOF(lngFileNum) - 1&) As Byte
        Get lngFileNum, , bytRtnVal
        Close lngFileNum
    Else
        Err.Raise 53
    End If
    GetFileBytes = bytRtnVal
    Erase bytRtnVal
End Function

Public Function FileToSHA1Integer(sFileName As String) As Long
    Dim enc As Object
    Dim bytes As Variant
    Dim pos As Long
    Dim checksum As Long
    
    Set enc = CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider")
    
    bytes = GetFileBytes(sFileName)
    bytes = enc.ComputeHash_2((bytes))
    
    checksum = 0
    For pos = 1 To LenB(bytes)
        checksum = checksum + AscB(MidB(bytes, pos, 1)) * pos
    Next
    
    ' VBA is rather lame and doesn't allow "overflow"
    
    Do While (checksum > 65535)
        checksum = checksum - 65535
    Loop
        
    FileToSHA1Integer = checksum
    
    Set enc = Nothing
End Function

Sub UpdateCalibrationValues(Strategy As Object, ByRef AtiFilePath As String)
    Dim misscount As Long, RowNum As Long, CSum As Long, i As Long
    Dim WorkSheetName As String, SettingName As String, SettingValue As String, missing_string As String
    Dim DataItem As Object
    Dim missingparams() As String
    
    ' Data bins for missing parameters
    misscount = 0
    ReDim missingparams(misscount)
    
    Progress.Label1.caption = "Modifying Calibration Values"
    Progress.Label2.caption = ""
    Progress.Repaint
    RowNum = 2
    WorkSheetName = "Cal Changes"
    Do While (Worksheets(WorkSheetName).Cells(RowNum, 1).Value <> "")
        SettingName = Trim(Worksheets(WorkSheetName).Cells(RowNum, 1).Value)
        SettingValue = Trim(Worksheets(WorkSheetName).Cells(RowNum, 2).Value)

        Progress.Label2.caption = SettingName
        Progress.Repaint

        Set DataItem = Strategy.FindOrSearchDataItem(SettingName)
        If (Not (DataItem Is Nothing)) Then
            If (SettingValue = "CHECKSUM") Then
                CSum = FileToSHA1Integer(AtiFilePath)
                
                DataItem.ActualValue = CSum
                DataItem.BaseValue = CSum
                DataItem.TargetValue = CSum
            Else
                DataItem.ActualValue = SettingValue
                DataItem.BaseValue = SettingValue
                DataItem.TargetValue = SettingValue
            End If
        Else
            ReDim Preserve missingparams(misscount)
            missingparams(misscount) = SettingName
            misscount = misscount + 1
        End If
        
        RowNum = RowNum + 1
    Loop
    
    Progress.Label2.caption = ""
    Progress.Repaint
    
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
End Sub

Sub AddMemoryRegions(Strategy As Object, ByRef ShowMisMatchError As Boolean, Optional BatchMode As Boolean = False)
    Dim RowNum As Long, x As Long, RegionBitMap As Long
    Dim Ret As Boolean, RegionExists As Boolean
    Dim WorkSheetName As String
    Dim MemoryMap As Object
    
    Progress.Label1.caption = "Adding Memory Regions"
    Progress.Repaint
    RowNum = 2
    WorkSheetName = "Memory Regions"
    Dim RegionName As String
    Dim StartAddress As String
    Dim RegionSize As String
    Dim RegionFlags As String
    Dim ChecksumType As Long
    
    Do While (Worksheets(WorkSheetName).Cells(RowNum, 1).Value <> "")
        RegionName = Worksheets(WorkSheetName).Cells(RowNum, 1).Value
        StartAddress = Worksheets(WorkSheetName).Cells(RowNum, 2).Value
        RegionSize = Worksheets(WorkSheetName).Cells(RowNum, 3).Value
        RegionFlags = Worksheets(WorkSheetName).Cells(RowNum, 4).Value
        ChecksumType = Worksheets(WorkSheetName).Cells(RowNum, 5).Value
        
        Progress.Label2.caption = RegionName
        Progress.Repaint

        If (StartAddress = "DELETE") Then
            Ret = Strategy.RemoveMemoryRegion(RegionName)
            If (Ret = False) Then
                ActivateExcel
                msgLogDisp "Error deleting memory region " & RegionName
            End If
        Else
            ' See if region already exists
            RegionExists = False
            For x = 1 To Strategy.StrategyMemoryMap.Count
                Set MemoryMap = Strategy.StrategyMemoryMap.Item(x)
                If MemoryMap.RegionName = RegionName Then
                    RegionExists = True
                    Exit For
                End If
            Next x
            
            If Not RegionExists Then
                ' Add Memory Region
                DoMemoryRegionWithWorkArounds Strategy, RegionName, StartAddress, RegionSize, RegionFlags, ChecksumType, ShowMisMatchError, BatchMode
            Else
                ' Update Existing Region
                If (StartAddress <> "") Then
                    MemoryMap.Address = Hex2Lng(StartAddress)
                End If
                If (RegionSize <> "") Then
                    MemoryMap.Size = Hex2Lng(RegionSize)
                End If
                If (ChecksumType > 0) Then
                    MemoryMap.SupportsChecksum = True
                    MemoryMap.ChecksumType = ChecksumType
                End If
                If (RegionFlags <> "") Then
                    ' Translate alpha flags to bitmap
                    RegionBitMap = 0
                    If (InStr(RegionFlags, "F") > 0) Then RegionBitMap = RegionBitMap + 2 ^ 0
                    If (InStr(RegionFlags, "D") > 0) Then RegionBitMap = RegionBitMap + 2 ^ 1
                    If (InStr(RegionFlags, "U") > 0) Then RegionBitMap = RegionBitMap + 2 ^ 2
                    If (InStr(RegionFlags, "T") > 0) Then RegionBitMap = RegionBitMap + 2 ^ 3
                    If (InStr(RegionFlags, "P") > 0) Then RegionBitMap = RegionBitMap + 2 ^ 4
                    If (InStr(RegionFlags, "R") > 0) Then RegionBitMap = RegionBitMap + 2 ^ 5
                    If (InStr(RegionFlags, "C") > 0) Then RegionBitMap = RegionBitMap + 2 ^ 6
                    If (InStr(RegionFlags, "K") > 0) Then RegionBitMap = RegionBitMap + 2 ^ 7
                    If (InStr(RegionFlags, "H") > 0) Then RegionBitMap = RegionBitMap + 2 ^ 8
                    MemoryMap.FlagsValue = RegionBitMap
                End If
            End If ' region exists
        End If ' delete or add/update
        
        RowNum = RowNum + 1
    Loop
End Sub

Sub DoMemoryRegionWithWorkArounds(Strategy As Object, ByVal RegionName As String, ByVal StartAddress As String, ByVal RegionSize As String, ByVal RegionFlags As String, ByVal ChecksumType As Long, ByRef ShowMisMatchError As Boolean, BatchMode As Boolean)
    Dim x As Long
    Dim MemoryMap As Object
    
    'Add the memory region with a start address of 0x1
    Strategy.AddMemoryRegion RegionName, "0x1", RegionSize, RegionFlags
    
    'Go back and verify that the region was added
    Dim RegionAdded As Boolean
    RegionAdded = False
    For x = 1 To Strategy.StrategyMemoryMap.Count
        Set MemoryMap = Strategy.StrategyMemoryMap.Item(x)
        If MemoryMap.RegionName = RegionName Then
            'Set the address - call takes a Long, so must convert to VB Hex notation.
            MemoryMap.Address = Hex2Lng(StartAddress)
            If ChecksumType > 0 Then
                MemoryMap.SupportsChecksum = True
                MemoryMap.ChecksumType = ChecksumType
            End If
            RegionAdded = True
        Exit For
        End If
    Next x
        
    If Not RegionAdded Then
        If BatchMode Then
            If ShowMisMatchError Then
                msgLogDisp "Unable to add the memory region " & RegionName, vbCritical, "AddMemoryRegions"
                Dim resp As VbMsgBoxResult
                resp = msgLogDisp("Do you want to ignore all batch build files error for Memory Region and State Var Colors", vbYesNo, "Skip Errors and Continue", vbNo)
                If resp = vbYes Then
                    ShowMisMatchError = False
                End If
            Else
                If AutomatedMode Then
                    msgLogDisp "Unable to add the memory region " & RegionName, vbCritical, "AddMemoryRegions"
                End If
            End If
        Else
            msgLogDisp "Unable to add the memory region " & RegionName, vbCritical, "AddMemoryRegions"
        End If
    End If
End Sub

Function FindItemsGroup(GDII As Object, ByVal ParamName As String) As Object
    Dim tmpGroup As Object
    Dim GroupNum As Long
    
    For GroupNum = 1 To GDII.Groups.Count
        If (GDII.Groups(GroupNum).DataItemName <> "Functions") Then
            If Not (GDII.Groups(GroupNum).FindDataItem(ParamName) Is Nothing) Then
                Set FindItemsGroup = GDII.Items(GroupNum)
                Exit For
            ElseIf (GDII.Groups(GroupNum).Groups.Count > 0) Then
                Set tmpGroup = FindItemsGroup(GDII.Groups(GroupNum), ParamName)
                If Not (tmpGroup Is Nothing) Then
                    Set FindItemsGroup = tmpGroup
                    Exit For
                End If
            End If
        End If
    Next
End Function

Sub BatchVstUpdate()
    Dim Strategy As Object
    Dim Ret As VISION_STATUS_CODES
    Dim sFilePaths As Variant, sFilePath As Variant
    Dim vstName As String
    Dim BatchMode As Boolean
    Dim ShowMisMatchError As Boolean
    ShowMisMatchError = True
    
    ' Choose which VST
    sFilePaths = Application.GetOpenFilename(FileFilter:="VISION Strategy files (*.vst),*.vst", MultiSelect:=True)
    If Not (IsArray(sFilePaths)) Then
        If (sFilePaths = "False") Then
            Exit Sub
        End If
    End If
    
    ' Progress Box
    Progress.Label1.caption = "Initializing..."
    Progress.Label2.caption = ""
    Progress.Show
    Application.Wait (Now + TimeValue("0:00:02"))

    If UBound(sFilePaths) > 1 Then
        BatchMode = True
    Else
        BatchMode = False
    End If

    For Each sFilePath In sFilePaths
        vstName = FileParts(sFilePath, "nameonly")

        'Create the ActiveX object
        Progress.Label1.caption = "Opening VST..."
        Progress.Label2.caption = vstName
        Set Strategy = CreateObject("Vision.StrategyFileInterface")
        Ret = Strategy.Open(sFilePath)
        If (Ret <> VISION__OK) Then
            ActivateExcel
            msgLogDisp "Error opening strategy file."
            Exit Sub
        End If
        
        ' Add memory regions if desired.  On new VSTs, this will get called within BuildVSTFile
        AddMemoryRegions Strategy, ShowMisMatchError, BatchMode
        
        ' Do updates
        VSTParameterTemplate ShowMisMatchError, Strategy, "", BatchMode
        
        'Update the state variable colour
        changecolour Strategy, BatchMode, ShowMisMatchError
        
        ' Save Strategy
        Progress.Label1.caption = "Saving changes to VST..."
        Progress.Label2.caption = vstName
        Progress.Repaint
        Strategy.Save
        Set Strategy = Nothing
    Next
        
CleanUp:
    Progress.Label1.caption = "Cleaning up..."
    Progress.Label2.caption = ""
    Progress.Repaint
    
    Set Strategy = Nothing
    
    ' Close
    Progress.Label1.caption = "Done!"
    Progress.Repaint
    Application.Wait (Now + TimeValue("0:00:02"))
    Progress.Hide
End Sub
