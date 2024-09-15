Attribute VB_Name = "MapMath"
Option Explicit

Function GetAddressFromMap(ByVal MAPFilePath As Variant, ByVal Region As String, Optional ByVal Processor As Long = 1, Optional ByVal DispError As Boolean = True, Optional ByVal NextAddress As Boolean = False) As String
    Dim RegExpResult() As String
    Dim MyLine As String, FirstSpace As String
    
    ' Bosch TriCore Example:
    ' 0x801ad668   0x801ad668     0 g __ASW0_FREE_START                                    asw0_flash       .asw0_free                 .asw0_free                 .\delivery\D1701_020_F1BQ.elf

    ' Conti TriCore Example:
    ' | fmc_code_csum_end                                   | 0x80181034 |              |
    ' | 0x80181034 | fmc_code_csum_end                                   |              |
    
    ' PPC Example:
    '   ROM_3_End            00178e90  00000000            0   0153db6
        
    ' Conti UTCU3 Example:
    ' __PPC_ROM_END2                                                    g 0x80105beb 0x80105bee    4 CODE_FLASH        .protram_copy_sec  .protram_copy_sec                                 tpca3_x07.elf
    ' 0x80105beb 0x80105bee    4 g __PPC_ROM_END2                                                    CODE_FLASH        .protram_copy_sec  .protram_copy_sec                                 tpca3_x07.elf

    Progress.Label2.caption = Region
    Progress.Repaint
    
    'Create object to read the file
    Dim objFSO As FileSystemObject
    Set objFSO = New FileSystemObject
    'Open the file using the object created
    Dim objFile As TextStream
    Set objFile = objFSO.OpenTextFile(MAPFilePath, 1)
    
    'Open MAPFilePath For Input As #1
    'Do While Not EOF(1)
    Do Until objFile.AtEndOfStream
        'Read the line
        MyLine = objFile.ReadLine
        'Line Input #1, MyLine
        If InStr(MyLine, Region) Then
                
            If (NextAddress) Then
                ' Get the next line instead
                MyLine = objFile.ReadLine
                'Line Input #1, MyLine
                If (MyLine = "") Then
                    ' Blank line, read one more
                    MyLine = objFile.ReadLine
                    'Line Input #1, MyLine
                End If
            End If
                
            Select Case Processor
                Case 1, 5, 6 ' Bosch TriCore
                    FirstSpace = InStr(MyLine, " ")
                    GetAddressFromMap = Left(MyLine, FirstSpace - 1)
                
                Case 2 ' PPC
                    RegExpResult = RegExpFind(MyLine, "\b[\da-f]+\b")
                    GetAddressFromMap = RegExpResult(0)
                
                Case 3, 4 ' Conti TriCore
                    RegExpResult = RegExpFind(MyLine, "\b0x[\da-f]+\b")
                    GetAddressFromMap = RegExpResult(0)
                               
            End Select
            objFile.Close
            'Close #1
            Exit Function
        
        End If
    Loop
    
    ' If we got here we didn't find it
    objFile.Close
    'Close #1
    GetAddressFromMap = "0"
    If (DispError) Then
        msgLogDisp "Could not find the region """ & Region & """", vbExclamation
    End If
End Function

Sub TestGetAddress()
    Dim mapfile As Variant
    Dim Region As String, Address As String
    mapfile = Application.GetOpenFilename("Memory Map File (*.map),*.map")
    Region = "__PTA_DATA_END"
    
    Address = GetAddressFromMap(mapfile, Region, 1, True, True)
    
    msgLogDisp Address
End Sub

Function HexSubtract(ByVal H1 As String, ByVal H2 As String) As String
    Dim DecResult As Double
    Dim HexResult As String
    
    H1 = "&H" & Right(H1, Len(H1) - 2)
    H2 = "&H" & Right(H2, Len(H2) - 2)
    
    DecResult = Val(H2) - Val(H1)
    
    HexResult = Hex(DecResult)
    
    HexSubtract = "0x" & HexResult
End Function

Function HexAdd(ByVal H1 As String, ByVal H2 As String) As String
    Dim DecResult As Double
    Dim HexResult As String
    
    H1 = "&H" & Right(H1, Len(H1) - 2)
    H2 = "&H" & Right(H2, Len(H2) - 2)
    
    DecResult = Val(H2) + Val(H1)
    
    HexResult = Hex(DecResult)
    
    HexAdd = "0x" & HexResult
End Function

Function HexRound(ByVal H1 As String, UpDown As Boolean) As String
    Dim Dec As Double, DecMod As Double
    Dim ModVal As Long
    Dim HexResult As String
    
    If (Left(H1, 2) = "0x") Then
        H1 = Right(H1, Len(H1) - 2)
    End If

    H1 = "&H" & H1

    Dec = Val(H1)
    
    If (Not UpDown) Then
        Dec = Dec + 1
    End If
        
    ModVal = Dec Mod 4
    If (ModVal < 0) Then
        ModVal = ModVal + 4
    End If
    
    If (ModVal <> 0) Then
        If (UpDown) Then
            DecMod = Dec + (4 - ModVal)
        Else
            DecMod = Dec - ModVal
        End If
    Else
        DecMod = Dec
    End If
    
    HexResult = Hex(DecMod)
    
    HexRound = "0x" & HexResult
End Function

Function Hex2Lng(ByVal H1 As String) As Long
    Hex2Lng = CLng(Replace(H1, "0x", "&H"))
End Function

Function RegExpFind(FindIn As String, FindWhat As String, Optional IgnoreCase As Boolean = False) As String()
    Dim i As Long
    Dim RE As RegExp, allMatches As MatchCollection
    Dim rslt() As String
    
    Set RE = New RegExp
    
    RE.Pattern = FindWhat
    RE.IgnoreCase = IgnoreCase
    RE.Global = True
    
    Set allMatches = RE.Execute(FindIn)
    
    If (allMatches.Count > 0) Then
        ReDim rslt(0 To allMatches.Count - 1)
        For i = 0 To allMatches.Count - 1
            rslt(i) = allMatches(i).Value
        Next i
    Else
        ReDim rslt(0 To 0)
        rslt(0) = ""
    End If
    RegExpFind = rslt

End Function

Function DetermineProcessor(ByVal MAPFilePath As String) As Long
    Dim LineNum As Long, Processor As Long
    Dim MyLine As String
    Dim Response As VbMsgBoxResult
    
    ' Read first few lines of MAP file to figure out what kind of processor we have
    'Open MAPFilePath For Input As #1
    
    '**************Added by TCS offshore************************
    'Create the object for the file to read it
    Dim objFSO As FileSystemObject
    Set objFSO = New FileSystemObject
    'Open the file using the object created
    Dim objFile As TextStream
    Set objFile = objFSO.OpenTextFile(MAPFilePath, 1)
    
    LineNum = 1
    Processor = 0
    'Check until the end of the file
    Do Until objFile.AtEndOfStream
    'Do While Not EOF(1)
        'Line Input #1, MyLine
        'Read the line
        MyLine = objFile.ReadLine
        If InStr(MyLine, "Allocating bit symbols") Then
            Processor = 1 ' Bosch TriCore
            Exit Do
        ElseIf InStr(MyLine, "Image Summary") Then
            Processor = 2 ' Conti PPC
            Exit Do
        ElseIf InStr(MyLine, "Link Date") Then
            Processor = 2 ' Conti PPC
            Exit Do
        ElseIf InStr(MyLine, "Processed Files Part") Then
            Processor = 3 ' Conti TriCore
            Exit Do
        ElseIf InStr(MyLine, "Tool and Invocation") Then
            Processor = 4 ' Conti TriCore, option 2
            Exit Do
        ElseIf InStr(MyLine, "BEGIN EXTENDED MAP LISTING") Then
            Processor = 5 ' Bosch MG1 Multicore
            Exit Do
        ElseIf InStr(MyLine, "Archive member included") Then
            Processor = 6 ' Conti TC1277 / UTCU3
            Exit Do
        End If
    
        LineNum = LineNum + 1
        If (LineNum > 6) Then
            Exit Do
        End If
    Loop
    'Close the file after use
    objFile.Close
    'Close #1

    If (Len(MyLine) > 200) Then
        Response = msgLogDisp("The MAP file may not have been transfered correctly from VMS." & vbCrLf & "As a result, the following steps may be very slow and not work correctly." & _
        vbCrLf & "Do you wish to continue?", vbYesNo, , vbNo)
        If (Response = vbNo) Then
            DetermineProcessor = 0
            Exit Function
        End If
    End If

    DetermineProcessor = Processor

End Function

Sub FindMemoryForNoHooks(ByVal MAPFilePath As String, ByRef Strategy As Object)
    Dim Processor As Long, x As Long, FirstNum As Long, y As Long
    Dim Ret As VISION_STATUS_CODES
    Dim RamStart As String, RamEnd As String, CodeStart As String, CodeEnd As String, CalStart As String, CalEnd As String, CalStart2 As String, CalEnd2 As String, SmallData As String, SmallData2 As String, SmallData3 As String, SmallData4 As String, CodeLength As String, RamLength As String, CalLength As String, CalLength2 As String, EntryName As String, EntryValue As String
    Dim msgret As VbMsgBoxResult
    Dim Found As Boolean

    ' First figure out if TriCore or PPC
    Progress.Label2.caption = "Determining processor type..."
    Progress.Repaint
    Processor = DetermineProcessor(MAPFilePath)
    
    Progress.Label2.caption = "Searching for addresses..."
    Progress.Repaint
    Select Case Processor
        Case 0 ' Not Determined
            msgLogDisp "Can't determine type of processor.  NoHooks settings failed.", vbExclamation
            Exit Sub
        
        Case 1 ' Bosch TriCore
            ' Use EDRAM
            RamStart = "8FF04000"
            RamEnd = "8FF08000"
            'RamStart = GetAddressFromMap(MAPFilePath, "_ASW_RAM0_FREE_START", Processor)
            'RamEnd = GetAddressFromMap(MAPFilePath, "_ASW_RAM0_FREE_END", Processor)
            CodeStart = GetAddressFromMap(MAPFilePath, "_ASW0_FREE_START", Processor)
            CodeEnd = GetAddressFromMap(MAPFilePath, "_ASW0_FREE_END", Processor)
            CalStart = GetAddressFromMap(MAPFilePath, "_DS0_FREE_START", Processor)
            CalEnd = GetAddressFromMap(MAPFilePath, "_DS0_FREE_END", Processor)
            CalStart2 = GetAddressFromMap(MAPFilePath, "__PTA_DATA_END", Processor)
            CalEnd2 = GetAddressFromMap(MAPFilePath, "__PTA_DATA_END", Processor, True, True)
            SmallData = GetAddressFromMap(MAPFilePath, "_SMALL_DATA_", Processor)
            SmallData2 = GetAddressFromMap(MAPFilePath, "_SMALL_DATA2_", Processor)
        
        Case 2 ' PPC
            RamStart = "200A0000"
            RamEnd = "200A2000"
            CodeStart = GetAddressFromMap(MAPFilePath, "ROM_3_End", Processor)
            CodeEnd = GetAddressFromMap(MAPFilePath, "CAL_Beg", Processor)
            CalStart = GetAddressFromMap(MAPFilePath, "CAL_End", Processor)
            CalEnd = GetAddressFromMap(MAPFilePath, "IFLASH_End", Processor)
            
        Case 3 ' Conti TriCore
            ' Using same block of EDRAM as Bosch module, but using non-cached due to suggestion from ATI
            RamStart = "AFF04000"
            RamEnd = "AFF08000"
            ' This region is not safe as it may be used for stack space
            'RamStart = GetAddressFromMap(MAPFilePath, "_DEF_LDRAM_CSA_BEG", Processor)
            'RamEnd = GetAddressFromMap(MAPFilePath, "_DEF_LDRAM_CSA_END", Processor)
            CodeStart = GetAddressFromMap(MAPFilePath, "fmc_code_csum_end", Processor)
            CodeEnd = GetAddressFromMap(MAPFilePath, "_lc_gb_ecu_rom__ecu_sig", Processor)
            CalStart = GetAddressFromMap(MAPFilePath, "fmc_cal_csum_end", Processor)
            CalEnd = GetAddressFromMap(MAPFilePath, "_lc_gb_cal_rom__susp", Processor)
            SmallData = GetAddressFromMap(MAPFilePath, "_SMALL_DATA_", Processor)
            
        Case 4 ' Conti TriCore EMS24xx
            ' Using second block of EDRAM because ECU seems to be using first block for ShadowTables
            RamStart = "BF010000"
            RamEnd = "BF020000"
            ' Flash 2 area
            'CodeStart = "809A0000"
            'CodeEnd = "809C0000"
            'MsgBox "Using a fixed code memory region as there are no good markers available in the MAP file to determine available space.  Please manually verify that 0x809A0000-0x809C0000 does not contain any data and adjust as necessary. A memory region starting at 0x808D0000 may be another option.", vbExclamation
            CodeStart = GetAddressFromMap(MAPFilePath, "_lc_ge_ram_init_const", Processor)
            CodeEnd = GetAddressFromMap(MAPFilePath, "_lc_u_ECU_ROM_END", Processor)
            CalStart = GetAddressFromMap(MAPFilePath, "_lc_ge_cal_rom__cal", Processor)
            CalEnd = GetAddressFromMap(MAPFilePath, "_lc_gb_cal_rom__susp", Processor)
            SmallData = GetAddressFromMap(MAPFilePath, "_SMALL_DATA_", Processor)
            SmallData2 = GetAddressFromMap(MAPFilePath, "_LITERAL_DATA_", Processor)
            
        Case 5 ' Bosch MG1 Multicore
            ' 2014-09-16 Starting with settings similar to Bosch TriCore, but these have not been tested
            ' 2015-04-28 Updating based on email from Jeff Fry and his work with ATI
            ' 4/5/19 JWHIT434 update: Jira issue PCCNDATA-231. Updated RamStart and RamEnd values to auto-update from the user selected .map file relating to Jim Mitchell's information.
            msgLogDisp "No Hooks settings for this processor are purely experimental and not proven at this point.", vbExclamation
            'RamStart = "0xB9000000"
            'RamEnd = "0xB9008000"
            RamStart = GetAddressFromMap(MAPFilePath, "_EDRAM_DYN0_START", Processor)
            RamEnd = Left(RamStart, 6) + "8000"
            CodeStart = GetAddressFromMap(MAPFilePath, "_ASW0_FREE_START", Processor)
            CodeEnd = GetAddressFromMap(MAPFilePath, "_ASW0_FREE_END", Processor)
            CalStart = GetAddressFromMap(MAPFilePath, "_DS0_FREE_START", Processor)
            CalEnd = GetAddressFromMap(MAPFilePath, "_DS0_FREE_END", Processor)
            'CalStart2 = GetAddressFromMap(MAPFilePath, "__PTA_DATA_END", Processor)
            'CalEnd2 = GetAddressFromMap(MAPFilePath, "__PTA_DATA_END", Processor, True, True)
            SmallData = GetAddressFromMap(MAPFilePath, "_SMALL_DATA_", Processor)
            SmallData2 = GetAddressFromMap(MAPFilePath, "_SMALL_DATA2_", Processor)
            SmallData3 = GetAddressFromMap(MAPFilePath, "_SMALL_DATA3_", Processor)
            SmallData4 = GetAddressFromMap(MAPFilePath, "_SMALL_DATA4_", Processor)
            
         Case 6 ' Conti Multicore
            ' First check for FORD_RP regions in MAP. If any of these fails, will jump to next section and try other marker types
            ' Details provided by Brian Clary on 2020-09-16
            RamStart = GetAddressFromMap(MAPFilePath, "__FORD_RP_RAM_START1", Processor, False)
            If RamStart <> "0" Then
                RamEnd = GetAddressFromMap(MAPFilePath, "__FORD_RP_RAM_END1", Processor)
                CodeStart = GetAddressFromMap(MAPFilePath, "__FORD_RP_CODE_START0", Processor)
                CodeEnd = GetAddressFromMap(MAPFilePath, "__FORD_RP_CODE_END0", Processor)
                CalStart = GetAddressFromMap(MAPFilePath, "__FORD_RP_CAL_START", Processor)
                CalEnd = GetAddressFromMap(MAPFilePath, "__FORD_RP_CAL_END", Processor)
            Else
                ' Using first block of EDRAM per Jim Mitchell's suggestion
                RamStart = "0xBF000000"
                RamEnd = "0xBF008000"
                CodeStart = GetAddressFromMap(MAPFilePath, "_ASW0_FREE_START", Processor)
                CodeEnd = GetAddressFromMap(MAPFilePath, "_ASW0_FREE_END", Processor)
                CalStart = GetAddressFromMap(MAPFilePath, "_DS0_FREE_START", Processor)
                CalEnd = GetAddressFromMap(MAPFilePath, "_DS0_FREE_END", Processor)
            End If
            SmallData = GetAddressFromMap(MAPFilePath, "_SMALL_DATA_", Processor)
            SmallData2 = GetAddressFromMap(MAPFilePath, "_SMALL_DATA2_", Processor)
            SmallData3 = GetAddressFromMap(MAPFilePath, "_SMALL_DATA3_", Processor)
            SmallData4 = GetAddressFromMap(MAPFilePath, "_SMALL_DATA4_", Processor)
    End Select
    
    ' If we found all the regions we need, continue.  Otherwise abort here.
    If Not ((CodeStart <> "0") And (CodeEnd <> "0") And (RamStart <> "0") And (RamEnd <> "0") And (((CalStart <> "0") And (CalEnd <> "0")) Or ((CalStart2 <> "0") And (CalEnd2 <> "0")))) Then
        msgLogDisp "Can't find required regions in MAP file", vbExclamation
        Exit Sub
    Else
        ' Round regions to 4 byte boundaries
        CodeStart = HexRound(CodeStart, True)
        CodeEnd = HexRound(CodeEnd, False)
        RamStart = HexRound(RamStart, True)
        RamEnd = HexRound(RamEnd, False)
        CalStart = HexRound(CalStart, True)
        CalEnd = HexRound(CalEnd, False)
        CalStart2 = HexRound(CalStart2, True)
        CalEnd2 = HexRound(CalEnd2, False)
        
        ' Calculate region lengths
        CodeLength = HexSubtract(CodeStart, CodeEnd)
        RamLength = HexSubtract(RamStart, RamEnd)
        CalLength = HexSubtract(CalStart, CalEnd)
        CalLength2 = HexSubtract(CalStart2, CalEnd2)
        
        ' For Bosch TriCore, query user which CAL region to use
        If (Processor = 1) Then
            msgret = msgLogDisp("Would you like to use the space after PTA_DATA_END for CalData rather than the normal free space definition?" & Chr(10) _
                & "Yes: PTA_DATA_END to next address   = " & Val("&H" & Right(CalLength2, Len(CalLength2) - 2)) & " bytes" & Chr(10) _
                & "No:  DSO_FREE_START to DSO_FREE_END = " & Val("&H" & Right(CalLength, Len(CalLength) - 2)) & " bytes", vbQuestion + vbYesNo, , vbNo)
            If (msgret = vbYes) Then
                CalStart = CalStart2
                CalEnd = CalEnd2
                CalLength = CalLength2
            End If
        End If
        
        ' Display memory available
        msgLogDisp "Available Memory For NoHooks:" & Chr(10) & Chr(10) _
            & "Code Region: " & Val("&H" & Right(CodeLength, Len(CodeLength) - 2) & "&") & " Bytes" & Chr(10) _
            & "Cal Region: " & Val("&H" & Right(CalLength, Len(CalLength) - 2) & "&") & " Bytes" & Chr(10) _
            & "RAM Region: " & Val("&H" & Right(RamLength, Len(RamLength) - 2) & "&") & " Bytes" & Chr(10) & Chr(10), _
            vbInformation
                
        ' Import MAP file
        Progress.Label2.caption = "Importing MAP File..."
        Progress.Repaint
        Ret = Strategy.Import(MAPFilePath)
            If (Ret <> VISION__OK) Then
            msgLogDisp "Error importing MAP file."
            Exit Sub
        End If
    
        ' Set Hook Params
        Progress.Label2.caption = "Configuring No Hooks Settings..."
        Progress.Repaint
        
        Dim errMsg As Dictionary
        Set errMsg = New Dictionary
    
        Ret = Strategy.SetStrategySetting("ReloCALRegion", CalStart & "," & CalLength & ",0")
        If (Ret <> VISION__OK) Then errMsg.Add errMsg.Count, "ReloCALRegion"
        Ret = Strategy.SetStrategySetting("ReloRAMRegion", RamStart & "," & RamLength)
        If (Ret <> VISION__OK) Then errMsg.Add errMsg.Count, "ReloRAMRegion"
        Ret = Strategy.SetStrategySetting("ReloCodeRegion", CodeStart & "," & CodeLength)
        If (Ret <> VISION__OK) Then errMsg.Add errMsg.Count, "ReloCodeRegion"
        
        If (Processor = 1) Or (Processor = 3) Or (Processor = 4) Or (Processor = 5) Or (Processor = 6) Then
            ' TriCore
            Ret = Strategy.SetStrategySetting("NoHooksProcessorType", "TRICORE")
            If (Ret <> VISION__OK) Then errMsg.Add errMsg.Count, "NoHooksProcessorType"
            ' SDA_BASE
            Ret = Strategy.SetStrategySetting("NoHooks.SDA_BASE", SmallData)
            If (Ret <> VISION__OK) Then errMsg.Add errMsg.Count, "NoHooks.SDA_BASE"
            ' Using EDRAM, so tell A7 not to use it
            If (Processor = 1) Or (Processor = 3) Then
                Ret = Strategy.SetStrategySetting("A7_OverlayRegisterMask", "0x3")
                If (Ret <> VISION__OK) Then errMsg.Add errMsg.Count, "A7_OverlayRegisterMask"
            ElseIf (Processor = 4) Then
                Ret = Strategy.SetStrategySetting("A7_OverlayRegisterMask", "0x2")
                If (Ret <> VISION__OK) Then errMsg.Add errMsg.Count, "A7_OverlayRegisterMask"
            ElseIf (Processor = 5) Then
                Ret = Strategy.SetStrategySetting("A_OverlayRegisterMask", "0xC0000001")
                ' I'm being lazy here. Newer processors can use 0x80000001. Older processors need to mask off an extra block (0xC0000001)
                ' Just using the "C" version because it's safer and easier than figuring out which is which.
                'Ret = Strategy.SetStrategySetting("A_OverlayRegisterMask", "0x80000001")
                If (Ret <> VISION__OK) Then errMsg.Add errMsg.Count, "A_OverlayRegisterMask"
            ElseIf (Processor = 6) Then
                Ret = Strategy.SetStrategySetting("A_OverlayRegisterMask", "0x1")
                If (Ret <> VISION__OK) Then errMsg.Add errMsg.Count, "A_OverlayRegisterMask"
            End If
            If (Processor = 3) Or (Processor = 4) Then
                msgLogDisp "Warning:  NoHooks is using EDRAM which may not be readable using synchronous data acquisition.  You may need to switch to asynchronous to read RAM data.", vbExclamation
            End If
            ' SDA2_BASE
            If (Processor = 1) Or (Processor = 4) Or (Processor = 5) Or (Processor = 6) Then
                Ret = Strategy.SetStrategySetting("NoHooks.SDA2_BASE", SmallData2)
                If (Ret <> VISION__OK) Then errMsg.Add errMsg.Count, "NoHooks.SDA2_BASE"
            End If
            ' SDA3_BASE
            If (Processor = 5) Or (Processor = 6) Then
                Ret = Strategy.SetStrategySetting("NoHooks.SDA3_BASE", SmallData3)
                If (Ret <> VISION__OK) Then errMsg.Add errMsg.Count, "NoHooks.SDA3_BASE"
            End If
            ' SDA4_BASE
            If (Processor = 5) Or (Processor = 6) Then
                Ret = Strategy.SetStrategySetting("NoHooks.SDA4_BASE", SmallData4)
                If (Ret <> VISION__OK) Then errMsg.Add errMsg.Count, "NoHooks.SDA4_BASE"
            End If
            
        ElseIf (Processor = 2) Then
            ' PPC
            Ret = Strategy.SetStrategySetting("NoHooksProcessorType", "PPC")
            If (Ret <> VISION__OK) Then errMsg.Add errMsg.Count, "NoHooksProcessorType"
            Ret = Strategy.SetStrategySetting("TargetRAMSize", "0x2000")
            If (Ret <> VISION__OK) Then errMsg.Add errMsg.Count, "TargetRAMSize"
            Ret = Strategy.SetStrategySetting("TargetRAMAddress", "0x200A0000")
            If (Ret <> VISION__OK) Then errMsg.Add errMsg.Count, "TargetRAMAddress"
        
            Strategy.AddMemoryRegion "Model_ROM", CodeStart, CodeLength, "FUT"
            Strategy.AddMemoryRegion "Model_RAM", RamStart, RamLength, "R"
            Strategy.AddMemoryRegion "Model_CAL", CalStart, CalLength, "FDUT"
        
            ' Additional settings for XCP
            Strategy.GetDeviceSettingString "XCP_AddressMapping1", Ret
            If (Ret = VISION__OK) Then
                ' Must be XCP because the device setting exists
                
                ' Find last address mapping
                Found = False
                x = 1
                Do While (Not Found)
                    Strategy.GetDeviceSettingString "XCP_AddressMapping" & x, Ret
                    If (Ret <> VISION_OK) Then
                        Found = True
                        FirstNum = x
                    Else
                        x = x + 1
                    End If
                Loop
                    
                Dim RegionNames(3) As String
                Dim StartAddress(3) As String
                Dim Length(3) As String
                    
                RegionNames(1) = "Model_ROM"
                RegionNames(2) = "Model_RAM"
                RegionNames(3) = "Model_CAL"
                StartAddress(1) = CodeStart
                StartAddress(2) = RamStart
                StartAddress(3) = CalStart
                Length(1) = CodeLength
                Length(2) = RamLength
                Length(3) = CalLength
                
                For y = 1 To 3
                    x = y + FirstNum - 1
                    Strategy.SetStrategySetting "XCP_AddressMapping" & x, StartAddress(y) & "," & StartAddress(y) & "," & Length(y)
                    
                    'Set the XcpSegment#.... settings
                    EntryName = "XcpSegment" & x & "Checksum"
                    EntryValue = "6,0," '6=XCP_ADD_44
                    Strategy.SetStrategySetting EntryName, EntryValue
                    
                    EntryName = "XcpSegment" & x & "Config"
                    EntryValue = RegionNames(y) & "," & (x - 1) & ",1,2,0"
                    Strategy.SetStrategySetting EntryName, EntryValue
                    
                    EntryName = "XcpSegment" & x & "Mapping1"
                    EntryValue = StartAddress(y) & "," & StartAddress(y) & "," & Length(y)
                    Strategy.SetStrategySetting EntryName, EntryValue
                    
                    EntryName = "XcpSegment" & x & "NumMappings"
                    EntryValue = "1"
                    Strategy.SetStrategySetting EntryName, EntryValue
                    
                    EntryName = "XcpSegment" & x & "NumPages"
                    EntryValue = "1"
                    Strategy.SetStrategySetting EntryName, EntryValue
                    
                    EntryName = "XcpSegment" & x & "Page1"
                    EntryValue = "0,3,3,3,0"
                    Strategy.SetStrategySetting EntryName, EntryValue
                    
                    EntryName = "XcpSegment" & x & "PgmVerify"
                    EntryValue = "0"
                    Strategy.SetStrategySetting EntryName, EntryValue
                Next y
            End If ' XCP Module
        End If ' processor type
        If errMsg.Count > 0 Then msgLogDisp "Error setting:" & vbCrLf & Join(errMsg.Items, vbCrLf)
    End If ' Found memory regions
End Sub
