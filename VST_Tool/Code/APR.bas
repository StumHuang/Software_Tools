Attribute VB_Name = "APR"
Option Explicit

Sub DoAprs(StrategyHandle As Object, A2lPath As String)
    Dim success As Boolean, Fetch As Boolean, Found As Boolean
    Dim LineTotal As Long, LineNum As Long
    Dim LineTmp As String
    Dim ff As Long
    Dim A2lFile As String
    Dim CharName() As String, LineArrayTmp() As String
    Dim CurrentParameterName As String
    Dim i As Long
    Dim MinVal As Double, MaxVal As Double
    Dim Strategy As Object
    Set Strategy = StrategyHandle
            
    Dim AprPath As String
    Dim DataItem As Object
    
    ' Check for path for APR file
    If (ThisWorkbook.Worksheets("File Paths").Range("B16").Value = "") Then
        ActivateExcel
        msgLogDisp "To use the APR feature, you need to select a location to save a local copy of APR values.  You may save the file anywhere on your local drive.", vbInformation
        If Not AutomatedMode Then AprPath = Application.GetSaveAsFilename("apr_summary.csv", "CSV Files (*.csv), *.csv", 1, "Select APR save file location") Else AprPath = False
        If (AprPath <> "False") Then
            ThisWorkbook.Worksheets("File Paths").Range("B16").Value = AprPath
        End If
    Else
        AprPath = ThisWorkbook.Worksheets("File Paths").Range("B16").Value
    End If
            
    ' Download file
    success = True
    If (Dir(AprPath) = "") Then
        Fetch = True
    ElseIf (FileDateTime(AprPath) < (Now - 1)) Then
        Fetch = True
    Else
        Fetch = False
    End If
    If Fetch Then
        Progress.Label2.caption = "Fetching Latest APR List"
        Progress.Repaint
        success = GetAprSummary
    End If
    
    ' Warn if failed to download
    If Not success Then
        Progress.Label2.caption = "Failed to Fetch Latest APR List"
        Progress.Repaint
        Application.Wait (Now + TimeValue("0:00:05"))
    End If
    
    ' Make sure the file exists
    If Dir(AprPath) = "" Then
        Progress.Label1.caption = "Please respond to open message box..."
        Progress.Label2.caption = ""
        Progress.Repaint
        ActivateExcel
        msgLogDisp "Could not find list of APRs.  Skipping APRs."
        Exit Sub
    ElseIf (FileDateTime(AprPath) < (Now - 14)) Then
        ActivateExcel
        msgLogDisp "Warning: APR list is over 2 weeks old.  You should connect to the network to get a new copy.", vbExclamation
    End If
    
    ' Read APR file to count total lines.  Used for progress box.
    Open AprPath For Input As 1
    Do While Not EOF(1)
        LineTotal = LineTotal + 1
        Line Input #1, LineTmp
    Loop
    Seek 1, 1
    
    ' Open A2L file.  Pre-filter the APRs to only parameters in the strategy.
    ff = FreeFile
    Open A2lPath For Input As #ff
    
    ' Read it all
    A2lFile = Input$(LOF(ff), ff)
    
    ' Find characteristics
    CharName = RegExpFind(A2lFile, "begin CHARACTERISTIC \S+", True)
    Close ff
    ' Strip off "begin characterisitic"
    For i = 0 To UBound(CharName)
        CharName(i) = Mid(CharName(i), 22, Len(CharName(i)) - 21)
    Next i

    ' Loop through the APR list
    Do While Not EOF(1)
        
        LineNum = LineNum + 1
        If ((LineNum Mod 50) = 0) Then
            Progress.Label2.caption = "#" & LineNum & " / " & LineTotal
            Progress.Repaint
            DoEvents
        End If
        
        Line Input #1, LineTmp
        If (LineTmp <> "") Then
            LineArrayTmp = Split(LineTmp, ",")
            CurrentParameterName = LineArrayTmp(0)
            MinVal = LineArrayTmp(1)
            MaxVal = LineArrayTmp(2)
            
            ' See if its in the strategy
            Found = True
            If (UBound(Filter(CharName, CurrentParameterName)) < 0) Then Found = False
        Else
            Found = False
        End If
        
        ' If it's in the strategy, set APR range limits
        If (Found) Then
            ' Fudge values slightly to account for F32 rounding errors
            If (MinVal < 0) Then
                MinVal = MinVal * (1 + 0.0000001)
            Else
                MinVal = MinVal * (1 - 0.0000001)
            End If
            If (MaxVal < 0) Then
                MaxVal = MaxVal * (1 - 0.0000001)
            Else
                MaxVal = MaxVal * (1 + 0.0000001)
            End If
            
            
            'Find the Data Item in the strategy, and set to DataItem
            Set DataItem = Strategy.FindOrSearchDataItem(CurrentParameterName)
            
            If (Not (DataItem Is Nothing)) Then
                If (DataItem.Type = VISION_DATAITEM_SCALAR) Then
                    DataItem.MinimumThreshold = MinVal
                    DataItem.MaximumThreshold = MaxVal
                    DataItem.MinimumText = "APR"
                    DataItem.MaximumText = "APR"
                    DataItem.EnableRangeColors = True
                    DataItem.MinimumColor = 16777215
                    DataItem.MaximumColor = 16777215
                    If (LineArrayTmp(3) = 0) Then
                        DataItem.MinimumBkgrdColor = 255
                        DataItem.MaximumBkgrdColor = 255
                    Else
                        DataItem.MinimumBkgrdColor = 98303
                        DataItem.MaximumBkgrdColor = 98303
                    End If
                
                ElseIf (DataItem.Type = VISION_DATAITEM_TABLE2D) Then
                    DataItem.YAxisMinimumThreshold = MinVal
                    DataItem.YAxisMaximumThreshold = MaxVal
                    DataItem.YAxisMinimumText = "APR"
                    DataItem.YAxisMaximumText = "APR"
                    DataItem.YAxisEnableRangeColors = True
                    DataItem.YAxisMinimumColor = 16777215
                    DataItem.YAxisMaximumColor = 16777215
                    If (LineArrayTmp(3) = 0) Then
                        DataItem.YAxisMinimumBkgrdColor = 255
                        DataItem.YAxisMaximumBkgrdColor = 255
                    Else
                        DataItem.YAxisMinimumBkgrdColor = 98303
                        DataItem.YAxisMaximumBkgrdColor = 98303
                    End If
    
                ElseIf (DataItem.Type = VISION_DATAITEM_TABLE3D) Then
                    DataItem.ZAxisMinimumThreshold = MinVal
                    DataItem.ZAxisMaximumThreshold = MaxVal
                    DataItem.ZAxisMinimumText = "APR"
                    DataItem.ZAxisMaximumText = "APR"
                    DataItem.ZAxisEnableRangeColors = True
                    DataItem.ZAxisMinimumColor = 16777215
                    DataItem.ZAxisMaximumColor = 16777215
                    If (LineArrayTmp(3) = 0) Then
                        DataItem.ZAxisMinimumBkgrdColor = 255
                        DataItem.ZAxisMaximumBkgrdColor = 255
                    Else
                        DataItem.ZAxisMinimumBkgrdColor = 98303
                        DataItem.ZAxisMaximumBkgrdColor = 98303
                    End If
                End If ' DataItem.Type
            End If ' Data item found
        End If ' DataItem is Nothing
    Loop
    
    Close #1 ' Close file.

End Sub

Sub TestApr()
    Dim success As Boolean
    success = GetAprSummary
    msgLogDisp success
End Sub

Function GetAprSummary() As Boolean
    Dim AprSummaryUrl As String, AprPath As String, Data As String
    Dim ff As Long
    ' Location of APR summary document
    AprSummaryUrl = "http://www.pd2.ford.com/calibration/mscannel/apr_summary.csv"
                    
    ' Local location
    AprPath = ThisWorkbook.Worksheets("File Paths").Range("B16").Value
    If (AprPath = "") Then
        msgLogDisp "No local APR path defined", vbCritical
        GetAprSummary = False
        Exit Function
    End If
                    
    ' Create a HTTP request
    Dim WinHttpReq As WinHttpRequest
    Set WinHttpReq = New WinHttpRequest

    ' Send the HTTP Request.
    On Error GoTo TheEnd
    WinHttpReq.Open "GET", AprSummaryUrl, False
    WinHttpReq.send
    Data = WinHttpReq.ResponseText
    
    ' Write it to a file
    If (Not (Data = "")) Then
        ff = FreeFile
        Open AprPath For Output As #ff
        Print #ff, Data
        Close ff
        GetAprSummary = True
        Exit Function
    Else
        GetAprSummary = False
    End If

TheEnd:
    GetAprSummary = False
    
End Function
