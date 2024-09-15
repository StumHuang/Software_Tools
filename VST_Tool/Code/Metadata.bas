Attribute VB_Name = "Metadata"
Option Explicit

Public Type H32Meta
    UserName As String
    ImageName As String
    StrategyName As String
    TimeStamp As String
    Group As String
    Project As String
    Comment As String
End Type

Sub AddH32Meta(ByVal H32FilePath As String, Strategy As Object)
    Dim MyH32Meta As H32Meta

    MyH32Meta = ReadMetaFromH32(H32FilePath, Strategy)
    Strategy.SetStrategySetting "FMC_UserName", MyH32Meta.UserName
    Strategy.SetStrategySetting "FMC_ImageName", MyH32Meta.ImageName
    Strategy.SetStrategySetting "FMC_StrategyName", MyH32Meta.StrategyName
    Strategy.SetStrategySetting "FMC_TimeStamp", MyH32Meta.TimeStamp
    Strategy.SetStrategySetting "FMC_Group", MyH32Meta.Group
    Strategy.SetStrategySetting "FMC_Project", MyH32Meta.Project
    Strategy.SetStrategySetting "FMC_Comment", MyH32Meta.Comment
End Sub

Sub TestReadH32()
    Dim H32FilePath As Variant
    ' Choose File
    H32FilePath = Application.GetOpenFilename("H32 File (*.h32),*.h32", 1, "Select H32 File")
    If (H32FilePath = False) Then
        Exit Sub
    End If

    ReadMetaFromH32 H32FilePath, Nothing
End Sub

Function ReadMetaFromH32(ByVal H32FilePath As String, Strategy As Object) As H32Meta
    Dim H32File As TextStream
    Dim MyH32Meta As H32Meta
    Dim FileSys As FileSystemObject
    Dim InMeta As Boolean, InMetaInfo As Boolean
    Dim MyLine As String, YearTemp As String, DayTemp As String, MonthStringTemp As String, MonthTemp As String, MetaDataText As String, writeData As String, valueData As String
    Dim cfxMetaData As Variant
    Dim iMeta As Long
    
    Set FileSys = New FileSystemObject
    Set H32File = FileSys.OpenTextFile(H32FilePath, 1)

    InMeta = False
    InMetaInfo = False
    Do While Not H32File.AtEndOfStream
        MyLine = H32File.ReadLine
        If (Not InMeta And (Left(MyLine, 10) = "User name:")) Then
            InMeta = True
        End If
        
        If (InMeta) And InMetaInfo = False Then
            If (MyStrComp(MyLine, "User name:", True)) Then
                MyH32Meta.UserName = Mid(MyLine, 12)
            ElseIf (MyStrComp(MyLine, " Image Name:", True)) Then
                MyH32Meta.ImageName = Mid(MyLine, 17)
            ElseIf (MyStrComp(MyLine, " Strategy Name:", True)) Then
                MyH32Meta.StrategyName = Mid(MyLine, 17)
            ElseIf (MyStrComp(MyLine, " Time Stamp:", True)) Then
                YearTemp = Mid(MyLine, 37, 4)
                DayTemp = Mid(MyLine, 25, 2)
                If (Left(DayTemp, 1) = " ") Then
                    DayTemp = "0" & Right(DayTemp, 1)
                End If
                MonthStringTemp = Mid(MyLine, 21, 3)
                Select Case MonthStringTemp
                    Case "Jan"
                        MonthTemp = "01"
                    Case "Feb"
                        MonthTemp = "02"
                    Case "Mar"
                        MonthTemp = "03"
                    Case "Apr"
                        MonthTemp = "04"
                    Case "May"
                        MonthTemp = "05"
                    Case "Jun"
                        MonthTemp = "06"
                    Case "Jul"
                        MonthTemp = "07"
                    Case "Aug"
                        MonthTemp = "08"
                    Case "Sep"
                        MonthTemp = "09"
                    Case "Oct"
                        MonthTemp = "10"
                    Case "Nov"
                        MonthTemp = "11"
                    Case "Dec"
                        MonthTemp = "12"
                End Select
                MyH32Meta.TimeStamp = YearTemp & MonthTemp & DayTemp
            ElseIf (MyStrComp(MyLine, "        Group:", True)) Then
                MyH32Meta.Group = Mid(MyLine, 18)
            ElseIf (MyStrComp(MyLine, "        Project:", True)) Then
                MyH32Meta.Project = Mid(MyLine, 18)
            ElseIf (MyStrComp(MyLine, "        Comment:", True)) Then
                MyH32Meta.Comment = Mid(MyLine, 18)
            ElseIf (MyStrComp(Trim(MyLine), "Metadata Information:", True)) Then
                InMetaInfo = True
            End If
        End If

        If (InMetaInfo) Then
            If MyLine <> "" And Left(MyLine, 1) <> " " Then
                InMetaInfo = False
            Else
                MetaDataText = MetaDataText & MyLine
            End If
        End If
    Loop
    
    ' Parse CFX Metadata
    cfxMetaData = RegExpTokens(MetaDataText, "(\w+) has \d+ assignments? : ('.+?')\s", False, True)
    If IsArray(cfxMetaData) Then
        For iMeta = 0 To UBound(cfxMetaData, 1)
            writeData = "FMC_" & cfxMetaData(iMeta, 0)
            valueData = Replace(cfxMetaData(iMeta, 1), " ", "")
            valueData = Replace(valueData, "'", "")
            'Debug.Print writeData & " = " & valueData
            Strategy.SetStrategySetting writeData, valueData
        Next
    End If
    
    ReadMetaFromH32 = MyH32Meta
End Function

Function MyStrComp(String1 As String, String2 As String, Optional UseLength As Boolean = False) As Boolean
    If (UseLength = False) Then
        MyStrComp = (StrComp(String1, String2) = 0)
    Else
        MyStrComp = (StrComp(Left(String1, Len(String2)), String2) = 0)
    End If
End Function

Function RegExpTokens(ByVal FindIn As String, ByVal FindWhat As String, Optional IgnoreCase As Boolean = False, Optional GlobalSearch As Boolean = False) As Variant
    Dim i As Long
    Dim J As Long
    Dim RE As RegExp, allMatches As MatchCollection
    Dim rslt1d As Variant, rslt As Variant
    
    Set RE = New RegExp
    
    RE.Pattern = FindWhat
    RE.IgnoreCase = IgnoreCase
    RE.Global = GlobalSearch
    
    Set allMatches = RE.Execute(FindIn)
    
    If (allMatches.Count = 1) Then
        ReDim rslt1d(0 To allMatches(0).SubMatches.Count - 1)
        For i = 0 To allMatches(0).SubMatches.Count - 1
            rslt1d(i) = allMatches(0).SubMatches(i)
        Next i
        RegExpTokens = rslt1d
    ElseIf (allMatches.Count > 1) Then
        ReDim rslt(0 To allMatches.Count - 1, 0 To allMatches(0).SubMatches.Count - 1)
        For i = 0 To allMatches.Count - 1
            For J = 0 To allMatches(0).SubMatches.Count - 1
                If J > UBound(rslt, 2) Then
                    ReDim Preserve rslt(0 To allMatches.Count - 1, 0 To allMatches(i).SubMatches.Count)
                End If
                rslt(i, J) = allMatches(i).SubMatches(J)
            Next J
        Next i
        RegExpTokens = rslt
    Else
        ReDim rslt1d(0 To 0)
        rslt1d(0) = ""
        RegExpTokens = ""
    End If
End Function
