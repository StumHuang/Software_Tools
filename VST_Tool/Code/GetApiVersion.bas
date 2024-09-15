Attribute VB_Name = "GetApiVersion"
Option Explicit

Function GetNumericVisionApiVersion() As Double
    Dim VersionString As String
    Dim V() As String
    Dim MultTemp As Double, Version As Double
    Dim i As Long
    On Error GoTo TheEnd

    Dim MyObj As Object
    Set MyObj = CreateObject("Vision.ProjectInterface")
    If (MyObj Is Nothing) Then
        GetNumericVisionApiVersion = 0
    Else
        VersionString = MyObj.Version
        Set MyObj = Nothing
        
        V = Split(VersionString, ".")
        
        ' Strip the "V" off the version number
        If Not IsNumeric(V(0)) Then
            V(0) = Mid(V(0), 2)
        End If
        
        MultTemp = 1000000#
        Version = 0
        For i = 0 To UBound(V)
            Version = Version + V(i) * MultTemp
            MultTemp = MultTemp / 1000
        Next i
        
        GetNumericVisionApiVersion = Version
    End If

    Exit Function

TheEnd:
    GetNumericVisionApiVersion = 0

End Function

Sub TestGetNumericVisionApiVersion()
    Dim Vers As Double
    Vers = GetNumericVisionApiVersion
    msgLogDisp Vers
End Sub

Sub CheckVisionVersion()
    Dim Vers As Double
    Vers = GetNumericVisionApiVersion
    If (Vers = 4001000) Then
        ' Issue with building VST files in Vision 4.1
        msgLogDisp "It appears you are using Vision 4.1. Vision 4.1 has a bug related to arrays when building VST files. It is not recommended to use this version of Vision for building VST files. You should revert to using an older version of Vision, or a new one (if available).", vbCritical, "VST Tool"
    End If
End Sub
