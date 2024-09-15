Attribute VB_Name = "Log"
Option Explicit

Public AutomatedMode As Boolean
Public debugLog As String

'either log or display a message depending on whether the tool is being run manually or automated
Public Function msgLogDisp(ByVal msgText As String, Optional ByVal msgStyle As VbMsgBoxStyle = vbOKOnly, Optional ByVal msgCaption As String = vbNullString, Optional ByVal defaultResponse As Variant) As VbMsgBoxResult
    Dim resp As String
    If AutomatedMode Then
        debugLog = debugLog & msgText & vbCrLf
        If Not IsMissing(defaultResponse) Then
            Select Case defaultResponse
            Case vbOK
                resp = "OK"
            Case vbCancel
                resp = "Cancel"
            Case vbAbort
                resp = "Abort"
            Case vbRetry
                resp = "Retry"
            Case vbIgnore
                resp = "Ignore"
            Case vbYes
                resp = "Yes"
            Case vbNo
                resp = "No"
            End Select
            debugLog = debugLog & vbCrLf & "Reponse: " & resp & vbCrLf & vbCrLf
            msgLogDisp = defaultResponse
        Else
            msgLogDisp = vbOK
        End If
    Else
        msgLogDisp = MsgBox(msgText, msgStyle, msgCaption)
    End If
End Function
