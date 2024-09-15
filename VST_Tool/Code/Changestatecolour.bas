Attribute VB_Name = "Changestatecolour"
Option Explicit

Sub changecolour(ByRef Strategy As Variant, BatchMode As Boolean, ByRef ShowMisMatchError As Boolean)
    Dim rownumber As Long, GroupNum As Long, ip As Long
    Dim WorkSheetName As String, parametername As String
    Dim parametervalue As Double, textcolour As Double, backgrndcolor As Double
    Dim MeasurementGroup As Object, temphandle As Object, stateitemdata As Object
    ' Initialize the row number
    rownumber = 2
    ' initialize the sheet name
    WorkSheetName = "State Var Colors"
    Progress.Label1.caption = "Modifying State Var Colors"
    Progress.Label2.caption = ""
    Progress.Repaint
    'Loop through all the parametrs and update the color accordingly
    On Error GoTo ErrorHandler
    Do While (Worksheets(WorkSheetName).Cells(rownumber, 1).Value <> "")
       ' Get the parametername,text color and background color
       parametername = Worksheets(WorkSheetName).Cells(rownumber, 1).Value
       parametervalue = Worksheets(WorkSheetName).Cells(rownumber, 2).Value
       textcolour = Worksheets(WorkSheetName).Cells(rownumber, 3).Value
       backgrndcolor = Worksheets(WorkSheetName).Cells(rownumber, 4).Value
       'Get the measurement group
       For GroupNum = 1 To Strategy.GroupDataItem.Items.Count
            If (Strategy.GroupDataItem.Items(GroupNum).DataItemName = "Measurements") Then
                Set MeasurementGroup = Strategy.GroupDataItem.Items(GroupNum)
                Exit For
            End If
       Next
       'Check if parameter present
       Set temphandle = MeasurementGroup.FindDataItem(parametername)
       ' Update the color data accordingly
       If Not (temphandle Is Nothing) Then
            For ip = 1 To temphandle.StateTable.Items.Count
                Set stateitemdata = temphandle.StateTable.Items.Item(ip)
                If stateitemdata.MinimumValue <= parametervalue And stateitemdata.MaximumValue >= parametervalue Then
                    stateitemdata.BackgroundColor = backgrndcolor
                    stateitemdata.TextColor = textcolour
                End If
            Next ip
       End If
       ' update the rrownumber
       rownumber = rownumber + 1
    Loop
    Exit Sub
ErrorHandler:
    If BatchMode Then
        If ShowMisMatchError Then
            msgLogDisp "State Var Color value is invalid on row " & CStr(rownumber), vbCritical, "Change State Var Colors"
            Dim resp As VbMsgBoxResult
            resp = msgLogDisp("Do you want to ignore State Var Color error for all batch build files", vbYesNo, "Skip Errors and Continue", vbNo)
            If resp = vbYes Then
                ShowMisMatchError = False
            End If
        Else
            If AutomatedMode Then
                msgLogDisp "State Var Color value is invalid on row " & CStr(rownumber), vbCritical, "Change State Var Colors"
            End If
        End If
    Else
        msgLogDisp "State Var Color value is invalid on row " & CStr(rownumber), vbCritical, "Change State Var Colors"
    End If
End Sub
