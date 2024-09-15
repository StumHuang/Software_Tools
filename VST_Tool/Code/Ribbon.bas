Attribute VB_Name = "Ribbon"
Option Explicit

Public Rib As IRibbonUI
Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set Rib = ribbon
    Application.OnTime Now + TimeValue("00:00:01"), "DoIt"
End Sub

Sub DoIt()
    Rib.ActivateTab "TabVstTool"
End Sub

'Callback for customButton1 onAction
Sub NewVstRibbon(control As IRibbonControl)
    BuildVSTFile False, True
End Sub

'Callback for customButton2 onAction
Sub BatchVstRibbon(control As IRibbonControl)
    BuildVSTFile True, True
End Sub

'Callback for customButton3 onAction
Sub UpdateVstRibbon(control As IRibbonControl)
    BatchVstUpdate
End Sub

'Callback for customButton4 onAction
Sub CopySettingsRibbon(control As IRibbonControl)
    CopySettings
End Sub

'Callback for customButton5 onAction
Sub ClearSettingsRibbon(control As IRibbonControl)
    ClearSettings
End Sub

'Callback for label
Sub getCurrentVersion(control As IRibbonControl, ByRef returnedVal As Variant)
    Dim CurrVer As String
    CurrVer = DateValToStr(GetCurrentVersionVal)
    returnedVal = "Current Version: " & CurrVer
End Sub

'Callback for label
Sub getLatestVersion(control As IRibbonControl, ByRef returnedVal As Variant)
    Dim TmpVer As Double
    Dim CurrVer As String
    TmpVer = GetWebVersionVal
    If (TmpVer > 0) Then
        CurrVer = DateValToStr(GetWebVersionVal)
    Else
        CurrVer = "Unknown"
    End If
    returnedVal = "Latest Version: " & CurrVer
End Sub

'Callback for label
Sub getUpdateStatus(control As IRibbonControl, ByRef returnedVal As Variant)
    Dim CurrVer As Double, WebVer As Double
    CurrVer = GetCurrentVersionVal
    WebVer = GetWebVersionVal
    If (WebVer < 1) Then
        returnedVal = "Error checking for updates"
    ElseIf (WebVer > CurrVer) Then
        returnedVal = "Update Available!"
    Else
        returnedVal = "Up-to-date"
    End If
End Sub

' Utility function
Function DateValToStr(DateVal As Double) As String
    Dim CurYear As Long, CurMonth As Long, CurDay As Long
    CurYear = Int(DateVal / 10000)
    CurMonth = Int((DateVal - (CurYear * 10000)) / 100)
    If (CurMonth < 10) Then CurMonth = "0" & CurMonth
    CurDay = DateVal - CurYear * 10000 - CurMonth * 100
    If (CurDay < 10) Then CurDay = "0" & CurDay
    DateValToStr = "20" & CurYear & "-" & CurMonth & "-" & CurDay
End Function
