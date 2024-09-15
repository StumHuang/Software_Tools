VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub UpdateVstButton_Click()
    BatchVstUpdate
End Sub

Private Sub BatchBuildButton_Click()
    BuildVSTFile True, True
End Sub

Private Sub BuildVstButton_Click()
    BuildVSTFile False, True
End Sub

Private Sub ClearSettingsButton_Click()
    ClearSettings
End Sub

Private Sub CopySettingsButton_Click()
    CopySettings
End Sub
