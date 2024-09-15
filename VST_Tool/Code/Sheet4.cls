VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub CommandButton1_Click()
    Dim ColorValue As Double
    ColorValue = PickNewColor
    ThisWorkbook.Worksheets("Color Picker").Range("C4").Value = ColorValue
End Sub

'Picks new color
Function PickNewColor(Optional i_OldColor As Double = xlNone) As Double
Const BGColor As Long = 13160660  'background color of dialogue
Const ColorIndexLast As Long = 32 'index of last custom color in palette

Dim myOrgColor As Double          'original color of color index 32
Dim myRGB_R As Long            'RGB values of the color that will be
Dim myRGB_G As Long            'displayed in the dialogue as
Dim myRGB_B As Long            '"Current" color
  
  'save original palette color, because we don't really want to change it
  myOrgColor = ActiveWorkbook.Colors(ColorIndexLast)
  
  If i_OldColor = xlNone Then
    'get RGB values of background color, so the "Current" color looks empty
    Color2RGB BGColor, myRGB_R, myRGB_G, myRGB_B
  Else
    'get RGB values of i_OldColor
    Color2RGB i_OldColor, myRGB_R, myRGB_G, myRGB_B
  End If
  
  'call the color picker dialogue
  If Application.Dialogs(xlDialogEditColor).Show(ColorIndexLast, myRGB_R, myRGB_G, myRGB_B) = True Then
    '"OK" was pressed, so Excel automatically changed the palette
    'read the new color from the palette
    PickNewColor = ActiveWorkbook.Colors(ColorIndexLast)
    'reset palette color to its original value
    ActiveWorkbook.Colors(ColorIndexLast) = myOrgColor
  Else
    '"Cancel" was pressed, palette wasn't changed
    'return old color (or xlNone if no color was passed to the function)
    PickNewColor = i_OldColor
  End If
End Function

'Converts a color to RGB values
Sub Color2RGB(ByVal i_Color As Long, o_R As Long, o_G As Long, o_B As Long)
  o_R = i_Color Mod 256
  i_Color = i_Color \ 256
  o_G = i_Color Mod 256
  i_Color = i_Color \ 256
  o_B = i_Color Mod 256
End Sub
