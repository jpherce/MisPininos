VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Ultima modificación:
Option Explicit

Private Sub Worksheet_Activate()
    On Error GoTo CE
    If Not Range("Desarrollador!B21") = Date Then _
      Application.Run "ADEL" 'Mod2
    'Application.Run "FCH" 'Mod2
    On Error GoTo 0
CE:
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo CE
    Application.Run "FCH" 'Mod2
    On Error GoTo 0
CE:
End Sub
