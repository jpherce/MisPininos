VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Ultima modificación: 2-Oct-2015
Option Explicit

Private Sub Worksheet_Activate()
    On Error GoTo ControlErrores
    'Application.Run "AE" 'Mod2
    If Not Range("Desarrollador!B21") = Date Then _
      Application.Run "FCR" 'Mód2
    On Error GoTo 0
ControlErrores:
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo CE
    Application.Run "FCR" 'Mod2
    On Error GoTo 0
CE:
End Sub
