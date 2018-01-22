VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Dim mValorInicial
Dim mContenido As Boolean

Private Sub Worksheet_Change(ByVal Target As Range)
' Genera automáticamente el indice de la tabla
    Dim rCelda As Range
    Dim i, iPos As Long
    Dim sPrefijo, sTabla As String
    sTabla = "Tabla6[Indice]"
    sPrefijo = vbNullString
    iPos = 1
    ' Terminar si el rango no tiene celdas en blanco
    If WorksheetFunction.CountBlank(Range(sTabla)) = 0 Then Exit Sub
    i = 0
    For Each rCelda In Range(sTabla)
        If Val(Mid(rCelda.Offset(0, 0), iPos)) > i Then _
          i = Val(Mid(rCelda.Offset(0, 0), iPos))
        If rCelda.Offset(0, 0) = vbNullString Then
            i = i + 1
            rCelda.Offset(0, 0) = sPrefijo & i
        End If
    Next
'Private Sub Worksheet_Change(ByVal Target As Range)
' Compara el valor de la celda, y si éste ha cambiado, _
  entonces el color de la celda cambia
    'Module1.jpUserApplication
    If mContenido = True Then Exit Sub
    On Error GoTo MuchasCeldas
    If Target <> mValorInicial Then
        Target.Font.ColorIndex = 5
    End If
MuchasCeldas:
End Sub


'Private Sub Worksheet_Change(ByVal Target As Range)
' Compara el valor de la celda, y si éste ha cambiado, _
  entonces el color de la celda cambia
    'Module1.jpUserApplication
'    If mContenido = True Then Exit Sub
'    On Error GoTo MuchasCeldas
'    If Target <> mValorInicial Then
'        Target.Font.ColorIndex = 5
'    End If
'MuchasCeldas:
'End Sub


Private Sub aWorksheet_SelectionChange(ByVal Target As Range)
' Al posicionarse en la celda, se guarda el valor de la celda
    'Module1.jpUserApplication
    mContenido = False
    If IsEmpty(Target) Then
            mContenido = True
            Exit Sub
        Else
            On Error GoTo MuchasCeldas
            mValorInicial = Target
    End If
MuchasCeldas:
End Sub






