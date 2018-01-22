Attribute VB_Name = "ModFertilidad"
Option Explicit
' Ultima modificación: 21.11.17
Dim iR As Long
Dim ws As Worksheet
Dim rCelda As Range
Dim sTabla As String

Private Sub AnalisisFertilidad()
    Application.ScreenUpdating = CBool(Range("Desarrollador!B6"))
    Application.DisplayStatusBar = True
    With Worksheets("InventarioSemen")
        .Visible = xlSheetVisible
        .Activate
        .Select
        Application.Run "Desproteger" 'Mod2
        .Range("A2").Select
    End With
    BT3
    FertilidadToros
    Range("A2").Select
    Application.Run "Proteger" 'Mod2
    Application.ScreenUpdating = True
End Sub

Private Sub BT3()
Dim oldStatusBar As String
    oldStatusBar = Application.DisplayStatusBar
    'Range("Tabla3").Clear
    Application.StatusBar = "Borrando Tabla3..."
    Set ws = Worksheets("InventarioSemen")
    sTabla = "Tabla3"
    BT
    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar
End Sub

Private Sub BT()
Dim iR As Long
Dim i As Long
    'ws.Visible = xlSheetVisible
    'ws.Select
    'Application.Run "Desproteger" 'Mód2
    If ws.Name = "Hato" Or ws.Name = "Reemplazos" Then _
      Range("Desarrollador!B20") = "T" 'Range("XFD1") = "T"
    Range(sTabla).Select
    If WorksheetFunction.Count(Range(sTabla)) = 0 Then _
     iR = 0 Else iR = WorksheetFunction.Count(Range(sTabla)) '- 1
    On Error Resume Next
    For i = 1 To iR
        Selection.ListObject.ListRows(1).Delete
    Next i
    On Error GoTo 0
End Sub

Private Sub FertilidadToros()
' Evaluación de todos los toros en la Base de Datos
Dim rCelda1 As Range
Dim lTot, lC As Long
    'Application.ScreenUpdating = CBool(Range("Desarrollador!B6"))
    'Application.DisplayStatusBar = True
    Set ws = Worksheets("InventarioSemen")
    'ws.Select
    'Desproteger
    lTot = Range("Tabla6").Rows.Count
    'Range("tabla3[[Calc]]").Clear
    For Each rCelda In Range("Tabla6[Arete]")
        Application.StatusBar = _
          "Obteniendo toros utilizados " & _
          Format(lC / lTot, "0%")
        If rCelda.Offset(0, 2) = "Serv" Then
        ' Localizar Toro
            For Each rCelda1 In Range("Tabla3[Toro]")
                If rCelda1 = rCelda.Offset(0, 3) Then
                    'iR = rCelda1.Offset.Row
                    GoTo 2123
                End If
            Next rCelda1
            'Agregar Toro
            iR = TamañoTabla("Tabla3") + 2
            ws.Cells(iR, 1) = rCelda.Offset(0, 3)
                'Escribir fórmulas
                'CalcularFertilidad
2123:
        End If
        lC = lC + 1
    Next
    lC = 0
    lTot = Range("Tabla3").Rows.Count
    For Each rCelda In Range("Tabla3[Toro]")
        Application.StatusBar = _
          "Calculando fertilidad por toro " & _
          Format(lC / lTot, "0%")
        iR = rCelda.Offset.Row
        CalcularFertilidad
        lC = lC + 1
    Next
    Application.StatusBar = False
    'Application.ScreenUpdating = True
End Sub

Private Sub CalcularFertilidad()
    On Error Resume Next
    With ws
        '% de Fertilidad
        .Cells(iR, 2) = FertilidadToro(rCelda.Offset(0, 0))
        ' % de Fertilidad al 1° Servicio
        .Cells(iR, 3) = WorksheetFunction.IfError( _
        WorksheetFunction.CountIfs _
          (Range("Tabla6[Observaciones]"), _
          rCelda.Offset(0, 0), _
          Range("Tabla6[Fecha]"), "<=" & Format(Date - _
          Range("Configuracion!C5"), "dd-mmm-yyyy"), _
          Range("Tabla6[Metadatos]"), "01-*", _
          Range("Tabla6[Metadatos]"), "*-P") / _
          WorksheetFunction.CountIfs _
          (Range("Tabla6[Observaciones]"), _
          rCelda.Offset(0, 0), _
          Range("Tabla6[Fecha]"), "<=" & Format(Date - _
          Range("Configuracion!C5"), "dd-mmm-yyyy"), _
          Range("Tabla6[Metadatos]"), "01-*"), "ND")
        ' Servicios realizados
        .Cells(iR, 4) = WorksheetFunction.CountIfs _
          (Range("Tabla6[Observaciones]"), _
          rCelda.Offset(0, 0))
        ' Servicios después del Dx Gestación
        .Cells(iR, 5) = WorksheetFunction.CountIfs _
          (Range("Tabla6[Observaciones]"), _
          rCelda.Offset(0, 0), _
          Range("Tabla6[Fecha]"), "<" & Format(Date - _
          Range("Configuracion!C5"), "dd-mmm-yyyy"))
        ' Gestantes
        .Cells(iR, 6) = WorksheetFunction.CountIfs _
          (Range("Tabla6[Observaciones]"), _
          rCelda.Offset(0, 0), _
          Range("Tabla6[Metadatos]"), "*-P") + _
          WorksheetFunction.CountIfs _
          (Range("Tabla6[Observaciones]"), _
          rCelda.Offset(0, 0), _
          Range("Tabla6[Metadatos]"), "*-R")
        ' Vacías
        .Cells(iR, 7) = WorksheetFunction.CountIfs _
          (Range("Tabla6[Observaciones]"), _
          rCelda.Offset(0, 0), _
          Range("Tabla6[Fecha]"), "<=" & Format(Date - _
          Range("Configuracion!C5"), "dd-mmm-yyyy"), _
          Range("Tabla6[Metadatos]"), "<>" & "*-P", _
          Range("Tabla6[Metadatos]"), "<>" & "*-R")
        ' % de reabsorciones
        .Cells(iR, 8) = WorksheetFunction.IfError _
          (WorksheetFunction.CountIfs _
          (Range("Tabla6[Observaciones]"), _
          rCelda.Offset(0, 0), _
          Range("Tabla6[Metadatos]"), "*-R") / _
          (WorksheetFunction.CountIfs _
          (Range("Tabla6[Observaciones]"), _
          rCelda.Offset(0, 0), _
          Range("Tabla6[Metadatos]"), "*-P") + _
          WorksheetFunction.CountIfs _
          (Range("Tabla6[Observaciones]"), _
          rCelda.Offset(0, 0), _
          Range("Tabla6[Metadatos]"), "*-R")), 0)
        ' Última dosis utilizada
        .Cells(iR, 9) = WorksheetFunction.MaxIfs _
          (Range("Tabla6[Fecha]"), _
          Range("Tabla6[Observaciones]"), _
          rCelda.Offset(0, 0))
        ' Pendientes de Dx de gestación
        .Cells(iR, 10) = WorksheetFunction.CountIfs _
          (Range("Tabla6[Observaciones]"), _
          rCelda.Offset(0, 0), _
          Range("Tabla6[Fecha]"), ">=" & Format(Date - _
          Range("Configuracion!C5"), "dd-mmm-yyyy"))
        ' Num. de hijas en el hato
        .Cells(iR, 11) = WorksheetFunction.CountIf _
          (Range("Tabla8[Padre]"), _
          rCelda.Offset(0, 0))
        ' % Eq. Madurez de las hijas
        .Cells(iR, 12) = ""
    End With
    On Error GoTo 0
End Sub
