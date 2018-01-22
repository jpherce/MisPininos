Attribute VB_Name = "ModExportarDHI"
'Ultima modificación: 22-Oct-2017
'Módilo para exportación de datos a Holstein de México
Option Explicit
Dim sArch As String
Dim nArch As String
Dim rCelda As Range
Dim sEvento As String

Sub ExportarHolstein()
    Dim c As String
    Dim a As Range
    Dim lTotal, lCounter As Long
    Application.ScreenUpdating = CBool(Range("Desarrollador!B6"))
    Application.DisplayStatusBar = True
    Set a = Range("A1:" & Range("A1").End(xlDown).Address)
    c = "A"
    Application.DisplayStatusBar = True
    lTotal = a.Rows.Count
    For Each rCelda In Range("Tabla6[Indice]")
        If rCelda.Offset(0, 1) = vbNullString Then
            Application.StatusBar = "Exportando... " & _
             Format(lCounter / lTotal, "0%")
            Select Case rCelda.Offset(0, -7)
                Case "Calor"
                    nArch = "CapturaCalor.csv"
                    sEvento = "H"
                    LoginRecord
                    rCelda.Offset(0, 1) = c
                Case "DxGst"
                    nArch = "CapturaDxGestacion.csv"
                    sEvento = "P"
                    LoginRecord
                    rCelda.Offset(0, 1) = c
                Case "Calor"
                    nArch = "CapturaEstadio.csv"
                    sEvento = "H"
                    LoginRecord
                    rCelda.Offset(0, 1) = c
                Case "Parto"
                    nArch = "CapturaParto.csv"
                    sEvento = 2
                    LoginRecord
                    rCelda.Offset(0, 1) = c
                Case "Prod"
                    nArch = "CapturaPesadas.csv"
                    sEvento = "Prod"
                    LoginRecord
                    rCelda.Offset(0, 1) = c
                Case "Rev"
                    'nArch = "CapturaRevisiones.csv"
                    'LoginRecord
                    'rCelda.Offset(0, 1) = c
                Case "Seca"
                    nArch = "CapturaSecados.csv"
                    sEvento = 6
                    LoginRecord
                    rCelda.Offset(0, 1) = c
                Case "Serv"
                    nArch = "CapturaEstadios.csv"
                    sEvento = "B"
                    LoginRecord
                    rCelda.Offset(0, 1) = c
            End Select
        End If
        lCounter = lCounter + 1
    Next
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub

Private Sub LoginRecord()
' Bitácora de Eventos en formato CSV
'    Dim sPath As String
    'Dim sArch As String
    Dim bArch As Boolean
'    sPath = Application.ActiveWorkbook.Path
'    sArch = Dir("C:\Users\yo\Documents\My Box Files\INIFAP\Log101.txt")
    sArch = Dir(nArch)
    If sArch <> vbNullString Then bArch = True
    On Error GoTo 100
'    Open "C:\Users\yo\Documents\My Box Files\INIFAP\Log101.txt" For Append As #1
    Open nArch For Append As #1
    On Error GoTo 0
200
    If bArch = False Then Write #1, "IdHato", "Arete", "Fecha", _
      "Evento", "Observaciones", "Responsable"
    Write #1, Range("Configuracion!D3"), rCelda.Offset(0, -9), rCelda.Offset(0, -8), _
      sEvento, rCelda.Offset(0, -6), rCelda.Offset(0, -5)
    Close #1
    Exit Sub
100
    Open nArch For Append As #1
    GoTo 200
End Sub


