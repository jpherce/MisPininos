VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usrEstadísticas 
   Caption         =   "Control de Establos"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8040
   OleObjectBlob   =   "usrEstadísticas.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "usrEstadísticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Ultima Modificacion: 29-Oct-2017
'
Option Explicit
Dim bCambios, bProteccion As Boolean
Dim ren, iMes As Long
Dim mArrayA(10, 13) 'ren, col
Dim mArrayB(12, 13) 'ren, col
Dim dDxPct As Double
Dim sN, sD As Long

Private Sub cmndVersionImpresora_Click()
    CommandButton1_Click
    Worksheets("Estadísticas").Visible = True
    Worksheets("Estadísticas").Select
    Application.Run "CalcularEstadisticas"
    Range("B1").Select
End Sub

Private Sub CommandButton1_Click()
' Cerrar formulario
    'Application.Run "MostrarHojas" 'ModSeguridad
    Application.ScreenUpdating = True
    Unload Me
End Sub

Private Sub PoblarLstBoxDesechos()
    Dim rCelda As Range
    Dim R, c As Long
    ' Poblar Encabezados
    mArrayB(0, 0) = "CAUSAS"
    mArrayB(0, 1) = "Ene"
    mArrayB(0, 2) = "Feb"
    mArrayB(0, 3) = "Mar"
    mArrayB(0, 4) = "Abr"
    mArrayB(0, 5) = "May"
    mArrayB(0, 6) = "Jun"
    mArrayB(0, 7) = "Jul"
    mArrayB(0, 8) = "Ago"
    mArrayB(0, 9) = "Sep"
    mArrayB(0, 10) = "Oct"
    mArrayB(0, 11) = "Nov"
    mArrayB(0, 12) = "Dic"
    mArrayB(0, 13) = "TOT."
    mArrayB(1, 0) = vbNullString
    mArrayB(2, 0) = "Producción"
    mArrayB(3, 0) = "Reproducción"
    mArrayB(4, 0) = "Ubres"
    mArrayB(5, 0) = "Locomoción"
    mArrayB(6, 0) = "Lesiones"
    mArrayB(7, 0) = "Neumonía"
    mArrayB(8, 0) = "Diarrea"
    mArrayB(9, 0) = "Otras Causas"
    mArrayB(10, 0) = "Totales"
    mArrayB(11, 0) = vbNullString
    mArrayB(12, 0) = "Machos"
    'PoblarMatrizB
    For Each rCelda In Range("Tabla6[Arete]")
        If CDate(rCelda.Offset(0, 1)) > Date - 365 Then
            If rCelda.Offset(0, 2) = "Baja" Or _
              rCelda.Offset(0, 2) = "Parto" Then
                iMes = Month(CDate(rCelda.Offset(0, 1)))
                Select Case rCelda.Offset(0, 3)
                    Case Is = "Producción"
                        ren = 2
                        PoblarMatrizDesechos
                    Case Is = "Reproducción"
                        ren = 3
                        PoblarMatrizDesechos
                    Case Is = "Mastitis"
                        ren = 4
                        PoblarMatrizDesechos
                    Case Is = "Gabarro"
                        ren = 5
                        PoblarMatrizDesechos
                    Case Is = "Lesiones"
                        ren = 6
                        PoblarMatrizDesechos
                    Case Is = "Neumonía"
                        ren = 7
                        PoblarMatrizDesechos
                    Case Is = "Diarrea"
                        ren = 8
                        PoblarMatrizDesechos
                    Case Is = "Otra"
                        ren = 9
                        PoblarMatrizDesechos
                    Case Is = "M"
                        ren = 12
                        PoblarMatrizDesechos
                End Select
            End If
        End If
    Next
    ' Totalizar
    'For r = 2 To 12
    '    For c = 1 To 11
    '        If mArrayB(c, r) > 0 Then _
    '          mArrayB(10, r) = mArrayB(10, r) + mArrayB(c, r)
    '    Next c
    'Next r
    ' Problar ListBox
    With Me.LstBoxDesechos
        .Clear
        .ColumnWidths = "66;21;21;21;21;22;21;21;21;21;21;21;21;30"
        .List = mArrayB()
    End With
End Sub

Private Sub PoblarLstBoxMorbilidad()
    Dim rCelda As Range
    ' Poblar Encabezados
    mArrayA(0, 0) = "CAUSAS"
    mArrayA(0, 1) = "Ene"
    mArrayA(0, 2) = "Feb"
    mArrayA(0, 3) = "Mar"
    mArrayA(0, 4) = "Abr"
    mArrayA(0, 5) = "May"
    mArrayA(0, 6) = "Jun"
    mArrayA(0, 7) = "Jul"
    mArrayA(0, 8) = "Ago"
    mArrayA(0, 9) = "Sep"
    mArrayA(0, 10) = "Oct"
    mArrayA(0, 11) = "Nov"
    mArrayA(0, 12) = "Dic"
    mArrayA(0, 13) = "TOT."
    mArrayA(1, 0) = vbNullString
    mArrayA(2, 0) = "Ubres"
    mArrayA(3, 0) = "Ret.Placentarias"
    mArrayA(4, 0) = "Metritis"
    mArrayA(5, 0) = "Despl.Abomazo"
    mArrayA(6, 0) = "Locomoción"
    mArrayA(7, 0) = "Neumonía"
    mArrayA(8, 0) = "Diarrea"
    mArrayA(9, 0) = "Lesiones"
    mArrayA(10, 0) = "Otras Causas"
    'PoblarMatrizA
    For Each rCelda In Range("Tabla6[Arete]")
        If CDate(rCelda.Offset(0, 1)) > Date - 365 Then
            If InStr(rCelda.Offset(0, 2), "Enf") Then
                iMes = Month(CDate(rCelda.Offset(0, 1)))
                Select Case rCelda.Offset(0, 2)
                    Case Is = "Enf-MA"
                        ren = 2
                        PoblarMatrizMorbilidad
                    Case Is = "Enf-RP"
                        ren = 3
                        PoblarMatrizMorbilidad
                    Case Is = "Enf-UM"
                        ren = 4
                        PoblarMatrizMorbilidad
                    Case Is = "Enf-DA"
                        ren = 5
                        PoblarMatrizMorbilidad
                    Case Is = "Enf-Ga"
                        ren = 6
                        PoblarMatrizMorbilidad
                    Case Is = "Enf-NE"
                        ren = 7
                        PoblarMatrizMorbilidad
                    Case Is = "Enf-Di"
                        ren = 8
                        PoblarMatrizMorbilidad
                    Case Is = "Enf-He"
                        ren = 9
                        PoblarMatrizMorbilidad
                    Case Is = "Enf-Ot"
                        ren = 10
                        PoblarMatrizMorbilidad
                End Select
            End If
        End If
    Next
    ' Problar ListBox
    With Me.LstBoxMorbilidad
        .Clear
        .ColumnWidths = "66;21;21;21;21;22;21;21;21;21;21;21;21;30"
        .List = mArrayA()
    End With
End Sub

Private Sub PoblarProgramacion()
    With Me
        .lblMes1 = Format(Month(Date), "0#") & "-" & Format(Right(Year(Date), 2), "##")
        .lblAxPM1 = AnimPorParir2(1)
        .lblApS1 = AnimPorSecar(1)
        .lblMes2 = Format(Month(Date + 30), "0#") & "-" & _
          Format(Right(Year(Date + 30), 2), "##")
        .lblAxPM2 = AnimPorParir2(2)
        .lblApS2 = AnimPorSecar(2)
        .lblMes3 = Format(Month(Date + 60), "0#") & "-" & _
          Format(Right(Year(Date + 60), 2), "##")
        .lblAxPM3 = AnimPorParir2(3)
        .lblApS3 = AnimPorSecar(3)
        .lblMes4 = Format(Month(Date + 90), "0#") & "-" & _
          Format(Right(Year(Date + 90), 2), "##")
        .lblAxPM4 = AnimPorParir2(4)
        .lblApS4 = AnimPorSecar(4)
        .lblMes5 = Format(Month(Date + 120), "0#") & "-" & _
          Format(Right(Year(Date + 120), 2), "##")
        .lblAxPM5 = AnimPorParir2(5)
        .lblApS5 = AnimPorSecar(5)
        .lblMes6 = Format(Month(Date + 150), "0#") & "-" & _
          Format(Right(Year(Date + 150), 2), "##")
        .lblAxPM6 = AnimPorParir2(6)
        .lblApS6 = AnimPorSecar(6)
        .LblMes7 = Format(Month(Date + 180), "0#") & "-" & _
          Format(Right(Year(Date + 180), 2), "##")
        .lblAxPM7 = AnimPorParir2(7)
        .lblApS7 = AnimPorSecar(7)
        .lblMes8 = Format(Month(Date + 210), "0#") & "-" & _
          Format(Right(Year(Date + 210), 2), "##")
        .lblAxPM8 = AnimPorParir2(8)
        .lblApS8 = AnimPorSecar(8)
    End With
End Sub

Private Sub UserForm_Initialize()
    Dim iTotalVacas, iVacasProd, iVacasSeca As Long
    Dim iTotReemplazos, iLact, iDes, iNov, iVaq, iVaqGest As Long
    Application.ScreenUpdating = CBool(Range("Desarrollador!B6"))
    sD = 365
    ' Estadísticas del Hato
    With Me
        ' Tomar Valores
        .Caption = Range("Configuracion!C3")
        .Label83 = "Ver. 15.1004"
        On Error Resume Next
        .txtPicoProd = Format(WorksheetFunction. _
          Max(Range("Tabla15[30d]:Tabla15[300d]")), "#.0")
        .txtProdDiaria = Format(WorksheetFunction. _
          Sum(Range("Tabla1[Prod.]")), "#,#")
        .txtPromDEL = Format(pDEL(), "#")
        If .txtProdDiaria > Range("Configuracion!B73") Then _
          .txtPromDEL.ForeColor = RGB(255, 0, 0)
        .txtPromD1S = Format(pD1S(), "#")
        If .txtPromD1S > Range("Configuracion!B76") Then _
          .txtPromD1S.ForeColor = RGB(255, 0, 0)
        .txtD1SG = Format(pD1S(, "P"), "#")
        If .txtD1SG > Range("Configuracion!B77") Then _
          .txtD1SG.ForeColor = RGB(255, 0, 0)
        Select Case Val(.txtPromD1S)
            Case Is > Val(.txtD1SG)
                .Label139 = "k" '"è"
                .Label139.ForeColor = RGB(255, 0, 0)
            Case Is = Val(.txtD1SG)
                .Label139 = "g" '"ì"
            Case Is < Val(.txtD1SG)
                .Label139 = "m" '"î"
        End Select
        .txtPromDAb = Format(pDAb(), "#")
        If .txtPromDAb > Range("Configuracion!C75") Then _
          .txtPromDAb.ForeColor = RGB(255, 0, 0)
        .txtPromProd = Format(pProdLinea, "#.0")
        .txtProy305d = Format(pProy305d, "#,#")
        '.txtServConcep = Format(pServConcep(), "#.0")
        .txtPersist = Int(WorksheetFunction.Average _
          (Range("Tabla15[Persistencia]"))) & "%"
        If .txtPersist < Range("Configuracion!B74") Then _
          .txtPersist.ForeColor = RGB(255, 0, 0)
        .txtServConcep = Format(pServicios(1, "P"), "#.0")
        If .txtServConcep > Range("Configuracion!B79") Then _
          .txtServConcep.ForeColor = RGB(255, 0, 0)
        '.txtServVaca = Format(pServVaca(), "#.0")
        .txtServVaca = Format(pServicios(1), "#.0")
        If .txtServVaca > Range("Configuracion!B78") Then _
          .txtServVaca.ForeColor = RGB(255, 0, 0)
        Select Case Val(.txtServVaca)
            Case Is > Val(.txtServConcep)
                .Label140 = "k" '"è"
                .Label140.ForeColor = RGB(255, 0, 0)
            Case Is = Val(.txtServConcep)
                .Label140 = "g" '"ì"
            Case Is < Val(.txtServConcep)
                .Label140 = "m" '"î"
        End Select
        .txtPromProdHato = Format(WorksheetFunction. _
          Sum(Range("Tabla1[Prod.]")) / tAnimales(1), "#,#.0")
        .txtTotalVacas = Format(tAnimales(1), "#,#")
        .TxtVacasProd = Format(tVacasProd, "#,#")
        .txtVacasSecas = Format(tVacasSecas, "#,#")
        If tAnimales(1) > 0 Then
            .Label33 = Format(tVacasProd / tAnimales(1), "0%")
            .Label34 = Format(tVacasSecas / tAnimales(1), "0%")
        End If
        '.lblIG1 = "% Dx Gest. Positivos"
        '.txtIG1 = Format(DxGstPositivos, "#") & "%"
        .lblIG1 = "Tasa de embarazo"
        .lblIG2 = "% Calores detectados"
        .lblIG3 = "Intervalo entre servicios"
        .lblIG4 = "% Gest. 1° Serv."
        '.lblIG4 = "Prom. de calores perdidos"
        .lblIG5 = "% Dx gest. positivos último mes"
        .lblIG6 = "Abortos último año"
        .txtIG1 = Format(TasaEmbarazo(), "0%")
        If .txtIG1 < Range("Configuracion!B82") Then _
          .txtIG1.ForeColor = RGB(255, 0, 0)
        .txtIG2 = Format(HeatsDetected, "0%")
        If .txtIG2 < Range("Configuracion!C83") Then _
          .txtIG2.ForeColor = RGB(255, 0, 0)
        .txtIG3 = Format(BreedingInterval, "0") & " d"
        If BreedingInterval > Range("Configuracion!C84") Then _
          .txtIG3.ForeColor = RGB(255, 0, 0)
        .txtIG4 = Format(pctGest1Serv, "0%")
        If .txtIG4 < Range("Configuracion!C85") Then _
          .txtIG4.ForeColor = RGB(255, 0, 0)
        '.txtIG4 = Format(nCaloresPerdidos, "#.0")
        .txtIG5 = Format(DxGstPositivos(30), "#") & "%"
        If Val(.txtIG5) < Range("Configuracion!C87") Then _
          .txtIG5.ForeColor = RGB(255, 0, 0)
        .txtIG6 = numAbortos
        .txtIG61 = Format(pctAbortos, "0%")
        If Val(.txtIG61) > Range("Configuracion!B88") Then _
          .txtIG61.ForeColor = RGB(255, 0, 0)
        .Label131 = Format(WorksheetFunction. _
          CountIf(Range("Tabla1[Status]"), "=P"), "#,0")
        .Label132 = Format(WorksheetFunction. _
          CountIf(Range("Tabla1[Servicio]"), ">0"), "#,0") _
          - WorksheetFunction. _
          CountIf(Range("Tabla1[Status]"), "=P")
        .Label133 = tProblema(1)
        .Label134 = tRepetidoras(1)
        .Label135 = Format((Val(Me.Label131) / tAnimales(1)), "0%")
        .Label136 = Format((Val(Me.Label132) / tAnimales(1)), "0%")
        .Label137 = Format((Val(Me.Label133) / tAnimales(1)), "0%")
        If Val(.Label137) > Range("Configuracion!B80") Then _
          .Label137.ForeColor = RGB(255, 0, 0)
        .Label138 = Format((Val(Me.Label134) / tAnimales(1)), "0%")
        If Val(.Label138) > Range("Configuracion!B81") Then _
          .Label138.ForeColor = RGB(255, 0, 0)
        .Label154 = Format(pD1Calor(1), "#d")
        If Val(.Label154) > Range("Configuracion!B89") Then _
          .Label154.ForeColor = RGB(255, 0, 0)
        .Label155 = Format(pD1Calor(2), "#d")
        If Val(.Label155) > Range("Configuracion!B90") Then _
          .Label155.ForeColor = RGB(255, 0, 0)
        .Label159 = cvProy305d & "%"
        On Error GoTo 0
    End With
    ' Estadísticas de Reemplazos
    With Me
        On Error Resume Next        ' Tomar Valores
        .Label64 = tLactantes
        .Label65 = tDesarrollo
        .Label69 = tNovillas
        .Label73 = tVaquillas("O")
        .Label72 = tVaquillas("P")
        .Label75 = tAnimales(2)
        If tAnimales(2) > 0 Then
            .Label66 = Format(tLactantes / tAnimales(2), "0%")
            .Label67 = Format(tDesarrollo / tAnimales(2), "0%")
            .Label76 = Format(tNovillas / tAnimales(2), "0%")
            .Label77 = Format(tVaquillas("O") / tAnimales(2), "0%")
            .Label78 = Format(tVaquillas("P") / tAnimales(2), "0%")
        End If
        sN = pEdadAlParto
        .Label45 = Format(Int(sN / sD), "0") & "-" & _
        Format(Int((sN - sD * Int(sN / sD)) / 30.4), "00")
        sN = pEdad1Serv()
        If Not pEdad1Serv() = "ND" Then _
          .Label46 = Format(Int(sN / sD), "0") & "-" & _
          Format(Int((sN - sD * Int(sN / sD)) / 30.4), "00") _
          Else .Label46 = pEdad1Serv()
        sN = pEdad1Serv("P")
        If Not pEdad1Serv("P") = "ND" Then _
          .txtEdad1SGest = Format(Int(sN / sD), "0") & "-" & _
          Format(Int((sN - sD * Int(sN / sD)) / 30.4), "00") _
          Else .txtEdad1SGest = pEdad1Serv("P")
        Select Case Val(.txtEdad1SGest)
            Case Is > Val(.Label46)
                .Label141 = "m" '"è"
            Case Is = Val(.Label46)
                .Label141 = "g" '"ì"
            Case Is < Val(.Label46)
                .Label141 = "k" '"î"
        End Select
        '.Label47 = Format(pServReemplazos, "#.0")
        .Label47 = Format(pServicios(2), "#.0")
        '.Label48 = Format(pServConcepR, "#.0")
        .Label48 = Format(pServicios(2, "P"), "#.0")
        Select Case Val(.Label48)
            Case Is > Val(.Label47)
                .Label142 = "m" '"è"
            Case Is = Val(.Label47)
                .Label142 = "g" '"ì"
            Case Is < Val(.Label47)
                .Label142 = "k" '"î"
        End Select
        
        .Label144 = tVaquillas("P")
        .Label145 = tVaquillas("O")
        .Label146 = tProblema(2)
        .Label147 = tRepetidoras(2)
        .Label148 = Format(Val(.Label144) / tAnimales(2), "0%")
        .Label149 = Format(Val(.Label145) / tAnimales(2), "0%")
        .Label150 = Format(Val(.Label146) / tAnimales(2), "0%")
        .Label151 = Format(Val(.Label147) / tAnimales(2), "0%")
        '.Label51 = Format(DxGstPositivos, "#") & "%"
    On Error GoTo 0
    End With
    PoblarLstBoxMorbilidad
    PoblarLstBoxDesechos
    PoblarProgramacion
    bCambios = False
End Sub

Private Sub UserForm_QueryClose(mCancel As Integer, _
                                mCloseMode As Integer)
' Deshabilitar el Botón X para cerrar el cuadro de diálogo
    If mCloseMode <> vbFormCode Then
        MsgBox "Utiliza 'Cerrar' para salir de este formulario.", _
          vbExclamation, "JP's Automatización de Aplicaciones"
        mCancel = True
    End If
End Sub

Private Sub PoblarMatrizMorbilidad()
    mArrayA(ren, iMes) = mArrayA(ren, iMes) + 1
    mArrayA(ren, 13) = mArrayA(ren, 13) + 1
End Sub

Private Sub PoblarMatrizDesechos()
    mArrayB(ren, iMes) = mArrayB(ren, iMes) + 1
    mArrayB(ren, 13) = mArrayB(ren, 13) + 1
    If Not ren = 12 Then _
      mArrayB(10, iMes) = mArrayB(10, iMes) + 1
End Sub
