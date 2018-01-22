Attribute VB_Name = "ModTareas"
'Última modificación: 24-Ene-2016
Option Explicit
Dim a, i, iAR As Long
Dim mContenido(1 To 15)
Dim ws As Worksheet

Private Sub TareasSemanales()
    'Application.ScreenUpdating = CBool(Range("Desarrollador!B6"))
    'Dim a
    'Mostrar hoja
    Application.StatusBar = "Prepararando Informe ..."
    With Sheets("TareasSemanal")
        .Visible = True
        .Select
        .Cells.ClearContents
    End With
    For a = 1 To 100: Next a
    Application.StatusBar = "Configurando impresora ..."
    ConfigPagImpresa
    'Escribir Nueva información
    'Revisiones
    Application.StatusBar = "Recabando Vacas por revisar ..."
    i = 0
    Range("A1").Select
    Encabezado ("VACAS A REVISION")
    Encabezado1
    For a = 1 To 10
    LlenarContenido
    Next a
    'Dx de Gestación
    Application.StatusBar = "Recabando Diagnósticos de Gestación ..."
    Encabezado ("ANIMALES A DIAGNOSTICO DE GESTACIÓN")
    Encabezado1
    For a = 1 To 10
    LlenarContenido
    Next a
    'Vacas por Secar
    Application.StatusBar = "Recabando Vacas por Secar ..."
    Encabezado ("VACAS POR SECAR")
    Encabezado1
    For a = 1 To 10
    LlenarContenido
    Next a
    'Vacas por Servir
    Application.StatusBar = "Recabando Vacas por Servir ..."
    Encabezado ("ANIMALES POR SERVIR")
    Encabezado1
    For a = 1 To 10
    LlenarContenido
    Next a
    'Animales por Parir
    Application.StatusBar = "Recabando Animales por Parir ..."
    Encabezado ("ANIMALES POR PARIR")
    Encabezado1
    For a = 1 To 10
    LlenarContenido
    Next a
    'Animales por Destetar
    Application.StatusBar = "Recabando Animales por Destetar ..."
    Encabezado ("ANIMALES POR DESTETAR")
    Encabezado1
    For a = 1 To 10
    LlenarContenido
    Next a
    'Animales por Vacunar
    Application.StatusBar = "Recabando Animales por Vacunar ..."
    Encabezado ("ANIMALES POR VACUNAR")
    Encabezado1
    For a = 1 To 10
    LlenarContenido
    Next a
    'Animales por Imantar
    Application.StatusBar = "Recabando Animales por Imanatar ..."
    Encabezado ("ANIMALES POR IMANTAR")
    Encabezado1
    For a = 1 To 10
    LlenarContenido
    Next a
    'Vacas Atrasadas
    Application.StatusBar = "Recabando Animales Retrasados ..."
    Encabezado ("ANIMALES ATRASADOS")
    Encabezado1
    For a = 1 To 10
    LlenarContenido
    Next a
    ActiveCell.Offset(i + 1, 0) = "*EOF()*"
    ' Configurar Columnas
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub

Private Function Encabezado(mensaje)
    i = i + 1
    With ActiveCell
        .Offset(i, 0).Font.Size = 12
        .Offset(i, 0).Font.Bold = True
        .Offset(i, 0) = mensaje
    End With
    RenglonAbajo
End Function

Private Function Encabezado1()
    With ActiveCell
        .Offset(i, 0) = "Arete"
        .Offset(i, 1) = "Corral"
        .Offset(i, 2) = "Prod"
        .Offset(i, 3) = "DEL"
        .Offset(i, 4) = "Parto"
        .Offset(i, 5) = "F.Parto"
        .Offset(i, 6) = "Serv."
        .Offset(i, 7) = "F.Servicio"
        .Offset(i, 8) = "Semental"
        .Offset(i, 9) = "Técnico"
        .Offset(i, 10) = "FxSecar"
        .Offset(i, 11) = "FxParir"
        .Offset(i, 12) = "Observaciones"
    End With
    RenglonAbajo
End Function

Private Sub TomarDatos()
    Set ws = Worksheets("Hato")
    Dim Renglon As Long
    Renglon = 2
    
    With ws
        mContenido(1) = .Cells(Renglon, 1) 'Arete
        mContenido(2) = .Cells(Renglon, 2) 'Corral
        mContenido(3) = .Cells(Renglon, 3) 'Prod
        mContenido(4) = .Cells(Renglon, 4) 'DEL
        mContenido(5) = .Cells(Renglon, 5) 'Parto
        mContenido(6) = .Cells(Renglon, 6) 'F.Parto
        mContenido(7) = .Cells(Renglon, 7) 'Serv
        mContenido(8) = .Cells(Renglon, 8) 'F.Serv
        mContenido(9) = .Cells(Renglon, 9) 'Semental
        mContenido(10) = .Cells(Renglon, 10) 'Técnico
        mContenido(11) = .Cells(Renglon, 11) 'FxSecar
        mContenido(12) = .Cells(Renglon, 12) 'FxParir
    End With
End Sub

Private Sub LlenarContenido()
    TomarDatos
    With ActiveCell
        .Offset(i, 0) = mContenido(1) 'Arete
        .Offset(i, 1) = mContenido(2) 'Corral
        .Offset(i, 2) = mContenido(3) 'Prod
        .Offset(i, 3) = mContenido(4) 'DEL
        .Offset(i, 4) = mContenido(5) 'Parto
        .Offset(i, 5) = mContenido(6) 'F.Parto
        .Offset(i, 6) = mContenido(7) 'Serv
        .Offset(i, 7) = mContenido(8) 'F.Serv
        .Offset(i, 8) = mContenido(9) 'Semental
        .Offset(i, 9) = mContenido(10) 'Técnico
        .Offset(i, 10) = mContenido(11) 'FxSecar
        .Offset(i, 11) = mContenido(12) 'FxParir
        .Offset(i, 12) = "_______________________________"
    End With
    RenglonAbajo
End Sub

Private Function RenglonAbajo()
    i = i + 1
End Function

Private Sub ConfigPagImpresa()
    On Error Resume Next
    'Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        '.PrintTitleRows = "$1:$2"
        '.PrintTitleColumns = vbnullstring
        '.PrintArea = vbnullstring
        .LeftHeader = _
          Range("Configuracion!C3")
        .CenterHeader = vbNullString
        .RightHeader = _
          "Tareas por realizar: " _
          & Format(Date, "dd-mmm-yy")
        .LeftFooter = _
          "Control de Establos"
        .CenterFooter = vbNullString
        .RightFooter = _
          "Página &P de &N"
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1)
        .HeaderMargin = Application.CentimetersToPoints(1)
        .FooterMargin = Application.CentimetersToPoints(1)
        '.PrintHeadings = True
        .PrintGridlines = False
        '.PrintComments = xlPrintNoComments
        .PrintQuality = 300
        '.CenterHorizontally = True
        '.CenterVertically = False
        '.Orientation = xlLandscape
        .Draft = True
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = True
        .Zoom = 100
        '.PrintErrors = xlPrintErrorsDisplayed
        '.OddAndEvenPagesHeaderFooter = False
        '.DifferentFirstPageHeaderFooter = False
        '.ScaleWithDocHeaderFooter = True
        '.AlignMarginsHeaderFooter = True
        '.EvenPage.LeftHeader.Text = vbnullstring
        '.EvenPage.CenterHeader.Text = vbnullstring
        '.EvenPage.RightHeader.Text = vbnullstring
        '.EvenPage.LeftFooter.Text = vbnullstring
        '.EvenPage.CenterFooter.Text = vbnullstring
        '.EvenPage.RightFooter.Text = vbnullstring
        '.FirstPage.LeftHeader.Text = vbnullstring
        '.FirstPage.CenterHeader.Text = vbnullstring
        '.FirstPage.RightHeader.Text = vbnullstring
        '.FirstPage.LeftFooter.Text = vbnullstring
        '.FirstPage.CenterFooter.Text = vbnullstring
        '.FirstPage.RightFooter.Text = vbnullstring
    End With
    'Application.PrintCommunication = True
    On Error GoTo 0
End Sub

