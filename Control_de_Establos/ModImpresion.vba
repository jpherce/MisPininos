Attribute VB_Name = "ModImpresion"
Option Private Module
Option Explicit


Sub ConfigHojaImpresion()

    On Error Resume Next
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        '.PrintTitleRows = "$1:$2"
        '.PrintTitleColumns = vbnullstring
        '.PrintArea = vbnullstring
        If .LeftHeader = vbNullString Then .LeftHeader = _
          Range("Configuracion!C3")
        '.CenterHeader = vbnullstring
        If .RightHeader = vbNullString Then .RightHeader = _
          "Situación del Establo al día: " _
          & Format(Date, "dd-mmm-yy")
        If .LeftFooter = vbNullString Then .LeftFooter = _
          "Control de Establos"
        '.CenterFooter = vbnullstring
        If .RightFooter = vbNullString Then .RightFooter = _
          "Página &P de &N"
        '.LeftMargin = Application.InchesToPoints(0.31496062992126)
        '.RightMargin = Application.InchesToPoints(0.31496062992126)
        '.TopMargin = Application.InchesToPoints(0.748031496062992)
        '.BottomMargin = Application.InchesToPoints(0.748031496062992)
        '.HeaderMargin = Application.InchesToPoints(0.31496062992126)
        '.FooterMargin = Application.InchesToPoints(0.31496062992126)
        '.PrintHeadings = False
        'If .PrintGridlines <> False Then .PrintGridlines = False
        '.PrintComments = xlPrintNoComments
        If .PrintQuality <> 600 Then .PrintQuality = 600
        If .CenterHorizontally <> True Then .CenterHorizontally = True
        If .CenterVertically <> False Then .CenterVertically = False
        If .Orientation <> xlLandscape Then .Orientation = xlLandscape
        If .Draft <> True Then .Draft = True
        If .PaperSize <> xlPaperLetter Then .PaperSize = xlPaperLetter
        If .FirstPageNumber <> xlAutomatic Then _
          .FirstPageNumber = xlAutomatic
        If .Order <> xlDownThenOver Then .Order = xlDownThenOver
        If .BlackAndWhite <> True Then .BlackAndWhite = True
        If .Zoom <> 100 Then .Zoom = 100
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
    Application.PrintCommunication = True
    On Error GoTo 0
End Sub
