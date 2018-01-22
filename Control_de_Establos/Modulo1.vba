Attribute VB_Name = "Modulo1"
' Ultima modificación: 19.11.2017
' Corrección PosibleCalor 19.11.17
Option Explicit
Dim ws As Worksheet
Dim mTabla As String

Sub MostrarMenu()
Attribute MostrarMenu.VB_Description = "Mostrar Menú Principal de Control de Establos"
Attribute MostrarMenu.VB_ProcData.VB_Invoke_Func = "m\n14"
    If VerifAccesoMod1 >= 4 Then
        usrVacas.Show
        'usrMenu.Show
      Else
        Application.Run "MsgAccesoNegado"
    End If
End Sub

Sub MostrarKardex()
Attribute MostrarKardex.VB_Description = "Mostrar la información contenida en el sistema de un animal determinado"
Attribute MostrarKardex.VB_ProcData.VB_Invoke_Func = "o\n14"
    Cells(ActiveCell.Row, 1).Select 'Columna A
    If IsEmpty(ActiveCell) Then
        MsgBox "No está posicionado en ninguna Tabla", _
          vbCritical, _
          "Consulta de Registro Individual"
        Exit Sub
    End If
    usrKardex.Show
End Sub

Private Sub Atrasadas()
    ' DEL >=60, Servicio=vbnullstring, Not DNB
    Dim sDate As String
    sDate = Format(Date - 334, "dd-mmm-yyyy")
    FiltrosParteI
    Worksheets("Reemplazos").Select
    Application.Run "Desproteger" 'Módulo2
    With ActiveSheet.ListObjects("Tabla2").Range
        .AutoFilter
        'Edad
        .AutoFilter Field:=4, Criteria1:=">=" & _
          sDate, Operator:=xlAnd
        'F.Servicio
        .AutoFilter Field:=7, Criteria1:="="
        'Clave1
        .AutoFilter Field:=12, Criteria1:="<>*DNB*", _
          Operator:=xlAnd
        ' Sexo
        .AutoFilter Field:=14, Criteria1:="=H", _
          Operator:=xlAnd
        .Range("A1").Select
    End With
    If Not CBool(Range("Configuracion!C39")) Then _
      Application.Run "Proteger" 'Módulo2
    Worksheets("Hato").Select
    Application.Run "Desproteger" 'Módulo2
    With ActiveSheet.ListObjects("Tabla1").Range
        .AutoFilter
        'DEL
        .AutoFilter Field:=4, Criteria1:=">=60", _
          Operator:=xlAnd
        'F.Servicio
        .AutoFilter Field:=8, Criteria1:="="
        'Clave1
        .AutoFilter Field:=14, Criteria1:="<>*DNB*", _
          Operator:=xlAnd
        .Range("A1").Select
    End With
    If Not CBool(Range("Configuracion!C39")) Then _
      Application.Run "Proteger" 'Módulo2
    FiltrosParteIII
End Sub

Private Sub PosibleCalor()
    Dim dDesde, dHasta As String
    
    dHasta = Format(Date - 18, "dd-mmm-yyyy")
    dDesde = Format(Date - 24, "dd-mmm-yyyy")
    
    FiltrosParteI
    Worksheets("Reemplazos").Select
    Application.Run "Desproteger" 'Módulo2
    With ActiveSheet.ListObjects("Tabla2").Range
        .AutoFilter
        ' F.Servicio
        .AutoFilter Field:=7, Criteria1:=">=" & dDesde, _
          Criteria2:="<=" & dHasta, Operator:=xlAnd
        'Clave1
        .AutoFilter Field:=12, Criteria1:="<>*DNB*", _
          Operator:=xlAnd
        .Range("A1").Select
    End With
    If Not CBool(Range("Configuracion!C39")) Then _
      Application.Run "Proteger" 'Módulo2
    Worksheets("Hato").Select
    Application.Run "Desproteger" 'Módulo2
    With ActiveSheet.ListObjects("Tabla1").Range
        .AutoFilter
        ' F.Servicio
        .AutoFilter Field:=8, Criteria1:=">=" & dDesde, _
          Criteria2:="<=" & dHasta, Operator:=xlAnd
        'Clave1
        .AutoFilter Field:=14, Criteria1:="<>*DNB*", _
          Operator:=xlAnd
        .Range("A1").Select
    End With
    If Not CBool(Range("Configuracion!C39")) Then _
      Application.Run "Proteger" 'Módulo2
    FiltrosParteIII
End Sub

Private Sub PorServir()
    ' DEL >=45, Servicio=vbnullstring, Not DNB
    Dim sDate As String
    sDate = Format(Date - 304, "dd-mmm-yyyy")
    FiltrosParteI
    Worksheets("Reemplazos").Select
    Application.Run "Desproteger" 'Módulo2
    With ActiveSheet.ListObjects("Tabla2").Range
        .AutoFilter
        ' Edad
        .AutoFilter Field:=5, Criteria1:="<=" & _
          sDate, Operator:=xlAnd
        ' F.Servicio
        .AutoFilter Field:=7, Criteria1:="=", _
          Operator:=xlAnd
        'Clave1
        .AutoFilter Field:=12, Criteria1:="<>*DNB*", _
          Operator:=xlAnd
        ' Sexo
        .AutoFilter Field:=14, Criteria1:="=H", _
          Operator:=xlAnd
        .Range("A1").Select
    End With
    If Not CBool(Range("Configuracion!C39")) Then _
      Application.Run "Proteger" 'Módulo2
    Worksheets("Hato").Select
    Application.Run "Desproteger" 'Módulo2
    With ActiveSheet.ListObjects("Tabla1").Range
        .AutoFilter
        'DEL
        .AutoFilter Field:=4, Criteria1:=">=" & _
          Range("Configuracion!C6"), _
          Operator:=xlAnd
        'F.Servicio
        .AutoFilter Field:=8, Criteria1:="="
        'Clave1
        .AutoFilter Field:=14, Criteria1:="<>*DNB*", _
          Operator:=xlAnd
        .Range("A2").Select
    End With
    If Not CBool(Range("Configuracion!C39")) Then _
      Application.Run "Proteger" 'Módulo2
    FiltrosParteIII
End Sub

Private Sub PorDxGestacion()
    ' DEL >=Configuracion!C5, Servicio Not "P", Not DNB
    Dim sDate As String
    sDate = Format(Date - _
      Range("Configuracion!C5"), "dd-mmm-yyyy")
    FiltrosParteI
    Worksheets("Reemplazos").Select
    Application.Run "Desproteger" 'Módulo2
    With ActiveSheet.ListObjects("Tabla2").Range
        .AutoFilter
        'F.Servicio
        .AutoFilter Field:=7, Criteria1:="<=" & _
          sDate, Operator:=xlAnd
        'Servicio
        .AutoFilter Field:=8, Criteria1:="<>*Calor*", _
          Operator:=xlAnd
        'Estatus
        .AutoFilter Field:=10, Criteria1:="<>P"
        'Clave1
        .AutoFilter Field:=12, Criteria1:="<>*DNB*", _
          Operator:=xlAnd
        .Range("A2").Select
    End With
    If Not CBool(Range("Configuracion!C39")) Then _
      Application.Run "Proteger" 'Módulo2
    Worksheets("Hato").Select
    Application.Run "Desproteger" 'Módulo2
    
    With ActiveSheet.ListObjects("Tabla1").Range
        .AutoFilter
        'F.Servicio
        .AutoFilter Field:=8, Criteria1:="<=" & _
          sDate, Operator:=xlAnd
        'Servicio
        .AutoFilter Field:=9, Criteria1:="<>*Calor*", _
          Operator:=xlAnd
        'Estatus
        .AutoFilter Field:=11, Criteria1:="<>P", _
          Operator:=xlOr
        'Clave1
        .AutoFilter Field:=14, Criteria1:="<>*DNB*", _
          Operator:=xlAnd
        .Range("A2").Select
    End With
    If Not CBool(Range("Configuracion!C39")) Then _
      Application.Run "Proteger" 'Módulo2
    FiltrosParteIII
End Sub

Private Sub PorSecar()
    Dim sDate As String
    sDate = Format(Date, "dd-mmm-yyyy")
    FiltrosParteI
    Worksheets("Hato").Select
    Application.Run "Desproteger" 'Módulo2
    With ActiveSheet.ListObjects("Tabla1").Range
        .AutoFilter
        'F.Servicio
        '.AutoFilter Field:=8, Criteria1:="<=" & _
          sDate, Operator:=xlAnd
        'Status
        '.AutoFilter Field:=11, Criteria1:="=P", _
          Operator:=xlAnd
        'FxSecar
        .AutoFilter Field:=12, Criteria1:="<=" & _
          sDate, _
          Operator:=xlAnd
        .Range("A2").Select
    End With
    If Not CBool(Range("Configuracion!C39")) Then _
      Application.Run "Proteger" 'Módulo2
    FiltrosParteII
End Sub

Private Sub PorParir()
    Dim sDate As String
    sDate = Format(Date - 266, "dd-mmm-yyyy")
    FiltrosParteI
    Worksheets("Reemplazos").Select
    Application.Run "Desproteger" 'Módulo2
    With ActiveSheet.ListObjects("Tabla2").Range
        .AutoFilter
        'F.Servicio
        .AutoFilter Field:=7, Criteria1:="<=" & _
          sDate, Operator:=xlAnd
        'Status
        .AutoFilter Field:=10, Criteria1:="=P", _
          Operator:=xlAnd
        .Range("A2").Select
    End With
    If Not CBool(Range("Configuracion!C39")) Then _
      Application.Run "Proteger" 'Módulo2
    Worksheets("Hato").Select
    Application.Run "Desproteger" 'Módulo2
    With ActiveSheet.ListObjects("Tabla1").Range
        .AutoFilter
        'F.Servicio
        .AutoFilter Field:=8, Criteria1:="<=" & _
          sDate, Operator:=xlAnd
        'Status
        .AutoFilter Field:=11, Criteria1:="=P", _
          Operator:=xlAnd
        .Range("A2").Select
    End With
    If Not CBool(Range("Configuracion!C39")) Then _
      Application.Run "Proteger" 'Módulo2
    FiltrosParteIII
End Sub

Private Sub PorDestetar()
    'Edad>=2, Corral<>8
    Dim sDate As String
    sDate = Format(Date - _
      Range("Configuracion!C34"), "dd-mmm-yyyy")
    FiltrosParteI
    Worksheets("Reemplazos").Select
    Application.Run "Desproteger" 'Módulo2
    With ActiveSheet.ListObjects("Tabla2").Range
        .AutoFilter
        ' F.Nacim
        .AutoFilter Field:=5, Criteria1:="<=" & _
          sDate, Operator:=xlAnd
        'Corral
        .AutoFilter Field:=2, Criteria1:="=8", _
          Operator:=xlAnd
        .Range("A1").Select
    End With
    If Not CBool(Range("Configuracion!C39")) Then _
      Application.Run "Proteger" 'Módulo2
    FiltrosParteII
End Sub

Private Sub PorVacunar()
    'Edad>=90 días
    Dim sDate As String
    sDate = Format(Date - 90, "dd-mmm-yyyy")
    FiltrosParteI
    Worksheets("Reemplazos").Select
    Application.Run "Desproteger" 'Módulo2
    With ActiveSheet.ListObjects("Tabla2").Range
        .AutoFilter
        ' F.Nacim
        .AutoFilter Field:=5, Criteria1:="<=" & _
          sDate, Operator:=xlAnd
        'F.Servicio
        .AutoFilter Field:=18, Criteria1:="=", _
          Operator:=xlAnd
        .Range("A1").Select
    End With
    If Not CBool(Range("Configuracion!C39")) Then _
      Application.Run "Proteger" 'Módulo2
    FiltrosParteII
End Sub

Private Sub PorImantar()
    'Edad>=1 Año, Iman=F
    Dim sDate As String
    sDate = Format(Date - 365, "dd-mmm-yyyy")
    Worksheets("Reemplazos").Select
    Application.Run "Desproteger" 'Módulo2
    With ActiveSheet.ListObjects("Tabla2").Range
        .AutoFilter
        'F.Nacim
        .AutoFilter Field:=5, Criteria1:="<=" & _
          sDate, Operator:=xlAnd
        'Clave1
        .AutoFilter Field:=12, Criteria1:="<>*DNB*", _
          Operator:=xlAnd
        'Sexo
        .AutoFilter Field:=14, Criteria1:="H", _
          Operator:=xlAnd
        'F.Imantación
        '.AutoFilter Field:=19, Criteria1:="=", _
          Operator:=xlAnd
        .Range("A1").Select
    End With
    If Not CBool(Range("Configuracion!C39")) Then _
      Application.Run "Proteger" 'Módulo2
    Worksheets("Hato").Select
    Application.Run "Desproteger" 'Módulo2
    With ActiveSheet.ListObjects("Tabla1").Range
        .AutoFilter
        'Clave1
        .AutoFilter Field:=14, Criteria1:="<>*DNB*", _
          Operator:=xlAnd
        'F.Imantación
        '.AutoFilter Field:=16, Criteria1:="=", _
          Operator:=xlAnd
        .Range("A1").Select
    End With
    If Not CBool(Range("Configuracion!C39")) Then _
      Application.Run "Proteger" 'Módulo2
    FiltrosParteIII
End Sub

Private Sub Repetidoras()
    'Servicio>=4, Dx Not "P", not DNB
    FiltrosParteI
    Worksheets("Reemplazos").Select
    Application.Run "Desproteger" 'Módulo2
    With ActiveSheet.ListObjects("Tabla2").Range
        .AutoFilter
        'Servicios
        .AutoFilter Field:=6, Criteria1:=">=3", _
          Operator:=xlAnd
        'Estatus
        .AutoFilter Field:=10, Criteria1:="<>P", _
          Operator:=xlAnd
        'Clave1
        .AutoFilter Field:=12, Criteria1:="<>*DNB*", _
          Operator:=xlAnd
        .Range("A1").Select
    End With
    If Not CBool(Range("Configuracion!C39")) Then _
      Application.Run "Proteger" 'Módulo2
    Worksheets("Hato").Select
    Application.Run "Desproteger" 'Módulo2
    With ActiveSheet.ListObjects("Tabla1").Range
        .AutoFilter
        'Servicios
        .AutoFilter Field:=7, Criteria1:=">=4", _
          Operator:=xlAnd
        'Estatus
        .AutoFilter Field:=11, Criteria1:="<>P", _
          Operator:=xlAnd
        'Clave1
        .AutoFilter Field:=14, Criteria1:="<>*DNB*", _
          Operator:=xlAnd
        .Range("A1").Select
    End With
    If Not CBool(Range("Configuracion!C39")) Then _
      Application.Run "Proteger" 'Módulo2
    FiltrosParteIII
End Sub

Private Sub BajasProductoras()
    'DEL < 90, Prod <= Configuracion!C24, not Tb
    FiltrosParteI
    Worksheets("Hato").Select
    Application.Run "Desproteger" 'Módulo2
    With ActiveSheet.ListObjects("Tabla1").Range
        .AutoFilter
        'Producción
        .AutoFilter Field:=3, Criteria1:="<" _
          & Range("Configuracion!C24"), Operator:=xlAnd
        'Días en Leche
        .AutoFilter Field:=4, Criteria1:=">90", _
          Operator:=xlAnd
        .Range("A1").Select
    End With
    If Not CBool(Range("Configuracion!C39")) Then _
      Application.Run "Proteger" 'Módulo2
    FiltrosParteII
End Sub

Private Sub QuitarFiltros2()
' Se quitan filtros de ambas hojas y se mantiene en hoja
    FiltrosParteI
    Worksheets("Hato").Activate
    Application.Run "Desproteger" 'Módulo2
    mTabla = "Tabla1"
    QuitarFiltros3
    Application.Run "Proteger" 'Módulo2
    Worksheets("Reemplazos").Activate
    Application.Run "Desproteger" 'Módulo2
    mTabla = "Tabla2"
    QuitarFiltros3
    Application.Run "Proteger" 'Módulo2
    FiltrosParteIII
End Sub
    
Private Sub QuitarFiltros3()
    With ActiveSheet.ListObjects(mTabla).Range
        .AutoFilter
        .AutoFilter
    End With
End Sub

Private Sub FiltrosParteI()
    Application.ScreenUpdating = _
      CBool(Range("Desarrollador!B6"))
    Range("Desarrollador!B20") = "T"
    Set ws = ActiveSheet
End Sub

Private Sub FiltrosParteII()
    Range("Desarrollador!B20").Clear
    Application.ScreenUpdating = True
End Sub

Private Sub FiltrosParteIII()
    ws.Activate
    FiltrosParteII
End Sub

Private Function VerifAccesoMod1()
' Private para evitar su visualización
    VerifAccesoMod1 = 0
    On Error Resume Next
    If Range("Configuracion!C49") = "HERCE" Then _
      VerifAccesoMod1 = 14: GoTo 100
    VerifAccesoMod1 = Application.WorksheetFunction. _
      VLookup(Range("Configuracion!C49"), Range("Tabla7"), 3, False)
    On Error GoTo 0
100:
End Function

