Attribute VB_Name = "ModBarraComandos"
' Ultima modificación: 19.11.17
Option Explicit
Dim iFiltro As Long

Private Sub CreateToolBar()
    Dim cbar As Office.CommandBar
    Dim bExists As Boolean
    
    'Para evitar duplicar la barra de comandos
    bExists = False
    For Each cbar In Application.CommandBars
        If cbar.Name = "ControlEstablos" Then bExists = True
    Next
    
    If Not bExists Then
        CreatePopup
        'cbar.Visible = True
    End If
End Sub

Private Sub CreatePopup()
    ' Orden de la Barra de Control
    Dim cbpop As CommandBarControl
    Dim cbctl As CommandBarControl 'Mostrar Todo
    Dim cbsub1 As CommandBarControl 'Filtrar Por ...
    Dim cbctl1 As CommandBarControl 'Posible Calor
    Dim cbctl2 As CommandBarControl 'Por Servir
    Dim cbctl3 As CommandBarControl 'Dx Gestacion
    Dim cbctl4 As CommandBarControl 'Por Secar
    Dim cbctl5 As CommandBarControl 'Por Parir
    Dim cbctl6 As CommandBarControl 'Por Destetar
    Dim cbctl7 As CommandBarControl 'Por Vacunar
    Dim cbctl8 As CommandBarControl 'Por Imantar
    Dim cbctl9 As CommandBarControl 'Repetidoras
    Dim cbctl10 As CommandBarControl 'Bajas Productoras
    Dim cbsub3 As CommandBarControl 'Captura Informacion ...
    Dim cbct2 As CommandBarControl 'Hato ...
    Dim cbct31 As CommandBarControl 'Reemplazos ...
    Dim cbct32 As CommandBarControl 'Guardar parámetros mensuales
    Dim cbct33 As CommandBarControl 'Importar Datos Externos ...
    Dim cbsub4 As CommandBarControl 'Consultar ...
    Dim cbct41 As CommandBarControl 'Actividades Semanales
    Dim cbct42 As CommandBarControl 'Notas de atención
    Dim cbct43 As CommandBarControl 'Bitácora de Eventos ...
    Dim cbct44 As CommandBarControl 'Última captura por animal ...
    Dim cbct50 As CommandBarControl 'Kardex ...
    Dim cbct60 As CommandBarControl 'Parámetros del Hato
    Dim cbsub7 As CommandBarControl 'Complementos ...
    Dim cbct71 As CommandBarControl 'Tablero de control
    Dim cbct72 As CommandBarControl 'Análisis Día de Prueba
    Dim cbct73 As CommandBarControl 'Análisis de Fertilidad
    Dim cbct74 As CommandBarControl 'Análisis Pesaje de Reemplazos
    Dim cbsub8 As CommandBarControl 'Utilerías ...
    Dim cbct81 As CommandBarControl 'Modificar información capturada ...
    Dim cbct82 As CommandBarControl 'Respaldar Base de Datos
    Dim cbct83 As CommandBarControl 'Importar Respaldo ...
    Dim cbct84 As CommandBarControl 'Reparar Base de Datos
    Dim cbct90 As CommandBarControl 'Configuraciones
    Dim cbct100 As CommandBarControl 'Acerca de Control de Establo
    
    ' Create a popup control on the main menu bar
    Set cbpop = Application.CommandBars("Worksheet Menu Bar"). _
      Controls.Add(Type:=msoControlPopup, Temporary:=True)
    cbpop.Caption = "ControlEstablos"
    cbpop.Visible = True
    
    'Add a menu item
    Set cbctl = cbpop.Controls.Add(Type:=msoControlButton)
    With cbctl
        .Visible = True
        .Style = msoButtonCaption  'Required for caption
        .Caption = "Quitar filtros"
        .OnAction = "Opcioncbct10" 'Action to perform
        If VerificarAcceso() = 0 Then .Enabled = False
    End With
    
    'Add a popup for a submenu
    Set cbsub1 = cbpop.Controls.Add(Type:=msoControlPopup)
    With cbsub1
        .Visible = True
        .Caption = "Filtrar por ..."
        If VerificarAcceso() = 0 Then .Enabled = False
    End With
    
    'Add Item for a submenu
    Set cbctl1 = cbsub1.Controls.Add(Type:=msoControlButton)
    With cbctl1
        .Visible = True
        .Style = msoButtonCaption 'Required for caption
        .Caption = "Posible Calor"
        .OnAction = "Opcioncbct11" 'Action to perform
    End With
    
    Set cbctl2 = cbsub1.Controls.Add(Type:=msoControlButton)
    With cbctl2
        .Visible = True
        .Style = msoButtonCaption 'Required for caption
        .Caption = "por Servir"
        .OnAction = "Opcioncbct12" 'Action to perform
    End With
    
    Set cbctl3 = cbsub1.Controls.Add(Type:=msoControlButton)
    With cbctl3
        .Visible = True
        .Style = msoButtonCaption 'Required for caption
        .Caption = "Dx Gestación"
        .OnAction = "Opcioncbct13" 'Action to perform
    End With
    
    Set cbctl4 = cbsub1.Controls.Add(Type:=msoControlButton)
    With cbctl4
        .Visible = True
        .Style = msoButtonCaption 'Required for caption
        .Caption = "por Secar"
        .OnAction = "Opcioncbct14" 'Action to perform
    End With
    
    Set cbctl5 = cbsub1.Controls.Add(Type:=msoControlButton)
    With cbctl5
        .Visible = True
        .Style = msoButtonCaption 'Required for caption
        .Caption = "por Parir"
        .OnAction = "Opcioncbct15" 'Action to perform
    End With
    
    Set cbctl6 = cbsub1.Controls.Add(Type:=msoControlButton)
    With cbctl6
        .Visible = True
        .Style = msoButtonCaption 'Required for caption
        .Caption = "por Destetar"
        .OnAction = "Opcioncbct16" 'Action to perform
    End With

    Set cbctl7 = cbsub1.Controls.Add(Type:=msoControlButton)
    With cbctl7
        .Visible = True
        .Style = msoButtonCaption 'Required for caption
        .Caption = "por Vacunar"
        .OnAction = "Opcioncbct17" 'Action to perform
    End With

    Set cbctl8 = cbsub1.Controls.Add(Type:=msoControlButton)
    With cbctl8
        .Visible = True
        .Style = msoButtonCaption 'Required for caption
        .Caption = "por &Imantar"
        .OnAction = "Opcioncbct18" 'Action to perform
    End With

    Set cbctl9 = cbsub1.Controls.Add(Type:=msoControlButton)
    With cbctl9
        .Visible = True
        .Style = msoButtonCaption
        .Caption = "Repetidoras"
        .OnAction = "Opcioncbct19" 'Action to perform
    End With
    
    Set cbctl10 = cbsub1.Controls.Add(Type:=msoControlButton)
    With cbctl10
        .Visible = True
        .Style = msoButtonCaption 'Required for caption
        .Caption = "Bajas Productoras"
        .OnAction = "Opcioncbct110" 'Action to perform
    End With
    
    Set cbct2 = cbpop.Controls.Add(Type:=msoControlButton)
    With cbct2
        .Visible = True
        .Style = msoButtonCaption
        .Caption = "Registro de Eventos ..."
        .OnAction = "Opcioncbct2" 'Action to perform
        If VerificarAcceso() < 4 Then .Enabled = False
    End With
    
    'Add a popup for a submenu
    Set cbsub3 = cbpop.Controls.Add(Type:=msoControlPopup)
    With cbsub3
        .Visible = True
        .Caption = "Otras capturas ..."
    End With
    
    Set cbct31 = cbsub3.Controls.Add(Type:=msoControlButton)
    With cbct31
        .Visible = True
        .Style = msoButtonCaption 'Required for caption
        .Caption = "Compras de semen ..."
        .OnAction = "ExampleMacro" 'Action to perform
        If VerificarAcceso() < 4 Then .Enabled = False
    End With
    
    Set cbct32 = cbsub3.Controls.Add(Type:=msoControlButton)
    With cbct32
        .Visible = True
        .Style = msoButtonCaption 'Required for caption
        .Caption = "Generar Indicadores del mes"
        .OnAction = "Opcioncbct32" 'Action to perform
        If VerificarAcceso() < 4 Then .Enabled = False
    End With
    
    Set cbct33 = cbsub3.Controls.Add(Type:=msoControlButton)
    With cbct33
        .Visible = True
        .Style = msoButtonCaption 'Required for caption
        .Caption = "Importar datos externos ..."
        .OnAction = "Opcioncbct33" 'Action to perform
        If VerificarAcceso() < 8 Then .Enabled = False
    End With
    
    'Add a popup for a submenu
    Set cbsub4 = cbpop.Controls.Add(Type:=msoControlPopup)
    With cbsub4
        .Visible = True
        .Caption = "Consultar ..."
    End With
       
     'Add a menu item
    Set cbct41 = cbsub4.Controls.Add(Type:=msoControlButton)
    With cbct41
        .Visible = True
        .Style = msoButtonCaption 'Required for caption
        .Caption = "Actividades por realizar próximos 8 días"
        .OnAction = "Opcioncbct41" 'Action to perform
        If VerificarAcceso() = 0 Then .Enabled = False
    End With
    
     'Add a menu item
    Set cbct42 = cbsub4.Controls.Add(Type:=msoControlButton)
    With cbct42
        .Visible = True
        .Style = msoButtonCaption 'Required for caption
        .Caption = "Notas de atención generados por sistema ..."
        .OnAction = "ExampleMacro" 'Action to perform
        If VerificarAcceso() = 0 Then .Enabled = False
    End With
    
     'Add a menu item
    Set cbct43 = cbsub4.Controls.Add(Type:=msoControlButton)
    With cbct43
        .Visible = True
        .Style = msoButtonCaption 'Required for caption
        .Caption = "Consulta a Base de Datos ..."
        .OnAction = "Opcioncbct43" 'Action to perform
        If VerificarAcceso() = 0 Then .Enabled = False
    End With
    
     'Add a menu item
    Set cbct44 = cbsub4.Controls.Add(Type:=msoControlButton)
    With cbct44
        .Visible = True
        .Style = msoButtonCaption 'Required for caption
        .Caption = "Última captura por animal ..."
        .OnAction = "Opcioncbct44" 'Action to perform
        If VerificarAcceso() = 0 Then .Enabled = False
    End With
    
    'Add a menu item
    Set cbct50 = cbpop.Controls.Add(Type:=msoControlButton)
    With cbct50
        .Visible = True
        .Style = msoButtonCaption 'Required for caption
        .Caption = "Kardex ..."
        .OnAction = "Opcioncbct50" 'Action to perform
        'If VerificarAcceso() = 0 Then .Enabled = False
    End With
    
    'Add a menu item
    Set cbct60 = cbpop.Controls.Add(Type:=msoControlButton)
    With cbct60
        .Visible = True
        .Style = msoButtonCaption 'Required for caption
        .Caption = "Indicadores del desempeño"
        .OnAction = "Opcioncbct60" 'Action to perform
        If VerificarAcceso() = 0 Then .Enabled = False
    End With
    
    'Add a menu item
    Set cbsub7 = cbpop.Controls.Add(Type:=msoControlPopup)
    With cbsub7
        .Visible = True
        '.Style = msoButtonCaption 'Required for caption
        .Caption = "Análisis complementarios ..."
        '.Enabled = False
        '.OnAction = "Opcioncbct40" 'Action to perform
    End With
    
    'Add a popup for a submenu
    Set cbct71 = cbsub7.Controls.Add(Type:=msoControlButton)
    With cbct71
        .Visible = True
        .Style = msoButtonCaption
        .Caption = "Análisis del hato"
        .OnAction = "AbrirTableroDeControl"
        '.OnAction = "ExampleMacro" 'Action to perform
        If VerificarAcceso() < 8 Then .Enabled = False
    End With
    
    'Add a popup for a submenu
    Set cbct72 = cbsub7.Controls.Add(Type:=msoControlButton)
    With cbct72
        .Visible = True
        .Style = msoButtonCaption
        .Caption = "Día de prueba"
        .OnAction = "ExampleMacro" 'Action to perform
        If VerificarAcceso() < 8 Then .Enabled = False
    End With
    
    'Add a popup for a submenu
    Set cbct73 = cbsub7.Controls.Add(Type:=msoControlButton)
    With cbct73
        .Visible = True
        .Style = msoButtonCaption
        .Caption = "Fertilidad de toros"
        .OnAction = "Opcioncbct73" 'Action to perform
        If VerificarAcceso() < 8 Then .Enabled = False
    End With

    'Add a popup for a submenu
    Set cbct74 = cbsub7.Controls.Add(Type:=msoControlButton)
    With cbct74
        .Visible = True
        .Style = msoButtonCaption
        .Caption = "Pesaje de Reemplazos"
        .OnAction = "ExampleMacro" 'Action to perform
        If VerificarAcceso() < 8 Then .Enabled = False
    End With
    
    'Add a menu item
    Set cbsub8 = cbpop.Controls.Add(Type:=msoControlPopup)
    With cbsub8
        .Visible = True
        '.Style = msoButtonCaption 'Required for caption
        .Caption = "Utilerías ..."
        '.Enabled = False
        '.OnAction = "Opcioncbct40" 'Action to perform
    End With
    
    'Add a popup for a submenu
    Set cbct81 = cbsub8.Controls.Add(Type:=msoControlButton)
    With cbct81
        .Visible = True
        .Style = msoButtonCaption
        .Caption = "Editar información capturada ..."
        .OnAction = "ExampleMacro" 'Action to perform
        If VerificarAcceso() < 8 Then .Enabled = False
    End With
        
    'Add a popup for a submenu
    Set cbct82 = cbsub8.Controls.Add(Type:=msoControlButton)
    With cbct82
        .Visible = True
        .Style = msoButtonCaption
        .Caption = "Respaldar información"
        .OnAction = "ExampleMacro" 'Action to perform
        If VerificarAcceso() < 4 Then .Enabled = False
    End With
    
    'Add a popup for a submenu
    Set cbct83 = cbsub8.Controls.Add(Type:=msoControlButton)
    With cbct83
        .Visible = True
        .Style = msoButtonCaption
        .Caption = "Subir respaldo ..."
        .OnAction = "ExampleMacro" 'Action to perform
        If VerificarAcceso() < 8 Then .Enabled = False
    End With
    
    'Add a popup for a submenu
    Set cbct84 = cbsub8.Controls.Add(Type:=msoControlButton)
    With cbct84
        .Visible = True
        .Style = msoButtonCaption
        .Caption = "Reparar base de datos"
        .Enabled = True
        .OnAction = "Opcioncbct84" 'Action to perform
        If VerificarAcceso() < 8 Then .Enabled = False
    End With
        
    'Add a menu item
    Set cbct90 = cbpop.Controls.Add(Type:=msoControlButton)
    With cbct90
        .Visible = True
        .Style = msoButtonCaption 'Required for caption
        .Caption = "Configuraciones ..."
        .OnAction = "Opcioncbct90" 'Action to perform
        If VerificarAcceso() < 8 Then .Enabled = False
    End With
    
    'Add a menu item
    Set cbct100 = cbpop.Controls.Add(Type:=msoControlButton)
    With cbct100
        .Visible = True
        .Style = msoButtonCaption 'Required for caption
        .Caption = "Acerca de Control de Establos"
        .OnAction = "Opcioncbct100" 'Action to perform
    End With
    
End Sub
    
Private Sub Opcioncbct10()
    iFiltro = 2048
    Application.Run "QuitarFiltros2" 'Modulo1
    BarraDeEstado
End Sub

Private Sub Opcioncbct11()
    iFiltro = 1024
    Application.Run "PosibleCalor" 'Modulo1
    BarraDeEstado
End Sub

Private Sub Opcioncbct12()
    iFiltro = 512
    Application.Run "PorServir" 'Modulo1
    BarraDeEstado
End Sub

Private Sub Opcioncbct13()
    iFiltro = 256
    Application.Run "PorDxGestacion" 'Modulo1
    BarraDeEstado
End Sub

Private Sub Opcioncbct14()
    iFiltro = 128
    Application.Run "PorSecar" 'Modulo1
    BarraDeEstado
End Sub

Private Sub Opcioncbct15()
    iFiltro = 64
    Application.Run "PorParir" 'Modulo1
    BarraDeEstado
End Sub

Private Sub Opcioncbct16()
    iFiltro = 32
    Application.Run "PorDestetar" 'Modulo1
    BarraDeEstado
End Sub

Private Sub Opcioncbct17()
    iFiltro = 16
    Application.Run "PorVacunar" 'Modulo1
    BarraDeEstado
End Sub

Private Sub Opcioncbct18()
    iFiltro = 8
    Application.Run "PorImantar" 'Modulo1
    BarraDeEstado
End Sub

Private Sub Opcioncbct19()
    iFiltro = 4
    Application.Run "Repetidoras" 'Modulo1
    BarraDeEstado
End Sub

Private Sub Opcioncbct110()
    iFiltro = 2
    Application.Run "BajasProductoras" 'Modulo1
    BarraDeEstado
End Sub

Private Sub Opcioncbct2()
    usrVacas.Show
End Sub

Private Sub Opcioncbct31()
    'usrReemplazos.Show
End Sub

Private Sub Opcioncbct32()
    Application.Run "Instamatic" 'ModParametros
End Sub

Private Sub Opcioncbct33()
    Application.Run "PrepImportarDatos" 'ModImportarInfo
End Sub
    
Private Sub Opcioncbct41()
    ExampleMacro
End Sub

Private Sub Opcioncbct43()
    usrQuery.Show
End Sub

Private Sub Opcioncbct44()
    usrUltimaCaptura.Show
End Sub

Private Sub Opcioncbct50()
    MostrarKardex
End Sub

Private Sub Opcioncbct60()
    usrEstadísticas.Show
End Sub

Private Sub Opcioncbct73()
    Application.Run "AnalisisFertilidad" 'ModFertilidad
End Sub

Private Sub Opcioncbct84()
    Application.Run "RepararBaseDatos" 'ModUtilerias
End Sub

Private Sub Opcioncbct90()
    usrContraseña.Show
End Sub

Private Sub Opcioncbct100()
    Application.Run "Info"
    If Application.UserName = "JPHC" Then
        MsgBox "Versión " & Range("Desarrollador!B19") & " (En desarrollo)" _
          & Chr(13) & "Se implementa navegación en Kardex" _
          & Chr(13) & "Se trabaja con Filtros", _
          vbInformation, "JP´s Development Labs"
    End If
End Sub

Private Sub ExampleMacro()
    Application.Run "UnderConstruction"
End Sub

Private Sub BarraDeEstado()
' Muestra una leyenda en la Barra de Estado de Acuerdo al filtro aplicado
    Dim sMensaje As String
    Application.DisplayStatusBar = True
    Select Case iFiltro
        Case 2048
            Application.StatusBar = False
        Case 1024
            Application.StatusBar = _
              "Filtrado por Animales en Posible Calor"
        Case 512
            Application.StatusBar = _
              "Filtrado por Animales por Servir"
        Case 256
            Application.StatusBar = _
              "Filtrado por Animales por Dx. de Gestación"
        Case 128
            Application.StatusBar = _
              "Filtrado por Animales por Secar"
        Case 64
            Application.StatusBar = _
              "Filtrado por Animales por Parir"
        Case 32
            Application.StatusBar = _
              "Filtrado por Animales por Destetar"
        Case 16
            Application.StatusBar = _
              "Filtrado por Animales por Vacunar"
        Case 8
            Application.StatusBar = _
              "Filtrado por Animales por Imantar"
        Case 4
            Application.StatusBar = _
              "Filtrado por Animales Repetidores"
        Case 2
            Application.StatusBar = _
              "Filtrado por Bajas Productoras"
    End Select
End Sub

Private Function VerificarAcceso()
    VerificarAcceso = 0
    On Error Resume Next
    If Range("Configuracion!C49") = "HERCE" Then _
      VerificarAcceso = 14: GoTo 100
    VerificarAcceso = Application.WorksheetFunction. _
      VLookup(Range("Configuracion!C49"), Range("Tabla7"), 3, False)
    On Error GoTo 0
100:
End Function

Private Sub AbrirTableroDeControl()
    Dim sPath, sArchivo As String
    'Actualizar datos que van a ser analizados
    sPath = Application.ActiveWorkbook.Path
    sArchivo = "Análisis.xlsm"
        'sArchivo = "DashBoardControlDeEstablos.xlsx"
    Workbooks.Open sPath & "\" & sArchivo
End Sub
