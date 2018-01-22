Attribute VB_Name = "ModSeguridad"
' Ultima modificación: 17-Nov-2017
' Incluir Application.Run "PrepararDistribucion" en _
  Auto_Open
Option Explicit
Dim FSO As New FileSystemObject
Dim bDemoProof As Boolean
Dim sTabla As String
Dim ws As Worksheet
' Debe estar habilitado el Microsoft Scripting Runtime

Private Sub Auto_Open()
' Creo una variable del tipo Disco
Dim dEsteDisco As Drive
Dim sRutaSO, sPath As String
Dim Serial

    Application.ScreenUpdating = CBool(Range("Desarrollador!B6"))
    ' Directorio del sistema
    sPath = Application.ActiveWorkbook.Path
    ' MsgBox (sPath)
    sRutaSO = Environ("windir")
    ' MsgBox (sRutaSO)
    ' Establecer este disco
    sRutaSO = Left(sRutaSO, 3)
    ' MsgBox (sRutaSO)
    Set dEsteDisco = FSO.GetDrive(sRutaSO)
    Serial = dEsteDisco.SerialNumber
    'Si la celda está vacía se guarda el Num. de Serie y Nombre de Archivo
    If Range("Desarrollador!A1") = vbNullString Then
            Range("Desarrollador!A1") = Serial
            Range("Desarrollador!A2") = Application.ActiveWorkbook.Path
            Range("Desarrollador!A3") = ThisWorkbook.Name
            If Range("Desarrollador!B13") = True And _
              IsEmpty(Range("Desarrollador!B14")) Then _
              Range("Desarrollador!B14") = Date
        Else
            ' Deben coincidir Num. de Serie y Nombre de Archivo
            If Range("Desarrollador!A1") <> Serial Or _
              Range("Desarrollador!A2") <> Application.ActiveWorkbook.Path _
              Or Range("Desarrollador!A3") <> ThisWorkbook.Name Then
                    MsgBox "Ud no está autorizado para usar este programa." & _
                     Chr(13) & Chr(13) & _
                    "Si desea una copia contacte a" & Chr(13) _
                    & Chr(13) & _
                    "     jpherce@gmail.com", vbCritical + vbOKOnly, _
                      "JP's Automatización de Aplicaciones"
            Application.Run "PrepararDistribucion" 'ModSeguridad
                    ActiveWorkbook.Close xlDoNotSaveChanges
            Else
                    ' Si es un demo
                    DemoProof
                    If bDemoProof = False Then Exit Sub
                    ' Mostrar hoja de inicio
                    MsgBox "Sistema de Control de Establos" _
                      & Chr(13) & "Versión: " & _
                      Range("Desarrollador!B19") _
                      & Chr(13) & Chr(13) & "Num. de Serie: " & _
                      dEsteDisco.SerialNumber, _
                      vbInformation, Range("Configuracion!C3")
                    If Application.UserName = "JPHC" Then
                        'Range("Desarrollador!A2") = 1
                    End If
                    If CBool(Range("Configuracion!C28")) Then 'Contraseña requerida
                        usrLogin.Show
                    
                    End If
            End If
    End If
    Range("Desarrollador!B21") = Date
    'Application.Run "CreateToolBar"
    Application.Run "CreatePopup" 'Crear Barra de Comandos
    MostrarHojas
    Application.Run "QuitarFiltros2" 'Módulo1
    'Application.Run "FCH"
    Range("A2").Select
    PrepararBaseDatos
    ActiveWorkbook.Save
    ' Liberar recursos del sistema
    Set FSO = Nothing
    Set dEsteDisco = Nothing
    Application.ScreenUpdating = True
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
' **** Este código debe estar en 'ThisWorkBook'" ****
    Application.Run "PrepararDistribucion" 'ModSeguridad
    Application.Run "CerrarTodo" 'Módulo2
    Range("A1").Activate
    ActiveWorkbook.Save
End Sub

Private Sub BorrarTodo()
Dim oldStatusBar As String
' Borrar todo y dejar sistema para nueva instalación
    oldStatusBar = Application.DisplayStatusBar
    ' Borrar Tabla1
    Application.StatusBar = "Borrando Tabla1..."
    Set ws = Worksheets("Hato")
    sTabla = "Tabla1"
    BorrarTabla
    'Range("XFD1").Clear
    Range("Desarrollador!B20").Clear
    ' Borrar Tabla2
    Application.StatusBar = "Borrando Tabla2..."
    Set ws = Worksheets("Reemplazos")
    sTabla = "Tabla2"
    BorrarTabla
    'Range("XFD1").Clear
    Range("Desarrollador!B20") = UCase("T")
    ' Borrar Tabla3
    Application.StatusBar = "Borrando Tabla3..."
    Set ws = Worksheets("InventarioSemen")
    sTabla = "Tabla3"
    BorrarTabla
    ' Borrar Tabla4
    Application.StatusBar = "Borrando Tabla4..."
    Set ws = Worksheets("LactanciasAnteriores")
    sTabla = "Tabla4"
    BorrarTabla
    ' Borrar Tabla5
    Application.StatusBar = "Borrando Tabla5..."
    Set ws = Worksheets("BajaReemplazos")
    sTabla = "Tabla5"
    BorrarTabla
    ' Borrar Tabla6
    Application.StatusBar = "Borrando Tabla6..."
    Set ws = Worksheets("Eventos")
    sTabla = "Tabla6"
    BorrarTabla
    ' Borrar Tabla15
    Application.StatusBar = "Borrando Tabla7..."
    Set ws = Worksheets("Hato2")
    sTabla = "Tabla15"
    BorrarTabla
    Application.StatusBar = "Borrando Tabla8..."
    Set ws = Worksheets("InfoVitalicia")
    sTabla = "Tabla8"
    BorrarTabla
    Range("Desarrollador!B6") = False 'ModoDepuracion
    Range("Desarrollador!B7") = False 'Display FormulaBar
    Range("Desarrollador!B8") = False 'Display Headings
    Range("Desarrollador!B9") = False 'Display Gridlines
    Range("Desarrollador!B11") = 4321 'ClaveConfigurador
    Range("Desarrollador!B12") = False 'UsuarioConfigurar
    Range("Desarrollador!B13") = True 'VersionDemo
    Range("Desarrollador!B14") = vbNullString 'FechaInicioDemo
    Range("Desarrollador!B15") = 1234 'ClaveUsuario
    Range("Configuracion!C3") = "Aquí Va Su Nombre" 'Empresa
    Range("Configuracion!C5") = 45 'dDxGest
    Range("Configuracion!C6") = 45 'dWait
    Range("Configuracion!C7") = False 'ReqMagnet
    Range("Configuracion!C9") = 6 'Secas
    Range("Configuracion!C10") = 7 'Preparacion
    Range("Configuracion!C11") = 2 'RecienParidas
    Range("Configuracion!C12") = 1 'Vaquillas
    Range("Configuracion!C13") = 8 'Lactancia
    Range("Configuracion!C15") = False 'ReqSemental
    Range("Configuracion!C16") = False 'ReqTecnico
    Range("Configuracion!C17") = False 'ReqInventario
    Range("Configuracion!C19") = False 'NomPadre
    Range("Configuracion!C20") = False 'AreteMadre
    Range("Configuracion!C21") = False 'ReqRaza
    Range("Configuracion!C22") = False 'ReqFNacim
    Range("Configuracion!C24") = 20 'ProdMin
    Range("Configuracion!C25") = True 'RespaldoCSV
    Range("Configuracion!C27") = False 'Caprturista
    Range("Configuracion!C28") = False 'Contraseña
    Range("Configuracion!C30") = True 'ConsecutivoReemplazos
    Range("Configuracion!C31") = 1000 'Id Inicial Hembras
    Range("Configuracion!C32") = True 'Control de Reemplazos
    Range("Configuracion!C33") = False 'Consecutivo Machos
    Range("Configuracion!C34") = 45 'Días para destete
    Range("Configuracion!C35") = False 'Control peso corporal hembras
    Range("Configuracion!C36") = 200 'Ajuste de curva lactancia
    Range("Configuracion!C38") = False 'Hojas visibles
    Range("Configuracion!C39") = False 'Modificar Tablas por usuario
    Range("Configuracion!B40") = True 'Hato
    Range("Configuracion!B41") = True 'Recrías
    Range("Configuracion!B42") = False 'Semen
    Range("Configuracion!B43") = False 'LactAnt
    Range("Configuracion!B44") = False 'BajaReemplazos
    MostrarHojas
    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar
End Sub

Private Sub BorrarTabla()
Dim iR As Long
Dim i As Long
    ws.Visible = xlSheetVisible
    ws.Select
    Application.Run "Desproteger" 'Mód2
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

Private Sub DemoProof()
' Checar y mostrar los días restantes dde la funcion demo
    Dim d As Long
    Dim t As String
    bDemoProof = True
    If CBool(Range("Desarrollador!B13")) Then
        d = 30 - (Date - CDate(Range("Desarrollador!B14")))
        If Date - CDate(Range("Desarrollador!b14")) < 30 Then _
          t = "Restan " & d & " Días a la licencia Demostrativa."
        If d < 0 Then
            t = "Licencia demostrativa ya vencida."
            bDemoProof = False
        End If
        MsgBox t & Chr(13) & Chr(13) & _
            "Contacte con su distribuidor al correo:" & Chr(13) & _
            Range("Desarrollador!B16").Text, vbInformation, _
            "Control de Establos 2.1"
    End If
End Sub

Private Sub Info()
Attribute Info.VB_Description = "Mostrar la información del sistema."
Attribute Info.VB_ProcData.VB_Invoke_Func = " \n14"
' Crear Variable
Dim dEsteDisco As Drive
Dim sRutaSO As String
    ' veamos donde está el directorio del sistema:
    sRutaSO = Environ("windir")
    ' y tomemos los primeors tres caracteres, que corresponden a la
    ' letra del disco:
    sRutaSO = Left(sRutaSO, 3)
    'seteo a dEsteDisco com el disco del sistema, así me garantizo de
    ' estar siempre leyendo la unidad principal de la PC:
    Set dEsteDisco = FSO.GetDrive(sRutaSO)
    MsgBox _
      "Sistema integral para el control del ganado lechero" _
      & Chr(13) & Chr(13) & _
      "Versión del sistema: " & Range("Desarrollador!B19") _
      & Chr(13) & Chr(13) & _
      "Licencia Otorgada a " & Range("Configuracion!C3") _
      & Chr(13) & _
      "No. Serie: " & dEsteDisco.SerialNumber & Chr(13) & _
      "Archivo: " & ThisWorkbook.Name & Chr(13) & _
      "Usuario: " & Application.UserName & Chr(13) & Chr(13) & _
      "Advertencia: Este programa está protegido por las leyes de derecho" _
      & Chr(13) & _
      "de autor y otros tratados internacionales. La reproducción o la" _
      & Chr(13) & _
      "distribución no autorizadas de este programa, o de cualquier parte" _
      & Chr(13) & _
      "del mismo, está penada por la ley con severas sanciones civiles y" _
      & Chr(13) & _
      "penales, y será objeto de todas las acciones judiciales que correspondan." _
      & Chr(13) & Chr(13) & Chr(13) & _
      "Contacto con el desarrollador: jpherce@gmail.com", _
      vbInformation, "Control de Establos"
    
    'Set FSO = Nothing
    Set dEsteDisco = Nothing
End Sub

Private Sub MostrarHojas()
' Mostrar Hojas de Acuerdo a lo establecido en Tabla Configuración
    ' Sólo se muestran en modo Depuración
    Sheets("Inicio").Visible = True
    If CBool(Range("Desarrollador!B6")) Then
            Sheets("Desarrollador").Visible = True
            Sheets("Configuracion").Visible = True
            Sheets("Eventos").Visible = True
            Sheets("Hato2").Visible = True
            'Sheets("Reportes").Visible = True
            Sheets("InfoVitalicia").Visible = True
        Else
            Sheets("Desarrollador").Visible = xlVeryHidden
            Sheets("Configuracion").Visible = xlVeryHidden
            Sheets("Eventos").Visible = xlVeryHidden
            Sheets("Hato2").Visible = xlVeryHidden
            'Sheets("Reportes").Visible = xlVeryHidden
            Sheets("InfoVitalicia").Visible = xlVeryHidden
    End If
    ' Se muestran de acuerdo a Configuración de usuario
    If CBool(Range("Desarrollador!B6")) Then
            Sheets("Hato").Visible = True
        Else
            If CBool(Range("Configuracion!B40")) Then _
              Sheets("Hato").Visible = True Else _
              Sheets("Hato").Visible = xlVeryHidden
    End If
    If CBool(Range("Desarrollador!B6")) Then
            Sheets("Reemplazos").Visible = True
        Else
            If CBool(Range("Configuracion!B41")) Then _
              Sheets("Reemplazos").Visible = True Else _
              Sheets("Reemplazos").Visible = xlVeryHidden
    End If
    If CBool(Range("Desarrollador!B6")) Then
            Sheets("InventarioSemen").Visible = True
        Else
            If CBool(Range("Configuracion!B42")) Then _
              Sheets("InventarioSemen").Visible = True Else _
              Sheets("InventarioSemen").Visible = xlVeryHidden
    End If
    If CBool(Range("Desarrollador!B6")) Then
            Sheets("LactanciasAnteriores").Visible = True
        Else
            If CBool(Range("Configuracion!B43")) Then _
              Sheets("LactanciasAnteriores").Visible = True _
              Else Sheets("LactanciasAnteriores").Visible = _
              xlVeryHidden
    End If
    If CBool(Range("Desarrollador!B6")) Then
            Sheets("BajaReemplazos").Visible = True
        Else
            If CBool(Range("Configuracion!B44")) Then _
              Sheets("BajaReemplazos").Visible = True Else _
              Sheets("BajaReemplazos").Visible = xlVeryHidden
    End If
    ' Se cierra hasta el final para evitar errores de _
      compilación
    If Not CBool(Range("Desarrollador!B6")) Then _
      Sheets("Hato2").Visible = xlVeryHidden
    If Not Sheets("Hato").Visible Then
            Sheets("Inicio").Visible = True
        Else
            Sheets("Inicio").Visible = xlVeryHidden
            Sheets("Hato").Select
    End If
    Application.Run "FinalLista" 'Módulo2
'    Range("A2").Select
End Sub

Private Sub NombreArchivo()
    MsgBox ThisWorkbook.Name
    If Not Range("Desarrollador!A3") = ThisWorkbook.Name Then
        MsgBox "No coinciden los nombres"
    End If
End Sub

Private Sub PrepararBaseDatos()
' Quitar Encabezados de las Hojas con Dases de Datos
    Application.DisplayFormulaBar = _
      CBool(Range("Desarrollador!B7"))
    ActiveWindow.DisplayHeadings = _
      CBool(Range("Desarrollador!B8"))
    ActiveWindow.DisplayGridlines = _
      CBool(Range("Desarrollador!B9"))
End Sub

Private Sub PrepararDistribucion()
    ' Quitar los seguros para la distribución de la aplicación
    Dim sRespuesta As String
    ' Control de Versiones
    If Application.UserName = "JPHC" Then
        sRespuesta = _
          MsgBox( _
          "¿Quitar seguros para que se pueda distribuir la aplicación?", _
          vbYesNo + vbDefaultButton2 + vbQuestion, _
          "JP's Automatización de Aplicaciones")
        If sRespuesta = vbYes Then
            With Sheets("Desarrollador")
                .Range("A1:A3").Clear
                '.Range("A2").Clear
                '.Range("A3").Clear
                .Range("B6") = False
                .Range("B13") = True
                sRespuesta = _
                  MsgBox("¿Borrar la Fecha de Inicio del Demo?", _
                  vbYesNo + vbDefaultButton2 + vbQuestion, _
                  "JP's Automatización de Aplicaciones")
                If sRespuesta = vbYes Then .Range("B14").Clear
            End With
        End If
    End If
End Sub

Private Sub QuitarBanderaDemo()
    ' Quitar la Bandera de Demo
    Dim sRespuesta As String
    sRespuesta = _
      InputBox("Introducir Clave de Activación", _
      "JP's Automatización de Aplicaciones")
    If sRespuesta = Hex(Date * 1691) Then
            Range("Desarrollador!B13") = False
            MsgBox "Se ha concedido licencia de uso de este sistema", _
            vbExclamation, _
            "JP's Automatización de Aplicaciones"
        Else
            MsgBox "Lo siento, la clave es incorrecta", _
            vbCritical, _
            "JP's Automatización de Aplicaciones"
    End If
End Sub

Private Sub ReestablecerCondicionesOriginales()
' Reponer Encabezados de las Hojas con Dases de Datos
    Application.DisplayFormulaBar = True
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayGridlines = True
End Sub

Private Sub GetReady()
    Dim i As Long
    Worksheets("Reemplazos").Select
    PrepararBaseDatos
    Worksheets("Hato").Select
    PrepararBaseDatos
End Sub

Private Sub MsgAccesoNegado()
    MsgBox "Ud. no tiene acceso a esta función", _
      vbCritical, _
      "JP's Automatización de Aplicaciones"
End Sub

