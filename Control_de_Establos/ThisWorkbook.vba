VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Ultima modificación: 15-Sep-2015
Option Explicit
    Dim FuncName As String
    Dim FunDesc As String
    Dim Category As String
    Dim ArgDesc(1 To 3) As String

Private Sub Workbook_BeforeClose(Cancel As Boolean)
' **** Este código debe estar en 'ThisWorkBook'" ****
    Application.ScreenUpdating = _
      CBool(Range("Desarrollador!B6"))
    Application.Run "PrepararDistribucion" 'ModSeguridad
    If Not CBool(Range("Desarrollador!B6")) Then _
      Application.Run "CerrarTodo"  'Módulo2"
    Range("Desarrollador!B21").Clear 'FechaInicioSesion
    Range("A1").Activate
    Application.Run "ReestablecerCondicionesOriginales" 'ModSeguridad
    ActiveWorkbook.Save
    Application.ScreenUpdating = True
End Sub

Private Sub Workbook_BeforePrint(Cancel As Boolean)
' Configuración de hojas de impresión
    'Impresión.ConfigHojaImpresion
End Sub

Private Sub Workbook_Open()
    Dim sRespuesta As String
    If Application.UserName = "JPHC" Then
        sRespuesta = _
          MsgBox( _
          "¿Actualizar código desde Control de Versiones?", _
          vbYesNo + vbDefaultButton2 + vbQuestion, _
          "JP's Automatización de Aplicaciones")
        If sRespuesta = vbYes Then
            Application.Run "ImportCodeMod" 'Mod ControlVersiones
        End If
    End If
' Establece ayuda en las UDF
    'Dim FuncName As String
    'Dim FunDesc As String
    'Dim Category As String
    'Dim ArgDesc(1 To 3) As String
    
    'FuncName = pDAb
    'FunDesc = "Calcula el  promedio de días abiertos"
    'Category = 14
    'ArgDesc(1) = "Número de lactancia optativo"
    'ArgDesc(2) = vbnullstring
    'ArgDesc(3) = vbnullstring
    'DefinirAyudaUDF
    
    'FuncName = pD1S
    'FunDesc = "Calcula el  promedio de días a 1er Servicio"
    'Category = 14
    'ArgDesc(1) = "Número de lactancia optativo"
    'ArgDesc(2) = "Gestantes o Vacías"
    'ArgDesc(3) = vbnullstring
    'DefinirAyudaUDF
    
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If Application.UserName = "JPHC" Then
        Dim sRespuesta As String
        sRespuesta = _
          MsgBox( _
          "¿Enviar código a Control de Versiones?", _
          vbYesNo + vbDefaultButton2 + vbQuestion, _
          "JP's Automatización de Aplicaciones")
        If sRespuesta = vbYes Then
            Application.Run "SaveCodeMod" 'Mod ControlVersiones
        End If
    End If
End Sub


Private Sub DefinirAyudaUDF()
    On Error Resume Next
    Application.MacroOptions _
      Macro:=FuncName, _
      Description:=FunDesc, _
      Category:=14, _
      ArgumentDescriptions:=ArgDesc
    On Error GoTo 0
End Sub
