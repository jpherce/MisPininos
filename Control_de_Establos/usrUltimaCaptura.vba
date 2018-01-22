VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usrUltimaCaptura 
   Caption         =   "Control de Establos"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4305
   OleObjectBlob   =   "usrUltimaCaptura.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "usrUltimaCaptura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Ultima modificación: 7-Ene-16
Option Explicit

Private Sub cmboIdArete_AfterUpdate()
' Busca y actualiza
    Dim iArete As Long
    If Me.cmboIdArete = vbNullString Then Exit Sub
    iArete = Me.cmboIdArete
    'iArete = CDbl(Me.cmboIdArete)
    If BuscarUltimaOcurrencia(iArete, Range("Tabla6[Arete]"), 1) = _
      "No existe" Then
        With Me
            .txtFecha = "No existe"
            .txtEvento = vbNullString
            .txtObservaciones = vbNullString
            .txtResponsable = vbNullString
            .txtCapturista = vbNullString
            .txtFechaCaptura = vbNullString
            .cmboIdArete.SetFocus
        End With
        Exit Sub
    End If
    With Me
        .txtFecha = _
          Format( _
          BuscarUltimaOcurrencia(iArete, Range("Tabla6[Arete]"), 2), _
          "d-mmm-yy")
        .txtEvento = _
          BuscarUltimaOcurrencia(iArete, Range("Tabla6[Arete]"), 3)
        .txtObservaciones = _
          BuscarUltimaOcurrencia(iArete, Range("Tabla6[Arete]"), 4)
        .txtResponsable = _
          BuscarUltimaOcurrencia(iArete, Range("Tabla6[Arete]"), 5)
        .txtCapturista = _
          BuscarUltimaOcurrencia(iArete, Range("Tabla6[Arete]"), 6)
        .txtFechaCaptura = _
          Format( _
          BuscarUltimaOcurrencia(iArete, Range("Tabla6[Arete]"), 7), _
          "d-mmm-yy") _
          & " " & _
          Format( _
          BuscarUltimaOcurrencia(iArete, Range("Tabla6[Arete]"), 8), _
          "hh:mm")
    End With
End Sub

Private Sub CommandButton2_Click()
' Borrar
    With Me
        .cmboIdArete = vbNullString
        .txtFecha = vbNullString
        .txtEvento = vbNullString
        .txtObservaciones = vbNullString
        .txtResponsable = vbNullString
        .txtCapturista = vbNullString
        .txtFechaCaptura = vbNullString
        .cmboIdArete.SetFocus
    End With
End Sub

Private Sub CommandButton3_Click()
' Cerrar
    Application.ScreenUpdating = True
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim mRenglones As Double
    Dim rCelda As Range
    Application.ScreenUpdating = CBool(Range("Desarrollador!B6"))
    Me.Label9.Caption = "Ver. 15.1004" ' Indicar Versión del módulo
    ' Quitar los filtros de Tabla6
    With Range("Tabla6")
        .AutoFilter
        .AutoFilter
    End With
    ' Poblar el ComboBox con aretes de Hato y Reemplazos
    For Each rCelda In Range("Tabla1[Arete]")
        With Me.cmboIdArete
            .AddItem rCelda.Offset(0, 0)
        End With
    Next rCelda
    For Each rCelda In Range("Tabla2[Arete]")
        With Me.cmboIdArete
            .AddItem rCelda.Offset(0, 0)
        End With
    Next rCelda
End Sub

Private Sub UserForm_QueryClose(mCancel As Integer, _
                                mCloseMode As Integer)
' Deshabilitar el Botón X para cerrar el cuadro de diálogo
    If mCloseMode <> vbFormCode Then
        MsgBox _
          "Utiliza 'Cerrar' para salir de este formulario.", _
          vbExclamation, _
          "JP's Automatización de Aplicaciones"
        mCancel = True
    End If
End Sub


