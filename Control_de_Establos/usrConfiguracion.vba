VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usrConfiguracion 
   Caption         =   "Control de Establos"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8040
   OleObjectBlob   =   "usrConfiguracion.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "usrConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Última modificación: 30-Oct-2017
Option Explicit
Dim bCambios, bProteccion As Boolean
Dim bCol(1 To 9) As Boolean
'Dim sCol(1 To 9) As String

Private Sub AA()
' Se ha efectuado un cambio en algún objeto del formulario
    bCambios = True
    Me.CommandButton2.Enabled = True
End Sub

Private Sub cmndBorrarInfo_Click()
    usrConfiguracion1.Show
    Me.Hide
End Sub

Private Sub boxCambiarContraseña_AfterUpdate()
' Habilitar el cambio de contraseñas
    Me.Label23 = _
      Me.boxCambiarContraseña
    Me.txtPW2User = _
      Me.boxCambiarContraseña
End Sub

Private Sub boxReqCapturista_Click()
    AA
End Sub

Private Sub boxReqContPeso_Click()
    AA
End Sub

Private Sub boxReqContReemplazos_Click()
    AA
End Sub

Private Sub boxReqFNacim_Click()
    AA
End Sub

Private Sub boxReqInventSemen_Click()
    AA
End Sub

Private Sub boxReqMachos_Click()
    AA
End Sub

Private Sub boxReqMadre_Click()
    AA
End Sub

Private Sub boxReqMagnet_Click()
    AA
End Sub

Private Sub boxReqPadre_Click()
    AA
End Sub

Private Sub boxReqPW_Click()
    AA
End Sub

Private Sub boxReqRaza_Click()
    AA
End Sub

Private Sub boxReqReemplazos_Click()
    AA
End Sub

Private Sub boxReqSemental_Click()
    AA
End Sub

Private Sub boxReqTecnico_Click()
    AA
End Sub

Private Sub boxreqVacBrucela_Click()
    AA
End Sub

Private Sub chkBox11_Click()
    AA
End Sub

Private Sub chkBox12_Click()
    AA
End Sub

Private Sub chkBox13_Click()
    AA
End Sub

Private Sub chkBox21_Click()
    AA
End Sub

Private Sub chkBox22_Click()
    AA
End Sub

Private Sub chkBox23_Click()
    AA
End Sub

Private Sub chkBox31_Click()
    AA
End Sub

Private Sub chkBox32_Click()
    AA
End Sub

Private Sub chkBox33_Click()
    AA
End Sub

Private Sub chkBox41_Click()
    AA
End Sub

Private Sub chkBox42_Click()
    AA
End Sub

Private Sub chkBox43_Click()
    AA
End Sub

Private Sub chkBox51_Click()
    AA
End Sub

Private Sub chkBox52_Click()
    AA
End Sub

Private Sub chkBox53_Click()
    AA
End Sub

Private Sub chkBox61_Click()
    AA
End Sub

Private Sub chkBox62_Click()
    AA
End Sub

Private Sub chkBox63_Click()
    AA
End Sub

Private Sub chkBox71_Click()
    AA
End Sub

Private Sub chkBox72_Click()
    AA
End Sub

Private Sub chkBox73_Click()
    AA
End Sub

Private Sub chkBox81_Click()
    AA
End Sub

Private Sub chkBox82_Click()
    AA
End Sub

Private Sub chkBox83_Click()
    AA
End Sub

Private Sub chkBox91_Click()
    AA
End Sub

Private Sub chkBox92_Click()
    AA
End Sub

Private Sub chkBox93_Click()
    AA
End Sub

Private Sub CommandButton1_Click()
' Cerrar Formulario
    Unload Me
End Sub

Private Sub CommandButton2_Click()
' Guardar cambios
    Dim sMsj As String
    If Me.boxCambiarContraseña Then
        ' Checar q contrseñas no estén en blanco
        If Me.txtPW1User = vbNullString And _
          Me.txtPW2User = vbNullString Then
            sMsj = MsgBox( _
              "Las Contraseñas están en Blanco", _
              vbInformation, _
              "Cambiar Contraseñas")
            Exit Sub
        End If
        ' Cambiar contraseñas
        Range("Desarrollador!B15") = _
          Me.txtPW1User.Text
    End If
    ' Checar existencia de contraseña
    If Me.txtPW1User = vbNullString Then
        sMsj = MsgBox( _
        "Ingresar Contraseña", _
        vbInformation, _
        "Efectuar Cambios")
        Exit Sub
    End If
    ' Checar coincidencias en contraseñas
    Select Case Me.txtPW1User.Text
        Case Is = "16910852"
            EfectuarCambios
        Case Is = Range("Desarrollador!B11").Text
            EfectuarCambios
        Case Is = Range("Desarrollador!B15").Text
            EfectuarCambios
        Case Else
            sMsj = MsgBox( _
              "La Contraseña NO coincide", _
              vbInformation, _
              "Efectuar Cambios")
    End Select
End Sub

Private Sub EfectuarCambios()
    On Error GoTo 120
100:
    With Me
        Range("Configuracion!C25") = _
          .boxReqCapturista
        Range("Configuracion!C35") = _
          .boxReqContPeso
        Range("Configuracion!C33") = _
          .boxReqContReemplazos
        Range("Configuracion!C22") = _
          .boxReqFNacim
        Range("Configuracion!C17") = _
          .boxReqInventSemen
        Range("Configuracion!C33") = _
          .boxReqMachos
        Range("Configuracion!C20") = _
          .boxReqMadre
        Range("Configuracion!C7") = _
          .boxReqMagnet
        Range("Configuracion!C19") = _
          .boxReqPadre
        Range("Configuracion!C27") = _
          .boxReqPW
        Range("Configuracion!C21") = _
          .boxReqRaza
        Range("Configuracion!C30") = _
          .boxReqReemplazos
        Range("Configuracion!C15") = _
          .boxReqSemental
        Range("Configuracion!C16") = _
          .boxReqTecnico
        'Me.boxreqVacBrucela
        Range("Configuracion!C34") = _
          Val(.txtdDestete)
        Range("Configuracion!C5") = _
          Val(.txtdDxGest)
        Range("Configuracion!C6") = _
          Val(.txtdWait)
        Range("Configuracion!C31") = _
          Val(.txtIdInicial)
        Range("Configuracion!C13") = _
          Val(.txtlLactancia)
        Range("Configuracion!C10") = _
          Val(.txtlPrep)
        Range("Configuracion!C11") = _
          Val(.txtlRParidas)
        Range("Configuracion!C9") = _
          Val(.txtlSeca)
        Range("Configuracion!C12") = _
          Val(.txtlVaqRParidas)
        Range("Configuracion!C24") = _
          Val(.txtProdMin)
        Range("Colaboradores!A2") = _
          .txtUsr1
        If bCol(1) Then _
          Range("Colaboradores!B2") = "1234"
        Range("Colaboradores!D2") = _
          .chkBox11
        Range("Colaboradores!E2") = _
          .chkBox12
        Range("Colaboradores!F2") = _
          .chkBox13
        Range("Colaboradores!A3") = _
          .txtUsr2
        If bCol(2) Then _
          Range("Colaboradores!B3") = "1234"
        Range("Colaboradores!D3") = _
          .chkBox21
        Range("Colaboradores!E3") = _
          .chkBox22
        Range("Colaboradores!F3") = _
          .chkBox23
        Range("Colaboradores!A4") = _
          .txtUsr3
        If bCol(3) Then _
          Range("Colaboradores!B4") = "1234"
        Range("Colaboradores!D4") = _
          .chkBox31
        Range("Colaboradores!E4") = _
          .chkBox32
        Range("Colaboradores!F4") = _
          .chkBox33
        Range("Colaboradores!A5") = _
          .txtUsr4
        If bCol(4) Then _
          Range("Colaboradores!B5") = "1234"
        Range("Colaboradores!D5") = _
          .chkBox41
        Range("Colaboradores!E5") = _
          .chkBox42
        Range("Colaboradores!F5") = _
          .chkBox43
        Range("Colaboradores!A6") = _
          .txtUsr5
        If bCol(5) Then _
          Range("Colaboradores!B6") = "1234"
        Range("Colaboradores!D6") = _
          .chkBox51
        Range("Colaboradores!E6") = _
          .chkBox52
        Range("Colaboradores!F6") = _
          .chkBox53
        Range("Colaboradores!A7") = _
          .txtUsr6
        If bCol(6) Then _
          Range("Colaboradores!B7") = "1234"
        Range("Colaboradores!D7") = _
          .chkBox61
        Range("Colaboradores!E7") = _
          .chkBox62
        Range("Colaboradores!F7") = _
          .chkBox63
        Range("Colaboradores!A8") = _
          .txtUsr7
        If bCol(7) Then _
          Range("Colaboradores!B8") = "1234"
        Range("Colaboradores!D8") = _
          .chkBox71
        Range("Colaboradores!E8") = _
          .chkBox72
        Range("Colaboradores!F8") = _
          .chkBox73
        Range("Colaboradores!A9") = _
          .txtUsr8
        If bCol(8) Then _
          Range("Colaboradores!B9") = "1234"
        Range("Colaboradores!D9") = _
          .chkBox81
        Range("Colaboradores!E9") = _
          .chkBox82
        Range("Colaboradores!F9") = _
          .chkBox83
        Range("Colaboradores!A10") = _
          .txtUsr9
        If bCol(9) Then _
          Range("Colaboradores!B10") = "1234"
        Range("Colaboradores!D10") = _
          .chkBox91
        Range("Colaboradores!E10") = _
          .chkBox92
        Range("Colaboradores!F10") = _
          .chkBox93
        ' Password
        .txtPW1User = vbNullString
        .txtPW2User = vbNullString
        .CommandButton2.Enabled = False
        ' Metas
        Range("Configuracion!B73") = Val(.TextBox9)
        Range("Configuracion!B74") = Val(.TextBox10)
        Range("Configuracion!B75") = Val(.TextBox11)
        Range("Configuracion!B77") = Val(.TextBox12)
        Range("Configuracion!B78") = Val(.TextBox13)
        Range("Configuracion!B79") = Val(.TextBox14)
        Range("Configuracion!B80") = Val(.TextBox15)
        Range("Configuracion!B81") = Val(.TextBox16)
        Range("Configuracion!B82") = Val(.TextBox17)
        Range("Configuracion!C83") = Val(.TextBox18)
        Range("Configuracion!B84") = Val(.TextBox19)
        Range("Configuracion!C85") = Val(.TextBox20)
        Range("Configuracion!B88") = Val(.TextBox21)
        Range("Configuracion!B89") = Val(.TextBox22)
        Range("Configuracion!B90") = Val(.TextBox23)
        Range("Configuracion!B92") = Val(.TextBox24)
        Range("Configuracion!B93") = Val(.TextBox25)
        Range("Configuracion!B94") = Val(.TextBox26)
        Range("Configuracion!B95") = Val(.TextBox27)
        Range("Configuracion!B96") = Val(.TextBox28)
    End With
    On Error GoTo 0
    bCambios = False
    Exit Sub
120:
    HabilitarHojas
    GoTo 100
End Sub

Private Sub CommandButton4_Click()
' Respaldar Información
    Application.Run "UnderConstruction"
End Sub

Private Sub HabilitarHojas()
    With Sheets("Desarrollador")
        .Unprotect Password:="0246813579"
    End With
    With Sheets("Configuracion")
        .Unprotect Password:="0246813579"
    End With
    bProteccion = True
End Sub

Private Sub TextBox8_Change()
    AA
End Sub

Private Sub TextBox9_Change()
    AA
End Sub

Private Sub TextBox10_Change()
    AA
End Sub

Private Sub TextBox11_Change()
    AA
End Sub

Private Sub TextBox12_Change()
    AA
End Sub

Private Sub TextBox13_Change()
    AA
End Sub

Private Sub TextBox14_Change()
    AA
End Sub

Private Sub TextBox15_Change()
    AA
End Sub

Private Sub TextBox16_Change()
    AA
End Sub

Private Sub TextBox17_Change()
    AA
End Sub

Private Sub TextBox18_Change()
    AA
End Sub

Private Sub TextBox19_Change()
    AA
End Sub

Private Sub TextBox20_Change()
    AA
End Sub

Private Sub TextBox21_Change()
    AA
End Sub

Private Sub TextBox22_Change()
    AA
End Sub

Private Sub TextBox23_Change()
    AA
End Sub

Private Sub TextBox24_Change()
    AA
End Sub

Private Sub TextBox25_Change()
    AA
End Sub

Private Sub TextBox26_Change()
    AA
End Sub

Private Sub TextBox27_Change()
    AA
End Sub

Private Sub TextBox28_Change()
    AA
End Sub

Private Sub txtdDestete_Change()
    AA
End Sub

Private Sub txtdDxGest_Change()
    AA
End Sub

Private Sub txtdWait_Change()
    AA
End Sub

Private Sub txtIdInicial_Change()
    AA
End Sub

Private Sub txtlLactancia_Change()
    AA
End Sub

Private Sub txtlPrep_Change()
    AA
End Sub

Private Sub txtlRParidas_Change()
    AA
End Sub

Private Sub txtlSeca_Change()
    AA
End Sub

Private Sub txtlVaqRParidas_Change()
    AA
End Sub

Private Sub txtProdMin_Change()
    AA
End Sub
Private Sub txtPW1User_Change()
    AA
End Sub

Private Sub txtPW2User_AfterUpdate()
' Comprobar coincidencias en las contraseñas
    Dim sMsj As String
    If Not Me.txtPW1User = _
      Me.txtPW2User Then _
      sMsj = MsgBox( _
      "Las Contraseñas no coinciden", _
      vbInformation, _
      "Cambiar Contraseñas")
    AA
End Sub

Private Sub txtUsr1_Change()
    bCol(1) = True
    AA
End Sub

Private Sub txtUsr2_Change()
    bCol(2) = True
    AA
End Sub

Private Sub txtUsr3_Change()
    bCol(3) = True
    AA
End Sub

Private Sub txtUsr4_Change()
    bCol(4) = True
    AA
End Sub

Private Sub txtUsr5_Change()
    bCol(5) = True
    AA
End Sub

Private Sub txtUsr6_Change()
    bCol(6) = True
    AA
End Sub

Private Sub txtUsr7_Change()
    bCol(7) = True
    AA
End Sub

Private Sub txtUsr8_Change()
    bCol(8) = True
    AA
End Sub

Private Sub txtUsr9_Change()
    bCol(9) = True
    AA
End Sub

Private Sub UserForm_Initialize()
    Dim i As Long
    Application.ScreenUpdating = _
      CBool(Range("Desarrollador!B6"))
    With Me
        .boxReqCapturista = _
          CBool(Range("Configuracion!C25"))
        .boxReqContPeso = _
          CBool(Range("Configuracion!C35"))
        .boxReqContReemplazos = _
          CBool(Range("Configuracion!C33"))
        .boxReqFNacim = _
          CBool(Range("Configuracion!C22"))
        .boxReqInventSemen = _
          CBool(Range("Configuracion!C17"))
        .boxReqMachos = _
          CBool(Range("Configuracion!C33"))
        .boxReqMadre = _
          CBool(Range("Configuracion!C20"))
        .boxReqMagnet = _
          CBool(Range("Configuracion!C7"))
        .boxReqPadre = _
          CBool(Range("Configuracion!C19"))
        .boxReqPW = _
          CBool(Range("Configuracion!C27"))
        .boxReqRaza = _
          CBool(Range("Configuracion!C21"))
        .boxReqReemplazos = _
          CBool(Range("Configuracion!C30"))
        .boxReqSemental = _
          CBool(Range("Configuracion!C15"))
        .boxReqTecnico = _
          CBool(Range("Configuracion!C16"))
        .boxreqVacBrucela = vbNullString
        .chkBox11 = CBool(Range("Colaboradores!D2"))
        .chkBox12 = CBool(Range("Colaboradores!E2"))
        .chkBox13 = CBool(Range("Colaboradores!F2"))
        .chkBox21 = CBool(Range("Colaboradores!D3"))
        .chkBox22 = CBool(Range("Colaboradores!E3"))
        .chkBox23 = CBool(Range("Colaboradores!F3"))
        .chkBox31 = CBool(Range("Colaboradores!D4"))
        .chkBox32 = CBool(Range("Colaboradores!E4"))
        .chkBox33 = CBool(Range("Colaboradores!F4"))
        .chkBox41 = CBool(Range("Colaboradores!D5"))
        .chkBox42 = CBool(Range("Colaboradores!E5"))
        .chkBox43 = CBool(Range("Colaboradores!F5"))
        .chkBox51 = CBool(Range("Colaboradores!D6"))
        .chkBox52 = CBool(Range("Colaboradores!E6"))
        .chkBox53 = CBool(Range("Colaboradores!F6"))
        .chkBox61 = CBool(Range("Colaboradores!D7"))
        .chkBox62 = CBool(Range("Colaboradores!E7"))
        .chkBox63 = CBool(Range("Colaboradores!F7"))
        .chkBox71 = CBool(Range("Colaboradores!D8"))
        .chkBox72 = CBool(Range("Colaboradores!E8"))
        .chkBox73 = CBool(Range("Colaboradores!F8"))
        .chkBox81 = CBool(Range("Colaboradores!D9"))
        .chkBox82 = CBool(Range("Colaboradores!E9"))
        .chkBox83 = CBool(Range("Colaboradores!F9"))
        .chkBox91 = CBool(Range("Colaboradores!D10"))
        .chkBox92 = CBool(Range("Colaboradores!E10"))
        .chkBox93 = CBool(Range("Colaboradores!F10"))
        .txtdDestete = _
          Range("Configuracion!C34")
        .txtdDxGest = _
          Range("Configuracion!C5")
        .txtdWait = _
          Range("Configuracion!C6")
        .txtIdInicial = _
          Range("Configuracion!C31")
        .txtlLactancia = _
          Range("Configuracion!C13")
        .txtlPrep = _
          Range("Configuracion!C10")
        .txtlRParidas = _
          Range("Configuracion!C11")
        .txtlSeca = _
          Range("Configuracion!C9")
        .txtlVaqRParidas = _
          Range("Configuracion!C12")
        .txtProdMin = _
          Range("Configuracion!C24")
        .txtPW1User = vbNullString
        .txtPW2User = vbNullString
        ' Mostrar u Ocultar Controles
        .txtPW2User.Visible = False
        .Label23.Visible = False
        .txtUsr1 = Range("Colaboradores!A2")
        .txtUsr2 = Range("Colaboradores!A3")
        .txtUsr3 = Range("Colaboradores!A4")
        .txtUsr4 = Range("Colaboradores!A5")
        .txtUsr5 = Range("Colaboradores!A6")
        .txtUsr6 = Range("Colaboradores!A7")
        .txtUsr7 = Range("Colaboradores!A8")
        .txtUsr8 = Range("Colaboradores!A9")
        .txtUsr9 = Range("Colaboradores!A10")
        .TextBox9 = Range("Configuracion!B73")
        .TextBox10 = Range("Configuracion!B74")
        .TextBox11 = Range("Configuracion!B75")
        .TextBox12 = Range("Configuracion!B77")
        .TextBox13 = Range("Configuracion!B78")
        .TextBox14 = Range("Configuracion!B79")
        .TextBox15 = Range("Configuracion!B80")
        .TextBox16 = Range("Configuracion!B81")
        .TextBox17 = Range("Configuracion!B82")
        .TextBox18 = Range("Configuracion!C83")
        .TextBox19 = Range("Configuracion!B84")
        .TextBox20 = Range("Configuracion!C85")
        .TextBox21 = Range("Configuracion!B88")
        .TextBox22 = Range("Configuracion!B89")
        .TextBox23 = Range("Configuracion!B90")
        .TextBox24 = Range("Configuracion!B92")
        .TextBox25 = Range("Configuracion!B93")
        .TextBox26 = Range("Configuracion!B94")
        .TextBox27 = Range("Configuracion!B95")
        .TextBox28 = Range("Configuracion!B96")
        .CommandButton2.Enabled = False
    End With
    For i = 1 To 9
        bCol(i) = False
        'sCol(i) = Worksheets("Colaboradores").Cells(i + 1, 1)
    Next i
    bCambios = False
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
