VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usrQuery 
   Caption         =   "UserForm1"
   ClientHeight    =   8145
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4080
   OleObjectBlob   =   "usrQuery.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "usrQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Ultima modificación: 31-10-2017


Private Sub CommandButton1_Click()
    Range("Configuracion!U3:AE4").Clear
    If Me.TextBox1 = "" Then GoTo 1001
    Range("Configuracion!U3:U4") = FormatNumber(TextBox1)
1001
    If Me.TextBox2 = "" Then GoTo 2002
    Range("Configuracion!V3:V4") = ">=" & FormatNumber(CDate(Me.TextBox2))
2002
    If Me.TextBox3 = "" Then GoTo 3003
    Range("Configuracion!W3:W4") = "<=" & FormatNumber(CDate(Me.TextBox3))
3003
    If Me.ComboBox2 = "" And Me.ComboBox3 = "" Then GoTo 4004
    If Me.ComboBox2 = "" Then Me.ComboBox2 = Me.ComboBox3
    If Me.ComboBox3 = "" Then Me.ComboBox3 = Me.ComboBox2
    Range("Configuracion!X3") = Me.ComboBox2
    Range("Configuracion!X4") = Me.ComboBox3
4004
    usrQuery
    CommandButton2_Click
End Sub

Private Sub CommandButton2_Click()
' Cerrar
    Application.ScreenUpdating = True
    Unload Me
End Sub

Private Sub textbox2_AfterUpdate()
    On Error Resume Next
    Me.TextBox2 = Format(CDate(Me.TextBox2), "d-mmm-yy")
    On Error GoTo 0
End Sub

Private Sub textbox3_AfterUpdate()
    On Error Resume Next
    Me.TextBox3 = Format(CDate(Me.TextBox3), "d-mmm-yy")
    On Error GoTo 0
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

Private Sub UserForm_Initialize()
    Application.ScreenUpdating = CBool(Range("Desarrollador!B6"))
    Me.Caption = Range("Configuracion!C3")
    With Me
        .CheckBox1 = True
        '.CheckBox1.Enabled = False
        .CheckBox2 = True
        '.CheckBox2.Enabled = False
        .CheckBox3 = True
        '.CheckBox3.Enabled = False
        .CheckBox4 = False
        '.CheckBox4.Enabled = False
        .CheckBox5 = False
        '.CheckBox5.Enabled = False
        .CheckBox6 = False
        '.CheckBox6.Enabled = False
        .CheckBox7 = False
        '.CheckBox7.Enabled = False
        .CheckBox8 = False
        '.CheckBox8.Enabled = False
    End With
    With Me.ComboBox2
        .AddItem ""
        .AddItem "Serv"
        .AddItem "Calor"
        .AddItem "Prod"
        .AddItem "Movimiento"
        .AddItem "Enfermedad"
        .AddItem "Revisión"
        .AddItem "DxGst"
        .AddItem "Seca"
        .AddItem "Nota"
        .AddItem "Parto"
        .AddItem "Aborto"
        .AddItem "Imantación"
        .AddItem "Otro"
        .AddItem "Baja"
        .AddItem "Pesaje"
        .AddItem "Destete"
        .AddItem "Alta"
    End With
    With Me.ComboBox3
        .AddItem ""
        .AddItem "Serv"
        .AddItem "Calor"
        .AddItem "Prod"
        .AddItem "Movimiento"
        .AddItem "Enfermedad"
        .AddItem "Revisión"
        .AddItem "DxGst"
        .AddItem "Seca"
        .AddItem "Nota"
        .AddItem "Parto"
        .AddItem "Aborto"
        .AddItem "Imantación"
        .AddItem "Otro"
        .AddItem "Baja"
        .AddItem "Pesaje"
        .AddItem "Destete"
        .AddItem "Alta"
    End With
End Sub

Private Sub usrQuery()
    Sheets("Eventos").Range("Tabla6[#All]").AdvancedFilter Action:=xlFilterCopy, _
      CriteriaRange:=Range("Query"), CopyToRange:=Sheets("Query").Range("A1:F1"), _
      Unique:=False
    With Sheets("Query")
        .Visible = True
        .Activate
    End With
End Sub


