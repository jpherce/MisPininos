VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usrKardex 
   Caption         =   "Control de Establos"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8040
   OleObjectBlob   =   "usrKardex.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "usrKardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Ultima modificación: 21.11.17
' Mostrar DCC
' Mod formas de presentción de EM305d, Proy305d, ValRel
' Mod presentación de valores en LactAnterirores
Option Explicit
Dim iNoLact As Long
Dim ws As Worksheet
Dim rCelda As Range
Dim mArray() As String
Dim sN, sD As Long
Dim mArete As Variant

Private Sub BorrarDatos()
' Borrar Datos del formulario
    With Me
        .TextBox7 = vbNullString
        .TextBox8 = vbNullString
        .TextBox9 = vbNullString
        .TextBox10 = vbNullString
        .txtStatus = vbNullString
        .txtCorral = vbNullString
        .txtCria = vbNullString
        .txtD1Ser = vbNullString
        .txtDEL = vbNullString
        .txtDEL2 = vbNullString
        .txtDiasAb = vbNullString
        .txtDiasSeca = vbNullString
        .txtEdad = vbNullString
        .txtEdad1Parto = vbNullString
        .txtEdad1Serv = vbNullString
        .txtFNacim = vbNullString
        .txtFParir = vbNullString
        .txtFParto = vbNullString
        .txtFSecar = vbNullString
        .txtFServicio = vbNullString
        .txtFVacBrucela = vbNullString
        .txtParto = vbNullString
        .txtProy305d = vbNullString
        .txtProdActual = vbNullString
        .txtProdAcum = vbNullString
        .txtRaza = vbNullString
        .txtServicio = vbNullString
        .txtSexo = vbNullString
        .txtTecnico = vbNullString
        .txtPartoTipo = vbNullString
        .txtToro = vbNullString
        .txtValorRelativo = vbNullString
        .txtEM305d = vbNullString
        .txtEdad = vbNullString
        .txtFNacim = vbNullString
        .txtSexo = vbNullString
        .txtPadre = vbNullString
        .txtMadre = vbNullString
        .txtRaza = vbNullString
        .txtFVacBrucela = vbNullString
        .txtFIman = vbNullString
    End With
End Sub

Private Sub CommandButton1_Click()
' Cerrar formulario
    Application.ScreenUpdating = True
    Unload Me
End Sub

Private Sub CommandButton2_Click()
' Actualizar información de Eventos
    With Me.listEventos
        .Clear
        .ColumnCount = 4
        .ColumnWidths = "50;40;110;30"
        '.ColumnHeads = True
    End With
    ' Checar si existe información
    If WorksheetFunction.CountIf(Range("Tabla6[Arete]"), _
      mArete) > 0 Then
        For Each rCelda In Range("Tabla6[Arete]")
            If rCelda.Offset(0, 0) = mArete Then
                Select Case rCelda.Offset(0, 2)
                    Case "Parto", "Aborto"
                        PoblarLstEventos
                    Case "Serv", "Calor" 'Calores y Servicios
                        If CBool(Me.cboxServicios) Then _
                          PoblarLstEventos
                    Case "Prod", "Seca" 'Producciones
                        If CBool(Me.cboxProd) Then PoblarLstEventos
                    Case "Mov", "Seca" 'Movimientos
                        If CBool(Me.cboxMov) Then PoblarLstEventos
                    Case "Rev", "DxGst" 'Revisiones Médicas
                        If CBool(Me.cboxRevisiones) Then _
                          PoblarLstEventos
                    Case Else 'Otros
                        If CBool(Me.cboxOtros) Then PoblarLstEventos
                End Select
            End If
        Next rCelda
    End If
End Sub

Private Sub MuestraDatos()
' Mostrar Datos renglón Actual
    Dim mProm305d, mValRelativo As Double
    BorrarDatos
    PoblarLstLactAnteriores 'LactanciasAnteriores
    PoblarLstBenchmarking 'Comparativos
    CommandButton2_Click 'Kardex
    
    On Error Resume Next
    ' seleccionar caso
    Select Case ws.Name
        Case Is = "Hato"
            Select Case WorksheetFunction.VLookup(mArete, _
              Range("Tabla1"), 11, False)
                Case Is = "P"
                    Me.txtStatus = "Gestante"
                Case Is = "O"
                    Me.txtStatus = "Vacía"
                Case Is = vbNullString
                    If Date - CDate(WorksheetFunction.VLookup( _
                      mArete, Range("Tabla1"), 6, False)) > _
                      Range("Configuracion!C5") And Not _
                      WorksheetFunction.VLookup(mArete, _
                      Range("Tabla1"), 7, False) = vbNullString Then _
                        Me.txtStatus = "Servida"
                    If Date - CDate(WorksheetFunction.VLookup( _
                    mArete, Range("Tabla1"), 6, False)) > _
                      Range("Configuracion!C6") And _
                      WorksheetFunction.VLookup(mArete, _
                      Range("Tabla1"), 7, False) = vbNullString Then _
                      Me.txtStatus = "Sin servir"
                Case Else
                    Me.txtStatus = vbNullString
            End Select
            If Not WorksheetFunction.VLookup(mArete, _
              Range("Tabla1"), 8, False) = vbNullString Then _
              Me.txtDiasCarga = Date - CDate(WorksheetFunction. _
              VLookup(mArete, Range("Tabla1"), 8, False)) & "d"
            Me.txtCorral = WorksheetFunction.VLookup(mArete, _
              Range("Tabla1"), 2, False)
            Me.txtCria = WorksheetFunction.VLookup(mArete, _
              Range("Tabla1"), 2, False)
            Me.txtD1Ser = WorksheetFunction.VLookup(mArete, _
              Range("Tabla15"), 2, False)
            If WorksheetFunction.VLookup(mArete, _
              Range("Tabla15"), 16, False) = vbNullString Then
                    Me.txtDEL = _
                      Date - CDate(WorksheetFunction. _
                      VLookup(mArete, Range("Tabla1"), 6, _
                      False))
                    Me.txtDEL2 = _
                      Date - CDate(WorksheetFunction.VLookup( _
                      mArete, Range("Tabla1"), 6, False))
                Else
                    Me.txtDEL = _
                      CDate(WorksheetFunction. _
                      VLookup(mArete, Range("Tabla15"), _
                      16, False)) - CDate(WorksheetFunction. _
                      VLookup(mArete, Range("Tabla1"), 6, _
                      False))
                    Me.txtDEL2 = _
                      CDate(WorksheetFunction. _
                      VLookup(mArete, Range("Tabla15"), 16, _
                      False)) - CDate(WorksheetFunction. _
                      VLookup(mArete, Range("Tabla1"), 6, _
                      False))
            End If
            Me.txtDiasAb = WorksheetFunction. _
              VLookup(mArete, Range("Tabla15"), 3, False)
            If WorksheetFunction.VLookup(mArete, _
              Range("Tabla15"), 16, False) <> vbNullString Then _
              Me.txtDiasSeca = Date - CDate(WorksheetFunction. _
              VLookup(mArete, Range("Tabla15"), 16, False))
            Me.txtFParir = _
              Format(WorksheetFunction.VLookup(mArete, _
              Range("Tabla1"), 13, False), "dd-mmm-yy")
            Me.txtFParto = _
              Format(WorksheetFunction.VLookup(mArete, _
              Range("Tabla1"), 6, False), "dd-mmm-yy")
            Me.txtFSecar = _
              Format(WorksheetFunction.VLookup(mArete, _
              Range("Tabla1"), 12, False), "dd-mmm-yy")
            Me.txtFServicio = _
              Format(WorksheetFunction.VLookup(mArete, _
              Range("Tabla1"), 8, False), "dd-mmm-yy")
            Me.txtParto = WorksheetFunction.VLookup(mArete, _
              Range("Tabla1"), 5, False)
            Me.txtPartoTipo = _
              WorksheetFunction. _
              VLookup(mArete, Range("Tabla15"), 18, False)
            Me.txtProdActual = _
              Format(WorksheetFunction.VLookup(mArete, _
              Range("Tabla1"), 3, False), "#.0")
            Me.txtProdAcum = Format(WorksheetFunction. _
              VLookup(mArete, Range("Tabla15"), 14, False), _
              "#,#")
            Me.txtServicio = WorksheetFunction. _
              VLookup(mArete, Range("Tabla1"), 7, False)
            Me.txtTecnico = WorksheetFunction. _
              VLookup(mArete, Range("Tabla1"), 10, False)
            Me.txtToro = WorksheetFunction.VLookup(mArete, _
              Range("Tabla1"), 9, False)
            'Me.txtValorRelativo = Format((WorksheetFunction. _
              VLookup(mArete, Range("Tabla15"), 14, False) / _
              Date - CDate(WorksheetFunction.VLookup(mArete, _
              Range("Tabla1"), 6, False))), "#,#")
            
                With Me
                    ' Proyección a 305d
                    If CBool(Range("Configuracion!B65")) Then
                        .lblProy305d.Visible = True
                        .txtProy305d.Visible = True
                        .txtProy305d = Format(WorksheetFunction. _
                          VLookup(mArete, Range("Tabla15"), 15, False), _
                        "#,#")
                    End If
                    ' Valor Relativo
                    If CBool(Range("Configuracion!B66")) Then
                        .lblValorRelativo.Visible = True
                        .txtValorRelativo.Visible = True
                        .txtValorRelativo = Format((WorksheetFunction. _
                          VLookup(mArete, Range("Tabla1"), 17, False)), "#,#")
                    End If
                    ' Equivalente Madurez
                    If CBool(Range("Configuracion!B67")) Then
                        .lblEM305d.Visible = True
                        .txtEM305d.Visible = True
                        Select Case Val(.txtParto)
                            Case Is = 1
                                .txtEM305d = Format(WorksheetFunction. _
                                  VLookup(mArete, Range("Tabla1"), 16, _
                                  False) * Range("Configuracion!L3"), "#,#")
                            Case Is = 2
                                .txtEM305d = Format(WorksheetFunction. _
                                  VLookup(mArete, Range("Tabla1"), 16, _
                                  False) * Range("Configuracion!L4"), "#,#")
                            Case Is >= 3
                                .txtEM305d = Format(WorksheetFunction. _
                                  VLookup(mArete, Range("Tabla1"), 16, _
                                  False) * Range("Configuracion!L5"), "#,#")
                        End Select
                    End If
                End With
            If Not WorksheetFunction.VLookup(mArete, _
              Range("Tabla15"), 22, False) = vbNullString Then Me.txtEdad = _
              Int((Date - CDate(WorksheetFunction. _
              VLookup(mArete, _
              Range("Tabla15"), 22, False))) / 365) & " años" 'Edad
            Me.txtFNacim = _
              Format(WorksheetFunction.VLookup(mArete, _
              Range("Tabla15"), 22, False), "dd-mmm-yy") 'F.Nacim
            Me.txtSexo = "H"
            Me.txtPadre = WorksheetFunction.VLookup(mArete, _
              Range("Tabla15"), 19, False)
            Me.txtMadre = WorksheetFunction.VLookup(mArete, _
              Range("Tabla15"), 20, False)
            Me.txtRaza = WorksheetFunction.VLookup(mArete, _
              Range("Tabla15"), 21, False)
            Me.txtFVacBrucela = vbNullString
            Me.txtFIman = _
              Format(WorksheetFunction.VLookup(mArete, _
              Range("Tabla1"), 16, False), "dd-mmm-yy") 'F.Imantación
            Me.txtProdAcumVitalica = _
              Format(Application.WorksheetFunction _
              .SumIfs(Range("Tabla4[ProdAcum]"), _
              Range("Tabla4[Arete]"), Me.txtArete), "#,#")
            Me.txtNumLact = Application.WorksheetFunction. _
              CountIf(Range("Tabla4[Arete]"), Me.txtArete)
            Me.txtProdPromVitalicia = Format(Int(Application. _
              WorksheetFunction.AverageIfs(Range("Tabla4[ProdAcum]"), _
              Range("Tabla4[Arete]"), Me.txtArete)), "#,#")
            Me.txtDiasProduccion = Application.WorksheetFunction. _
              SumIfs(Range("Tabla4[DiasLactancia]"), _
              Range("Tabla4[Arete]"), Me.txtArete)
            Me.txtDiasSecaVitalicia = Application.WorksheetFunction _
              .SumIfs(Range("Tabla4[DíasSeca]"), _
              Range("Tabla4[Arete]"), Me.txtArete)
            Me.txtPromServicios = Format(Application.WorksheetFunction _
              .AverageIfs(Range("Tabla4[Servicio]"), _
              Range("Tabla4[Arete]"), Me.txtArete), "#.0")
        Case Is = "Reemplazos"
            MostrarReemplazos
        Case Is = "InfoVitalicia"
            Me.txtStatus = "BAJA"
        Case Else
    End Select 'ws.Name
    MuestraInfoVitalicia
    On Error GoTo 0
End Sub

Private Sub MuestraInfoVitalicia()
' Mostrar Datos Vitalicios
    sD = 365
    On Error Resume Next
    'If Not WorksheetFunction.VLookup(mArete, _
      Range("Tabla8"), 3, False) = vbNullString Then _
      Me.txtEdad = Format(Int((Date - _
      WorksheetFunction.VLookup(mArete, _
      Range("Tabla8"), 3, False)) / 30.4), "###")
    If Not WorksheetFunction.VLookup(mArete, _
      Range("Tabla8"), 3, False) = vbNullString Then
        sN = Date - WorksheetFunction.VLookup(mArete, _
        Range("Tabla8"), 3, False)
        Me.txtEdad = Format(Int(sN / sD), "0") & "-" & _
        Format(Int((sN - sD * Int(sN / sD)) / 30.4), "00")
      End If
    'On Error GoTo 0
    Me.txtFNacim = _
      Format(WorksheetFunction.VLookup(mArete, _
      Range("Tabla8"), 3, False), "dd-mmm-yy")
    Me.txtRaza = _
      UCase(WorksheetFunction.VLookup(mArete, _
      Range("Tabla8"), 6, False))
    Me.txtPadre = _
      UCase(WorksheetFunction.VLookup(mArete, _
      Range("Tabla8"), 4, False))
    Me.txtMadre = _
      WorksheetFunction.VLookup(mArete, _
      Range("Tabla8"), 5, False)
    'Me.txtEdad1Serv = _
      WorksheetFunction.VLookup(mArete, _
      Range("Tabla8"), 10, False)
    If Not WorksheetFunction.VLookup(mArete, _
      Range("Tabla8"), 10, False) = vbNullString Then
        sN = WorksheetFunction.VLookup(mArete, _
        Range("Tabla8"), 10, False)
        Me.txtEdad1Serv = Format(Int(sN / sD), "0") & "-" & _
        Format(Int((sN - sD * Int(sN / sD)) / 30.4), "00")
    End If
    
    'Me.txtEdad1Parto = _
      WorksheetFunction.VLookup(mArete, _
      Range("Tabla8"), 11, False) & "m"
    If Not WorksheetFunction.VLookup(mArete, _
      Range("Tabla8"), 11, False) = vbNullString Then
        sN = WorksheetFunction.VLookup(mArete, _
        Range("Tabla8"), 11, False)
        Me.txtEdad1Parto = Format(Int(sN / sD), "0") & "-" & _
        Format(Int((sN - sD * Int(sN / sD)) / 30.4), "00")
    End If
    On Error GoTo 0
    
    Me.txtFVacBrucela = _
      Format(WorksheetFunction.VLookup(mArete, _
      Range("Tabla8"), 8, False), "dd-mmm-yy")
    Me.txtFIman = _
      Format(WorksheetFunction.VLookup(mArete, _
      Range("Tabla8"), 9, False), "dd-mmm-yy")
    Me.TextBox9 = WorksheetFunction.VLookup(mArete, _
      Range("Tabla8"), 2, False)
    Me.TextBox10 = WorksheetFunction.VLookup(mArete, _
      Range("Tabla8"), 7, False)
    Me.TextBox7 = Format(WorksheetFunction.VLookup(mArete, _
      Range("Tabla8"), 14, False), "dd-mmm-yy")
    Me.TextBox8 = WorksheetFunction.VLookup(mArete, _
      Range("Tabla8"), 15, False)
    Me.txtArete.SetFocus
End Sub

Private Sub MostrarReemplazos()
' Mostrar Datos de Hoja Reemplazos
    Me.txtCorral = _
      WorksheetFunction.VLookup(mArete, _
      Range("Tabla2"), 2, False)
    If ws.Name = "Reemplazos" Then
        Select Case WorksheetFunction.VLookup(mArete, _
          Range("Tabla2"), 10, False)
            Case Is = "P"
                Me.txtStatus = "Gestante"
            Case Is = "O"
                Me.txtStatus = "Vacía"
            Case Is = vbNullString
                If Date - CDate(WorksheetFunction.VLookup( _
                  mArete, Range("Tabla2"), 5, False)) > _
                  Range("Configuracion!C14") * 30.4 And Not _
                  WorksheetFunction.VLookup(mArete, _
                  Range("Tabla2"), 7, False) = vbNullString Then _
                  Me.txtStatus = "Servida"
                If Date - CDate(WorksheetFunction.VLookup( _
                  mArete, Range("Tabla2"), 5, False)) > _
                  Range("Configuracion!C14") * 30.4 And _
                  WorksheetFunction.VLookup(mArete, _
                  Range("Tabla2"), 6, False) = vbNullString Then _
                  Me.txtStatus = "Sin servir"
                If Date - CDate(WorksheetFunction.VLookup( _
                  mArete, Range("Tabla2"), 5, False)) < _
                  Range("Configuracion!C14") * 30.4 Then _
                  Me.txtStatus = vbNullString
            Case Else
                Me.txtStatus = vbNullString
        End Select
    End If
End Sub

Private Sub PoblarLstEventos()
' Adiciona info
If ws.Name = "Hato" Then
  If CBool(Me.optLacActual) = True And _
  CDate(rCelda.Offset(0, 1)) < _
  CDate(WorksheetFunction.VLookup(mArete, _
  Range("Tabla1"), 6, False)) Then Exit Sub
End If
    With Me.listEventos
        If rCelda.Offset(0, 2) = "Parto" Or rCelda.Offset(0, 2) = _
          "Aborto" Then
                .AddItem "*" & Format(rCelda.Offset(0, 1), _
                "d-mmm-yy")
            Else
                .AddItem "  " & Format(rCelda.Offset(0, 1), _
                "d-mmm-yy")
        End If
        .List(.ListCount - 1, 1) = rCelda.Offset(0, 2)
        Select Case rCelda.Offset(0, 2)
            Case Is = "Serv", "Calor" 'Servicios y Calores
                .List(.ListCount - 1, 2) = rCelda.Offset(0, 3)
                .List(.ListCount - 1, 3) = rCelda.Offset(0, 4) _
                  & " " & Mid(rCelda.Offset(0, 8), 4, 3) & "d" '**
            Case Is = "Prod" 'Producciones
                .List(.ListCount - 1, 2) = _
                  Format(rCelda.Offset(0, 3), "#.0") '**
                '.List(.ListCount - 1, 3) = rCelda.Offset(0, 4)
                If Not Mid(rCelda.Offset(0, 8), 5, 3) = "000" Then _
                  .List(.ListCount - 1, 3) = _
                  Format(Mid(rCelda.Offset(0, 8), 5, 3), "000") & "%"
            Case Is = "Mov" 'Movimientos
                .List(.ListCount - 1, 2) = _
                  "Corral-> " & rCelda.Offset(0, 3) '**
                .List(.ListCount - 1, 3) = Val(rCelda.Offset(0, 8)) & "d"
            Case Else 'Otros
                .List(.ListCount - 1, 2) = rCelda.Offset(0, 3)
                .List(.ListCount - 1, 3) = rCelda.Offset(0, 4)
        End Select
    End With
End Sub

Private Sub PoblarLstBenchmarking()
    Dim i, j As Long
    'Dias1Serv, DiasAb, ProdAcum, Proy305d
    Dim mD1S, mDA, mPA, mP305 As Double
    Dim cD1S, cDA, cPA, cP305 As Long
    'ProdAcum, Servicios, DiasLact, DiasSecos, Dias1Serv, DiasAb
    Dim xPA, xS, xDL, xDS, xD1S, xDAb As Double
    Dim cPAc, cS, cDL, cDS, cD1Se, cDAb As Long
    'LactComputadas
    Dim cLC As Long
    Dim rCelda1 As Range
    Dim mArrayB(5, 28)
    ' Identifica todos los animales con mismo parto
    If ws.Name = "Hato" Then _
      iNoLact = WorksheetFunction.VLookup(mArete, _
        Range("Tabla1"), 5, False) Else iNoLact = 0
      'iNoLact = ActiveCell.Offset(0, 4) Else _
      iNoLact = 0
    'iNoLact = ActiveCell.Offset(0, 4) Else
    
    If iNoLact = 0 Then Exit Sub ' ¡No hay contra quién comparar!
    ReDim mArrayC(iNoLact)
    For Each rCelda In Range("Tabla1[Arete]")
    'If rCelda = 502 Then MsgBox "HastaAQui"
        If rCelda.Offset(0, 4) = iNoLact Then
            On Error Resume Next
            If Not IsEmpty(WorksheetFunction.VLookup( _
              rCelda.Offset(0, 0), Range("Tabla15"), 2, False)) _
              Then cD1S = cD1S + 1
            mD1S = mD1S + WorksheetFunction.VLookup( _
              rCelda.Offset(0, 0), Range("Tabla15"), 2, False)
            If Not IsEmpty(WorksheetFunction.VLookup( _
              rCelda.Offset(0, 0), Range("Tabla15"), 3, False)) _
              Then cDA = cDA + 1
            mDA = mDA + WorksheetFunction.VLookup( _
              rCelda.Offset(0, 0), Range("Tabla15"), 3, False)
            If Not IsEmpty(WorksheetFunction.VLookup( _
              rCelda.Offset(0, 0), Range("Tabla15"), 14, False)) _
              Then cPA = cPA + 1
            mPA = mPA + WorksheetFunction.VLookup( _
              rCelda.Offset(0, 0), Range("Tabla15"), 14, False)
            If Not IsEmpty(WorksheetFunction.VLookup( _
              rCelda.Offset(0, 0), Range("Tabla15"), 15, False)) _
              Then cP305 = cP305 + 1
            mP305 = mP305 + WorksheetFunction.VLookup( _
              rCelda.Offset(0, 0), Range("Tabla15"), 15, False)
            On Error GoTo 0
            For Each rCelda1 In Range("Tabla4[Arete]")
                If rCelda1.Offset(0, 0) = rCelda.Offset(0, 0) Then
                    If Not IsEmpty(WorksheetFunction.VLookup( _
                      rCelda1.Offset(0, 0), Range("Tabla4"), _
                      13, False)) Then
                        cPAc = cPAc + 1
                        xPA = xPA + WorksheetFunction.VLookup( _
                          rCelda1.Offset(0, 0), Range("Tabla4"), _
                          13, False)
                    End If
                    If Not IsEmpty(WorksheetFunction.VLookup( _
                      rCelda1.Offset(0, 0), Range("Tabla4"), _
                      3, False)) Then
                        cS = cS + 1
                        xS = xS + WorksheetFunction.VLookup( _
                          rCelda1.Offset(0, 0), Range("Tabla4"), _
                          3, False) '***OJO
                    End If
                    If Not IsEmpty(WorksheetFunction.VLookup( _
                      rCelda1.Offset(0, 0), Range("Tabla4"), _
                      11, False)) Then
                        cDL = cDL + 1
                        xDL = xDL + WorksheetFunction.VLookup( _
                          rCelda1.Offset(0, 0), Range("Tabla4"), _
                          11, False)
                    End If
                    If Not IsEmpty(WorksheetFunction.VLookup( _
                      rCelda1.Offset(0, 0), Range("Tabla4"), _
                      12, False)) Then
                        cDS = cDS + 1
                        xDS = xDS + WorksheetFunction.VLookup( _
                          rCelda1.Offset(0, 0), Range("Tabla4"), _
                          12, False)
                    End If
                    If Not IsEmpty(WorksheetFunction.VLookup( _
                      rCelda1.Offset(0, 0), Range("Tabla4"), _
                      15, False)) Then
                        cD1Se = cD1Se + 1
                        xD1S = xD1S + WorksheetFunction.VLookup( _
                          rCelda1.Offset(0, 0), Range("Tabla4"), _
                          15, False)
                    End If
                    If Not IsEmpty(WorksheetFunction.VLookup( _
                      rCelda1.Offset(0, 0), Range("Tabla4"), _
                      16, False)) Then
                        cDAb = cDAb + 1
                        xDAb = xDAb + WorksheetFunction.VLookup( _
                          rCelda1.Offset(0, 0), Range("Tabla4"), _
                          16, False)
                    End If
                End If
            Next rCelda1
        End If
    Next rCelda
    On Error Resume Next
'2431:
    ' Promediar
    mD1S = mD1S / cD1S
    mDA = mDA / cDA
    mPA = mPA / cPA
    mP305 = mP305 / cP305
    xPA = xPA / cPA
    xS = xS / cS
    xDL = xDL / cDL
    xDS = xDS / cDS
    xD1S = xD1S / cD1Se
    xDAb = xDAb / cDAb

    mArrayB(1, 1) = "No. Parto"
    mArrayB(2, 1) = iNoLact
    mArrayB(3, 1) = iNoLact
    mArrayB(1, 2) = vbNullString
    mArrayB(1, 3) = "LACT. ACTUAL"
    mArrayB(1, 4) = "Prod. Diaria"
    mArrayB(2, 4) = Format(WorksheetFunction. _
      VLookup(mArete, Range("Tabla1"), 3, False), "#.0")
    mArrayB(3, 4) = Format(Application.WorksheetFunction. _
      AverageIfs(Range("Tabla1[Prod.]"), Range("Tabla1[Parto]"), _
      iNoLact), "#.0")
    mArrayB(4, 4) = Format(Application.WorksheetFunction. _
      Average(Range("Tabla1[Prod.]")), "#.0")
    mArrayB(1, 5) = "Días en Leche"
    mArrayB(2, 5) = WorksheetFunction.VLookup(mArete, _
      Range("Tabla1"), 4, False)
    mArrayB(3, 5) = Format(Application.WorksheetFunction. _
      AverageIfs(Range("Tabla1[DEL]"), Range("Tabla1[Parto]"), _
      iNoLact), "#")
    mArrayB(4, 5) = Format(Application.WorksheetFunction. _
      Average(Range("Tabla1[DEL]")), "#")
    mArrayB(1, 6) = "Prod. Acumulada"
    mArrayB(2, 6) = Format(Int(WorksheetFunction.VLookup( _
      mArete, Range("Tabla15"), 14, False) / 10) * 10, "#,#")
    mArrayB(3, 6) = Format(Int(mPA / 10) * 10, "#,#0")
    mArrayB(4, 6) = Format(Int(Application.WorksheetFunction. _
      Average(Range("Tabla15[ProdAcum]")) / 10) * 10, "#,#0")
    mArrayB(1, 7) = "Proy. a 305 Días"
    mArrayB(2, 7) = Format(Int(WorksheetFunction.VLookup( _
      mArete, Range("Tabla15"), 15, False) / 10) * 10, "#,#")
    mArrayB(3, 7) = Format(Int(mP305 / 10) * 10, "#,#")
    mArrayB(4, 7) = Format(Int(Application.WorksheetFunction. _
      Average(Range("Tabla15[Proy305d]")) / 10) * 10, "#,#0")
    mArrayB(1, 8) = "Valor Relativo"
    If WorksheetFunction.VLookup(mArete, _
      Range("Tabla15"), 15, False) = vbNullString Then
            mArrayB(3, 8) = "N.D."
            mArrayB(4, 8) = "N.D."
        Else
            mArrayB(3, 8) = Format((mArrayB(2, 7) / mArrayB(3, 7)) _
              * 100, "#")
            mArrayB(4, 8) = Format((mArrayB(2, 7) / mArrayB(4, 7)) _
              * 100, "#")
    End If
    mArrayB(1, 9) = "Días Seca"
    If WorksheetFunction.VLookup(mArete, _
      Range("Tabla1"), 12, False) = "**SECA**" Then mArrayB(2, 9) = _
      Date - CDate(WorksheetFunction.VLookup(mArete, _
      Range("Tabla15"), 16, False))
    mArrayB(1, 10) = "Serv. x Vaca"
    mArrayB(2, 10) = WorksheetFunction.VLookup( _
      mArete, Range("Tabla1"), 7, False)
    mArrayB(3, 10) = Format(Application.WorksheetFunction. _
      AverageIfs(Range("Tabla1[Servicio]"), Range("Tabla1[Parto]"), _
      iNoLact), "#.0")
    mArrayB(4, 10) = Format(Application.WorksheetFunction. _
      Average(Range("Tabla1[Servicio]")), "#.0")
    mArrayB(1, 11) = "Serv. x Concepción"
    If WorksheetFunction.VLookup(mArete, _
      Range("Tabla1"), 11, False) = "P" Then _
      mArrayB(2, 11) = WorksheetFunction.VLookup( _
      mArete, Range("Tabla1"), 7, False)
    mArrayB(3, 11) = Format(Application.WorksheetFunction. _
      AverageIfs(Range("Tabla1[Servicio]"), Range("Tabla1[Parto]"), _
      iNoLact, Range("Tabla1[Status]"), "P"), "#.0")
    mArrayB(4, 11) = Format(Application.WorksheetFunction. _
      AverageIfs(Range("Tabla1[Servicio]"), Range("Tabla1[Status]"), _
      "P"), "#.0")
    mArrayB(1, 12) = "Días Abiertos"
    mArrayB(2, 12) = WorksheetFunction.VLookup( _
       mArete, Range("Tabla15"), 3, False)
    mArrayB(3, 12) = Format(mDA, "#")
    mArrayB(4, 12) = Format(Application.WorksheetFunction. _
      Average(Range("Tabla15[dAbiertos]")), "#")
    mArrayB(5, 12) = 114
    mArrayB(1, 13) = "Días a 1er Serv."
    mArrayB(2, 13) = WorksheetFunction.VLookup( _
      mArete, Range("Tabla15"), 2, False)
    mArrayB(3, 13) = Format(mD1S, "#")
    mArrayB(4, 13) = Format(Application.WorksheetFunction. _
      Average(Range("Tabla15[d1S]")), "#")
    mArrayB(5, 13) = 45
    mArrayB(1, 14) = vbNullString
'****************
    mArrayB(1, 15) = "LACT. TERMINADAS"
    mArrayB(1, 16) = "Lact. Computadas"
    cLC = Int(Application. _
      WorksheetFunction.CountIf(Range("Tabla4[Arete]"), _
      Me.txtArete))
    mArrayB(2, 16) = cLC
    'mArrayB(2, 16) = Format(Int(Application. _
      WorksheetFunction.CountIf(Range("Tabla4[Arete]"), _
      Me.txtArete)), "#")
    mArrayB(1, 17) = "Prom. Prod. x Lact."
    mArrayB(2, 17) = Format(Int(Application. _
      WorksheetFunction.AverageIfs(Range("Tabla4[ProdAcum]"), _
      Range("Tabla4[Arete]"), Me.txtArete)), "#,#")
    
    If iNoLact > 0 Then _
      mArrayB(3, 17) = Format(xPA / cLC, "#,#")
    
    mArrayB(1, 18) = "Prod. Acumulada"
    mArrayB(2, 18) = Format(Int(Application. _
      WorksheetFunction.SumIfs(Range("Tabla4[ProdAcum]"), _
      Range("Tabla4[Arete]"), Me.txtArete)), "#,#")
    mArrayB(3, 18) = Format(xPA, "#,#")
    mArrayB(1, 19) = "Días en Leche Acum."
    mArrayB(2, 19) = Format(Application.WorksheetFunction. _
      SumIfs(Range("Tabla4[DiasLactancia]"), Range("Tabla4[Arete]"), _
      Me.txtArete), "#,#")
    mArrayB(3, 19) = Format(xDL, "#")
    mArrayB(1, 20) = "Días Seca Acum."
    mArrayB(2, 20) = Application.WorksheetFunction _
      .SumIfs(Range("Tabla4[DíasSeca]"), Range("Tabla4[Arete]"), _
      Me.txtArete)
    mArrayB(3, 20) = Format(xDS, "#")
    mArrayB(1, 21) = "Prom. Serv. x Lact."
    mArrayB(2, 21) = Format(Application.WorksheetFunction _
      .AverageIfs(Range("Tabla4[Servicio]"), _
      Range("Tabla4[Arete]"), Me.txtArete), "#.0")
    
    mArrayB(1, 22) = "Prom. Días Ab. x Lact."
    mArrayB(2, 22) = Format(Application.WorksheetFunction _
      .AverageIfs(Range("Tabla4[DiasAbierta]"), _
      Range("Tabla4[Arete]"), Me.txtArete), "#")
    mArrayB(3, 22) = Format(xDAb, "#")
    mArrayB(1, 23) = "Prom. Días a 1er Serv."
    mArrayB(2, 23) = Format(Application.WorksheetFunction _
      .AverageIfs(Range("Tabla4[DIas1Serv]"), _
      Range("Tabla4[Arete]"), Me.txtArete), "#")
    mArray(3, 23) = Format(xD1S, "#")
    mArrayB(1, 24) = vbNullString
    mArrayB(1, 25) = "CRIANZA"
    mArrayB(1, 26) = "Edad"
    If Not WorksheetFunction.VLookup(mArete, _
              Range("Tabla15"), 22, False) = vbNullString Then mArrayB(2, 26) = _
              Int((Date - CDate(WorksheetFunction. _
              VLookup(mArete, _
              Range("Tabla15"), 22, False))) / 365) & " años" 'Edad
    mArrayB(1, 27) = "Edad al Parto"
    mArrayB(1, 28) = "Edad al 1er Serv."
    If Not WorksheetFunction.VLookup(mArete, _
      Range("Tabla5"), 18, False) = vbNullString Then mArrayB(2, 28) = _
      WorksheetFunction.VLookup(mArete, _
      Range("Tabla5"), 18, False)
    
   On Error GoTo 0
   
    'PoblarMatrizB
    With Me.ListBenchmarking
        .Clear
        .ColumnCount = 5
        .ColumnWidths = "90;0;50;50;50;50"
        For i = 1 To 28
            .AddItem mArrayB(1, i)
            For j = 2 To 5
                .List(.ListCount - 1, j) = mArrayB(j, i)
            Next j
        Next i
    End With
End Sub

Private Sub PoblarLstLactAnteriores()
    Dim i, j As Long
    iNoLact = Application.WorksheetFunction. _
      CountIf(Range("Tabla4[Arete]"), mArete)
    ReDim mArray(iNoLact + 1, 18)
    mArray(1, 1) = "No. Parto"
    mArray(1, 2) = "Fecha Parto"
    mArray(1, 3) = "No. Servicio"
    mArray(1, 4) = "Fecha Servicio"
    mArray(1, 5) = "Toro"
    mArray(1, 6) = "Técnico"
    mArray(1, 7) = "Estatus Reprod."
    mArray(1, 8) = "Clave 1"
    mArray(1, 9) = "Clave 2"
    mArray(1, 10) = "Días en Leche"
    mArray(1, 11) = "Días Seca"
    mArray(1, 12) = "Prod. Acumulada"
    mArray(1, 13) = "Proyección 305d"
    mArray(1, 14) = "Días a 1er Serv."
    mArray(1, 15) = "Días Abierta"
    mArray(1, 16) = "FechaTerminación"
    mArray(1, 17) = "Tipo Terminación"
    mArray(1, 18) = "Causa Terminación"
    PoblarMatrizLA
    With Me.ListLactAnteriores
        .Clear
        If iNoLact >= 6 Then .ColumnCount = iNoLact + 2 Else _
          .ColumnCount = iNoLact + 1
        Select Case iNoLact
            Case Is <= 5
                .ColumnWidths = "75;50;50;50;50;50"
            Case Is = 6
                .ColumnWidths = "75;50;50;50;50;50;50;75"
            Case Is = 7
                .ColumnWidths = "75;50;50;50;50;50;50;50;75"
            Case Is = 8
                .ColumnWidths = "75;50;50;50;50;50;50;50;50;75"
            Case Is > 8
                .ColumnWidths = "75;50;50;50;50;50;50;50;50;50;75"
        End Select
        For i = 1 To 18
            .AddItem mArray(1, i)
            For j = 1 To iNoLact
                .List(.ListCount - 1, j) = mArray(j + 1, i)
                If iNoLact >= 6 Then _
                  .List(.ListCount - 1, j + 1) = mArray(1, i)
            Next j
        Next i
        .AddItem mArray(1, 1) 'Se Repite Para Facilitar Lectura
        For j = 1 To iNoLact
            .List(.ListCount - 1, j) = mArray(j + 1, 1)
        Next j
    End With
End Sub

Private Sub PoblarMatrizLA() 'PoblarLstLactAnteriores
    ' Llena Matriz con Info. de Lact. Terminadas
    Dim col, i As Long
      col = 1 'Columna de la Matriz
    For Each rCelda In Range("Tabla4[Arete]")
        If rCelda.Offset(0, 0) = mArete Then
            For i = 1 To 18 'Renglones de la Matriz
                If i = 2 Or i = 4 Or i = 16 Then
                        mArray(col + 1, i) = _
                          Format(rCelda.Offset(0, i), "dd-mmm-yy")
                    Else
                        mArray(col + 1, i) = rCelda.Offset(0, i)
                End If
                If i = 12 Or i = 13 Then mArray(col + 1, i) = _
                  Format(rCelda.Offset(0, i), "#,#")
            Next i
            col = col + 1
        End If
    Next rCelda
End Sub

Private Sub RegistroActual()
    Me.txtArete = rCelda.Offset
    MuestraDatos
End Sub

Private Sub txtArete_AfterUpdate()
'Buscar Arete
    BorrarDatos
    Me.ListBenchmarking.Clear
    Me.listEventos.Clear
    Me.ListLactAnteriores.Clear
    
    If IsNumeric(Me.txtArete) Then _
      mArete = CDbl(Me.txtArete) Else _
      mArete = Me.txtArete
    If Application.WorksheetFunction. _
      CountIf(Range("Tabla1[Arete]"), mArete) > 0 Then
        For Each rCelda In Range("Tabla1[Arete]")
            If rCelda.Offset(0, 0) = mArete Then
                Set ws = Worksheets("Hato")
                GoTo MostrarInfo
            End If
        Next rCelda
    End If
    'Buscar en Reemplazos
    If Application.WorksheetFunction. _
      CountIf(Range("Tabla2[Arete]"), mArete) > 0 Then
        For Each rCelda In Range("Tabla2[Arete]")
            If rCelda.Offset(0, 0) = mArete Then
                Set ws = Worksheets("Reemplazos")
                GoTo MostrarInfo
            End If
        Next rCelda
    End If
    'Buscar en InfoVitalicia
    If Application.WorksheetFunction. _
      CountIf(Range("Tabla6[Arete]"), mArete) > 0 Then
        For Each rCelda In Range("Tabla6[Arete]")
            If rCelda.Offset(0, 0) = mArete Then
                Set ws = Worksheets("InfoVitalicia")
                GoTo MostrarInfo
            End If
        Next rCelda
    End If
    
    MsgBox "Este animal no se encuentra en la Base de Datos", _
      vbCritical, "Consulta de Registros Individuales"
    Exit Sub
    
MostrarInfo:
    mArete = rCelda.Offset
    MuestraDatos
    Me.txtArete.SetFocus
End Sub

Private Sub UserForm_Initialize()
    Dim mTabla As Range
    Application.ScreenUpdating = _
      CBool(Range("Desarrollador!B6"))
    Application.Run "OrdenarEventos" 'Mod2
    Me.txtArete = ActiveCell
    txtArete_AfterUpdate
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
