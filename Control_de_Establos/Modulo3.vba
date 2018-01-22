Attribute VB_Name = "Modulo3"
' Ultima modificación: 21-Ene-16
Option Explicit

Function IndiceTabla(nArete, nTabla)
Attribute IndiceTabla.VB_ProcData.VB_Invoke_Func = " \n14"
    'Devuelve el renglón donde se encuentra el arete
    Dim rCelda As Object
    If IsNumeric(nArete) Then nArete = CDbl(nArete)
    IndiceTabla = 0
    nTabla = nTabla & "[Arete]"
    If Application.WorksheetFunction.CountIf(Range(nTabla), _
      nArete) = 0 Then Exit Function 'Salir si no existe
    For Each rCelda In Range(nTabla)
        If rCelda = nArete Then
            IndiceTabla = rCelda.Offset.Row
            Exit Function
        End If
    Next rCelda
End Function

Function TamañoTabla(sTabla)
' Devuelve el número de integrantes de la tabla
     If IsEmpty(Application.WorksheetFunction.Index(Range(sTabla), 1, 1)) _
      Then _
      TamañoTabla = 0 Else _
      TamañoTabla = Range(sTabla).Rows.Count 'No considera encabezado
End Function

Function BuscarEvento(Arete As Variant, _
  Evento As Variant, Fecha As Date)
' Devuelve el valor del renglón donde se encuentra el evento
' Ejemplo: =BuscarEvento(1084,"Serv", "1-Nov-2015")
    Dim rw As Long
    Dim Ocurrencia As Double
    Dim rCelda As Range
    rw = 0
    BuscarEvento = 0
    ' Para ahorrar tiempo, contar las ocurrencias
    If WorksheetFunction.CountIfs( _
      Range("Tabla6[Arete]"), Arete, _
      Range("Tabla6[Clave]"), Evento, _
      Range("Tabla6[Fecha]"), Fecha) = 0 Then GoTo 1234
    ' Localiza la ocurrencia
    For Each rCelda In Range("Tabla6[Arete]")
        If Val(rCelda.Offset(rw, 0)) = Arete And _
          rCelda.Offset(rw, 2) = Evento And _
          rCelda.Offset(rw, 1) = Fecha Then
          BuscarEvento = rCelda.Row
        End If
    Next
1234:
End Function

Function BuscarUltimaOcurrencia(Valor_Buscado As Long, _
  Matriz_Buscar As Range, Indicador_Columna As Long)
' Valor_Buscado: El valor que estamos buscando.
' Matriz_Buscar: El rango de celdas con los datos. La búsqueda, al igual _
  que con la funcion BUSCARV, se hará en la primera columna del rango.
' Ocurrencia: El número de ocurrencia del Valor_Buscado que requerimos.
' Indicador_Columna: El número de columna que va a devolver la función.
' Ejemplo: =BuscarUltimaOcurrencia(1084, Tabla6[Arete], 1)
    Dim cont, i, Ocurrencia As Long
    Ocurrencia = WorksheetFunction. _
      CountIf(Matriz_Buscar, Valor_Buscado)
    BuscarUltimaOcurrencia = "No existe"
    For i = 1 To Matriz_Buscar.Rows.Count
       If Matriz_Buscar.Cells(i, 1) = Valor_Buscado Then
           cont = cont + 1
           If cont = Ocurrencia Then
               BuscarUltimaOcurrencia = _
                 Matriz_Buscar.Cells(i, Indicador_Columna)
               Exit Function
           End If
       End If
    Next
End Function

Function BuscarUltimoEvento(Arete_Buscado As Variant, _
  Evento_Buscado As String)
' Devuelve la fecha del último evento buscado
' Arete_Buscado
' Evento_Buscado:
' Ejemplo: =BuscarUltimoEvento(1084,"Serv")
    Dim cont, i As Long
    Dim Ocurrencia As Double
    Dim rCelda As Range
    ' Contar las ocurrencias de estos eventos
    Ocurrencia = WorksheetFunction. _
      CountIfs(Range("Tabla6[Arete]"), Arete_Buscado, _
        Range("Tabla6[Clave]"), Evento_Buscado)
    BuscarUltimoEvento = "27-jun-1959" 'comienzo del mundo :)
    For Each rCelda In Range("Tabla6[Arete]")
       If Val(rCelda.Offset(i, 0)) = _
         Arete_Buscado And _
         rCelda.Offset(i, 2) = Evento_Buscado Then
           cont = cont + 1
           ' Si es la última ocurrencia del evento
           If cont = Ocurrencia Then
               ' Devuelve fecha del ultimo evento
               BuscarUltimoEvento = _
                 rCelda.Offset(i, 1)
               GoTo 100
           End If
       End If
    Next
100:
End Function

Function BuscarValorInverso(ValorBuscado As Variant, Rango As Range, _
  Posicion As Long)
Attribute BuscarValorInverso.VB_Description = "Devuelve el valor de la columna indicada por posición.\r\nEjemplo BuscarValorInverso(""A"",M1:M1000,-5)"
  ' BUSCARVI
  ' Devuelve el valor de la columna indicada por posición
  ' Ejemplo BUSCARVI("A",m1:m2000,-5), Indicará el valor de la columna H
    Dim rCelda As Range
    For Each rCelda In Rango
        If rCelda.Offset(0, 0) = ValorBuscado Then
            BuscarValorInverso = rCelda.Offset(0, Posicion)
            GoTo 100
        End If
    Next
100:
End Function

Function BuscarHoja(nombreHoja As String) As Boolean
'https://exceltotal.com/comprobar-si-existe-una-hoja-de-excel-desde-vba/
' Indica si existe una hoja determinada
    Dim i As Long
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = nombreHoja Then
            BuscarHoja = True
            Exit Function
        End If
    Next
    BuscarHoja = False
End Function

Function XLMod(a, b)
    ' This replicates the Excel MOD function
    XLMod = a - b * Int(a / b)
End Function

