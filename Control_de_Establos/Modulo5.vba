Attribute VB_Name = "Modulo5"
Option Explicit

Private Sub DBReader()
    Dim cnt As New ADODB.Connection
    Dim rst As New Recordset
    Dim strConnectStr As String
    Dim Qry As String
    
    strConnectStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Temp\Test.mdb;"
    
    Qry = "SELECT tblOurTable.[Product Name], tblOurTable.[Product ID], tblOurTable.[Price Each] FROM tblOurTable;"
    
    ActiveSheet.Cells.ClearContents
    
    cnt.Open strConnectStr
    
    rst.Open Qry, cnt
    
    Range("B3:D3").CopyFromRecordset rst
    
End Sub
