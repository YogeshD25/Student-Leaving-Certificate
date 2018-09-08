Attribute VB_Name = "Module1"
Function NewTCNo()
    Dim Conn As New ADODB.Connection
    Dim Rs As New ADODB.Recordset
    Dim id As Integer
    
    Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Studdetail.mdb;Persist Security Info=False"
    Conn.Open
    Rs.Open "Select Last(TCNo) from Stud_reg", Conn
    If Rs.EOF = True Then
        student_reg.RegiNo.SelText = "1"
    Else
        id = Rs.Fields(0)
        student_reg.tcNo.SelText = id + 1
    End If
    Rs.Close
    Conn.Close
End Function
  
