Attribute VB_Name = "modDB"
Public connect As ADODB.Connection

Public Sub AperturaConnessione()
    Set connect = New ADODB.Connection
    connect.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Documents and Settings\Administrator\Desktop\ProgettoGara\DataBase\gareSci.mdb"
End Sub
