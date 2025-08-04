Attribute VB_Name = "modDB"
Public connect As ADODB.Connection

Public Sub AperturaConnessione()
    Set connect = New ADODB.Connection
    connect.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=(percorso del file mdb)"
End Sub

