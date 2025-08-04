VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDiscesa 
   Caption         =   "Form2"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9360
   LinkTopic       =   "Form2"
   ScaleHeight     =   8970
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid grdClassifica 
      Height          =   2415
      Left            =   240
      TabIndex        =   9
      Top             =   6480
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4260
      _Version        =   393216
   End
   Begin VB.ComboBox cmbPartecipante 
      Height          =   315
      Left            =   360
      TabIndex        =   8
      Text            =   "Seleziona un partecipante"
      Top             =   5400
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid grdIscritti 
      Height          =   1575
      Left            =   360
      TabIndex        =   7
      Top             =   2640
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   2778
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSalvaRisultato 
      Caption         =   "Salva Discesa"
      Height          =   615
      Left            =   7440
      TabIndex        =   6
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CheckBox chkCaduto 
      Caption         =   "Caduto"
      Height          =   615
      Left            =   5280
      TabIndex        =   5
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CheckBox chkSqualificato 
      Caption         =   "Squalificato"
      Height          =   615
      Left            =   5280
      TabIndex        =   4
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox txtTempo 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Text            =   "Tempo ( minuti.secondi)"
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton cmdCaricaClassifica 
      Caption         =   "Mostra Classifica"
      Height          =   615
      Left            =   3360
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
   Begin VB.ComboBox cmbGara 
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Text            =   "Seleziona una gara"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Registro Tempi di Discesa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   11
      Top             =   360
      Width           =   6975
   End
   Begin VB.Label Label1 
      Caption         =   "Scegli i partecipanti ed inserisci il tempo di discesa, spuntando una delle checkbox se è caduto o è stato squalificato"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   4560
      Width           =   6855
   End
   Begin VB.Label lblListaGare 
      Caption         =   "Scegli la gara che vuoi modificare"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
   End
End
Attribute VB_Name = "frmDiscesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCaricaClassifica_Click()
    Dim idGara As Integer
    
    If cmbGara.ListIndex = -1 Then
        MsgBox "Seleziona una gara."
        Exit Sub
    End If

    idGara = cmbGara.ItemData(cmbGara.ListIndex)
    Call CaricaClassifica(idGara)
End Sub

Private Sub Form_Load()
    Dim rs As New ADODB.Recordset
    cmbGara.Clear

    rs.Open "SELECT ID, Titolo FROM Gare ORDER BY DataInizio DESC", connect, adOpenStatic, adLockReadOnly

    Do While Not rs.EOF
        cmbGara.AddItem rs!titolo
        cmbGara.ItemData(cmbGara.NewIndex) = rs!ID
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
End Sub
Private Sub cmbGara_Click()
    Dim idGara As Integer
    Dim rs As New ADODB.Recordset

    idGara = cmbGara.ItemData(cmbGara.ListIndex)
    grdIscritti.Clear

    ' Intestazioni
    With grdIscritti
        .Cols = 4
        .Rows = 1
        .TextMatrix(0, 0) = "ID"
        .TextMatrix(0, 1) = "Nome"
        .TextMatrix(0, 2) = "Cognome"
        .TextMatrix(0, 3) = "Data Nascita"
    End With

    ' Query iscritti
    rs.Open "SELECT DISTINCT S.ID, S.Nome, S.Cognome, S.DataNascita " & _
            "FROM Iscrizioni I INNER JOIN Sciatori S ON I.IDSciatore = S.ID " & _
            "WHERE I.IDGara = " & idGara, connect, adOpenStatic, adLockReadOnly

    Dim r As Integer: r = 1
    Do While Not rs.EOF
        grdIscritti.Rows = grdIscritti.Rows + 1
        grdIscritti.TextMatrix(r, 0) = rs!ID
        grdIscritti.TextMatrix(r, 1) = rs!Nome
        grdIscritti.TextMatrix(r, 2) = rs!Cognome
        grdIscritti.TextMatrix(r, 3) = rs!DataNascita
        rs.MoveNext
        r = r + 1
    Loop
    rs.Close
    Set rs = Nothing

    ' Carica la Combo partecipante
    Call CaricaPartecipanti(idGara)
End Sub
Private Sub CaricaPartecipanti(idGara As Integer)
    Dim rs As New ADODB.Recordset
    cmbPartecipante.Clear

    rs.Open "SELECT DISTINCT S.ID, S.Nome & ' ' & S.Cognome AS NomeCompleto " & _
            "FROM Iscrizioni I INNER JOIN Sciatori S ON I.IDSciatore = S.ID " & _
            "WHERE I.IDGara = " & idGara, connect, adOpenStatic, adLockReadOnly

    Do While Not rs.EOF
        cmbPartecipante.AddItem rs!NomeCompleto
        cmbPartecipante.ItemData(cmbPartecipante.NewIndex) = rs!ID
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
End Sub

Function ConvertiTempo(testo As String) As Double
    Dim parti() As String
    Dim minuti As Integer
    Dim secondi As Integer

    ' Gestisce input tipo "1.10", "0.45", "2.00"
    If InStr(testo, ".") > 0 Then
        parti = Split(testo, ".")
        If UBound(parti) = 1 Then
            minuti = Val(parti(0))
            secondi = Val(parti(1))
        Else
            ConvertiTempo = 0
            Exit Function
        End If
    Else
        ' Se non c'è punto, lo tratta come minuti interi
        minuti = Val(testo)
        secondi = 0
    End If

    ' Se seconds >= 60, considera errore
    If secondi >= 60 Then
        MsgBox "I secondi devono essere minori di 60", vbExclamation
        ConvertiTempo = 0
        Exit Function
    End If

    ConvertiTempo = minuti * 60 + secondi
End Function


Private Sub cmdSalvaRisultato_Click()
    Dim idGara As Integer
    Dim idSciatore As Integer
    Dim tempo As String
    Dim isCaduto As String
    Dim isSqualificato As String
    Dim sql As String
    Dim rsCheck As New ADODB.Recordset

    If cmbPartecipante.ListIndex = -1 Then
        MsgBox "Seleziona un partecipante."
        Exit Sub
    End If

    idGara = cmbGara.ItemData(cmbGara.ListIndex)
    idSciatore = cmbPartecipante.ItemData(cmbPartecipante.ListIndex)

    ' Determina tempo o 0 se caduto/squalificato
    If chkCaduto.Value = 1 Or chkSqualificato.Value = 1 Then
        tempo = "0"
    Else
        If Trim(txtTempo.Text) = "" Then
            MsgBox "Inserisci il tempo oppure seleziona caduto/squalificato", vbExclamation
            Exit Sub
        End If
        tempo = ConvertiTempo(Replace(txtTempo.Text, ",", "."))
    End If

    isCaduto = IIf(chkCaduto.Value = 1, "True", "False")
    isSqualificato = IIf(chkSqualificato.Value = 1, "True", "False")

    On Error GoTo errore

    ' Verifica se esiste già un record per lo stesso sciatore e gara
    sql = "SELECT * FROM Discesa WHERE IDGara = " & idGara & " AND IDSciatore = " & idSciatore
    rsCheck.Open sql, connect, adOpenStatic, adLockReadOnly

    If Not rsCheck.EOF Then
        ' Record già presente ? aggiorna
        sql = "UPDATE Discesa SET Tempo = " & tempo & _
              ", Caduto = " & isCaduto & ", Squalificato = " & isSqualificato & _
              " WHERE IDGara = " & idGara & " AND IDSciatore = " & idSciatore
    Else
        ' Nuovo record ? inserisci
        sql = "INSERT INTO Discesa (IDSciatore, IDGara, Tempo, Caduto, Squalificato) " & _
              "VALUES (" & idSciatore & ", " & idGara & ", " & tempo & ", " & isCaduto & ", " & isSqualificato & ")"
    End If
    
    Call CaricaClassifica(idGara)
    
    rsCheck.Close
    Set rsCheck = Nothing

    connect.Execute sql

    MsgBox "Risultato salvato!", vbInformation
    txtTempo.Text = ""
    chkCaduto.Value = 0
    chkSqualificato.Value = 0
    Exit Sub

errore:
    MsgBox "Errore: " & Err.Description
End Sub

Private Sub CaricaClassifica(idGara As Integer)
    Dim rs As New ADODB.Recordset
    Dim sql As String
    Dim r As Integer

    sql = "SELECT S.ID, S.Nome, S.Cognome, D.Tempo, D.Caduto, D.Squalificato " & _
          "FROM Discesa D INNER JOIN Sciatori S ON D.IDSciatore = S.ID " & _
          "WHERE D.IDGara = " & idGara & " " & _
          "ORDER BY IIf(D.Tempo = 0, 999999, D.Tempo) ASC"

    rs.Open sql, connect, adOpenStatic, adLockReadOnly

    ' Imposta intestazioni
    With grdClassifica
        .Rows = 1
        .Cols = 6
        .TextMatrix(0, 0) = "ID"
        .TextMatrix(0, 1) = "Nome"
        .TextMatrix(0, 2) = "Cognome"
        .TextMatrix(0, 3) = "Tempo"
        .TextMatrix(0, 4) = "Caduto"
        .TextMatrix(0, 5) = "Squalificato"
        
        ' Riempie righe
        Do While Not rs.EOF
            .Rows = .Rows + 1
            r = .Rows - 1
            .TextMatrix(r, 0) = rs!ID
            .TextMatrix(r, 1) = rs!Nome
            .TextMatrix(r, 2) = rs!Cognome
            
            If rs!tempo = 0 Then
                .TextMatrix(r, 3) = "-"
            Else
                .TextMatrix(r, 3) = rs!tempo
            End If

            .TextMatrix(r, 4) = IIf(rs!Caduto = True, "Sì", "")
            .TextMatrix(r, 5) = IIf(rs!Squalificato = True, "Sì", "")
            rs.MoveNext
        Loop
    End With

    rs.Close
    Set rs = Nothing
End Sub

