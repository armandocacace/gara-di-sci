VERSION 5.00
Begin VB.Form frmGare 
   Caption         =   "Form2"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8775
   LinkTopic       =   "Form2"
   ScaleHeight     =   6345
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbCategorie 
      Height          =   315
      Left            =   4440
      TabIndex        =   4
      Text            =   "Clicca per selezionare una categoria"
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox txtDescrizioneGara 
      Height          =   735
      Left            =   4440
      TabIndex        =   3
      Text            =   "Descrizione"
      Top             =   4200
      Width           =   2895
   End
   Begin VB.TextBox txtDataGara 
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Text            =   "Data"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox txtTitoloGara 
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Text            =   "Titolo"
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton cmdSalvaGara 
      Caption         =   "Salva gara"
      Height          =   615
      Left            =   3240
      TabIndex        =   0
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Scrivi una descrizione dettagliata per non lasciare nulla al caso:"
      Height          =   495
      Index           =   3
      Left            =   840
      TabIndex        =   9
      Top             =   4320
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Seleziona la categoria che potrà gareggiare:"
      Height          =   495
      Index           =   2
      Left            =   840
      TabIndex        =   8
      Top             =   3360
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Indica la data di inizio della gara (formato gg/mm/aaaa):"
      Height          =   495
      Index           =   1
      Left            =   840
      TabIndex        =   7
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Scegli un titlolo accattivante:"
      Height          =   495
      Index           =   0
      Left            =   840
      TabIndex        =   6
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Crea una Nuova Gara"
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
      Left            =   1920
      TabIndex        =   5
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "frmGare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim rs As New ADODB.Recordset
    cmbCategorie.Clear

    rs.Open "SELECT ID, NomeCategoria FROM Categorie", connect, adOpenStatic, adLockReadOnly

    Do While Not rs.EOF
        cmbCategorie.AddItem rs!NomeCategoria
        cmbCategorie.ItemData(cmbCategorie.NewIndex) = rs!ID
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
End Sub

Private Sub cmdSalvaGara_Click()
    Dim sql As String
    Dim dataGara As Date
    Dim titolo As String
    Dim descrizione As String
    Dim idCategoria As Integer
    Dim rsID As New ADODB.Recordset
    Dim idGara As Integer

    If Trim(txtTitoloGara.Text) = "" Or Trim(txtDataGara.Text) = "" Or _
       Trim(txtDescrizioneGara.Text) = "" Or cmbCategorie.ListIndex = -1 Then
        MsgBox "Compila tutti i campi", vbExclamation
        Exit Sub
    End If

    If Not IsDate(txtDataGara.Text) Then
        MsgBox "La data inserita non è valida. Usa il formato GG/MM/AAAA.", vbExclamation
        Exit Sub
    End If

    ' Recupera e prepara i dati
    dataGara = Format(CDate(txtDataGara.Text), "yyyy-mm-dd")
    titolo = Replace(txtTitoloGara.Text, "'", "''")
    descrizione = Replace(txtDescrizioneGara.Text, "'", "''")
    idCategoria = cmbCategorie.ItemData(cmbCategorie.ListIndex)

    ' Query SQL
    sql = "INSERT INTO Gare (Titolo, DataInizio, Descrizione) " & _
          "VALUES ('" & titolo & "', #" & dataGara & "#, '" & descrizione & "')"
    connect.Execute sql
    
    ' 2. Recupera l'ID appena inserito (ultimo ID)
    rsID.Open "SELECT MAX(ID) AS IDGara FROM Gare", connect, adOpenStatic, adLockReadOnly
    idGara = rsID!idGara
    rsID.Close
    Set rsID = Nothing

    ' 3. Inserisce nella tabella ponte
    sql = "INSERT INTO CategorieGara (IDGara, IDCategoria) VALUES (" & idGara & ", " & idCategoria & ")"
    connect.Execute sql
    
    On Error GoTo errore
    connect.Execute sql
    MsgBox "Gara salvata!"

    ' Pulisce i campi
    txtTitoloGara.Text = ""
    txtDataGara.Text = ""
    txtDescrizioneGara.Text = ""
    cmbCategorie.ListIndex = -1
    Exit Sub

errore:
    MsgBox "Errore: " & Err.Description
End Sub
