VERSION 5.00
Begin VB.Form frmSciatori 
   Caption         =   "Form2"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8055
   LinkTopic       =   "Form2"
   ScaleHeight     =   6060
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalva 
      Caption         =   "Aggiungi al DB"
      Height          =   615
      Left            =   3480
      TabIndex        =   4
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CheckBox chkVisita 
      Caption         =   "Visita fatta"
      Height          =   735
      Left            =   3600
      TabIndex        =   3
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox txtDataNascita 
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Text            =   "Data di nascita"
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox txtCognome 
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Text            =   "Cognome"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtNome 
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Text            =   "Nome"
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Controlla i dati inseriti e salva le modifiche"
      Height          =   735
      Left            =   720
      TabIndex        =   11
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "- Spunta la casella se il nuovo iscritto  ha già fatto la visita medica"
      Height          =   735
      Left            =   600
      TabIndex        =   10
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label Label5 
      Caption         =   "- Inserisci la data di nascita"
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "- Inserisci il cognome"
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "- Inserire il nominativo"
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Inserisci i dati anagrafici del nuovo iscritto:"
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   960
      Width           =   6015
   End
   Begin VB.Label Label1 
      Caption         =   "Iscrizione"
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
      Left            =   600
      TabIndex        =   5
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmSciatori"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdSalva_Click()
    Dim sql As String
    Dim visita As Integer
    Dim eta As Integer
    Dim idCategoria As Integer
    Dim rs As ADODB.Recordset


    If connect Is Nothing Then
        MsgBox "Connessione mancante!"
        Exit Sub
    End If
    
    eta = DateDiff("yyyy", CDate(txtDataNascita.Text), Date)
    If Month(Date) < Month(CDate(txtDataNascita.Text)) Or _
       (Month(Date) = Month(CDate(txtDataNascita.Text)) And Day(Date) < Day(CDate(txtDataNascita.Text))) Then
        eta = eta - 1
    End If

    ' Trova la categoria in base all’età
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM Categorie WHERE EtaMinima <= " & eta & " AND EtaMassima >= " & eta, connect, adOpenStatic, adLockReadOnly
    
    
    If rs.EOF Then
        MsgBox "Nessuna categoria trovata per età: " & eta, vbExclamation
        rs.Close
        Set rs = Nothing
        Exit Sub
    Else
        idCategoria = rs!ID
        rs.Close
        Set rs = Nothing
    End If
    
    ' Converte checkbox in 0/1
    If chkVisita.Value = 1 Then
        visita = 1
    Else
        visita = 0
    End If
    
    sql = "INSERT INTO Sciatori (Nome, Cognome, DataNascita, VisitaMedica, IDCategoria) " & _
          "VALUES ('" & txtNome.Text & "', '" & txtCognome.Text & "', #" & txtDataNascita.Text & "#, " & visita & "," & idCategoria & ")"

    On Error GoTo errore
    connect.Execute sql
    MsgBox "Sciatore salvato con successo!"
    
    ' Pulisce i campi
    txtNome.Text = ""
    txtCognome.Text = ""
    txtDataNascita.Text = ""
    chkVisita.Value = 0
    Exit Sub

errore:
    MsgBox "Errore: " & Err.Description
End Sub

