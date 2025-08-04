VERSION 5.00
Begin VB.Form frmCategorie 
   Caption         =   "Form2"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8325
   LinkTopic       =   "Form2"
   ScaleHeight     =   5865
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalvaCategoria 
      Caption         =   "Salva"
      Height          =   615
      Left            =   2880
      TabIndex        =   3
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtEtaMax 
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Text            =   "Età massima"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox txtEtaMin 
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Text            =   "Età minima"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox txtNomeCat 
      Height          =   375
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmCategorie.frx":0000
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Crea Categorie"
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
      Left            =   720
      TabIndex        =   7
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label Label1 
      Caption         =   "Inserisci l'età massima per partecipare:"
      Height          =   615
      Index           =   2
      Left            =   5040
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Inserisci il nome della categoria da creare:"
      Height          =   615
      Index           =   1
      Left            =   720
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "inserisci l'età minima per la categoria:"
      Height          =   615
      Index           =   0
      Left            =   2760
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
End
Attribute VB_Name = "frmCategorie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSalvaCategoria_Click()
    Dim sql As String

    If Trim(txtNomeCat.Text) = "" Or Trim(txtEtaMin.Text) = "" Or Trim(txtEtaMax.Text) = "" Then
        MsgBox "Compila tutti i campi", vbExclamation
        Exit Sub
    End If

    sql = "INSERT INTO Categorie (NomeCategoria, EtaMinima, EtaMassima) " & _
          "VALUES ('" & txtNomeCat.Text & "', " & txtEtaMin.Text & ", " & txtEtaMax.Text & ")"

    On Error GoTo errore
    connect.Execute sql
    MsgBox "Categoria salvata!"
    
    ' Pulisce i campi
    txtNomeCat.Text = ""
    txtEtaMin.Text = ""
    txtEtaMax.Text = ""
    Exit Sub

errore:
    MsgBox "Errore: " & Err.Description
End Sub

