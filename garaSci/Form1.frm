VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form2"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6375
   LinkTopic       =   "Form2"
   ScaleHeight     =   7125
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClassificheStatistiche 
      Caption         =   "Classifiche e Statistiche"
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton cmdIscrizioni 
      Caption         =   "Gestione Iscrizioni"
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdApriDiscesa 
      Caption         =   "Registra Discesa"
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton apriGare 
      Caption         =   "Registra Gara"
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton apriCategorie 
      Caption         =   "Registra Categorie"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton apriSciatori 
      Caption         =   "Registra sciatore"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Gestione Impianto Sciistisco per Gare Amatoriali"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub apriCategorie_Click()
    frmCategorie.Show
End Sub

Private Sub apriGare_Click()
    frmGare.Show
End Sub


Private Sub cmdApriDiscesa_Click()
    frmDiscesa.Show
End Sub

Private Sub cmdClassificheStatistiche_Click()
    frmClassifiche.Show
End Sub

Private Sub cmdIscrizioni_Click()
    frmIscrizioni.Show
End Sub

Private Sub Form_Load()
    Call modDB.AperturaConnessione
    MsgBox "Connessione Riuscita!"
  
End Sub

Private Sub Label1_Click()
    Label1.Caption = "Gara di Sci"
End Sub

Private Sub apriSciatori_Click()
    frmSciatori.Show
End Sub

