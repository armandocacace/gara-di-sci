VERSION 5.00
Begin VB.Form frmStatistiche 
   Caption         =   "Form2"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5475
   LinkTopic       =   "Form2"
   ScaleHeight     =   4590
   ScaleWidth      =   5475
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblMedia 
      Caption         =   "Label9"
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label lblPeggiore 
      Caption         =   "Label8"
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lblMigliore 
      Caption         =   "Label7"
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label lblPartecipanti 
      Caption         =   "Label6"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Tempo medio di discesa:"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Tempo peggiore"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Tempo migliore:"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Numero partecipanti:"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Statistiche"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "frmStatistiche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public idGaraSelezionata As Integer


Private Sub Form_Load()
    Dim rs As New ADODB.Recordset
    Dim sql As String

    ' Numero partecipanti (escludendo caduti e squalificati)
    sql = "SELECT COUNT(*) AS Totale FROM Discesa WHERE IDGara = " & idGaraSelezionata & _
          " AND Caduto = False AND Squalificato = False"
    rs.Open sql, connect, adOpenStatic, adLockReadOnly
    lblPartecipanti.Caption = rs!Totale
    rs.Close

    ' Miglior tempo, peggior tempo, media tempo
    sql = "SELECT MIN(Tempo) AS Migliore, MAX(Tempo) AS Peggiore, AVG(Tempo) AS Media " & _
          "FROM Discesa WHERE IDGara = " & idGaraSelezionata & " AND Caduto = False AND Squalificato = False"
    rs.Open sql, connect, adOpenStatic, adLockReadOnly

    lblMigliore.Caption = Format(rs!Migliore, "0.00")
    lblPeggiore.Caption = Format(rs!Peggiore, "0.00")
    lblMedia.Caption = Format(rs!Media, "0.00")

    rs.Close
    Set rs = Nothing
End Sub

