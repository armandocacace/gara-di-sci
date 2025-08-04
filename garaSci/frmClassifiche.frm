VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmClassifiche 
   Caption         =   "Form2"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8070
   LinkTopic       =   "Form2"
   ScaleHeight     =   5625
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdApriStatistiche 
      Caption         =   "Statistiche"
      Height          =   615
      Left            =   2760
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grdClassifica 
      Height          =   2295
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   4048
      _Version        =   393216
   End
   Begin VB.ComboBox cmbGara 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Text            =   "Scegli una gara"
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Elenco delle gare"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.Image imgCoppa 
      Height          =   375
      Left            =   5880
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Classifiche e Statistiche"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmClassifiche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdApriStatistiche_Click()
    If cmbGara.ListIndex = -1 Then
        MsgBox "Seleziona una gara prima di visualizzare le statistiche.", vbExclamation
        Exit Sub
    End If

    frmStatistiche.idGaraSelezionata = cmbGara.ItemData(cmbGara.ListIndex)
    frmStatistiche.Show
End Sub

Private Sub Form_Load()
    Dim rs As New ADODB.Recordset
    cmbGara.Clear

    rs.Open "SELECT DISTINCT G.ID, G.Titolo, G.DataInizio " & _
            "FROM Gare G INNER JOIN Discesa D ON G.ID = D.IDGara " & _
            "ORDER BY G.DataInizio DESC", connect, adOpenStatic, adLockReadOnly

    Do While Not rs.EOF
        cmbGara.AddItem rs!titolo & " (" & Format(rs!DataInizio, "dd/mm/yyyy") & ")"
        cmbGara.ItemData(cmbGara.NewIndex) = rs!ID
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
End Sub

Private Sub cmbGara_Click()
    Dim rs As New ADODB.Recordset
    Dim sql As String
    Dim idGara As Integer
    Dim posizione As Integer
    Dim topPos As Integer, leftPos As Integer

    idGara = cmbGara.ItemData(cmbGara.ListIndex)

    ' Ottieni i risultati ordinati per tempo crescente, escludendo caduti e squalificati
    sql = "SELECT S.Nome, S.Cognome, D.Tempo " & _
          "FROM Discesa D INNER JOIN Sciatori S ON D.IDSciatore = S.ID " & _
          "WHERE D.IDGara = " & idGara & " AND D.Caduto = False AND D.Squalificato = False " & _
          "ORDER BY D.Tempo ASC"

    rs.Open sql, connect, adOpenStatic, adLockReadOnly

    ' Imposta intestazioni
    With grdClassifica
        .Clear
        .Rows = 1
        .Cols = 5
        .TextMatrix(0, 0) = "Pos."
        .TextMatrix(0, 1) = "Nome"
        .TextMatrix(0, 2) = "Cognome"
        .TextMatrix(0, 3) = "Tempo"
        .TextMatrix(0, 4) = "Premio"
        .ColWidth(4) = 1500
    End With

    posizione = 1
    Do While Not rs.EOF
        grdClassifica.Rows = grdClassifica.Rows + 1
        grdClassifica.TextMatrix(posizione, 0) = posizione
        grdClassifica.TextMatrix(posizione, 1) = rs!Nome
        grdClassifica.TextMatrix(posizione, 2) = rs!Cognome
        grdClassifica.TextMatrix(posizione, 3) = Format(rs!tempo, "0.00")
        grdClassifica.TextMatrix(posizione, 4) = "" ' Lascia vuota la colonna premio per tutti tranne il primo
        
        
        If posizione = 1 Then
            grdClassifica.TextMatrix(posizione, 4) = "gara omaggio"
            ' Calcola posizione della prima riga dati nella griglia
            topPos = grdClassifica.Top + grdClassifica.RowHeight(0) * posizione
            ' Somma larghezza colonne precedenti per la colonna "Premio"
            leftPos = grdClassifica.Left
            Dim i As Integer
            For i = 0 To 3
                leftPos = leftPos + grdClassifica.ColWidth(i)
            Next i

            ' Imposta immagine coppa
            imgCoppa.Visible = True
            imgCoppa.ZOrder 0
            imgCoppa.Picture = LoadPicture(App.Path & "\Immagine1.bmp")

            imgCoppa.Width = 360
            imgCoppa.Height = 360
        End If

        posizione = posizione + 1
        rs.MoveNext
    Loop

    If posizione = 1 Then
        ' Nessun risultato, nascondi coppa
        imgCoppa.Visible = False
    End If

    rs.Close
    Set rs = Nothing
End Sub
