VERSION 5.00
Begin VB.Form frmIscrizioni 
   Caption         =   "Form2"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5955
   LinkTopic       =   "Form2"
   ScaleHeight     =   5130
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdIscrivi 
      Caption         =   "Iscrivi"
      Height          =   615
      Left            =   3360
      TabIndex        =   2
      Top             =   2880
      Width           =   1575
   End
   Begin VB.ComboBox cmbSciatori 
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Text            =   "Nessun atleta selezionato"
      Top             =   3000
      Width           =   2295
   End
   Begin VB.ComboBox cmbGara 
      Height          =   315
      Left            =   600
      TabIndex        =   0
      Text            =   "Clicca per selezionare una gara"
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Sciatori idonei"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Form Iscrizioni"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   3
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "frmIscrizioni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Dim rs As New ADODB.Recordset
    cmbGara.Clear
    
    rs.Open "SELECT ID, Titolo FROM Gare WHERE DataInizio >= Date()", connect, adOpenStatic, adLockReadOnly

    Do While Not rs.EOF
        cmbGara.AddItem rs!titolo
        cmbGara.ItemData(cmbGara.NewIndex) = rs!ID
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
End Sub

Private Sub cmbGara_Click()
    Dim rs As New ADODB.Recordset
    Dim idGara As Integer
    Dim sql As String
    
    If cmbGara.ListIndex = -1 Then Exit Sub
    idGara = cmbGara.ItemData(cmbGara.ListIndex)
    
    cmbSciatori.Clear
    
    sql = "SELECT DISTINCT S.ID, S.Nome, S.Cognome, S.DataNascita " & _
          "FROM Sciatori S " & _
          "INNER JOIN CategorieGara CG ON S.IDCategoria = CG.IDCategoria " & _
          "WHERE CG.IDGara = " & idGara & " AND S.VisitaMedica = True"
    
    rs.Open sql, connect, adOpenStatic, adLockReadOnly
    
    Do While Not rs.EOF
        Dim visualizza As String
        visualizza = "[" & rs!ID & "] " & rs!Cognome & " " & rs!Nome & " - " & Format(rs!DataNascita, "dd/mm/yyyy")
        
        cmbSciatori.AddItem visualizza
        cmbSciatori.ItemData(cmbSciatori.NewIndex) = rs!ID
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub cmdIscrivi_Click()
    Dim idGara As Integer
    Dim idSciatore As Integer
    Dim sql As String
    
    If cmbGara.ListIndex = -1 Then
        MsgBox "Seleziona una gara", vbExclamation
        Exit Sub
    End If
    
    If cmbSciatori.ListIndex = -1 Then
        MsgBox "Seleziona uno sciatore", vbExclamation
        Exit Sub
    End If
    
    idGara = cmbGara.ItemData(cmbGara.ListIndex)
    idSciatore = cmbSciatori.ItemData(cmbSciatori.ListIndex)
    
    sql = "INSERT INTO Iscrizioni (IDGara, IDSciatore) VALUES (" & idGara & ", " & idSciatore & ")"
    
    On Error GoTo errore
    connect.Execute sql
    MsgBox "Iscrizione avvenuta con successo!"
    Exit Sub
    
errore:
    MsgBox "Errore durante l'iscrizione: " & Err.Description
End Sub

