VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmComunicado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comunicado de Cobrança Judicial"
   ClientHeight    =   4230
   ClientLeft      =   12075
   ClientTop       =   3690
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4230
   ScaleWidth      =   6720
   Begin VB.Frame Frame1 
      Caption         =   "Execuções Fiscais"
      ForeColor       =   &H00000080&
      Height          =   2670
      Left            =   180
      TabIndex        =   5
      Top             =   945
      Width           =   6225
      Begin VB.ListBox mainList 
         Appearance      =   0  'Flat
         Height          =   2055
         Left            =   180
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   405
         Width           =   5865
      End
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   960
      MaxLength       =   6
      TabIndex        =   1
      Top             =   150
      Width           =   1065
   End
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   960
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   525
      Width           =   5490
   End
   Begin prjChameleon.chameleonButton btCodigo 
      Height          =   315
      Left            =   2070
      TabIndex        =   4
      ToolTipText     =   "Consulta cadastro"
      Top             =   135
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmComunicado.frx":0000
      PICN            =   "frmComunicado.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton btPrint 
      Height          =   360
      Left            =   5265
      TabIndex        =   7
      ToolTipText     =   "Imprimir Comunicado"
      Top             =   3735
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "Imprimir"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmComunicado.frx":0176
      PICN            =   "frmComunicado.frx":0192
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código......:"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   3
      Top             =   195
      Width           =   855
   End
   Begin VB.Label lblRS 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome........:"
      Height          =   225
      Left            =   90
      TabIndex        =   2
      Top             =   555
      Width           =   855
   End
End
Attribute VB_Name = "frmComunicado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nTipo As Integer

Private Sub btCodigo_Click()
Dim codigo As Long
mainList.Clear
codigo = Val(txtCodigo.Text)
If codigo = 0 Then
    MsgBox "Digite o código!", vbCritical, "Erro"
    Exit Sub
End If

CarregaContribuinte codigo
If nTipo = 0 Then
    MsgBox "Cadastro não localizado!", vbCritical, "Erro"
    Exit Sub
End If

End Sub

Private Sub btPrint_Click()
Dim codigo As Long, x As Integer, bFind As Boolean

codigo = Val(txtCodigo.Text)
If codigo = 0 Then
    MsgBox "Digite o código!", vbCritical, "Erro"
    Exit Sub
End If

If mainList.ListCount = 0 Then
    MsgBox "Nenhuma excução fiscal selecionada!", vbCritical, "Erro"
    Exit Sub
End If

bFind = False
For x = 0 To mainList.ListCount - 1
    If mainList.Selected(x) = True Then
        bFind = True
        Exit For
    End If
Next

If Not bFind Then
    MsgBox "Nenhuma excução fiscal selecionada!", vbCritical, "Erro"
    Exit Sub
End If

frmReport.ShowReport "COMUNICADOJUDICIAL", frmMdi.HWND, Me.HWND

End Sub

Private Sub Form_Load()
Centraliza Me
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
Tweak txtCodigo, KeyAscii, IntegerPositive
End Sub

Private Sub CarregaContribuinte(nCodReduz As Long)
Dim Sql As String, RdoAux As rdoResultset, sNome As String
txtNome.Text = ""
mainList.Clear

If nCodReduz < 100000 Then
    nTipo = 1
ElseIf nCodReduz >= 100000 And nCodReduz < 300000 Then
    nTipo = 2
ElseIf nCodReduz >= 500000 And nCodReduz < 700000 Then
    nTipo = 3
Else
    nTipo = 0
End If

Ocupado
If nTipo = 1 Then
    Sql = "select * from vwfullimovel where codreduzido=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
            nTipo = 0
            GoTo Fim
        Else
            sNome = !Nomecidadao
        End If
       .Close
    End With
ElseIf nTipo = 2 Then
    Sql = "select * from mobiliario where codigomob=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
            nTipo = 0
            GoTo Fim
        Else
            sNome = !RazaoSocial
        End If
       .Close
    End With
ElseIf nTipo = 3 Then
    Sql = "select * from cidadao where codcidadao=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
            nTipo = 0
            GoTo Fim
        Else
            sNome = !Nomecidadao
        End If
       .Close
    End With
End If

Sql = "SELECT DISTINCT processocnj FROM debitoparcela WHERE codreduzido=" & Val(txtCodigo.Text) & " AND processocnj  IS not null ORDER BY processocnj"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        mainList.AddItem !processocnj
       .MoveNext
    Loop
   .Close
End With

Fim:
txtNome.Text = sNome
Liberado

End Sub
