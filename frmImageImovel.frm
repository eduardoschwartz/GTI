VERSION 5.00
Begin VB.Form frmImageImovel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Foto(s) do Imóvel Selecionado"
   ClientHeight    =   7875
   ClientLeft      =   8865
   ClientTop       =   3525
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   11865
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11835
      Begin VB.CommandButton uFoto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1590
         Picture         =   "frmImageImovel.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Próxima Foto"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton pFoto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         Picture         =   "frmImageImovel.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Foto Anterior"
         Top             =   180
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "de"
         Height          =   195
         Left            =   870
         TabIndex        =   5
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblFotoAte 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   195
         Left            =   1140
         TabIndex        =   4
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblFotoDe 
         Caption         =   "0"
         Height          =   195
         Left            =   600
         TabIndex        =   3
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.TextBox txtWait 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1620
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmImageImovel.frx":0294
      Top             =   2340
      Width           =   4215
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   2400
      TabIndex        =   0
      Top             =   3420
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Image Img 
      Height          =   4065
      Left            =   60
      Stretch         =   -1  'True
      Top             =   840
      Width           =   6525
   End
End
Attribute VB_Name = "frmImageImovel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aFoto() As tFoto
Dim nQtdeFoto As Integer

Dim sSubPath As String
Dim nSetor As Integer
Dim nQuadra As Integer
Dim nLote As Integer
Dim nSeq As Integer
Dim aSeq() As Integer, nPos As Integer, nTotal As Integer

Private Sub Form_Activate()
Dim RdoAux As rdoResultset
On Error GoTo Erro

lblFotoDe.Caption = "0"

Dim nTotal As Integer
txtWait.Visible = False
'ConectaBinary
Sql = "select COUNT(*)AS CONTADOR from foto_imovel where codigo=" & Val(Left$(frmCadImob.lblCodReduz.Caption, 7))
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nTotal2 = RdoAux!contador
If nTotal2 = 0 Then
    MsgBox "Não existem fotos para este imóvel."
    lblFotoDe.Caption = "0"
    lblFotoAte.Caption = "0"
Else
    ReDim aSeq(0)
    nPos = 1
    Sql = "select seq from foto_imovel where codigo=" & Val(Left$(frmCadImob.lblCodReduz.Caption, 7))
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    Do Until RdoAux.EOF
        ReDim Preserve aSeq(UBound(aSeq) + 1)
        aSeq(UBound(aSeq)) = RdoAux!Seq
        RdoAux.MoveNext
    Loop
    RdoAux.Close
    lblFotoDe.Caption = "1"
    lblFotoAte.Caption = UBound(aSeq)
    'Carrega_Foto aSeq(nPos)
    CarregaFoto
End If


Exit Sub
Erro:
MsgBox "O diretório de fotos do servidor não está disponível", vbCritical, "Erro"
Resume Next
Unload Me
End Sub

Private Sub Carrega_Foto_Old(Seq As Integer)
Dim mStream As New ADODB.Stream
Dim rst As New ADODB.Recordset
Dim adoConn As New ADODB.Connection

adoConn.CursorLocation = adUseClient
adoConn.Open cnBinary.Connect

rst.Open "Select seq,foto from Foto_imovel where codigo=" & Val(frmCadImob.lblCodReduz.Caption) & " and seq=" & Seq, adoConn, adOpenKeyset, adLockOptimistic
If rst.EOF Then
   MsgBox "Não existem fotos para este imóvel.", vbCritical, "Atenção"
Else
    With mStream
        .Type = adTypeBinary
        .Open
        If Not IsNull(rst("foto")) Then
            .Write rst("foto")
            img.DataField = "foto"
            Set img.DataSource = rst
        End If
    End With

    Set mStream = Nothing
    pbWidthPercentage = img.Width / Me.Width
    pbHeightPercentage = img.Height / Me.Height
    pbAspectRatio = img.Height / img.Width
End If

End Sub


Private Sub Form_DblClick()
frmZoom.show
frmZoom.ZOrder 0
frmZoom.Left = 500
frmZoom.Top = 500
End Sub

Private Sub Form_Resize()

Frame1.Width = Me.Width
img.Width = Me.Width - 50
img.Height = Me.Height - 600
img.Top = 600
img.Left = 0
img.Stretch = True
img.Refresh
    
End Sub

Private Sub pFoto_Click()

On Error GoTo Erro

Dim nFotoAtual As Integer
nFotoAtual = Val(lblFotoDe.Caption)

If nFotoAtual > 1 Then
    nFotoAtual = nFotoAtual - 1
    lblFotoDe.Caption = nFotoAtual
    sPathOrigem = sPathAnexo & "09\" & Format(aFoto(nFotoAtual).Pasta, "00") & "\" & aFoto(nFotoAtual).Arquivo
    img.Picture = LoadPicture(sPathOrigem)
    img.Visible = True
End If

Exit Sub
Erro:
MsgBox "Erro ao carregar a foto do imóvel.", vbCritical, "Atenção"

End Sub

Private Sub uFoto_Click()
On Error GoTo Erro

Dim nFotoAtual As Integer
nFotoAtual = Val(lblFotoDe.Caption)

If nFotoAtual < nQtdeFoto Then
    nFotoAtual = nFotoAtual + 1
    lblFotoDe.Caption = nFotoAtual
    sPathOrigem = sPathAnexo & "09\" & Format(aFoto(nFotoAtual).Pasta, "00") & "\" & aFoto(nFotoAtual).Arquivo
    img.Picture = LoadPicture(sPathOrigem)
    img.Visible = True
End If

Exit Sub
Erro:
MsgBox "Erro ao carregar a foto do imóvel.", vbCritical, "Atenção"

End Sub

Private Sub CarregaFoto()
Dim Sql As String, RdoAux As rdoResultset, nCodReduz As Long, sPath As String
Dim sPathOrigem As String, fso As New FileSystemObject
    
On Error GoTo Erro

ReDim aFoto(0)
nCodReduz = Val(frmCadImob.lblCodReduz.Caption)

Sql = "select * from foto_imovel where codigo=" & nCodReduz & " order by seq"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aFoto(UBound(aFoto) + 1)
        aFoto(UBound(aFoto)).Seq = !Seq
        aFoto(UBound(aFoto)).Pasta = !Pasta
        aFoto(UBound(aFoto)).Arquivo = !Arquivo
       .MoveNext
    Loop
   .Close
End With
    
nQtdeFoto = UBound(aFoto)
If nQtdeFoto = 0 Then
    lblFotoDe.Caption = "0"
    lblFotoAte.Caption = "0"
    img.Visible = False
Else
    lblFotoDe.Caption = "1"
    lblFotoAte.Caption = nQtdeFoto
    
    sPathOrigem = sPathAnexo & "09\" & Format(aFoto(1).Pasta, "00") & "\" & aFoto(1).Arquivo
    img.Picture = LoadPicture(sPathOrigem)
    img.Visible = True
End If
    
Exit Sub
Erro:
MsgBox "Erro ao carregar a foto do imóvel.", vbCritical, "Atenção"
    
End Sub






