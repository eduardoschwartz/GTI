VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmEmissaoGuia 
   BackColor       =   &H00E8F7F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão de guias/2ª via"
   ClientHeight    =   3255
   ClientLeft      =   13680
   ClientTop       =   2385
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin Tributacao.jcFrames jcFrames3 
      Height          =   825
      Left            =   60
      Top             =   60
      Width           =   6980
      _ExtentX        =   12303
      _ExtentY        =   1455
      FrameColor      =   8421504
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   ""
      TextColor       =   128
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   5
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.ComboBox cmbEnd 
         BackColor       =   &H00E8F7F0&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmEmissaoGuia.frx":0000
         Left            =   5250
         List            =   "frmEmissaoGuia.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   90
         Visible         =   0   'False
         Width           =   1650
      End
      Begin VB.TextBox txtDoc 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5250
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   435
         Width           =   1605
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   960
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   435
         Width           =   4185
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   960
         MaxLength       =   6
         TabIndex        =   26
         Top             =   110
         Width           =   1065
      End
      Begin prjChameleon.chameleonButton cmdCnsImovel 
         Height          =   315
         Left            =   2070
         TabIndex        =   33
         ToolTipText     =   "Consulta Imóvel"
         Top             =   90
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
         MICON           =   "frmEmissaoGuia.frx":0026
         PICN            =   "frmEmissaoGuia.frx":0042
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
         Caption         =   "Tipo de endereço..:"
         Height          =   195
         Index           =   5
         Left            =   3750
         TabIndex        =   32
         Top             =   150
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label lblRS 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome........:"
         Height          =   225
         Left            =   90
         TabIndex        =   28
         Top             =   465
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código......:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   27
         Top             =   150
         Width           =   855
      End
   End
   Begin Tributacao.jcFrames jcFrames2 
      Height          =   975
      Left            =   60
      Top             =   2220
      Width           =   6980
      _ExtentX        =   12303
      _ExtentY        =   1720
      FrameColor      =   8421504
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Endereço de entrega"
      TextColor       =   128
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.TextBox txtUFEnt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6420
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   600
         Width           =   405
      End
      Begin VB.TextBox txtCidadeEnt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3990
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   600
         Width           =   1875
      End
      Begin VB.TextBox txtCepEnt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5940
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   270
         Width           =   885
      End
      Begin VB.TextBox txtBairroEnt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   960
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtEnderecoEnt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   960
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   270
         Width           =   4395
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "UF.:"
         Height          =   225
         Index           =   13
         Left            =   6060
         TabIndex        =   20
         Top             =   630
         Width           =   345
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade..:"
         Height          =   225
         Index           =   12
         Left            =   3270
         TabIndex        =   19
         Top             =   630
         Width           =   690
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço..:"
         Height          =   225
         Index           =   10
         Left            =   60
         TabIndex        =   18
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cep..:"
         Height          =   225
         Index           =   8
         Left            =   5460
         TabIndex        =   17
         Top             =   300
         Width           =   480
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro........:"
         Height          =   225
         Index           =   7
         Left            =   60
         TabIndex        =   16
         Top             =   630
         Width           =   870
      End
   End
   Begin Tributacao.jcFrames jcFrames1 
      Height          =   1305
      Left            =   60
      Top             =   900
      Width           =   6980
      _ExtentX        =   12303
      _ExtentY        =   2302
      FrameColor      =   8421504
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Endereço de localização"
      TextColor       =   128
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   1
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.TextBox txtInscricao 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   960
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   930
         Width           =   2175
      End
      Begin VB.TextBox txtLote 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5730
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   930
         Width           =   1080
      End
      Begin VB.TextBox txtUF 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6420
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   600
         Width           =   405
      End
      Begin VB.TextBox txtCidade 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3990
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   600
         Width           =   1875
      End
      Begin VB.TextBox txtQuadra 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3990
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   930
         Width           =   1080
      End
      Begin VB.TextBox txtCep 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         Height          =   285
         Left            =   5940
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   270
         Width           =   885
      End
      Begin VB.TextBox txtBairro 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   960
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtEndereco 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   960
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   270
         Width           =   4395
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inscrição...:"
         Height          =   225
         Index           =   14
         Left            =   60
         TabIndex        =   14
         Top             =   990
         Width           =   870
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lote..:"
         Height          =   225
         Index           =   4
         Left            =   5190
         TabIndex        =   6
         Top             =   990
         Width           =   510
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Quadra..:"
         Height          =   225
         Index           =   3
         Left            =   3270
         TabIndex        =   5
         Top             =   990
         Width           =   690
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "UF.:"
         Height          =   225
         Index           =   2
         Left            =   6060
         TabIndex        =   4
         Top             =   645
         Width           =   345
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade..:"
         Height          =   225
         Index           =   1
         Left            =   3270
         TabIndex        =   3
         Top             =   645
         Width           =   690
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço..:"
         Height          =   225
         Index           =   6
         Left            =   60
         TabIndex        =   2
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cep..:"
         Height          =   225
         Index           =   9
         Left            =   5460
         TabIndex        =   1
         Top             =   300
         Width           =   510
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro........:"
         Height          =   225
         Index           =   11
         Left            =   60
         TabIndex        =   0
         Top             =   650
         Width           =   870
      End
   End
   Begin VB.Menu mnuGuia 
      Caption         =   "MenuGuia"
      Visible         =   0   'False
      Begin VB.Menu mnuMobiliario 
         Caption         =   "Mobiliário"
      End
      Begin VB.Menu mnuImovel 
         Caption         =   "Imobiliário"
      End
      Begin VB.Menu mnuCidadao 
         Caption         =   "Cidadão"
      End
   End
End
Attribute VB_Name = "frmEmissaoGuia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nTipo As Integer

Private Sub cmdCnsImovel_Click()
PopupMenu mnuGuia
End Sub

Private Sub Form_Activate()

If Val(CodImovel) > 0 Then
     txtCodigo.Text = Val(Left$(CodImovel, 7))
     CarregaContribuinte Val(CodImovel)
     CodImovel = 0
Else
    If Val(CodEmpresa) > 0 Then
         txtCodigo.Text = Val(Left$(CodEmpresa, 7))
         CarregaContribuinte Val(CodEmpresa)
         CodEmpresa = 0
    Else
        If Val(CodCidadao) > 0 Then
             Unload frmCnsCidadao
             If cGetInputState() <> 0 Then DoEvents
             txtCodigo.Text = Val(CodCidadao)
             CarregaContribuinte Val(CodCidadao)
             CodCidadao = 0
        End If
    End If
End If

End Sub

Private Sub Form_Load()
Centraliza Me
Me.Top = Me.Top - 2000
Me.Left = Me.Left - 2000
nTipo = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set xImovel = Nothing
Set m_cMenuContrib = Nothing

End Sub

Private Sub txtCodigo_Change()
Limpa
Unload frmEmissaoGuia2
Unload frmEmissaoGuia3
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
Dim nCodigo As Long
nCodigo = Val(txtCodigo.Text)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    If nCodigo > 0 Then
        CarregaContribuinte nCodigo
    End If
Else
    Tweak txtCodigo, KeyAscii, IntegerPositive
End If

End Sub

Private Sub Limpa()

txtNome.Text = ""
txtDoc.Text = ""
txtEndereco.Text = ""
txtCep.Text = ""
txtBairro.Text = ""
txtCidade.Text = ""
txtUF.Text = ""
txtInscricao.Text = ""
txtQuadra.Text = ""
txtLote.Text = ""
txtEnderecoEnt.Text = ""
txtCepEnt.Text = ""
txtBairroEnt.Text = ""
txtCidadeEnt.Text = ""
txtUFEnt.Text = ""
cmbEnd.Enabled = False
cmbEnd.BackColor = &HE8F7F0
cmbEnd.ListIndex = -1
End Sub

Private Sub CarregaContribuinte(nCodReduz As Long)
Dim Sql As String, RdoAux As rdoResultset, sNome As String, sDoc As String, sLote As String, sQuadra As String, sInscricao As String, tTipoEnd As SeqEndereco
Dim xImovel As clsImovel, sEndereco As String, nNum As Integer, sComplemento As String, sBairro As String, sCidade As String, sUF As String, sCep As String
Dim sEnderecoEnt As String, nNumEnt As Integer, sComplementoEnt As String, sBairroEnt As String, sCidadeEnt As String, sUFEnt As String, sCEPEnt As String, nTipoEnd As Integer
Set xImovel = New clsImovel

If nCodReduz < 100000 Then
    nTipo = 1
    tTipoEnd = Imobiliario
ElseIf nCodReduz >= 100000 And nCodReduz < 300000 Then
    nTipo = 2
    tTipoEnd = Mobiliario
ElseIf nCodReduz >= 500000 And nCodReduz < 700000 Then
    nTipo = 3
    tTipoEnd = cidadao
Else
    nTipo = 0
End If
sDoc = ""

Ocupado
If tTipoEnd = Imobiliario Then
    Sql = "select * from vwfullimovel where codreduzido=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
            GoTo SemCadastro
        Else
            sNome = !nomecidadao
            sDoc = SubNull(!Cnpj)
            If sDoc = "" Then
                sDoc = Format(Val(SubNull(!CPF)), "00000000000")
            End If
            sQuadra = SubNull(!Li_Quadras)
            sLote = SubNull(!Li_Lotes)
            sInscricao = !Inscricao
            nTipoEnd = !Ee_TipoEnd
        End If
       .Close
    End With
ElseIf tTipoEnd = Mobiliario Then
    Sql = "select * from mobiliario where codigomob=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
            GoTo SemCadastro
        Else
            sNome = !razaosocial
            sDoc = SubNull(!Cnpj)
            If sDoc = "" Then
                sDoc = Format(Val(SubNull(!CPF)), "00000000000")
            End If
            sQuadra = "": sLote = ""
            sInscricao = SubNull(!inscestadual)
        
        End If
       .Close
    End With
ElseIf tTipoEnd = cidadao Then
    Sql = "select * from cidadao where codcidadao=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
            GoTo SemCadastro
        Else
            sNome = !nomecidadao
            sDoc = SubNull(!Cnpj)
            If sDoc = "" Then
                sDoc = Format(Val(SubNull(!CPF)), "00000000000")
            End If
            sQuadra = "": sLote = ""
            sInscricao = ""
            
        End If
       .Close
    End With
Else
    GoTo SemCadastro
End If


xImovel.RetornaEndereco nCodReduz, tTipoEnd, Localizacao
sEndereco = xImovel.Endereco
nNum = Val(xImovel.Numero)
sComplemento = xImovel.Complemento
sBairro = xImovel.Bairro
sCidade = xImovel.Cidade
sUF = xImovel.UF
sCep = xImovel.Cep

If tTipoEnd = Imobiliario Then
    If nTipoEnd = 0 Then
        xImovel.RetornaEndereco nCodReduz, Imobiliario, Localizacao
    ElseIf nTipoEnd = 1 Then
        xImovel.RetornaEndereco nCodReduz, Imobiliario, cadastrocidadao
    Else
        xImovel.RetornaEndereco nCodReduz, Imobiliario, Entrega
    End If
Else
    If tTipoEnd = Mobiliario Then
        xImovel.RetornaEndereco nCodReduz, tTipoEnd, Entrega
    Else
        sEnderecoEnt = sEndereco
        nNumEnt = nNum
        sComplementoEnt = sComplemento
        sBairroEnt = sBairro
        sCidadeEnt = sCidade
        sUFEnt = sUF
        sCEPEnt = sCep
        GoTo Continua
    End If
End If
sEnderecoEnt = xImovel.Endereco
nNumEnt = Val(xImovel.Numero)
sComplementoEnt = xImovel.Complemento
sBairroEnt = xImovel.Bairro
sCidadeEnt = xImovel.Cidade
sUFEnt = xImovel.UF
sCEPEnt = xImovel.Cep

Continua:
txtNome.Text = sNome
If Len(sDoc) = 11 Then
    sDoc = Format(sDoc, "000\.000\.000-00")
ElseIf Len(sDoc) = 14 Then
    sDoc = Format(sDoc, "00\.000\.000/0000-00")
End If
txtDoc.Text = sDoc
txtQuadra.Text = sQuadra
txtLote.Text = sLote
txtInscricao.Text = sInscricao

txtEndereco.Text = sEndereco & ", " & nNum & " " & sComplemento
txtBairro.Text = sBairro
txtCidade.Text = sCidade
txtUF.Text = sUF
txtCep.Text = sCep

txtEnderecoEnt.Text = sEnderecoEnt & ", " & nNumEnt & " " & sComplementoEnt
txtBairroEnt.Text = sBairroEnt
txtCidadeEnt.Text = sCidadeEnt
txtUFEnt.Text = sUFEnt
txtCepEnt.Text = sCEPEnt

Liberado

If Trim(txtDoc.Text) = "" Or Val(RetornaNumero(txtDoc.Text)) = 0 Then
    MsgBox "CPF/CNPJ obrigatório para emissão de guia.", vbCritical, "Erro"
    Exit Sub
End If

If sEndereco = "" Then
    MsgBox "Endereço obrigatório para emissão de guia.", vbCritical, "Erro"
    Exit Sub
End If

If FlagForm = 1 Then
    frmEmissaoGuia2.show vbModeless
Else
    CarregaListaDebitoGeral
    If UBound(aListaDebitoGeral) = 0 Then
        MsgBox "Não existem lançamentos não pagos a serem emitidos.", vbCritical, "Atenção"
    Else
        frmEmissaoGuia3.show
    End If
End If

Exit Sub
SemCadastro:
Liberado
MsgBox "Inscrição não cadastrada.", vbCritical, "Erro"


End Sub

Private Sub txtCodigo_KeyUp(KeyCode As Integer, Shift As Integer)
Dim nCodigo As Long
nCodigo = Val(txtCodigo.Text)
If nCodigo >= 500000 Then
    cmbEnd.Enabled = True
    cmbEnd.BackColor = Branco
    cmbEnd.ListIndex = 0
Else
    cmbEnd.ListIndex = -1
    cmbEnd.Enabled = False
    cmbEnd.BackColor = &HE8F7F0
End If

End Sub

Private Sub mnuCidadao_Click()
    Set frm = frmCnsCidadao
    frm.sForm = "frmEmissaoGuia"
    frm.show
    frm.ZOrder 0
End Sub

Private Sub mnuImovel_Click()
    sForm = "EG"
    frmCnsImovel.show
    frmCnsImovel.ZOrder 0
End Sub

Private Sub mnuMobiliario_Click()
    sFormMob = "EG2"
    frmCnsMob.show
    frmCnsMob.ZOrder 0
End Sub

Private Sub CarregaListaDebitoGeral()
Dim nCodigo As Long, nEval As Integer, Achou As Boolean, x As Integer, k As Integer
Dim Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, qd As New rdoQuery, bFind As Boolean

Ocupado
nCodigo = Val(txtCodigo.Text)
Achou = False
ReDim aListaDebitoGeral(0)

Set qd.ActiveConnection = cn
qd.QueryTimeout = 0
On Error Resume Next
RdoAux.Close
On Error GoTo 0
qd.Sql = "{ Call spEXTRATONAOPAGO(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
qd(0) = nCodigo
qd(1) = nCodigo
qd(2) = 1950: qd(3) = 2050
qd(4) = 0: qd(5) = 99
qd(6) = 0: qd(7) = 9999
qd(8) = 0: qd(9) = 999
qd(10) = 0: qd(11) = 99
qd(12) = 0: qd(13) = 99
qd(14) = Format(CDate(Right$(frmMdi.Sbar.Panels(6).Text, 10)), "mm/dd/yyyy")
qd(15) = NomeDeLogin
Set RdoAux = qd.OpenResultset(rdOpenKeyset)
With RdoAux
    If .RowCount > 0 Then
'        ReDim Preserve aListaDebitoGeral(UBound(aListaDebitoGeral) + 1)
        nEval = UBound(aListaDebitoGeral)
        
        Do Until .EOF
            nEval = UBound(aListaDebitoGeral)
            Achou = False
            For x = 1 To nEval
                If aListaDebitoGeral(x).nAno = !AnoExercicio And aListaDebitoGeral(x).nLanc = !CodLancamento And aListaDebitoGeral(x).nSeq = !SeqLancamento And _
                   aListaDebitoGeral(x).nParc = !NumParcela And aListaDebitoGeral(x).nCompl = !CODCOMPLEMENTO Then
                   Achou = True
                   Exit For
                End If
            Next
            
            If Not Achou Then
                ReDim Preserve aListaDebitoGeral(UBound(aListaDebitoGeral) + 1)
                nEval = UBound(aListaDebitoGeral)
                aListaDebitoGeral(nEval).nAno = !AnoExercicio
                aListaDebitoGeral(nEval).nLanc = !CodLancamento
                aListaDebitoGeral(nEval).sLanc = !DESCLANCAMENTO
                aListaDebitoGeral(nEval).nSeq = !SeqLancamento
                aListaDebitoGeral(nEval).nParc = !NumParcela
                aListaDebitoGeral(nEval).nCompl = !CODCOMPLEMENTO
                aListaDebitoGeral(nEval).nSituacao = !statuslanc
                aListaDebitoGeral(nEval).sSituacao = !Situacao
                aListaDebitoGeral(nEval).sVencto = Format(!DataVencimento, "dd/mm/yyyy")
                aListaDebitoGeral(nEval).sDA = IIf(IsNull(!datainscricao), "N", "S")
                aListaDebitoGeral(nEval).sAj = IIf(IsNull(!dataajuiza), "N", "S")
                aListaDebitoGeral(nEval).nCodTributo = !CodTributo
                aListaDebitoGeral(nEval).nValorTributo = FormatNumber(!ValorTributo, 2)
                aListaDebitoGeral(nEval).nValorAtual = FormatNumber(!ValorTotal, 2)
            Else
                bFind = False
                For k = 1 To UBound(aListaDebitoGeral)
                    If aListaDebitoGeral(k).nAno = !AnoExercicio And aListaDebitoGeral(k).nLanc = !CodLancamento And _
                       aListaDebitoGeral(k).nSeq = !SeqLancamento And aListaDebitoGeral(k).nParc = !NumParcela And _
                       aListaDebitoGeral(k).nCompl = !CODCOMPLEMENTO And aListaDebitoGeral(k).nCodTributo = !CodTributo Then
                       bFind = True
                       Exit For
                    End If
                Next
                
                If Not bFind Then
                    aListaDebitoGeral(x).nValorTributo = FormatNumber(aListaDebitoGeral(x).nValorTributo + !ValorTributo, 2)
                    aListaDebitoGeral(x).nValorAtual = FormatNumber(aListaDebitoGeral(x).nValorAtual + !ValorTotal, 2)
                End If
            End If
Proximo:
            
           .MoveNext
        Loop
    End If
   .Close
End With
Liberado

End Sub


