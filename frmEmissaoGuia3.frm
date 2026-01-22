VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmEmissaoGuia3 
   BackColor       =   &H00FBFBE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Composição da guia"
   ClientHeight    =   4050
   ClientLeft      =   8970
   ClientTop       =   7035
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   Begin prjChameleon.chameleonButton cmdTeste 
      Height          =   330
      Left            =   3735
      TabIndex        =   12
      ToolTipText     =   "Imprimir as parcelas"
      Top             =   4230
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      BTYPE           =   3
      TX              =   "TESTE"
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
      MICON           =   "frmEmissaoGuia3.frx":0000
      PICN            =   "frmEmissaoGuia3.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox lstIPTU 
      Appearance      =   0  'Flat
      Height          =   420
      Left            =   8940
      TabIndex        =   9
      Top             =   180
      Width           =   1155
   End
   Begin VB.Frame frDDList 
      BackColor       =   &H00FBFBE3&
      Height          =   375
      Left            =   5490
      TabIndex        =   6
      Top             =   90
      Width           =   1230
      Begin VB.ListBox lstAno 
         Height          =   1635
         Left            =   45
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   405
         Width           =   1140
      End
      Begin prjChameleon.chameleonButton cmdDDList 
         Height          =   240
         Left            =   270
         TabIndex        =   8
         ToolTipText     =   "Exibir Lista"
         Top             =   45
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   423
         BTYPE           =   14
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   14869218
         BCOLO           =   14869218
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmEmissaoGuia3.frx":0176
         PICN            =   "frmEmissaoGuia3.frx":0192
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
   End
   Begin Tributacao.jcFrames PainelTop 
      Height          =   915
      Left            =   30
      Top             =   60
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   1614
      FrameColor      =   12829635
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   ""
      TextColor       =   13579779
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
      ColorFrom       =   16514019
      ColorTo         =   0
      Begin VB.ComboBox cmbLanc 
         Height          =   315
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   3690
      End
      Begin VB.ComboBox cmbTipoGuia 
         Height          =   315
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   90
         Width           =   2745
      End
      Begin prjChameleon.chameleonButton cmdPrint 
         Height          =   330
         Left            =   7260
         TabIndex        =   11
         ToolTipText     =   "Imprimir as parcelas"
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
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
         MICON           =   "frmEmissaoGuia3.frx":02EC
         PICN            =   "frmEmissaoGuia3.frx":0308
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdSelectAll 
         Height          =   330
         Left            =   5445
         TabIndex        =   13
         ToolTipText     =   "Imprimir as parcelas"
         Top             =   495
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         BTYPE           =   3
         TX              =   "Selecionar tudo"
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
         MICON           =   "frmEmissaoGuia3.frx":0462
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Selecionar Anos..:"
         Height          =   240
         Index           =   3
         Left            =   4080
         TabIndex        =   5
         Top             =   165
         Width           =   1320
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lançamento..:"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   3
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de guia..:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   2
         Top             =   150
         Width           =   1095
      End
   End
   Begin Tributacao.jcFrames jcFrames1 
      Height          =   2745
      Left            =   30
      Top             =   960
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   4842
      FrameColor      =   12829635
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   ""
      TextColor       =   13579779
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
      ColorFrom       =   16514019
      ColorTo         =   0
      Begin MSComctlLib.ListView lvDebito 
         Height          =   2595
         Left            =   60
         TabIndex        =   4
         Top             =   60
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   4577
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Ano"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Lanc"
            Object.Width           =   3882
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Seq"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Pc."
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Co."
            Object.Width           =   811
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Vencto."
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Valor"
            Object.Width           =   2187
         EndProperty
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      Caption         =   "Todas as parcelas serão registradas durante a noite. Caso necessite pagar no mesmo dia, optar pela emissão de DAM."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   30
      TabIndex        =   10
      Top             =   3750
      Width           =   8655
   End
End
Attribute VB_Name = "frmEmissaoGuia3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbLanc_Click()
Dim sql As String, RdoAux As rdoResultset

CarregaDebito

End Sub

Private Sub cmbTipoGuia_Click()
lvDebito.ListItems.Clear
If cmbTipoGuia.ListIndex = -1 Then Exit Sub
If cmbTipoGuia.ListIndex = 1 Then
    CarregaLancamento 1
ElseIf cmbTipoGuia.ListIndex = 2 Then
    CarregaLancamento 6
ElseIf cmbTipoGuia.ListIndex = 3 Then
    CarregaLancamento 14
ElseIf cmbTipoGuia.ListIndex = 4 Then
    CarregaLancamento 13
Else
    CarregaLancamento 0
End If

End Sub

Private Sub cmdDDList_Click()
Dim nAno As Integer, x As Integer, y As Integer

If cmdDDList.value = True Then
    frDDList.Height = 2130
Else
    frDDList.Height = 375
    For x = 0 To lstAno.ListCount - 1
        nAno = lstAno.List(x)
        If lstAno.Selected(x) = True Then
            For y = 1 To lvDebito.ListItems.Count
                If lvDebito.ListItems(y).Text = nAno Then
                    lvDebito.ListItems(y).Checked = True
                End If
            Next
        Else
            For y = 1 To lvDebito.ListItems.Count
                If lvDebito.ListItems(y).Text = nAno Then
                    lvDebito.ListItems(y).Checked = False
                End If
            Next
        End If
    Next

End If

End Sub

Private Sub cmdPrint_Click()
Dim nQtdeParc As Integer, x As Integer, nSid As Long, bFind As Boolean

bFind = False
For x = 0 To Forms.Count - 1
    If Forms(x).Name = "frmEmissaoGuia" Then
        bFind = True
    End If
Next

If Not bFind Then
    MsgBox "Você fechou a tela principal da emissão de 2 ª via, você deverá reiniciar novamente a operação.", vbExclamation, "Atenção"
    Unload Me
    Exit Sub
End If


If frmMdi.frTeste.Visible = True Then
    MsgBox "Emissão de 2ª via não liberado na base de testes.", vbCritical, "Erro"
    Exit Sub
End If

nQtdeParc = 0
For x = 1 To lvDebito.ListItems.Count
    If lvDebito.ListItems(x).Checked And Val(lvDebito.ListItems(x).SubItems(3)) > 0 And lvDebito.ListItems(x).SubItems(6) = "NÃO PAGO" Then
        nQtdeParc = nQtdeParc + 1
    End If
Next

If nQtdeParc = 0 Then
    MsgBox "Selecione a(s) parcela(s) à serem impressas.", vbCritical, "Erro"
Else
    nSid = Int(Rnd(100) * 1000000)
    If cmbTipoGuia.ListIndex = 1 Then
        nQtdeParc = 12
    End If
    Grava_Boleto nSid, nQtdeParc
    
    
    
    
    If cmbTipoGuia.ListIndex <> 1 Then
        frmReport.ShowReport3 "FICHACOMPENSACAO", frmMdi.HWND, Me.HWND, nSid
    ElseIf cmbTipoGuia.ListIndex = 1 Then
   '     MsgBox "Emissão de IPTU apenas pelo Site.", vbCritical, "Erro"

        Calculo_IPTU
        frmReport.ShowReport3 "FICHACOMPENSACAO_IPTU", frmMdi.HWND, Me.HWND, nSid
        
    End If
    
    sql = "delete from ficha_compensacao where sid=" & nSid
    cn.Execute sql, rdExecDirect
    Unload Me
End If

End Sub

Private Sub cmdSelectAll_Click()
'If cmdSelectAll.value = True Then
For x = 1 To lvDebito.ListItems.Count
    lvDebito.ListItems(x).Checked = cmdSelectAll.value
Next
End Sub

Private Sub cmdTeste_Click()
If NomeDeLogin <> "SCHWARTZ" Then Exit Sub
EmiteBoletoTeste
End Sub

Private Sub Form_Load()

Me.Top = frmEmissaoGuia.Top + 2500
Me.Left = frmEmissaoGuia.Left + 1000

cmbTipoGuia.AddItem "(Lançamentos diversos)"
'cmbTipoGuia.AddItem "IPTU/ITU"
cmbTipoGuia.AddItem "Taxa de Licença"
cmbTipoGuia.AddItem "ISS Fixo"
cmbTipoGuia.AddItem "Vigilância Sanitária"
cmbTipoGuia.ListIndex = 0
CarregaDebito

End Sub

Private Sub CarregaLancamento(nCodigo As Integer)
Dim sql As String, RdoAux As rdoResultset

cmbLanc.Clear
sql = "select codlancamento, descreduz from lancamento "
If nCodigo = 1 Or nCodigo = 6 Or nCodigo = 13 Or nCodigo = 14 Then
    sql = sql & "where codlancamento=" & nCodigo
Else
    sql = sql & "where codlancamento not in (1,2,3,5,6,8,12,13,14,20,21,30) "
End If
sql = sql & "order by descreduz"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbLanc.AddItem !descreduz
        cmbLanc.ItemData(cmbLanc.NewIndex) = !CodLancamento
       .MoveNext
    Loop
   .Close
End With

If cmbLanc.ListCount > 0 Then cmbLanc.ListIndex = 0

End Sub

Private Sub CarregaDebito()
Dim itmX As ListItem, x As Integer, nLanc As Integer, Achou As Boolean

lstAno.Clear
lvDebito.ListItems.Clear
nLanc = cmbLanc.ItemData(cmbLanc.ListIndex)
If nLanc = 1 Then
    CarregaListaDebitoGeral
End If
If UBound(aListaDebitoGeral) = 0 Then CarregaListaDebitoGeral
For x = 1 To UBound(aListaDebitoGeral)
    With aListaDebitoGeral(x)
        If .nLanc = nLanc Then
            Achou = False
            For t = 0 To lstAno.ListCount - 1
                If lstAno.List(t) = .nAno Then
                    Achou = True
                    Exit For
                End If
            Next
            If Not Achou Then
                If nLanc = 1 Then 'IPTU APENAS ANO ATUAL
                    If .nAno = Year(Now) Then
'                    If CDate(.sVencto) >= CDate(Format(Now, "dd/mm/yyyy")) Then
                        lstAno.AddItem .nAno
                    End If
                End If
            End If
        
            If nLanc = 1 Then 'IPTU APENAS ANO ATUAL
                If .nAno <> Year(Now) Then
                'If CDate(.sVencto) < CDate(Format(Now, "dd/mm/yyyy")) Then
                '    GoTo Proximo
                End If
            End If
                    
                    
                    
            If (CDate(Format(.sVencto, "dd/mm/yyyy")) >= CDate(Format(Now, "dd/mm/yyyy"))) Or nLanc = 1 Then
                Set itmX = lvDebito.ListItems.Add(, CStr(.nAno) & Format(.nLanc, "00") & Format(.nSeq, "00") & Format(.nParc, "00") & Format(.nCompl, "00"), CStr(.nAno))
                itmX.SubItems(1) = Format(.nLanc, "00") & "-" & .sLanc
                itmX.SubItems(2) = Format(.nSeq, "00")
                itmX.SubItems(3) = Format(.nParc, "00")
                itmX.SubItems(4) = Format(.nCompl, "00")
                itmX.SubItems(5) = .sVencto
                itmX.SubItems(6) = .sSituacao
                itmX.SubItems(7) = Format(.nValorAtual, "#0.00")
            End If
        End If
    End With
    
Proximo:
Next

End Sub

Private Sub Calculo_IPTU()
Dim qd As New rdoQuery, sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, nValorFinal As Double, nQtdeParc As Integer, nCod As Long
Dim nAno As Integer, x As Integer
lstIPTU.Clear
Set qd.ActiveConnection = cn
nAno = Year(Now)
For x = 1 To lvDebito.ListItems.Count
    If lvDebito.ListItems(x).Checked Then
        nAno = Val(lvDebito.ListItems(x).Text)
        Exit For
    End If
Next

nCod = Val(frmEmissaoGuia.txtCodigo.Text)
qd.sql = "{ Call spCalculo(?,?) }"
qd(0) = nCod
qd(1) = nAno
Set RdoAux = qd.OpenResultset(rdOpenKeyset)
With RdoAux
    lstIPTU.AddItem Format(nCod, "000000")
    lstIPTU.AddItem FormatNumber(!FRACAO, 2)
    lstIPTU.AddItem !Natureza
    lstIPTU.AddItem FormatNumber(!AreaTerreno, 2)
    lstIPTU.AddItem FormatNumber(!AreaPredial, 2)
    lstIPTU.AddItem FormatNumber(!TESTADAPRINC, 2)
    lstIPTU.AddItem FormatNumber(!vvt, 2)
    lstIPTU.AddItem FormatNumber(!vvp, 2)
    lstIPTU.AddItem FormatNumber(!vvi, 2)
    lstIPTU.AddItem FormatNumber(!ValorIPTU, 2)
    lstIPTU.AddItem FormatNumber(!valoritu, 2)
    lstIPTU.AddItem FormatNumber(!QtdeParc * !valorparcela, 2)
    lstIPTU.AddItem FormatNumber(!valorunica, 2)
    lstIPTU.AddItem FormatNumber(!valorunica2, 2)
    lstIPTU.AddItem FormatNumber(!valorunica3, 2)
   .Close
End With

End Sub

Private Sub Grava_Boleto(nSid As Long, nQtdeParc As Integer)
Dim sql As String, RdoAux As rdoResultset, x As Integer, nCodReduz As Long, nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer
Dim nCompl As Integer, sDataVencto As String, nValorParc As Double, sNossoNumero As String, sDigitavel As String, sDv As String
Dim dDataBase As String, nFatorVencto As Integer, sQuintoGrupo As String, sBarra As String, sDigitavel2 As String, nNumDoc As Long
Dim sObs As String, sNome As String, sDoc As String, sEndereco As String, sBairro As String, sCidade As String, sCep As String, sInsc As String
Dim sQuadra As String, sLote As String, sUF As String, nPos As Integer, sCampo1 As String
Dim sCampo2 As String, sCampo3 As String, sCampo4 As String, sCampo5 As String, sParcela As String



With frmEmissaoGuia
    sNome = .txtNome.Text
    sEndereco = .txtEndereco.Text
    sBairro = .txtBairro.Text
    sDoc = .txtDoc.Text
    sCidade = .txtCidade.Text
    sUF = .txtUF.Text
    sCep = .txtCep.Text
    sQuadra = .txtQuadra.Text
    sLote = .txtLote.Text
    sInsc = .txtInscricao.Text
End With

sql = "delete from ficha_compensacao where sid=" & nSid
cn.Execute sql, rdExecDirect

nCodReduz = Val(frmEmissaoGuia.txtCodigo.Text)
sql = "select * from laseriptu where codreduzido=" & nCodReduz & " and ano=" & Year(Now)
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    nQtdeParc = RdoAux!QtdeParc
Else
    nQtdeParc = 1
End If

nPos = 1
nQtdeParc = 0
Dim h As Integer
For h = 1 To (lvDebito.ListItems.Count)
    nParc = lvDebito.ListItems(h).SubItems(3)
    If nParc <> 0 And nParc <> 92 And nParc <> 93 Then
        nQtdeParc = nQtdeParc + 1
    End If
Next

If cmbTipoGuia.ListIndex = 1 Then
    nQtdeParc = 12
End If


For x = 1 To lvDebito.ListItems.Count
    If lvDebito.ListItems(x).Checked Then
        With lvDebito.ListItems(x)
            nAno = .Text
            nLanc = Left(.SubItems(1), 2)
            nSeq = .SubItems(2)
            nParc = .SubItems(3)
            nCompl = .SubItems(4)
            sDataVencto = .SubItems(5)
            nValorParc = CDbl(.SubItems(7))
            
            sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
            Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                nNumDoc = !maximo + 1
            End With
            'GRAVA NA TABELA NUMDOCUMENTO
            sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC,emissor,valorguia) VALUES("
            sql = sql & nNumDoc & ",'" & Format(Now, "mm/dd/yyyy") & "',0,0,0,0,'" & NomeDeLogin & " (2ª via)" & "'," & Virg2Ponto(RemovePonto(CStr(nValorParc))) & ")"
            cn.Execute sql, rdExecDirect
            'GRAVA NA TABELA PARCELADOCUMENTO
            sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES("
            sql = sql & nCodReduz & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & "," & nNumDoc & ")"
            cn.Execute sql, rdExecDirect
                                    
            'GRAVA DOCUMENTO PARA REGISTRO
            sql = "insert ficha_compensacao_documento(numero_documento,data_vencimento,valor_documento,nome,cpf,endereco,bairro,cep,cidade,uf) values(" & nNumDoc & ",'"
            sql = sql & Format(sDataVencto, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(nValorParc)) & ",'" & Mask(Left(sNome, 40)) & "','" & RetornaNumero(sDoc) & "','"
            sql = sql & Mask(Left(sEndereco, 40)) & "','" & Mask(Left(sBairro, 15)) & "','" & RetornaNumero(sCep) & "','" & Mask(Left(sCidade, 30)) & "','" & sUF & "')"
            cn.Execute sql, rdExecDirect
                                    
           '**** GERADOR DE CÓDIGO DE BARRAS ********
            sNossoNumero = "2873532"
            dDataBase = "07/10/1997"
            nFatorVencto = CDate(sDataVencto) - CDate(dDataBase)
            
            If CDate(sDataVencto) >= "22/02/2025" Then
                dDataBase = "29/05/2022"
                nFatorVencto = CDate(sDataVencto) - CDate(dDataBase)
            End If
            
            sQuintoGrupo = Format(nFatorVencto, "0000")
            sQuintoGrupo = sQuintoGrupo & Format(RetornaNumero(FormatNumber(nValorParc, 2)), "0000000000")
            sBarra = "0019" & Format(nFatorVencto, "0000") & Format(RetornaNumero(FormatNumber(nValorParc, 2)), "0000000000") & "000000287353200"
            sBarra = sBarra & CStr(nNumDoc) & "17"
            
            sCampo1 = "0019" & Mid(sBarra, 20, 5)
            sDigitavel = sCampo1 & Val(Calculo_DV10(sCampo1))
            sCampo2 = Mid(sBarra, 24, 10)
            sDigitavel = sDigitavel & sCampo2 & Val(Calculo_DV10(sCampo2))
            sCampo3 = Mid(sBarra, 34, 10)
            sDigitavel = sDigitavel & sCampo3 & Val(Calculo_DV10(sCampo3))
            sCampo5 = Format(nFatorVencto, "0000") & Format(RetornaNumero(FormatNumber(nValorParc, 2)), "0000000000")
            sCampo4 = Val(Calculo_DV11(sBarra))
            sDigitavel = sDigitavel & sCampo4 & sCampo5
            sBarra = Left(sBarra, 4) & sCampo4 & Mid(sBarra, 5, Len(sBarra) - 4)
            sDigitavel2 = Left(sDigitavel, 5) & "." & Mid(sDigitavel, 6, 5) & " " & Mid(sDigitavel, 11, 5) & "." & Mid(sDigitavel, 16, 6) & " "
            sDigitavel2 = sDigitavel2 & Mid(sDigitavel, 22, 5) & "." & Mid(sDigitavel, 27, 6) & " " & Mid(sDigitavel, 33, 1) & " " & Right(sDigitavel, 14)
            sBarra = Gera2of5Str(sBarra)
            
            If nParc = 0 Then
                If nCompl = 0 Then
                    sParc = "UNICA 1"
                ElseIf nCompl = 91 Then
                    sParc = "UNICA 2"
                ElseIf nCompl = 92 Then
                    sParc = "UNICA 3"
                End If
            Else
                sParc = Format(nParc, "00") & "/" & Format(nQtdeParc, "00")
            End If
            
            sObs = ""
            sNumDoc = "287353200" & Format(nNumDoc, "00000000")
            sql = "Insert FICHA_COMPENSACAO(SID,SEQ,CODIGO,NOME,CPF,ENDERECO,BAIRRO,CIDADE,CEP,DOCUMENTO,VALOR,VENCIMENTO,PARCELA,DIGITAVEL,CODBARRA,OBS,INSCRICAO,QUADRA,LOTE,UF,PARCELA_DESC) VALUES("
            sql = sql & nSid & "," & x & "," & nCodReduz & ",'" & Left(Mask(sNome), 80) & "','" & RetornaNumero(sDoc) & "','" & Mask(sEndereco) & "','" & Left(Mask(sBairro), 25) & "','"
            sql = sql & Mask(sCidade) & "','" & sCep & "'," & sNumDoc & "," & Virg2Ponto(Format(nValorParc, "#0.00")) & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','"
            'sql = sql & Format(nPos, "00") & "/" & Format(nQtdeParc, "00") & "','" & sDigitavel2 & "','" & Mask(sBarra) & "','" & sObs & "','" & sInsc & "','"
            sql = sql & sParc & "','" & sDigitavel2 & "','" & Mask(sBarra) & "','" & sObs & "','" & sInsc & "','"
            sql = sql & Mask(sQuadra) & "','" & Mask(sLote) & "','" & sUF & "','" & sParc & "')"
            cn.Execute sql, rdExecDirect
                       
            nPos = nPos + 1
            
        End With
    End If
Next

End Sub

Private Sub CarregaListaDebitoGeral()
Dim nCodigo As Long, nEval As Integer, Achou As Boolean, x As Integer, k As Integer
Dim sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, qd As New rdoQuery, bFind As Boolean

Ocupado
nCodigo = Val(frmEmissaoGuia.txtCodigo.Text)
Achou = False
ReDim aListaDebitoGeral(0)

Set qd.ActiveConnection = cn
qd.QueryTimeout = 0
On Error Resume Next
RdoAux.Close
On Error GoTo 0

'qd.Sql = "{ Call spEXTRATO2VIAIPTU(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
qd.sql = "{ Call spEXTRATONAOPAGO(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
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
                If !statuslanc = 3 Or !statuslanc = 42 Or !statuslanc = 43 Then
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
                    aListaDebitoGeral(nEval).nValorTributo = FormatNumber(!VALORTRIBUTO, 2)
                    aListaDebitoGeral(nEval).nValorAtual = FormatNumber(!ValorTotal, 2)
                End If
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
                    aListaDebitoGeral(x).nValorTributo = FormatNumber(aListaDebitoGeral(x).nValorTributo + !VALORTRIBUTO, 2)
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


Private Sub EmiteBoletoTeste()

Dim sql As String, RdoAux As rdoResultset, x As Integer, nCodReduz As Long, nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer
Dim nCompl As Integer, sDataVencto As String, nValorParc As Double, sNossoNumero As String, sDigitavel As String, sDv As String
Dim dDataBase As String, nFatorVencto As Integer, sQuintoGrupo As String, sBarra As String, sDigitavel2 As String, nNumDoc As Long
Dim sObs As String, sNome As String, sDoc As String, sEndereco As String, sBairro As String, sCidade As String, sCep As String, sInsc As String
Dim sQuadra As String, sLote As String, sUF As String, nPos As Integer, sCampo1 As String
Dim sCampo2 As String, sCampo3 As String, sCampo4 As String, sCampo5 As String, sParcela As String, nSid As Long
nSid = Int(Rnd(100) * 1000000)

With frmEmissaoGuia
    sNome = .txtNome.Text
    sEndereco = .txtEndereco.Text
    sBairro = .txtBairro.Text
    sDoc = .txtDoc.Text
    sCidade = .txtCidade.Text
    sUF = .txtUF.Text
    sCep = .txtCep.Text
    sQuadra = .txtQuadra.Text
    sLote = .txtLote.Text
    sInsc = .txtInscricao.Text
End With

sql = "delete from ficha_compensacao where sid=" & nSid
cn.Execute sql, rdExecDirect

nCodReduz = Val(frmEmissaoGuia.txtCodigo.Text)

nPos = 1
nQtdeParc = 1
nNumDoc = 20993830
nValorParc = 10.79
sDataVencto = "15/02/2023"
                                    
'**** GERADOR DE CÓDIGO DE BARRAS ********
sNossoNumero = "2873532"
dDataBase = "07/10/1997"
nFatorVencto = CDate(sDataVencto) - CDate(dDataBase)

If CDate(sDataVencto) >= "22/02/2025" Then
    dDataBase = "29/05/2022"
    nFatorVencto = CDate(sDataVencto) - CDate(dDataBase)
End If

sQuintoGrupo = Format(nFatorVencto, "0000")
sQuintoGrupo = sQuintoGrupo & Format(RetornaNumero(FormatNumber(nValorParc, 2)), "0000000000")
sBarra = "0019" & Format(nFatorVencto, "0000") & Format(RetornaNumero(FormatNumber(nValorParc, 2)), "0000000000") & "000000287353200"
sBarra = sBarra & CStr(nNumDoc) & "17"

sCampo1 = "0019" & Mid(sBarra, 20, 5)
sDigitavel = sCampo1 & Val(Calculo_DV10(sCampo1))
sCampo2 = Mid(sBarra, 24, 10)
sDigitavel = sDigitavel & sCampo2 & Val(Calculo_DV10(sCampo2))
sCampo3 = Mid(sBarra, 34, 10)
sDigitavel = sDigitavel & sCampo3 & Val(Calculo_DV10(sCampo3))
sCampo5 = Format(nFatorVencto, "0000") & Format(RetornaNumero(FormatNumber(nValorParc, 2)), "0000000000")
sCampo4 = Val(Calculo_DV11(sBarra))
sDigitavel = sDigitavel & sCampo4 & sCampo5
sBarra = Left(sBarra, 4) & sCampo4 & Mid(sBarra, 5, Len(sBarra) - 4)
sDigitavel2 = Left(sDigitavel, 5) & "." & Mid(sDigitavel, 6, 5) & " " & Mid(sDigitavel, 11, 5) & "." & Mid(sDigitavel, 16, 6) & " "
sDigitavel2 = sDigitavel2 & Mid(sDigitavel, 22, 5) & "." & Mid(sDigitavel, 27, 6) & " " & Mid(sDigitavel, 33, 1) & " " & Right(sDigitavel, 14)
sBarra = Gera2of5Str(sBarra)

If nParc = 0 Then
    If nCompl = 0 Then
        sParc = "UNICA 1"
    ElseIf nCompl = 91 Then
        sParc = "UNICA 2"
    ElseIf nCompl = 92 Then
        sParc = "UNICA 3"
    End If
Else
    sParc = Format(nParc, "00") & "/" & Format(nQtdeParc, "00")
End If

sObs = ""
sNumDoc = "287353200" & Format(nNumDoc, "00000000")
sql = "Insert FICHA_COMPENSACAO(SID,SEQ,CODIGO,NOME,CPF,ENDERECO,BAIRRO,CIDADE,CEP,DOCUMENTO,VALOR,VENCIMENTO,PARCELA,DIGITAVEL,CODBARRA,OBS,INSCRICAO,QUADRA,LOTE,UF,PARCELA_DESC) VALUES("
sql = sql & nSid & "," & x & "," & nCodReduz & ",'" & Left(Mask(sNome), 80) & "','" & RetornaNumero(sDoc) & "','" & Mask(sEndereco) & "','" & Left(Mask(sBairro), 25) & "','"
sql = sql & Mask(sCidade) & "','" & sCep & "'," & sNumDoc & "," & Virg2Ponto(Format(nValorParc, "#0.00")) & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','"
sql = sql & Format(nPos, "00") & "/" & Format(nQtdeParc, "00") & "','" & sDigitavel2 & "','" & Mask(sBarra) & "','" & sObs & "','" & sInsc & "','"
sql = sql & Mask(sQuadra) & "','" & Mask(sLote) & "','" & sUF & "','" & sParc & "')"
cn.Execute sql, rdExecDirect
           
frmReport.ShowReport3 "FICHACOMPENSACAO", frmMdi.HWND, Me.HWND, nSid
sql = "delete from ficha_compensacao where sid=" & nSid
cn.Execute sql, rdExecDirect

End Sub

