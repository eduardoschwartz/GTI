VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmServico 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerenciador de Serviços do G.T.I."
   ClientHeight    =   2100
   ClientLeft      =   7680
   ClientTop       =   6105
   ClientWidth     =   5280
   Icon            =   "frmServico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   WindowState     =   1  'Minimized
   Begin Tributacao.XP_ProgressBar PBar 
      Height          =   195
      Left            =   315
      TabIndex        =   7
      Top             =   1710
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   16750899
   End
   Begin VB.TextBox txtMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   315
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmServico.frx":08CA
      Top             =   990
      Width           =   3615
   End
   Begin prjChameleon.chameleonButton btVerificar 
      Height          =   330
      Left            =   4185
      TabIndex        =   0
      Top             =   1620
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   582
      BTYPE           =   14
      TX              =   "Verificar"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmServico.frx":08F2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblV 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   3825
      TabIndex        =   6
      Top             =   180
      Width           =   1230
   End
   Begin VB.Label lblNext 
      BackColor       =   &H00FFFFFF&
      Caption         =   "00:00:00"
      ForeColor       =   &H00000040&
      Height          =   240
      Left            =   2205
      TabIndex        =   5
      Top             =   495
      Width           =   1590
   End
   Begin VB.Label lblLast 
      BackColor       =   &H00FFFFFF&
      Caption         =   "00:00:00"
      ForeColor       =   &H00000040&
      Height          =   240
      Left            =   2205
      TabIndex        =   4
      Top             =   180
      Width           =   1590
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Próxima verificação..:"
      Height          =   195
      Index           =   1
      Left            =   405
      TabIndex        =   3
      Top             =   495
      Width           =   1680
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Última verificação.....:"
      Height          =   195
      Index           =   0
      Left            =   405
      TabIndex        =   2
      Top             =   180
      Width           =   1680
   End
   Begin VB.Image Image 
      Height          =   1005
      Left            =   3870
      Picture         =   "frmServico.frx":090E
      Top             =   405
      Width           =   1350
   End
End
Attribute VB_Name = "frmServico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim sql As String, RdoAux As rdoResultset
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal HWND As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_QUIT As Long = &H12
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA = 48
Private Const MsgDefault = "Pronto para atualização dos sistemas."
Private Type NOTIFYICONDATA
   cbSize As Long
   HWND As Long
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type
   
Private Const NIM_ADD = 0
Private Const NIM_MODIFY = 1
Private Const NIM_DELETE = 2
Private Const NIF_MESSAGE = 1
Private Const NIF_ICON = 2
Private Const NIF_TIP = 4
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
    
Private Type RegistroProcessado
    Cnae As String
    CodigoEmpresa As Long
    CodigoSocio As Long
    Existe As Boolean
    Novo As Boolean
End Type
    
Private Sub CallPb(nVal As Long, nTot As Long)
If nVal > 0 Then
    PBar.Color = &HC0C000
Else
    PBar.Color = vbWhite
End If
If ((nVal * 100) / nTot) <= 100 Then
   PBar.value = (nVal * 100) / nTot
Else
   PBar.value = 100
End If

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

End Sub
    
Private Function CheckUpdate() As Boolean
Dim sql As String, RdoAux As rdoResultset

'Sql = "select last,lock,version from eicon_timer"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'If RdoAux!version > App.Revision Then
'    End
'End If
'RdoAux.Close

sql = "select last,lock,version from eicon_timer"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
'If RdoAux!lock = True Or RdoAux!version > App.Revision Then
If RdoAux!lock = True Then
    CheckUpdate = True
Else
    CheckUpdate = False
End If
RdoAux.Close

End Function

Private Sub Form_Initialize()
Dim sql As String
lMajor = App.Major
lMinor = App.Minor
lBuild = App.Revision
lblV.Caption = "Versão: " & lMajor & "." & lMinor & "." & lBuild
If App.PrevInstance Then
    Dim handle As Long
    handle = FindWindow("GTI_Server.exe", vbNullString)
    If handle Then
        PostMessage handle, WM_QUIT, 0&, 0&
    End If
    ' MsgBox "Já existe uma cópia deste programa rodando.", vbCritical, "ATENÇÃO"
     End
End If

On Error Resume Next
If NomeDeLogin = "SCHWARTZ" Then
    Exit Sub
End If


Me.show
Me.Refresh
Conecta UL, UP
'Sql = "update machines set gti_server_version='" & App.Major & "." & App.Minor & "." & App.Revision & "' where computer='" & NomeDoComputador & "'"
'cn.Execute Sql, rdExecDirect

'Timer.Interval = 60000 '1 min
'Timer.Enabled = True
'With nId
'    .cbSize = Len(nId)
'    .hwnd = Me.hwnd
'    .uID = vbNull
'    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
'    .uCallbackMessage = WM_MOUSEMOVE
'    .hIcon = Me.Icon
'    .szTip = "Gerenciador de Serviços do G.T.I." & vbNullChar
'End With
'Shell_NotifyIcon NIM_ADD, nId

'If NomeDoComputador <> "GTI" Then
'    btVerificar.Enabled = False
'End If

sql = "select last,next from eicon_timer"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount = 0 Then
    sql = "insert eicon_timer(last,next,computer,lock) values('" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "','" & Format(DateAdd("n", 10, Now), "mm/dd/yyyy hh:mm:ss") & "','" & NomeDoComputador & "'," & 0 & ")"
    cn.Execute sql, rdExecDirect
Else
    If RdoAux!Next < Now Then
        lblNext.Caption = DateAdd("n", 10, Now)
        sql = "update eicon_timer set lock=0,last='" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "',next='" & Format(DateAdd("n", 10, Now), "mm/dd/yyyy hh:mm:ss") & "'"
        cn.Execute sql, rdExecDirect
        lblLast.Caption = Format(RdoAux!Last, "dd/mm/yyyy hh:mm:ss")
    Else
        lblLast.Caption = Format(RdoAux!Last, "dd/mm/yyyy hh:mm:ss")
        lblNext.Caption = Format(RdoAux!Next, "dd/mm/yyyy hh:mm:ss")
    End If
End If
RdoAux.Close
End Sub

Private Sub Timer_Timer()
Static iMin As Integer
iMin = iMin + 1
If iMin = 20 Then
    iMin = 0
    If CDate(Now) > CDate(lblNext.Caption) Then
'        btVerificar_Click
    End If
End If
End Sub

Private Sub TimerSecond_Timer()
If CDate(Now) > CDate(lblNext.Caption) Then
    'btVerificar_Click
End If
End Sub
   
Public Sub btVerificar_Click()
Dim sql As String, RdoAux As rdoResultset
'Exit Sub
If CheckUpdate Then
    'MsgBox "Atualização em andamento. Aguarde!!!", vbCritical, "ALERTA"
   txtMsg.Text = "Atualização em andamento.Aguarde!!!"
   Exit Sub
End If

'Exit Sub

If NomeDeLogin = "SCHWARTZ" Then Exit Sub

'If NomeDeLogin = "SCHWARTZ" Then
'    AtualizaBaixaGiss
'    Exit Sub
'End If

'Exit Sub
btVerificar.Enabled = False
sql = "update eicon_timer set lock=1,computer='" & NomeDoComputador & "'"
cn.Execute sql, rdExecDirect

Ocupado
txtMsg.Text = "Verificando sistema..."
lblLast.Caption = Now
sql = "select codigo from eicon_empresa order by codigo"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    RdoAux.Close
    AtualizaEmpresa
End If
On Error Resume Next
RdoAux.Close
On Error GoTo 0

sql = "select codigo from eicon_socio order by codigo"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    RdoAux.Close
    AtualizaSocio
End If
On Error Resume Next
RdoAux.Close
On Error GoTo 0

sql = "select codigo from eicon_suspensao order by codigo"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    RdoAux.Close
    AtualizaSuspensao
End If
On Error Resume Next
RdoAux.Close


sql = "select * from periodomei where data_exportacao is null"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    AtualizaMei
End If
RdoAux.Close

sql = "select * from optante_simples where data_exportacao is null"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    AtualizaSN
End If
RdoAux.Close

'Sql = "select * from tb_inter_empr_mei_giss where controle is null"
'Set RdoAux = cnEicon.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'If RdoAux.RowCount > 0 Then
'    RdoAux.Close
    'AtualizaMei
'End If
'On Error Resume Next
'RdoAux.Close
'
'Sql = "select * from tb_inter_empr_snacional_giss where controle is null"
'Set RdoAux = cnEicon.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'If RdoAux.RowCount > 0 Then
'    RdoAux.Close
'    AtualizaSN
'End If
'On Error Resume Next
'RdoAux.Close

ConectaEicon
sql = "select * from tb_inter_empresas_giss where controle is null"
Set RdoAux = cnEicon.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    RdoAux.Close
    AtualizaEmpresaFora
End If
On Error Resume Next

RdoAux.Close

sql = "select * from tb_inter_boletos_giss where controle is null"
Set RdoAux = cnEicon.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    RdoAux.Close
    AtualizaGuias
End If
On Error Resume Next
RdoAux.Close

sql = "select * from tb_inter_bol_descartados_giss where controle is null"
Set RdoAux = cnEicon.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    RdoAux.Close
    AtualizaGuiasCanceladas
End If
On Error Resume Next
RdoAux.Close

'Divida_Ativa:

sql = "SELECT debitoparcela.codreduzido, debitoparcela.anoexercicio, debitoparcela.codlancamento, debitoparcela.seqlancamento, debitoparcela.numparcela, "
sql = sql & "debitoparcela.codcomplemento, debitoparcela.statuslanc, debitoparcela.datavencimento, debitoparcela.datadebase, debitoparcela.codmoeda,"
sql = sql & "debitoparcela.numerolivro, debitoparcela.paginalivro, debitoparcela.numcertidao, debitoparcela.datainscricao, debitoparcela.dataajuiza, debitoparcela.valorjuros,"
sql = sql & "debitoparcela.numprocesso, debitoparcela.intacto, debitoparcela.notificado, debitoparcela.usuario, debitoparcela.numexecfiscal, debitoparcela.anoexecfiscal,"
sql = sql & "debitoparcela.processocnj , debitoparcela.simplesnacional, parceladocumento.NumDocumento FROM debitoparcela INNER JOIN "
sql = sql & "parceladocumento ON debitoparcela.codreduzido = parceladocumento.codreduzido AND debitoparcela.anoexercicio = parceladocumento.anoexercicio AND "
sql = sql & "debitoparcela.codlancamento = parceladocumento.codlancamento AND debitoparcela.seqlancamento = parceladocumento.seqlancamento AND "
sql = sql & "debitoparcela.NumParcela = parceladocumento.NumParcela And debitoparcela.CODCOMPLEMENTO = parceladocumento.CODCOMPLEMENTO "
sql = sql & "WHERE (debitoparcela.codreduzido BETWEEN 100000 AND 300000) AND (debitoparcela.codlancamento = 5) AND (debitoparcela.datainscricao IS NOT NULL) AND "
sql = sql & "(debitoparcela.statuslanc < 5) AND (parceladocumento.numdocumento BETWEEN 2000000 AND 3000000) AND (parceladocumento.numdocumento NOT IN "
sql = sql & "(SELECT num_documento FROM GTI_Eicon.dbo.tb_inter_boletos_cdas))"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    RdoAux.Close
    Atualiza_DividaAtiva
End If
On Error Resume Next
RdoAux.Close


Exit Sub

sql = "SELECT DISTINCT origemreparc.anoproc, origemreparc.numproc FROM origemreparc INNER JOIN parceladocumento ON origemreparc.codreduzido = parceladocumento.codreduzido AND origemreparc.anoexercicio = parceladocumento.anoexercicio AND "
sql = sql & "origemreparc.codlancamento = parceladocumento.codlancamento AND origemreparc.numsequencia = parceladocumento.seqlancamento AND origemreparc.numparcela = parceladocumento.numparcela AND origemreparc.codcomplemento = parceladocumento.codcomplemento INNER JOIN "
sql = sql & "processoreparc ON origemreparc.numproc = processoreparc.numproc AND origemreparc.anoproc = processoreparc.anoproc WHERE (origemreparc.codlancamento = 5) AND (parceladocumento.numdocumento BETWEEN 2000000 AND 3000000) AND (processoreparc.data_exportacao IS NULL) "
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    RdoAux.Close
    Atualiza_Parcelamento
End If
On Error Resume Next
RdoAux.Close


cnEicon.Close

lblNext.Caption = DateAdd("n", 20, Now)
sql = "update eicon_timer set lock=0,last='" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "',next='" & Format(DateAdd("n", 10, Now), "mm/dd/yyyy hh:mm:ss") & "'"
cn.Execute sql, rdExecDirect

btVerificar.Enabled = True

txtMsg.Text = MsgDefault
Liberado
End Sub



Private Sub Form_Resize()
'this is necessary to assure that the minimized window is hidden
If Me.WindowState = vbMinimized Then Me.Hide

End Sub

Private Sub Form_Unload(Cancel As Integer)
'this removes the icon from the system tray
'Shell_NotifyIcon NIM_DELETE, nId
End Sub

Private Sub mnuCancel_Click()
lblNext.Caption = DateAdd("n", 10, Now)
sql = "update eicon_timer set lock=0,last='" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "',next='" & Format(DateAdd("n", 10, Now), "mm/dd/yyyy hh:mm:ss") & "'"
cn.Execute sql, rdExecDirect
End Sub

Private Sub mnuCriaListaAtividade_Click()
Dim sql As String, RdoAux As rdoResultset, nPos As Long, nTot As Long

txtMsg.Text = "Criando lista de atividades."
Ocupado
ConectaEicon2
sql = "truncate table tb_inter_lista_atividades"
cnEicon2.Execute sql, rdExecDirect

sql = "select distinct atividadeiss.codatividade, atividadeiss.descatividade, tabelaiss.aliquota * 100 as aliquota, tabelaiss.data "
sql = sql & "from atividadeiss inner join tabelaiss on atividadeiss.codatividade = tabelaiss.codigoativ where (atividadeiss.codatividade >= 200) and (atividadeiss.imprimir = 1) order by atividadeiss.codatividade, tabelaiss.data"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nPos = 1
    nTot = .RowCount
    Do Until .EOF
        CallPb nPos, nTot
        sql = "insert tb_inter_lista_atividades(cod_atividade,desc_atividade,data_inicio,prestacao_servico,aliquota) "
        sql = sql & "values(" & !codatividade & ",'" & Mask(!descatividade) & "','" & Format(!Data, "mm/dd/yyyy") & "','S'," & Virg2Ponto(Format(!Aliquota, "0.0")) & ")"
        cnEicon2.Execute sql, rdExecDirect
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

txtMsg.Text = MsgDefault
PBar.value = 0
Liberado
cnEicon2.Close
End Sub

Private Sub mPopExit_Click()
'called when user clicks the popup menu Exit command
Unload Me
End Sub



Sub PositionForm()
    Dim WindowRect As RECT
    SystemParametersInfo SPI_GETWORKAREA, 0, WindowRect, 0
    Me.Top = WindowRect.Bottom * Screen.TwipsPerPixelY - Me.Height
    Me.Left = WindowRect.Right * Screen.TwipsPerPixelX - Me.Width
End Sub

Private Sub AtualizaEmpresa()
Dim sql As String, RdoAux As rdoResultset, RdoEmp As rdoResultset, RdoAux2 As rdoResultset, RdoProp As rdoResultset, RdoAux3 As rdoResultset, sCnae As String
Dim nCodigo As Long, sIE As String, sRazao As String, sFantasia As String, sNumProcesso As String, sTipoEmpresa As String, nArea As Double, t As Integer
Dim sDoc As String, sDataAbertura As String, sDataEncerramento As String, sTipoLog As String, sTitLog As String, sNomeLog As String, sRegime As String
Dim nNumImovel As Integer, sCompl As String, sBairro As String, sCep As String, sCidade As String, sUF As String, sFone As String, sFax As String, sEmail As String, sDDD As String
Dim nCodCidadao As Long, sNome As String, nCodLogr As Long, nTipoEnd As String, nPos As Long, nTot As Long, aAtiv() As RegistroProcessado, bFind As Boolean
Dim aSocio() As Long, z As Long, nValor As Double

ConectaEicon
txtMsg.Text = "Atualizando Empresas..."
DoEvents


nPos = 1
sql = "select codigo from eicon_empresa order by codigo"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     nTot = .RowCount
     Do Until .EOF
        CallPb nPos, nTot
        nCodigo = !Codigo
        
        '******* DADOS DA EMPRESA **************************
        
        sql = "select * from vwfullempresa3 where codigomob=" & nCodigo
        Set RdoEmp = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoEmp
            If RdoEmp.RowCount = 0 Then GoTo Proximo
            sIE = RetornaNumero(SubNull(!inscestadual))
            sRazao = !RazaoSocial
            sFantasia = SubNull(!NOMEFANTASIA)
            sNumProcesso = SubNull(!NumProcesso)
            If Len(SubNull(!Cnpj)) = 14 Then
                sTipoEmpresa = "J"
                sDoc = RetornaNumero(!Cnpj)
            Else
                sTipoEmpresa = "F"
                sDoc = RetornaNumero(SubNull(!cpf))
            End If
            If Len(sDoc) < 2 Then sTipoEmpresa = "J"
            sDataAbertura = Format(!DataAbertura, "dd/mm/yyyy")
            If IsNull(!dataencerramento) Then
                sDataEncerramento = ""
            Else
                sDataEncerramento = Format(!dataencerramento, "dd/mm/yyyy")
            End If
            sTipoLog = Trim(SubNull(!AbrevTipoLog))
            sTitLog = Trim(SubNull(!AbrevTitLog))
            If SubNull(Val(!CodLogradouro)) > 0 Then
                sNomeLog = Trim(SubNull(!NomeLogradouro))
            Else
                sNomeLog = Trim(SubNull(!Logradouro))
            End If
            nNumImovel = Val(SubNull(!Numero))
            sCompl = SubNull(!Complemento)
            sBairro = SubNull(!DescBairro)
            If sBairro = "ZONA RURAL" Then
                If Not IsNull(!Cep) Then
                    sCep = Format(!Cep, "00000000")
                Else
                    sCep = "00000000"
                End If
            Else
                sCep = RetornaNumero(RetornaCEP(!CodLogradouro, !Numero))
            End If
            sCidade = SubNull(!descCidade)
            sUF = SubNull(!SiglaUF)
            sDDD = SubNull(!ddd_nf)
            sFone = Left(SubNull(!telefone_nf), 15)
            sFax = Left(RetornaNumero(SubNull(!faxcontato)), 15)
            sEmail = SubNull(!emailcontato)
            
            sql = "select codtributo from mobiliarioatividadeiss where codmobiliario=" & nCodigo
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount > 0 Then
                If RdoAux2!CodTributo = 11 Then
                    sRegime = "F"
                ElseIf RdoAux2!CodTributo = 12 Then
                    sRegime = "E"
                ElseIf RdoAux2!CodTributo = 13 Then
                    sRegime = "V"
                Else
                    sRegime = "A"
                End If
            Else
                sRegime = "A"
            End If
            RdoAux2.Close
            If sRegime = "V" Then sRegime = "A"
            If sRegime = "E" Then sRegime = "T"
            nArea = IIf(IsNull(!areatl), 0, !areatl)
           .Close
           
        End With
        
        sql = "insert tb_inter_empresas(cod_cliente,num_cadastro,timestamp,inscricao,inscricao_estadual,nome_empresa,nome_fantasia,"
        sql = sql & "num_processo,tipo_empresa,cpf_cnpj,data_abertura,data_encerramento,tipo_logradouro,titulo_logradouro,logradouro,"
        sql = sql & "num_imovel,complemento,bairro,cep,cidade,estado,ddd,telefone,fax,email,regime_empresa,status_empresa,classificacao,area_ocupada) "
        sql = sql & "values(2177," & nCodigo & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "'," & nCodigo & "," & IIf(Val(sIE) > 0, Val(sIE), "Null") & ",'" & Left(Mask(sRazao), 120) & "',"
        sql = sql & IIf(sFantasia <> "", "'" & Mask(sFantasia) & "'", "Null") & "," & IIf(sNumProcesso <> "", "'" & sNumProcesso & "'", "Null") & ",'" & sTipoEmpresa & "'," & IIf(Val(sDoc) > 0, Val(sDoc), "Null") & ",'" & Format(sDataAbertura, "m/dd/yyyy") & "',"
        sql = sql & IIf(IsDate(sDataEncerramento), "'" & Format(sDataEncerramento, "mm/dd/yyyy") & "'", "Null") & "," & IIf(sTipoLog <> "", "'" & sTipoLog & "'", "Null") & ","
        sql = sql & IIf(sTitLog <> "", "'" & sTitLog & "'", "Null") & ",'" & Mask(sNomeLog) & "'," & IIf(nNumImovel > 0, "'" & CStr(nNumImovel) & "'", "Null") & "," & IIf(sCompl <> "", "'" & Mask(Left(sCompl, 40)) & "'", "Null") & ",'"
        sql = sql & Mask(sBairro) & "'," & IIf(Val(sCep) > 0, Val(sCep), "Null") & ",'" & sCidade & "','" & sUF & "'," & IIf(Val(sDDD) > 0, Val(sDDD), "Null") & "," & IIf(Val(sFone) > 0, Val(sFone), "Null") & "," & IIf(Val(sFax) > 0, Val(sFax), "Null") & "," & IIf(sEmail <> "", "'" & sEmail & "'", "Null") & ","
        sql = sql & IIf(sRegime <> "", "'" & sRegime & "'", "Null") & ",'" & IIf(IsDate(sDataEncerramento), "E", "A") & "'," & "Null" & "," & RetornaNumero(FormatNumber(nArea, 2)) & ")"
        cnEicon.Execute sql, rdExecDirect
        
        '******* DADOS DOS SÓCIOS **************************
        ReDim aSocio(0)
        
        sql = "SELECT * FROM mobiliarioproprietario Where mobiliarioproprietario.codmobiliario = " & nCodigo
        Set RdoProp = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoProp
            Do Until .EOF
                nCodCidadao = !CodCidadao
                z = BinarySearchLong(aSocio(), nCodCidadao)
                If z = -1 Then
                    ReDim Preserve aSocio(UBound(aSocio) + 1)
                    aSocio(UBound(aSocio)) = nCodCidadao
                End If
               .MoveNext
            Loop
            RdoProp.Close
        End With
                
        sql = "select cod_socio from tb_inter_socios where num_cadastro=" & nCodigo
        Set RdoAux2 = cnEicon.OpenResultset(sql, rdOpenDynamic, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                z = BinarySearchLong(aSocio(), !Cod_socio)
                If z = -1 Then
                    ReDim Preserve aSocio(UBound(aSocio) + 1)
                    aSocio(UBound(aSocio)) = !Cod_socio
                End If
               .MoveNext
            Loop
           .Close
        End With
                
        For t = 1 To UBound(aSocio)
            sql = "select codigo from eicon_socio where codigo=" & aSocio(t)
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount = 0 Then
                sql = "insert eicon_socio(codigo) values(" & aSocio(t) & ")"
                cn.Execute sql, rdExecDirect
                AtualizaSocio
            End If
            RdoAux2.Close
        Next
                
        '****** DADOS DAS ATIVIDADES ************************
        ReDim aAtiv(0)
        sql = "select cod_cliente,cod_atividade from tb_inter_atividades where num_cadastro=" & nCodigo & " and data_fim is null"
        Set RdoAux2 = cnEicon.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                ReDim Preserve aAtiv(UBound(aAtiv) + 1)
                aAtiv(UBound(aAtiv)).Cnae = !cod_atividade
                aAtiv(UBound(aAtiv)).Existe = False
                aAtiv(UBound(aAtiv)).Novo = False
               .MoveNext
            Loop
           .Close
        End With
        
        sCnae = ""
        sql = "select * from mobiliariocnae where codmobiliario=" & nCodigo
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                sCnae = Format(!divisao, "00") & !grupo & Left(Format(!classe, "00"), 1) & "-" & Right$(Format(!classe, "00"), 1) & "/" & Format(!subclasse, "00")
                bFind = False
                For t = 1 To UBound(aAtiv)
                    If sCnae = aAtiv(t).Cnae Then
                        bFind = True
                        Exit For
                    End If
                Next
                If Not bFind Then
                    ReDim Preserve aAtiv(UBound(aAtiv) + 1)
                    aAtiv(UBound(aAtiv)).Cnae = sCnae
                    aAtiv(UBound(aAtiv)).Novo = True
                    aAtiv(UBound(aAtiv)).Existe = False
                Else
                    aAtiv(t).Existe = True
                End If
               .MoveNext
            Loop
           .Close
        End With
        
        For t = 1 To UBound(aAtiv)
            If aAtiv(t).Novo = True Then
                sql = "insert tb_inter_atividades(cod_cliente,num_cadastro,cod_atividade,timestamp,data_inicio) values("
                sql = sql & 2177 & "," & nCodigo & ",'" & aAtiv(t).Cnae & "','" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "','" & Format(Now, "mm/dd/yyyy") & "')"
                cnEicon.Execute sql, rdExecDirect
            End If
            If aAtiv(t).Existe = False And aAtiv(t).Novo = False Then
                sql = "update tb_inter_atividades set data_fim='" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "' where num_cadastro=" & nCodigo & " and cod_atividade='" & aAtiv(t).Cnae & "' and data_fim is null"
                cnEicon.Execute sql, rdExecDirect
            End If
        Next
        
        
        '***** ESTIMATIVA *******
        nValor = 0
        ReDim aAtiv(0)
        sql = "select cod_cliente,cod_atividade from tb_inter_estimativa where num_cadastro=" & nCodigo & " and dt_fim is null"
        Set RdoAux2 = cnEicon.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                ReDim Preserve aAtiv(UBound(aAtiv) + 1)
                aAtiv(UBound(aAtiv)).Cnae = !cod_atividade
                aAtiv(UBound(aAtiv)).Existe = False
                aAtiv(UBound(aAtiv)).Novo = False
               .MoveNext
            Loop
           .Close
        End With
        
        sCnae = ""
        sql = "select * from mobiliariocnae where codmobiliario=" & nCodigo
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                sCnae = Format(!divisao, "00") & !grupo & Left(Format(!classe, "00"), 1) & "-" & Right$(Format(!classe, "00"), 1) & "/" & Format(!subclasse, "00")
                bFind = False
                For t = 1 To UBound(aAtiv)
                    If sCnae = aAtiv(t).Cnae Then
                        bFind = True
                        Exit For
                    End If
                Next
                If Not bFind Then
                    ReDim Preserve aAtiv(UBound(aAtiv) + 1)
                    aAtiv(UBound(aAtiv)).Cnae = sCnae
                    aAtiv(UBound(aAtiv)).Novo = True
                    aAtiv(UBound(aAtiv)).Existe = False
                Else
                    aAtiv(t).Existe = True
                End If
               .MoveNext
            Loop
           .Close
        End With
        
        
        If sRegime = "T" Then
            sql = "select valortributo from debitotributo where codreduzido=" & nCodigo & " and numparcela=1 and codtributo=12 and anoexercicio=" & Year(Now)
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount > 0 Then
                nValor = RdoAux2!VALORTRIBUTO * 12
            End If
        End If
        
        
        
        For t = 1 To UBound(aAtiv)
            If aAtiv(t).Novo = True And sRegime = "T" Then
                sql = "insert tb_inter_estimativa(cod_cliente,num_cadastro,dt_inicio,vlr_estimativa,cod_atividade,timestamp) values("
                sql = sql & 2177 & "," & nCodigo & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "'," & Virg2Ponto(Format(nValor, "#0.00")) & ",'" & aAtiv(t).Cnae & "','" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "')"
                cnEicon.Execute sql, rdExecDirect
            End If
            If aAtiv(t).Existe = False And aAtiv(t).Novo = False Then
                sql = "update tb_inter_atividades set data_fim='" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "' where num_cadastro=" & nCodigo & " and cod_atividade='" & aAtiv(t).Cnae & "' and data_fim is null"
                cnEicon.Execute sql, rdExecDirect
            End If
        Next
        
        
        
Proximo:
        nPos = nPos + 1
        DoEvents
        RdoAux.MoveNext
    Loop
    RdoAux.Close
End With
PBar.value = 0
sql = "delete from eicon_empresa"
cn.Execute sql, rdExecDirect

cnEicon.Close

End Sub

Private Sub AtualizaSocio()

Dim sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, aSocio() As RegistroProcessado, RdoProp As rdoResultset, t As Integer, bFind As Boolean
Dim nCodigo As Long, sTipoLog As String, sTitLog As String, sNomeLog As String, nCodEmpresa As Long, sNome As String, nCodLogr As Long, nTipoEnd As String
Dim nNumImovel As Integer, sCompl As String, sBairro As String, sCep As String, sCidade As String, sUF As String, sFone As String, sFax As String, sEmail As String

ConectaEicon2

sql = "select codigo from eicon_socio order by codigo"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        nCodigo = !Codigo
        ReDim aSocio(0)
        
        sql = "select num_cadastro,cod_socio from tb_inter_socios where cod_socio=" & nCodigo & " and data_fim is null"
        Set RdoProp = cnEicon2.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoProp
            Do Until .EOF
                ReDim Preserve aSocio(UBound(aSocio) + 1)
                aSocio(UBound(aSocio)).CodigoEmpresa = !num_cadastro
                aSocio(UBound(aSocio)).CodigoSocio = nCodigo
                aSocio(UBound(aSocio)).Existe = False
                aSocio(UBound(aSocio)).Novo = False
               .MoveNext
            Loop
           .Close
        End With
        bFind = False
        
        sql = "SELECT codmobiliario,codcidadao from mobiliarioproprietario where codcidadao = " & nCodigo
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                bFind = False
                For t = 1 To UBound(aSocio)
                    If aSocio(t).CodigoEmpresa = !codmobiliario And aSocio(t).CodigoSocio = !CodCidadao Then
                        bFind = True
                        Exit For
                    End If
                Next
                 If Not bFind Then
                    ReDim Preserve aSocio(UBound(aSocio) + 1)
                    aSocio(UBound(aSocio)).CodigoEmpresa = !codmobiliario
                    aSocio(UBound(aSocio)).CodigoSocio = !CodCidadao
                    aSocio(UBound(aSocio)).Novo = True
                    aSocio(UBound(aSocio)).Existe = False
                Else
                    aSocio(t).Existe = True
                End If
               .MoveNext
            Loop
           .Close
        End With
       
        For t = 1 To UBound(aSocio)
            If aSocio(t).Novo = True Then
                sql = "select * from vwfullcidadao where codcidadao=" & aSocio(t).CodigoSocio
                Set RdoProp = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                With RdoProp
                    sNome = !nomecidadao
                    sql = "select num_cadastro from tb_inter_socios where num_cadastro=" & nCodEmpresa & " and nome_socio='" & Mask(sNome) & "' and controle is null"
                    Set RdoAux2 = cnEicon2.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                    If RdoAux2.RowCount > 0 Then
                        RdoAux2.Close
                        GoTo Proximo
                    End If
                    RdoAux2.Close
                    
                    sDoc = RetornaNumero(SubNull(!cpf))
                    
                    If Not IsNull(!etiqueta2) Then
                        sTipoEnd = "C"
                    Else
                        sTipoEnd = "R"
                    End If
                                    
                    If sTipoEnd = "R" Then
                        sTipoLog = SubNull(!AbrevTipoLog)
                        sTitLog = SubNull(!AbrevTitLog)
                        sNomeLog = SubNull(!NomeLogradouro)
                        If sNomeLog = "" Then
                            sNomeLog = SubNull(!NOMELOGRADOURO2)
                        End If
                        nNumImovel = Val(SubNull(!NUMIMOVEL))
                        sCompl = SubNull(!Complemento)
                        sBairro = SubNull(!DescBairro)
                        nCodLogr = Val(SubNull(!CodLogradouro))
                        If nCodLogr > 0 Then
                            sCep = RetornaNumero(RetornaCEP(nCodLogr, nNumImovel))
                        Else
                            sCep = RetornaNumero(SubNull(!Cep))
                        End If
                        sCidade = SubNull(!descCidade)
                        sUF = SubNull(!SiglaUF)
                        sFone = SubNull(!telefone)
                        sEmail = SubNull(!Email)
                    Else
                        sTipoLog = SubNull(!AbrevTipoLogC)
                        sTitLog = SubNull(!AbrevTitLogC)
                        sNomeLog = SubNull(!NomeLogradouroC)
                        nNumImovel = Val(SubNull(!NUMIMOVEL2))
                        sCompl = SubNull(!Complemento2)
                        sBairro = SubNull(!DescBairroC)
                        nCodLogr = Val(SubNull(!CodLogradouro2))
                        If nCodLogr > 0 Then
                            sCep = RetornaNumero(RetornaCEP(nCodLogr, nNumImovel))
                        Else
                            sCep = RetornaNumero(SubNull(!Cep2))
                        End If
                        sCidade = SubNull(!desccidadeC)
                        sUF = SubNull(!SiglaUF2)
                        sFone = SubNull(!Telefone2)
                        sEmail = SubNull(!EMAIL2)
                    End If
                   .Close
                End With
                sql = "insert tb_inter_socios(cod_cliente,num_cadastro,inscricao,cod_socio,nome_socio,timestamp,cpf,data_inicio,tipo_logradouro,titulo_logradouro,logradouro,num_imovel,complemento,bairro,cep,"
                sql = sql & "cidade,estado,telefone,email) "
                sql = sql & "values(2177," & aSocio(t).CodigoEmpresa & "," & aSocio(t).CodigoEmpresa & "," & aSocio(t).CodigoSocio & ",'" & Mask(sNome) & "','" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "'," & IIf(Val(sDoc) > 0, Val(sDoc), "Null") & ",'" & Format(Now, "mm/dd/yyyy") & "',"
                sql = sql & IIf(sTipoLog <> "", "'" & Trim(sTipoLog) & "'", "Null") & "," & IIf(sTitLog <> "", "'" & Trim(sTitLog) & "'", "Null") & ",'" & Mask(sNomeLog) & "'," & IIf(nNumImovel > 0, "'" & CStr(nNumImovel) & "'", "Null") & ","
                sql = sql & IIf(sCompl <> "", "'" & Mask(Left(sCompl, 30)) & "'", "Null") & "," & IIf(sBairro <> "", "'" & sBairro & "'", "Null") & "," & IIf(Val(sCep) > 0, Val(sCep), "Null") & ",'" & sCidade & "','" & sUF & "',"
                sql = sql & IIf(Val(sFone) > 0, Val(RemovePonto(Left(sFone, 12))), "Null") & "," & IIf(sEmail <> "", "'" & sEmail & "'", "Null") & ")"
                cnEicon2.Execute sql, rdExecDirect
            End If
            If aSocio(t).Existe = False And aSocio(t).Novo = False Then
                sql = "update tb_inter_socios set data_fim='" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "' where num_cadastro=" & aSocio(t).CodigoEmpresa & " and cod_socio=" & aSocio(t).CodigoSocio & " and data_fim is null"
                cnEicon2.Execute sql, rdExecDirect
            End If
        Next
       
       
Proximo:
        sql = "delete from eicon_socio where codigo=" & nCodigo
        cn.Execute sql, rdExecDirect
       
        DoEvents
        RdoAux.MoveNext
    Loop
    RdoAux.Close
End With

cnEicon2.Close

End Sub

Private Sub AtualizaSuspensao()

Dim sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, sNumProcesso As String, nNumproc As Long, nAnoproc As Long, nDigito As Integer
Dim nCodigo As Long, sDataIni As String, sDataFim As String, nPos As Integer

ConectaEicon2

sql = "select codigo from eicon_suspensao order by codigo"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        nCodigo = !Codigo
        sql = "select num_cadastro from tb_inter_empr_suspensa where num_cadastro=" & nCodigo & " and controle is null"
        Set RdoAux2 = cnEicon2.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            RdoAux2.Close
            GoTo Proximo
        End If
        RdoAux2.Close
                
        nPos = 1
        sDataIni = ""
        sDataFim = ""
        
        sql = "SELECT mobiliarioevento.codmobiliario, mobiliarioevento.codtipoevento, mobiliarioevento.dataevento, mobiliarioevento.numprocevento "
        sql = sql & "FROM mobiliarioevento INNER JOIN mobiliario ON mobiliarioevento.codmobiliario = mobiliario.codigomob Where (mobiliarioevento.codtipoevento > 1) And (mobiliario.codcidade = 413) and "
        sql = sql & "mobiliarioevento.codmobiliario=" & nCodigo & " ORDER BY mobiliarioevento.codmobiliario, mobiliarioevento.dataevento DESC"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                If nPos = 3 Then Exit Do
                sNumProcesso = SubNull(!NUMPROCEVENTO)
                If InStr(1, sNumProcesso, "/") = 0 Then
                    sNumProcesso = ""
                    nAnoproc = 0
                    nNumproc = 0
                Else
                    nAnoproc = Val(Mid(Trim$(sNumProcesso), InStr(1, Trim$(sNumProcesso), "/", vbBinaryCompare) + 1, 4))
                    sNumProcesso = Left$(Trim$(sNumProcesso), InStr(1, Trim$(sNumProcesso), "/", vbBinaryCompare) - 1)
                    nNumproc = Val(Left(sNumProcesso, Len(sNumProcesso) - 1))
                    nDigito = Val(Right(sNumProcesso, 1))
                End If
                
                If nPos = 1 Then
                    If !CODTIPOEVENTO = 2 Then
                        sDataIni = Format(!DATAEVENTO, "dd/mm/yyyy")
                    Else
                        sDataFim = Format(!DATAEVENTO, "dd/mm/yyyy")
                    End If
                ElseIf nPos = 2 Then
                    If !CODTIPOEVENTO = 2 Then
                        sDataIni = Format(!DATAEVENTO, "dd/mm/yyyy")
                    Else
                        sDataFim = Format(!DATAEVENTO, "dd/mm/yyyy")
                    End If
                End If
                nPos = nPos + 1
               .MoveNext
            Loop
           .Close
        
        End With
        
        If sDataIni = "" Then sDataIni = Format(Now, "mm/dd/yyyy")
        sql = "insert tb_inter_empr_suspensa(cod_cliente,num_cadastro,timestamp,num_processo,dig_processo,ano_processo,data_suspensao,data_suspensao_fim) "
        sql = sql & "values(2177," & nCodigo & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "'," & IIf(sNumProcesso <> "", "'" & CStr(nNumproc) & "'", "Null") & ","
        sql = sql & IIf(sNumProcesso <> "", "'" & CStr(nDigito) & "'", "Null") & "," & IIf(sNumProcesso <> "", nAnoproc, "Null") & "," & IIf(IsDate(sDataIni), "'" & Format(sDataIni, "mm/dd/yyyy") & "'", "Null") & ","
        sql = sql & IIf(IsDate(sDataFim), "'" & Format(sDataFim, "mm/dd/yyyy") & "'", "Null") & ")"
        cnEicon2.Execute sql, rdExecDirect
        
        sql = "delete from eicon_suspensao where codigo=" & nCodigo
        cn.Execute sql, rdExecDirect
                 
Proximo:
       DoEvents
       RdoAux.MoveNext
    Loop
    RdoAux.Close
End With

cnEicon2.Close

End Sub

Private Sub AtualizaMei()
On Error GoTo Erro
Dim sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset
Dim nCodigo As Long, sDataIni As String, sDataFim As String, nPos As Long, nTot As Long, sCnpj As String

ConectaEicon2

PBar.value = 0
nPos = 0
sql = "SELECT periodomei.id, periodomei.codigo, periodomei.datainicio, periodomei.datafim, periodomei.cnpj_base, periodomei.data_exportacao, mobiliario.cnpj "
sql = sql & "FROM periodomei INNER JOIN mobiliario ON periodomei.codigo = mobiliario.codigomob where data_exportacao is null order by id"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    Do Until .EOF
        If nPos > nTot Then Exit Do
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        nCodigo = !Codigo
        sDataIni = Format(!DataInicio, "dd/mm/yyyy")
        sDataFim = IIf(IsNull(!Datafim), "", Format(!Datafim, "dd/mm/yyyy"))
        sCnpj = !Cnpj_Base
        
        sql = "select * from tb_inter_empr_mei where num_cadastro=" & nCodigo & " and data_inicio='" & Format(sDataIni, "mm/dd/yyyy") & "' and data_final='" & Format(sDataFim, "mm/dd/yyyy")
        Set RdoAux2 = cnEicon.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount = 0 Then
            If sDataFim = "" Then
                sql = "insert tb_inter_empr_mei(cod_cliente,num_cadastro,inscricao,base_cnpj,data_inicio,[ timestamp]) values(2177,"
                sql = sql & nCodigo & "," & nCodigo & ",'" & sCnpj & "','" & Format(sDataIni, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "')"
            Else
                sql = "insert tb_inter_empr_mei(cod_cliente,num_cadastro,inscricao,base_cnpj,data_inicio,data_fim,[ timestamp]) values(2177,"
                sql = sql & nCodigo & "," & nCodigo & ",'" & sCnpj & "','" & Format(sDataIni, "mm/dd/yyyy") & "','" & Format(sDataFim, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "')"
            End If
            cnEicon2.Execute sql, rdExecDirect
        End If
        
        sql = "update periodomei set data_exportacao='" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "' where id=" & RdoAux!id
        cn.Execute sql, rdExecDirect
        nPos = nPos + 1
        DoEvents
        RdoAux.MoveNext
Proximo:
    Loop
    RdoAux.Close
End With




'PBar.value = 0
'nPos = 0
'Sql = "select * from tb_inter_empr_mei_giss where controle is null order by inscricao"
'Set RdoAux = cnEicon2.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    nTot = .RowCount
'    Do Until .EOF
'        If nPos > nTot Then Exit Do
'        If nPos Mod 10 = 0 Then
'            CallPb nPos, nTot
'        End If
'        nCodigo = !Inscricao
'        sDataIni = Format(!Data_Inicio, "mm/dd/yyyy")
'        sDataFim = IIf(IsNull(!data_fim), "", Format(!data_fim, "mm/dd/yyyy"))
'
'        Sql = "select * from mei where codigo=" & nCodigo & " and datainicio='" & Format(sDataIni, "mm/dd/yyyy") & "'"
'        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'        If RdoAux.RowCount = 0 Then
'            Sql = "insert mei(codigo,datainicio,datafim) values(" & nCodigo & ",'" & Format(sDataIni, "mm/dd/yyyy") & "',"
'            Sql = Sql & IIf(IsDate(sDataFim), "'" & Format(sDataFim, "mm/dd/yyyy") & "'", "Null") & ")"
'        Else
'            Sql = "update mei set datafim=" & IIf(IsDate(sDataFim), "'" & Format(sDataFim, "mm/dd/yyyy") & "'", "Null") & " where codigo=" & nCodigo & " and datainicio='" & Format(sDataIni, "mm/dd/yyyy") & "'"
'        End If
'        cn.Execute Sql, rdExecDirect
'
'        Sql = "update tb_inter_empr_mei_giss set controle=1 where inscricao=" & nCodigo
'        cnEicon2.Execute Sql, rdExecDirect
'        nPos = nPos + 1
'        DoEvents
'        RdoAux.MoveNext
'proximo:
'    Loop
'    RdoAux.Close
'End With

cnEicon2.Close
PBar.value = 0
Exit Sub
Erro:
'MsgBox Err.Description
Resume Next
End Sub

Private Sub AtualizaSN()
On Error GoTo Erro
Dim sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset
Dim nCodigo As Long, sDataIni As String, sDataFim As String, nPos As Long, nTot As Long

'Exit Sub

ConectaEicon2

PBar.value = 0
nPos = 0
sql = "SELECT optante_simples.codigo, optante_simples.data_inicio, optante_simples.data_final, optante_simples.cnpj_base, optante_simples.timestamp, optante_simples.data_exportacao, mobiliario.cnpj "
sql = sql & "FROM optante_simples INNER JOIN mobiliario ON optante_simples.codigo = mobiliario.codigomob where data_exportacao is null order by timestamp"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    Do Until .EOF
        If nPos > nTot Then Exit Do
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        nCodigo = !Codigo
        sDataIni = Format(!Data_Inicio, "dd/mm/yyyy")
        sDataFim = IIf(IsNull(!Data_Final), "", Format(!Data_Final, "dd/mm/yyyy"))
        sCnpj = !Cnpj
        
        If sDataFim = "" Then
            sql = "insert tb_inter_empr_snacional(cod_cliente,num_cadastro,inscricao,cnpj,data_inicio,[timestamp]) values(2177,"
            sql = sql & nCodigo & "," & nCodigo & ",'" & sCnpj & "','" & Format(sDataIni, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "')"
        Else
            sql = "insert tb_inter_empr_snacional(cod_cliente,num_cadastro,inscricao,cnpj,data_inicio,data_fim,[timestamp]) values(2177,"
            sql = sql & nCodigo & "," & nCodigo & ",'" & sCnpj & "','" & Format(sDataIni, "mm/dd/yyyy") & "','" & Format(sDataFim, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "')"
        End If
        cnEicon2.Execute sql, rdExecDirect
        
        sql = "update optante_simples set data_exportacao='" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "' "
        sql = sql & "where codigo=" & nCodigo & " and data_inicio='" & Format(sDataIni, "mm/dd/yyyy") & "' and data_final='" & sDataFim & "'"
        cn.Execute sql, rdExecDirect
        nPos = nPos + 1
        DoEvents
        RdoAux.MoveNext
Proximo:
    Loop
    RdoAux.Close
End With



'PBar.value = 0
'nPos = 0
'Sql = "select * from tb_inter_empr_snacional_giss where controle is null and ip<>'1.1.1.1' order by inscricao"
'Set RdoAux = cnEicon2.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    nTot = .RowCount
'    Do Until .EOF
'        If nPos Mod 10 = 0 Then
'            CallPb nPos, nTot
'        End If
'        If !Inscricao > 300000 Then GoTo proximo
'        nCodigo = !Inscricao
'        sDataIni = Format(!Data_Inicio, "mm/dd/yyyy")
'        sDataFim = IIf(IsNull(!data_fim), "", Format(!data_fim, "mm/dd/yyyy"))
'
'        Sql = "insert periodosn(codigo,dataini,datafim) values(" & nCodigo & ",'" & Format(sDataIni, "mm/dd/yyyy") & "',"
'        Sql = Sql & IIf(IsDate(sDataFim), "'" & Format(sDataFim, "mm/dd/yyyy") & "'", "Null") & ")"
'        cn.Execute Sql, rdExecDirect
'
'        Sql = "update tb_inter_empr_snacional_giss set controle=1 where inscricao=" & nCodigo
'        cnEicon2.Execute Sql, rdExecDirect
'proximo:
'        nPos = nPos + 1
'        DoEvents
'        RdoAux.MoveNext
'    Loop
'    RdoAux.Close
'End With

cnEicon2.Close
PBar.value = 0
Exit Sub
Erro:
MsgBox Err.Description
Resume Next
End Sub

Private Sub AtualizaEmpresaFora()
On Error GoTo Erro
Dim sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, sDoc As String, dDateTime As Date, nJuridica As Integer
Dim nCodigo As Long, sDataIni As String, sDataFim As String, nPos As Long, nTot As Long, RdoAux3 As rdoResultset, nMaxBairro As Integer
Dim nMaxCod As Long, sUF As String, nCodCidade As Integer, nCodBairro As Integer, sEndereco As String, nNumero As Integer, sCompl As String
Dim sCep As String, sFone As String, sEmail As String, sBairro As String, sCidade As String, sNome As String, sCPF As String, sCnpj As String
Dim nCodCidadao As Long

ConectaEicon2
PBar.value = 0
nPos = 0
sql = "select * from tb_inter_empresas_giss where controle is null order by cpf_cnpj"
Set RdoAux = cnEicon2.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    Do Until .EOF
        sCPF = "": sCnpj = ""
        nCodigo = !Inscricao
        dDateTime = !TimeStamp
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        sDoc = !cpf_cnpj
        
        nCodCidadao = 0
        If !tipo_empresa = "J" Then
            nJuridica = 1
            sCnpj = sDoc
            sql = "select * from cidadao where cnpj='" & Format(sDoc, "00000000000000") & "'"
        Else
            nJuridica = 0
            sCPF = sDoc
            sql = "select * from cidadao where cpf='" & sDoc & "'"
        End If
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            nCodCidadao = RdoAux2!CodCidadao
        End If
        sNome = !nome_empresa
        sEndereco = UCase(SubNull(!tipo_logradouro) & " " & SubNull(!titulo_logradouro) & " " & SubNull(!Logradouro))
        nNumero = Val(SubNull(!num_imovel))
        sCompl = SubNull(!Complemento)
        sBairro = SubNull(!Bairro)
        sCidade = SubNull(!Cidade)
        sUF = SubNull(!estado)
        sFone = SubNull(!DDD) & SubNull(!telefone)
        sEmail = SubNull(!Email)
        
        nCodCidade = 999
        sql = "select codcidade from cidade where siglauf='" & sUF & "' and desccidade='" & sCidade & "'"
        Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux3.RowCount > 0 Then
            nCodCidade = RdoAux3!CodCidade
        End If
        RdoAux3.Close
        
        nCodBairro = 999
        sql = "select codbairro from bairro where siglauf='" & sUF & "' and codcidade=" & nCodCidade & " and descbairro='" & sBairro & "'"
        Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux3.RowCount > 0 Then
            nCodBairro = RdoAux3!CodBairro
        End If
        RdoAux3.Close
        If nCodBairro = 999 And nCodCidade <> 413 Then
            sql = "select max(codbairro) as maximo from bairro where siglauf='" & sUF & "' and codcidade=" & nCodCidade
            Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If IsNull(RdoAux3!maximo) Then
                nMaxBairro = 1
            Else
                nMaxBairro = RdoAux3!maximo + 1
            End If
            If (nCodCidade <> 999) Then
                sql = "insert bairro(siglauf,codcidade,codbairro,descbairro) values('" & sUF & "'," & nCodCidade & "," & nMaxBairro & ",'" & Mask(UCase(sBairro)) & "')"
                cn.Execute sql, rdExecDirect
                nCodBairro = nMaxBairro
            End If
        End If
        
        
        If sCompl = "null" Then sCompl = ""
        If RdoAux2.RowCount = 0 Then
            sql = "select max(codcidadao) as maximo from cidadao"
            Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            nCodCidadao = RdoAux3!maximo + 1
            RdoAux3.Close
            sql = "INSERT CIDADAO(CODCIDADAO,NOMECIDADAO,CPF,CNPJ,NUMIMOVEL,COMPLEMENTO,CODBAIRRO,CODCIDADE,"
            sql = sql & "SIGLAUF,CEP,TELEFONE,EMAIL,NOMELOGRADOURO,JURIDICA,ETIQUETA) VALUES(" & nCodCidadao & ",'"
            sql = sql & Mask(sNome) & "','" & sCPF & "','" & sCnpj & "'," & nNumero & ",'" & Mask(sCompl) & "'," & nCodBairro & ","
            sql = sql & nCodCidade & ",'" & sUF & "','" & sCep & "','" & sFone & "','" & sEmail & "','" & sEndereco & "'," & nJuridica & ",'" & "S" & "')"
        Else
            sql = "UPDATE CIDADAO SET NUMIMOVEL=" & nNumero & ",COMPLEMENTO='" & Mask(sCompl) & "',CODBAIRRO=" & nCodBairro & ",CODCIDADE=" & nCodCidade & ","
            sql = sql & "SIGLAUF='" & sUF & "',CEP='" & sCep & "',TELEFONE='" & sFone & "',EMAIL='" & sEmail & "',NOMELOGRADOURO='" & sEndereco & "',JURIDICA=" & nJuridica
            sql = sql & " WHERE CODCIDADAO=" & nCodCidadao
        End If
        cn.Execute sql, rdExecDirect
        
        If RdoAux2.RowCount = 0 Then
'            Sql = "insert historicocidadao(codigo,data,usuario,obs) values(" & nCodCidadao & ",'" & Format(Now, sDataFormat & " hh:mm:ss") & "','" & NomeDeLogin & "','"
'            Sql = Sql & "Cidadão criado através da Tela de Serviço - Atualiza Empresas de Fora')"
            sql = "insert historicocidadao(codigo,data,userid,obs) values(" & nCodCidadao & ",'" & Format(Now, sDataFormat & " hh:mm:ss") & "'," & RetornaUsuarioID(NomeDeLogin) & ",'"
            sql = sql & "Cidadão criado através da Tela de Serviço - Atualiza Empresas de Fora')"
            cn.Execute sql, rdExecDirect
        End If
        RdoAux2.Close
        
        sql = "update tb_inter_empresas_giss set inscricao_gti=" & nCodCidadao & ",controle=1 where inscricao=" & nCodigo & " and timestamp='" & Format(dDateTime, "mm/dd/yyyy hh:mm:ss") & "'"
        cnEicon2.Execute sql, rdExecDirect
Proximo:
        nPos = nPos + 1
        DoEvents
        RdoAux.MoveNext
    Loop
    RdoAux.Close
End With

cnEicon2.Close
PBar.value = 0
Exit Sub
Erro:
MsgBox Err.Description
Resume Next
End Sub

Private Sub AtualizaGuias()
On Error GoTo Erro
Dim sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, nAno As Integer
Dim nCodigo As Long, bEmpresa As Boolean, nDoc As Long, nMaxSeq As Integer, sCPFCNPJ As String, nValor As Double

ConectaEicon2

sql = "select * from tb_inter_boletos_giss where controle is null"
Set RdoAux = cnEicon2.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        bEmpresa = False
        nCodigo = !num_cadastro
        nDoc = !num_documento
        nValor = Ponto2Virg(!valor_imposto)
        sql = "select num_cadastro from tb_inter_empresas where num_cadastro=" & nCodigo
        Set RdoAux2 = cnEicon2.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            bEmpresa = True
        End If
        RdoAux2.Close
        If Not bEmpresa Then
            sql = "select num_cadastro,cpf_cnpj from tb_inter_empresas_giss where num_cadastro=" & nCodigo
            Set RdoAux2 = cnEicon2.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount = 0 Then
                sql = "update tb_inter_boletos_giss set controle=1 where num_documento=" & nDoc
                'cnEicon2.Execute Sql, rdExecDirect
                GoTo Proximo
            Else
                sql = "select * from tb_inter_empresas_giss where num_cadastro=" & nCodigo
                Set RdoAux2 = cnEicon2.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                If RdoAux2.RowCount > 0 Then
                    If (Not IsNull(RdoAux2!Inscricao_gti)) Then
                        nCodigo = RdoAux2!Inscricao_gti
                    Else
                        GoTo Proximo
                    End If
                Else
                    sCPFCNPJ = RdoAux2!cpf_cnpj
                    RdoAux2.Close
                    If Len(sCPFCNPJ) > 11 Then
                        sql = "select codcidadao from cidadao where cnpj='" & Format(sCPFCNPJ, "00000000000000") & "'"
                    Else
                        sql = "select codcidadao from cidadao where cpf='" & sCPFCNPJ & "'"
                    End If
                    Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                    If RdoAux2.RowCount > 0 Then
                        nCodigo = RdoAux2!CodCidadao
                        RdoAux2.Close
                    Else
                        GoTo Proximo
                    End If
                
                End If
                
            End If
        End If
                 
        '** Encontrou a empresa **
        '** Busca última sequencia de guia para lancamento no exercício **"
        nAno = !ano_competencia
        sql = "select max(seqlancamento) as maximo from debitoparcela where codreduzido=" & nCodigo & " and anoexercicio=" & nAno & " and codlancamento=5"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux2!maximo) Then
            nMaxSeq = 0
        Else
            nMaxSeq = RdoAux2!maximo + 1
        End If
        RdoAux2.Close
        
NOVAMENTE:
        '** Grava Parcela **
'        Sql = "insert debitoparcela(codreduzido,anoexercicio,codlancamento,seqlancamento,numparcela,codcomplemento,statuslanc,datavencimento,datadebase,usuario) values("
'        Sql = Sql & nCodigo & "," & nAno & ",5," & nMaxSeq & ",1,0,3,'" & Format(!Data_Vencimento, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "','Giss Online')"
        sql = "insert debitoparcela(codreduzido,anoexercicio,codlancamento,seqlancamento,numparcela,codcomplemento,statuslanc,datavencimento,datadebase,userid) values("
        sql = sql & nCodigo & "," & nAno & ",5," & nMaxSeq & ",1,0,3,'" & Format(!Data_Vencimento, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "',477)"
        cn.Execute sql, rdExecDirect
        
        sql = "select * from debitoparcela where codreduzido=" & nCodigo & " and anoexercicio=" & nAno & " and codlancamento=5 and seqlancamento=" & nMaxSeq
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount = 0 Then
            GoTo NOVAMENTE
        End If
        RdoAux2.Close
        'se não conseguiu gravar tenta novamente até conseguir
        
        
        '** Grava Tributo **
        sql = "insert debitotributo(codreduzido,anoexercicio,codlancamento,seqlancamento,numparcela,codcomplemento,codtributo,valortributo) values("
        sql = sql & nCodigo & "," & nAno & ",5," & nMaxSeq & ",1,0,13," & Virg2Ponto(Format(nValor, "#0.00")) & ")"
        
'        If !valor_imposto < 1 Then
'            Sql = "insert debitotributo(codreduzido,anoexercicio,codlancamento,seqlancamento,numparcela,codcomplemento,codtributo,valortributo) values("
'            Sql = Sql & nCodigo & "," & nAno & ",5," & nMaxSeq & ",1,0,13," & Virg2Ponto(Format(!valor_imposto, "#0.00")) & ")"
'        Else
'            Sql = "insert debitotributo(codreduzido,anoexercicio,codlancamento,seqlancamento,numparcela,codcomplemento,codtributo,valortributo) values("
'            Sql = Sql & nCodigo & "," & nAno & ",5," & nMaxSeq & ",1,0,13," & Virg2Ponto(Format(!valor_imposto / 100, "#0.00")) & ")"
'        End If
        cn.Execute sql, rdExecDirect
        
        '** Grava Parcela-Documento **
        sql = "insert parceladocumento(codreduzido,anoexercicio,codlancamento,seqlancamento,numparcela,codcomplemento,numdocumento) values("
        sql = sql & nCodigo & "," & nAno & ",5," & nMaxSeq & ",1,0," & nDoc & ")"
        cn.Execute sql, rdExecDirect
        
          On Error Resume Next
        '** Grava Documento **
        sql = "insert numdocumento(numdocumento,datadocumento,emissor,valorguia) values("
        sql = sql & nDoc & ",'" & Format(Now, "mm/dd/yyyy") & "','" & "Giss(frmServico)" & "'," & Virg2Ponto(Format(nValor, "#0.00")) & ")"
        cn.Execute sql, rdExecDirect
        On Error GoTo 0

        sql = "update tb_inter_boletos_giss set controle=1 where num_documento=" & nDoc
        cnEicon2.Execute sql, rdExecDirect
        
Proximo:
       DoEvents
       RdoAux.MoveNext
    Loop
    RdoAux.Close
End With

cnEicon2.Close
Exit Sub
Erro:
MsgBox Err.Description
Resume Next
End Sub

Private Sub AtualizaGuiasCanceladas()
On Error GoTo Erro
Dim sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, nAno As Integer, nSeq As Integer, RdoAux3 As rdoResultset
Dim nCodigo As Long, bEmpresa As Boolean, nDoc As Long, nMaxSeq As Integer, sCPFCNPJ As String, sTexto As String, sUser As String

ConectaEicon2

sql = "select * from tb_inter_bol_descartados_giss where controle is null"
Set RdoAux = cnEicon2.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        nCodigo = !num_cadastro
        nDoc = !num_documento
        sql = "select * from parceladocumento where numdocumento=" & nDoc
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                sql = "update debitoparcela set statuslanc=37 where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and "
                sql = sql & "codlancamento=" & !CodLancamento & " and seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO
                cn.Execute sql, rdExecDirect
                
                sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and "
                sql = sql & "codlancamento=" & !CodLancamento & " and seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO
                Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                With RdoAux3
                    If IsNull(!maximo) Then
                        nSeq = 1
                    Else
                        nSeq = !maximo + 1
                    End If
                   .Close
                End With
                sTexto = SubNull(RdoAux!descricao_desc)
                sUser = SubNull(RdoAux!USUARIO)
                If sUser = "" Then sUser = "Eicon --> GTI"
'                Sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USUARIO,DATA) VALUES(" & !CODREDUZIDO & "," & !AnoExercicio & ","
'                Sql = Sql & !CodLancamento & "," & !SeqLancamento & "," & !NumParcela & "," & !CODCOMPLEMENTO & "," & nSeq & ",'" & Mask(sTexto) & "','" & sUser & "','" & Format(sData, "mm/dd/yyyy") & "')"
                sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & !CODREDUZIDO & "," & !AnoExercicio & ","
                sql = sql & !CodLancamento & "," & !SeqLancamento & "," & !NumParcela & "," & !CODCOMPLEMENTO & "," & nSeq & ",'" & Mask(sTexto) & "'," & 477 & ",'" & Format(sData, "mm/dd/yyyy") & "')"
                cn.Execute sql, rdExecDirect
                
                
            End If
            On Error Resume Next
           .MoveNext
           On Error GoTo Erro
           .Close
        End With
                        
        sql = "update tb_inter_bol_descartados_giss set controle=1 where num_documento=" & nDoc
        cnEicon2.Execute sql, rdExecDirect
                        
        DoEvents
       .MoveNext
    Loop
    RdoAux.Close
End With

cnEicon2.Close
Exit Sub
Erro:
MsgBox Err.Description
Resume Next
End Sub

Private Sub Atualiza_DividaAtiva()
Dim sql As String, RdoAux As rdoResultset
On Error GoTo Erro
ConectaEicon2

sql = "SELECT debitoparcela.codreduzido, debitoparcela.anoexercicio, debitoparcela.codlancamento, debitoparcela.seqlancamento, debitoparcela.numparcela, "
sql = sql & "debitoparcela.codcomplemento, debitoparcela.statuslanc, debitoparcela.datavencimento, debitoparcela.datadebase, debitoparcela.codmoeda,"
sql = sql & "debitoparcela.numerolivro, debitoparcela.paginalivro, debitoparcela.numcertidao, debitoparcela.datainscricao, debitoparcela.dataajuiza, debitoparcela.valorjuros,"
sql = sql & "debitoparcela.numprocesso, debitoparcela.intacto, debitoparcela.notificado,  debitoparcela.numexecfiscal, debitoparcela.anoexecfiscal,"
sql = sql & "debitoparcela.processocnj , debitoparcela.simplesnacional, parceladocumento.NumDocumento FROM debitoparcela INNER JOIN "
sql = sql & "parceladocumento ON debitoparcela.codreduzido = parceladocumento.codreduzido AND debitoparcela.anoexercicio = parceladocumento.anoexercicio AND "
sql = sql & "debitoparcela.codlancamento = parceladocumento.codlancamento AND debitoparcela.seqlancamento = parceladocumento.seqlancamento AND "
sql = sql & "debitoparcela.NumParcela = parceladocumento.NumParcela And debitoparcela.CODCOMPLEMENTO = parceladocumento.CODCOMPLEMENTO "
sql = sql & "WHERE (debitoparcela.codreduzido BETWEEN 100000 AND 300000) AND (debitoparcela.codlancamento = 5) AND (debitoparcela.datainscricao IS NOT NULL) AND "
sql = sql & "(debitoparcela.statuslanc < 5) AND (parceladocumento.numdocumento BETWEEN 2000000 AND 3000000) AND (parceladocumento.numdocumento NOT IN "
sql = sql & "(SELECT num_documento FROM GTI_Eicon.dbo.tb_inter_boletos_cdas))"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        sql = "insert tb_inter_boletos_cdas(cod_cliente,num_cadastro,num_documento,mes_competencia,ano_competencia,ins_divida,livro,folha,volume,"
        sql = sql & "fundamento,data_inscricao,status,processo,timestamp) values("
        sql = sql & "2177," & !CODREDUZIDO & "," & !NumDocumento & "," & Month(!DataVencimento) & "," & Year(!DataVencimento) & "," & !numcertidao & ","
        sql = sql & !numerolivro & "," & !paginalivro & ",0,Null,'" & Format(!datainscricao, "mm/dd/yyyy") & "','','','" & Format(Now, "mm/dd/yyyy") & "')"
        cnEicon2.Execute sql, rdExecDirect
       .MoveNext
    Loop
   .Close
End With

cnEicon2.Close
Exit Sub
Erro:
MsgBox Err.Description
Resume Next

End Sub

Private Sub Atualiza_Parcelamento()
Dim sql As String, RdoAux As rdoResultset, nAnoproc As Integer, nNumproc As Long, RdoAux2 As rdoResultset
Dim nQtdeParcela As Integer, sDataParc As String, bCancelado As Boolean, nValorTotal As Double
On Error GoTo Erro
Exit Sub
ConectaEicon2

sql = "SELECT DISTINCT origemreparc.anoproc, origemreparc.numproc FROM origemreparc INNER JOIN parceladocumento ON origemreparc.codreduzido = parceladocumento.codreduzido AND origemreparc.anoexercicio = parceladocumento.anoexercicio AND "
sql = sql & "origemreparc.codlancamento = parceladocumento.codlancamento AND origemreparc.numsequencia = parceladocumento.seqlancamento AND origemreparc.numparcela = parceladocumento.numparcela AND origemreparc.codcomplemento = parceladocumento.codcomplemento INNER JOIN "
sql = sql & "processoreparc ON origemreparc.numproc = processoreparc.numproc AND origemreparc.anoproc = processoreparc.anoproc WHERE (origemreparc.codlancamento = 5) AND (parceladocumento.numdocumento BETWEEN 2000000 AND 3000000) AND (processoreparc.data_exportacao IS NULL) "
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        nAnoproc = !AnoProc
        nNumproc = !NumProc
        
        sql = "SELECT numproc, anoproc, datareparc, qtdeparcela, cancelado, data_exportacao From processoreparc Where AnoProc=" & nAnoproc & " And NumProc=" & nNumproc
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        nQtdeParcela = rdoaux2Qqtdeparcela
        sDataParc = Format(RdoAux2!datareparc, "dd/mm/yyyy")
        bCancelado = IIf(RdoAux2!Cancelado = 1, True, False)
        RdoAux2.Close
        
        sql = "SELECT SUM(debitotributo.valortributo) AS Soma FROM debitotributo INNER JOIN destinoreparc ON debitotributo.codreduzido = destinoreparc.codreduzido AND debitotributo.anoexercicio = destinoreparc.anoexercicio AND "
        sql = sql & "debitotributo.codlancamento = destinoreparc.codlancamento AND debitotributo.seqlancamento = destinoreparc.numsequencia AND debitotributo.NumParcela = destinoreparc.NumParcela And debitotributo.CODCOMPLEMENTO = destinoreparc.CODCOMPLEMENTO "
        sql = sql & "Where destinoreparc.NumProc = " & nNumproc & " And destinoreparc.AnoProc = " & nAnoproc
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        nValorTotal = RdoAux2!soma
        RdoAux2.Close
        
        
       .MoveNext
    Loop
   .Close
End With

cnEicon2.Close
Exit Sub
Erro:
MsgBox Err.Description
Resume Next

End Sub

Private Sub AtualizaBaixaGiss()
Dim RdoAux As rdoResultset, sql As String, nCodReduz As Long, nNumDoc As Long, nSeqPag As Integer, RdoAux3 As rdoResultset
Dim RdoAux2 As rdoResultset, nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer

ConectaEicon2
'baixa ok
sql = "SELECT num_documento,num_cadastro,valor_pago,data_pagamento FROM tb_inter_baixa_detalhe_giss WHERE controle IS NULL AND num_cadastro BETWEEN 100000 AND 200000"
Set RdoAux = cnEicon2.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If !num_documento = 2224949 Then MsgBox "teste"
        nCodReduz = !num_cadastro
        nNumDoc = !num_documento
        sql = "select * from parceladocumento where numdocumento=" & nNumDoc
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                nAno = !AnoExercicio
                nLanc = !CodLancamento
                nSeq = !SeqLancamento
                nParc = !NumParcela
                nCompl = !CODCOMPLEMENTO
            End If
        End With
    
        sql = "SELECT MAX(SEQPAG) AS MAXIMO FROM DEBITOPAGO WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND "
        sql = sql & "SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO = " & nCompl
        Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux3
            If IsNull(!maximo) Then
                nSeqPag = 0
            Else
                If .RowCount = 0 Then
                   nSeqPag = 0
                Else
                   nSeqPag = !maximo + 1
                End If
            End If
            .Close
        End With
    
        sql = "INSERT DEBITOPAGO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQPAG,DATAPAGAMENTO,DATARECEBIMENTO,VALORPAGO,VALORPAGOREAL,NUMDOCUMENTO) VALUES(" & nCodReduz & ","
        sql = sql & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & "," & nSeqPag & ",'" & Format(RdoAux!Data_Pagamento, "mm/dd/yyyy") & "','" & Format(RdoAux!Data_Pagamento, "mm/dd/yyyy") & "',"
        sql = sql & Virg2Ponto(CStr(RdoAux!valor_pago)) & "," & Virg2Ponto(CStr(RdoAux!valor_pago)) & "," & nNumDoc & ")"
        cn.Execute sql, rdExecDirect
    
        sql = "UPDATE DEBITOPARCELA SET STATUSLANC=2 WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND "
        sql = sql & "SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO = " & nCompl
        cn.Execute sql, rdExecDirect
    
       .MoveNext
    Loop
   .Close
End With


cnEicon2.Close




End Sub
