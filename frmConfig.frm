VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmConfig 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuração"
   ClientHeight    =   5250
   ClientLeft      =   12765
   ClientTop       =   3615
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   5310
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   5610
      TabIndex        =   14
      Top             =   480
      Width           =   2625
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   8295
      MultiSelect     =   2  'Extended
      TabIndex        =   13
      Top             =   495
      Width           =   3630
   End
   Begin VB.CommandButton cmdPagos 
      Caption         =   "Pagos"
      Enabled         =   0   'False
      Height          =   330
      Left            =   2610
      TabIndex        =   11
      Top             =   4770
      Width           =   960
   End
   Begin prjChameleon.chameleonButton btArquivos 
      Height          =   315
      Index           =   1
      Left            =   180
      TabIndex        =   9
      ToolTipText     =   "IP"
      Top             =   4770
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Le Arq"
      ENAB            =   0   'False
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmConfig.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton btCorrige 
      Height          =   315
      Index           =   0
      Left            =   4140
      TabIndex        =   8
      ToolTipText     =   "IP"
      Top             =   4740
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Corrige"
      ENAB            =   0   'False
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmConfig.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CommandButton cmdClearBD 
      Caption         =   "Limpa BD"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1740
      TabIndex        =   5
      Top             =   7800
      Width           =   1935
   End
   Begin VB.TextBox txtOld 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      TabIndex        =   4
      Top             =   4125
      Width           =   5085
   End
   Begin VB.TextBox txtValor 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      TabIndex        =   3
      Top             =   3000
      Width           =   5085
   End
   Begin VB.ListBox lstParam 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      ItemData        =   "frmConfig.frx":0038
      Left            =   90
      List            =   "frmConfig.frx":003A
      TabIndex        =   2
      Top             =   90
      Width           =   5085
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   4140
      TabIndex        =   0
      ToolTipText     =   "Cancelar Edição"
      Top             =   3420
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Cancelar"
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmConfig.frx":003C
      PICN            =   "frmConfig.frx":0058
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdGravar 
      Height          =   315
      Left            =   3060
      TabIndex        =   1
      ToolTipText     =   "Gravar os Dados"
      Top             =   3420
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Gravar"
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmConfig.frx":01B2
      PICN            =   "frmConfig.frx":01CE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ProgressBar Pb 
      Height          =   165
      Left            =   135
      TabIndex        =   6
      Top             =   3825
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   291
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   2265
      Left            =   30
      TabIndex        =   12
      Top             =   5220
      Width           =   16815
      _ExtentX        =   29660
      _ExtentY        =   3995
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Arquivo"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ShortName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Banco"
         Object.Width           =   1305
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Dt.Rec."
         Object.Width           =   1835
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "CNPJ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Codigo"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Ano"
         Object.Width           =   953
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Mes"
         Object.Width           =   953
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Dt.Venc"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Valor"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "Exer."
         Object.Width           =   953
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Text            =   "Sq"
         Object.Width           =   776
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "Dup"
         Object.Width           =   776
      EndProperty
   End
   Begin VB.CommandButton btCadastro 
      Caption         =   "baixar"
      Height          =   330
      Left            =   1380
      TabIndex        =   10
      Top             =   4770
      Width           =   960
   End
   Begin VB.Image img 
      Height          =   885
      Left            =   5850
      Top             =   3090
      Width           =   1065
   End
   Begin VB.Label lblPB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100 %"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2295
      TabIndex        =   7
      Top             =   3825
      Width           =   480
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type idUser
    id As Integer
    Nome As String
End Type
Dim aIdUser() As idUser

Private Type Proc
    Numero As Long
    ano As Integer
    Cancelado As Boolean
End Type

Private Type tLaser
    Codigo As Long
    ano As Integer
    Area_Terreno As Double
    Area_Predial As Double
End Type

Private Type Registro
    nNumDoc As Long
    nSeq As Integer
    sDataDoc As String
    sDataPag As Date
    sDataCred As Date
    nValorPago As Double
    sAgencia As String
    nValorTarifa As Double
    sSitRetorno As String
    bExiste As Boolean
    bIsentoMJ As Boolean
    sCnpj As String
    nAno As Integer
    nMes As Integer
    sDataVencto As Date
    nValorTarifaBancaria As Double
    nSomaTributo As Double
    sDataPagCalc As Date
End Type

Private Type Documento
    nNumDoc As Long
    nSeqDoc As Integer
    nCodReduz As Long
    nAno As Integer
    nLanc As Integer
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
    sDataVencto As String
    sSit As String
    nNumeroLivro As Integer
    nPaginaLivro As Integer
    bAjuizado As Boolean
    nValorPrincipal As Double
    nValorMulta As Double
    nValorJuros As Double
    nValorCorrecao As Double
    nValorTotal As Double
    nValorTarifa As Double
    nValorDif As Double
    nValorCompensado As Double
    sBx As String
    sDp As String
    nSeqReg As Integer
    bExiste As Boolean
    sCnpj As String
End Type

Private Type tAREA
    nSeq As Integer
    
End Type

Private Type SIMPLES
    nCodigo As Long
    nAno As Integer
End Type

Private Type Inad
    nQtdeAtrasado As Long
    nLanc As Integer
    nAnoLc As Integer
    nMesLc As Integer
    nAnoPg As Integer
    nMesPg As Integer
    nValor As Double
End Type

Private Type tProtocolo
    nCodReduz As Long
    nAno As Integer
    nLanc As Integer
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
    nCodTributo As Integer
End Type

Private Type tCalculoResumo
    nValor0 As Double
    nValor91 As Double
    nValor92 As Double
    nValor1 As Double
    nValor2 As Double
    nValor3 As Double
    nValor4 As Double
    nValor5 As Double
    nValor6 As Double
    nValor7 As Double
    nValor8 As Double
    nValor9 As Double
    nValor10 As Double
    nValor11 As Double
    nValor12 As Double
    nDoc0 As Long
    nDoc91 As Long
    nDoc92 As Long
    nDoc1 As Long
    nDoc2 As Long
    nDoc3 As Long
    nDoc4 As Long
    nDoc5 As Long
    nDoc6 As Long
    nDoc7 As Long
    nDoc8 As Long
    nDoc9 As Long
    nDoc10 As Long
    nDoc11 As Long
    nDoc12 As Long
    DataVento1 As String
    DataVento2 As String
    DataVento3 As String
    DataVento4 As String
    DataVento5 As String
    DataVento6 As String
    DataVento7 As String
    DataVento8 As String
    DataVento9 As String
    DataVento10 As String
    DataVento11 As String
    DataVento12 As String
End Type

Private Type Debito
    nAno As Integer
    nLanc As Integer
    sLanc As String
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
    nSituacao As Integer
    sSituacao As String
    sVencto As String
    sDA As String
    sAj As String
    nCodTributo As Double
    nValorTributo As Double
    nValorJuros As Double
    nValorMulta As Double
    nValorCorrecao As Double
    nValorAtual As Double
    nValorHon As Double
    nValorJurApl As Double
    nSaldo As Double
    nCodBanco As Integer
    dDataPag As Date
    sNotificado As String
    sExFiscal As String
    nProt_certidao As Long
    nProt_dtremessa As Date
End Type

Private Type tTramite
    nAno As Integer
    nNumero As Integer
    nSeq As Integer
    nCCusto As Integer
End Type

Private Type tMei
    id As Integer
    Codigo As Long
    datainicio As String
    datafim As String
    apagar As Boolean
End Type


Private Sub btArquivos_Click(Index As Integer)
Dim sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, x As Integer
Dim sNomeArq As String, nCodBanco As Integer, sFullPath As String, sReg As String, sCnpj As String
Dim sDataIni As String, sDataFim As String, sEncerrada As String, sSuspensa As String
Dim nCodReduzido As Long, sClasse As String, sSimples As String, RdoAux3 As rdoResultset, nNumDoc As Long, FF1 As Integer

If NomeDeLogin <> "SCHWARTZ" Then Exit Sub
'GoTo Parte2
sql = "truncate table mei2"
cn.Execute sql, rdExecDirect


On Error Resume Next
FF1 = FreeFile()
Open "c:\trabalho\simples.txt" For Binary Access Read Write As FF1

    While Not EOF(FF1)
        Input #FF1, sReg
        sCnpj = Left(sReg, 8)
        sDataIni = Mid(sReg, 9, 8)
        sDataFim = Mid(sReg, 17, 8)
        sDataIni = ConvDataSerial(sDataIni)
        sDataFim = ConvDataSerial(sDataFim)
        If sDataFim = "00/00/0000" Then sDataFim = "01/01/1900"
        
        sql = "select cnpj from simples_codigo where cnpj='" & sCnpj & "'"
        Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If rdoAux.RowCount > 0 Then
            rdoAux.Close
            GoTo Proximo
        End If
        rdoAux.Close
        
        
        sql = "SELECT codigomob,dataencerramento From mobiliario WHERE SUBSTRING(cnpj, 1, 8) = '" & sCnpj & "'"
        Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If rdoAux.RowCount > 0 Then
            nCodReduzido = rdoAux!codigomob
            If IsNull(rdoAux!dataencerramento) Then
                sEncerrada = "N"
            Else
                sEncerrada = "S"
            End If
        Else
            nCodReduzido = 0
            sEncerrada = "N"
        End If
        rdoAux.Close

        sql = "insert mei2 (cnpj,codigo,datainicio,datafim,encerrada) values('" & sCnpj & "'," & nCodReduzido & ",'"
        sql = sql & Format(sDataIni, sDataFormat) & "','" & Format(sDataFim, sDataFormat) & "','" & sEncerrada & "')"
 '      Sql = "insert simples_codigo (cnpj) values('" & sCNPJ & "')"
        cn.Execute sql, rdExecDirect
        DoEvents
        
       'suspenção
        sql = "SELECT CODTIPOEVENTO,DATAEVENTO FROM MOBILIARIOEVENTO WHERE CODMOBILIARIO=" & nCodReduzido
        sql = sql & " ORDER BY DATAEVENTO DESC"
        Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With rdoAux
            If .RowCount = 0 Then
                sSuspensa = "N"
            Else
                If !CODTIPOEVENTO = 2 Then
                    sSuspensa = "S"
                Else
                    sSuspensa = "N"
                End If
            End If
           .Close
        End With
        sql = "update mei2 set suspensa='" & sSuspensa & "' where codigo=" & nCodReduzido
        cn.Execute sql, rdExecDirect
Proximo:
    Wend
CloseFile2:
Close #FF1

'Exit Sub
Parte2:
sql = "select codigo from mei2 where codigo > 0"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
Do Until rdoAux.EOF
    nCodReduzido = rdoAux!Codigo
    sSimples = IIf(SNCheck2(nCodReduzido), "S", "N")
    sql = "update mei2 set esimples='" & sSimples & "' where codigo=" & nCodReduzido
    cn.Execute sql, rdExecDirect
    rdoAux.MoveNext
Loop
rdoAux.Close


MsgBox "fim"

End Sub

Private Sub btCadastro_Click()
Dim sql As String, rdoAux As rdoResultset, nAno1 As Integer, nLanc1 As Integer, RdoAux2 As rdoResultset, x As Integer, RdoAux3 As rdoResultset
Dim nCodReduz As Long, aCodigo(17) As Integer, nCodigo1 As Long, nSeq1 As Integer, nParc1 As Integer, nCompl1 As Integer, sDataVencto As String, nValor1 As Double
Dim aOrigem() As Documento, bFind As Boolean, y As Integer, nCodigo2 As Long, nSeq2 As Integer, nParc2 As Integer, nCompl2 As Integer


If NomeDeLogin <> "SCHWARTZ" Then Exit Sub
ReDim aIdUser(0)

sql = "select id,nomelogin from usuario order by nomelogin"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    Do Until .EOF
        ReDim Preserve aIdUser(UBound(aIdUser) + 1)
        aIdUser(UBound(aIdUser)).id = !id
        aIdUser(UBound(aIdUser)).Nome = !NomeLogin
       .MoveNext
    Loop
   .Close
End With
VerificaTaxaMei
'CorrigeMei2
'VerificaProtestoCancelado
'VerificaProtestoPago
'VerificaAreaGeoCalculada
'CorrigeEnviadoProtestoPago
'CorrigeParcelamento2026
'CorrigeParcelamentoBloqueio
'Incluir_AreaGeo
'ResumoIPTU_GEO
'Corrige_Telefone
'ExtrairTelefones
'Atualiza_Protesto
'Corrige_Complemento
'ListaEmpresasVereador
'BaseEgati
'PagamentoCdas
'RelPagamentoISS
'Corrige_Tramite
'Baixa2Eicon
'ParcelamentoParaBase
'RelatorioRefis
'MudaStatus2019
'CorrigeCPFCNPJ
'Incluir_para_registro
'Corrige_BaixaDocGiss
'Corrige_BaixaGiss
'Corrige_CdasProtesto
'RemoveIPTUCalculo
'RemoveProtocolo
'CorrigeStatusPago
'Cancela_200Reais
'Corrige_Cep_Cdas
'Cdas_Nao_Ajuizadas
'Corrige_EnderecoLaser
'AreaLaserIptu
'LimpezaBD
'Corrige_Vencimentos_Parcelamento
'CorrigeCep
'Corrige_Permei
'Corrige_ParcelaPix
'Nova_RazaoSocial
'Corrige_Importacao_Eicon
'Corrige_TaxaLic
'Corrige_IPTU
'ISS_Errado
'Rel_Ana
'Apaga_Taxa_Protocolo
'Usuario_web_autoriza
'Imovel_Limarfe
'Imovel_Arbor
'Imovel_Santander
'Imovel_Cem
'Imovel_Wegg
'Imovel_Caixa
'Imovel_Cunha
'Imovel_DePaula
'Digito_Carta
'Lista_Empresa_Devedora
'Inscricao_Estadual
'Corrige_Tabela_Tributo
'Corrige_Cnae_Bar
'Descarte_Processo
'Corrige_Endereco_Empresa
'Corrige_Unica_2021
'InadimplenteValor
'RelProcesso
'Simples_Cnpj
Exit Sub

aCodigo(0) = 1992
aCodigo(1) = 22498
aCodigo(2) = 22499
aCodigo(3) = 22501
aCodigo(4) = 22502
aCodigo(5) = 22503
aCodigo(6) = 22504
aCodigo(7) = 22505
aCodigo(8) = 22506
aCodigo(9) = 22507
aCodigo(10) = 22509
aCodigo(11) = 22510
aCodigo(12) = 22511
aCodigo(13) = 22512
aCodigo(14) = 22513
aCodigo(15) = 22514
aCodigo(16) = 22508


'carrega origem
ReDim aOrigem(0)
sql = "SELECT debitoparcela.codreduzido, debitoparcela.anoexercicio, debitoparcela.codlancamento, debitoparcela.seqlancamento, debitoparcela.numparcela,debitoparcela.codcomplemento, debitoparcela.datavencimento, SUM(debitotributo.valortributo) AS Soma "
sql = sql & "FROM debitoparcela INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND "
sql = sql & "debitoparcela.NumParcela = debitotributo.NumParcela And debitoparcela.CODCOMPLEMENTO = debitotributo.CODCOMPLEMENTO WHERE (debitotributo.codtributo <> 3) GROUP BY debitoparcela.codreduzido, debitoparcela.anoexercicio, debitoparcela.codlancamento, debitoparcela.seqlancamento, debitoparcela.numparcela,"
sql = sql & "debitoparcela.CODCOMPLEMENTO , debitoparcela.DataVencimento Having (debitoparcela.CODREDUZIDO = 38258) ORDER BY debitoparcela.anoexercicio, debitoparcela.codlancamento, debitoparcela.seqlancamento, debitoparcela.numparcela, debitoparcela.codcomplemento"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    Do Until .EOF
        ReDim Preserve aOrigem(UBound(aOrigem) + 1)
        aOrigem(UBound(aOrigem)).nAno = !AnoExercicio
        aOrigem(UBound(aOrigem)).nLanc = !CodLancamento
        aOrigem(UBound(aOrigem)).nSeq = !SeqLancamento
        aOrigem(UBound(aOrigem)).nParc = !NumParcela
        aOrigem(UBound(aOrigem)).nCompl = !CODCOMPLEMENTO
        aOrigem(UBound(aOrigem)).sDataVencto = Format(!DataVencimento, "dd/mm/yyyy")
        aOrigem(UBound(aOrigem)).nValorPrincipal = Round(!soma, 2)
       .MoveNext
    Loop
   .Close
End With
'GoTo fim
'fim origem

sql = "delete from transfere_debito"
cn.Execute sql, rdExecDirect

For x = 0 To 16
    lblPB.Caption = x
    nCodReduz = aCodigo(x)
    sql = "select * FROM debitoparcela where codreduzido=" & nCodReduz & " and statuslanc=13"
    Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
    With rdoAux
        Do Until .EOF
            nAno1 = !AnoExercicio
            nLanc1 = !CodLancamento
            nSeq1 = !SeqLancamento
            nParc1 = !NumParcela
            nCompl1 = !CODCOMPLEMENTO
            sDataVencto = Format(!DataVencimento, "dd/mm/yyyy")
            
            sql = "select sum(valortributo) as soma from debitotributo where codreduzido=" & nCodReduz & " and anoexercicio=" & nAno1 & " and "
            sql = sql & "codlancamento=" & nLanc1 & " and seqlancamento=" & nSeq1 & " and numparcela=" & nParc1 & " and codcomplemento=" & nCompl1 & " and "
            sql = sql & "codtributo<>3"
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            nValor1 = Round(RdoAux2!soma, 2)
            RdoAux2.Close
            
            'Localizar estes valores na matriz de origem
            bFind = False
            For y = 1 To UBound(aOrigem)
                With aOrigem(y)
                    If .nAno = nAno1 And .nLanc = nLanc1 And .nParc = nParc1 And .nCompl = nCompl1 And .sDataVencto = sDataVencto And .nValorPrincipal = nValor1 Then
                        bFind = True
                        sql = "insert transfere_debito (codigo1,ano1,lanc1,seq1,parc1,comp1,datavencto1,valor1,codigo2,ano2,lanc2,seq2,parc2,comp2,datavencto2,valor2,statuslanc) "
                        sql = sql & "values(" & 38258 & "," & .nAno & "," & .nLanc & "," & .nSeq & "," & .nParc & "," & .nCompl & ",'" & Format(.sDataVencto, "mm/dd/yyyy") & "',"
                        sql = sql & Virg2Ponto(CStr(.nValorPrincipal)) & "," & nCodReduz & "," & nAno1 & "," & nLanc1 & "," & nSeq1 & "," & nParc1 & "," & nCompl1 & ",'"
                        sql = sql & Format(sDataVencto, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(nValor1)) & "," & rdoAux!statuslanc & ")"
                        cn.Execute sql, rdExecDirect
                        
                        Exit For
                    End If
                End With
            Next
           
            If Not bFind Then
                MsgBox "não achei"
            End If
           
           
            DoEvents
           .MoveNext
        Loop
        rdoAux.Close
    End With
Next

Etapa2:

sql = "SELECT transfere_debito.codigo1, transfere_debito.ano1, transfere_debito.lanc1, transfere_debito.seq1, transfere_debito.parc1, transfere_debito.comp1,transfere_debito.datavencto1, transfere_debito.valor1,"
sql = sql & "transfere_debito.codigo2, transfere_debito.ano2, transfere_debito.lanc2, transfere_debito.seq2,transfere_debito.parc2, transfere_debito.comp2, transfere_debito.datavencto2, transfere_debito.valor2, transfere_debito.statuslanc,"
sql = sql & "debitoparcela.statuslanc AS sit2 FROM transfere_debito INNER JOIN debitoparcela ON transfere_debito.codigo1 = debitoparcela.codreduzido AND transfere_debito.ano1 = debitoparcela.anoexercicio AND "
sql = sql & "transfere_debito.lanc1 = debitoparcela.codlancamento AND transfere_debito.seq1 = debitoparcela.seqlancamento AND transfere_debito.parc1 = debitoparcela.NumParcela And transfere_debito.comp1 = debitoparcela.CODCOMPLEMENTO "
sql = sql & "ORDER BY transfere_debito.codigo2, transfere_debito.ano2, transfere_debito.lanc2, transfere_debito.parc2"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    Do Until .EOF
        nCodigo1 = !Codigo1: nAno1 = !ano1: nLanc1 = !lanc1: nSeq1 = !seq1: nParc1 = !parc1: nCompl1 = !comp1
        nCodigo2 = !Codigo2: nAno2 = !ano2: nLanc2 = !lanc2: nSeq2 = !seq2: nParc2 = !parc2: nCompl2 = !comp2
        
        'ATUALIZA PARCELADOCUMENTO
        sql = "UPDATE PARCELADOCUMENTO SET CODREDUZIDO=" & nCodigo2 & ",SEQLANCAMENTO=" & nSeq2 & " WHERE CODREDUZIDO=" & nCodigo1 & " AND "
        sql = sql & "ANOEXERCICIO=" & nAno1 & " AND CODLANCAMENTO=" & nLanc1 & " AND SEQLANCAMENTO=" & nSeq1 & " AND NUMPARCELA=" & nParc1 & " AND "
        sql = sql & "CODCOMPLEMENTO=" & nCompl1
        cn.Execute sql, rdExecDirect
        'ATUALIZA DEBITOPAGO
        sql = "UPDATE DEBITOPAGO SET CODREDUZIDO=" & nCodigo2 & ",SEQLANCAMENTO=" & nSeq2 & " WHERE CODREDUZIDO=" & nCodigo1 & " AND "
        sql = sql & "ANOEXERCICIO=" & nAno1 & " AND CODLANCAMENTO=" & nLanc1 & " AND SEQLANCAMENTO=" & nSeq1 & " AND NUMPARCELA=" & nParc1 & " AND "
        sql = sql & "CODCOMPLEMENTO=" & nCompl1
        cn.Execute sql, rdExecDirect
        'ATUALIZA OBS
        sql = "UPDATE obsparcela SET CODREDUZIDO=" & nCodigo2 & ",SEQLANCAMENTO=" & nSeq2 & " WHERE CODREDUZIDO=" & nCodigo1 & " AND "
        sql = sql & "ANOEXERCICIO=" & nAno1 & " AND CODLANCAMENTO=" & nLanc1 & " AND SEQLANCAMENTO=" & nSeq1 & " AND NUMPARCELA=" & nParc1 & " AND "
        sql = sql & "CODCOMPLEMENTO=" & nCompl1
        cn.Execute sql, rdExecDirect
                
        sql = "UPDATE DEBITOPARCELA SET STATUSLANC=13 WHERE CODREDUZIDO=" & nCodigo1 & " AND "
        sql = sql & "ANOEXERCICIO=" & nAno1 & " AND CODLANCAMENTO=" & nLanc1 & " AND SEQLANCAMENTO=" & nSeq1 & " AND NUMPARCELA=" & nParc1 & " AND "
        sql = sql & "CODCOMPLEMENTO=" & nCompl1
        cn.Execute sql, rdExecDirect
        
        sql = "UPDATE DEBITOPARCELA SET STATUSLANC=" & !SIT2 & " WHERE CODREDUZIDO=" & nCodigo2 & " AND "
        sql = sql & "ANOEXERCICIO=" & nAno2 & " AND CODLANCAMENTO=" & nLanc2 & " AND SEQLANCAMENTO=" & nSeq2 & " AND NUMPARCELA=" & nParc2 & " AND "
        sql = sql & "CODCOMPLEMENTO=" & nCompl2
        cn.Execute sql, rdExecDirect
        
        DoEvents
       .MoveNext
    Loop
   .Close
End With


Fim:
MsgBox "fim"

Exit Sub
Erro:
MsgBox rdoErrors(0).Description

End Sub

Private Sub cmdFase4_Click()

Dim sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset
Dim nPagas As Integer


If NomeDeLogin <> "SCHWARTZ" Then Exit Sub
Pb.value = 0: lblPB.Caption = "0 %": nPos = 1


sql = "SELECT * from daf_reg order by codreduzido"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    Do Until .EOF
        If nPos Mod 100 = 0 Then CallPb nPos, nTot
                
        sql = "select count(*) as contador from debitopago where codreduzido=" & !CODREDUZIDO & " and "
        sql = sql & "datapagamento='" & Format(!datapagto, "mm/dd/yyyy") & "'"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux2!contador) Then
            nPagas = 0
        Else
            nPagas = RdoAux2!contador
        End If
                
        sql = "update daf_reg set pagas=" & nPagas & " where codreduzido=" & !CODREDUZIDO & " and "
        sql = sql & "datapagto='" & Format(!datapagto, "mm/dd/yyyy") & "'"
        cn.Execute sql, rdExecDirect
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With



MsgBox "fim"

Exit Sub
Erro:
MsgBox rdoErrors(0).Description



End Sub

Private Sub cmdFase5_Click()
Dim sql As String, rdoAux As rdoResultset
If NomeDeLogin <> "SCHWARTZ" Then
    MsgBox "Erro fatal."
    Exit Sub
End If
cmdFase5.Enabled = False

sql = "DELETE FROM SIMPLESCNPJ"
cn.Execute sql, rdExecDirect

sql = "SELECT * FROM RESUMOARQSN WHERE nome='CNPJ NÃO LOCALIZADO' ORDER BY CNPJ"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    Do Until .EOF
        sql = "INSERT SIMPLESCNPJ (CNPJ,ARQUIVOSHORT,BANCO,DATAARRECADA,DATAVENCTO,ANOCOMP,MESCOMP,PRINCIPAL,JUROS,"
        sql = sql & "MULTA,AGENCIA,CODREDUZIDO) VALUES('" & RetornaNumero(!Cnpj) & "','" & !ArquivoShort & "'," & !Banco & ",'" & Format(!DataArrecada, "mm/dd/yyyy") & "','"
        sql = sql & Format(!DataVencto, "mm/dd/yyyy") & "'," & !AnoComp & "," & !MesComp & "," & Virg2Ponto(!principal) & "," & Virg2Ponto(!Juros) & "," & Virg2Ponto(!multa) & ",'"
        sql = sql & !Agencia & "'," & 0 & ")"
        cn.Execute sql, rdExecDirect
       .MoveNext
    Loop
   .Close
End With

MsgBox "FIM"
End Sub

Private Sub btCorrige_Click(Index As Integer)
If NomeDeLogin <> "SCHWARTZ" Then Exit Sub
'Restaura
'CorrigeProcesso
'BaixaEicon
NaoPagoParaPago
End Sub

Private Sub Restaura()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nAno As Integer, nLc As Integer, nSq As Integer, nPc As Integer, nCp As Integer
Dim nTot As Long, nPos As Long, sNumProc As String, nNumproc As Long, nAnoproc As Integer

ConectaBkp

GoTo FASE8

FASE1:
sql = "SELECT * from facequadra order by coddistrito,codsetor,codquadra,codface"
Set rdoAux = cnBkp.OpenResultset(sql, rdOpenKeyset, rydConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        sql = "select * from facequadra where coddistrito=" & !CODDISTRITO & " and codsetor=" & !CODSETOR & " and codquadra=" & !CODQUADRA & " and codface=" & !CODFACE
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            If RdoAux2!CodLogr <> !CodLogr Then
                DoEvents
                sql = "update facequadra set codlogr=" & !CodLogr & " where coddistrito=" & !CODDISTRITO & " and codsetor=" & !CODSETOR & " and codquadra=" & !CODQUADRA & " and codface=" & !CODFACE
                cn.Execute sql, rdExecDirect
            End If
        Else
            sql = "insert facequadra select coddistrito,codsetor,codquadra,codface,codlogr,codagrupa,pavimento,quadras from tributacaobkp..facequadra where "
            sql = sql & "coddistrito=" & !CODDISTRITO & " and codsetor=" & !CODSETOR & " and codquadra=" & !CODQUADRA & " and codface=" & !CODFACE
            cn.Execute sql, rdExecDirect
        End If
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

FASE2:
sql = "SELECT * from logradouro"
Set rdoAux = cnBkp.OpenResultset(sql, rdOpenKeyset, rydConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        sql = "select * from logradouro where codlogradouro=" & !CodLogradouro
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            If RdoAux2!NomeLogradouro <> !NomeLogradouro Then
                DoEvents
         '       MsgBox "BKP:" & !CodLogradouro & "-" & !NomeLogradouro & " -> Atual:" & RdoAux2!CodLogradouro & "-" & RdoAux2!NomeLogradouro
 '               Sql = "update logradouro set codtipolog=" & Val(SubNull(!codtipolog)) & ",codtitlog=" & Val(SubNull(!codtitlog)) & ",nomelogradouro='" & Mask(!NomeLogradouro) & "',endereco='" & Mask(!Endereco) & "' where codlogradouro=" & !CodLogradouro
'                cn.Execute Sql, rdExecDirect
            End If
        Else
            'MsgBox "Não existe:" & !CodLogradouro & "-" & !NomeLogradouro
            sql = "insert logradouro select codlogradouro,endereco,dataofic,numofic,codtipolog,codtitlog,nomelogradouro from tributacaobkp..logradouro where "
            sql = sql & "codlogradouro=" & !CodLogradouro
            cn.Execute sql, rdExecDirect
        End If
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

FASE3:
sql = "SELECT * from bairro where siglauf='SP' and codcidade=413 order by codbairro"
Set rdoAux = cnBkp.OpenResultset(sql, rdOpenKeyset, rydConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        sql = "select * from bairro where siglauf='SP' and codcidade=413 and codbairro=" & !CodBairro
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            If RdoAux2!DescBairro <> !DescBairro Then
                DoEvents
'                MsgBox "BKP:" & !CodBairro & "-" & !DescBairro & " -> Atual:" & RdoAux2!CodBairro & "-" & RdoAux2!DescBairro
                sql = "update bairro set descbairro='" & Mask(!DescBairro) & "' where siglauf='SP' and codcidade=413 and codbairro=" & !CodBairro
                cn.Execute sql, rdExecDirect
            End If
        Else
 '           MsgBox "Não existe:" & !CodBairro & "-" & !DescBairro
            sql = "insert bairro select siglauf,codcidade,codbairro,descbairro from tributacaobkp..bairro where "
            sql = sql & "siglauf='SP' and codcidade=413 and codbairro=" & !CodBairro
            cn.Execute sql, rdExecDirect
        End If
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With


FASE4:
sql = "SELECT * from cadimob order by codreduzido"
Set rdoAux = cnBkp.OpenResultset(sql, rdOpenKeyset, rydConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        sql = "select * from cadimob where codreduzido=" & !CODREDUZIDO
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            If RdoAux2!Li_CodBairro <> !Li_CodBairro Then
                DoEvents
              '  MsgBox "BKP:" & !CODREDUZIDO & "-" & !Li_CodBairro & " -> Atual:" & RdoAux2!CODREDUZIDO & "-" & RdoAux2!Li_CodBairro
                sql = "update cadimob set li_codbairro=" & !Li_CodBairro & " where codreduzido=" & !CODREDUZIDO
                cn.Execute sql, rdExecDirect
            End If
        Else
            MsgBox "Não existe:" & !CODREDUZIDO
'            Sql = "insert bairro select siglauf,codcidade,codbairro,descbairro from tributacaobkp..bairro where "
'            Sql = Sql & "siglauf='SP' and codcidade=413 and codbairro=" & !CodBairro
'            cn.Execute Sql, rdExecDirect
        End If
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With


FASE5:
sql = "SELECT * from cidadao order by codcidadao"
Set rdoAux = cnBkp.OpenResultset(sql, rdOpenKeyset, rydConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        sql = "select * from cidadao where codcidadao=" & !CodCidadao
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            If RdoAux2!CodLogradouro <> !CodLogradouro Or RdoAux2!CodBairro <> !CodBairro Then
                DoEvents
                If Val(SubNull(!CodLogradouro)) > 0 Then
                    'MsgBox !CodCidadao & " BKP:" & !CodLogradouro & "-" & !CodBairro & " -> Atual:" & RdoAux2!CodLogradouro & "-" & RdoAux2!CodBairro
                sql = "update cidadao set codlogradouro=" & !CodLogradouro & ",codbairro=" & IIf(IsNull(!CodBairro), 999, !CodBairro) & " where codcidadao=" & !CodCidadao
                cn.Execute sql, rdExecDirect
               End If
            End If
        Else
            MsgBox "Não existe:" & !CodCidadao
'            Sql = "insert bairro select siglauf,codcidade,codbairro,descbairro from tributacaobkp..bairro where "
'            Sql = Sql & "siglauf='SP' and codcidade=413 and codbairro=" & !CodBairro
'            cn.Execute Sql, rdExecDirect
        End If
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With
GoTo Fim
FASE6:
sql = "SELECT * from endentrega where ee_cidade=413 order by codreduzido"
Set rdoAux = cnBkp.OpenResultset(sql, rdOpenKeyset, rydConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        sql = "select * from endentrega where ee_cidade=413 and codreduzido=" & !CODREDUZIDO
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            If RdoAux2!Ee_CodLog <> !Ee_CodLog Or RdoAux2!Ee_Bairro <> !Ee_Bairro Then
                DoEvents
                If Val(SubNull(!Ee_CodLog)) > 0 Then
               '     MsgBox !CODREDUZIDO & " BKP:" & !Ee_CodLog & "-" & !Ee_Bairro & " -> Atual:" & RdoAux2!Ee_CodLog & "-" & RdoAux2!Ee_Bairro
                sql = "update endentrega set ee_codlog=" & !Ee_CodLog & ",ee_bairro=" & IIf(IsNull(!Ee_Bairro), 999, !Ee_Bairro) & " where codreduzido=" & !CODREDUZIDO
                cn.Execute sql, rdExecDirect
               End If
            End If
        Else
'            MsgBox "Não existe:" & !CODREDUZIDO
'            Sql = "insert bairro select siglauf,codcidade,codbairro,descbairro from tributacaobkp..bairro where "
'            Sql = Sql & "siglauf='SP' and codcidade=413 and codbairro=" & !CodBairro
'            cn.Execute Sql, rdExecDirect
        End If
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

GoTo Fim

FASE7:
sql = "SELECT * from mobiliario order by codigomob"
Set rdoAux = cnBkp.OpenResultset(sql, rdOpenKeyset, rydConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        sql = "select * from mobiliario where codigomob=" & !codigomob
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            If RdoAux2!CodLogradouro <> !CodLogradouro Or RdoAux2!CodBairro <> !CodBairro Then
                DoEvents
                If Val(SubNull(!CodLogradouro)) > 0 Then
                 '   MsgBox !CODigomob & " BKP:" & !CodLogradouro & "-" & !CodBairro & " -> Atual:" & RdoAux2!CodLogradouro & "-" & RdoAux2!CodBairro
                sql = "update mobiliario set codlogradouro=" & !CodLogradouro & ",codbairro=" & IIf(IsNull(!CodBairro), 999, !CodBairro) & " where codigomob=" & !codigomob
                cn.Execute sql, rdExecDirect
               End If
            End If
        Else
'            MsgBox "Não existe:" & !CODREDUZIDO
'            Sql = "insert bairro select siglauf,codcidade,codbairro,descbairro from tributacaobkp..bairro where "
'            Sql = Sql & "siglauf='SP' and codcidade=413 and codbairro=" & !CodBairro
'            cn.Execute Sql, rdExecDirect
        End If
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With
GoTo Fim

FASE8:
sql = "SELECT * from mobiliarioendentrega order by codmobiliario"
Set rdoAux = cnBkp.OpenResultset(sql, rdOpenKeyset, rydConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        sql = "select * from mobiliarioendentrega where codmobiliario=" & !codmobiliario
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            If RdoAux2!CodLogradouro <> !CodLogradouro Or RdoAux2!CodBairro <> !CodBairro Then
                DoEvents
                If Val(SubNull(!CodLogradouro)) > 0 Then
                  '  MsgBox !CODmobiliario & " BKP:" & !CodLogradouro & "-" & !CodBairro & " -> Atual:" & RdoAux2!CodLogradouro & "-" & RdoAux2!CodBairro
                sql = "update mobiliarioendentrega set codlogradouro=" & !CodLogradouro & ",codbairro=" & IIf(IsNull(!CodBairro), 999, !CodBairro) & " where codmobiliario=" & !codmobiliario
                cn.Execute sql, rdExecDirect
               End If
            End If
        Else
'            MsgBox "Não existe:" & !CODREDUZIDO
'            Sql = "insert bairro select siglauf,codcidade,codbairro,descbairro from tributacaobkp..bairro where "
'            Sql = Sql & "siglauf='SP' and codcidade=413 and codbairro=" & !CodBairro
'            cn.Execute Sql, rdExecDirect
        End If
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With
GoTo Fim
FASE9:
sql = "SELECT * from processoend order by Ano,numprocesso"
Set rdoAux = cnBkp.OpenResultset(sql, rdOpenKeyset, rydConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        sql = "select * from processoend where ano=" & !ano & " and numprocesso=" & !NumProcesso
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            If RdoAux2!CodLogr <> !CodLogr Then
                DoEvents
                On Error Resume Next
                If Val(SubNull(!CodLogr)) > 0 Then
        '            MsgBox !numprocesso & "/" & !Ano & " BKP:" & !CodLogr & " -> Atual:" & RdoAux2!CodLogr
                sql = "update processoend set codlogr=" & !CodLogr & " where ano=" & !ano & " and numprocesso=" & !NumProcesso & " and numero=" & !Numero
  '              cn.Execute Sql, rdExecDirect
               End If
            End If
        Else
'            MsgBox "Não existe:" & !CODREDUZIDO
'            Sql = "insert bairro select siglauf,codcidade,codbairro,descbairro from tributacaobkp..bairro where "
'            Sql = Sql & "siglauf='SP' and codcidade=413 and codbairro=" & !CodBairro
'            cn.Execute Sql, rdExecDirect
        End If
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With


Fim:
MsgBox "fim"

End Sub


Private Sub CorrigeProcesso()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nAno As Integer, nLc As Integer, nSq As Integer, nPc As Integer, nCp As Integer
Dim nTot As Long, nPos As Long, sNumProc As String, nNumproc As Long, nAnoproc As Integer


FASE1:
sql = "SELECT distinct NUMPROCESSO FROM origemreparc WHERE ANOPROC IS NULL ORDER BY NUMPROCESSO"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rydConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        sNumProc = !NumProcesso
        nNumproc = Val(Left$(sNumProc, InStr(1, sNumProc, "/", vbBinaryCompare) - 1))
        nAnoproc = Val(Right$(sNumProc, 4))
        
        sql = "update origemreparc set numproc=" & nNumproc & ",anoproc=" & nAnoproc & " where numprocesso='" & sNumProc & "'"
        cn.Execute sql, rdExecDirect

        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

FASE2:

sql = "SELECT distinct NUMPROCESSO FROM destinoreparc WHERE ANOPROC IS NULL ORDER BY NUMPROCESSO"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rydConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        sNumProc = !NumProcesso
        nNumproc = Val(Left$(sNumProc, InStr(1, sNumProc, "/", vbBinaryCompare) - 1))
        nAnoproc = Val(Right$(sNumProc, 4))
        
        sql = "update destinoreparc set numproc=" & nNumproc & ",anoproc=" & nAnoproc & " where numprocesso='" & sNumProc & "'"
        cn.Execute sql, rdExecDirect

        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With


Fim:
MsgBox "fim"

End Sub


Private Sub cmdGravar_Click()
Dim rdoAux As rdoResultset, x As Integer

s = Mid(lstParam.Text, 2, 6)

If txtOld.Text = txtValor.Text Then
    MsgBox "Nenhuma alteração foi feita neste parâmetro.", vbInformation, "Atenção"
    Exit Sub
Else
    If MsgBox("Deseja alterar o Valor de " & txtOld.Text & " para " & txtValor.Text & " ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
        sql = "UPDATE PARAMETROS SET VALPARAM='" & txtValor.Text & "' WHERE NOMEPARAM='" & s & "'"
        cn.Execute sql, rdExecDirect
        txtOld.Text = txtValor.Text
    End If
End If

End Sub


Private Sub cmdPagos_Click()

Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nAno As Integer, nLc As Integer, nSq As Integer, RdoAux3 As rdoResultset
Dim nPc As Integer, nCp As Integer, sCnae As String, nPos As Long, nTot As Long, nIni As Integer, nFim As Integer, sMotivo As String, nNumDoc As Long
If NomeDeLogin <> "SCHWARTZ" Then Exit Sub

BaixaEicon
'CancelaUnica
'EmiteBoleto
Exit Sub

frmComercioEletronico.BoletoNome = "DANIELA LAURENTIZ TERRA"
frmComercioEletronico.BoletoCidade = "CENTRO"
frmComercioEletronico.BoletoCep = "14887-888"
frmComercioEletronico.BoletoCpfCnpj = "151.729.278-67"
frmComercioEletronico.BoletoEndereco = "AV. MARECHAL DEODORO, 573 BLOCO A"
frmComercioEletronico.BoletoNumDoc = 15712545
frmComercioEletronico.BoletoUF = "PR"
frmComercioEletronico.BoletoValor = 12658.3
frmComercioEletronico.BoletoVencto = "17/01/2018"
frmComercioEletronico.show 1
'CorrigeRefis
'NaoPagoParaPago
'MsgBox "fim"


Exit Sub


sql = "truncate table taxa_lixo"
cn.Execute sql, rdExecDirect

sql = "SELECT areas.codreduzido, SUM(areas.areaconstr) AS soma, vwFULLIMOVEL.INSCRICAO, vwFULLIMOVEL.LOGRADOURO,    vwFULLIMOVEL.Li_Num , vwFULLIMOVEL.Li_Compl, vwFULLIMOVEL.DescBairro, vwFULLIMOVEL.CodLogr "
sql = sql & "FROM areas INNER JOIN vwFULLIMOVEL ON areas.codreduzido = vwFULLIMOVEL.codreduzido Where (vwFULLIMOVEL.Inativo = 0) GROUP BY areas.codreduzido, vwFULLIMOVEL.INSCRICAO, vwFULLIMOVEL.LOGRADOURO, vwFULLIMOVEL.li_num, vwFULLIMOVEL.li_compl,"
sql = sql & "vwFULLIMOVEL.DescBairro , vwFULLIMOVEL.CodLogr ORDER BY areas.codreduzido"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rydConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        On Error Resume Next
        sql = "select TOP (1) areas.codreduzido, usoconstr.descusoconstr FROM areas INNER JOIN usoconstr ON areas.usoconstr = usoconstr.codusoconstr Where CODREDUZIDO = " & !CODREDUZIDO
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        sql = "insert taxa_lixo (codigo_imovel,inscricao,codigo_logradouro,endereco,numero,complemento,bairro,cep,area,uso) values(" & !CODREDUZIDO & ",'" & !Inscricao & "'," & !CodLogr & ",'" & !Logradouro & "',"
        sql = sql & !Li_Num & ",'" & Left(SubNull(!Li_Compl), 50) & "','" & !DescBairro & "','" & RetornaCEP(!CodLogr, !Li_Num) & "'," & Virg2Ponto(CStr(Round(!soma, 2))) & ",'" & RdoAux2!descusoconstr & "')"
        cn.Execute sql, rdExecDirect

Proximo2:
        nPos = nPos + 1
       .MoveNext
    Loop
    
   .Close
End With
'PrintExcel
Fim:
MsgBox "fim"
End Sub


Private Sub Dir1_Change()

File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo Erro
Dir1.Path = "D:\Trabalho\GTI\Fotos"
Exit Sub
Erro:
MsgBox Err.Description
End Sub

Private Sub Form_Load()
'Dir1.Path = "D:\Trabalho\GTI\Documentos"

'File1.Path = "c:\trabalho\daf\Arq1\"
Centraliza Me
With lstParam
    .AddItem "(SEQ237) Sequência Arquivo DA Bradesco"
    .AddItem "(SEQ341) Sequência Arquivo DA Itaú"
    .AddItem "(SEQ409) Sequência Arquivo DA Unibanco"
    .AddItem "(SEQ033) Sequência Arquivo DO Banespa"
    .AddItem "(SEQ399) Sequência Arquivo DO HSBC"
End With
lstParam.ListIndex = 0

End Sub

Private Sub lstParam_Click()
Dim s As String, sql As String, rdoAux As rdoResultset

If lstParam.ListIndex = -1 Then Exit Sub
s = Mid(lstParam.Text, 2, 6)

sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='" & s & "'"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    If .RowCount > 0 Then
        txtValor.Text = !valparam
    Else
        txtValor.Text = "Não Cadastrado"
    End If
   .Close
End With
txtOld.Text = txtValor.Text

End Sub

Private Function FillSpace(sPalavra As String, nTamanho As Integer) As String
Dim sTmp As String

   On Error GoTo FillSpace_Error

If Len(sPalavra) > nTamanho Then sPalavra = Left(sPalavra, nTamanho)
sTmp = sPalavra & Space(nTamanho - Len(sPalavra))
FillSpace = sTmp

   On Error GoTo 0
   Exit Function

FillSpace_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FillSpace of Formulário frmConfig"

End Function

Private Sub CallPb(nPosF As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents

If ((nPosF * 100) / nTotal) <= 100 Then
   Pb.value = (nPosF * 100) / nTotal
Else
   Pb.value = 100
End If
lblPB.Caption = FormatNumber(Pb.value, 2)

'Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub

Private Function ConvDataSerial(sData As String) As String
If Len(sData) = 8 Then
   ConvDataSerial = Right$(sData, 2) & "/" & Mid$(sData, 5, 2) & "/" & Left$(sData, 4)
Else
   ConvDataSerial = Left$(sData, 2) & "/" & Mid$(sData, 3, 2) & "/20" & Right$(sData, 2)
End If
End Function

Public Function SNCheck(nCodigo As Long) As Boolean
Dim rdoAux As rdoResultset, sql As String
sql = "SELECT " & NomeBaseDados & ".dbo.RETORNASN(" & Format(nCodigo, "000000") & ",'" & Format(Now, "mm/dd/yyyy") & "') AS RETORNO"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurReadOnly)
With rdoAux
     If rdoAux!RETORNO = 1 Then
        SNCheck = True
     Else
        SNCheck = False
     End If
    .Close
End With

End Function

Private Sub simples_ano()
sql = "delete from simples_ano"
cn.Execute sql, rdExecDirect
ReDim aAno(0)
sql = "SELECT * from periodosn order by codigo,dataini"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    Do Until .EOF
        
        nCodReduz = !Codigo
        nAnoIni = Year(!DATAINI)
        If Not IsNull(!datafim) Then
            nAnoFim = Year(!datafim)
             
            For y = nAnoIni To nAnoFim
                nAnoTmp = y
                GoSub AddMatrix
            Next
        Else
            nAnoTmp = nAnoIni
            GoSub AddMatrix
        End If
        
       .MoveNext
        DoEvents
    Loop
   .Close
   
    For x = 1 To UBound(aAno)
        sql = "insert simples_ano (codigo,ano) values(" & aAno(x).nCodigo & "," & aAno(x).nAno & ")"
        cn.Execute sql, rdExecDirect
    Next
   
 MsgBox "fim"
Exit Sub

AddMatrix:
    bFind = False
    For x = 0 To UBound(aAno)
        If aAno(x).nCodigo = nCodReduz And aAno(x).nAno = nAnoTmp Then
            bFind = True
            Exit For
        End If
    Next
    If Not bFind Then
        ReDim Preserve aAno(UBound(aAno) + 1)
        aAno(UBound(aAno)).nCodigo = nCodReduz
        aAno(UBound(aAno)).nAno = nAnoTmp
    End If
    Return
   
End With
End Sub

Private Sub PrintExcel()

If lvMain.ListItems.Count = 0 Then Exit Sub

Dim x As Long, y As Long, ax As String, Scr_hdc As Long, z As Long
Dim cnExcel As ADODB.Connection, Rs As ADODB.Recordset, nCont As Integer, sFile As String
Scr_hdc = GetDesktopWindow()
Set cnExcel = New ADODB.Connection
sFile = "Rel" & Format(Now, "ddmmyyyyhhmmss") & ".xls"
cnExcel.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0; data source=" & sPathBin & "\" & sFile & "; Extended Properties=""Excel 8.0;HDR=YES"""
cnExcel.Open

ax = ""
For y = 1 To lvMain.ColumnHeaders.Count
    ax = ax & RemoveSpace(lvMain.ColumnHeaders(y).Text) & " char(255), "
Next
ax = Left(ax, Len(ax) - 2)
cnExcel.Execute "Create Table Table1(" & ax & ")"

Set Rs = New ADODB.Recordset
Rs.Open "Table1$", cnExcel, adOpenDynamic, adLockOptimistic, adCmdTable


For x = 1 To lvMain.ListItems.Count
    Rs.AddNew
    nCont = 0
    Rs.Fields(nCont).value = lvMain.ListItems(x).Text
    nCont = nCont + 1
    For y = 2 To lvMain.ColumnHeaders.Count
         
         Rs.Fields(nCont).value = lvMain.ListItems(x).SubItems(y - 1)
         nCont = nCont + 1
    
        
    Next
    Rs.Update
Next


 cnExcel.Close
Set Rs = Nothing
Set cnExcel = Nothing

z = ShellExecute(Scr_hdc, "Open", sFile, "", sPathBin, SW_SHOWNORMAL)


End Sub

Private Sub LeArquivo(sFullPath As String, sArq As String, nCodBanco As Integer, sDataCredito As String)

Dim sReg As String, FF1 As Integer, bExec As Boolean, sTipoArq As String, kk As Integer
Dim nValorPrincipal As Double, nValorJuros As Double, nValorMulta As Double, nValorGuia As Double, nNumDoc As Long, nErro As Integer, RdoAux4 As rdoResultset
Dim sAno As String, sMes As String, sAgencia As String, bLayoutNovo As Boolean, rdoAux As rdoResultset, RdoAux2 As rdoResultset, sql As String, RdoAux3 As rdoResultset
Dim nNumParc As Integer, bAchou As Boolean, nSeq As Integer, nCompl As Integer, nCodReduz As Long, sDataVencto As String, nRetorno As Integer, sRetorno As String
Dim nValorEfetivo As Double, nSeqReg As Integer, itmX As ListItem, nValorTaxa As Double, R As Integer, sDataGeracao As String, sLinhaT As String, sLinhaU As String, aRegistro() As Registro, aDoc() As Documento


ReDim aRegistro(0): ReDim aDoc(0)
nSeqReg = 1

'*** VERIFICA EXISTENCIA DO ARQUIVO

sFullPath = Replace(sFullPath, "/", "\")

If Dir$(sFullPath) = "" Then
    MsgBox "Não localizado o arquivo em " & sFullPath, vbCritical, "ERRO FATAL !!!"
    Exit Sub
End If


Ocupado

sReg = ""

'*****************************************
'****** ARQUIVO DO SIMPLES NACIONAL ******
'*****************************************
FF1 = FreeFile()
Open sFullPath For Binary Access Read Write As FF1

    While Not EOF(FF1)
        On Error GoTo CloseFile2
        bExec = False
        If Left(sReg, 1) = "9" Then GoTo CloseFile2
        Input #FF1, sReg
        If Left(sReg, 1) = "1" Then
            sSeqArq = Mid(sReg, 2, 8)
        ElseIf Left(sReg, 1) = "2" Then
           'LE OS REGISTROS
            With grdReg
                nValorPrincipal = CDbl(Mid(sReg, 107, 17)) / 100
                nValorJuros = CDbl(Mid(sReg, 124, 17)) / 100
                nValorMulta = CDbl(Mid(sReg, 141, 17)) / 100
                nValorGuia = nValorPrincipal + nValorJuros + nValorMulta
                sAno = Mid(sReg, 101, 4)
                sMes = Mid(sReg, 105, 2)
                sAgencia = Mid(sReg, 223, 4)
                sCnpj = Mid(sReg, 75, 14)
                
                nSeq = 0
                bAchou = False
                For R = 1 To UBound(aRegistro)
                    If aRegistro(R).sCnpj = Mid(sReg, 75, 14) Then
                        bAchou = True
                        nSeq = nSeq + 1
                    End If
                Next
                
                ReDim Preserve aRegistro(UBound(aRegistro) + 1)
                aRegistro(UBound(aRegistro)).sDataVencto = ConvDataSerial(Mid(sReg, 18, 8))
                aRegistro(UBound(aRegistro)).sDataCred = ConvDataSerial(Mid(sReg, 10, 8))
                aRegistro(UBound(aRegistro)).sDataPag = ConvDataSerial(Mid(sReg, 10, 8))
                aRegistro(UBound(aRegistro)).nValorPago = nValorGuia
                aRegistro(UBound(aRegistro)).sAgencia = Mid(sReg, 223, 4)
                aRegistro(UBound(aRegistro)).sCnpj = Mid(sReg, 75, 14)
                aRegistro(UBound(aRegistro)).nAno = Val(Mid(sReg, 101, 4))
                aRegistro(UBound(aRegistro)).nMes = Val(Mid(sReg, 105, 2))
                aRegistro(UBound(aRegistro)).nValorTarifaBancaria = 0
                aRegistro(UBound(aRegistro)).sSitRetorno = "CNPJ: " & Format(aRegistro(UBound(aRegistro)).sCnpj, "0#\.###\.###/####-##")
                aRegistro(UBound(aRegistro)).bExiste = True
                aRegistro(UBound(aRegistro)).sDataPagCalc = RetornaDiaUtil(CDate(aRegistro(UBound(aRegistro)).sDataPag))
                If Not bAchou Then
                    aRegistro(UBound(aRegistro)).nSeq = 0
                Else
                    aRegistro(UBound(aRegistro)).nSeq = nSeq
                End If
                
                With aRegistro(UBound(aRegistro))
                    
                    
                    Set itmX = lvMain.ListItems.Add(, , sFullPath)
                    itmX.SubItems(1) = sArq
                    itmX.SubItems(2) = nCodBanco
                    itmX.SubItems(3) = sDataCredito
                    itmX.SubItems(4) = .sCnpj
                    itmX.SubItems(6) = .nAno
                    itmX.SubItems(7) = .nMes
'                    itmX.SubItems(5) = !CODREDUZIDO
                    itmX.SubItems(8) = aRegistro(UBound(aRegistro)).sDataVencto
                    itmX.SubItems(9) = nValorGuia

                    
                End With
                'PROCURA SE O DEBITO JA FOI BAIXADO
                sql = "SELECT * FROM COMPLEMENTOSIMPLES WHERE ARQUIVOBANCO='" & sArq & "' AND DATACREDITO='" & Format(ConvDataSerial(Mid(sReg, 10, 8)), "mm/dd/yyyy") & "' AND "
                sql = sql & "CNPJ='" & Mid(sReg, 75, 14) & "' AND ANO=" & Val(Mid(sReg, 101, 4)) & " AND MES=" & Val(Mid(sReg, 105, 2))
                Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                With rdoAux
                    If .RowCount > 0 Then
                        'CARREGA PARCELA GRAVADA
                        ReDim Preserve aDoc(UBound(aDoc) + 1)
                        aDoc(UBound(aDoc)).sCnpj = aRegistro(UBound(aRegistro)).sCnpj
                        aDoc(UBound(aDoc)).nCodReduz = !CODREDUZIDO
                        aDoc(UBound(aDoc)).nAno = !AnoExercicio
                        aDoc(UBound(aDoc)).nLanc = !CodLancamento
                        aDoc(UBound(aDoc)).nSeq = !SeqLancamento
                        aDoc(UBound(aDoc)).nParc = !NumParcela
                        aDoc(UBound(aDoc)).nCompl = !CODCOMPLEMENTO
                        aDoc(UBound(aDoc)).sDataVencto = aRegistro(UBound(aRegistro)).sDataVencto
                        aDoc(UBound(aDoc)).sSit = 2
                        aDoc(UBound(aDoc)).nValorPrincipal = nValorPrincipal
                        aDoc(UBound(aDoc)).nValorMulta = nValorMulta
                        aDoc(UBound(aDoc)).nValorJuros = nValorJuros
                        aDoc(UBound(aDoc)).nValorCorrecao = 0
                        aDoc(UBound(aDoc)).nValorTotal = nValorGuia
                        aDoc(UBound(aDoc)).nValorTarifa = 0
                        aDoc(UBound(aDoc)).nValorDif = 0
                        aDoc(UBound(aDoc)).nValorCompensado = nValorGuia
                        aDoc(UBound(aDoc)).sBx = "S"
                        aDoc(UBound(aDoc)).sDp = "N"
                        aDoc(UBound(aDoc)).bExiste = True
                        aDoc(UBound(aDoc)).nSeqReg = aRegistro(UBound(aRegistro)).nSeq
                    Else
                        'DEFINIR NOVA PARCELA
                        'BUSCA CÓDIGO
                        sql = "SELECT CODIGOMOB,CNPJ FROM MOBILIARIO WHERE DATAENCERRAMENTO IS NULL and CONVERT(BIGINT, cnpj) = " & Val(aRegistro(UBound(aRegistro)).sCnpj)
                        sql = sql & " OR CNPJ='" & Format(aRegistro(UBound(aRegistro)).sCnpj, "00\.000\.000/0000-00") & "' AND DATAENCERRAMENTO IS NULL ORDER BY CODIGOMOB DESC"
                        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux2
                            If .RowCount > 0 Then
                                nCodReduz = !codigomob
                                .Close
                            Else
                                .Close
                                sql = "SELECT CODCIDADAO,CNPJ FROM CIDADAO WHERE CNPJ = '" & RetornaNumero(aRegistro(UBound(aRegistro)).sCnpj) & "' OR "
                                sql = sql & "CNPJ='" & Format(aRegistro(UBound(aRegistro)).sCnpj, "00\.000\.000/0000-00") & "' ORDER BY CODCIDADAO DESC"
                                Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                                With RdoAux2
                                    If .RowCount > 0 Then
                                        nCodReduz = !CodCidadao
                                    Else
                                        'CNPJ NÃO LOCALIZADO
                                        aRegistro(UBound(aRegistro)).bExiste = False
                                        sql = "SELECT * FROM SIMPLESCNPJ WHERE CNPJ='" & aRegistro(UBound(aRegistro)).sCnpj & "' AND ANOCOMP=" & aRegistro(UBound(aRegistro)).nAno & " AND MESCOMP=" & aRegistro(UBound(aRegistro)).nMes
                                        Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                                        If RdoAux3.RowCount = 0 Then
                                            sql = "INSERT SIMPLESCNPJ (CNPJ,ARQUIVOSHORT,BANCO,DATAARRECADA,DATAVENCTO,ANOCOMP,MESCOMP,PRINCIPAL,JUROS,"
                                            sql = sql & "MULTA,AGENCIA,CODREDUZIDO) VALUES('" & RetornaNumero(aRegistro(UBound(aRegistro)).sCnpj) & "','" & lstArq.Text & "'," & Val(Left(lblBanco.Caption, 3)) & ",'" & Format(aRegistro(UBound(aRegistro)).sDataCred, "mm/dd/yyyy") & "','"
                                            sql = sql & Format(aRegistro(UBound(aRegistro)).sDataVencto, "mm/dd/yyyy") & "'," & aRegistro(UBound(aRegistro)).nAno & "," & aRegistro(UBound(aRegistro)).nMes & "," & Virg2Ponto(CStr(aRegistro(UBound(aRegistro)).nValorPago)) & "," & Virg2Ponto(0) & "," & Virg2Ponto(0) & ",'"
                                            sql = sql & aRegistro(UBound(aRegistro)).sAgencia & "'," & 0 & ")"
                      '                      cn.Execute Sql, rdExecDirect
                                        End If
                                        RdoAux3.Close
                                        GoTo CONTSN
                                    End If
                                    
                                    
                                    
                                End With
                            End If
                           
                        End With
                                
                        'BUSCA LANCAMENTO
                         sql = "SELECT debitoparcela.codreduzido,debitoparcela.anoexercicio, debitoparcela.codlancamento,DEBITOPARCELA.SEQLANCAMENTO,debitoparcela.numparcela,DEBITOPARCELA.CODCOMPLEMENTO,debitoparcela.datavencimento, debitoparcela.statuslanc, debitotributo.valortributo "
                         sql = sql & "FROM debitoparcela INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
                         sql = sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND debitoparcela.NumParcela = debitotributo.NumParcela And "
                         sql = sql & "debitoparcela.CODCOMPLEMENTO = debitotributo.CODCOMPLEMENTO WHERE (debitoparcela.codreduzido = " & nCodReduz & ") AND (debitoparcela.codlancamento = 5) AND (MONTH(debitoparcela.datavencimento) = " & Month(CDate(aRegistro(UBound(aRegistro)).sDataVencto)) & ") AND "
                         sql = sql & "(YEAR(debitoparcela.datavencimento) = " & Year(CDate(aRegistro(UBound(aRegistro)).sDataVencto)) & ") AND (debitotributo.codtributo = 13) and debitotributo.valortributo =" & Virg2Ponto(CStr(nValorGuia)) & " AND statuslanc<>6"
                         Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                         With RdoAux3
                            'EXISTE LANCAMENTO NESTE MÊS/ANO?
                             If .RowCount > 0 Then 'SIM
                                 nNumParc = !NumParcela 'CAPTURA A PARCELA
                                 bAchou = False
                                'TEM ALGUMA QUE NÃO ESTA PAGA?
                                 Do Until .EOF
                                     If !statuslanc = 3 Then
                                         bAchou = True
                                         Exit Do
                                     End If
                                    .MoveNext
                                 Loop
                                 
                                'SE ACHOU PEGA A PARCELA
                                 If bAchou Then
                                     nSeq = !SeqLancamento
                                     nCompl = !CODCOMPLEMENTO '---------------> PARCELA PRONTA PARA USO
                                 Else
                                    'SE NÃO ACHAR
                                    .MoveFirst
                                     nCompl = 0
                                    'BUSCAR A ÚLTIMA SEQUENCIA DE LANCAMENTO PARA EVITAR DUPLICIDADE
                                     sql = "SELECT MAX(SEQLANCAMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE (codreduzido = " & nCodReduz & ") AND (codlancamento = 5) AND (MONTH(datavencimento) = " & Month(dDataVencto) & ") AND "
                                     sql = sql & "(YEAR(datavencimento) = " & Year(dDataVencto) & ")"
                                     Set RdoAux4 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                                     With RdoAux4
                                         If IsNull(!maximo) Then
                                             nSeq = 0
                                         Else
                                             nSeq = !maximo + 1
                                         End If
                                        .Close
                                     End With
                                 End If
                             
                             Else
                                'NÃO ACHOU LANCAMENTOS NESTE MÊS/ANO
                                'AUMENTA O LANCAMENTO
                                 sql = "SELECT MAX(SEQLANCAMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE (codreduzido = " & nCodReduz & ") AND (codlancamento = 5) AND (ANOEXERCICIO = " & Val(sAno) & ")"
                                 Set RdoAux4 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                                 With RdoAux4
                                     If IsNull(!maximo) Then
                                         nSeq = 1
                                     Else
                                         nSeq = !maximo + 1
                                     End If
                                    .Close
                                 End With
                                 'VERIFICA SE A SEQ JA NÃO EXISTE NA MATRIZ
                                 For R = 1 To UBound(aDoc)
                                    If aDoc(R).nCodReduz = nCodReduz And aDoc(R).nAno = Val(sAno) Then
                                        nSeq = aDoc(R).nSeq + 1
                                    End If
                                 Next
                                 
                                 nCompl = 0
                                 nNumParc = 1
                             End If
                             ReDim Preserve aDoc(UBound(aDoc) + 1)
                             aDoc(UBound(aDoc)).sCnpj = aRegistro(UBound(aRegistro)).sCnpj
                             aDoc(UBound(aDoc)).nCodReduz = nCodReduz
                             aDoc(UBound(aDoc)).nAno = Val(sAno)
                             aDoc(UBound(aDoc)).nLanc = 5
                             aDoc(UBound(aDoc)).nSeq = nSeq
                             aDoc(UBound(aDoc)).nParc = nNumParc
                             aDoc(UBound(aDoc)).nCompl = nCompl
                             aDoc(UBound(aDoc)).sDataVencto = aRegistro(UBound(aRegistro)).sDataVencto
                             aDoc(UBound(aDoc)).sSit = 3
                             aDoc(UBound(aDoc)).nValorPrincipal = nValorPrincipal
                             aDoc(UBound(aDoc)).nValorMulta = nValorMulta
                             aDoc(UBound(aDoc)).nValorJuros = nValorJuros
                             aDoc(UBound(aDoc)).nValorCorrecao = 0
                             aDoc(UBound(aDoc)).nValorTotal = nValorGuia
                             aDoc(UBound(aDoc)).nValorTarifa = 0
                             aDoc(UBound(aDoc)).nValorDif = 0
                             aDoc(UBound(aDoc)).nValorCompensado = nValorGuia
                             aDoc(UBound(aDoc)).sBx = ""
                             aDoc(UBound(aDoc)).sDp = ""
                             aDoc(UBound(aDoc)).bExiste = True
                             aDoc(UBound(aDoc)).nSeqReg = aRegistro(UBound(aRegistro)).nSeq
                            itmX.SubItems(5) = nCodReduz
                            itmX.SubItems(10) = Val(sAno)
                            itmX.SubItems(11) = nSeq
                            sql = "SELECT debitoparcela.codreduzido,debitotributo.valortributo From  debitoparcela INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
                            sql = sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND debitoparcela.NumParcela = debitotributo.NumParcela And debitoparcela.CODCOMPLEMENTO = debitotributo.CODCOMPLEMENTO "
                            sql = sql & "WHERE debitoparcela.codreduzido = " & nCodReduz & " AND debitoparcela.anoexercicio = " & Val(sAno) & " AND (debitoparcela.codlancamento = 5) and codtributo=13 AND datavencimento = '" & Format(CDate(aRegistro(UBound(aRegistro)).sDataVencto), "mm/dd/yyyy") & "' and valortributo=" & Virg2Ponto(CStr(nValorGuia))
                            Set RdoAux4 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                            If RdoAux4.RowCount > 0 Then
                                itmX.SubItems(12) = "S"
                            Else
                                itmX.SubItems(12) = "N"
                            End If
                            RdoAux4.Close
                            .Close
                         End With
CONTSN:
'**********************************
                    End If
                   .Close
                End With
               
            End With
        ElseIf Left(sReg, 1) = "9" Then
           'LE O RODAPÉ DO ARQUIVO
'            lblNumReg.Caption = Format(Val(Mid(sReg, 10, 6)) - 2, "000000")
'            lblValorTotal.Caption = FormatNumber(CDbl(Mid(sReg, 16, 17) / 100), 2)
        End If
         
        
        'nPos = nPos + 1
    Wend
CloseFile2:
Close #FF1
'Pb.value = 0
Liberado
nErro = 0




End Sub


Private Sub cmdPagosold_Click()

Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nAno As Integer, nLc As Integer, nSq As Integer
Dim nPc As Integer, nCp As Integer, sEvento As String, nPos As Long, nTot As Long, nIni As Integer, nFim As Integer, sMotivo As String
If NomeDeLogin <> "SCHWARTZ" Then Exit Sub


GoTo Rotina2
sql = "SELECT DISTINCT seq, datahoraevento, computador, usuario, form, evento, secevento, logevento "
sql = sql & "From logevento WHERE (evento = 3) AND (secevento = 2) order by seq"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rydConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        sEvento = !LOGEVENTO
        nIni = InStr(sEvento, " em ")
        If nIni = 0 Then GoTo Proximo
        nFim = InStr(sEvento, " Ano:")
        If nFim = 0 Then GoTo Proximo
        nCodReduz = Val(Mid(sEvento, nIni + 4, nFim - nIni - 4))
        
        If nCodReduz < 500000 Or nCodReduz > 700000 Then GoTo Proximo
        sMotivo = Left(sEvento, nIni - 1)
        
        nAno = Val(Mid(sEvento, InStr(sEvento, "Ano:") + 4, 4))
        nLc = Val(Mid(sEvento, InStr(sEvento, "Lc:") + 3, 2))
        nSq = Val(Mid(sEvento, InStr(sEvento, "Sq:") + 3, 2))
        nPc = Val(Mid(sEvento, InStr(sEvento, "Pc:") + 3, 2))
        nCp = Val(Mid(sEvento, InStr(sEvento, "Cp:") + 3, 2))
        sEvento = sMotivo & " - Ex:" & nAno & " Lc:" & nLc & " Sq:" & nSq & " Pc:" & nPc & " Cp:" & nCp
        
        sql = "insert historicocidadao(codigo,data,usuario,obs) values(" & nCodReduz & ",'" & Format(!DATAHORAEVENTO, sDataFormat & " hh:mm:ss") & "','" & !USUARIO & "','" & sEvento & "')"
        'cn.Execute Sql, rdExecDirect

Proximo:
        nPos = nPos + 1
       .MoveNext
    Loop
    
   .Close
End With

Rotina2:
sql = "SELECT DISTINCT seq, datahoraevento, computador, usuario, form, evento, secevento, logevento "
sql = sql & "From logevento WHERE (evento = 3) AND (secevento = 3) order by seq"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rydConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        sEvento = !LOGEVENTO
        nIni = InStr(sEvento, "Código:")
        If nIni = 0 Then GoTo Proximo2
        nFim = InStr(sEvento, "Lançamento:")
        If nFim = 0 Then GoTo Proximo2
        nCodReduz = Val(Mid(sEvento, InStr(sEvento, "Código:") + 7, 6))
        'nCodReduz = Val(Mid(sEvento, nIni + 4, nFim - nIni - 4))
        
        If nCodReduz < 500000 Or nCodReduz > 700000 Then GoTo Proximo2
        'sMotivo = Left(sEvento, nIni - 1)
        sMotivo = Mid(sEvento, InStr(sEvento, "Motivo:") + 8, Len(sEvento) - InStr(sEvento, "Motivo:"))
        nAno = Val(Mid(sEvento, InStr(sEvento, "Ano:") + 4, 4))
        nLc = Val(Mid(sEvento, InStr(sEvento, "Lançamento:") + 11, 2))
        nSq = Val(Mid(sEvento, InStr(sEvento, "Seq:") + 4, 2))
        nPc = Val(Mid(sEvento, InStr(sEvento, "Parcela:") + 8, 2))
        nCp = Val(Mid(sEvento, InStr(sEvento, "Compl:") + 5, 2))
        sEvento = sMotivo & " - Ex:" & nAno & " Lc:" & nLc & " Sq:" & nSq & " Pc:" & nPc & " Cp:" & nCp
        
        'Sql = "insert historicocidadao(codigo,data,usuario,obs) values(" & nCodReduz & ",'" & Format(!DATAHORAEVENTO, sDataFormat & " hh:mm:ss") & "','" & !Usuario & "','" & Mask(sEvento) & "')"
        sql = "insert historicocidadao(codigo,data,userid,obs) values(" & nCodReduz & ",'" & Format(!DATAHORAEVENTO, sDataFormat & " hh:mm:ss") & "'," & RetornaUsuarioID(!USUARIO) & ",'" & Mask(sEvento) & "')"
        cn.Execute sql, rdExecDirect

Proximo2:
        nPos = nPos + 1
       .MoveNext
    Loop
    
   .Close
End With
'PrintExcel

MsgBox "fim"
End Sub

Private Sub EmiteBoleto()

Dim nValorTaxa As Double, x As Integer, nSituacao As Integer, dDataProc As Date, sDescImposto As String, RdoAux2 As rdoResultset, y As Integer, nPercTrib As Double
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, sDataVencto As String, nCodTrib As Integer, nValorTributo As Double
Dim NumBarra1 As String, StrBarra1 As String, NumBarra2 As String, NumBarra2a As String, NumBarra2b As String, NumBarra2c As String, NumBarra2d As String, StrBarra2 As String
Dim sCodReduz As String, sNomeResp As String, sTipoImposto As String, sEndImovel As String, nNumImovel As Integer, sComplImovel As String, sBairroImovel As String
Dim nCodLogr As Long, sEndEntrega As String, nNumEntrega As Integer, sBairroEntrega As String, sComplEntrega As String, sCepEntrega As String, sCidadeEntrega As String
Dim sUFEntrega As String, sNumInsc As String, sValorParc As String, nNumDoc As Long, sBarra As String, sDigitavel2 As String, nValorGuia As Double, nNumGuia As Long
Dim sNumDoc2 As String, sNumDoc3 As String, nFatorVencto As Long, sNumDoc As String, nSid As Long, sDigitavel As String, sNossoNumero As String, sCPF As String, sObs As String
Dim clsImovel As New clsImovel, nCodReduz As Long, sSetor As String, sRG As String, dDataPrimeiraParc As String, nValorTotalHon As Double, RdoAux3 As rdoResultset
Dim nPagina As Integer, nLivro As Integer, sDataDam As String, xImovel As clsImovel


'LIMPA TEMPORARIO
nSid = Int(Rnd(100) * 1000000)

sql = "delete from boletoguia where sid=" & nSid
cn.Execute sql, rdExecDirect

sql = "delete from boletoguiacapa where sid=" & nSid
cn.Execute sql, rdExecDirect

sLib = "LIBERACAO"


'sNumProc = lblNumProc.Caption & "/" & lblAnoProc.Caption
'dDataProc = lblDataParc.Caption
sql = "SELECT cadimob.codreduzido, parceladocumento.anoexercicio, parceladocumento.codlancamento, parceladocumento.seqlancamento, parceladocumento.numparcela, "
sql = sql & "parceladocumento.CODCOMPLEMENTO , parceladocumento.NumDocumento, debitoparcela.DataVencimento, debitotributo.ValorTributo FROM cadimob INNER JOIN "
sql = sql & "parceladocumento ON cadimob.codreduzido = parceladocumento.codreduzido INNER JOIN debitoparcela ON parceladocumento.codreduzido = debitoparcela.codreduzido AND parceladocumento.anoexercicio = debitoparcela.anoexercicio AND "
sql = sql & "parceladocumento.codlancamento = debitoparcela.codlancamento AND parceladocumento.seqlancamento = debitoparcela.seqlancamento AND parceladocumento.numparcela = debitoparcela.numparcela AND parceladocumento.codcomplemento = debitoparcela.codcomplemento INNER JOIN "
sql = sql & "debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
sql = sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND "
sql = sql & "debitoparcela.NumParcela = debitotributo.NumParcela And debitoparcela.CODCOMPLEMENTO = debitotributo.CODCOMPLEMENTO "
'Sql = Sql & "Where (cadimob.codreduzido in (35654,35565,35566)) "
sql = sql & "Where (cadimob.li_codbairro =1069) "
sql = sql & " And (parceladocumento.AnoExercicio = 2018) And (parceladocumento.CodLancamento =1) And (parceladocumento.SeqLancamento = 0) "
sql = sql & "AND  statuslanc=3 ORDER BY cadimob.codreduzido, parceladocumento.numparcela"

Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    x = 1
   Set xImovel = New clsImovel
    Do Until .EOF
        DoEvents
        nCodReduz = !CODREDUZIDO
        sTipoImposto = "Cont.Ilum.Pub."
        sSetor = "IMOBILIÁRIO"
        xImovel.CarregaImovel nCodReduz
        sNumInsc = xImovel.Inscricao
        sCodReduz = nCodReduz
        sNomeResp = xImovel.NomePropPrincipal
        sQuadra = xImovel.Li_Quadras
        sLote = xImovel.Li_Lotes
        xImovel.RetornaEndereco nCodReduz, Imobiliario, Localizacao
        sEndImovel = xImovel.Endereco
        nNumImovel = xImovel.Numero
        sComplImovel = xImovel.Complemento
        sBairroImovel = xImovel.Bairro
        sEndEntrega = xImovel.Ee_NomeLog
        nNumEntrega = xImovel.Ee_NumImovel
        sComplEntrega = xImovel.Ee_Complemento
        sBairroEntrega = xImovel.Ee_Bairro
        sCidadeEntrega = "JABOTICABAL"
        sUFEntrega = "SP"
        sCepEntrega = xImovel.Ee_Cep
        sql = "SELECT cidadao.codcidadao, cidadao.nomecidadao, cidadao.cpf, cidadao.cnpj, cidadao.rg "
        sql = sql & "FROM cidadao INNER JOIN proprietario ON cidadao.codcidadao = proprietario.codcidadao "
        sql = sql & "WHERE(proprietario.codreduzido = " & nCodReduz & ") AND (proprietario.tipoprop = 'P') AND (proprietario.principal = 1)"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                sCPF = SubNull(!cpf)
                If Trim(sCPF) = "" Then
                   sCPF = SubNull(!Cnpj)
                End If
             Else
                sCPF = ""
             End If
             sRG = SubNull(!rg)
            .Close
        End With
        
        nAno = !AnoExercicio
        nLanc = !CodLancamento
        nSeq = !SeqLancamento
        nParc = !NumParcela
        nCompl = !CODCOMPLEMENTO
        sDataDam = Format(!DataVencimento, "dd/mm/yyyy")
        nNumDoc = !NumDocumento
        nValorParc = !VALORTRIBUTO

        nNumGuia = nNumDoc

        sNumDoc = CStr(nNumGuia) & "-" & RetornaDVNumDoc(nNumGuia)
        sNumDoc2 = CStr(nNumGuia) & RetornaDVNumDoc(nNumGuia)
        sNumDoc3 = CStr(nNumGuia) & Modulo11(nNumGuia)


        
        sValorParc = Format(nValorParc, "#0.00")
        nValorGuia = sValorParc
        nValorDoc = nValorGuia
    '**** GERADOR DE CÓDIGO DE BARRAS ********
    sNossoNumero = "2678478"
    sDigitavel = "001900000"
    sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
    sDigitavel = sDigitavel & sDv & "0" & sNossoNumero & "01"
    sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
    sDigitavel = sDigitavel & sDv & Right(sNumDoc3, 8) & "18"
    sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
    sDigitavel = sDigitavel & sDv
    
    dDataBase = "07/10/1997"
    nFatorVencto = CDate(sDataDam) - CDate(dDataBase)
    
    If CDate(sDataVencto) >= "22/02/2025" Then
        dDataBase = "29/05/2022"
        nFatorVencto = CDate(sDataVencto) - CDate(dDataBase)
    End If
    
    sQuintoGrupo = Format(nFatorVencto, "0000")
    sQuintoGrupo = sQuintoGrupo & Format(RetornaNumero(FormatNumber(nValorDoc, 2)), "0000000000")
    sBarra = "0019" & Format(nFatorVencto, "0000") & Format(RetornaNumero(FormatNumber(nValorDoc, 2)), "0000000000") & "00000026784780"
    sBarra = sBarra & sNumDoc3 & "18"
    sDv = Trim(Calculo_DV11(sBarra))
    sBarra = Left(sBarra, 4) & sDv & Mid(sBarra, 5, Len(sBarra) - 4)
    
    sDigitavel = sDigitavel & sDv & sQuintoGrupo
    
    sDigitavel2 = Left(sDigitavel, 5) & "." & Mid(sDigitavel, 6, 5) & " " & Mid(sDigitavel, 11, 5) & "." & Mid(sDigitavel, 16, 6) & " "
    sDigitavel2 = sDigitavel2 & Mid(sDigitavel, 22, 5) & "." & Mid(sDigitavel, 27, 6) & " " & Mid(sDigitavel, 33, 1) & " " & Right(sDigitavel, 14)
    sBarra = Gera2of5Str(sBarra)
    
    '*******************************************

        sql = "insert boletoguia(usuario,computer,sid,seq,codreduzido,nome,cpf,endereco,numimovel,complemento,bairro,cidade,uf,fulllanc,numdoc,numparcela,totparcela,datavencto,numdoc2,"
        sql = sql & "digitavel,codbarra,valorguia,obs,numproc) values('" & NomeDeLogin & "','" & NomeDoComputador & "'," & nSid & "," & x & "," & nCodReduz & ",'" & Left(Mask(sNomeResp), 80) & "','" & sCPF & "','"
        sql = sql & Left(Mask(sEndImovel), 80) & "'," & nNumImovel & ",'" & Left(Mask(sComplImovel), 30) & "','" & Left(Mask(sBairroImovel), 25) & "','" & Mask(sCidadeEntrega) & "','" & sUFEntrega & "','" & Mask(sDescImposto) & "','"
        sql = sql & CStr(nNumGuia) & "'," & IIf(nParc = 0, 0, nParc) & "," & 12 & ",'" & Format(!DataVencimento, "mm/dd/yyyy") & "','" & sNumDoc & "','" & sDigitavel2 & "','" & Mask(sBarra) & "',"
        sql = sql & Virg2Ponto(Format(nValorGuia, "#0.00")) & "," & "'','')"
        'cn.Execute Sql, rdExecDirect
        x = x + 1
       .MoveNext
    Loop
   .Close
End With

frmReport.ShowReport2 "boletoguia2", frmMdi.HWND, Me.HWND, nSid
Liberado

End Sub

Private Function SNCheck2(nCodigo As Long) As Boolean
Dim rdoAux As rdoResultset, sql As String, sReturn As Boolean
ConectaEicon
sql = "select * from  tb_inter_empr_snacional_giss Where NUM_CADASTRO=" & nCodigo & " order by timestamp desc"
Set rdoAux = cnEicon.OpenResultset(sql, rdOpenKeyset, rdConcurReadOnly)
With rdoAux
    If rdoAux.RowCount > 0 Then
        If IsNull(!Data_Fim) Then
            sReturn = True
        Else
            sReturn = False
        End If
     Else
        sReturn = False
     End If
    .Close
End With
cnEicon.Close
SNCheck2 = sReturn
End Function

Private Sub TransfereLancamento()
Dim sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, x As Integer
Dim nCodReduz As Long, nFicha As Long, sProc As String, nAno As Integer, nNumero As Long, sStatus As String, bCancelado As Boolean
Dim sCep As String

If NomeDeLogin <> "SCHWARTZ" Then Exit Sub
Pb.value = 0
nPos = 1

sql = "SELECT codreduzido, anoexercicio, codlancamento, seqlancamento, numparcela, codcomplemento From debitotributo Where CodTributo = 527 AND CODLANCAMENTO=11 ORDER BY codreduzido, anoexercicio, numparcela"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        sql = "update debitoparcela set codlancamento=81 where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and codlancamento=11 and "
        sql = sql & "seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO
        cn.Execute sql, rdExecDirect
        
        sql = "update debitotributo set codlancamento=81 where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and codlancamento=11 and "
        sql = sql & "seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO
        cn.Execute sql, rdExecDirect
        
        sql = "update parceladocumento set codlancamento=81 where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and codlancamento=11 and "
        sql = sql & "seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO
        cn.Execute sql, rdExecDirect
        
        sql = "update debitopago set codlancamento=81 where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and codlancamento=11 and "
        sql = sql & "seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO
        cn.Execute sql, rdExecDirect
        
        sql = "update obsparcela set codlancamento=81 where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and codlancamento=11 and "
        sql = sql & "seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO
        cn.Execute sql, rdExecDirect
        
        sql = "update origemreparc set codlancamento=81 where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and codlancamento=11 and "
        sql = sql & "numsequencia=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO
        cn.Execute sql, rdExecDirect
        
        
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
    rdoAux.Close
End With

MsgBox "fim"

Exit Sub
Erro:
MsgBox rdoErrors(0).Description

End Sub

Private Sub NaoPagoParaPago()
Dim sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long
Pb.value = 0
nPos = 1

sql = "SELECT DISTINCT  debitopago.codreduzido, debitopago.contacorrente, debitoparcela.statuslanc, debitopago.anoexercicio, debitopago.seqlancamento, debitopago.numparcela,"
sql = sql & "debitopago.CODCOMPLEMENTO , debitopago.CodLancamento FROM  debitopago INNER JOIN debitoparcela ON debitopago.codreduzido = debitoparcela.codreduzido AND debitopago.anoexercicio = debitoparcela.anoexercicio AND "
sql = sql & "debitopago.codlancamento = debitoparcela.codlancamento AND debitopago.seqlancamento = debitoparcela.seqlancamento AND "
sql = sql & "debitopago.NumParcela = debitoparcela.NumParcela And debitopago.CODCOMPLEMENTO = debitoparcela.CODCOMPLEMENTO Where (debitoparcela.statuslanc = 3)"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        sql = "update debitoparcela set statuslanc=2 where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and codlancamento=" & !CodLancamento & " and "
        sql = sql & "seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO
        cn.Execute sql, rdExecDirect
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub BaixaEicon()
Dim sql As String, rdoAux As rdoResultset, RdoAux3 As rdoResultset

ConectaEicon

sql = "SELECT codreduzido, anoexercicio, codlancamento, seqlancamento, numparcela, codcomplemento, seqpag, datapagamento, datarecebimento, valorpago,CodBanco,"
sql = sql & "CodAgencia, restituido, NumDocumento, valorpagoreal, intacto, ValorTarifa, arquivobanco, valordif, datapagamentocalc, dataintegracao, contacorrente "
sql = sql & "From debitopago WHERE (numdocumento BETWEEN 2000000 AND 2200000) AND (numdocumento NOT IN (SELECT num_documento FROM GTI_Eicon.dbo.tb_inter_baixa_detalhe)) "
sql = sql & " AND (anoexercicio > 2015) ORDER BY numdocumento"

'Sql = "SELECT codreduzido, anoexercicio, codlancamento, seqlancamento, numparcela, codcomplemento, seqpag, datapagamento, datarecebimento, valorpago, codbanco, "
'Sql = Sql & "CodAgencia , restituido, NumDocumento, valorpagoreal, intacto, ValorTarifa, arquivobanco, valordif, datapagamentocalc, dataintegracao, contacorrente "
'Sql = Sql & "From debitopago WHERE (codreduzido BETWEEN 100000 AND 300000) AND (datapagamento BETWEEN '03/01/2017' AND '03/31/2017') AND (codlancamento = 5)"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    Do Until .EOF
        DoEvents
        'baixa ok
        sql = "SELECT * FROM tb_inter_baixa_detalhe WHERE num_documento=" & nNumDoc
        Set RdoAux3 = cnEicon.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux3.RowCount = 0 Then
        
            '***** GRAVA BAIXA NA GISS ***************
            sql = "insert tb_inter_baixa(cod_cliente,cod_banco,num_sequencia,timestamp,data_geracao,nome_arquivo,data_movimento) values("
            sql = sql & 2177 & "," & !CodBanco & "," & 0 & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "','" & Format(Now, "mm/dd/yyyy") & "','"
            sql = sql & !arquivobanco & "','" & Format(!datarecebimento, "mm/dd/yyyy") & "')"
            cnEicon.Execute sql, rdExecDirect
            
            sql = "insert tb_inter_baixa_detalhe(cod_cliente,cod_banco,num_sequencia,num_documento,linha,timestamp,valor_titulo,valor_pago,data_pagamento,valor_encargos,"
            sql = sql & "descricao_linha_t,descricao_linha_u) values(" & 2177 & "," & !CodBanco & "," & 0 & "," & !NumDocumento & "," & !SEQPAG & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "',"
            sql = sql & Virg2Ponto(CStr(!ValorPagoreal)) & "," & Virg2Ponto(CStr(!ValorPagoreal)) & ",'" & Format(!DataPagamento, "mm/dd/yyyy") & "'," & 0 & ",'"
            sql = sql & "" & "','" & "" & "')"
            cnEicon.Execute sql, rdExecDirect
            
        End If
       .MoveNext
    Loop
   .Close
End With
                   
End Sub

Private Sub BaixaISSPagoPorDam()
Dim sql As String, rdoAux As rdoResultset, nNumDoc As Long, nNumDocISS As Long, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, nValorDoc As Double

ConectaEicon

sql = "SELECT DISTINCT docdam,dociss From damiss where baixado=0 ORDER BY docdam"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    Do Until .EOF
        nNumDoc = !docdam
        nNumDocISS = !dociss
        sql = "select * from debitopago where numdocumento=" & nNumDoc & " and codlancamento=5"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
    
            sql = "SELECT parceladocumento.codreduzido, debitotributo.valortributo FROM parceladocumento INNER JOIN debitotributo ON parceladocumento.codreduzido = debitotributo.codreduzido AND parceladocumento.anoexercicio = debitotributo.anoexercicio AND parceladocumento.codlancamento = debitotributo.codlancamento AND "
            sql = sql & "parceladocumento.SeqLancamento = debitotributo.SeqLancamento And parceladocumento.NumParcela = debitotributo.NumParcela And parceladocumento.CODCOMPLEMENTO = debitotributo.CODCOMPLEMENTO Where parceladocumento.NumDocumento = " & nNumDoc
            Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            nValorDoc = RdoAux3!VALORTRIBUTO
            RdoAux3.Close
'baixa ok
           '***** GRAVA BAIXA NA GISS ***************
            sql = "SELECT * FROM tb_inter_baixa_detalhe WHERE num_documento=" & nNumDoc
            Set RdoAux3 = cnEicon.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If RdoAux3.RowCount = 0 Then
            
                sql = "insert tb_inter_baixa(cod_cliente,cod_banco,num_sequencia,timestamp,data_geracao,nome_arquivo,data_movimento) values("
                sql = sql & 2177 & "," & RdoAux2!CodBanco & "," & 0 & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "','" & Format(Now, "mm/dd/yyyy") & "','"
                sql = sql & RdoAux2!arquivobanco & "','" & Format(RdoAux2!datarecebimento, "mm/dd/yyyy") & "')"
                cnEicon.Execute sql, rdExecDirect
                
                sql = "insert tb_inter_baixa_detalhe(cod_cliente,cod_banco,num_sequencia,num_documento,linha,timestamp,valor_titulo,valor_pago,data_pagamento,valor_encargos,"
                sql = sql & "descricao_linha_t,descricao_linha_u) values(" & 2177 & "," & RdoAux2!CodBanco & "," & 0 & "," & nNumDocISS & "," & RdoAux2!SEQPAG & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "',"
                sql = sql & Virg2Ponto(CStr(nValorDoc)) & "," & Virg2Ponto(CStr(nValorDoc)) & ",'" & Format(RdoAux2!DataPagamento, "mm/dd/yyyy") & "'," & 0 & ",'"
                sql = sql & "" & "','" & "" & "')"
                cnEicon.Execute sql, rdExecDirect
            End If
            sql = "update damiss set baixado=1 where dociss=" & nNumDocISS
            cn.Execute sql, rdExecDirect
        End If
       .MoveNext
    Loop
   .Close
End With
                   
End Sub


Private Sub ISSpagoPorDAM()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nAno As Integer, nLc As Integer, nSq As Integer, RdoAux3 As rdoResultset
Dim nPc As Integer, nCp As Integer, sCnae As String, nPos As Long, nTot As Long, nIni As Integer, nFim As Integer, sMotivo As String, nNumDoc As Long
If NomeDeLogin <> "SCHWARTZ" Then Exit Sub

nPos = 1
sql = "SELECT * From debitopago WHERE (codreduzido between 100000 and 200000) and (codlancamento = 5) AND (numdocumento > 4000000) AND (codbanco NOT IN (90, 91, 92, 93, 94, 95, 96, 97, 98, 99)) AND "
sql = sql & " (seqpag = 0) AND (codcomplemento = 0) and year(datapagamento)>2017 order by numdocumento"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If

        nNumDoc = !NumDocumento
        sql = "select * from parceladocumento where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and codlancamento=" & !CodLancamento & " and "
        sql = sql & "seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO & " and numdocumento between 2000000 and 3000000"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                sql = "select * from parceladocumento where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and codlancamento=" & !CodLancamento & " and "
                sql = sql & "seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO & " and numdocumento <> " & rdoAux!NumDocumento
                Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                Do Until RdoAux3.EOF
                    sql = "insert damiss (docdam,dociss,baixado) values(" & nNumDoc & "," & RdoAux3!NumDocumento & ",0)"
                    cn.Execute sql, rdExecDirect
                    
                    RdoAux3.MoveNext
                Loop
                RdoAux3.Close
            End If
           .Close
        End With
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With
MsgBox "fim"

End Sub

Private Sub CorrigeRefis()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPercIsencao As Integer, nPlano As Integer, nSq As Integer, RdoAux3 As rdoResultset


sql = "SELECT DISTINCT parceladocumento.plano, numdocumento.percisencao, numdocumento.emissor, numdocumento.numdocumento FROM parceladocumento INNER JOIN "
sql = sql & "numdocumento ON parceladocumento.numdocumento = numdocumento.numdocumento Where (parceladocumento.plano = 0) And (NumDocumento.percisencao > 0) And (Year(NumDocumento.datadocumento) = 2017)"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    Do Until .EOF
        nNumDoc = !NumDocumento
        nPercIsencao = !percisencao
        If nPercIsencao = 100 Then
            nPlano = 16
        ElseIf nPercIsencao = 80 Then
            nPlano = 17
        ElseIf nPercIsencao = 60 Then
            nPlano = 18
        ElseIf nPercIsencao = 50 Then
            nPlano = 19
        End If
        sql = "update parceladocumento set plano=" & nPlano & " where numdocumento=" & nNumDoc
        cn.Execute sql, rdExecDirect
        
       .MoveNext
    Loop
   .Close
End With
MsgBox "fim"



End Sub


Private Sub CancelaUnica()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPercIsencao As Integer, nPlano As Integer, nSq As Integer, RdoAux3 As rdoResultset


sql = "SELECT * FROM debitoparcela WHERE (codreduzido < 100000) AND (anoexercicio = 2018) AND (codlancamento = 1) AND (numparcela > 0) AND (statuslanc = 2) order by codreduzido"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    Do Until .EOF
        sql = "SELECT * FROM debitoparcela WHERE codreduzido = " & !CODREDUZIDO & " AND (anoexercicio = 2018) AND (codlancamento = 1) AND (numparcela = 0)  AND (statuslanc = 3)"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            sql = "update debitoparcela set statuslanc=5 WHERE codreduzido = " & !CODREDUZIDO & " AND (anoexercicio = 2018) AND (codlancamento = 1) AND (numparcela = 0)  AND (statuslanc = 3)"
            cn.Execute sql, rdExecDirect
        End If
'        Sql = "update parceladocumento set plano=" & nPlano & " where numdocumento=" & nNumDoc
'        cn.Execute Sql, rdExecDirect
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "fim"



End Sub

Private Sub Relatorio_SanMarino()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String

sql = "truncate table relatorio_sanmarino"
cn.Execute sql, rdExecDirect

sql = "SELECT codreduzido, nomecidadao, LOGRADOURO, li_num, descbairro, li_quadras, li_lotes From vwFULLIMOVEL2 WHERE (li_codbairro IN (81, 1056, 1062, 1064, 1074, 1075, 1077)) ORDER BY li_codbairro, codreduzido"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    Do Until .EOF
        sExercicio = ""
        sql = "SELECT distinct anoexercicio FROM debitoparcela WHERE codreduzido = " & !CODREDUZIDO & " AND (statuslanc = 3) and datavencimento<getdate()  order by anoexercicio"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            Do Until RdoAux2.EOF
                sExercicio = sExercicio & RdoAux2!AnoExercicio & ","
                RdoAux2.MoveNext
            Loop
            RdoAux2.Close
        End If
        If sExercicio <> "" Then
            sExercicio = Left(sExercicio, Len(sExercicio) - 1)
        End If
        sql = "INSERT relatorio_sanmarino (codreduzido,nome,endereco,numero,bairro,quadras,lotes ,exercicio) values (" & !CODREDUZIDO & ",'" & Mask(!nomecidadao) & "','"
        sql = sql & !Logradouro & "'," & !Li_Num & ",'" & !DescBairro & "','" & !Li_Quadras & "','" & !Li_Lotes & "','" & sExercicio & "')"
        cn.Execute sql, rdExecDirect
        
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "fim"

End Sub

Private Sub Corrige_Obs()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String
Dim nPos As Long, nTot As Long
On Error GoTo Erro

GoTo Debito:

sql = "truncate table obsparcela2"
'cn.Execute Sql, rdExecDirect

sql = "SELECT * from obsparcela where anoexercicio>=2018 order by codreduzido,anoexercicio,codlancamento,seqlancamento,numparcela,codcomplemento,seq"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 100 = 0 Then
           CallPb nPos, nTot
        End If
        sql = "INSERT obsparcela2 (codreduzido,anoexercicio,codlancamento,seqlancamento,numparcela,codcomplemento,seq,obs,usuario,data) values(" & !CODREDUZIDO & ","
        sql = sql & !AnoExercicio & "," & !CodLancamento & "," & !SeqLancamento & "," & !NumParcela & "," & !CODCOMPLEMENTO & "," & !Seq & ",'" & Mask(!obs) & "','"
        sql = sql & !USUARIO & "','" & Format(!Data, "mm/dd/yyyy") & "')"
'        cn.Execute Sql, rdExecDirect
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

Debito:
sql = "SELECT * from debitoobservacao  order by codreduzido"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 100 = 0 Then
           CallPb nPos, nTot
        End If
        sql = "INSERT debitoobservacao2 (codreduzido,seq,usuario,dataobs,obs) values(" & !CODREDUZIDO & ","
        sql = sql & !Seq & ",'" & !USUARIO & "','" & Format(!DATAOBS, "mm/dd/yyyy") & "','" & Mask(!obs) & "')"
        cn.Execute sql, rdExecDirect
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With



MsgBox "fim"

Exit Sub

Erro:
'MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub ContaArea()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String
Dim nPos As Long, nTot As Long
Dim bR As Boolean, bC As Boolean, bId As Boolean, bIn As Boolean
Dim nCountR As Integer, nCountC As Integer, nCountId As Integer, nCountIn As Integer, nCountM As Integer, nCountTmp As Integer
On Error GoTo Erro


sql = "SELECT codreduzido from cadimob where inativo=0 order by codreduzido"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    nCountR = 0: nCountC = 0: nCountId = 0: nCountIn = 0: nCountM = 0
    Do Until .EOF
        If nPos Mod 100 = 0 Then
           CallPb nPos, nTot
        End If
        nCodReduz = !CODREDUZIDO
        nCountTmp = 0
        bR = False: bC = False: bId = False: bIn = False
        sql = "select * from areas where codreduzido=" & nCodReduz
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                If !USOCONSTR = 1 Then
                    bR = True
                ElseIf !USOCONSTR = 2 Then
                    bId = True
                ElseIf !USOCONSTR = 3 Then
                    bC = True
                ElseIf !USOCONSTR = 4 Then
                    bIn = True
                End If
               .MoveNext
            Loop
           .Close
        End With
        If bR Then
            nCountTmp = nCountTmp + 1
        End If
        If bC Then
            nCountTmp = nCountTmp + 1
        End If
        If bId Then
            nCountTmp = nCountTmp + 1
        End If
        If bIn Then
            nCountTmp = nCountTmp + 1
        End If
        If nCountTmp > 1 Then
            nCountM = nCountM + 1
        Else
            If bR Then
                nCountR = nCountR + 1
            End If
            If bC Then
                nCountC = nCountC + 1
            End If
            If bIn Then
                nCountIn = nCountIn + 1
            End If
            If bId Then
                nCountId = nCountId + 1
            End If
        End If
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With



MsgBox "Residencial: " & nCountR & vbCrLf & "Comercial: " & nCountC & vbCrLf & "Industrial: " & nCountId & vbCrLf & "Institucional: " & nCountIn & vbCrLf & "Misto: " & nCountM

Exit Sub

Erro:
'MsgBox rdoErrors(1).Description
Resume Next

End Sub


Private Sub Codigo_Usuario()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String, x As Integer
x = 1

sql = "SELECT * from usuario order by nomelogin"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    Do Until .EOF
        sql = "update usuario set id=" & x & " where nomelogin='" & !NomeLogin & "'"
        cn.Execute sql, rdExecDirect
        
        x = x + 1
       .MoveNext
    Loop
   .Close
End With
MsgBox "fim"

End Sub

Private Sub ProcessoGTI()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String
Dim nPos As Long, nTot As Long
On Error GoTo Erro

sql = "SELECT ano,numero,responsavel from processogti where responsavel is not null and userid is null order by ano,numero"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 100 = 0 Then
           CallPb nPos, nTot
        End If
        nCodReduz = idFromLogin(UCase(Trim(!RESPONSAVEL)))
        'nCodReduz = RetornaUsuarioID(UCase(Trim(!RESPONSAVEL)))
        If nCodReduz > 0 Then
            sql = "update processogti set userid=" & RetornaUsuarioID(!RESPONSAVEL) & " where ano=" & !ano & " and numero=" & !Numero
            cn.Execute sql, rdExecDirect
        End If
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"

Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub CorrigeTramite()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String
Dim nPos As Long, nTot As Long
On Error GoTo Erro

sql = "SELECT ano,numero,seq,usuario,usuario2,userid,userid2 from tramitacao where usuario2 is not null and userid2 is null order by ano,numero"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 100 = 0 Then
           CallPb nPos, nTot
        End If
        nCodReduz = 0
        
        If Not IsNull(!Usuario2) Then
            nCodReduz = idFromLogin(UCase(Trim(!Usuario2)))
        End If
        
        sql = "update tramitacao set userid2=" & nCodReduz & " where ano=" & !ano & " and numero=" & !Numero & " and seq=" & !Seq
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"

Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub


Private Sub CorrigeUsuarioCC()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String
Dim nPos As Long, nTot As Long
On Error GoTo Erro

sql = "SELECT nome,codigocc from usuariocc order by nome,codigocc"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 100 = 0 Then
           CallPb nPos, nTot
        End If
        
        nCodReduz = idFromLogin(UCase(Trim(!Nome)))
        
        sql = "update usuariocc  set userid=" & nCodReduz & " where nome='" & !Nome & "' and codigocc=" & !codigocc
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"

Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub CorrigeDebitoParcela()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String
Dim nPos As Long, nTot As Long
On Error GoTo Erro

sql = "SELECT distinct usuario from debitoparcela where usuario is not null and userid is null order by usuario"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        
        nCodReduz = idFromLogin(UCase(Trim(!USUARIO)))
        
        sql = "update debitoparcela  set userid=" & nCodReduz & " where usuario='" & !USUARIO & "'"
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"

Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub CorrigeDebitoObservacao()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String
Dim nPos As Long, nTot As Long
On Error GoTo Erro

sql = "SELECT distinct usuario from debitoobservacao where usuario is not null and userid =0 order by usuario"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        
        nCodReduz = idFromLogin(UCase(Trim(!USUARIO)))
        
        sql = "update debitoobservacao  set userid=" & nCodReduz & " where usuario='" & !USUARIO & "'"
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"

Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub CorrigeDebitoCancel()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String
Dim nPos As Long, nTot As Long
On Error GoTo Erro

sql = "SELECT distinct usuario from debitocancel where usuario is not null and userid is null order by usuario"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        
        nCodReduz = idFromLogin(UCase(Trim(!USUARIO)))
        
        sql = "update debitocancel  set userid=" & nCodReduz & " where usuario='" & !USUARIO & "'"
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"

Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub CorrigeObsCidadao()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String
Dim nPos As Long, nTot As Long
On Error GoTo Erro

sql = "SELECT distinct usuario from obsparcela where usuario is not null and userid =0 order by usuario"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        
        nCodReduz = idFromLogin(UCase(Trim(!USUARIO)))
        
        sql = "update obsparcela  set userid=" & nCodReduz & " where usuario='" & !USUARIO & "'"
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"

Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub CorrigeHistorico()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String
Dim nPos As Long, nTot As Long
On Error GoTo Erro

sql = "SELECT distinct usuario from Historicocidadao where usuario is not null and userid is null order by usuario"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        
        nCodReduz = idFromLogin(UCase(Trim(!USUARIO)))
        
        sql = "update Historicocidadao  set userid=" & nCodReduz & " where usuario='" & !USUARIO & "'"
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"

Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub CorrigeIsencao()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String
Dim nPos As Long, nTot As Long
On Error GoTo Erro

sql = "SELECT DISTINCT codreduzido From debitoparcela Where (CODREDUZIDO < 100000) And (AnoExercicio = 2018) And (CodLancamento = 1) And (NumParcela = 0) And (statuslanc = 2) ORDER BY codreduzido"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        
        nCodReduz = !CODREDUZIDO
        
        sql = "update debitoparcela set statuslanc=1 where codreduzido=" & nCodReduz & " and anoexercicio=2018 and codlancamento=1 and  statuslanc=3"
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"

Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub



Private Function idFromLogin(sNomeLogin As String) As Integer
Dim x As Integer, nRet As Integer

nRet = 0
For x = 1 To UBound(aIdUser)
    If aIdUser(x).Nome = sNomeLogin Then
        nRet = aIdUser(x).id
        Exit For
    End If
Next
idFromLogin = nRet

End Function


Private Sub EmiteBoletoCIP()

Dim nValorTaxa As Double, x As Integer, nSituacao As Integer, dDataProc As Date, sDescImposto As String, RdoAux2 As rdoResultset, y As Integer, nPercTrib As Double
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, sDataVencto As String, nCodTrib As Integer, nValorTributo As Double
Dim NumBarra1 As String, StrBarra1 As String, NumBarra2 As String, NumBarra2a As String, NumBarra2b As String, NumBarra2c As String, NumBarra2d As String, StrBarra2 As String
Dim sCodReduz As String, sNomeResp As String, sTipoImposto As String, sEndImovel As String, nNumImovel As Integer, sComplImovel As String, sBairroImovel As String
Dim nCodLogr As Long, sEndEntrega As String, nNumEntrega As Integer, sBairroEntrega As String, sComplEntrega As String, sCepEntrega As String, sCidadeEntrega As String
Dim sUFEntrega As String, sNumInsc As String, sValorParc As String, nNumDoc As Long, sBarra As String, sDigitavel2 As String, nValorGuia As Double, nNumGuia As Long
Dim sNumDoc2 As String, sNumDoc3 As String, nFatorVencto As Long, sNumDoc As String, nSid As Long, sDigitavel As String, sNossoNumero As String, sCPF As String, sObs As String
Dim clsImovel As New clsImovel, nCodReduz As Long, sSetor As String, sRG As String, dDataPrimeiraParc As String, nValorTotalHon As Double, RdoAux3 As rdoResultset
Dim nPagina As Integer, nLivro As Integer, sDataDam As String, xImovel As clsImovel


'LIMPA TEMPORARIO
nSid = Int(Rnd(100) * 1000000)

sql = "delete from boletoguia where sid=" & nSid
cn.Execute sql, rdExecDirect

sql = "delete from boletoguiacapa where sid=" & nSid
cn.Execute sql, rdExecDirect

sLib = "CIP"


'sNumProc = lblNumProc.Caption & "/" & lblAnoProc.Caption
'dDataProc = lblDataParc.Caption
sql = "SELECT cadimob.codreduzido, parceladocumento.anoexercicio, parceladocumento.codlancamento, parceladocumento.seqlancamento, parceladocumento.numparcela, "
sql = sql & "parceladocumento.CODCOMPLEMENTO , parceladocumento.NumDocumento, debitoparcela.DataVencimento, debitotributo.ValorTributo FROM cadimob INNER JOIN "
sql = sql & "parceladocumento ON cadimob.codreduzido = parceladocumento.codreduzido INNER JOIN debitoparcela ON parceladocumento.codreduzido = debitoparcela.codreduzido AND parceladocumento.anoexercicio = debitoparcela.anoexercicio AND "
sql = sql & "parceladocumento.codlancamento = debitoparcela.codlancamento AND parceladocumento.seqlancamento = debitoparcela.seqlancamento AND parceladocumento.numparcela = debitoparcela.numparcela AND parceladocumento.codcomplemento = debitoparcela.codcomplemento INNER JOIN "
sql = sql & "debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
sql = sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND "
sql = sql & "debitoparcela.NumParcela = debitotributo.NumParcela And debitoparcela.CODCOMPLEMENTO = debitotributo.CODCOMPLEMENTO "
sql = sql & "Where (cadimob.codreduzido in (select codigo from cip_semregistro where ano=2018)) "
'Sql = Sql & "Where (cadimob.li_codbairro =1069) "
sql = sql & " And (parceladocumento.AnoExercicio = 2018) And (parceladocumento.CodLancamento =79) And (parceladocumento.SeqLancamento = 0) "
sql = sql & "AND  statuslanc=18 ORDER BY cadimob.codreduzido, parceladocumento.numparcela"

Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    x = 1
   Set xImovel = New clsImovel
    Do Until .EOF
        DoEvents
        nCodReduz = !CODREDUZIDO
        sTipoImposto = "Cont.Ilum.Pub."
        sSetor = "IMOBILIÁRIO"
        xImovel.CarregaImovel nCodReduz
        sNumInsc = xImovel.Inscricao
        sCodReduz = nCodReduz
        sNomeResp = xImovel.NomePropPrincipal
        sQuadra = xImovel.Li_Quadras
        sLote = xImovel.Li_Lotes
        xImovel.RetornaEndereco nCodReduz, Imobiliario, Localizacao
        sEndImovel = xImovel.Endereco
        nNumImovel = xImovel.Numero
        sComplImovel = xImovel.Complemento
        sBairroImovel = xImovel.Bairro
        sEndEntrega = xImovel.Ee_NomeLog
        nNumEntrega = xImovel.Ee_NumImovel
        sComplEntrega = xImovel.Ee_Complemento
        sBairroEntrega = xImovel.Ee_Bairro
        sCidadeEntrega = "JABOTICABAL"
        sUFEntrega = "SP"
        sCepEntrega = xImovel.Ee_Cep
        sql = "SELECT cidadao.codcidadao, cidadao.nomecidadao, cidadao.cpf, cidadao.cnpj, cidadao.rg "
        sql = sql & "FROM cidadao INNER JOIN proprietario ON cidadao.codcidadao = proprietario.codcidadao "
        sql = sql & "WHERE(proprietario.codreduzido = " & nCodReduz & ") AND (proprietario.tipoprop = 'P') AND (proprietario.principal = 1)"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                sCPF = SubNull(!cpf)
                If Trim(sCPF) = "" Then
                   sCPF = SubNull(!Cnpj)
                End If
             Else
                sCPF = ""
             End If
             sRG = SubNull(!rg)
            .Close
        End With
        
        nAno = !AnoExercicio
        nLanc = !CodLancamento
        nSeq = !SeqLancamento
        nParc = !NumParcela
        nCompl = !CODCOMPLEMENTO
        sDataDam = Format(!DataVencimento, "dd/mm/yyyy")
        nNumDoc = !NumDocumento
        nValorParc = !VALORTRIBUTO

        nNumGuia = nNumDoc

        sNumDoc = CStr(nNumGuia) & "-" & RetornaDVNumDoc(nNumGuia)
        sNumDoc2 = CStr(nNumGuia) & RetornaDVNumDoc(nNumGuia)
        sNumDoc3 = CStr(nNumGuia) & Modulo11(nNumGuia)


        
        sValorParc = Format(nValorParc, "#0.00")
        nValorGuia = sValorParc
        nValorDoc = nValorGuia

    sValor = nValorDoc
    dDataVencto = CDate(sDataDam)
    nNumDoc = nNumGuia
    sDadosLanc = "CONTRIBUIÇÃO DE ILUMINAÇÃO PÚBLICA 2018"
    NumBarra2 = Gera2of5Cod(CStr(sValor), CDate(dDataVencto), CLng(nNumDoc), CLng(nCodReduz))
    NumBarra2a = Left$(NumBarra2, 13)
    NumBarra2b = Mid$(NumBarra2, 14, 13)
    NumBarra2c = Mid$(NumBarra2, 27, 13)
    NumBarra2d = Right$(NumBarra2, 13)

    StrBarra2 = Gera2of5Str(Left$(NumBarra2a, 11) & Left$(NumBarra2b, 11) & Left$(NumBarra2c, 11) & Left$(NumBarra2d, 11))
    sBarra = StrBarra2

    '*******************************************

        sql = "insert boletoguia(usuario,computer,sid,seq,codreduzido,nome,cpf,endereco,numimovel,complemento,bairro,cidade,uf,fulllanc,numdoc,numparcela,totparcela,datavencto,numdoc2,"
        sql = sql & "digitavel,codbarra,valorguia,obs,numbarra2a,numbarra2b,numbarra2c,numbarra2d) values('" & NomeDeLogin & "','" & NomeDoComputador & "'," & nSid & "," & x & "," & nCodReduz & ",'" & Left(Mask(sNomeResp), 80) & "','" & sCPF & "','"
        sql = sql & Left(Mask(sEndImovel), 80) & "'," & nNumImovel & ",'" & Left(Mask(sComplImovel), 30) & "','" & Left(Mask(sBairroImovel), 25) & "','" & Mask(sCidadeEntrega) & "','" & sUFEntrega & "','" & Mask(sDescImposto) & "','"
        sql = sql & CStr(nNumGuia) & "'," & IIf(nParc = 0, 0, nParc) & "," & 3 & ",'" & Format(!DataVencimento, "mm/dd/yyyy") & "','" & sNumDoc & "','" & sDigitavel2 & "','" & Mask(sBarra) & "',"
        sql = sql & Virg2Ponto(Format(nValorGuia, "#0.00")) & ",'" & "contrib" & "','" & NumBarra2a & "','" & NumBarra2b & "','" & NumBarra2c & "','" & NumBarra2d & "')"
        cn.Execute sql, rdExecDirect
        x = x + 1
       .MoveNext
    Loop
   .Close
End With

frmReport.ShowReport2 "BOLETOGUIA_CIP", frmMdi.HWND, Me.HWND, nSid, nNumGuia
Liberado

End Sub

Private Sub SuspendeEmpresa()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String
Dim nPos As Long, nTot As Long, nSeq As Integer
On Error GoTo Erro

sql = "SELECT codigo from codtmp order by codigo"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodReduz = !Codigo
        
        sql = "SELECT MAX(SEQEVENTO) AS MAXIMO FROM MOBILIARIOEVENTO WHERE CODMOBILIARIO=" & nCodReduz
        sql = sql & " AND CODTIPOEVENTO=2"
        Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With rdoAux
            If IsNull(!maximo) Then
                nSeq = 0
            Else
                nSeq = !maximo + 1
            End If
        End With

        sql = "INSERT MOBILIARIOEVENTO (CODMOBILIARIO,CODTIPOEVENTO,SEQEVENTO,DATAEVENTO,NUMPROCEVENTO,DATAPROCEVENTO,TIPOCALCULO) VALUES("
        sql = sql & nCodReduz & "," & 2 & "," & nSeq & ",'" & Format("18/05/2018", "mm/dd/yyyy") & "','" & "23273/2017" & "','" & Format("14/12/2017", "mm/dd/yyyy") & "'," & 0 & ")"
        cn.Execute sql, rdExecDirect


        sql = "SELECT MAX(SEQ) AS MAXIMO FROM MOBILIARIOHIST WHERE CODMOBILIARIO=" & nCodReduz
        Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If IsNull(rdoAux!maximo) Then
            nSeq = 0
        Else
            nSeq = rdoAux!maximo + 1
        End If
            
        sTexto1 = "A Empresa foi suspensa através do processo nº 23273-4/2017 em 18/05/2018."
            
        sql = "INSERT MOBILIARIOHIST(CODMOBILIARIO,SEQ,DATAHIST,OBS,USERID) VALUES(" & nCodReduz & "," & nSeq & ",'" & Format(Now, "mm/dd/yyyy") & "','" & Mask(CStr(sTexto1)) & "'," & RetornaUsuarioID(NomeDeLogin) & ")"
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"

Exit Sub

Erro:
MsgBox Err.Description
Resume Next

End Sub

Private Sub Mei()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String
Dim nPos As Long, nTot As Long, nSeq As Integer, sHist As String, nIni As Integer, nFim As Integer, nSize As Integer
Dim sFileName As String, mStream As New ADODB.Stream, rst As New ADODB.Recordset, adoConn As New ADODB.Connection, sArq As String
Dim nTipo As Integer, nAno As Integer, sSeq As String, nSeqTipo As Integer, sExt As String, sNome As String, sNome_Novo As String
Dim sTmp As String, sHex As String, sSeqTipo As String, nAnoArq As Integer, nMesArq As Integer
Dim f As File, s, dDataCreated As Date, fso As New FileSystemObject, FSfolder As Folder, sPath As String


'ConectaBinary
On Error GoTo Erro
nTot = File1.ListCount

For x = 0 To File1.ListCount - 1
    If nPos Mod 10 = 0 Then
       CallPb nPos, nTot
    End If

    sArq = Left(File1.List(x), Len(File1.List(x)) - 4)
    sNome = File1.List(x)
    sExt = LCase(Right(File1.List(x), 3))
    nTipo = Val(Left(sArq, 2))
    nAno = Val(Mid(sArq, 3, 4))
    sSeq = Mid(sArq, 7, Len(File1.List(x)) - 6)
    nSeqTipo = Left(sSeq, Len(sSeq) - 6)
    nCodReduz = Val(Right(sArq, 6))
    
    sql = "select max(seq) as maximo from anexos where codigo=" & nCodReduz & " and tipo=" & nTipo
    Set rdoAux = cnBinary.OpenResultset(sql, rdOpenKeyset, rdConcurReadOnly)
    If IsNull(rdoAux!maximo) Then
        nSeq = 0
    Else
        nSeq = rdoAux!maximo + 1
    End If
                        
    sNome_Novo = Format(nCodReduz, "000000") & Format(nTipo, "00") & Format(nSeq, "0000")
                        
    sFileName = File1.Path + "\" + File1.List(x)
     
    Set f = fso.GetFile(sFileName)
    dDataCreated = f.DateLastModified
    nAnoArq = Year(dDataCreated)
    nMesArq = Month(dDataCreated)
    
    sql = "insert anexos(codigo,tipo,seq,ano,mes,oldname,newname,ext) values(" & nCodReduz & "," & nTipo & ","
    sql = sql & nSeq & "," & nAnoArq & "," & nMesArq & ",'" & Mask(sNome) & "','" & sNome_Novo & "','" & sExt & "')"
    cnBinary.Execute sql, rdExecDirect
     
    sql = "insert anexos_controle(codigo,tipo,seq,data,userid) values(" & nCodReduz & "," & nTipo & ","
    sql = sql & nSeq & ",'" & Format(dDataCreated, "mm/dd/yyyy") & "'," & RetornaUsuarioID(NomeDeLogin) & ")"
    cnBinary.Execute sql, rdExecDirect
     
    sPath = sPathAnexo & Format(nTipo, "00")
    If fso.FolderExists(sPath) = False Then
        fso.CreateFolder (sPath)
    End If
    sPath = sPathAnexo & Format(nTipo, "00") & "\" & Format(nAnoArq, "0000")
    If fso.FolderExists(sPath) = False Then
        fso.CreateFolder (sPath)
    End If
    sPath = sPathAnexo & Format(nTipo, "00") & "\" & Format(nAnoArq, "0000") & "\" & Format(nMesArq, "00")
    If fso.FolderExists(sPath) = False Then
        fso.CreateFolder (sPath)
    End If
    
    sPath = sPath & "\" & sNome_Novo
    fso.CopyFile sFileName, sPath, False

    nPos = nPos + 1
    DoEvents
Proximo:
Next
cnBinary.Close

MsgBox "Fim"

Exit Sub

Erro:
MsgBox Err.Description
Resume Next

End Sub

Private Sub GravaFoto()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String
Dim nPos As Long, nTot As Long, nSeq As Integer, sHist As String, nIni As Integer, nFim As Integer, nSize As Integer
Dim sFileName As String, mStream As New ADODB.Stream, rst As New ADODB.Recordset, adoConn As New ADODB.Connection, sArq As String
Dim nTipo As Integer, nAno As Integer, sSeq As String, nSeqTipo As Integer, sExt As String, sNome As String, sNome_Novo As String
Dim sTmp As String, sHex As String, sSeqTipo As String, nAnoArq As Integer, nMesArq As Integer
Dim f As File, s, dDataCreated As Date, fso As New FileSystemObject, FSfolder As Folder, sPath As String, nFolder As Integer
Dim nPos1 As Long, nPos2 As Long

'ConectaBinary

adoConn.CursorLocation = adUseClient
adoConn.Open cnBinary.Connect

nPos1 = 32750: nPos2 = 32800
Inicio:
nPos = 1
rst.Open "Select codigo,seq,foto from Foto_imovel where codigo between " & nPos1 & " and " & nPos2 & " and controle is null order by codigo,seq", adoConn, adOpenKeyset, adLockOptimistic
nTot = rst.RecordCount
Do Until rst.EOF
    If nPos Mod 50 = 0 Then
    txtValor.Text = nCodReduz
       CallPb nPos, nTot
    End If

    nCodReduz = rst!Codigo
    If nCodReduz <= 5000 Then
        nFolder = 1
    ElseIf nCodReduz > 5000 And nCodReduz <= 10000 Then
        nFolder = 2
    ElseIf nCodReduz > 10000 And nCodReduz <= 15000 Then
        nFolder = 3
    ElseIf nCodReduz > 15000 And nCodReduz <= 20000 Then
        nFolder = 4
    ElseIf nCodReduz > 20000 And nCodReduz <= 25000 Then
        nFolder = 5
    ElseIf nCodReduz > 25000 And nCodReduz <= 30000 Then
        nFolder = 6
    ElseIf nCodReduz > 30000 And nCodReduz <= 35000 Then
        nFolder = 7
    ElseIf nCodReduz > 35000 And nCodReduz <= 40000 Then
        nFolder = 8
    End If
    
    sPath = sPathAnexo & "09" & "\" & Format(nFolder, "00")
    If fso.FolderExists(sPath) = False Then
        fso.CreateFolder (sPath)
    End If
    
    nSeq = rst!Seq
    With mStream
        .Type = adTypeBinary
        .Open
        .Write rst("foto")
         sArq = Format(nCodReduz, "000000") & "09" & Format(nSeq, "0000")
        .SaveToFile sPath & "\" & sArq, adSaveCreateOverWrite
    End With
    
    sql = "insert fotos (codigo,seq,pasta,arquivo) values(" & nCodReduz & "," & nSeq & "," & nFolder & ",'" & sArq & "')"
    cnBinary.Execute sql, rdExecDirect
    
    sql = "update foto_imovel set controle=1 where codigo=" & nCodReduz & " and seq=" & nSeq
    cnBinary.Execute sql, rdExecDirect
    
    nPos = nPos + 1
    Set mStream = Nothing
    rst.MoveNext
Loop
rst.Close
nPos1 = nPos1 + 50

nPos2 = nPos2 + 50
GoTo Inicio

Fim:
Exit Sub:
cnBinary.Close
MsgBox "fim"

End Sub

Private Sub SuspendeMei2015()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String, RdoAux3 As rdoResultset
Dim nPos As Long, nTot As Long, nSeqLanc As Integer, sData As String, sObs As String, RunOnce As Boolean
On Error GoTo Erro

sObs = "Débito suspenso conforme processo 12446-0/2018 (Taxa de licença lançado para empresa do MEI)"

sql = "SELECT distinct codreduzido from debitoparcela where codreduzido between 100000 and 300000 and anoexercicio=2015 and codlancamento=6 and statuslanc=3"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        RunOnce = False
        nCodReduz = !CODREDUZIDO
        If IsMEI(nCodReduz) Then
            sql = "select * from debitoparcela where codreduzido=" & nCodReduz & " and anoexercicio=2015 and codlancamento=6 and statuslanc=3"
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                Do Until .EOF
                    sql = "update debitoparcela set statuslanc=19 where codreduzido=" & nCodReduz & " and anoexercicio=" & !AnoExercicio & " and codlancamento=6 and seqlancamento=" & !SeqLancamento
                    sql = sql & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO & " and statuslanc=3"
                    cn.Execute sql, rdExecDirect
                    
                    sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA where codreduzido=" & nCodReduz & " and anoexercicio=" & !AnoExercicio & " and codlancamento=6 and seqlancamento=" & !SeqLancamento
                    sql = sql & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO
                    Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux3
                        If IsNull(!maximo) Then
                            nSeqLanc = 1
                        Else
                            nSeqLanc = !maximo + 1
                        End If
                       .Close
                    End With
                    
                    sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & !CODREDUZIDO & "," & !AnoExercicio & ","
                    sql = sql & !CodLancamento & "," & !SeqLancamento & "," & !NumParcela & "," & !CODCOMPLEMENTO & "," & nSeqLanc & ",'" & sObs & "'," & RetornaUsuarioID(NomeDeLogin) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
                    cn.Execute sql, rdExecDirect
                    
                    If Not RunOnce Then
                        sql = "insert mei_suspenso (codigo) values(" & nCodReduz & ")"
                        cn.Execute sql, rdExecDirect
                        RunOnce = True
                    End If
                    
                   .MoveNext
                Loop
               .Close
            End With
        End If
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"

Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

'Private Function IsMEI(nCodigo As Long) As Boolean
Dim nRet As Boolean, sql As String, rdoAux As rdoResultset
nRet = False

sql = "select * from mei where codigo=" & nCodigo & " order by datainicio desc"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    If IsNull(!datafim) Then
        nRet = True
    End If
   .Close
End With

IsMEI = nRet

End Function

Private Sub EmpresaNaoPago()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String
Dim nPos As Long, nTot As Long
On Error GoTo Erro

Open sPathBin & "\codigos.txt" For Output As #1
sql = "SELECT DISTINCT codigomob FROM mobiliario INNER JOIN debitoparcela ON mobiliario.codigomob = debitoparcela.codreduzido "
sql = sql & "Where (debitoparcela.AnoExercicio > 2016) And (mobiliario.dataencerramento Is Null) ORDER BY mobiliario.codigomob"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        
        nCodReduz = !codigomob
        
        sql = "select * from debitoparcela where codreduzido=" & nCodReduz & " and anoexercicio>2016 and statuslanc<3"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount = 0 Then
                 Print #1, nCodReduz & ","
            End If
           .Close
        End With
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
Close #1
MsgBox "Fim"

Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub CorrigeVS()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos2 As Integer
Dim nPos As Long, nTot As Long, sCNPJ_Base As String, sData_Inicio As String, sData_Final As String, sFone As String, sDDD As String
On Error GoTo Erro

ConectaEicon

sql = "SELECT codigomob, ddd_nf, telefone_nf From mobiliario WHERE ddd_nf IS NOT NULL"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodReduz = !codigomob
        sDDD = !ddd_nf
        sFone = Trim(!telefone_nf)
        
        sql = " SELECT TOP(1) cod_cliente, num_cadastro, timestamp, inscricao, inscricao_estadual, nome_empresa, nome_fantasia, num_processo, tipo_empresa, cpf_cnpj, data_abertura, data_encerramento, tipo_logradouro, titulo_logradouro,"
        sql = sql & "logradouro, num_imovel, complemento, bairro, cep, cidade, estado, ddd, telefone,  fax, email, regime_empresa, status_empresa, controle, classificacao, area_total, area_ocupada, bair_cod_bairro, logr_cod_logradouro,"
        sql = sql & "imob_num_cadastro From tb_inter_empresas Where num_cadastro = " & nCodReduz & " ORDER BY timestamp DESC"
        Set RdoAux2 = cnEicon.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            sql = "insert tb_inter_empresas(cod_cliente,num_cadastro,timestamp,inscricao,inscricao_estadual,nome_empresa,nome_fantasia,"
            sql = sql & "num_processo,tipo_empresa,cpf_cnpj,data_abertura,data_encerramento,tipo_logradouro,titulo_logradouro,logradouro,"
            sql = sql & "num_imovel,complemento,bairro,cep,cidade,estado,ddd,telefone,fax,email,regime_empresa,status_empresa,classificacao,area_ocupada) "
            sql = sql & "values(2177," & nCodReduz & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "'," & nCodReduz & "," & IIf(Val(SubNull(!Inscricao_Estadual)) > 0, !Inscricao_Estadual, "Null") & ",'" & Mask(!nome_empresa) & "',"
            sql = sql & IIf(SubNull(!nome_fantasia) <> "", "'" & Mask(SubNull(!nome_fantasia)) & "'", "Null") & "," & IIf(SubNull(!num_processo) <> "", "'" & !num_processo & "'", "Null") & ",'" & !tipo_empresa & "'," & IIf(Val(SubNull(!cpf_cnpj)) > 0, Val(SubNull(!cpf_cnpj)), "Null") & ",'" & Format(!data_abertura, "m/dd/yyyy") & "',"
            sql = sql & IIf(Not IsNull(!data_encerramento), "'" & Format(!data_encerramento, "mm/dd/yyyy") & "'", "Null") & "," & IIf(SubNull(!tipo_logradouro) <> "", "'" & !tipo_logradouro & "'", "Null") & ","
            sql = sql & IIf(SubNull(!titulo_logradouro) <> "", "'" & !titulo_logradouro & "'", "Null") & ",'" & Mask(!Logradouro) & "'," & IIf(Val(SubNull(!num_imovel)) > 0, "'" & !num_imovel & "'", "Null") & "," & IIf(SubNull(!Complemento) <> "", "'" & Mask(SubNull(!Complemento)) & "'", "Null") & ",'"
            sql = sql & !Bairro & "'," & IIf(Val(SubNull(!Cep)) > 0, Val(SubNull(!Cep)), "Null") & ",'" & !Cidade & "','" & !estado & "','" & sDDD & "','" & sFone & "'," & IIf(SubNull(!Fax) <> "", "'" & SubNull(!Fax) & "'", "Null") & "," & IIf(SubNull(!Email) <> "", "'" & Trim(!Email) & "'", "Null") & ","
            sql = sql & IIf(SubNull(!regime_empresa) <> "", "'" & !regime_empresa & "'", "Null") & ",'" & IIf(IsDate(!data_encerramento), "E", "A") & "'," & IIf(SubNull(!CLASSIFICACAO) <> "", "'N'", "Null") & "," & RetornaNumero(!area_ocupada) & ")"
            cnEicon.Execute sql, rdExecDirect
        End With
        
        sDDD = ""
        sFone = ""
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"

cnEicon.Close

Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub CorrigeMei()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos2 As Integer, bFind As Boolean
Dim nPos As Long, nTot As Long, sCNPJ_Base As String, sData_Inicio As String, sData_Final As String, sFone As String, sDDD As String
On Error GoTo Erro

ConectaEicon

sql = "SELECT DISTINCT codigo, datainicio, datafim,cnpj_base From periodomei ORDER BY codigo"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        If IsNull(!datafim) Then
            sql = "select * from tb_inter_empr_mei where num_cadastro=" & !Codigo & " and data_inicio='" & Format(!datainicio, "mm/dd/yyyy") & "' and data_fim is null"
        Else
            sql = "select * from tb_inter_empr_mei where num_cadastro=" & !Codigo & " and data_inicio='" & Format(!datainicio, "mm/dd/yyyy") & "' and data_fim='" & Format(!datafim, "mm/dd/yyyy") & "'"
        End If
        Set RdoAux2 = cnEicon.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount = 0 Then
            If Not IsNull(!datafim) Then
                sql = "insert tb_inter_empr_mei (cod_cliente,num_cadastro,inscricao,base_cnpj,data_inicio,data_fim,[ timestamp]) values(" & "2177" & ","
                sql = sql & !Codigo & "," & !Codigo & "," & !Cnpj_Base & ",'" & Format(!datainicio, "mm/dd/yyyy") & "','" & Format(!datafim, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy hh:mm") & "')"
            Else
                sql = "insert tb_inter_empr_mei (cod_cliente,num_cadastro,inscricao,base_cnpj,data_inicio,[ timestamp]) values(" & "2177" & ","
                sql = sql & !Codigo & "," & !Codigo & "," & !Cnpj_Base & ",'" & Format(!datainicio, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy hh:mm") & "')"
            End If
            cnEicon.Execute sql, rdExecDirect
        End If
        RdoAux2.Close
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"

cnEicon.Close

Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub CorrigeIE()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos2 As Integer, bFind As Boolean
Dim nPos As Long, nTot As Long, sCNPJ_Base As String, sData_Inicio As String, sData_Final As String, sFone As String, sDDD As String
On Error GoTo Erro


sql = "SELECT * From mobiliarioie ORDER BY f1"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        sql = "select * from mobiliario where INSCESTADUAL='" & !F1 & "'"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount = 0 Then
            sql = "update mobiliarioie set f3=1 where f1=" & !F1
            cn.Execute sql, rdExecDirect
        End If
        RdoAux2.Close
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"


Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub CorrigeCPF()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos2 As Integer, bFind As Boolean
Dim nPos As Long, nTot As Long, sCNPJ_Base As String, sData_Inicio As String, sData_Final As String, sFone As String, sDDD As String
On Error GoTo Erro


sql = "SELECT * From CARTA_COBRANCA WHERE REMESSA=4 ORDER BY CODIGO"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        If Len(!cpf_cnpj) = 11 Then
            sql = "update carta_cobranca set tipodoc=1 where remessa=4 and codigo=" & !Codigo
        Else
            sql = "update carta_cobranca set tipodoc=2 where remessa=4 and codigo=" & !Codigo
        End If
        cn.Execute sql, rdExecDirect
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"


Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub LaserIPTU()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos2 As Integer, bFind As Boolean
Dim nPos As Long, nTot As Long, a2019() As tLaser, a2020() As tLaser, x As Integer, y As Integer
On Error GoTo Erro

ReDim a2019(0): ReDim a2020(0)

sql = "SELECT * From laseriptu where ano=2019 order by codreduzido"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        ReDim Preserve a2019(UBound(a2019) + 1)
        a2019(UBound(a2019)).ano = 2019
        a2019(UBound(a2019)).Codigo = !CODREDUZIDO
        a2019(UBound(a2019)).Area_Terreno = !AreaTerreno
        a2019(UBound(a2019)).Area_Predial = !areaconstrucao
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

sql = "SELECT * From laseriptu where ano=2020 and codreduzido<38755 order by codreduzido"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        ReDim Preserve a2020(UBound(a2020) + 1)
        a2020(UBound(a2020)).ano = 2020
        a2020(UBound(a2020)).Codigo = !CODREDUZIDO
        a2020(UBound(a2020)).Area_Terreno = !AreaTerreno
        a2020(UBound(a2020)).Area_Predial = !areaconstrucao
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

nPos = 1
nTot = UBound(a2020)
For x = 1 To UBound(a2020)
    If nPos Mod 50 = 0 Then
       CallPb nPos, nTot
    End If
    bFind = False
    nCodReduz = a2020(x).Codigo
    For y = 1 To UBound(a2019)
        If a2019(y).Codigo = nCodReduz Then
            If a2020(x).Area_Terreno <> a2019(y).Area_Terreno Or a2020(x).Area_Predial <> a2019(y).Area_Predial Then
                sql = "update laseriptu set alterado=1 where ano=2020 and codreduzido=" & nCodReduz
                cn.Execute sql, rdExecDirect
            End If
            Exit For
        End If
    Next
    nPos = nPos + 1
Next


MsgBox "Fim"


Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub


Private Sub SENHA()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos2 As Integer, bFind As Boolean
Dim nPos As Long, nTot As Long, sCNPJ_Base As String, sData_Inicio As String, sData_Final As String, sFone As String, sDDD As String
On Error GoTo Erro


sql = "SELECT * from usuario ORDER BY nomelogin"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        If Not IsNull(!SENHA) Then
            sql = "update usuario set senha2='" & Decrypt128(!SENHA, UP) & "' where nomelogin='" & !NomeLogin & "'"
            cn.Execute sql, rdExecDirect
        End If
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"


Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub


Private Sub Suspender()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos2 As Integer, nSeq As Integer
Dim nPos As Long, nTot As Long, sCNPJ_Base As String, sData_Inicio As String, sData_Final As String, sFone As String, sDDD As String
On Error GoTo Erro

sql = "SELECT codmobiliario FROM mobiliarioevento WHERE numprocevento = '15905/2018'"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        
        sql = "SELECT MAX(SEQ) AS MAXIMO FROM MOBILIARIOHIST WHERE CODMOBILIARIO=" & !codmobiliario
        Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If IsNull(rdoAux!maximo) Then
            nSeq = 0
        Else
            nSeq = rdoAux!maximo + 1
        End If
                    
        sql = "INSERT MOBILIARIOHIST(CODMOBILIARIO,SEQ,DATAHIST,OBS,USERID) VALUES("
        sql = sql & !codmobiliario & "," & nSeq & ",'" & Format(Now, "mm/dd/yyyy") & "','Empresa suspensa conforme processo nº 15905-1/2018',236)"
        cn.Execute sql, rdExecDirect
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"


Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub Simples_Cnpj()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos2 As Integer, bFind As Boolean
Dim nPos As Long, nTot As Long
On Error GoTo Erro

sql = "SELECT cnpj From simplestmp order by cnpj"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 50 = 0 Then
           CallPb nPos, nTot
        End If
        sql = "insert simples_cnpj_receita(cnpj) values('" & !Cnpj & "')"
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With


MsgBox "Fim"


Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub Numero_Certidao()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos2 As Integer, bFind As Boolean
Dim nPos As Long, nTot As Long, sCnpj As String
On Error GoTo Erro

sql = "SELECT DISTINCT cnpj From importacao_banco Where Cnpj Is Not Null ORDER BY cnpj"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 50 = 0 Then
           CallPb nPos, nTot
        End If
        nCodReduz = 0
        sCnpj = rdoAux!Cnpj
        sql = "select codigomob from mobiliario where cnpj='" & sCnpj & "'"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            nCodReduz = RdoAux2!codigomob
        End If
        RdoAux2.Close
        If nCodReduz = 0 Then
            sql = "select codcidadao from cidadao where cnpj='" & sCnpj & "'"
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount > 0 Then
                nCodReduz = RdoAux2!CodCidadao
            End If
            RdoAux2.Close
        End If
        If nCodReduz > 0 Then
            sql = "UPDATE IMPORTACAO_BANCO SET CODIGO_REDUZIDO=" & nCodReduz & " WHERE CNPJ='" & sCnpj & "'"
            cn.Execute sql, rdExecDirect
        End If
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With


MsgBox "Fim"


Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub Corrige_Protesto()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos2 As Integer, bFind As Boolean, RdoAux3 As rdoResultset
Dim nPos As Long, nTot As Long, nCodProtesto As Long, nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer
On Error GoTo Erro
ConectaIntegrativa

sql = "SELECT distinct iddevedor,cod_protesto FROM Protesto_remessa WHERE YEAR(dtLeitura)=2020 ORDER BY cod_protesto"
Set rdoAux = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        nCodReduz = !iddevedor
        nCodProtesto = !Cod_protesto
        sql = "SELECT * FROM Protesto_Debitos WHERE Cod_protesto=" & nCodProtesto
        Set RdoAux2 = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                nAno = !exercicio
                nLanc = !lancamento
                nSeq = !Seq
                nParc = !nroparcela
                nCompl = !complparcela
                
                sql = "select * from debitoparcela where codreduzido=" & nCodReduz & " and anoexercicio=" & nAno & " and codlancamento=" & nLanc & " and "
                sql = sql & "seqlancamento=" & nSeq & " and numparcela=" & nParc & " and statuslanc=6"
                Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                If RdoAux3.RowCount > 0 Then
                    Debug.Print nCodReduz
                End If
                RdoAux3.Close
               .MoveNext
               'nPos = nPos + 1
            Loop
           .Close
        End With
               
      
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

cnInt.Close
MsgBox "Fim"


Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub Corrige_Livro90()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos2 As Integer, bFind As Boolean, RdoAux3 As rdoResultset
Dim nPos As Long, nTot As Long, nCodProtesto As Long, nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, nLivro As Integer
On Error GoTo Erro

sql = "SELECT DISTINCT codreduzido FROM debitoparcela WHERE anoexercicio=2019 AND numerolivro=91 ORDER BY codreduzido"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    nLivro = 8828
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        nCodReduz = !CODREDUZIDO
        sql = "update debitoparcela set numcertidao=" & nLivro & " where codreduzido=" & nCodReduz & " and anoexercicio=2019 and numerolivro=91"
        cn.Execute sql, rdExecDirect
        'Debug.Print nCodReduz
        nLivro = nLivro + 1
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"


Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub Descarte_Processo()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos2 As Integer, bFind As Boolean
Dim nPos As Long, nTot As Long, sCNPJ_Base As String, sData_Inicio As String, sData_Final As String, sFone As String, sDDD As String
Dim sNumProcesso As String, nNumero As Long, nAno As Integer
On Error GoTo Erro

sql = "select ano,numero,data from codtmp2 order by ano,numero"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 100 = 0 Then
           CallPb nPos, nTot
        End If
        nAno = !ano
        nNumero = !Numero
        sql = "update processogti set datadescarte='" & Format(!Data, "mm/dd/yyyy") & "' where ano=" & nAno & " and numero=" & nNumero
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"


Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub Conta_Domicilio()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos2 As Integer, bFind As Boolean
Dim nPos As Long, nTot As Long, sCNPJ_Base As String, sData_Inicio As String, sData_Final As String, sFone As String, sDDD As String
Dim nContaImovel As Long, nContaDomicilio As Long, nNumDoc As Long, nValor As Double
On Error GoTo Erro
GoTo 2

sql = "select documento, SUM(valor) AS soma FROM resumo_pagto_banco_ficha GROUP  BY documento ORDER BY documento"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        nNumDoc = !Documento
        nValor = !soma
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        'Sql = "select sum(valorpagoreal) as soma from debitopago where numdocumento=" & nNumDoc
        sql = "select valorpago as soma from numdocumento where numdocumento=" & nNumDoc
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2!soma - nValor > 1 Then
            MsgBox nNumDoc & "   Valor doc: " & RdoAux2!soma & "   Valor analise: " & nValor
        End If
        RdoAux2.Close
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

2:
sql = "select distinct numdocumento FROM resumo_pagto_banco_ficha GROUP  BY documento ORDER BY documento"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        nNumDoc = !Documento
        nValor = !soma
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        'Sql = "select sum(valorpagoreal) as soma from debitopago where numdocumento=" & nNumDoc
        sql = "select valorpago as soma from numdocumento where numdocumento=" & nNumDoc
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2!soma - nValor > 1 Then
            MsgBox nNumDoc & "   Valor doc: " & RdoAux2!soma & "   Valor analise: " & nValor
        End If
        RdoAux2.Close
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With


Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub CorrigeCep()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos2 As Integer, bFind As Boolean, RdoAux3 As rdoResultset
Dim nPos As Long, nTot As Long, sUF As String, sCidade As String, sBairro As String, sCep As String, nCodBairro As Integer, nCodCidade As Integer
On Error GoTo Erro

sql = "SELECT cep,codlogr,LOGRADOURO,codbairro FROM CEP INNER JOIN VWlogradouro ON CEP.codlogr = codlogradouro "
sql = sql & "WHERE cep <> 14870000 AND codbairro IS NOT null and cep NOT IN (SELECT cep FROM cepDB) ORDER BY cep"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 30 = 0 Then
            CallPb nPos, nTot
            DoEvents
        End If
        '****grava cep ************
        sql = "INSERT cepdb (cep,uf,cidadecodigo,bairrocodigo,logradouro,func,userid) values('" & !Cep & "','SP',"
        sql = sql & 413 & "," & !CodBairro & ",'" & Mask(!Logradouro) & "',0,0)"
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"

Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub RelProcesso()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos2 As Integer, nSeq As Integer
Dim nPos As Long, nTot As Long, sNumProc As String, nAno As Integer, nNumero As Long, nUserId As Integer
On Error GoTo Erro

sql = "delete from processotmp"
cn.Execute sql, rdExecDirect

sql = "SELECT ano,numero,p.COMPLEMENTO,p.OBSERVACAO,p.DATAENTRADA,userid FROM processogti p "
sql = sql & "WHERE DATAENTRADA BETWEEN '02/01/2021' AND '02/28/2021' AND (p.USERID IN (96,414) OR p.CENTROCUSTO IN (113,117,179))"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nUserId = Val(SubNull(!UserId))
        sql = "select compl from processotmp where ano=" & !ano & " and numero=" & !Numero
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount = 0 Then
            sNumProc = !Numero & "-" & RetornaDVProcesso(!Numero) & "/" & !ano
            sql = "INSERT processotmp(ano,numero,anonumero,compl,obs,data,userid) VALUES(" & !ano & "," & !Numero & ",'" & sNumProc & "','"
            sql = sql & UCase(Mask(!Complemento)) & "','" & UCase(Mask(!OBSERVACAO)) & "','" & Format(!DATAENTRADA, "mm/dd/yyyy") & "'," & nUserId & ")"
            cn.Execute sql, rdExecDirect
        End If
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

sql = "SELECT p.ano,p.numero,p.COMPLEMENTO,p.OBSERVACAO,p.DATAENTRADA,t.userid FROM processogti p INNER JOIN tramitacao t ON p.ANO = t.ano AND p.NUMERO = t.numero "
sql = sql & "WHERE DATAENTRADA BETWEEN '02/01/2021' AND '02/28/2021' AND t.ccusto IN (113,117,179)"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nUserId = Val(SubNull(!UserId))
        sql = "select compl from processotmp where ano=" & !ano & " and numero=" & !Numero
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount = 0 Then
            sNumProc = !Numero & "-" & RetornaDVProcesso(!Numero) & "/" & !ano
            sql = "INSERT processotmp(ano,numero,anonumero,compl,obs,data,userid) VALUES(" & !ano & "," & !Numero & ",'" & sNumProc & "','"
            sql = sql & UCase(Mask(!Complemento)) & "','" & UCase(Mask(!OBSERVACAO)) & "','" & Format(!DATAENTRADA, "mm/dd/yyyy") & "'," & nUserId & ")"
            cn.Execute sql, rdExecDirect
        End If
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"


Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub Inadimplente()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos As Long, bFind As Boolean, nTot As Long
Dim aMat() As Inad, nAnoLc As Integer, nMesLc As Integer, nAnoPg As Integer, nMesPg As Integer, tPos As Integer, ax As String
On Error GoTo Erro

ReDim Preserve aMat(0)

'Carrega matriz com meses entre 01/2019 até 02/2021
For nAnoLc = 2019 To 2021
    For nMesLc = 1 To 12
        If nAnoLc = 2021 And nMesLc = 3 Then
            GoTo FimAno
        End If
        For nAnoPg = 2019 To 2021
            For nMesPg = 1 To 12
                If nAnoPg = 2021 And nMesPg = 3 Then
                    GoTo FimAnoPg
                End If
                ReDim Preserve aMat(UBound(aMat) + 1)
                aMat(UBound(aMat)).nQtdeAtrasado = 0
                aMat(UBound(aMat)).nAnoLc = nAnoLc
                aMat(UBound(aMat)).nMesLc = nMesLc
                aMat(UBound(aMat)).nAnoPg = nAnoPg
                aMat(UBound(aMat)).nMesPg = nMesPg
            Next
        Next
FimAnoPg:
    Next
Next
FimAno:

sql = "SELECT DISTINCT debitoparcela.CODREDUZIDO , DataVencimento, DataPagamento From dbo.debitoparcela INNER JOIN dbo.debitopago ON "
sql = sql & "debitoparcela.codreduzido = debitopago.codreduzido AND debitoparcela.anoexercicio = debitopago.anoexercicio AND "
sql = sql & "debitoparcela.codlancamento = debitopago.codlancamento AND debitoparcela.seqlancamento = debitopago.seqlancamento AND "
sql = sql & "debitoparcela.numparcela = debitopago.numparcela AND debitoparcela.codcomplemento = debitopago.codcomplemento "
sql = sql & "Where debitoparcela.datavencimento BETWEEN ('01/01/2019') AND ('02/28/2021') AND debitoparcela.codlancamento = 1 AND "
sql = sql & "debitoparcela.numparcela >0 AND debitoparcela.seqlancamento = 0 AND debitoparcela.datavencimento<debitopago.datapagamento"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 50 = 0 Then
            CallPb nPos, nTot
            DoEvents
        End If
        nAnoLc = Year(!DataVencimento)
        nMesLc = Month(!DataVencimento)
        nAnoPg = Year(!DataPagamento)
        nMesPg = Month(!DataPagamento)
        
        For tPos = 1 To UBound(aMat)
            If aMat(tPos).nAnoLc = nAnoLc And aMat(tPos).nMesLc = nMesLc And aMat(tPos).nAnoPg = nAnoPg And aMat(tPos).nMesPg = nMesPg Then
                aMat(tPos).nQtdeAtrasado = aMat(tPos).nQtdeAtrasado + 1
                Exit For
            End If
        Next
        
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

Open sPathBin & "\inad2.txt" For Output As #2

For nAnoLc = 2019 To 2021
    For nMesLc = 1 To 12
        If nAnoLc = 2021 And nMesLc = 3 Then
            Close #2
            GoTo Fim
        End If
        ax = nAnoLc & ";" & nMesLc & ";"
        For tPos = 1 To UBound(aMat)
            With aMat(tPos)
                If .nAnoLc = nAnoLc And .nMesLc = nMesLc Then
                    ax = ax & .nQtdeAtrasado & ";"
                End If
            End With
        Next
        ax = Left(ax, Len(ax) - 1)
        Print #2, ax
    Next
Next


Close #2


Fim:
MsgBox "Fim"

Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub


Private Sub InadimplenteValor()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos As Long, bFind As Boolean, nTot As Long
Dim aMat() As Inad, nAnoLc As Integer, nMesLc As Integer, nAnoPg As Integer, nMesPg As Integer, tPos As Integer, ax As String
On Error GoTo Erro

ReDim Preserve aMat(0)

'Carrega matriz com meses entre 01/2019 até 02/2021
For nAnoLc = 2019 To 2021
    For nMesLc = 1 To 12
        If nAnoLc = 2021 And nMesLc = 3 Then
            GoTo FimAno
        End If
        For nAnoPg = 2019 To 2021
            For nMesPg = 1 To 12
                If nAnoPg = 2021 And nMesPg = 3 Then
                    GoTo FimAnoPg
                End If
                ReDim Preserve aMat(UBound(aMat) + 1)
                aMat(UBound(aMat)).nQtdeAtrasado = 0
                aMat(UBound(aMat)).nAnoLc = nAnoLc
                aMat(UBound(aMat)).nMesLc = nMesLc
                aMat(UBound(aMat)).nAnoPg = nAnoPg
                aMat(UBound(aMat)).nMesPg = nMesPg
            Next
        Next
FimAnoPg:
    Next
Next
FimAno:

sql = "SELECT debitoparcela.CODREDUZIDO , DataVencimento, DataPagamento,ValorPagoReal From dbo.debitoparcela INNER JOIN dbo.debitopago ON "
sql = sql & "debitoparcela.codreduzido = debitopago.codreduzido AND debitoparcela.anoexercicio = debitopago.anoexercicio AND "
sql = sql & "debitoparcela.codlancamento = debitopago.codlancamento AND debitoparcela.seqlancamento = debitopago.seqlancamento AND "
sql = sql & "debitoparcela.numparcela = debitopago.numparcela AND debitoparcela.codcomplemento = debitopago.codcomplemento "
sql = sql & "Where debitoparcela.datavencimento BETWEEN ('01/01/2019') AND ('02/28/2021') AND debitoparcela.codlancamento = 1 AND "
sql = sql & "debitoparcela.numparcela >0 AND debitoparcela.seqlancamento = 0 AND debitoparcela.datavencimento<debitopago.datapagamento"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 50 = 0 Then
            CallPb nPos, nTot
            DoEvents
        End If
        nAnoLc = Year(!DataVencimento)
        nMesLc = Month(!DataVencimento)
        nAnoPg = Year(!DataPagamento)
        nMesPg = Month(!DataPagamento)
        
        For tPos = 1 To UBound(aMat)
            If aMat(tPos).nAnoLc = nAnoLc And aMat(tPos).nMesLc = nMesLc And aMat(tPos).nAnoPg = nAnoPg And aMat(tPos).nMesPg = nMesPg Then
                aMat(tPos).nQtdeAtrasado = aMat(tPos).nQtdeAtrasado + 1
                    aMat(tPos).nValor = aMat(tPos).nValor + !ValorPagoreal
                Exit For
            End If
        Next
        
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

Open sPathBin & "\inad2.txt" For Output As #2

For nAnoLc = 2019 To 2021
    For nMesLc = 1 To 12
        If nAnoLc = 2021 And nMesLc = 3 Then
            Close #2
            GoTo Fim
        End If
        ax = nAnoLc & ";" & nMesLc & ";"
        For tPos = 1 To UBound(aMat)
            With aMat(tPos)
                If .nAnoLc = nAnoLc And .nMesLc = nMesLc Then
                    ax = ax & .nValor & ";"
                End If
            End With
        Next
        ax = Left(ax, Len(ax) - 1)
        Print #2, ax
    Next
Next


Close #2


Fim:
MsgBox "Fim"

Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub


Private Sub Corrige_Unica_2021()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos2 As Integer, bFind As Boolean
Dim nPos As Long, nTot As Long, sCNPJ_Base As String, sData_Inicio As String, sData_Final As String, sFone As String, sDDD As String
Dim sNumProcesso As String
On Error GoTo Erro

sql = "SELECT * from debitopago WHERE  anoexercicio=2021 AND codlancamento =14 "
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
    '    Sql = "SELECT * from debitoparcela WHERE codreduzido=" & !CODREDUZIDO & " and anoexercicio=2021 AND codlancamento =6"
    '    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    '    With RdoAux2
     '       Do Until .EOF
     
                If !NumParcela = 0 Then
                    sql = "update debitoparcela set statuslanc=1 where codreduzido=" & !CODREDUZIDO & " and anoexercicio=2021 and (codlancamento=6 or codlancamento=14)"
                    cn.Execute sql, rdExecDirect
                Else
                    sql = "update debitoparcela set statuslanc=2 where codreduzido=" & !CODREDUZIDO & " and anoexercicio=2021 and (codlancamento=6 or codlancamento=14) and numparcela=" & !NumParcela
                    cn.Execute sql, rdExecDirect
                End If
                'RdoAux2.MoveNext
      '      Loop
     '   End With
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"


Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub Corrige_Endereco_Empresa()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos2 As Integer, bFind As Boolean
Dim nPos As Long, nTot As Long, sCNPJ_Base As String, sData_Inicio As String, sData_Final As String, sFone As String, sDDD As String
Dim sNumProcesso As String, nCodImovel As Long
On Error GoTo Erro

sql = "  SELECT CODIGOMob,codlogradouro,numero FROM mobiliario WHERE dataencerramento IS NULL AND areatl>0 "
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodReduz = !codigomob
        
        sql = "SELECT * FROM vwfullimovel WHERE codlogr=" & !CodLogradouro & " AND li_num=" & !Numero
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            nCodImovel = RdoAux2!CODREDUZIDO
        Else
            nCodImovel = 0
        End If
        
        sql = "update mobiliario set imovel=" & nCodImovel & " where codigomob=" & nCodReduz
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"


Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub Corrige_Cnae_Bar()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos As Long, nTot As Long

On Error GoTo Erro

sql = "select DISTINCT codreduzido FROM debitoparcela WHERE codreduzido BETWEEN 100000 AND 200000 and  anoexercicio=2022 AND codlancamento=6 AND statuslanc=3"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodReduz = !CODREDUZIDO
        sql = "select * FROM debitotributo WHERE codreduzido =" & nCodReduz & " and  anoexercicio=2022 AND codlancamento=6"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount = 0 Then
            sql = "update debitoparcela set statuslanc=5 WHERE codreduzido=" & nCodReduz & " and  anoexercicio=2022 AND codlancamento=6"
            cn.Execute sql, rdExecDirect
        End If
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"


Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub


Private Sub Corrige_Tabela_Tributo()
Dim nCodTributo As Long, sDescTributo As String, sAbrevTributo As String, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos As Long, nTot As Long, DA As Boolean
Dim Ficha As Long, nat_ficha As String, FichaJrMulta As Long, nat_jrmulta As String, FichaDivida As Long, nat_divida As String, FichaDaJrMul As Long, nat_dajrmul As String
Dim FichaDaEnca As Long, nat_daenca As String, FichaAjuiza As Long, nat_ajuiza As String, FichaAjJrMul As Long, nat_ajjrmul As String, FichaAjEnca As Long, nat_ajenca As String, Juros As Boolean, multa As Boolean, livro As Byte
Dim ficha_new As Long, nat_ficha_new As String, fichajrmulta_new As Long, nat_jrmulta_new As String, fichadivida_new As Long, nat_divida_new As String, fichadajrmul_new As Long, nat_dajrmul_new As String
Dim fichadaenca_new As Long, nat_daenca_new As String, fichaajuiza_new As Long, nat_ajuiza_new As String, fichaajjrmul_new As Long, nat_ajjrmul_new As String, fichaajenca_new As Long, nat_ajenca_new As String


On Error GoTo Erro

sql = "delete from tributonew"
cn.Execute sql, rdExecDirect

sql = "SELECT * from tributo order by codtributo"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        ficha_new = 0: nat_ficha_new = "": fichajrmulta_new = 0: nat_jrmulta_new = "": fichadivida_new = 0: nat_divida_new = "": fichadajrmul_new = 0: nat_dajrmul_new = ""
        fichadaenca_new = 0: nat_daenca_new = "": fichaajuiza_new = 0: nat_ajuiza_new = "": fichaajjrmul_new = 0: nat_ajjrmul_new = "": fichaajenca_new = 0: nat_ajenca_new = ""
        
        nCodTributo = !CodTributo
        sDescTributo = Mask(!desctributo)
        sAbrevTributo = Mask(!abrevTributo)
        Juros = !Juros
        multa = !multa
        DA = !DA
        livro = Val(SubNull(!livro))
        
        Ficha = !Ficha
        nat_ficha = SubNull(!nat_ficha)
        ficha_jr_multa = !FichaJrMulta
        nat_jrmulta = SubNull(!nat_jrmulta)
        FichaDivida = !FichaDivida
        nat_divida = SubNull(!nat_divida)
        FichaDaJrMul = !FichaDaJrMul
        nat_dajrmul = SubNull(!nat_dajrmul)
        FichaDaEnca = !FichaDaEnca
        nat_daenca = SubNull(!nat_daenca)
        FichaAjuiza = !FichaAjuiza
        nat_ajuiza = SubNull(!nat_ajuiza)
        FichaAjJrMul = !FichaAjJrMul
        nat_ajjrmul = SubNull(!nat_ajjrmul)
        FichaAjEnca = !FichaAjEnca
        nat_ajenca = SubNull(!nat_ajenca)
        Juros = !Juros
        multa = !multa
        livro = Val(SubNull(!livro))
        
        sql = "SELECT ficha,natureza FROM fichacontabil  WHERE ficha_old=" & Ficha
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            ficha_new = RdoAux2!Ficha
            nat_ficha_new = RdoAux2!Natureza
        End If
        RdoAux2.Close
                
        sql = "SELECT ficha,natureza FROM fichacontabil  WHERE ficha_old=" & ficha_jr_multa
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            fichajrmulta_new = RdoAux2!Ficha
            nat_jrmulta_new = RdoAux2!Natureza
        End If
        RdoAux2.Close
                
        sql = "SELECT ficha,natureza FROM fichacontabil  WHERE ficha_old=" & FichaDivida
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            fichadivida_new = RdoAux2!Ficha
            nat_divida_new = RdoAux2!Natureza
        End If
        RdoAux2.Close
                
        sql = "SELECT ficha,natureza FROM fichacontabil  WHERE ficha_old=" & FichaDaJrMul
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            fichadajrmul_new = RdoAux2!Ficha
            nat_dajrmul_new = RdoAux2!Natureza
        End If
        RdoAux2.Close
                
        sql = "SELECT ficha,natureza FROM fichacontabil  WHERE ficha_old=" & FichaDaEnca
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            fichadaenca_new = RdoAux2!Ficha
            nat_daenca_new = RdoAux2!Natureza
        End If
        RdoAux2.Close
                
        sql = "SELECT ficha,natureza FROM fichacontabil  WHERE ficha_old=" & FichaAjuiza
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            fichaajuiza_new = RdoAux2!Ficha
            nat_ajuiza_new = RdoAux2!Natureza
        End If
        RdoAux2.Close
                
        sql = "SELECT ficha,natureza FROM fichacontabil  WHERE ficha_old=" & FichaAjJrMul
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            fichaajjrmul_new = RdoAux2!Ficha
            nat_ajjrmul_new = RdoAux2!Natureza
        End If
        RdoAux2.Close
                
        sql = "SELECT ficha,natureza FROM fichacontabil  WHERE ficha_old=" & FichaAjEnca
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            fichaajenca_new = RdoAux2!Ficha
            nat_ajenca_new = RdoAux2!Natureza
        End If
        RdoAux2.Close
                        
        sql = "INSERT TRIBUTONEW(CODTRIBUTO,DESCTRIBUTO,ABREVTRIBUTO,DA,JUROS,Multa,livro, ficha,nat_ficha,fichajrmulta,nat_jrmulta,fichadivida,nat_divida,fichadajrmul,nat_dajrmul,fichadaenca,nat_daenca,"
        sql = sql & "fichaajuiza,nat_ajuiza,fichaajjrmul,nat_ajjrmul,fichaajenca,nat_ajenca) VALUES("
        sql = sql & nCodTributo & ",'" & sDescTributo & "','" & sAbrevTributo & "'," & IIf(!DA, 1, 0) & "," & IIf(!Juros, 1, 0) & "," & IIf(!multa, 1, 0) & "," & Val(SubNull(!livro)) & "," & ficha_new & ",'"
        sql = sql & nat_ficha_new & "'," & fichajrmulta_new & ",'" & nat_jrmulta_new & "'," & fichadivida_new & ",'" & nat_divida_new & "'," & fichadajrmul_new & ",'" & nat_dajrmul_new & "',"
        sql = sql & fichadaenca_new & ",'" & nat_daenca_new & "'," & fichaajuiza_new & ",'" & nat_ajuiza_new & "'," & fichaajjrmul_new & ",'" & nat_ajjrmul_new & "'," & fichaajenca_new & ",'" & nat_ajenca_new & "')"
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"


Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next


End Sub

Private Sub Inscricao_Estadual()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos As Long, nTot As Long

On Error GoTo Erro

sql = "select inscricao from codtmp5"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        
        sql = "select codigomob FROM mobiliario WHERE inscestadual ='" & !Inscricao & "'"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            nCodReduz = RdoAux2!codigomob
            sql = "update CODTMP5 set CODIGO=" & nCodReduz & " WHERE inscricao='" & !Inscricao & "'"
            cn.Execute sql, rdExecDirect
        End If
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"


Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub Lista_Empresa_Devedora()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos As Long, nTot As Long, Sql2 As String
Dim sCodigos As String, aSuspensoCod() As Long, lResult As Long, aTemPago() As Long

On Error GoTo Erro

ReDim aTemPago(0): ReDim aSuspensoCod(0)
sql = "SELECT codmobiliario From vwMOBILIARIOSUSPENSO Where (codtipoevento = 2) order by codmobiliario"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurReadOnly)
With rdoAux
    Do Until .EOF
        ReDim Preserve aSuspensoCod(UBound(aSuspensoCod) + 1)
        aSuspensoCod(UBound(aSuspensoCod)) = !codmobiliario
       .MoveNext
    Loop
   .Close
End With

sql = "SELECT distinct codreduzido From debitopago Where (codreduzido between 100000 and 200000 and anoexercicio>2019) order by codreduzido"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurReadOnly)
With rdoAux
    Do Until .EOF
        ReDim Preserve aTemPago(UBound(aTemPago) + 1)
        aTemPago(UBound(aTemPago)) = !CODREDUZIDO
       .MoveNext
    Loop
   .Close
End With




sql = "truncate table codtmp"
cn.Execute sql, rdExecDirect
sCodigos = ""
sql = "select codigomob from mobiliario where dataencerramento is null"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    sql = "INSERT INTO codtmp(codigo) "
    Do Until .EOF
        nCodReduz = !codigomob
        lResult = BinarySearchLong(aSuspensoCod(), !codigomob)
        If lResult > -1 Then
            GoTo Proximo
        End If

        lResult = BinarySearchLong(aTemPago(), !codigomob)
        If lResult > -1 Then
            GoTo Proximo
        End If

        If nPos Mod 100 = 0 Then
            CallPb nPos, nTot
            sCodigos = sCodigos & " SELECT " & nCodReduz & " UNION ALL "
            sCodigos = Left(sCodigos, Len(sCodigos) - 10)
            Sql2 = sql & sCodigos
            cn.Execute Sql2, rdExecDirect
            sCodigos = ""
        Else
            sCodigos = sCodigos & " SELECT " & nCodReduz & " UNION ALL "
        End If

Proximo:
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
    sCodigos = Left(sCodigos, Len(sCodigos) - 10)
    Sql2 = sql & sCodigos
    cn.Execute Sql2, rdExecDirect
   .Close
End With
MsgBox "Fim"


Grava:


Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub Digito_Carta()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, sSoma As String, nDV As Integer
Dim sCep As String, nMax As Integer, nSoma As Integer, x As Integer, nDigito As Integer, sCepNovo As String, nCodigo As Long

On Error GoTo Erro

sql = "select codigo,cep_entrega FROM carta_cobranca WHERE remessa=8 order by codigo"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodigo = !Codigo
        sCep = RetornaNumero(!cep_entrega)
        nSoma = 0
        For x = 1 To 8
            nSoma = nSoma + Val(Mid(sCep, x, 1))
        Next
        sSoma = CStr(nSoma)
        nDV = Val(Right(sSoma, 1))
        If nDV = 0 Then
            nDigito = 0
        Else
            nDigito = 10 - nDV
        End If

        sCepNovo = "/" & sCep & nDigito & "\"
        sql = " update carta_cobranca set cepnet='" & sCepNovo & "' where remessa=8 and codigo=" & nCodigo
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"


Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub Imovel_Caixa()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, sSoma As String, nDV As Integer
Dim sCep As String, nMax As Integer, nSoma As Integer, x As Integer, nDigito As Integer, sCepNovo As String, nCodigo As Long
Dim RdoAux2 As rdoResultset, RdoAux3 As rdoResultset

'Sql = "delete from codtmp"
'cn.Execute Sql, rdExecDirect

sql = "SELECT vwSomaArea.CODREDUZIDO,cadimob.resideimovel From dbo.vwSomaArea LEFT OUTER JOIN dbo.cadimob ON vwSomaArea.codreduzido = cadimob.codreduzido "
sql = sql & "WHERE vwSomaArea.codreduzido IN (SELECT Proprietario.CODREDUZIDO From dbo.Proprietario Where Proprietario.CodCidadao = 500239 AND proprietario.tipoprop = 'C') "
sql = sql & "AND vwSomaArea.soma < 65 AND resideimovel=1"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodigo = !CODREDUZIDO
        
        sql = "SELECT codreduzido FROM proprietario WHERE codcidadao=500239 AND tipoprop='C'"
        Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        
        
        sql = "insert calculo_removido(ano,codigo,nome) values(" & Year(Now) & "," & nCodigo & ",'CAIXA')"
        'Sql = "insert codtmp(codigo) values(" & nCodigo & ")"
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"



End Sub

Private Sub Imovel_Wegg()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, sSoma As String, nDV As Integer
Dim sCep As String, nMax As Integer, nSoma As Integer, x As Integer, nDigito As Integer, sCepNovo As String, nCodigo As Long
Dim RdoAux2 As rdoResultset, RdoAux3 As rdoResultset

sql = "delete from codtmp"
cn.Execute sql, rdExecDirect

sql = "SELECT vwSomaArea.CODREDUZIDO,cadimob.resideimovel From dbo.vwSomaArea LEFT OUTER JOIN dbo.cadimob ON vwSomaArea.codreduzido = cadimob.codreduzido "
sql = sql & "WHERE vwSomaArea.codreduzido IN (SELECT Proprietario.CODREDUZIDO From dbo.Proprietario Where Proprietario.CodCidadao = 578651 AND proprietario.tipoprop = 'C') "
sql = sql & "AND vwSomaArea.soma < 65 AND resideimovel=1"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodigo = !CODREDUZIDO
        
        sql = "SELECT codreduzido FROM proprietario WHERE codcidadao=500239 AND tipoprop='C'"
        Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        
        
        sql = "insert calculo_removido(ano,codigo,nome) values(" & Year(Now) & "," & nCodigo & ",'WEGG')"
        'Sql = "insert codtmp(codigo) values(" & nCodigo & ")"
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"



End Sub


Private Sub Usuario_web_autoriza()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, nCodigo As Long
Dim RdoAux2 As rdoResultset

sql = "delete from usuario_web_analise_doc"
cn.Execute sql, rdExecDirect

sql = "SELECT * from usuario_web_analise"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodigo = !id
        
        sql = "SELECT id FROM usuario_web_analise_doc WHERE id=" & nCodigo
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount = 0 Then
            If Not IsNull(!data_autorizado) Then
                sql = "insert usuario_web_analise_doc(id,data_envio,autorizado,data_autorizado,autorizado_por) values(" & nCodigo & ",'"
                sql = sql & Format(!data_envio, "mm/dd/yyyy") & "'," & IIf(!autorizado, 1, 0) & ",'" & Format(!data_autorizado, "mm/dd/yyyy") & "','" & !autorizado_por & "')"
            Else
                sql = "insert usuario_web_analise_doc(id,data_envio,autorizado,autorizado_por) values(" & nCodigo & ",'"
                sql = sql & Format(!data_envio, "mm/dd/yyyy") & "'," & IIf(!autorizado, 1, 0) & ",'" & !autorizado_por & "')"
            End If
            cn.Execute sql, rdExecDirect
            
        End If
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"


End Sub

Private Sub Apaga_Taxa_Protocolo()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, nCodigo As Long
Dim RdoAux2 As rdoResultset, aProt() As tProtocolo, bFind As Boolean, x As Integer, nCodTrib As Integer
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, t As Integer

ReDim aProt(0)

sql = "SELECT dp.CODREDUZIDO , dp.AnoExercicio, dp.CodLancamento, dp.SeqLancamento, dp.NumParcela, dp.CODCOMPLEMENTO ,debitotributo.codtributo,dp.statuslanc "
sql = sql & "FROM dbo.debitoparcela dp INNER JOIN dbo.debitotributo ON dp.codreduzido = debitotributo.codreduzido AND dp.anoexercicio = debitotributo.anoexercicio AND "
sql = sql & "dp.codlancamento = debitotributo.codlancamento AND dp.seqlancamento = debitotributo.seqlancamento AND dp.numparcela = debitotributo.numparcela AND "
sql = sql & "dp.codcomplemento = debitotributo.codcomplemento Where dp.CodLancamento = 11 AND dp.datavencimento < '01/01/2023' AND dp.statuslanc = 3 ORDER BY "
sql = sql & "dp.codreduzido,dp.anoexercicio,dp.codlancamento,dp.seqlancamento,dp.numparcela,codtributo"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodReduz = !CODREDUZIDO
        nAno = !AnoExercicio
        nLanc = !CodLancamento
        nSeq = !SeqLancamento
        nParc = !NumParcela
        nCompl = !CODCOMPLEMENTO
        nCodTrib = !CodTributo
        
        bFind = False
        t = UBound(aProt)
        For x = 1 To t
            If aProt(x).nCodReduz = nCodReduz And aProt(x).nAno = nAno And aProt(x).nLanc = nLanc And aProt(x).nSeq = nSeq And aProt(x).nParc = nParc And aProt(x).nCompl = nCompl And aProt(x).nCodTributo = nCodTrib Then
                bFind = True
                Exit For
            End If
        Next
        If (Not bFind) Then
            ReDim Preserve aProt(t + 1)
            t = UBound(aProt)
            aProt(t).nCodReduz = nCodReduz
            aProt(t).nAno = nAno
            aProt(t).nLanc = nLanc
            aProt(t).nSeq = nSeq
            aProt(t).nParc = nParc
            aProt(t).nCompl = nCompl
            aProt(t).nCodTributo = nCodTrib
        End If
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

t = UBound(aProt)
For z = 1 To t
    
Next

MsgBox "Fim"


End Sub

Private Sub Rel_Ana()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, nCodigo As Long
Dim RdoAux2 As rdoResultset, nCodCidadao As Long, RdoAux3 As rdoResultset

sql = "delete from codtmp"
cn.Execute sql, rdExecDirect

sql = "SELECT vwSomaArea.CODREDUZIDO,vwSomaArea.soma,proprietario.codcidadao From dbo.vwSomaArea INNER JOIN dbo.proprietario ON "
sql = sql & "vwSomaArea.codreduzido = proprietario.codreduzido INNER JOIN dbo.cadimob ON vwSomaArea.codreduzido = cadimob.codreduzido "
sql = sql & "Where vwSomaArea.soma < 65 AND proprietario.tipoprop = 'P' AND proprietario.principal = 1 AND cadimob.imune <> 1 AND cadimob.inativo <> 1 ORDER BY vwSomaArea.codreduzido"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodReduz = !CODREDUZIDO
        nCodCidadao = !CodCidadao
        
        sql = "SELECT * FROM vwPROPRIETARIODUPLICADO WHERE CODPROPRIETARIO=" & nCodCidadao
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            
            sql = "select codreduzido from debitoparcela where codreduzido=" & nCodReduz & " and anoexercicio=2023 and codlancamento=1"
            Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If RdoAux3.RowCount = 0 Then
                sql = "insert codtmp(codigo) values(" & nCodReduz & ")"
                cn.Execute sql, rdExecDirect
            End If
            RdoAux3.Close
            
        End If
        RdoAux2.Close
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"


End Sub


Private Sub ISS_Errado()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, nAliquota_Errada As Double
Dim RdoAux2 As rdoResultset, nCodAtividade As Integer, RdoAux3 As rdoResultset, nAliquota_Correta As Double
GoTo Parte2
Exit Sub
sql = "delete from codtmp3"
cn.Execute sql, rdExecDirect

sql = "SELECT * FROM mobiliarioatividadeiss WHERE codmobiliario>=100000 and codtributo=11 AND  valoriss>14 ORDER BY codmobiliario"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodReduz = !codmobiliario
        nCodAtividade = !codatividade
        nAliquota_Errada = !valoriss
        
        
        sql = "SELECT aliquota FROM tabelaiss WHERE codigoativ=" & nCodAtividade
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        nAliquota_Correta = RdoAux2!Aliquota
        RdoAux2.Close
        
        sql = "SELECT * FROM debitoPARCELA WHERE codreduzido=" & nCodReduz & " AND anoexercicio=2023 AND CODLANCAMENTO IN (6,14) AND NUMPARCELA=1"
        Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux3.RowCount > 0 Then
            sql = "INSERT CODTMP3(CODIGO,ATIVIDADE,ALIQUOTA_ERRADA,ALIQUOTA_CORRETA) values(" & nCodReduz & "," & nCodAtividade & "," & Virg2Ponto(CStr(nAliquota_Errada)) & "," & Virg2Ponto(CStr(nAliquota_Correta)) & ")"
            cn.Execute sql, rdExecDirect
        End If
        RdoAux3.Close
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"

Parte2:
sql = "select * from codtmp3"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    Do Until .EOF
        nCodReduz = !Codigo
        nAliquota_Correta = !aliquota_correta
        sql = "update mobiliarioatividadeiss set valoriss=" & Virg2Ponto(Format(nAliquota_Correta, "#0.00")) & " where codmobiliario=" & nCodReduz & " and codtributo=11"
        cn.Execute sql, rdExecDirect
        
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"

End Sub

Private Sub Corrige_IPTU()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, nAliquota_Errada As Double
Dim RdoAux2 As rdoResultset, nCodAtividade As Integer, RdoAux3 As rdoResultset, nAliquota_Correta As Double

sql = "SELECT * FROM debitoparcela WHERE codreduzido<50000 and anoexercicio=2023 and codlancamento=1 and numparcela=0 and statuslanc=1 order by codreduzido"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodReduz = rdoAux!CODREDUZIDO
'        If nCodReduz = 286 Then
 '           MsgBox "teste"
  '      End If
        sql = "select * from debitopago where codreduzido=" & nCodReduz & " and anoexercicio=2023 and codlancamento=1 and seqlancamento=" & !SeqLancamento & " and numparcela=0 and codcomplemento=" & !CODCOMPLEMENTO
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount = 0 Then
            sql = "update debitoparcela set statuslanc=5 where codreduzido=" & nCodReduz & " and anoexercicio=2023 and codlancamento=1 and seqlancamento=" & !SeqLancamento & " and numparcela=0 and codcomplemento=" & !CODCOMPLEMENTO
            cn.Execute sql, rdExecDirect
        End If
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
End Sub

Private Sub Corrige_TaxaLic()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, nAliquota_Errada As Double
Dim RdoAux2 As rdoResultset, nCodAtividade As Integer, RdoAux3 As rdoResultset, nAliquota_Correta As Double

sql = "SELECT distinct codreduzido FROM debitoparcela WHERE codreduzido between 100000 and 200000 and anoexercicio=2025 and codlancamento=6 and statuslanc=3 order by codreduzido"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodReduz = rdoAux!CODREDUZIDO
        sql = "SELECT * FROM debitotributo WHERE codreduzido =" & nCodReduz & " and anoexercicio=2025 and codlancamento=6"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount = 0 Then
            sql = "update debitoparcela set statuslanc=5 WHERE codreduzido =" & nCodReduz & " and anoexercicio=2023 and codlancamento=6 and statuslanc=3"
            cn.Execute sql, rdExecDirect
        End If
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "fim"
End Sub

Private Sub Corrige_Importacao_Eicon()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, nAliquota_Errada As Double
Dim RdoAux2 As rdoResultset, nCodAtividade As Integer, RdoAux3 As rdoResultset, nAliquota_Correta As Double, nNumDocumento As Long

ConectaEicon
sql = "SELECT importacao_banco.*,debitopago.codreduzido From dbo.importacao_banco LEFT OUTER JOIN dbo.debitopago ON importacao_banco.Numero_Documento = debitopago.numdocumento "
sql = sql & "WHERE importacao_banco.Numero_Documento between 2200000 and 2300000 AND codreduzido IS null"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nNumDocumento = !numero_documento

        sql = "SELECT * FROM PARCELADOCUMENTO WHERE NUMDOCUMENTO=" & nNumDocumento
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)


        sql = "SELECT MAX(SEQPAG) AS MAXIMO FROM DEBITOPAGO WHERE CODREDUZIDO=" & RdoAux2!CODREDUZIDO & " AND "
        sql = sql & "ANOEXERCICIO=" & RdoAux2!AnoExercicio & " AND CODLANCAMENTO=5 AND SEQLANCAMENTO=" & RdoAux2!SeqLancamento
        sql = sql & " AND NUMPARCELA=" & RdoAux2!NumParcela & " AND CODCOMPLEMENTO=" & RdoAux2!CODCOMPLEMENTO
        Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With rdoAux
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


        sql = "INSERT DEBITOPAGO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,"
        sql = sql & "SEQPAG,DATAPAGAMENTO,DATARECEBIMENTO,VALORPAGO,CODBANCO,CODAGENCIA,NUMDOCUMENTO,VALORPAGOREAL,VALORTARIFA,ARQUIVOBANCO) VALUES(" & RdoAux2!CODREDUZIDO & ","
        sql = sql & RdoAux2!AnoExercicio & "," & RdoAux2!CodLancamento & "," & RdoAux2!SeqLancamento & "," & RdoAux2!NumParcela & "," & RdoAux2!CODCOMPLEMENTO & "," & nSeqPag & ",'"
        sql = sql & Format(!Data_Pagamento, "mm/dd/yyyy") & "','" & Format(!data_credito, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(!valor_pago)) & ","
        sql = sql & !Codigo_Banco & ",'" & SubNull(!Agencia) & "'," & nNumDocumento & "," & Virg2Ponto(CStr(!valor_pago)) & ","
        sql = sql & 0 & ",'" & Left(!nome_arquivo, 50) & "')"
        cn.Execute sql, rdExecDirect

        sql = "UPDATE DEBITOPARCELA SET STATUSLANC=2 WHERE CODREDUZIDO=" & RdoAux2!CODREDUZIDO & " AND ANOEXERCICIO=" & RdoAux2!AnoExercicio & " AND CODLANCAMENTO=5 AND SEQLANCAMENTO=" & RdoAux2!SeqLancamento
        sql = sql & " AND NUMPARCELA=" & RdoAux2!NumParcela & " AND CODCOMPLEMENTO=" & RdoAux2!CODCOMPLEMENTO
        cn.Execute sql, rdExecDirect

        'baixa ok
        '***** GRAVA BAIXA NA GISS ***************
        sql = "SELECT * FROM tb_inter_baixa_detalhe WHERE num_documento=" & nNumDocumento
        Set RdoAux3 = cnEicon.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux3.RowCount = 0 Then
        
            sql = "insert tb_inter_baixa(cod_cliente,cod_banco,num_sequencia,timestamp,data_geracao,nome_arquivo,data_movimento) values("
            sql = sql & 2177 & "," & !Codigo_Banco & "," & 0 & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "','" & Format(Now, "mm/dd/yyyy") & "','"
            sql = sql & !nome_arquivo & "','" & Format(!data_credito, "mm/dd/yyyy") & "')"
            cnEicon.Execute sql, rdExecDirect
            
            sql = "insert tb_inter_baixa_detalhe(cod_cliente,cod_banco,num_sequencia,num_documento,linha,timestamp,valor_titulo,valor_pago,data_pagamento,valor_encargos,"
            sql = sql & "descricao_linha_t,descricao_linha_u) values(" & 2177 & "," & !Codigo_Banco & "," & 0 & "," & nNumDocumento & "," & 10 & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "',"
            sql = sql & Virg2Ponto(CStr(!valor_pago)) & "," & Virg2Ponto(CStr(!valor_pago)) & ",'" & Format(!Data_Pagamento, "mm/dd/yyyy") & "'," & 0 & ",'"
            sql = sql & "" & "','" & "" & "')"
            cnEicon.Execute sql, rdExecDirect
        End If
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
Pb.value = 100
lblPB.Caption = 100
Me.Refresh
cnEicon.Close
MsgBox "fim"
End Sub


Private Sub Nova_RazaoSocial()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, nAliquota_Errada As Double, sObs As String
Dim RdoAux2 As rdoResultset, nCodAtividade As Integer, RdoAux3 As rdoResultset, sAntigo As String, sNovo As String
Dim nSeq As Integer

sql = "SELECT * from tributacao..codtmp4 order by codigo"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        sNovo = !Novo
        sAntigo = !Antigo
        nCodReduz = Val(!Inscricao)
        sObs = "Alteração da Razão Social conforme arquivo enviado pela Jucesp - Razão Social Antiga: " & sAntigo
                
        If sNovo <> sAntigo Then
            sql = "update mobiliario set razaosocial='" & Mask(sNovo) & "' WHERE codigomob=" & nCodReduz
            cn.Execute sql, rdExecDirect
            
            'grava histórico
            sql = "SELECT MAX(SEQ) AS MAXIMO FROM MOBILIARIOHIST WHERE CODMOBILIARIO=" & nCodReduz
            Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If IsNull(rdoAux!maximo) Then
                nSeq = 0
            Else
                nSeq = rdoAux!maximo + 1
            End If
           
            sql = "INSERT MOBILIARIOHIST(CODMOBILIARIO,SEQ,DATAHIST,OBS,USERID) VALUES("
            sql = sql & nCodReduz & "," & nSeq & ",'" & Format(Now, "mm/dd/yyyy") & "','" & Mask(sObs) & "',236)"
            cn.Execute sql, rdExecDirect

            'Integração_Eicon
            sql = "select codigo from eicon_empresa where codigo=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount = 0 Then
                sql = "insert eicon_empresa(codigo) values(" & nCodReduz & ")"
                cn.Execute sql, rdExecDirect
            Else
                MsgBox nCodReduz
            End If
            RdoAux2.Close
            
            
        End If
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
End Sub


Private Sub Corrige_ParcelaPix()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, nAliquota_Errada As Double, sObs As String
Dim RdoAux2 As rdoResultset, nCodAtividade As Integer, RdoAux3 As rdoResultset, sAntigo As String, sNovo As String, nNumDoc As Long, sGuid As String
Dim nSeq As Integer, nAno As Integer, nLanc As Integer, nParc As Integer, nCompl As Integer

sql = "SELECT DISTINCT dam_header.numero_documento,guid,codigo  FROM dam_header INNER JOIN importacao_banco ON dam_header.numero_documento = importacao_banco.Numero_Documento "
sql = sql & "WHERE dam_header.numero_documento = importacao_banco.Numero_Documento"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodReduz = Val(!Codigo)
        nNumDoc = !numero_documento
        sGuid = !guid
        sql = "select * from parceladocumento where numdocumento=" & nNumDoc
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            GoTo Proximo
        End If
        RdoAux2.Close
        
        sql = "select * from dam_data where guid='" & sGuid & "'"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                nAno = !exercicio
                nLanc = !lancamento
                nSeq = !Sequencia
                nParc = !Parcela
                nCompl = !Complemento
                sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO,PLANO) VALUES(" & nCodReduz & ","
                sql = sql & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & "," & nNumDoc & "," & 60 & ")"
                cn.Execute sql, rdExecDirect
             
                
                
               .MoveNext
            Loop
           .Close
        End With
        
        
        
Proximo:
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "fim"
End Sub


Private Sub Imovel_Cem()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, sSoma As String, nDV As Integer
Dim sCep As String, nMax As Integer, nSoma As Integer, x As Integer, nDigito As Integer, sCepNovo As String, nCodigo As Long
Dim RdoAux2 As rdoResultset, RdoAux3 As rdoResultset

sql = "delete from codtmp"
cn.Execute sql, rdExecDirect

sql = "SELECT vwSomaArea.CODREDUZIDO,cadimob.resideimovel From dbo.vwSomaArea LEFT OUTER JOIN dbo.cadimob ON vwSomaArea.codreduzido = cadimob.codreduzido "
sql = sql & "WHERE vwSomaArea.codreduzido IN (SELECT Proprietario.CODREDUZIDO From dbo.Proprietario Where Proprietario.CodCidadao in (500502,506788,519154) AND proprietario.tipoprop = 'C') "
sql = sql & "AND vwSomaArea.soma < 65 AND resideimovel=1"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodigo = !CODREDUZIDO
        
        sql = "SELECT codreduzido FROM proprietario WHERE codcidadao in (500502,506788,519154) AND tipoprop='C'"
        Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        
        
        sql = "insert calculo_removido(ano,codigo,nome) values(" & Year(Now) & "," & nCodigo & ",'CEM')"
        'Sql = "insert codtmp(codigo) values(" & nCodigo & ")"
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"



End Sub

Private Sub Imovel_Cunha()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, sSoma As String, nDV As Integer
Dim sCep As String, nMax As Integer, nSoma As Integer, x As Integer, nDigito As Integer, sCepNovo As String, nCodigo As Long
Dim RdoAux2 As rdoResultset, RdoAux3 As rdoResultset

sql = "delete from codtmp"
cn.Execute sql, rdExecDirect

sql = "SELECT vwSomaArea.CODREDUZIDO,cadimob.resideimovel From dbo.vwSomaArea LEFT OUTER JOIN dbo.cadimob ON vwSomaArea.codreduzido = cadimob.codreduzido "
sql = sql & "WHERE vwSomaArea.codreduzido IN (SELECT Proprietario.CODREDUZIDO From dbo.Proprietario Where Proprietario.CodCidadao in (635000,532915,522758) AND proprietario.tipoprop = 'C') "
sql = sql & "AND vwSomaArea.soma < 65 AND resideimovel=1"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodigo = !CODREDUZIDO
        
        sql = "SELECT codreduzido FROM proprietario WHERE codcidadao in (635000,532915,522758) AND tipoprop='C'"
        Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        
        
        
        sql = "insert codtmp(codigo) values(" & nCodigo & ")"
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"



End Sub

Private Sub Imovel_DePaula()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, sSoma As String, nDV As Integer
Dim sCep As String, nMax As Integer, nSoma As Integer, x As Integer, nDigito As Integer, sCepNovo As String, nCodigo As Long
Dim RdoAux2 As rdoResultset, RdoAux3 As rdoResultset

sql = "delete from codtmp"
cn.Execute sql, rdExecDirect

sql = "SELECT vwSomaArea.CODREDUZIDO,cadimob.resideimovel From dbo.vwSomaArea LEFT OUTER JOIN dbo.cadimob ON vwSomaArea.codreduzido = cadimob.codreduzido "
sql = sql & "WHERE vwSomaArea.codreduzido IN (SELECT Proprietario.CODREDUZIDO From dbo.Proprietario Where Proprietario.CodCidadao in (635000,532915,522758) AND proprietario.tipoprop = 'C') "
sql = sql & "AND vwSomaArea.soma < 65 AND resideimovel=1"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodigo = !CODREDUZIDO
        
        sql = "SELECT codreduzido FROM proprietario WHERE codcidadao in (635000,532915,522758) AND tipoprop='C'"
        Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        
        
        
        sql = "insert codtmp(codigo) values(" & nCodigo & ")"
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"



End Sub

Private Sub Imovel_Limarfe()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, sSoma As String, nDV As Integer
Dim sCep As String, nMax As Integer, nSoma As Integer, x As Integer, nDigito As Integer, sCepNovo As String, nCodigo As Long
Dim RdoAux2 As rdoResultset, RdoAux3 As rdoResultset

sql = "delete from codtmp"
cn.Execute sql, rdExecDirect

sql = "SELECT vwSomaArea.CODREDUZIDO,cadimob.resideimovel From dbo.vwSomaArea LEFT OUTER JOIN dbo.cadimob ON vwSomaArea.codreduzido = cadimob.codreduzido "
sql = sql & "WHERE vwSomaArea.codreduzido IN (SELECT Proprietario.CODREDUZIDO From dbo.Proprietario Where Proprietario.CodCidadao in (614496) AND proprietario.tipoprop = 'C') "
sql = sql & "AND vwSomaArea.soma < 65 AND resideimovel=1"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodigo = !CODREDUZIDO
        
        sql = "SELECT codreduzido FROM proprietario WHERE codcidadao in (614496) AND tipoprop='C'"
        Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        
        
        sql = "insert calculo_removido(ano,codigo,nome) values(" & Year(Now) & "," & nCodigo & ",'LIMARFE')"
        'Sql = "insert codtmp(codigo) values(" & nCodigo & ")"
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"



End Sub

Private Sub Imovel_Arbor()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, sSoma As String, nDV As Integer
Dim sCep As String, nMax As Integer, nSoma As Integer, x As Integer, nDigito As Integer, sCepNovo As String, nCodigo As Long
Dim RdoAux2 As rdoResultset, RdoAux3 As rdoResultset

sql = "delete from codtmp"
cn.Execute sql, rdExecDirect

sql = "SELECT vwSomaArea.CODREDUZIDO,cadimob.resideimovel From dbo.vwSomaArea LEFT OUTER JOIN dbo.cadimob ON vwSomaArea.codreduzido = cadimob.codreduzido "
sql = sql & "WHERE vwSomaArea.codreduzido IN (SELECT Proprietario.CODREDUZIDO From dbo.Proprietario Where Proprietario.CodCidadao in (614498) AND proprietario.tipoprop = 'C') "
sql = sql & "AND vwSomaArea.soma < 65 AND resideimovel=1"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodigo = !CODREDUZIDO
        
        sql = "SELECT codreduzido FROM proprietario WHERE codcidadao in (614498) AND tipoprop='C'"
        Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        
        
        sql = "insert calculo_removido(ano,codigo,nome) values(" & Year(Now) & "," & nCodigo & ",'ARBOR')"
        'Sql = "insert codtmp(codigo) values(" & nCodigo & ")"
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"



End Sub

Private Sub Imovel_Santander()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, sSoma As String, nDV As Integer
Dim sCep As String, nMax As Integer, nSoma As Integer, x As Integer, nDigito As Integer, sCepNovo As String, nCodigo As Long
Dim RdoAux2 As rdoResultset, RdoAux3 As rdoResultset

sql = "delete from codtmp"
cn.Execute sql, rdExecDirect

sql = "SELECT vwSomaArea.CODREDUZIDO,cadimob.resideimovel From dbo.vwSomaArea LEFT OUTER JOIN dbo.cadimob ON vwSomaArea.codreduzido = cadimob.codreduzido "
sql = sql & "WHERE vwSomaArea.codreduzido IN (SELECT Proprietario.CODREDUZIDO From dbo.Proprietario Where Proprietario.CodCidadao in (534364,627466,553296,558963,545926,591703) AND proprietario.tipoprop = 'C') "
sql = sql & "AND vwSomaArea.soma < 65 AND resideimovel=1"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodigo = !CODREDUZIDO
        
        sql = "SELECT codreduzido FROM proprietario WHERE codcidadao in (534364,627466,553296,558963,545926,591703) AND tipoprop='C'"
        Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        
        
        sql = "insert calculo_removido(ano,codigo,nome) values(" & Year(Now) & "," & nCodigo & ",'SANTANDER')"
        'Sql = "insert codtmp(codigo) values(" & nCodigo & ")"
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"



End Sub





Private Sub Corrige_Permei()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, sSoma As String, nDV As Integer
Dim nCodigo As Long, sDataInicio As String, sDataFim As String
Dim RdoAux2 As rdoResultset, RdoAux3 As rdoResultset

sql = "delete from periodomei2"
cn.Execute sql, rdExecDirect

sql = "SELECT * FROM periodomei ORDER BY codigo,datainicio desc"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodigo = !Codigo
        sDataInicio = Format(!datainicio, "dd/mm/yyyy")
        If Not IsNull(!datafim) Then
            sDataFim = Format(!datafim, "dd/mm/yyyy")
        Else
            sDataFim = ""
        End If
        
        
        sql = "SELECT * from periodomei2 where codigo=" & nCodigo & " and datainicio='" & Format(sDataInicio, "mm/dd/yyyy") & "' and "
        If sDataFim <> "" Then
            sql = sql & "datafim='" & Format(sDataFim, "mm/dd/yyyy") & "'"
        Else
            sql = sql & "datafim is null"
        End If
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        
        If RdoAux2.RowCount = 0 Then
            If sDataFim <> "" Then
                sql = "insert periodomei2(codigo,datainicio,datafim,cnpj_base,data_exportacao) values(" & nCodigo & ",'" & Format(sDataInicio, "mm/dd/yyyy") & "','"
                sql = sql & Format(sDataFim, "mm/dd/yyyy") & "','" & !Cnpj_Base & "','" & Format(!data_exportacao, "mm/dd/yyyy hh:mm:ss") & "')"
            Else
                sql = "insert periodomei2(codigo,datainicio,cnpj_base,data_exportacao) values(" & nCodigo & ",'" & Format(sDataInicio, "mm/dd/yyyy") & "','"
                sql = sql & !Cnpj_Base & "','" & Format(!data_exportacao, "mm/dd/yyyy hh:mm:ss") & "')"
            End If
        End If
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"



End Sub


Private Sub Corrige_Vencimentos_Parcelamento()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, nSeq As Integer
Dim RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, nErro As Integer, DataVenctoAtual As Date, DataVenctoAnterior As Date

GoTo Corrige

sql = "delete from codtmp6"
cn.Execute sql, rdExecDirect

sql = "SELECT DISTINCT codreduzido,seqlancamento FROM debitoparcela WHERE anoexercicio>=2023 and codlancamento=20 and (statuslanc=3 or statuslanc=18) AND datavencimento>getdate() ORDER BY codreduzido,seqlancamento"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodReduz = !CODREDUZIDO
        nSeq = !SeqLancamento
        
        nErro = 0
        DataVenctoAnterior = Now
        sql = "select codreduzido,numparcela, datavencimento from debitoparcela where codreduzido=" & nCodReduz & " and codlancamento=20 and seqlancamento=" & nSeq & " and statuslanc=18 order by numparcela"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                DataVenctoAtual = !DataVencimento
                If DataVenctoAnterior > DataVenctoAtual Then
                    nErro = 1
                    Exit Do
                End If
                DataVenctoAnterior = DataVenctoAtual
               .MoveNext
            Loop
           .Close
        End With
        If nErro = 1 Then
            sql = "insert codtmp6(codigo,seq,problem) values(" & nCodReduz & "," & nSeq & "," & nErro & ")"
            cn.Execute sql, rdExecDirect
        End If
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
Exit Sub
Corrige:

Dim DataVencimento As Date

sql = "select codigo,seq from codtmp6 where problem=1 order by codigo,seq"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        CallPb nPos, nTot
        
        nCodReduz = !Codigo
        nSeq = !Seq
                
        sql = "select codreduzido,datavencimento from debitoparcela where codreduzido=" & nCodReduz & " and codlancamento=20 and seqlancamento=" & nSeq & " and numparcela=1"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        DataVencimento = RdoAux2!DataVencimento
        RdoAux2.Close
        
        sql = "select codreduzido,numparcela from debitoparcela where codreduzido=" & nCodReduz & " and codlancamento=20 and seqlancamento=" & nSeq & " and codcomplemento=0 order by numparcela"
        Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux3
            Do Until .EOF
            
                If !NumParcela > 1 Then
                    sql = "update debitoparcela set datavencimento='" & Format(DataVencimento, "mm/dd/yyyy") & "' where codreduzido=" & nCodReduz & " and codlancamento=20 and "
                    sql = sql & "seqlancamento=" & nSeq & " and numparcela=" & !NumParcela
                    cn.Execute sql, rdExecDirect
                End If
                DataVencimento = DateAdd("m", 1, DataVencimento)
               .MoveNext
            Loop
           .Close
        End With
        DoEvents
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With


MsgBox "Fim"

End Sub

Private Sub LimpezaBD()
Dim nCodReduz As Long, nCodCidadao As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long
Dim sTmp As String, nMax As Integer, nSoma As Integer, x As Integer, nMaxCod As Long
Dim RdoAux2 As rdoResultset, RdoAux3 As rdoResultset
Dim nLog1 As Long, nLog2 As Long, nNum1 As Integer, nNum2 As Integer, sCepOld1 As String, sCepOld2 As String, sCepNew1 As String, sCepNew2 As String

sql = "SELECT codcidadao,codlogradouro,numimovel,codlogradouro2,numimovel2,cep,cep2 FROM cidadao WHERE codlogradouro IS NOT NULL OR codlogradouro2 IS NOT null"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodCidadao = !CodCidadao
           
        nLog1 = Val(SubNull(!CodLogradouro))
        nLog2 = Val(SubNull(!CodLogradouro2))
        nNum1 = Val(SubNull(!NUMIMOVEL))
        nNum2 = Val(SubNull(!NUMIMOVEL2))
        sCepOld1 = SubNull(!Cep)
        sCepOld2 = SubNull(!Cep2)
        
        If nLog1 > 0 Then
            sCepNew1 = RetornaNumero(RetornaCEP(nLog1, nNum1))
            If sCepNew1 <> sCepOld1 Then
                sql = "update cidadao set cep='" & sCepNew1 & "' where codcidadao=" & nCodCidadao
                cn.Execute sql, rdExecDirect
            End If
        End If
        
        If nLog2 > 0 Then
            sCepNew2 = RetornaNumero(RetornaCEP(nLog2, nNum2))
            If sCepNew2 <> sCepOld2 Then
                sql = "update cidadao set cep2='" & sCepNew2 & "' where codcidadao=" & nCodCidadao
                cn.Execute sql, rdExecDirect
            End If
        End If
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"



End Sub

Private Sub AreaLaserIptu()
Dim nCodReduz As Long, nCodCidadao As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long
Dim sTmp As String, nMax As Integer, nSoma As Integer, x As Long, nMaxCod As Long, nArea As Double
Dim RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, a2023() As tLaser, bFind As Boolean

ReDim a2023(0)

sql = "SELECT ano,codreduzido,areaconstrucao FROM laseriptu WHERE ano=2023 AND seq=1 order by codreduzido"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodReduz = !CODREDUZIDO
        nArea = !areaconstrucao
        
        ReDim Preserve a2023(UBound(a2023) + 1)
        a2023(UBound(a2023)).Codigo = nCodReduz
        a2023(UBound(a2023)).Area_Predial = nArea
        a2023(UBound(a2023)).Area_Terreno = 0
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

sql = "SELECT ano,codreduzido,areaconstrucao FROM laseriptu WHERE ano=2024 order by codreduzido"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodReduz = !CODREDUZIDO
        nArea = !areaconstrucao
        bFind = False
        
        For x = 1 To UBound(a2023)
            If nCodReduz = a2023(x).Codigo Then
                bFind = True
                Exit For
            End If
        Next
        
        If bFind = True Then
            a2023(x).Area_Terreno = nArea
        End If
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

Open sPathBin & "\codigos.txt" For Output As #1
    For x = 1 To UBound(a2023)
        If a2023(x).Area_Terreno > 0 Then
            If a2023(x).Area_Predial <> a2023(x).Area_Terreno Then
                Print #1, a2023(x).Codigo
            End If
        End If
    Next
Close #1

MsgBox "Fim"



End Sub

Private Sub Corrige_EnderecoLaser()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, sSoma As String, nDV As Integer
Dim sCep As String, nMax As Integer, nSoma As Integer, x As Integer, nDigito As Integer, nTipo As Integer, nCodigo As Long
Dim RdoAux2 As rdoResultset, RdoAux3 As rdoResultset

sql = "delete from codtmp"
'cn.Execute Sql, rdExecDirect

sql = "SELECT * from tributacaoteste..calculo_data order by codigo, tipo"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodigo = !Codigo
        nTipo = !Tipo
        
        sql = "select * from calculo_data where codigo=" & nCodigo & " and tipo=" & nTipo
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
        If RdoAux2.RowCount > 0 Then
            If !Enderecoc <> RdoAux2!Enderecoc Then
                sql = "update calculo_data set enderecoc='" & Mask(!Enderecoc) & "', numeroc=" & !numeroc & ", complc='" & SubNull(!complc) & "', cepc='" & SubNull(!cepc) & "',"
                sql = sql & "bairroc='" & SubNull(!bairroc) & "',cidadec='" & SubNull(!cidadec) & "',ufc='" & SubNull(!ufc) & "' where codigo=" & nCodigo & " and tipo=" & nTipo
                cn.Execute sql, rdExecDirect
            End If
        End If
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"



End Sub

Private Sub Cdas_Nao_Ajuizadas()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, sSoma As String, nDV As Integer
Dim sCep As String, nMax As Integer, nSoma As Integer, x As Integer, nDigito As Integer, nTipo As Integer, nCodigo As Long
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, nIdAjuizamento As Long, nValor As Double
Dim RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, sProcesso As String, bParcial As Boolean, idCda As Long

sql = "delete from codtmp"
cn.Execute sql, rdExecDirect
ConectaIntegrativa

sql = "SELECT Count (CDAs.idCDA) as contador From dbo.CDADebitos INNER JOIN dbo.CDAs ON CDADebitos.idCDA = CDAs.idCDA WHERE CDAs.idDevedor>=100000 and CDAs.idDevedor<=200000 and CDAs.DtLeitura >= '01/01/2024'"
Set rdoAux = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
nTot = rdoAux!contador
rdoAux.Close

sql = "delete from Cdas_Nao_Ajuizadas"
cnInt.Execute sql, rdExecDirect

sql = "SELECT CDAs.idCDA,CDAs.idDevedor,CDADebitos.Exercicio,CDADebitos.Lancamento,CDADebitos.Seq ,CDADebitos.NroParcela ,CDADebitos.ComplParcela ,CDAs.DtLeitura "
sql = sql & "From dbo.CDADebitos INNER JOIN dbo.CDAs ON CDADebitos.idCDA = CDAs.idCDA WHERE CDAs.idDevedor>=500000 and CDAs.DtLeitura >= '01/01/2024'"
'Sql = Sql & "From dbo.CDADebitos INNER JOIN dbo.CDAs ON CDADebitos.idCDA = CDAs.idCDA WHERE CDAs.idDevedor>=100000 and CDAs.idDevedor<=200000 and CDAs.DtLeitura >= '01/01/2024'"
Set rdoAux = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        idCda = !idCda
        nCodigo = !iddevedor
        nAno = !exercicio
        nLanc = !lancamento
        nSeq = !Seq
        nParc = !nroparcela
        nCompl = !complparcela
               
        sql = "select codreduzido, dataajuiza,processocnj from debitoparcela where codreduzido=" & nCodigo & " and anoexercicio=" & nAno & " and codlancamento=" & nLanc & " and "
        sql = sql & "seqlancamento=" & nSeq & " and numparcela=" & nParc & " and codcomplemento=" & nCompl
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux2!processocnj) Or RdoAux2!processocnj = "" Then
            On Error Resume Next
            sql = "insert Cdas_Nao_Ajuizadas(IdCda,IdDevedor,Exercicio,Lancamento,Sequencia,Parcela,Complemento) values(" & idCda & "," & nCodigo & "," & nAno & "," & nLanc & ","
            sql = sql & nSeq & "," & nParc & "," & nCompl & ")"
            cnInt.Execute sql, rdExecDirect
        End If
        
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"

cnInt.Close

End Sub

Private Sub Corrige_Cep_Cdas()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, sSoma As String, nDV As Integer
Dim sCep As String, nMax As Integer, nSoma As Integer, x As Integer, nDigito As Integer, nTipo As Integer, nCodigo As Long
Dim nCodLogr As Integer, nNum As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, nIdAjuizamento As Long, nValor As Double
Dim RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, sProcesso As String, bParcial As Boolean, idCda As Long, idCadastro As Long

ConectaIntegrativa

sql = "SELECT CDAS.idCDA,idDevedor, idCadastro FROM cadastro INNER JOIN CDAS ON Cadastro.idCDA=CDAS.idCDA WHERE idDevedor<50000 ORDER BY CDAs.idCDA"
Set rdoAux = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        idCda = !idCda
        idCadastro = !idCadastro
        nCodigo = !iddevedor
        
        sql = "SELECT codlogr,li_num FROM vwfullimovel WHERE codreduzido=" & nCodigo
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        nCodLogr = RdoAux2!CodLogr
        nNum = RdoAux2!Li_Num
        RdoAux2.Close
        sCep = RetornaNumero(RetornaCEP(CLng(nCodLogr), nNum))
        
        sql = "update cadastro set localcep='" & sCep & "' where idcadastro=" & idCadastro
        cnInt.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"

cnInt.Close

End Sub

Private Sub Cancela_200Reais()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, sObs As String, nSeqLanc As Integer
Dim sCep As String, nMax As Integer, nSoma As Integer, x As Integer, nDigito As Integer, nTipo As Integer, nCodigo As Long
Dim nCodLogr As Integer, nNum As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, nIdAjuizamento As Long, nValor As Double
Dim RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, sProcesso As String, bParcial As Boolean, idCda As Long, idCadastro As Long

sObs = "Débito cancelado conforme Art. 25 da LEI DO REFIS 2023, solicitado pelo Setor de Dívida Ativa."

sql = "SELECT * FROM DEVEDOR_CANCELAR WHERE CANCELAR=1"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodigo = !Codigo
        nAno = !ano
        nLanc = !Lanc
        nSeq = !Seq
        nParc = !Parc
        nCompl = !Compl
        
        sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA where codreduzido=" & nCodigo & " and anoexercicio=" & nAno & " and codlancamento=" & nLanc & " and "
        sql = sql & "seqlancamento=" & nSeq & " and numparcela=" & nParc & " and codcomplemento=" & nCompl
        Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux3
            If IsNull(!maximo) Then
                nSeqLanc = 1
            Else
                nSeqLanc = !maximo + 1
            End If
           .Close
        End With
        
        sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & nCodigo & "," & nAno & ","
        sql = sql & nLanc & "," & nSeq & "," & nParc & "," & nCompl & "," & nSeqLanc & ",'" & sObs & "'," & RetornaUsuarioID(NomeDeLogin) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
        cn.Execute sql, rdExecDirect
        
        sql = "update debitoparcela set statuslanc=5 where codreduzido=" & nCodigo & " and anoexercicio=" & nAno & " and codlancamento=" & nLanc & " and seqlancamento=" & nSeq & " and "
        sql = sql & "numparcela=" & nParc & " and codcomplemento=" & nCompl & " and statuslanc=38"
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"


End Sub

Private Sub CorrigeStatusPago()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, sObs As String, nSeqLanc As Integer
Dim sCep As String, nMax As Integer, nSoma As Integer, x As Integer, nDigito As Integer, nTipo As Integer, nCodigo As Long
Dim nCodLogr As Integer, nNum As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, nIdAjuizamento As Long, nValor As Double
Dim RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, sProcesso As String, bParcial As Boolean, idCda As Long, idCadastro As Long


sql = "SELECT * FROM parceladocumento WHERE numdocumento BETWEEN 22130366 AND 22160363 AND codlancamento=41 ORDER BY codreduzido"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodigo = !CODREDUZIDO
        nAno = !AnoExercicio
        nLanc = !CodLancamento
        nSeq = !SeqLancamento
        nParc = !NumParcela
        nCompl = !CODCOMPLEMENTO
        
        sql = "update debitoparcela set statuslanc=5 where codreduzido=" & nCodigo & " and anoexercicio=" & nAno & " and codlancamento=" & nLanc & " and seqlancamento<>" & nSeq & " and "
        sql = sql & "numparcela=" & nParc & " and codcomplemento=" & nCompl
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"


End Sub


Private Sub RemoveProtocolo()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, sObs As String, nSeqLanc As Integer
Dim sCep As String, nMax As Integer, nSoma As Integer, x As Integer, nDigito As Integer, nTipo As Integer, nCodigo As Long
Dim nCodLogr As Integer, nNum As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, nIdAjuizamento As Long, nValor As Double
Dim RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, sProcesso As String, bParcial As Boolean, idCda As Long, idCadastro As Long, nCodTributo As Integer

nCount = 1
sql = "SELECT distinct codreduzido from debitoparcela where codlancamento=11 and statuslanc=3"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodigo = !CODREDUZIDO
        
        sql = "select * from debitoparcela where codreduzido=" & nCodigo & " and statuslanc=3"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
        With RdoAux2
             If .RowCount = 1 Then
                nAno = !AnoExercicio
                nLanc = !CodLancamento
                nSeq = !SeqLancamento
                nParc = !NumParcela
                nCompl = !CODCOMPLEMENTO
                sql = "select codtributo from debitotributo where codreduzido=" & nCodigo & " and anoexercicio=" & nAno & " and codlancamento=" & nLanc & " and "
                sql = sql & "seqlancamento=" & nSeq & " and numparcela=" & nParc & " and codcomplemento=" & nCompl
                Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
                If RdoAux3.RowCount = 1 Then
                    nCodTributo = RdoAux3!CodTributo
                    If nCodTributo = 28 Then
                        sql = "insert codtmp(codigo) values(" & nCodigo & ")"
                        cn.Execute sql, rdExecDirect
                        nCount = nCount + 1
                    End If
                End If
             End If
             
        End With
                
                
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"


End Sub

Private Sub RemoveIPTUCalculo()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, sObs As String, nSeqLanc As Integer
Dim sCep As String, nMax As Integer, nSoma As Integer, x As Integer, nDigito As Integer, nTipo As Integer, nCodigo As Long
Dim nCodLogr As Integer, nNum As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, nIdAjuizamento As Long, nValor As Double
Dim RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, sProcesso As String, bParcial As Boolean, idCda As Long, idCadastro As Long, nCodTributo As Integer

nCount = 1
sql = "delete from laseriptu WHERE ano=2025 AND codreduzido IN (SELECT codigo FROM calculo_removido where ano=2024)"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)

sql = "delete from calculo_resumo WHERE ano=2025 AND codigo IN (SELECT codigo FROM calculo_removido where ano=2024)"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)

sql = "delete from debitoparcela WHERE anoexercicio=2025 AND codreduzido IN (SELECT codigo FROM calculo_removido where ano=2024) AND codlancamento=1"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)

sql = "delete from debitotributo WHERE anoexercicio=2025 AND codreduzido IN (SELECT codigo FROM calculo_removido where ano=2024) AND codlancamento=1"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)

sql = "delete from parceladocumento WHERE anoexercicio=2025 AND codreduzido IN (SELECT codigo FROM calculo_removido where ano=2024) AND codlancamento=1"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)

MsgBox "Fim"

End Sub


Private Sub Corrige_CdasProtesto()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, sSoma As String, nDV As Integer
Dim sCep As String, nMax As Integer, nSoma As Integer, x As Integer, nDigito As Integer, nTipo As Integer, nCodigo As Long
Dim nCodLogr As Integer, nNum As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, nIdAjuizamento As Long, nValor As Double
Dim RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, sProcesso As String, bParcial As Boolean, idCda As Long, idCadastro As Long
Dim nValorP As Double, nValorM As Double, nValorJ As Double, nValorC As Double, nValorT As Double

ConectaIntegrativa

sql = "SELECT idCDA FROM CDAs_Protesto WHERE YEAR(DtGeracao)=2024 AND vlrOriginal IS null"
Set rdoAux = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        idCda = !idCda
        
        sql = "SELECT * FROM CDADebitos_Protesto WHERE idCDA=" & idCda
        Set RdoAux2 = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        nValorP = RdoAux2!vlroriginal
        nValorJ = RdoAux2!vlrjuros
        nValorM = RdoAux2!vlrmultas
        nValorC = RdoAux2!vlrcorrecao
        nValorT = nValorP + nValorJ + nValorM + nValorC
        RdoAux2.Close
       
        sql = "update CDAs_Protesto set vlroriginal=" & Virg2Ponto(CStr(nValorP)) & ",vlrmultas=" & Virg2Ponto(CStr(nValorM)) & ",vlrjuros=" & Virg2Ponto(CStr(nValorJ)) & ","
        sql = sql & "vlrcorrecao=" & Virg2Ponto(CStr(nValorC)) & ",vlrtotal=" & Virg2Ponto(CStr(nValorT)) & ",dtleitura=null where idcda=" & idCda
        cnInt.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"

cnInt.Close

End Sub


Private Sub Corrige_BaixaGiss()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, nMaxDoc As Long

ConectaEicon

sql = "SELECT max(num_documento) as maximo FROM tb_inter_boletos_giss  WHERE num_documento>2000000"
Set rdoAux = cnEicon.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
nMaxDoc = rdoAux!maximo
rdoAux.Close

'Sql = "SELECT  pd.codreduzido,pd.anoexercicio,pd.codlancamento,pd.seqlancamento,pd.numparcela,pd.codcomplemento,numdocumento,statuslanc "
'Sql = Sql & "FROM parceladocumento pd INNER JOIN debitoparcela dp ON pd.codreduzido = dp.codreduzido AND pd.anoexercicio = dp.anoexercicio AND pd.codlancamento = dp.codlancamento "
'Sql = Sql & "AND pd.seqlancamento = dp.seqlancamento AND pd.numparcela = dp.numparcela AND pd.codcomplemento = dp.codcomplemento "
'Sql = Sql & "WHERE numdocumento BETWEEN 2000001 AND " & nMaxDoc & " AND statuslanc NOT IN (3,5,12,37)"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    nTot = .RowCount
'    nPos = 1
'    Do Until .EOF
'        If nPos Mod 10 = 0 Then
'           CallPb nPos, nTot
'        End If
'        On Error Resume Next
'        Sql = "insert giss_guia(documento,codigo,ano,lancamento,seq,parcela,complemento,situacao,enviado) values(" & !NumDocumento & "," & !CODREDUZIDO & ","
'        Sql = Sql & !AnoExercicio & "," & !CodLancamento & "," & !SeqLancamento & "," & !NumParcela & "," & !CODCOMPLEMENTO & "," & !statuslanc & ",0)"
'        cn.Execute Sql, rdExecDirect
'        nPos = nPos + 1
 '       DoEvents
 '      .MoveNext
 '   Loop
 '  .Close
'End With
'baixa ok
sql = "SELECT num_documento,SUM(valor_titulo) AS soma FROM tb_inter_baixa_detalhe  WHERE num_documento BETWEEN 2000000 AND 2282846 GROUP BY num_documento ORDER BY num_documento"
Set rdoAux = cnEicon.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        If !soma > 0 Then
            sql = "update giss_guia set enviado=1 where documento=" & Val(!num_documento)
            cn.Execute sql, rdExecDirect
        End If
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With



MsgBox "Fim"

cnEicon.Close

End Sub


Private Sub Corrige_BaixaDocGiss()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, nNumDoc As Long, RdoAux2 As rdoResultset, nValor As Double, RdoAux3 As rdoResultset

ConectaEicon

sql = "SELECT * FROM tb_inter_baixa_detalhe WHERE num_documento IN (SELECT CODIGO FROM TRIBUTACAO..CODTMP) ORDER BY num_documento"
Set rdoAux = cnEicon.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nNumDoc = !num_documento
        sql = "select valor from codtmp7 where documento=" & nNumDoc
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        nValor = RdoAux2!valor
        RdoAux2.Close
        sql = "SELECT * FROM tb_inter_baixa_detalhe WHERE num_documento=" & nNumDoc
        Set RdoAux3 = cnEicon.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux3.RowCount = 0 Then
'baixa ok
            sql = "insert tb_inter_baixa_detalhe(cod_cliente,cod_banco,num_sequencia,num_documento,linha,timestamp,valor_titulo,valor_pago,data_pagamento,valor_encargos,"
            sql = sql & "descricao_linha_t,descricao_linha_u) values(" & 2177 & "," & !cod_banco & "," & 0 & "," & nNumDoc & "," & 0 & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "',"
            sql = sql & Virg2Ponto(CStr(nValor)) & "," & Virg2Ponto(CStr(nValor)) & ",'" & Format(!Data_Pagamento, "mm/dd/yyyy") & "'," & 0 & ",'"
            sql = sql & "" & "','" & "" & "')"
            cnEicon.Execute sql, rdExecDirect
        End If


        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With



MsgBox "Fim"

cnEicon.Close

End Sub


Private Sub Incluir_para_registro()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, nDoc As Long
Dim nCodigo As Long, sCPF As String, nQtdeParc As Integer, aCalculo() As tCalculoResumo, sNome As String, sEndereco As String, sBairro As String

'nDoc = 22767075

'Sql = "SELECT codigo FROM calculo_resumo WHERE ano=2025 and codigo BETWEEN 42008 and 42368 order by codigo"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    Do Until .EOF
'        nCodigo = !Codigo
'        Sql = "update calculo_resumo set documento0=" & nDoc & " where ano=2025 and codigo=" & nCodigo
'        cn.Execute Sql, rdExecDirect
'        nPos = .AbsolutePosition
'        Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES("
'        Sql = Sql & nCodigo & ",2025,1,0,0,0," & nDoc & ")"
'        cn.Execute Sql, rdExecDirect
'
'        nDoc = nDoc + 1
'        '###########
'
'        Sql = "update calculo_resumo set documento91=" & nDoc & " where ano=2025 and codigo=" & nCodigo
'        cn.Execute Sql, rdExecDirect
'
'        Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES("
'        Sql = Sql & nCodigo & ",2025,1,0,0,91," & nDoc & ")"
'        cn.Execute Sql, rdExecDirect
'
'        nDoc = nDoc + 1
'        '###########
'
'        Sql = "update calculo_resumo set documento92=" & nDoc & " where ano=2025 and codigo=" & nCodigo
'        cn.Execute Sql, rdExecDirect
'
'        Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES("
'        Sql = Sql & nCodigo & ",2025,1,0,0,92," & nDoc & ")"
'        cn.Execute Sql, rdExecDirect
'
'        nDoc = nDoc + 1
'        '###########
'
'        Sql = "update calculo_resumo set documento1=" & nDoc & " where ano=2025 and codigo=" & nCodigo
'        cn.Execute Sql, rdExecDirect
'
'        Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES("
'        Sql = Sql & nCodigo & ",2025,1,0,1,0," & nDoc & ")"
'        cn.Execute Sql, rdExecDirect
'
'        nDoc = nDoc + 1
'        '###########
'
'        Sql = "update calculo_resumo set documento2=" & nDoc & " where ano=2025 and codigo=" & nCodigo
'        cn.Execute Sql, rdExecDirect
'
'        Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES("
'        Sql = Sql & nCodigo & ",2025,1,0,2,0," & nDoc & ")"
'        cn.Execute Sql, rdExecDirect
'
'        nDoc = nDoc + 1
'        '###########
'
'        Sql = "update calculo_resumo set documento3=" & nDoc & " where ano=2025 and codigo=" & nCodigo
'        cn.Execute Sql, rdExecDirect
'
'        Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES("
'        Sql = Sql & nCodigo & ",2025,1,0,3,0," & nDoc & ")"
'        cn.Execute Sql, rdExecDirect
'
'        nDoc = nDoc + 1
'        '###########
'
'        Sql = "update calculo_resumo set documento4=" & nDoc & " where ano=2025 and codigo=" & nCodigo
'        cn.Execute Sql, rdExecDirect
'
'        Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES("
'        Sql = Sql & nCodigo & ",2025,1,0,4,0," & nDoc & ")"
'        cn.Execute Sql, rdExecDirect
'
'        nDoc = nDoc + 1
'        '###########
'
'        Sql = "update calculo_resumo set documento5=" & nDoc & " where ano=2025 and codigo=" & nCodigo
'        cn.Execute Sql, rdExecDirect
'
'        Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES("
'        Sql = Sql & nCodigo & ",2025,1,0,5,0," & nDoc & ")"
'        cn.Execute Sql, rdExecDirect
'
'        nDoc = nDoc + 1
'        '###########
'
'        Sql = "update calculo_resumo set documento6=" & nDoc & " where ano=2025 and codigo=" & nCodigo
'        cn.Execute Sql, rdExecDirect
'
'        Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES("
'        Sql = Sql & nCodigo & ",2025,1,0,6,0," & nDoc & ")"
'        cn.Execute Sql, rdExecDirect
'
'        nDoc = nDoc + 1
'        '###########
'
'        Sql = "update calculo_resumo set documento7=" & nDoc & " where ano=2025 and codigo=" & nCodigo
'        cn.Execute Sql, rdExecDirect
'
'        Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES("
'        Sql = Sql & nCodigo & ",2025,1,0,7,0," & nDoc & ")"
'        cn.Execute Sql, rdExecDirect
'
'        nDoc = nDoc + 1
'        '###########
'
'        Sql = "update calculo_resumo set documento8=" & nDoc & " where ano=2025 and codigo=" & nCodigo
'        cn.Execute Sql, rdExecDirect
'
'        Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES("
'        Sql = Sql & nCodigo & ",2025,1,0,8,0," & nDoc & ")"
'        cn.Execute Sql, rdExecDirect
'
'        nDoc = nDoc + 1
'        '###########
'
'        Sql = "update calculo_resumo set documento9=" & nDoc & " where ano=2025 and codigo=" & nCodigo
'        cn.Execute Sql, rdExecDirect
'
'        Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES("
'        Sql = Sql & nCodigo & ",2025,1,0,9,0," & nDoc & ")"
'        cn.Execute Sql, rdExecDirect
'
'        nDoc = nDoc + 1
'        '###########
'
'        Sql = "update calculo_resumo set documento10=" & nDoc & " where ano=2025 and codigo=" & nCodigo
'        cn.Execute Sql, rdExecDirect
'
'        Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES("
'        Sql = Sql & nCodigo & ",2025,1,0,10,0," & nDoc & ")"
'        cn.Execute Sql, rdExecDirect
'
'        nDoc = nDoc + 1
'        '###########
'
'        Sql = "update calculo_resumo set documento11=" & nDoc & " where ano=2025 and codigo=" & nCodigo
'        cn.Execute Sql, rdExecDirect
'
'        Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES("
'        Sql = Sql & nCodigo & ",2025,1,0,11,0," & nDoc & ")"
'        cn.Execute Sql, rdExecDirect
'
'        nDoc = nDoc + 1
'        '###########
'
'
'        Sql = "update calculo_resumo set documento12=" & nDoc & " where ano=2025 and codigo=" & nCodigo
'        cn.Execute Sql, rdExecDirect
'
'        Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES("
'        Sql = Sql & nCodigo & ",2025,1,0,12,0," & nDoc & ")"
'        cn.Execute Sql, rdExecDirect
'
'        nDoc = nDoc + 1
'        '###########
'
'       .MoveNext
'    Loop
'   .Close
'End With
'
'Exit Sub
'

sql = "SELECT codreduzido,nomecidadao,cpf,cnpj,LOGRADOURO,li_num,li_compl,descbairro,codlogr FROM vwFULLIMOVEL WHERE codreduzido BETWEEN 42009 and 42368"
'Sql = "SELECT codreduzido,nomecidadao,cpf,cnpj,LOGRADOURO,li_num,li_compl,descbairro,codlogr FROM vwFULLIMOVEL WHERE codreduzido = 42008"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodigo = !CODREDUZIDO
        sCep = RetornaNumero(RetornaCEP(!CodLogr, !Li_Num))
        If SubNull(!Cnpj) = "" Then
            sCPF = !cpf
        Else
            sCPF = !Cnpj
        End If
        sCPF = RetornaNumero(sCPF)
        sNome = Mask(Left(!nomecidadao, 40))
        sEndereco = Left(Mask(!Logradouro) & ", " & !Li_Num & " " & !Li_Compl, 40)
        sBairro = Mask(Left(!DescBairro, 15))
        
        ReDim aCalculo(0)
        
        sql = "SELECT * FROM calculo_resumo WHERE ano=2025 AND codigo=" & nCodigo
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                nQtdeParc = !Qtde_Parcela
                aCalculo(0).nDoc0 = !documento0
                aCalculo(0).nDoc91 = !documento91
                aCalculo(0).nDoc92 = !documento92
                aCalculo(0).nDoc1 = !documento1
                aCalculo(0).nDoc2 = !documento2
                aCalculo(0).nDoc3 = !documento3
                aCalculo(0).nDoc4 = !documento4
                aCalculo(0).nDoc5 = !documento5
                aCalculo(0).nDoc6 = !documento6
                aCalculo(0).nDoc7 = !documento7
                aCalculo(0).nDoc8 = !documento8
                aCalculo(0).nDoc9 = !documento9
                aCalculo(0).nDoc10 = !documento10
                aCalculo(0).nDoc11 = !documento11
                aCalculo(0).nDoc12 = !documento12
                aCalculo(0).nValor0 = !valor0
                aCalculo(0).nValor91 = !valor91
                aCalculo(0).nValor92 = !valor92
                aCalculo(0).nValor1 = !VALOR1
                aCalculo(0).DataVento1 = Format(!vencimento1, "mm/dd/yyyy")
                aCalculo(0).DataVento2 = Format(!vencimento2, "mm/dd/yyyy")
                aCalculo(0).DataVento3 = Format(!vencimento3, "mm/dd/yyyy")
                aCalculo(0).DataVento4 = Format(!vencimento4, "mm/dd/yyyy")
                aCalculo(0).DataVento5 = Format(!vencimento5, "mm/dd/yyyy")
                aCalculo(0).DataVento6 = Format(!vencimento6, "mm/dd/yyyy")
                aCalculo(0).DataVento7 = Format(!vencimento7, "mm/dd/yyyy")
                aCalculo(0).DataVento8 = Format(!vencimento8, "mm/dd/yyyy")
                aCalculo(0).DataVento9 = Format(!vencimento9, "mm/dd/yyyy")
                aCalculo(0).DataVento10 = Format(!vencimento10, "mm/dd/yyyy")
                aCalculo(0).DataVento11 = Format(!vencimento11, "mm/dd/yyyy")
                aCalculo(0).DataVento12 = Format(!vencimento12, "mm/dd/yyyy")
               .MoveNext
            Loop
           .Close
        End With
                
                        
        With aCalculo(0)
            sql = "insert ficha_compensacao_documento(numero_documento,data_vencimento,valor_documento,nome,cpf,endereco,bairro,cep,cidade,uf) values(" & .nDoc0 & ",'"
            sql = sql & .DataVento1 & "'," & Virg2Ponto(CStr(.nValor0)) & ",'" & sNome & "','" & sCPF & "','" & sEndereco & "','" & sBairro & "','" & sCep & "','JABOTICABAL','SP')"
            cn.Execute sql, rdExecDirect
        
            sql = "insert ficha_compensacao_documento(numero_documento,data_vencimento,valor_documento,nome,cpf,endereco,bairro,cep,cidade,uf) values(" & .nDoc91 & ",'"
            sql = sql & .DataVento2 & "'," & Virg2Ponto(CStr(.nValor91)) & ",'" & sNome & "','" & sCPF & "','" & sEndereco & "','" & sBairro & "','" & sCep & "','JABOTICABAL','SP')"
            cn.Execute sql, rdExecDirect
        
            sql = "insert ficha_compensacao_documento(numero_documento,data_vencimento,valor_documento,nome,cpf,endereco,bairro,cep,cidade,uf) values(" & .nDoc92 & ",'"
            sql = sql & .DataVento3 & "'," & Virg2Ponto(CStr(.nValor92)) & ",'" & sNome & "','" & sCPF & "','" & sEndereco & "','" & sBairro & "','" & sCep & "','JABOTICABAL','SP')"
            cn.Execute sql, rdExecDirect
        
            sql = "insert ficha_compensacao_documento(numero_documento,data_vencimento,valor_documento,nome,cpf,endereco,bairro,cep,cidade,uf) values(" & .nDoc1 & ",'"
            sql = sql & .DataVento1 & "'," & Virg2Ponto(CStr(.nValor1)) & ",'" & sNome & "','" & sCPF & "','" & sEndereco & "','" & sBairro & "','" & sCep & "','JABOTICABAL','SP')"
            cn.Execute sql, rdExecDirect

            sql = "insert ficha_compensacao_documento(numero_documento,data_vencimento,valor_documento,nome,cpf,endereco,bairro,cep,cidade,uf) values(" & .nDoc2 & ",'"
            sql = sql & .DataVento2 & "'," & Virg2Ponto(CStr(.nValor1)) & ",'" & sNome & "','" & sCPF & "','" & sEndereco & "','" & sBairro & "','" & sCep & "','JABOTICABAL','SP')"
            cn.Execute sql, rdExecDirect

            sql = "insert ficha_compensacao_documento(numero_documento,data_vencimento,valor_documento,nome,cpf,endereco,bairro,cep,cidade,uf) values(" & .nDoc3 & ",'"
            sql = sql & .DataVento3 & "'," & Virg2Ponto(CStr(.nValor1)) & ",'" & sNome & "','" & sCPF & "','" & sEndereco & "','" & sBairro & "','" & sCep & "','JABOTICABAL','SP')"
            cn.Execute sql, rdExecDirect

            sql = "insert ficha_compensacao_documento(numero_documento,data_vencimento,valor_documento,nome,cpf,endereco,bairro,cep,cidade,uf) values(" & .nDoc4 & ",'"
            sql = sql & .DataVento4 & "'," & Virg2Ponto(CStr(.nValor1)) & ",'" & sNome & "','" & sCPF & "','" & sEndereco & "','" & sBairro & "','" & sCep & "','JABOTICABAL','SP')"
            cn.Execute sql, rdExecDirect

            sql = "insert ficha_compensacao_documento(numero_documento,data_vencimento,valor_documento,nome,cpf,endereco,bairro,cep,cidade,uf) values(" & .nDoc5 & ",'"
            sql = sql & .DataVento5 & "'," & Virg2Ponto(CStr(.nValor1)) & ",'" & sNome & "','" & sCPF & "','" & sEndereco & "','" & sBairro & "','" & sCep & "','JABOTICABAL','SP')"
            cn.Execute sql, rdExecDirect

            sql = "insert ficha_compensacao_documento(numero_documento,data_vencimento,valor_documento,nome,cpf,endereco,bairro,cep,cidade,uf) values(" & .nDoc6 & ",'"
            sql = sql & .DataVento6 & "'," & Virg2Ponto(CStr(.nValor1)) & ",'" & sNome & "','" & sCPF & "','" & sEndereco & "','" & sBairro & "','" & sCep & "','JABOTICABAL','SP')"
            cn.Execute sql, rdExecDirect

            sql = "insert ficha_compensacao_documento(numero_documento,data_vencimento,valor_documento,nome,cpf,endereco,bairro,cep,cidade,uf) values(" & .nDoc7 & ",'"
            sql = sql & .DataVento7 & "'," & Virg2Ponto(CStr(.nValor1)) & ",'" & sNome & "','" & sCPF & "','" & sEndereco & "','" & sBairro & "','" & sCep & "','JABOTICABAL','SP')"
            cn.Execute sql, rdExecDirect

            sql = "insert ficha_compensacao_documento(numero_documento,data_vencimento,valor_documento,nome,cpf,endereco,bairro,cep,cidade,uf) values(" & .nDoc8 & ",'"
            sql = sql & .DataVento8 & "'," & Virg2Ponto(CStr(.nValor1)) & ",'" & sNome & "','" & sCPF & "','" & sEndereco & "','" & sBairro & "','" & sCep & "','JABOTICABAL','SP')"
            cn.Execute sql, rdExecDirect

            sql = "insert ficha_compensacao_documento(numero_documento,data_vencimento,valor_documento,nome,cpf,endereco,bairro,cep,cidade,uf) values(" & .nDoc9 & ",'"
            sql = sql & .DataVento9 & "'," & Virg2Ponto(CStr(.nValor1)) & ",'" & sNome & "','" & sCPF & "','" & sEndereco & "','" & sBairro & "','" & sCep & "','JABOTICABAL','SP')"
            cn.Execute sql, rdExecDirect

            sql = "insert ficha_compensacao_documento(numero_documento,data_vencimento,valor_documento,nome,cpf,endereco,bairro,cep,cidade,uf) values(" & .nDoc10 & ",'"
            sql = sql & .DataVento10 & "'," & Virg2Ponto(CStr(.nValor1)) & ",'" & sNome & "','" & sCPF & "','" & sEndereco & "','" & sBairro & "','" & sCep & "','JABOTICABAL','SP')"
            cn.Execute sql, rdExecDirect

            sql = "insert ficha_compensacao_documento(numero_documento,data_vencimento,valor_documento,nome,cpf,endereco,bairro,cep,cidade,uf) values(" & .nDoc11 & ",'"
            sql = sql & .DataVento11 & "'," & Virg2Ponto(CStr(.nValor1)) & ",'" & sNome & "','" & sCPF & "','" & sEndereco & "','" & sBairro & "','" & sCep & "','JABOTICABAL','SP')"
            cn.Execute sql, rdExecDirect

            sql = "insert ficha_compensacao_documento(numero_documento,data_vencimento,valor_documento,nome,cpf,endereco,bairro,cep,cidade,uf) values(" & .nDoc12 & ",'"
            sql = sql & .DataVento12 & "'," & Virg2Ponto(CStr(.nValor1)) & ",'" & sNome & "','" & sCPF & "','" & sEndereco & "','" & sBairro & "','" & sCep & "','JABOTICABAL','SP')"
            cn.Execute sql, rdExecDirect



        End With
                
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"


End Sub

Private Sub CorrigeCPFCNPJ()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, sCPFCNPJ As String

sql = "update cidadao set cpf=null where cpf=''"
cn.Execute sql, rdExecDirect

sql = "update cidadao set cnpj=null where cnpj=''"
cn.Execute sql, rdExecDirect

sql = "update cidadao SET cnpj=NULL WHERE cnpj IS NOT NULL AND cpf IS NOT null"
cn.Execute sql, rdExecDirect


sql = "select codcidadao,cnpj from cidadao where cnpj is not null"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodigo = !CodCidadao
        sCPFCNPJ = RetornaNumero(!Cnpj)
        
        sql = "update cidadao set cnpj='" & sCPFCNPJ & "' where codcidadao=" & nCodigo
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"


End Sub

Private Sub MudaStatus2019()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos As Long, nTot As Long, Sql2 As String
Dim sCodigos As String, aSuspensoCod() As Long, lResult As Long, aTemPago() As Long

On Error GoTo Erro

ReDim aSuspensoCod(0)
sql = "SELECT codmobiliario From vwMOBILIARIOSUSPENSO Where (codtipoevento = 2) order by codmobiliario"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurReadOnly)
With rdoAux
    Do Until .EOF
        ReDim Preserve aSuspensoCod(UBound(aSuspensoCod) + 1)
        aSuspensoCod(UBound(aSuspensoCod)) = !codmobiliario
       .MoveNext
    Loop
   .Close
End With


sql = "truncate table codtmp"
cn.Execute sql, rdExecDirect
sCodigos = ""
sql = "select codigomob from mobiliario where dataencerramento is null"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    
    Do Until .EOF
        nCodReduz = !codigomob
        lResult = BinarySearchLong(aSuspensoCod(), !codigomob)
        If lResult > -1 Then
            GoTo Proximo
        End If

        sql = "INSERT  codtmp(codigo) values( " & !codigomob & ")"
        cn.Execute sql, rdExecDirect

Proximo:
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"


Grava:


Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub RelatorioRefis()
Dim sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, RdoAux4 As rdoResultset
Dim nCodigo As Long, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, nNumDoc As Long, qd As New rdoQuery
Dim aDebito() As Debito, dDataDoc As Date, nAnoTmp As Integer, bFind As Boolean, nEval As Integer, nPos As Long, nTot As Long, nValorTotal As Double
Dim nValor2024 As Double, nValorOutros As Double, nPlano As Integer, nPerc As Double, nValorP As Double, nValorM As Double, nValorJ As Double, nValorC As Double

sql = "DELETE FROM CODTMP3"
cn.Execute sql, rdExecDirect

Set qd.ActiveConnection = cn
qd.QueryTimeout = 0

sql = "select codigo from codtmp order by codigo"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
nPos = 1
nTot = rdoAux.RowCount
Do Until rdoAux.EOF
    If nPos Mod 10 = 0 Then
       CallPb nPos, nTot
    End If
    nNumDoc = rdoAux!Codigo
    sql = "select datapagamento from debitopago where numdocumento=" & nNumDoc
    Set RdoAux4 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
    dDataDoc = RdoAux4!DataPagamento
    RdoAux4.Close
    nPlano = 0
'    sql = "select plano from parceladocumento where numdocumento=" & nNumDoc
'    Set RdoAux4 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
'    If IsNull(RdoAux4!plano) Then
'        nPlano = 65
'    Else
'        nPlano = RdoAux4!plano
'    End If
'    RdoAux4.Close
    
    nValor2024 = 0: nValorOutros = 0
    
    sql = "select * from parceladocumento where numdocumento=" & nNumDoc
    Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        ReDim aDebito(0)
        Do Until .EOF
            nCodigo = !CODREDUZIDO
            nAno = !AnoExercicio
            nLanc = !CodLancamento
            nSeq = !SeqLancamento
            nParc = !NumParcela
            nCompl = !CODCOMPLEMENTO
            
            On Error Resume Next
            RdoAux3.Close
            On Error GoTo 0
            qd.sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
            qd(0) = nCodigo
            qd(1) = nCodigo
            qd(2) = nAno
            qd(3) = nAno
            qd(4) = nLanc
            qd(5) = nLanc
            qd(6) = nSeq
            qd(7) = nSeq
            qd(8) = nParc
            qd(9) = nParc
            qd(10) = nCompl
            qd(11) = nCompl
            qd(12) = 0
            qd(13) = 99
            qd(14) = Format(dDataDoc, "mm/dd/yyyy")
            qd(15) = "GTI"
            Set RdoAux3 = qd.OpenResultset(rdOpenKeyset)
            With RdoAux3
                Do Until .EOF
                    nValorP = !VALORTRIBUTO
                    nValorM = !ValorMulta
                    nValorJ = !ValorJuros
                    nValorC = !valorcorrecao

                    nPerc = 0
                    If nPlano = 65 Then
                       nPerc = 100
                    ElseIf nPlano = 66 Then
                       nPerc = 90
                    ElseIf nPlano = 67 Then
                       nPerc = 80
                    End If
                    nValorM = nValorM - nValorM * nPerc / 100
                    nValorJ = nValorJ - nValorJ * nPerc / 100
                    nValorTotal = nValorP + nValorM + nValorJ + nValorC
                    
                    If !AnoExercicio > 2023 Then
                        nValor2024 = nValor2024 + nValorTotal
                    Else
                        nValorOutros = nValorOutros + nValorTotal
                    End If
                   .MoveNext
                Loop
               .Close
            End With
            DoEvents
           .MoveNext
        Loop
       .Close
    End With
    
    sql = "INSERT CODTMP3 (DOCUMENTO,VALOR2024,VALOR_OUTROS) VALUES(" & nNumDoc & "," & Virg2Ponto(CStr(nValor2024)) & "," & Virg2Ponto(CStr(nValorOutros)) & ")"
    cn.Execute sql, rdExecDirect
        
    nPos = nPos + 1
    rdoAux.MoveNext
Loop
rdoAux.Close
MsgBox "fim"

End Sub

Private Sub ParcelamentoParaBase()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, nSeq As Integer, sProcesso As String
Dim bCancelado As Boolean, nAno As Integer, nNumero As Long, nDoc As Long
sql = "delete from parcelamento2eicon"
cn.Execute sql, rdExecDirect

sql = "SELECT DISTINCT p.Numprocesso FROM Origemreparc p INNER JOIN Processoreparc m ON p.Numprocesso = m.Numprocesso WHERE p.Codreduzido >= 100000 AND "
sql = sql & "  p.Codreduzido < 200000 AND p.Anoexercicio > 2017  AND p.Codlancamento = 5"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        sProcesso = !NumProcesso
        
        sql = "SELECT * FROM processoreparc WHERE numprocesso='" & sProcesso & "'"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        nCodigo = RdoAux2!CODIGORESP
        bCancelado = RdoAux2!Cancelado
        sProcesso = RdoAux2!NumProcesso
        nAno = RdoAux2!AnoProc
        nNumero = RdoAux2!NumProc
        RdoAux2.Close
        
        sql = "SELECT distinct numdocumento  FROM origemreparc INNER JOIN parceladocumento ON origemreparc.codreduzido = parceladocumento.codreduzido AND origemreparc.anoexercicio = parceladocumento.anoexercicio AND origemreparc.codlancamento = parceladocumento.codlancamento AND numsequencia=parceladocumento.seqlancamento AND origemreparc.numparcela = parceladocumento.numparcela AND origemreparc.codcomplemento = parceladocumento.codcomplemento "
        sql = sql & "WHERE numprocesso='" & sProcesso & "' AND numdocumento BETWEEN 2000000 AND 2500000 ORDER BY numdocumento"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                nDoc = RdoAux2!NumDocumento
                sql = "insert parcelamento2eicon(codigo,ano,numero,cancelado,documento) values(" & nCodigo & "," & nAno & "," & nNumero & "," & IIf(bCancelado, 1, 0) & "," & nDoc & ")"
                cn.Execute sql, rdExecDirect
                .MoveNext
            Loop
        End With
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"

End Sub

Private Sub Baixa2Eicon()
Dim nCodReduz As Long, sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, nSeq As Integer, sProcesso As String
Dim bCancelado As Boolean, nAno As Integer, nNumero As Long, nDoc As Long
sql = "delete from parcelamento2eicon"
cn.Execute sql, rdExecDirect

sql = "SELECT DISTINCT p.Numprocesso FROM Origemreparc p INNER JOIN Processoreparc m ON p.Numprocesso = m.Numprocesso WHERE p.Codreduzido >= 100000 AND "
sql = sql & "  p.Codreduzido < 200000 AND p.Anoexercicio > 2017  AND p.Codlancamento = 5"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        sProcesso = !NumProcesso
        
        sql = "SELECT * FROM processoreparc WHERE numprocesso='" & sProcesso & "'"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        nCodigo = RdoAux2!CODIGORESP
        bCancelado = RdoAux2!Cancelado
        sProcesso = RdoAux2!NumProcesso
        nAno = RdoAux2!AnoProc
        nNumero = RdoAux2!NumProc
        RdoAux2.Close
        
        sql = "SELECT distinct numdocumento  FROM origemreparc INNER JOIN parceladocumento ON origemreparc.codreduzido = parceladocumento.codreduzido AND origemreparc.anoexercicio = parceladocumento.anoexercicio AND origemreparc.codlancamento = parceladocumento.codlancamento AND numsequencia=parceladocumento.seqlancamento AND origemreparc.numparcela = parceladocumento.numparcela AND origemreparc.codcomplemento = parceladocumento.codcomplemento "
        sql = sql & "WHERE numprocesso='" & sProcesso & "' AND numdocumento BETWEEN 2000000 AND 2500000 ORDER BY numdocumento"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                nDoc = RdoAux2!NumDocumento
                sql = "insert parcelamento2eicon(codigo,ano,numero,cancelado,documento) values(" & nCodigo & "," & nAno & "," & nNumero & "," & IIf(bCancelado, 1, 0) & "," & nDoc & ")"
                cn.Execute sql, rdExecDirect
                .MoveNext
            Loop
        End With
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"

End Sub

Private Sub Corrige_Tramite()
Dim nAno As Integer, nNumero As Integer, nCCusto As Integer, nSeq As Integer, sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos As Long, nTot As Long
Dim aTramite() As tTramite, aTramite2() As tTramite, nCcusto2 As Integer, nSeq2 As Integer, x As Integer, bDiferente As Boolean, bFind As Boolean, y As Integer

sql = "delete from codtmp2"
'cn.Execute sql, rdExecDirect

sql = "select ano,numero from processogti where ano=2018 order by ano,numero"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nAno = !ano
        nNumero = !Numero
        ReDim aTramite(0): ReDim aTramite2(0)
        
        sql = "select ano,numero,seq,ccusto from tramitacao where ano=" & nAno & " and numero=" & nNumero & " order by seq"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                ReDim Preserve aTramite(UBound(aTramite) + 1)
                aTramite(UBound(aTramite)).nAno = nAno
                aTramite(UBound(aTramite)).nNumero = nNumero
                aTramite(UBound(aTramite)).nSeq = !Seq
                aTramite(UBound(aTramite)).nCCusto = !ccusto
               .MoveNext
            Loop
           .Close
        End With

        sql = "select ano,numero,seq,ccusto from tramitacaocc where ano=" & nAno & " and numero=" & nNumero & " order by seq"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                ReDim Preserve aTramite2(UBound(aTramite2) + 1)
                aTramite2(UBound(aTramite2)).nAno = nAno
                aTramite2(UBound(aTramite2)).nNumero = nNumero
                aTramite2(UBound(aTramite2)).nSeq = !Seq
                aTramite2(UBound(aTramite2)).nCCusto = !ccusto
               .MoveNext
            Loop
           .Close
        End With

       'busca as diferenças
        bDiferente = False
        If UBound(aTramite) <> UBound(aTramite2) Then
            bDiferente = True
        Else
            For x = 1 To UBound(aTramite2)
                nSeq = aTramite2(x).nSeq
                nCCusto = aTramite2(x).nCCusto
                If aTramite(x).nSeq <> nSeq Or aTramite(x).nCCusto <> nCCusto Then
                    bDiferente = True
                    Exit For
                End If
            Next
            If bDiferente Then
                sql = "insert codtmp2 (ano,numero) values(" & nAno & "," & nNumero & ")"
                cn.Execute sql, rdExecDirect
            End If
        End If

        


Proximo:
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "Fim"




End Sub

Private Sub RelPagamentoISS()
Dim sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos As Long, nTot As Long
Dim nNumDoc As Long, nValorGTI As Double, nValorGiss As Double, nValorTotalGiss As Double, sDataPagto As String, nCodReduz As Long, nValorTotal As Double

sql = "delete from codtmp3"
'cn.Execute sql, rdExecDirect
ConectaEicon
sql = "SELECT distinct tb_inter_baixa_detalhe.num_documento,valor_titulo,data_pagamento,num_cadastro FROM tb_inter_baixa_detalhe INNER JOIN tb_inter_boletos_giss on "
sql = sql & "tb_inter_baixa_detalhe.num_documento = tb_inter_boletos_giss.num_documento WHERE tb_inter_baixa_detalhe.NUM_DOCUMENTO BETWEEN 2276352 AND 2700000 and valor_pago>0 order by tb_inter_baixa_detalhe.num_documento"
'sql = "select numdocumento,valorpagoreal from debitopago where anoexercicio=2024 and codlancamento=5 "
Set rdoAux = cnEicon.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nValorGiss = 0: nValorTotalGiss = 0
        nNumDoc = !num_documento
        nValorGiss = Ponto2Virg(CStr(!VALOR_TITULO))
        sDataPagto = Format(!Data_Pagamento, "dd/mm/yyyy")
        nCodReduz = !num_cadastro
        
        
        sql = "select sum(valor_pago) as soma FROM tb_inter_baixa_detalhe WHERE num_documento=" & nNumDoc
        Set RdoAux2 = cnEicon.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        nValorTotal = Ponto2Virg(CStr(RdoAux2!soma))
        RdoAux2.Close
        If nValorTotal <> nValorGiss Then
            On Error Resume Next
            sql = "insert codtmp3 (documento,codigo,valorpago,datapagto,valortotal) values(" & nNumDoc & "," & nCodReduz & "," & Virg2Ponto(CStr(nValorGiss)) & ",'" & Format(sDataPagto, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(nValorTotal)) & ")"
            cn.Execute sql, rdExecDirect
        
        End If

        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

cnEicon.Close
MsgBox "Fim"

End Sub

Private Sub PagamentoCdas()
Dim sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos As Long, nTot As Long, nNumDoc As Long, nStatus As Integer

ConectaEicon
sql = "SELECT * FROM vwcdastatus ORDER BY num_documento"
Set rdoAux = cnEicon.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nNumDoc = !num_documento
        nStatus = !statuslanc
        
        If nStatus = 2 Then
            sql = "UPDATE tb_inter_boletos_cdas SET status='Q' WHERE num_documento=" & nNumDoc
            cnEicon.Execute sql, rdExecDirect
        ElseIf nStatus = 5 Or nStatus = 37 Then
            sql = "UPDATE tb_inter_boletos_cdas SET status='D' WHERE num_documento=" & nNumDoc
            cnEicon.Execute sql, rdExecDirect
        End If
        
'        If nValorTotal <> nValorGiss Then
'            sql = "insert codtmp3 (documento,codigo,valorpago,datapagto,valortotal) values(" & nNumDoc & "," & nCodReduz & "," & Virg2Ponto(CStr(nValorGiss)) & ",'" & Format(sDataPagto, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(nValorTotal)) & ")"
 '           cn.Execute sql, rdExecDirect
 '
 '       End If

        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

cnEicon.Close
MsgBox "Fim"

End Sub

Private Sub BaseEgati()
Dim sql As String, rdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos As Long, nTot As Long, x As Integer
Dim nSetor As Integer, nQuadra As Integer, nLote As Integer, nFace As Integer, nUnidade As Integer, nSubUnidade As Integer, nCodigo As Long, sQuadra As String, sLote As String
Dim nCodLogr As Integer, sTipoLogr As String, sNomeLogr As String, nNumImovel As Integer, sCompl As String, nCodBairro As Integer, sDescBairro As String, sCep As String, sUso As String, sTopografia As String
Dim sPedologia As String, bIsentoArea As Boolean, sIsencao As String, nTestada1 As Double, nTestadaN As Double, nAreaTerreno As Double, nAreaPredial As Double, tBairro As Bairro, nTipoEnd As Integer, sTipoEnd As String
Dim nArea1 As Double, nArea2 As Double, nArea3 As Double, nArea4 As Double, nArea5 As Double, sTipo1 As String, sTipo2 As String, sTipo3 As String, sTipo4 As String, sTipo5 As String
Dim sCateg1 As String, sCateg2 As String, sCateg3 As String, sCateg4 As String, sCateg5 As String, nCodCidadao As Long, sNome As String, sCPFCNPJ As String, sRG As String, sFone As String
Dim nEndereco_CodigoEE As Integer, sEndereco_NomeEE As String, nEndereco_NumEE As Integer, sEndreco_ComplEE As String, nBairro_CodigoEE As Integer, sBairro_NomeEE As String, nCidade_CodigoEE As Integer, sCidade_NomeEE As String
Dim sUFEE As String, sCepEE As String, nVVT As Double, nVVP As Double, nVVI As Double, nMatricula As Long, nCondominio As Integer, sCondominio As String

sql = "delete from exporta_imovel"
cn.Execute sql, rdExecDirect

sql = "SELECT  * FROM vwFULLIMOVEL2 WHERE ativo='S' ORDER BY codreduzido"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nTot = .RowCount
    nPos = 1
    
    Do Until .EOF
        If nPos Mod 50 = 0 Then
           CallPb nPos, nTot
        End If
        
        If !CODREDUZIDO = 7 Then
            'MsgBox "teste"
        End If
        
        nSetor = !Setor
        nQuadra = !Quadra
        nLote = !Lote
        nFace = !Seq
        nUnidade = !Unidade
        nSubUnidade = !SubUnidade
        nCodigo = !CODREDUZIDO
        sQuadra = Left(SubNull(!Li_Quadras), 15)
        sLote = Left(SubNull(!Li_Lotes), 15)
        nCodLogr = !CodLogr
        sTipoLogr = Trim(Left(SubNull(!AbrevTipoLog), 15))
        If SubNull(!AbrevTitLog) = "" Then
            sNomeLogr = !Logradouro3
        Else
            sNomeLogr = Trim(!AbrevTitLog) & " " & !Logradouro3
        End If
        nNumImovel = Val(SubNull(!Li_Num))
        sCompl = Left(SubNull(!Complemento), 50)
        tBairro = RetornaLogradouroBairro(nCodLogr, !Li_Num)
        nCodBairro = tBairro.Codigo
        sDescBairro = tBairro.Nome
        sCep = RetornaCEP(CLng(nCodLogr), !Li_Num)
        sUso = !DescUsoTerreno
        sTopografia = !DescTopografia
        sPedologia = !DescPedologia
        nAreaTerreno = !Dt_AreaTerreno
        nCodCidadao = !CodCidadao
        sNome = !nomecidadao
        sRG = SubNull(!rg)
        sCPFCNPJ = IIf(SubNull(!Cnpj) = "", SubNull(!cpf), SubNull(!Cnpj))
        sFone = SubNull(!telefone)
        nTipoEnd = !Ee_TipoEnd
        sIsencao = IIf(!Imune = True, "Imunidade", "")
        nMatricula = Val(SubNull(!NumMat))
        sCondominio = ""
        nCondominio = !CodCondominio
        If nCondominio <> 999 And nCondominio <> 0 Then
            sCondominio = Mask(!cd_nomecond)
        End If
                        
'        If sIsencao = "" Then
'            bIsentoArea = False
'            sql = "select codreduzido from vwisentoarea where codreduzido=" & !CODREDUZIDO
'            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
'            If RdoAux2.RowCount > 0 Then
'                bIsentoArea = True
'                sIsencao = "Isento Área"
'            End If
'            RdoAux2.Close
'        End If
        
 '       sql = "SELECT * FROM laseriptu WHERE ano=2025 AND codreduzido=" & !CODREDUZIDO
 '       Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
 '       With RdoAux2
 '           If RdoAux2.RowCount > 0 Then
 '               bIsentoArea = False
 '               sIsencao = ""
 '               nVVT = !vvt
 '               nVVP = !vvc
 '               nVVI = !vvi
 '           End If
 '          .Close
 '       End With
        
'        nTestada1 = 0: nTestadaN = 0
'        sql = "select * from testada where codreduzido=" & !CODREDUZIDO & " order by numface"
'        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
'        With RdoAux2
'            Do Until .EOF
'                If !NUMFACE = nFace Then
'                    nTestada1 = !AREATESTADA
 '               Else
 '                   nTestadaN = nTestadaN + !AREATESTADA
 '               End If
 ''              .MoveNext
  '          Loop
  '         .Close
  '      End With
        
        nAreaPredial = 0: x = 1
        nArea1 = 0: nArea2 = 0: nArea3 = 0: nArea4 = 0: nArea5 = 0
        sTipo1 = "": sTipo2 = "": sTipo3 = "": sTipo4 = "": sTipo5 = ""
        sCateg1 = "": sCateg2 = "": sCateg3 = "": sCateg4 = "": sCateg5 = ""
   '     sql = "select * from vwareas where codreduzido=" & !CODREDUZIDO
   '     Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
   '     With RdoAux2
   '         Do Until .EOF
   '             nAreaPredial = nAreaPredial + !AREACONSTR
   '             If x = 1 Then
   '                 nArea1 = !AREACONSTR
   '                 sTipo1 = !DESCTIPOCONSTR
   '                 sCateg1 = !desccategconstr
    ''            ElseIf x = 2 Then
    '                nArea2 = !AREACONSTR
    '                sTipo2 = !DESCTIPOCONSTR
    '                sCateg2 = !desccategconstr
    '            ElseIf x = 3 Then
    '                nArea3 = !AREACONSTR
    '                sTipo3 = !DESCTIPOCONSTR
     '               sCateg3 = !desccategconstr
     '           ElseIf x = 4 Then
     '               nArea4 = !AREACONSTR
     '               sTipo4 = !DESCTIPOCONSTR
     '               sCateg4 = !desccategconstr
     '           ElseIf x = 5 Then
     '               nArea5 = !AREACONSTR
     '               sTipo5 = !DESCTIPOCONSTR
     '               sCateg5 = !desccategconstr
     '           End If
      '          x = x + 1
      '         .MoveNext
      '      Loop
      '     .Close
       ' End With
        
        If nTipoEnd = 0 Then 'endereco imovel
            nEndereco_CodigoEE = nCodLogr
            sEndereco_NomeEE = sTipoLogr & " "
'            If SubNull(!AbrevTitLog) <> "" Then
'                sEndereco_NomeEE = sEndereco_NomeEE & Trim(!AbrevTitLog) & " "
 '           End If
            sEndereco_NomeEE = sEndereco_NomeEE & sNomeLogr
            nEndereco_NumEE = nNumImovel
            sEndreco_ComplEE = sCompl
            sBairro_NomeEE = sDescBairro
            sCidade_NomeEE = "JABOTICABAL"
            sUFEE = "SP"
            sCepEE = RetornaNumero(sCep)
        ElseIf nTipoEnd = 1 Then 'endereco prop
            sql = "SELECT CODCIDADAO,CODBAIRRO,CODBAIRRO2 FROM CIDADAO WHERE CODCIDADAO=" & nCodigo
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If Val(SubNull(RdoAux2!CodBairro)) > 0 Then
               sTipoEnd = "R"
            Else
               If Val(SubNull(RdoAux2!CodBairro2)) > 0 Then
                  sTipoEnd = "C"
               Else
                  sTipoEnd = "R"
               End If
            End If
            RdoAux2.Close
            
            If sTipoEnd = "R" Then
                sql = "SELECT CODCIDADAO,NOMECIDADAO,CPF,CNPJ,CODLOGRADOURO AS fCODLOGRADOURO,NUMIMOVEL AS fNUMIMOVEL,"
                sql = sql & "COMPLEMENTO AS fCOMPLEMENTO,CODBAIRRO AS fCODBAIRRO,CODCIDADE AS fCODCIDADE,SIGLAUF AS fSIGLAUF,"
                sql = sql & "CEP AS fCEP,TELEFONE AS fTELEFONE,EMAIL AS fEMAIL,RG AS fRG,NOMELOGRADOURO AS fNOMELOGRADOURO,ORGAO AS fORGAO"
                sql = sql & " FROM CIDADAO WHERE CODCIDADAO=" & nCodCidadao
            Else
                sql = "SELECT CODCIDADAO,NOMECIDADAO,CPF,CNPJ,CODLOGRADOURO2 AS fCODLOGRADOURO,NUMIMOVEL2 AS fNUMIMOVEL,"
                sql = sql & "COMPLEMENTO2 AS fCOMPLEMENTO,CODBAIRRO2 AS fCODBAIRRO,CODCIDADE2 AS fCODCIDADE,SIGLAUF2 AS fSIGLAUF,"
                sql = sql & "CEP2 AS fCEP,TELEFONE2 AS fTELEFONE,EMAIL2 AS fEMAIL,RG AS fRG,NOMELOGRADOURO2 AS fNOMELOGRADOURO,ORGAO AS fORGAO"
                sql = sql & " FROM CIDADAO WHERE CODCIDADAO=" & nCodCidadao
            End If
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount > 0 Then
                nEndereco_CodigoEE = Val(SubNull(RdoAux2!FCodLogradouro))
                sEndereco_NomeEE = Trim(SubNull(RdoAux2!FNomeLogradouro))
                nEndereco_NumEE = RdoAux2!fNUMIMOVEL
                sEndreco_ComplEE = Trim(SubNull(RdoAux2!fcomplemento))
                nBairro_CodigoEE = Val(SubNull(RdoAux2!fCodBairro))
                nCidade_CodigoEE = Val(SubNull(RdoAux2!fCodCidade))
                sUFEE = Trim(SubNull(RdoAux2!fsiglauf))
                sCepEE = Val(RetornaNumero(SubNull(RdoAux2!FCEP)))
            End If
            RdoAux2.Close
        ElseIf nTipoEnd = 2 Then 'endereco entrega
            sql = "select DISTINCT e.codreduzido, e.ee_codlog,v.endereco_resumido ,ee_numimovel,ee_complemento,ee_nomelog,e.ee_bairro,b.descbairro,e.ee_cidade,c.desccidade,ee_uf,ee_cep "
            sql = sql & "FROM endentrega e LEFT outer JOIN vwLOGRADOURO v ON ee_codlog = v.codlogradouro LEFT OUTER JOIN cidade c ON ee_uf=c.siglauf AND e.ee_cidade=c.codcidade "
            sql = sql & "LEFT OUTER JOIN bairro b ON e.ee_uf = b.siglauf AND e.ee_cidade = b.codcidade AND e.ee_bairro = codbairro where e.codreduzido=" & nCodigo
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                nEndereco_CodigoEE = Val(SubNull(RdoAux2!Ee_CodLog))
                sEndereco_NomeEE = Trim(SubNull(RdoAux2!endereco_resumido))
                nEndereco_NumEE = RdoAux2!Ee_NumImovel
                sEndreco_ComplEE = Trim(SubNull(RdoAux2!Ee_Complemento))
                nBairro_CodigoEE = Val(SubNull(RdoAux2!Ee_Bairro))
                nCidade_CodigoEE = Val(SubNull(RdoAux2!Ee_Cidade))
                sUFEE = Trim(SubNull(RdoAux2!Ee_Uf))
                sBairro_NomeEE = Trim(SubNull(RdoAux2!DescBairro))
                sCidade_NomeEE = Trim(SubNull(RdoAux2!descCidade))
                sCepEE = Val(RetornaNumero(SubNull(RdoAux2!Ee_Cep)))
               .Close
            End With
        End If
        
        If nTipoEnd > 0 Then
            If nEndereco_CodigoEE > 0 Then
                sql = "SELECT endereco_resumido FROM vwLOGRADOURO WHERE codlogradouro=" & nEndereco_CodigoEE
                Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                If RdoAux2.RowCount > 0 Then
                    sEndereco_NomeEE = SubNull(RdoAux2!endereco_resumido)
                End If
                RdoAux2.Close
            End If
            If nCidade_CodigoEE > 0 And sCidade_NomeEE = "" Then
                sql = "SELECT desccidade FROM cidade WHERE siglauf='" & sUFEE & "' and codcidade=" & nCidade_CodigoEE
                Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                If RdoAux2.RowCount > 0 Then
                    sCidade_NomeEE = SubNull(RdoAux2!descCidade)
                End If
                RdoAux2.Close
            End If
            If nBairro_CodigoEE > 0 And sBairro_NomeEE = "" Then
                sql = "SELECT descbairro FROM bairro WHERE siglauf='" & sUFEE & "' and codcidade=" & nCidade_CodigoEE & " and codbairro=" & nBairro_CodigoEE
                Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                If RdoAux2.RowCount > 0 Then
                    sBairro_NomeEE = SubNull(RdoAux2!DescBairro)
                End If
                RdoAux2.Close
            End If
            
        End If
        If Val(sCepEE) = 0 Then sCepEE = "14870000"
        If sCidade_NomeEE = "" Then sCidade_NomeEE = "JABOTICABAL"
        sql = "insert exporta_imovel (seq,setor,quadra,lote,face,unidade,subunidade,codigo_imovel,quadra_rf,lote_rf,logradouro_codigo,logradouro_tipo,logradouro_nome,logradouro_numero,logradouro_complemento,"
        sql = sql & "bairro_codigo,bairro_nome,cep,uso_imovel,topografia,pedologia,isencao,testadaprinc,testadasecun,area_terreno,area_predial,area1m2,area1tipo,area1categ,area2m2,area2tipo,area2categ,area3m2,"
        sql = sql & "area3tipo,area3categ,area4m2,area4tipo,area4categ,area5m2,area5tipo,area5categ,proprietario_nome,proprietario_cpfcnpj,proprietario_rg,proprietario_telefone,corresp_ruacodigo,corresp_ruanome,"
        sql = sql & "corresp_numero,corresp_complemento ,corresp_bairro_codigo,corresp_bairro_nome,corresp_cidade_codigo,corresp_cidade_nome,corresp_uf,corresp_cep,vlvenal_terreno,vlvenal_predial,vlvenal_imovel,matricula,"
        sql = sql & "condominio_codigo,condominio_nome) "
        sql = sql & "values(" & nPos & "," & nSetor & "," & nQuadra & "," & nLote & "," & nFace & "," & nUnidade & "," & nSubUnidade & "," & nCodigo & ",'" & Mask(sQuadra) & "','" & Mask(sLote) & "'," & nCodLogr & ",'"
        sql = sql & sTipoLogr & "','" & Mask(sNomeLogr) & "'," & nNumImovel & ",'" & sCompl & "'," & nCodLogr & ",'" & sDescBairro & "','" & sCep & "','" & sUso & "','" & sTopografia & "','" & sPedologia & "','"
        sql = sql & sIsencao & "'," & Virg2Ponto(CStr(nTestada1)) & "," & Virg2Ponto(CStr(nTestadaN)) & "," & Virg2Ponto(CStr(nAreaTerreno)) & "," & Virg2Ponto(CStr(nAreaPredial)) & ","
        sql = sql & Virg2Ponto(CStr(nArea1)) & ",'" & sTipo1 & "','" & sCateg1 & "'," & Virg2Ponto(CStr(nArea2)) & ",'" & sTipo2 & "','" & sCateg2 & "'," & Virg2Ponto(CStr(nArea3)) & ",'" & sTipo3 & "','"
        sql = sql & sCateg3 & "'," & Virg2Ponto(CStr(nArea4)) & ",'" & sTipo4 & "','" & sCateg4 & "'," & Virg2Ponto(CStr(nArea5)) & ",'" & sTipo5 & "','" & sCateg5 & "','" & Mask(sNome) & "','" & sCPFCNPJ & "','"
        sql = sql & sRG & "','" & sFone & "'," & nEndereco_CodigoEE & ",'" & Mask(sEndereco_NomeEE) & "'," & nEndereco_NumEE & ",'" & sEndreco_ComplEE & "'," & nBairro_CodigoEE & ",'" & Mask(sBairro_NomeEE) & "',"
        sql = sql & nCidade_CodigoEE & ",'" & Mask(sCidade_NomeEE) & "','" & sUFEE & "','" & sCepEE & "'," & Virg2Ponto(CStr(nVVT)) & "," & Virg2Ponto(CStr(nVVP)) & "," & Virg2Ponto(CStr(nVVI)) & "," & nMatricula & ","
        sql = sql & nCondominio & ",'" & Mask(sCondominio) & "')"
        cn.Execute sql, rdExecDirect

        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"

End Sub


Private Sub ListaEmpresasVereador()
Dim sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, sReg As String, FF1 As Integer, sCPFCNPJ As String, sSocio As String


FF1 = FreeFile()
Open "c:\work\empresas.txt" For Output As FF1


sql = "SELECT codigomob,razaosocial,cpf,cnpj,logradouro,numero,complemento,descbairro AS bairro,ativextenso AS atividade,telefone_nf AS telefone,email_nf AS email FROM vwFULLEMPRESA2 WHERE dataencerramento IS NULL AND codatividade>10000 AND codatividade<20000 ORDER BY razaosocial"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nPos = 1
    nTot = .RowCount
    Do Until .EOF
        If nPos Mod 30 = 0 Then
           CallPb nPos, nTot
        End If
        sSocio = ""
        sql = "SELECT c.nomecidadao AS nome FROM mobiliarioproprietario p inner JOIN cidadao c ON p.codcidadao = c.codcidadao  WHERE codmobiliario=" & !codigomob
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                sSocio = sSocio & !Nome & ","
               .MoveNext
            Loop
           .Close
        End With
        If sSocio <> "" Then
            sSocio = Left(sSocio, Len(sSocio) - 1)
        End If
        sCPFCNPJ = IIf(SubNull(!Cnpj) = "", SubNull(!cpf), SubNull(!Cnpj))
        sReg = !RazaoSocial & "#" & sCPFCNPJ & "#" & !Logradouro & "#" & !Numero & "#" & SubNull(!Complemento) & "#" & SubNull(!Bairro) & "#" & !Atividade & "#" & SubNull(!telefone) & "#" & !Email & "#" & sSocio
        Print #FF1, sReg
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

Close #FF1




MsgBox "fim"

End Sub

Private Sub Corrige_Complemento()
Dim sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, sReg As String, FF1 As Integer, sCPFCNPJ As String, sSocio As String


sql = "SELECT ano,numero,complemento from codtmp2 order by ano,numero"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nPos = 1
    nTot = .RowCount
    Do Until .EOF
        If nPos Mod 6 = 0 Then
           CallPb nPos, nTot
        End If
        sql = "update processogti set complemento='" & Mask(!Complemento) & "' where ano=" & !ano & " and numero=" & !Numero
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

MsgBox "fim"

End Sub

Private Sub Atualiza_Protesto()
Dim sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, sReg As String, FF1 As Integer, sCPFCNPJ As String, sSocio As String


sql = "SELECT guid FROM protesto_importar where dataremessa >'10/01/2025'"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nPos = 1
    nTot = .RowCount
    Do Until .EOF
        If nPos Mod 6 = 0 Then
           CallPb nPos, nTot
        End If
        sql = "update protesto_parcela SET atualizou_GTI=0 WHERE guid='" & !guid & "'"
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

MsgBox "fim"

End Sub

Private Function ExtrairTelefones()
    Dim rdoAux As rdoResultset, sql As String, nPos As Long, nTot As Long
    Dim partes() As String, telefoneBruto As String, telefoneLimpo As String
    Dim telefone As String, Codigo As String, DDD As String, Numero As String
    Dim i As Long

    ' Limpa tabela de destino
    sql = "DELETE FROM codtmp5"
    cn.Execute sql, rdExecDirect

    ' Seleciona registros com telefone
    sql = "SELECT codigomob AS codigo, fonecontato FROM mobiliário WHERE DATAENCERRAMENTO IS NULL AND fonecontato IS NOT NULL ORDER BY CODIGOMOB "
    Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)

    nPos = 1
    nTot = rdoAux.RowCount

    Do While Not rdoAux.EOF
        If nPos Mod 10 = 0 Then CallPb nPos, nTot
        'If nPos > 30 Then Exit Do
        Codigo = rdoAux!Codigo
        telefoneBruto = Trim(rdoAux!fonecontato)

        ' Divide por vírgula, barra ou ponto e vírgula
        partes = Split(Replace(Replace(Replace(telefoneBruto, "/", ","), ";", ","), "\", ","), ",")

        For i = 0 To UBound(partes)
            telefone = Trim(partes(i))
            telefoneLimpo = LimparTelefone(telefone)

            If Len(telefoneLimpo) >= 8 Then
                Call SepararDDDNumero(telefoneLimpo, DDD, Numero)

                sql = "INSERT INTO codtmp5(codigo, bruto, ddd, fone) VALUES(" & _
                      Val(Codigo) & ", '" & telefoneBruto & "', '" & DDD & "', '" & Numero & "')"
                cn.Execute sql, rdExecDirect
            End If
        Next i

        nPos = nPos + 1
        rdoAux.MoveNext
    Loop

    rdoAux.Close
    MsgBox "fim"
End Function
Function LimparTelefone(ByVal entrada As String) As String
    entrada = Replace(entrada, "(", "")
    entrada = Replace(entrada, ")", "")
    entrada = Replace(entrada, "-", "")
    entrada = Replace(entrada, ".", "")
    entrada = Replace(entrada, " ", "")
     
     ' Remove letras e mantém apenas números
    For i = 1 To Len(entrada)
        If Mid(entrada, i, 1) Like "#" Then
            resultado = resultado & Mid(entrada, i, 1)
        End If
    Next i

 ' Regras de formatação
    If Len(resultado) = 8 And Left(resultado, 1) <> "3" Then
        resultado = "16" & "9" & resultado
    ElseIf Len(resultado) = 9 Then
        resultado = "16" & resultado
    ElseIf Len(resultado) = 10 Or Len(resultado) = 11 Then
        ' já tem DDD
    ElseIf Len(resultado) < 8 Then
        resultado = ""
    End If

    LimparTelefone = resultado

End Function

Sub SepararDDDNumero(ByVal telefone As String, ByRef DDD As String, ByRef Numero As String)
    Select Case Len(telefone)
        Case 11 ' DDD + celular
            DDD = Left(telefone, 2)
            Numero = Mid(telefone, 3)
        Case 10 ' DDD + fixo
            DDD = Left(telefone, 2)
            Numero = Mid(telefone, 3)
        Case 9 ' celular sem DDD
            DDD = "16"
            Numero = telefone
        Case 8 ' sem DDD
            DDD = "16"
            If Left(telefone, 1) = "3" Then
                Numero = telefone ' fixo, não adiciona 9
            Else
                Numero = "9" & telefone ' celular, adiciona 9
            End If
        Case Else
            DDD = "16"
            Numero = telefone
    End Select
End Sub

Private Sub Corrige_TelefoneOld()
Dim sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, sReg As String, FF1 As Integer, sCPFCNPJ As String, sSocio As String
Dim nCodNew As Long, nCodOld As Long, nSeq As Integer
nCodOld = 0
sql = "SELECT id,codigo,bruto,ddd,fone from codtmp5 order by id"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nPos = 1
    nTot = .RowCount
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodNew = !Codigo
        If nCodOld = nCodNew Then
            nSeq = nSeq + 1
        Else
            nCodOld = nCodNew
            nSeq = 1
        End If
        
        sql = "update codtmp5 set seq=" & nSeq & " where id=" & !id
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

MsgBox "fim"

End Sub

Private Sub Corrige_Telefone()
Dim sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, sReg As String, FF1 As Integer, sCPFCNPJ As String, sSocio As String
Dim nSeq As Integer, nCodigo As Long
sql = "SELECT id,codigo,seq,bruto,ddd,fone from codtmp5 order by id"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nPos = 1
    nTot = .RowCount
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodigo = !Codigo
        nSeq = !Seq
        If nSeq = 1 Then
            sql = "update mobiliario set ddd1=" & !DDD & ",fone1=" & !fone & " where codigomob=" & nCodigo
        ElseIf nSeq = 2 Then
            sql = "update mobiliario set ddd2=" & !DDD & ",fone2=" & !fone & " where codigomob=" & nCodigo
        ElseIf nSeq = 3 Then
            sql = "update mobiliario set ddd3=" & !DDD & ",fone3=" & !fone & " where codigomob=" & nCodigo
        End If
        
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

MsgBox "fim"

End Sub

Private Sub ResumoIPTU_GEO()

Dim sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, sReg As String, FF1 As Integer

FF1 = FreeFile()
Open "C:\Work\GTI\Diversos\IPTU2026\ResumoIPTU2026.txt" For Output As FF1

sql = "SELECT codreduzido,natureza,areaconstrucao,(valortotalparc*qtdeparc) AS soma from laseriptu where ano=2026  order by codreduzido"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nPos = 1
    nTot = .RowCount
    Do Until .EOF
        If nPos Mod 30 = 0 Then
           CallPb nPos, nTot
        End If
        
        sql = "SELECT codreduzido,natureza,areaconstrucao,(valortotalparc*qtdeparc) AS soma from laseriptu_nogeo where ano=2026 and codreduzido=" & !CODREDUZIDO & " order by codreduzido"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            sReg = !CODREDUZIDO & "#" & !Natureza & "#" & Format(!areaconstrucao, "#0.00") & "#" & Format(!soma, "#0.00")
            sReg = sReg & "#" & RdoAux2!Natureza & "#" & Format(RdoAux2!areaconstrucao, "#0.00") & "#" & Format(RdoAux2!soma, "#0.00") & "#" & " "
        Else
            sql = "select sum(areaconstr) as soma from areas where codreduzido=" & !CODREDUZIDO & " and areageo=0"
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount > 0 Then
                sReg = !CODREDUZIDO & "#" & !Natureza & "#" & Format(!areaconstrucao, "#0.00") & "#" & Format(!soma, "#0.00")
                sReg = sReg & "#" & IIf(RdoAux2!soma > 0, "Predial", "Territorial") & "#" & Format(RdoAux2!soma, "#0.00") & "#" & Format(0, "#0.00") & "#" & "Sem Anterior"
            Else
                sReg = !CODREDUZIDO & "#" & !Natureza & "#" & Format(!areaconstrucao, "#0.00") & "#" & Format(!soma, "#0.00")
                sReg = sReg & "#" & !Natureza & "#" & Format(!areaconstrucao, "#0.00") & "#" & Format(0, "#0.00") & "#" & "Sem Anterior"
            End If
        End If
        Print #FF1, sReg
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

Close #FF1

MsgBox "fim"

End Sub

Private Sub Incluir_AreaGeo()
Dim sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, nCodReduz As Long, nMaxSeq As Integer, nArea As Double
Dim nUso As Integer, nTipo As Integer, nCat As Integer

sql = "select * from areas_geo2 where remover =0 order by codigo"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    nPos = 1
    nTot = .RowCount
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodReduz = !Codigo
        nArea = !Area
        
        sql = "select * from areas where codreduzido=" & nCodReduz
        Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux3.RowCount > 0 Then
            sql = "select max(seqarea) as maximo from areas where codreduzido=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            nMaxSeq = RdoAux2!maximo + 1
            nUso = RdoAux3!USOCONSTR
            nTipo = RdoAux3!TIPOCONSTR
            nCat = RdoAux3!CATCONSTR
        Else
            nMaxSeq = 1
            nUso = 1
            nCat = 1
            nTipo = 1
        End If
        
        sql = "INSERT AREAS (CODREDUZIDO,SEQAREA,DATAAPROVA,AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR,QTDEPAV,AREAGEO) VALUES(" & nCodReduz & "," & nMaxSeq & ",'" & Format(Now, "mm/dd/yyyy") & "',"
        sql = sql & Virg2Ponto(RemovePonto(CStr(nArea))) & "," & nUso & "," & nTipo & "," & nCat & ",1,1)"
        cn.Execute sql, rdExecDirect

        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

MsgBox "fim"

End Sub

Private Sub CorrigeParcelamentoBloqueio()
Dim sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, nCodReduz As Long, nValorTributo As Double
Dim nSeq As Integer, nParc As Integer, nCompl As Integer
nPos = 1
 sql = "SELECT numprocesso,  numproc,anoproc,codigoresp FROM processoreparc WHERE year(datareparc)>2015 and YEAR(datareparc)<2025"
 Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)

 With rdoAux
     nTot = .RowCount
     Do Until .EOF
        If .AbsolutePosition Mod 50 = 0 Then
            CallPb nPos, nTot
         End If
         
         sql = "select * from debitoparcela where codreduzido=" & !CODIGORESP & " and codlancamento=20 and numprocesso='" & !NumProcesso & "' and statuslanc in (3,18)"
         Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
         If RdoAux2.RowCount = 0 Then
            GoTo Proximo
         End If
         nSeq = RdoAux2!SeqLancamento
         nParc = RdoAux2!NumParcela
         nCompl = RdoAux2!CODCOMPLEMENTO
         RdoAux2.Close
         
         nValorTributo = 0
         sql = "select * from debitotributo where codreduzido=" & !CODIGORESP & " and anoexercicio=2025 and codlancamento=20 and seqlancamento=" & nSeq & " and "
         sql = sql & "numparcela=" & nParc & " and codcomplemento=" & nCompl & " and codtributo=587"
         Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
         If RdoAux2.RowCount > 0 Then
            nValorTributo = RdoAux2!VALORTRIBUTO
         End If
         RdoAux2.Close
         
         If nValorTributo > 0 Then
            sql = "update debitotributo set valortributo=" & Virg2Ponto(CStr(nValorTributo)) & " where codreduzido=" & !CODIGORESP & " and anoexercicio>2025 and codlancamento=20 and seqlancamento=" & nSeq & " and codtributo=587"
            cn.Execute sql, rdExecDirect
         End If
         
Proximo:
         nPos = nPos + 1
        .MoveNext
     Loop
    .Close
 End With
 

MsgBox "fim"

End Sub

Private Sub CorrigeParcelamento2026()
Dim sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, nCodReduz As Long, nValorTributo As Double
Dim nSeq As Integer, nParc As Integer, nCompl As Integer
nPos = 1
sql = "SELECT DISTINCT codreduzido,seqlancamento FROM debitoparcela WHERE anoexercicio=2026 AND codlancamento=20 AND statuslanc IN (3,18)"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)

 With rdoAux
     nTot = .RowCount
     Do Until .EOF
        If .AbsolutePosition Mod 50 = 0 Then
            CallPb nPos, nTot
         End If
         nCodReduz = !CODREDUZIDO
         nSeq = !SeqLancamento
         
         'If nCodReduz <> 34692 Then
         '   GoTo Proximo
         'End If
         
         nValorTributo = 0
         sql = "select * from debitotributo where codreduzido=" & nCodReduz & " and anoexercicio=2025 and codlancamento=20 and seqlancamento=" & nSeq & " and codtributo=587"
         Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
         If RdoAux2.RowCount > 0 Then
            nValorTributo = RdoAux2!VALORTRIBUTO
         End If
         RdoAux2.Close
         
         If nValorTributo > 0 Then
            sql = "update debitotributo set valortributo=" & Virg2Ponto(CStr(nValorTributo)) & " where codreduzido=" & nCodReduz & " and anoexercicio>2025 and codlancamento=20 and seqlancamento=" & nSeq & " and codtributo=587"
            cn.Execute sql, rdExecDirect
         End If
         
Proximo:
         nPos = nPos + 1
        .MoveNext
     Loop
    .Close
 End With
 

MsgBox "fim"

End Sub

Private Sub CorrigeEnviadoProtestoPago()
Dim sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, nCodReduz As Long, nValorTributo As Double
Dim nSeq As Integer, nParc As Integer, nCompl As Integer, nLanc As Integer, nAno As Integer
nPos = 1
sql = "SELECT * FROM debitoparcela WHERE statuslanc=39 AND YEAR(protesto_data_remessa)=2025 ORDER BY codreduzido"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)

 With rdoAux
     nTot = .RowCount
     Do Until .EOF
        If .AbsolutePosition Mod 50 = 0 Then
            CallPb nPos, nTot
         End If
         nCodReduz = !CODREDUZIDO
         nAno = !AnoExercicio
         nLanc = !CodLancamento
         nSeq = !SeqLancamento
         nParc = !NumParcela
         nCompl = !CODCOMPLEMENTO
         
         sql = "select * from debitopago where codreduzido=" & nCodReduz & " and anoexercicio=" & nAno & " and codlancamento=" & nLanc & " and seqlancamento=" & nSeq
         sql = sql & " and numparcela=" & nParc & " and codcomplemento=" & nCompl
         Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
         If RdoAux2.RowCount > 0 Then
            sql = "update debitoparcela set statuslanc=2 where codreduzido=" & nCodReduz & " and anoexercicio=" & nAno & " and codlancamento=" & nLanc & " and seqlancamento=" & nSeq
            sql = sql & " and numparcela=" & nParc & " and codcomplemento=" & nCompl
            cn.Execute sql, rdExecDirect
         End If
         RdoAux2.Close
         
         
Proximo:
         nPos = nPos + 1
        .MoveNext
     Loop
    .Close
 End With
 

MsgBox "fim"

End Sub

Private Sub VerificaAreaGeoCalculada()
Dim sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, nCodReduz As Long, nArea As Double
Dim bCalcular As Boolean, nSeq As Integer, RdoAux4 As rdoResultset

nPos = 1
sql = "SELECT codreduzido,seqarea FROM areas WHERE areageo=1"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)

 With rdoAux
     nTot = .RowCount
     Do Until .EOF
        If .AbsolutePosition Mod 50 = 0 Then
           CallPb nPos, nTot
        End If
        nCodReduz = !CODREDUZIDO
        nSeq = !SEQAREA
            
        sql = "SELECT codreduzido,imune FROM cadimob WHERE codreduzido=" & nCodReduz
        Set RdoAux4 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux4!Imune = 1 Then
            GoTo Proximo
        End If
            
        bCalcular = False
        sql = "SELECT codreduzido,areaconstrucao FROM laseriptu WHERE ano=2026 AND codreduzido=" & nCodReduz
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount = 0 Then
            GoTo Proximo
        Else
            nArea = RdoAux2!areaconstrucao
            sql = "select sum(areaconstr) as soma from areas where codreduzido=" & nCodReduz
            Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If RdoAux3!soma <> nArea Then
                bCalcular = True
            End If
        End If
         
        If bCalcular Then
            sql = "update areas set calcular=1 where codreduzido=" & nCodReduz & " and seqarea=" & nSeq
            cn.Execute sql, rdExecDirect
        End If
        
         
Proximo:
         nPos = nPos + 1
        .MoveNext
     Loop
    .Close
 End With
 

MsgBox "fim"

End Sub

Private Sub VerificaProtestoPago()
Dim sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, nCodReduz As Long, nValorTributo As Double, RdoAux4 As rdoResultset
Dim nSeq As Integer, nParc As Integer, nCompl As Integer, nLanc As Integer, nAno As Integer, guid As String, id As String, nAnoCertidao As Integer, nNumCertidao As Integer
nPos = 1
sql = "SELECT guid,id,cadastro,anocertidao,nrocertidao FROM protesto_importar ORDER BY cadastro,anocertidao"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)

 With rdoAux
     nTot = .RowCount
     Do Until .EOF
        If .AbsolutePosition Mod 50 = 0 Then
            CallPb nPos, nTot
         End If
         nCodReduz = !cadastro
         guid = !guid
         id = !id
         nAnoCertidao = !anocertidao
         nNumCertidao = !NroCertidao
         
         sql = "SELECT * FROM protesto_parcela WHERE guid='" & guid & "' and id='" & id & "'"
         Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
         With RdoAux2
            Do Until .EOF
                nAno = !exercicio
                nLanc = !lancamento
                nSeq = !Seq
                nParc = !nroparcela
                nCompl = !subparcela
               
                sql = "select * from debitopago where codreduzido=" & nCodReduz & " and anoexercicio=" & nAno & " and codlancamento=" & nLanc & " and seqlancamento=" & nSeq
                sql = sql & " and numparcela=" & nParc & " and codcomplemento=" & nCompl
                Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                If RdoAux3.RowCount > 0 Then
                    sql = "select codreduzido,statuslanc from debitoparcela where codreduzido=" & nCodReduz & " and anoexercicio=" & nAno & " and codlancamento=" & nLanc & " and seqlancamento=" & nSeq
                    sql = sql & " and numparcela=" & nParc & " and codcomplemento=" & nCompl
                    Set RdoAux4 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                    If RdoAux4!statuslanc <> 2 And RdoAux4!statuslanc <> 41 And RdoAux4!statuslanc <> 39 Then
                        sql = "insert codtmp6 (codigo,ano,numero) values(" & nCodReduz & "," & nAnoCertidao & "," & nNumCertidao & ")"
                    '    sql = "update debitoparcela set statuslanc=2 where codreduzido=" & nCodReduz & " and anoexercicio=" & nAno & " and codlancamento=" & nLanc & " and seqlancamento=" & nSeq
                     '   sql = sql & " and numparcela=" & nParc & " and codcomplemento=" & nCompl
                     On Error Resume Next
                        cn.Execute sql, rdExecDirect
                    End If
                    RdoAux4.Close
                End If
                RdoAux3.Close
               .MoveNext
            Loop
           .Close
           
         End With
Proximo:
         DoEvents
         nPos = nPos + 1
        .MoveNext
     Loop
    .Close
 End With
 

MsgBox "fim"

End Sub

Private Sub VerificaProtestoCancelado()
Dim sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, nCodReduz As Long, nValorTributo As Double, RdoAux4 As rdoResultset
Dim nSeq As Integer, nParc As Integer, nCompl As Integer, nLanc As Integer, nAno As Integer, guid As String, id As String, nAnoCertidao As Integer, nNumCertidao As Integer
nPos = 1
sql = "SELECT * from debitoparcela where statuslanc=39 order by codreduzido"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)

 With rdoAux
     nTot = .RowCount
     Do Until .EOF
        'If !CODREDUZIDO <> 123982 Then GoTo Proximo
        If .AbsolutePosition Mod 50 = 0 Then
           CallPb nPos, nTot
        End If
        nCodReduz = !CODREDUZIDO
        nAno = !AnoExercicio
        nLanc = !CodLancamento
        nSeq = !SeqLancamento
        nParc = !NumParcela
        nCompl = !CODCOMPLEMENTO
         
        sql = "SELECT * from debitocancel WHERE codreduzido=" & nCodReduz & " AND anoexercicio=" & nAno & " and codlancamento=" & nLanc & " and "
        sql = sql & "seqlancamento=" & nSeq & " and numparcela=" & nParc & " and codcomplemento=" & nCompl
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            sql = "update debitoparcela set statuslanc=8 where codreduzido=" & nCodReduz & " and anoexercicio=" & nAno & " and codlancamento=" & nLanc & " and seqlancamento=" & nSeq
            sql = sql & " and numparcela=" & nParc & " and codcomplemento=" & nCompl
            cn.Execute sql, rdExecDirect
        End If
         
Proximo:
        DoEvents
        nPos = nPos + 1
       .MoveNext
     Loop
    .Close
 End With
 

MsgBox "fim"

End Sub

Private Sub CorrigeMei2()
Dim sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, nCodReduz As Long, nValorTributo As Double, RdoAux4 As rdoResultset
Dim sDataInicio As String, sDataFim  As String, aMei() As tMei, x As Integer, id As Integer, y As Integer, z As Integer, bFind As Boolean

ReDim aMei(0)

nPos = 1
sql = "SELECT id, codigo,datainicio,datafim FROM periodomei INNER JOIN mobiliario ON codigomob = codigo Where dataencerramento Is Null ORDER BY codigo,datainicio desc"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)

 With rdoAux
     nTot = .RowCount
     Do Until .EOF

        If .AbsolutePosition Mod 50 = 0 Then
           CallPb nPos, nTot
        End If
        id = !id
        nCodReduz = !Codigo
        sDataInicio = Format(!datainicio, "dd/mm/yyyy")
        If IsNull(!datafim) Then
            sDataFim = ""
        Else
            sDataFim = Format(!datafim, "dd/mm/yyyy")
        End If
         
        ReDim Preserve aMei(UBound(aMei) + 1)
        x = UBound(aMei)
        aMei(x).id = id
        aMei(x).Codigo = nCodReduz
        aMei(x).datainicio = sDataInicio
        aMei(x).datafim = sDataFim
Proximo:
        DoEvents
        nPos = nPos + 1
       .MoveNext
     Loop
    .Close
End With
 
For x = 1 To UBound(aMei)
    For y = 1 To UBound(aMei)
        If aMei(x).Codigo = aMei(y).Codigo And aMei(x).datainicio = aMei(y).datainicio And aMei(x).datafim = "" Then
            For z = 1 To UBound(aMei)
                If aMei(z).datainicio = aMei(x).datainicio And aMei(z).datafim <> "" Then
                    MsgBox aMei(z).Codigo & " - " & aMei(z).id
                    Exit For
                End If
            Next
        End If
    Next
Next
 

MsgBox "fim"

End Sub


Private Sub VerificaTaxaMei()
Dim sql As String, rdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, nCodReduz As Long, nValorTributo As Double, RdoAux4 As rdoResultset
Dim nSeq As Integer, nParc As Integer, nCompl As Integer, nLanc As Integer, nAno As Integer, guid As String, id As String, nAnoCertidao As Integer, nNumCertidao As Integer
nPos = 1
sql = "SELECT codigomob FROM mobiliario WHERE isentotaxa=1"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)

 With rdoAux
     nTot = .RowCount
     Do Until .EOF
        If .AbsolutePosition Mod 50 = 0 Then
           CallPb nPos, nTot
        End If
        nCodReduz = !codigomob
         
        sql = "SELECT * from periodomei where codigo=" & nCodReduz
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            If Not IsMEI(nCodReduz) Then
                sql = "update mobiliario set isentotaxa=0 where codigomob=" & nCodReduz
                cn.Execute sql, rdExecDirect
            End If
        End If
Proximo:
        DoEvents
        nPos = nPos + 1
       .MoveNext
     Loop
    .Close
 End With
 

MsgBox "fim"

End Sub
Private Function IsMEI(nCodigo As Long) As Boolean
Dim nRet As Boolean, sql As String, rdoAux As rdoResultset
nRet = False

'ConectaEicon
'Sql = "SELECT * FROM tb_inter_empr_mei WHERE inscricao=" & nCodigo & " order by [ timestamp] desc"
sql = "select * from periodomei where codigo=" & nCodigo & " order by id desc"
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    If .RowCount > 0 Then
        If IsNull(!datafim) Then
            nRet = True
        End If
    End If
   .Close
End With

'cnEicon.Close
IsMEI = nRet

End Function
