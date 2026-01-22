VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmArquivoIntegrativa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Geração de Arquivos para o Sistema de Cobrança"
   ClientHeight    =   3060
   ClientLeft      =   12555
   ClientTop       =   5655
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3060
   ScaleWidth      =   5865
   Begin esMaskEdit.esMaskedEdit mskDataVencto 
      Height          =   285
      Left            =   3960
      TabIndex        =   3
      Top             =   1755
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      MouseIcon       =   "frmArquivoIntegrativa.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "99/99/9999"
      SelText         =   ""
      Text            =   "31/12/2022"
      HideSelection   =   -1  'True
   End
   Begin VB.CommandButton cmdIntegrar 
      Caption         =   "Integrar"
      Height          =   330
      Left            =   2070
      TabIndex        =   10
      Top             =   4140
      Width           =   825
   End
   Begin VB.CommandButton btfix 
      Caption         =   "Fix"
      Enabled         =   0   'False
      Height          =   330
      Left            =   210
      TabIndex        =   9
      Top             =   4095
      Width           =   825
   End
   Begin prjChameleon.chameleonButton cmdVerificar 
      Height          =   345
      Left            =   90
      TabIndex        =   7
      ToolTipText     =   "Verificar Base de Dados"
      Top             =   3195
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   609
      BTYPE           =   3
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16711935
      MPTR            =   1
      MICON           =   "frmArquivoIntegrativa.frx":001C
      PICN            =   "frmArquivoIntegrativa.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtArq 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   390
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "c:\CADASTRO\Bin\INTEGR01.txt"
      Top             =   2280
      Width           =   4035
   End
   Begin prjChameleon.chameleonButton cmdArq 
      Height          =   315
      Left            =   4560
      TabIndex        =   5
      Top             =   2250
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Arquivo"
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
      MCOL            =   16711935
      MPTR            =   1
      MICON           =   "frmArquivoIntegrativa.frx":00F8
      PICN            =   "frmArquivoIntegrativa.frx":0114
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cmbCadastro 
      Height          =   315
      ItemData        =   "frmArquivoIntegrativa.frx":01CF
      Left            =   1140
      List            =   "frmArquivoIntegrativa.frx":01DF
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1740
      Width           =   1395
   End
   Begin VB.ListBox lstTipo 
      Appearance      =   0  'Flat
      Height          =   1590
      ItemData        =   "frmArquivoIntegrativa.frx":020E
      Left            =   60
      List            =   "frmArquivoIntegrativa.frx":0210
      TabIndex        =   1
      Top             =   60
      Width           =   5745
   End
   Begin prjChameleon.chameleonButton cmdExec 
      Height          =   345
      Left            =   4560
      TabIndex        =   0
      ToolTipText     =   "Executar a operação selecionada"
      Top             =   2700
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Executar"
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
      MICON           =   "frmArquivoIntegrativa.frx":0212
      PICN            =   "frmArquivoIntegrativa.frx":022E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Tributacao.XP_ProgressBar PBar 
      Height          =   240
      Left            =   360
      TabIndex        =   8
      Top             =   2700
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   423
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
      Color           =   16777215
      Scrolling       =   1
      ShowText        =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "Vencidos até:"
      Height          =   225
      Left            =   2790
      TabIndex        =   11
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Cadastro...:"
      Height          =   225
      Left            =   180
      TabIndex        =   4
      Top             =   1800
      Width           =   915
   End
End
Attribute VB_Name = "frmArquivoIntegrativa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********* integração que será separada *****
Private Type tCDADebitoCorrecao
    idCdaIndex As Integer
    DataCorrecao As Date
End Type

Private Type tCDADebito
    idCdaIndex As Integer
    nCodReduz As Long
    nAno As Integer
    nLanc As Integer
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
    nCodTributo As Integer
    dDataCorrecao As Date
End Type

'********************************************

Private Type typeCDA
    nCDA As Long
    nCodReduz As Long
    nAno As Integer
    nLanc As Integer
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
    dDataInscricao As Date
    nNumCertidao As Long
    nNumLivro As Integer
    nNumPagina As Integer
End Type

Private Type typeCDANew
    nCDA As Long
    nCodReduz As Long
    nAno As Integer
    nLanc As Integer
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
    dDataInscricao As Date
    nNumCertidao As Long
    nPrincipal As Double
    nJuros  As Double
    nMulta As Double
    nCorrecao As Double
    nTotal As Double
    nNumLivro As Integer
    nNumPagina As Integer
End Type

Private Type typeCDADebito
    nCDA As Long
    nCodReduz As Long
    nAno As Integer
    nLanc As Integer
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
    nCodTributo As Integer
    sDescTributo As String
    nPrincipal As Double
    nMulta As Double
    nJuros As Double
    nCorrecao As Double
    nTotal As Double
    dDataVencto As Date
    nNumCertidao As Long
    nLivro As Integer
    nFolha As Integer
End Type

Private Type typeCDADebitoNew
    nCDA As Long
    nCodReduz As Long
    nAno As Integer
    nLanc As Integer
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
    nCodTributo As Integer
    sDescTributo As String
    nPrincipal As Double
    nMulta As Double
    nJuros As Double
    nCorrecao As Double
    nTotal As Double
    dDataVencto As Date
    nLivro As Integer
    nFolha As Integer
End Type

 
Dim cnn As ADODB.Connection

Private Type Reg00
    sTipoReg00 As String
    sCrc As String
    sNome As String
    sCPFCNPJ As String
    sRGIE As String
    sCodLogradouro As String
    sEndereco As String
    sNumero As String
    sComplemento As String
    sCodBairro As String
    sBairro As String
    sCodCidade As String
    sCidade As String
    sCep As String
    sEstado As String
    nValorTotal As Double
    sDataAtualizacao As String
    sIdAjuizamento As String
    sDescCadastro As String
    sCodCadastro As String
    sCadastroImob As String
    sFoneRes As String
    sFoneCom As String
    sCelular As String
    sFoneContato As String
    sEmail As String
    sCodLogradouroLocal As String
    sEnderecoLocal As String
    sNumeroLocal As String
    sCEPLocal As String
    sCodBairroLocal As String
    sBairroLocal As String
    sQuadra As String
    sLote As String
    sAtividade As String
    sMatricula As String
End Type

Private Type Reg01
    sTipoReg01 As String
    sCrcSocio As String
    sNomeSocio As String
    sCPFCNPJSocio As String
    sRGIE As String
    sCodLogradouro As String
    sEndereco As String
    sNumero As String
    sComplemento As String
    sCodBairro As String
    sBairro As String
    sCodCidade As String
    sCidade As String
    sCepSocio As String
    sEstadoSocio As String
    sClassificao As String
    sFoneRes As String
    sFoneCom As String
    sCelular As String
    sFoneContato As String
    sEmail As String
End Type

Private Type Reg10
    sTipoReg As String
    sTributo As String
    sExercicio As String
    sCodigoDivida As String
    sSubCodDivida As String
    sNumParcela As String
    sSubParcela As String
    sSeqInscricaoDA As String
    sNroCDA As String
    sFolha As String
    sLivro As String
    sDtVencimento As String
    sDtInscricao As String
    nPrincipal As Double
    nJuros As Double
    nMulta As Double
    nCorrecao As Double
    sHonorarios As String
    nTotal As Double
    sCodLogradouro As String
    sEndereco As String
    sNumero As String
    sQuadra As String
    sLote As String
    sCep As String
    sCodBairro As String
    sBairro As String
    sAtividade As String
    sMatricula As String
    sCartorio As String
    sProcessoAdm As String
    sAutoInfracao As String
    sNatureza As String
    sNroParcelamento As String
    sFundamento As String
    nTotalAcumulado As Double
    sNumExecFiscal As String
    sAnoExecFiscal As String
End Type

Private Type Reg20
    sTipoReg As String
    sExercicio As String
    sCodigoDivida As String
    sSubCodDivida As String
    sNumParcela As String
    sSubParcela As String
    sCodReceita As String
    sDescricao As String
    sValorOriginal As String
End Type

Dim xImovel As clsImovel

Enum FieldType
    Character = 0
    Numeric = 1
End Enum

Private Sub cmdIntegrar_Click()
Dim sql As String, RdoAcordo As rdoResultset, nAnoproc As Integer, nNumproc As Long, nCodReduz As Long, RdoDebito As rdoResultset, RdoAux2 As rdoResultset
Dim sTipoDivida As String, nCDA As Long, sRG As String, sCPF As String, sNumProc As String, dDataVencto As Date, nTot As Long, nPos As Long, RdoAux3 As rdoResultset
Dim sNome As String, sInscricao As String, sEndereco As String, nNumero As Integer, sComplemento As String, sComplementoEntrega As String
Dim sBairro As String, sCidade As String, sUF As String, sCep As String, sEnderecoEntrega As String, nNumEntrega As Integer, sBairroEntrega As String
Dim sCidadeEntrega As String, sUFEntrega As String, sCepEntrega As String, xImovel As clsImovel, sQuadra As String, sLote As String, nTipoEnd As Integer
Dim sTipoProp As String, nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, nLivro As Integer, nPag As Integer, nNumCertidao As Long, dDataInscricao As Date
Dim nNumExecFiscal As Long, nAnoExecFiscal As Integer, sNumExecFiscal As String, sRazaoSocial As String, RdoAcordoDebito As rdoResultset, dDataPrimeiraParc As Date
Dim nValorParc As Double, nValorHon As Double, RdoProcesso As rdoResultset, bCancelado As Boolean, nQtdeParc As Integer, dDataProc As Date, RdoTributo As rdoResultset
Dim nValorJuros As Double, nValorMulta As Double, nValorCorrecao As Double, qd As New rdoQuery, RdoJuros As rdoResultset
Exit Sub
Ocupado
Set xImovel = New clsImovel
nPos = 1
ConectaIntegrativa
If cnInt.Connect = "" Then
    MsgBox "Conexão falhou"
    Liberado
    Exit Sub
End If
Set qd.ActiveConnection = cn
qd.QueryTimeout = 0
GoTo Debitos
'Exit Sub

Ocupado
'Sql = "select * from acordos where iddevedor>500000 order by crcacordante"
'Set RdoAcordo = cnInt.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAcordo
'    nTot = .RowCount
'    nPos = 1
'    Do Until .EOF
'        If nPos Mod 50 = 0 Then
'            CallPb nPos, nTot
'        End If
'        nCodReduz = !crcacordante
'
'        Sql = "select nomecidadao from cidadao where codcidadao=" & nCodReduz
'        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'        With RdoAux2
'            Sql = "update acordos set nomeacordante='" & !nomecidadao & "' where iddevedor=" & nCodReduz
'            cnInt.Execute Sql, rdExecDirect
 '          .Close
 '       End With
 '
 '      nPos = nPos + 1
 '      .MoveNext
 '   Loop
 '  .Close
'end With
'liberado

'Exit Sub

'GoTo Debitos
'Exit Sub
'******** LIMPA TABELAS ******************
'Sql = "delete from acordobaixas"
'cnInt.Execute Sql, rdExecDirect

'Sql = "DBCC CHECKIDENT (acordobaixas, RESEED, 1)"
'cnInt.Execute Sql, rdExecDirect

'Sql = "delete from acordodebitos"
'cnInt.Execute Sql, rdExecDirect

'Sql = "DBCC CHECKIDENT (acordodebitos, RESEED, 1)"
'cnInt.Execute Sql, rdExecDirect

'Sql = "delete from acordostatus"
'cnInt.Execute Sql, rdExecDirect

'Sql = "DBCC CHECKIDENT (acordostatus, RESEED, 1)"
'cnInt.Execute Sql, rdExecDirect

'Sql = "delete from acordos"
'cnInt.Execute Sql, rdExecDirect

sql = "truncate table cadastro"
cnInt.Execute sql, rdExecDirect

sql = "DBCC CHECKIDENT (cadastro, RESEED, 1)"
cnInt.Execute sql, rdExecDirect

sql = "truncate table partes"
cnInt.Execute sql, rdExecDirect

sql = "DBCC CHECKIDENT (partes, RESEED, 1)"
cnInt.Execute sql, rdExecDirect

sql = "delete from cdadebitos"
cnInt.Execute sql, rdExecDirect

sql = "DBCC CHECKIDENT (cdadebitos, RESEED, 1)"
cnInt.Execute sql, rdExecDirect

sql = "delete from cdas"
cnInt.Execute sql, rdExecDirect

sql = "DBCC CHECKIDENT (cdas, RESEED, 1)"
cnInt.Execute sql, rdExecDirect
Exit Sub
GoTo Debitos
ACORDO:

'******** PARCELAMENTOS ******************

sql = "SELECT DISTINCT codreduzido, numprocesso From debitoparcela WHERE codreduzido=530446 "
'sql = "SELECT DISTINCT codreduzido, numprocesso From debitoparcela WHERE (codlancamento = 20) AND (statuslanc = 3 OR statuslanc = 18) order by codreduzido"
Set RdoAcordo = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAcordo
    nTot = .RowCount
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        nCodReduz = !CODREDUZIDO
        sNumProc = !NumProcesso
'        nAnoProc = Val(Mid(Trim$(sNumProc), InStr(1, Trim$(sNumProc), "/", vbBinaryCompare) + 1, 4))
'        nNumProc = Val(Left$(Trim$(sNumProc), InStr(1, Trim$(sNumProc), "/", vbBinaryCompare) - 1))
        nAnoproc = ExtraiNumeroProcesso(sNumProc)
        nAnoproc = ExtraiAnoProcesso(sNumProc)
        
       '******** DADOS DO PROCESSSO **********
       
        sql = "select * from processoreparc where numprocesso='" & sNumProc & "'"
        Set RdoProcesso = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoProcesso
            bCancelado = !Cancelado
            nQtdeParc = !qtdeparcela
            dDataProc = !DATAPROCESSO
           .Close
       End With
       
       '******** ENDEREÇO DO CONTRIBUINTE ***
        Select Case nCodReduz
            Case 1 To 99999
                sSetor = "IMOBILIÁRIO"
                xImovel.CarregaImovel nCodReduz
                sInscricao = xImovel.Inscricao
                sCodReduz = nCodReduz
                sRazaoSocial = xImovel.NomePropPrincipal
                sNome = sRazaoSocial
                sQuadra = xImovel.Li_Quadras
                sLote = xImovel.Li_Lotes
                
                xImovel.RetornaEndereco nCodReduz, Imobiliario, Localizacao
                sEndereco = xImovel.Endereco
                nNumero = xImovel.Numero
                sComplemento = xImovel.Complemento
                sBairro = xImovel.Bairro
                sCep = RetornaNumero(xImovel.Cep)
                sCidade = xImovel.Cidade
                sUF = xImovel.UF
                
                sEnderecoEntrega = xImovel.Ee_NomeLog
                nNumEntrega = xImovel.Ee_NumImovel
                sComplementoEntrega = xImovel.Ee_Complemento
                sBairroEntrega = xImovel.Ee_Bairro
                sCidadeEntrega = xImovel.Ee_Cidade
                sUFEntrega = xImovel.Ee_Uf
                sCepEntrega = RetornaNumero(xImovel.Ee_Cep)
                
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
        
            Case 100000 To 500000
                sSetor = "MOBILIÁRIO"
                sql = "select * from vwfullempresa3 where codigomob=" & nCodReduz
                Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                    sRazaoSocial = !RazaoSocial
                    sNome = sRazaoSocial
                    sInscricao = nCodReduz
                    sRG = SubNull(!inscestadual)
                    If Trim(sRG) = "" Then
                        sRG = SubNull(!rg)
                    End If
                    sCPF = SubNull(!Cnpj)
                    If Trim(sCPF) = "" Then
                        sCPF = SubNull(!cpf)
                    End If
                    sCodReduz = nCodReduz
                    sLote = ""
                    sQuadra = ""
                
                    xImovel.RetornaEndereco nCodReduz, Mobiliario, Localizacao
                    sEndereco = xImovel.Endereco
                    nNumero = xImovel.Numero
                    sComplemento = xImovel.Complemento
                    sBairro = xImovel.Bairro
                    sCep = RetornaNumero(xImovel.Cep)
                    sCidade = xImovel.Cidade
                    sUF = xImovel.UF
                    
                    sEnderecoEntrega = xImovel.Ee_NomeLog
                    nNumEntrega = xImovel.Ee_NumImovel
                    sComplementoEntrega = xImovel.Ee_Complemento
                    sBairroEntrega = xImovel.Ee_Bairro
                    sCidadeEntrega = xImovel.Ee_Cidade
                    sUFEntrega = xImovel.Ee_Uf
                    sCepEntrega = RetornaNumero(xImovel.Ee_Cep)
                   .Close
                End With
            Case 500000 To 800000
                sSetor = "TAXAS DIVERSAS"
                sql = "SELECT codcidadao,nomecidadao,cpf,cnpj,rg from cidadao WHERE CODCIDADAO=" & nCodReduz
                Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    If .RowCount > 0 Then
                        sNomeResp = !nomecidadao
                        sNome = sNomeResp
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
                sInscricao = nCodReduz
                sCodReduz = nCodReduz
                sLote = ""
                sQuadra = ""
                
                xImovel.RetornaEndereco nCodReduz, cidadao, cadastrocidadao
                sEndereco = xImovel.Endereco
                nNumero = xImovel.Numero
                sComplemento = xImovel.Complemento
                sBairro = xImovel.Bairro
                sCep = RetornaNumero(xImovel.Cep)
                sCidade = xImovel.Cidade
                sUF = xImovel.UF
                
                sEnderecoEntrega = xImovel.Ee_NomeLog
                nNumEntrega = xImovel.Ee_NumImovel
                sComplementoEntrega = xImovel.Ee_Complemento
                sBairroEntrega = xImovel.Ee_Bairro
                sCidadeEntrega = xImovel.Ee_Cidade
                sUFEntrega = xImovel.Ee_Uf
                sCepEntrega = RetornaNumero(xImovel.Ee_Cep)
        End Select
        
       '***** VALOR DO HONORÁRIO ************************
        
        sql = "SELECT * FROM debitoparcela WHERE (debitoparcela.codreduzido = " & nCodReduz & ") AND "
        sql = sql & "(debitoparcela.numparcela = 1) AND (debitoparcela.numprocesso = '" & sNumProc & "')"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            dDataPrimeiraParc = !DataVencimento
            nAno = !AnoExercicio
            nLanc = !CodLancamento
            nSeq = !SeqLancamento
            nParc = !NumParcela
            nCompl = !CODCOMPLEMENTO
            
            sql = "SELECT SUM(VALORTRIBUTO) AS SOMA FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND "
            sql = sql & "SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO <> 3"
            Set RdoAux4 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux4
                nValorParc = FormatNumber(!soma, 2)
               .Close
            End With
            
            sql = "SELECT valortributo FROM debitotributo WHERE codreduzido = " & nCodReduz & " and anoexercicio=" & !AnoExercicio & " and "
            sql = sql & "codlancamento=" & !CodLancamento & " and seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and "
            sql = sql & "codcomplemento=" & !CODCOMPLEMENTO & " and codtributo=90"
            Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux3
                If .RowCount > 0 Then
                    nValorHon = !VALORTRIBUTO * nQtdeParc
                Else
                    nValorHon = 0
                End If
               .Close
            End With
           .Close
        
        End With
        
       '******** GRAVA O ACORDO *************************
        
        sql = "insert acordos(idacordo,anoacordo,dtparcelamento,setordevedor,iddevedor,nroprocessoadm,crcacordante,nomeacordante,cpfcnpj,rginscrestadual,"
        sql = sql & "cep,endereco,numero,complemento,bairro, cidade,estado,vlrtotal,qtdparcelas,primeirovencimento,vlrtotalhonorarios,qtdparcelashonorarios,"
        sql = sql & "vlrparcelahonorarios,dtvenctohonorarios,VlrTotalDespesas, QtdParcelasDespesas, VlrParcelaDespesas, DtVenctoDespesas, DtGeracao) values ("
        sql = sql & nNumproc & "," & nAnoproc & ",'" & Format(dDataProc, "mm/dd/yyyy") & "','" & sSetor & "',"
        sql = sql & nCodReduz & ",'" & nNumproc & "/" & nAnoproc & "'," & nCodReduz & ",'" & Mask(sRazaoSocial) & "','" & sCPF & "','" & sRG & "','" & sCep & "','" & Mask(sEndereco) & "',"
        sql = sql & nNumero & ",'" & Mask(Left(sComplemento, 40)) & "','" & sBairro & "','" & Mask(sCidade) & "','" & sUF & "'," & Virg2Ponto(Round((nValorParc * nQtdeParc), 2)) & "," & nQtdeParc & ",'"
        sql = sql & Format(dDataPrimeiraParc, "mm/dd/yyyy") & "'," & Virg2Ponto(Round(nValorHon, 2)) & "," & IIf(nValorHon = 0, 0, nQtdeParc) & "," & Virg2Ponto(Round((nValorHon / nQtdeParc), 2)) & ","
        sql = sql & IIf(nValorHon = 0, "Null", "'" & Format(dDataPrimeiraParc, "mm/dd/yyyy") & "'") & "," & "0,0,0," & "Null" & ",'" & Format(Now, "mm/dd/yyyy") & "')"
       ' cnInt.Execute sql, rdExecDirect
        
        
       '******** GRAVA NA TABELA ACORDOSTATUS ***********
        sStatus = IIf(bCancelado = False, "PARCELAMENTO EM DIA", "PARCEL.CANCELADO")
        sql = "insert acordostatus(idacordo,anoacordo,dtocorrencia,ocorrencia,dtgeracao) values("
        sql = sql & nNumproc & "," & nAnoproc & ",'" & Format(Now, "mm/dd/yyyy") & "','" & sStatus & "','" & Format(Now, "mm/dd/yyyy") & "')"
        'cnInt.Execute sql, rdExecDirect
        
       
       '******** GRAVA OS DÉBITOS DO ACORDO *************
        sql = "SELECT DISTINCT origemreparc.numprocesso, origemreparc.codreduzido, origemreparc.anoexercicio, origemreparc.codlancamento, origemreparc.numsequencia,"
        sql = sql & "origemreparc.numparcela, origemreparc.codcomplemento, SUM(debitotributo.valortributo) AS Total, debitoparcela.numerolivro, debitoparcela.paginalivro,debitoparcela.dataajuiza, debitoparcela.numcertidao, debitoparcela.datainscricao, debitoparcela.datavencimento,numexecfiscal,anoexecfiscal "
        sql = sql & "FROM origemreparc INNER JOIN debitotributo ON origemreparc.codreduzido = debitotributo.codreduzido AND origemreparc.anoexercicio = debitotributo.anoexercicio AND "
        sql = sql & "origemreparc.codlancamento = debitotributo.codlancamento AND origemreparc.numsequencia = debitotributo.seqlancamento AND "
        sql = sql & "origemreparc.numparcela = debitotributo.numparcela AND origemreparc.codcomplemento = debitotributo.codcomplemento INNER JOIN debitoparcela ON origemreparc.codreduzido = debitoparcela.codreduzido AND "
        sql = sql & "origemreparc.anoexercicio = debitoparcela.anoexercicio AND origemreparc.codlancamento = debitoparcela.codlancamento AND origemreparc.numsequencia = debitoparcela.seqlancamento AND "
        sql = sql & "origemreparc.NumParcela = debitoparcela.NumParcela And origemreparc.CODCOMPLEMENTO = debitoparcela.CODCOMPLEMENTO GROUP BY origemreparc.numprocesso, origemreparc.codreduzido, origemreparc.anoexercicio, origemreparc.codlancamento, origemreparc.numsequencia,"
        sql = sql & "origemreparc.NumParcela , origemreparc.CODCOMPLEMENTO, debitoparcela.numerolivro, debitoparcela.paginalivro,debitoparcela.dataajuiza, debitoparcela.numcertidao, debitoparcela.datainscricao, debitoparcela.datavencimento,numexecfiscal,anoexecfiscal "
        sql = sql & "HAVING origemreparc.numprocesso = '" & sNumProc & "' AND origemreparc.codreduzido =" & nCodReduz
        sql = sql & " ORDER BY origemreparc.anoexercicio, origemreparc.numparcela"
        Set RdoAcordoDebito = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAcordoDebito
            Do Until .EOF
                nAno = !AnoExercicio
                nLanc = !CodLancamento
                nSeq = !numsequencia
                nParc = !NumParcela
                nCompl = !CODCOMPLEMENTO
                
                nPagina = Val(SubNull(!paginalivro))
                nLivro = Val(SubNull(!numerolivro))
                nNumCertidao = Val(SubNull(!numcertidao))
                dDataInscricao = IIf(IsNull(!datainscricao), CDate("01/01/1900"), !datainscricao)
                dDataVencto = !DataVencimento
                               
                nNumExecFiscal = Val(SubNull(!numexecfiscal))
                nAnoExecFiscal = Val(SubNull(!anoexecfiscal))
                If nNumExecFiscal > 0 Then
                    sNumExecFiscal = CStr(nNumExecFiscal) & "/" & CStr(nAnoExecFiscal)
                Else
                    sNumExecFiscal = ""
                End If
               
               
               'GRAVA NA TABELA ACORDODEBITO
                sql = "insert acordodebitos(idacordo,anoacordo,nrolivro,nrofolha,seq,lancamento,exercicio,vlroriginal,vlrcorrecao,vlrjuros,vlrmulta,vlrtotal,nroparcela,complparcela,ajuizado,dtgeracao) values("
                sql = sql & nNumproc & "," & nAnoproc & "," & nLivro & "," & nPagina & "," & nSeq & "," & nLanc & "," & nAno & "," & Virg2Ponto(Format(!Total, "#0.##")) & ",0,0,0," & Virg2Ponto(Format(!Total, "#0.##")) & ","
                sql = sql & nParc & "," & nCompl & "," & IIf(IsNull(!dataajuiza), 0, 1) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
         '       cnInt.Execute sql, rdExecDirect
               
               
               '*** GRAVA AS CDA, CDADebito, PARTES E CADASTRO ***
                    
                sql = "INSERT CDAs(IdDevedor,SetorDevedor,DtInscricao,NroCertidao,NroLivro,NroFolha,NroOrdem,DtGeracao) values("
                sql = sql & nCodReduz & ",'" & sSetor & "','" & Format(dDataInscricao, "mm/dd/yyyy") & "'," & nNumCertidao & ","
                sql = sql & nLivro & "," & nPagina & ",'" & sNumExecFiscal & "','" & Format(Now, "mm/dd/yyyy") & "')"
          '      cnInt.Execute sql, rdExecDirect
                
                sql = "select @@identity as LastKey"
                Set RdoAux2 = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                nCDA = RdoAux2!lastkey
                RdoAux2.Close
                    
                    
                If nCodReduz < 500000 Then
                    sql = "INSERT Cadastro(IdCDA,SetorDevedor,Crc,Nome,Inscricao,CPFCnpj,RgInscrEstadual,LocalCep,LocalEndereco,LocalNumero,LocalComplemento,"
                    sql = sql & "LocalBairro,LocalCidade,LocalEstado,Quadra,Lote,EntregaCep,EntregaEndereco,EntregaNumero,EntregaComplemento,EntregaBairro,"
                    sql = sql & "EntregaCidade,EntregaEstado,DtGeracao) values("
                    sql = sql & nCDA & ",'" & sSetor & "'," & nCodReduz & ",'" & Mask(SubNull(sNome)) & "','" & sNumInsc & "','" & sCPF & "','" & sRG & "','"
                    sql = sql & sCep & "','" & sEndereco & "'," & nNumero & ",'" & Mask(Left(sComplemento, 50)) & "','" & sBairro & "','" & sCidade & "','" & sUF & "','"
                    sql = sql & sQuadra & "','" & sLote & "','" & sCepEntrega & "','" & sEnderecoEntrega & "'," & nNumEntrega & ",'" & sComplementoEntrega & "','"
                    sql = sql & sBairroEntrega & "','" & sCidadeEntrega & "','" & sUFEntrega & "','" & Format(Now, "mm/dd/yyyy") & "')"
           '         cnInt.Execute sql, rdExecDirect
                Else
                    sql = "select * from vwFullCidadao where codcidadao=" & nCodReduz
                    Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        sCPF = SubNull(!Cnpj)
                        If Trim(sCPF) = "" Then
                            sCPF = SubNull(!cpf)
                        End If
                        sql = "INSERT Partes(idCDA,Tipo,Crc,Nome,CpfCnpj,RgInscrEstadual,Cep,Endereco,Numero,Complemento,Bairro,Cidade,Estado,DtGeracao) Values("
                        sql = sql & nCDA & ",'Principal'," & !CodCidadao & ",'" & Mask(!nomecidadao) & "','" & sCPF & "','" & SubNull(!rg) & "','" & SubNull(!Cep) & "','"
                        sql = sql & Mask(SubNull(!Endereco)) & "'," & Val(SubNull(!NUMIMOVEL)) & ",'" & Mask(SubNull(!Complemento)) & "','" & Mask(SubNull(!DescBairro)) & "','"
                        sql = sql & Mask(SubNull(!descCidade)) & "','" & SubNull(!SiglaUF) & "','" & Format(Now, "mm/dd/yyyy") & "')"
            '            cnInt.Execute sql, rdExecDirect
                       .Close
                    End With
                End If
                
                If nCodReduz < 100000 Then 'cadastra os proprietarios e compromissarios
                    sql = "SELECT cadimob.codreduzido, proprietario.codcidadao, proprietario.tipoprop, vwFULLCIDADAO.nomecidadao, vwFULLCIDADAO.cpf,"
                    sql = sql & "vwFULLCIDADAO.cnpj, vwFULLCIDADAO.numimovel, vwFULLCIDADAO.complemento, vwFULLCIDADAO.siglauf, vwFULLCIDADAO.cep,"
                    sql = sql & "vwFULLCIDADAO.rg , vwFULLCIDADAO.orgao, vwFULLCIDADAO.DescBairro, vwFULLCIDADAO.desccidade, vwFULLCIDADAO.Endereco "
                    sql = sql & "FROM cadimob INNER JOIN proprietario ON cadimob.codreduzido = proprietario.codreduzido INNER JOIN vwFULLCIDADAO ON "
                    sql = sql & "proprietario.codcidadao = vwFULLCIDADAO.codcidadao where cadimob.codreduzido=" & nCodReduz
                    Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        Do Until .EOF
                            sCPF = SubNull(!Cnpj)
                            If Trim(sCPF) = "" Then
                                sCPF = SubNull(!cpf)
                            End If
                            sTipoProp = IIf(!tipoprop = "P", "Principal", "Compromissário")
                            sql = "INSERT Partes(idCDA,Tipo,Crc,Nome,CpfCnpj,RgInscrEstadual,Cep,Endereco,Numero,Complemento,Bairro,Cidade,Estado,DtGeracao) Values("
                            sql = sql & nCDA & ",'" & sTipoProp & "'," & nCodReduz & ",'" & Mask(!nomecidadao) & "','" & sCPF & "','" & SubNull(!rg) & "','" & SubNull(!Cep) & "','"
                            sql = sql & Mask(SubNull(!Endereco)) & "'," & Val(SubNull(!NUMIMOVEL)) & ",'" & Mask(SubNull(!Complemento)) & "','" & Mask(SubNull(!DescBairro)) & "','"
                            sql = sql & Mask(SubNull(!descCidade)) & "','" & SubNull(!SiglaUF) & "','" & Format(Now, "mm/dd/yyyy") & "')"
             '               cnInt.Execute sql, rdExecDirect
                           .MoveNext
                        Loop
                       .Close
                    End With
                ElseIf nCodReduz >= 100000 And nCodReduz < 300000 Then   'cadastra os socios
                    sql = "SELECT * from vwmobiliarioproprietario where codmobiliario=" & nCodReduz
                    Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        Do Until .EOF
                            sCPF = SubNull(!Cnpj)
                            If Trim(sCPF) = "" Then
                                sCPF = SubNull(!cpf)
                            End If
                            sTipoProp = "Sócio"
                            sql = "INSERT Partes(idCDA,Tipo,Crc,Nome,CpfCnpj,RgInscrEstadual,Cep,Endereco,Numero,Complemento,Bairro,Cidade,Estado,DtGeracao) Values("
                            sql = sql & nCDA & ",'" & sTipoProp & "'," & nCodReduz & ",'" & Mask(!nomecidadao) & "','" & sCPF & "','" & SubNull(!rg) & "','" & SubNull(!Cep) & "','"
                            sql = sql & Mask(SubNull(!Endereco)) & "'," & Val(SubNull(!NUMIMOVEL)) & ",'" & Mask(SubNull(!Complemento)) & "','" & Mask(SubNull(!DescBairro)) & "','"
                            sql = sql & Mask(SubNull(!descCidade)) & "','" & SubNull(!SiglaUF) & "','" & Format(Now, "mm/dd/yyyy") & "')"
              '              cnInt.Execute sql, rdExecDirect
                           .MoveNext
                        Loop
                       .Close
                    End With
                End If
                
                '**** GRAVA CDADebitos *********************
                sql = "SELECT debitotributo.codtributo,valortributo,abrevtributo FROM debitotributo INNER JOIN tributo ON debitotributo.codtributo = tributo.codtributo "
                sql = sql & "where codreduzido=" & nCodReduz & " and anoexercicio=" & nAno & " and codlancamento=" & nLanc & " and seqlancamento=" & nSeq & " and "
                sql = sql & "numparcela=" & nParc & " and codcomplemento=" & nCompl
                Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    Do Until .EOF
                        
                        sql = "INSERT CDADebitos(idCDA,CodTributo,Tributo,Exercicio,Lancamento,Seq,NroParcela,ComplParcela,DtVencimento,VlrOriginal,DtGeracao) values("
                        sql = sql & nCDA & "," & !CodTributo & ",'" & Mask(!abrevTributo) & "'," & nAno & "," & nLanc & "," & nSeq & "," & nParc & ","
                        sql = sql & nCompl & ",'" & Format(dDataVencto, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(!VALORTRIBUTO)) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
               '         cnInt.Execute sql, rdExecDirect
                        
                       .MoveNext
                    Loop
                   .Close
                End With
               
               .MoveNext
            Loop
           .Close
        End With
        
        
       '**** GRAVA AcordosBaixa *********************
        sql = "select anoexercicio,codlancamento,seqlancamento,numparcela,codcomplemento from debitoparcela "
        sql = sql & "where codreduzido=" & nCodReduz & " and numprocesso='" & sNumProc & "'"
        Set RdoProcesso = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoProcesso
            Do Until .EOF
                sql = "SELECT * FROM debitopago WHERE codreduzido = " & nCodReduz & " and anoexercicio=" & !AnoExercicio & " and "
                sql = sql & "codlancamento=" & !CodLancamento & " and seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and "
                sql = sql & "codcomplemento=" & !CODCOMPLEMENTO
                Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                With RdoAux3
                    Do Until .EOF
                        'GRAVA NA TABELA ACORDOBAIXAS
                         sql = "insert acordobaixas(idAcordo, anoAcordo, DtBaixa, TipoBaixa, NroParcela, VlrOriginal, VlrCorrecao, VlrJuros, VlrMulta, VlrTotal, DtGeracao) values("
                         sql = sql & nNumproc & "," & nAnoproc & ",'" & Format(RdoAux3!DataPagamento, "mm/dd/yyyy") & "','PAGAMENTO'," & RdoProcesso!NumParcela & "," & Virg2Ponto(CStr(RdoAux3!ValorPagoreal)) & ","
                         sql = sql & 0 & "," & 0 & "," & 0 & "," & Virg2Ponto(CStr(RdoAux3!ValorPagoreal)) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
                         cnInt.Execute sql, rdExecDirect
                    
                        .MoveNext
                    Loop
                   .Close
                End With
               .MoveNext
            Loop
           .Close
        End With
        
        
       '***** PRÓXIMO PROCESSO ******************
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With



PBar.value = 0
Set xImovel = Nothing
Exit Sub
Debitos:
'*******  DÉBITOS ************************

sql = "select * from debitoparcela where codreduzido between 500000 and 600000 and codlancamento<>20 and  statuslanc=3 and datainscricao is not null and dataajuiza is null and codreduzido>0 order by codreduzido,anoexercicio,seqlancamento,numparcela"
'Sql = "select * from debitoparcela where codreduzido in (117094,117166,117474,114693,117028,117634,108369,117877,115812,113424) and codlancamento<>20 and  statuslanc=3 and datainscricao is not null and dataajuiza is null and codreduzido>0 order by codreduzido,anoexercicio,seqlancamento,numparcela"
Set RdoDebito = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoDebito
    nPos = 1
    nTot = .RowCount
    Do Until .EOF
        nCodReduz = !CODREDUZIDO
        sql = "select * from debitoparcela where codreduzido =" & !CODREDUZIDO & " and anoexercicio<2013 and codlancamento<>20 and  statuslanc=3 and datainscricao is not null and dataajuiza is null"
        Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
        If RdoAux3.RowCount = 0 Then
            DoEvents
            GoTo Proximo2
        End If
        
        If nPos Mod 50 = 0 Then
            CallPb nPos, nTot
            DoEvents
        End If
        
       '******** ENDEREÇO DO CONTRIBUINTE ***
        Select Case nCodReduz
            Case 1 To 99999
                sSetor = "IMOBILIÁRIO"
                xImovel.CarregaImovel nCodReduz
                sInscricao = xImovel.Inscricao
                sCodReduz = nCodReduz
                sRazaoSocial = xImovel.NomePropPrincipal
                sNome = sRazaoSocial
                sQuadra = xImovel.Li_Quadras
                sLote = xImovel.Li_Lotes
                
                xImovel.RetornaEndereco nCodReduz, Imobiliario, Localizacao
                sEndereco = xImovel.Endereco
                nNumero = xImovel.Numero
                sComplemento = xImovel.Complemento
                sBairro = xImovel.Bairro
                sCep = RetornaNumero(xImovel.Cep)
                sCidade = xImovel.Cidade
                sUF = xImovel.UF
                If xImovel.Ee_TipoEnd = 0 Then
                    xImovel.RetornaEndereco nCodReduz, Imobiliario, Localizacao
                ElseIf xImovel.Ee_TipoEnd = 1 Then
                    xImovel.RetornaEndereco nCodReduz, Imobiliario, cadastrocidadao
                Else
                    xImovel.RetornaEndereco nCodReduz, Imobiliario, Entrega
                End If
               sEnderecoEntrega = xImovel.Endereco
                nNumEntrega = Val(xImovel.Numero)
                sComplementoEntrega = xImovel.Complemento
                sBairroEntrega = xImovel.Bairro
                sCepEntrega = RetornaNumero(xImovel.Cep)
                sCidadeEntrega = xImovel.Cidade
                sUFEntrega = xImovel.UF
                
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
        
            Case 100000 To 500000
                sSetor = "MOBILIÁRIO"
                sql = "select * from vwfullempresa3 where codigomob=" & nCodReduz
                Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                    If RdoAux.RowCount = 0 Then GoTo Proximo2
                    sRazaoSocial = !RazaoSocial
                    sNome = sRazaoSocial
                    sInscricao = nCodReduz
                    sRG = SubNull(!inscestadual)
                    If Trim(sRG) = "" Then
                        sRG = SubNull(!rg)
                    End If
                    sCPF = SubNull(!Cnpj)
                    If Trim(sCPF) = "" Then
                        sCPF = SubNull(!cpf)
                    End If
                    sCodReduz = nCodReduz
                    sLote = ""
                    sQuadra = ""
                
                    xImovel.RetornaEndereco nCodReduz, Mobiliario, Localizacao
                    sEndereco = xImovel.Endereco
                    nNumero = xImovel.Numero
                    sComplemento = xImovel.Complemento
                    sBairro = xImovel.Bairro
                    sCep = RetornaNumero(xImovel.Cep)
                    sCidade = xImovel.Cidade
                    sUF = xImovel.UF
                    xImovel.RetornaEndereco nCodReduz, Mobiliario, Entrega
                    sEnderecoEntrega = xImovel.Endereco
                    nNumEntrega = xImovel.Numero
                    sComplementoEntrega = xImovel.Complemento
                    sBairroEntrega = xImovel.Bairro
                    sCidadeEntrega = xImovel.Cidade
                    sUFEntrega = xImovel.UF
                    sCepEntrega = xImovel.Cep
                   .Close
                End With
            Case 500000 To 800000
                sSetor = "TAXAS DIVERSAS"
                sql = "SELECT codcidadao,nomecidadao,cpf,cnpj,rg from cidadao WHERE CODCIDADAO=" & nCodReduz
                Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    If .RowCount > 0 Then
                        sNomeResp = !nomecidadao
                        sNome = sNomeResp
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
                sInscricao = nCodReduz
                sCodReduz = nCodReduz
                sLote = ""
                sQuadra = ""
                
                xImovel.RetornaEndereco nCodReduz, cidadao, cadastrocidadao
                sEndereco = xImovel.Endereco
                nNumero = Val(xImovel.Numero)
                sComplemento = xImovel.Complemento
                sBairro = xImovel.Bairro
                sCep = RetornaNumero(xImovel.Cep)
                sCidade = xImovel.Cidade
                sUF = xImovel.UF
                
                sEnderecoEntrega = xImovel.Ee_NomeLog
                nNumEntrega = xImovel.Ee_NumImovel
                sComplementoEntrega = xImovel.Ee_Complemento
                sBairroEntrega = xImovel.Ee_Bairro
                sCidadeEntrega = xImovel.Ee_Cidade
                sUFEntrega = xImovel.Ee_Uf
                sCepEntrega = RetornaNumero(xImovel.Ee_Cep)
        End Select
                
        nAno = !AnoExercicio
        nLanc = !CodLancamento
        nSeq = !SeqLancamento
        nParc = !NumParcela
        nCompl = !CODCOMPLEMENTO
        
        nPagina = Val(SubNull(!paginalivro))
        nLivro = Val(SubNull(!numerolivro))
        nNumCertidao = Val(SubNull(!numcertidao))
        dDataInscricao = IIf(IsNull(!datainscricao), CDate("01/01/1900"), !datainscricao)
        dDataVencto = !DataVencimento
                       
        nNumExecFiscal = Val(SubNull(!numexecfiscal))
        nAnoExecFiscal = Val(SubNull(!anoexecfiscal))
        If nNumExecFiscal > 0 Then
            sNumExecFiscal = CStr(nNumExecFiscal) & "/" & CStr(nAnoExecFiscal)
        Else
            sNumExecFiscal = ""
        End If
        
        
        sql = "INSERT CDAs(IdDevedor,SetorDevedor,DtInscricao,NroCertidao,NroLivro,NroFolha,NroOrdem,DtGeracao) values("
        sql = sql & nCodReduz & ",'" & sSetor & "','" & Format(dDataInscricao, "mm/dd/yyyy") & "'," & nNumCertidao & ","
        sql = sql & nLivro & "," & nPagina & ",'" & sNumExecFiscal & "','" & Format(Now, "mm/dd/yyyy") & "')"
        cnInt.Execute sql, rdExecDirect
        
        sql = "select @@identity as LastKey"
        Set RdoAux2 = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        nCDA = RdoAux2!lastkey
        RdoAux2.Close
            
            
        If nCodReduz < 500000 Then
            sql = "INSERT Cadastro(IdCDA,SetorDevedor,Crc,Nome,Inscricao,CPFCnpj,RgInscrEstadual,LocalCep,LocalEndereco,LocalNumero,LocalComplemento,"
            sql = sql & "LocalBairro,LocalCidade,LocalEstado,Quadra,Lote,EntregaCep,EntregaEndereco,EntregaNumero,EntregaComplemento,EntregaBairro,"
            sql = sql & "EntregaCidade,EntregaEstado,DtGeracao) values("
            sql = sql & nCDA & ",'" & sSetor & "'," & nCodReduz & ",'" & Left(Mask(SubNull(sNome)), 50) & "','" & sInscricao & "','" & sCPF & "','" & sRG & "','"
            sql = sql & sCep & "','" & Mask(sEndereco) & "'," & nNumero & ",'" & Mask(Left(sComplemento, 50)) & "','" & sBairro & "','" & sCidade & "','" & sUF & "','"
            sql = sql & Mask(sQuadra) & "','" & sLote & "','" & sCepEntrega & "','" & Mask(sEnderecoEntrega) & "'," & nNumEntrega & ",'" & Mask(Left(sComplementoEntrega, 50)) & "','"
            sql = sql & sBairroEntrega & "','" & sCidadeEntrega & "','" & sUFEntrega & "','" & Format(Now, "mm/dd/yyyy") & "')"
            cnInt.Execute sql, rdExecDirect
        Else
            sql = "select * from vwFullCidadao where codcidadao=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                sCPF = SubNull(!Cnpj)
                If Trim(sCPF) = "" Then
                    sCPF = SubNull(!cpf)
                End If
                sql = "INSERT Partes(idCDA,Tipo,Crc,Nome,CpfCnpj,RgInscrEstadual,Cep,Endereco,Numero,Complemento,Bairro,Cidade,Estado,DtGeracao) Values("
                sql = sql & nCDA & ",'Principal'," & !CodCidadao & ",'" & Mask(!nomecidadao) & "','" & sCPF & "','" & SubNull(!rg) & "','" & SubNull(!Cep) & "','"
                sql = sql & Mask(SubNull(!Endereco)) & "'," & Val(SubNull(!NUMIMOVEL)) & ",'" & Mask(SubNull(!Complemento)) & "','" & Mask(SubNull(!DescBairro)) & "','"
                sql = sql & Mask(SubNull(!descCidade)) & "','" & SubNull(!SiglaUF) & "','" & Format(Now, "mm/dd/yyyy") & "')"
                cnInt.Execute sql, rdExecDirect
               .Close
            End With
        End If
        
        If nCodReduz < 100000 Then 'cadastra os proprietarios e compromissarios
            sql = "SELECT cadimob.codreduzido, proprietario.codcidadao, proprietario.tipoprop, vwFULLCIDADAO.nomecidadao, vwFULLCIDADAO.cpf,"
            sql = sql & "vwFULLCIDADAO.cnpj, vwFULLCIDADAO.numimovel, vwFULLCIDADAO.complemento, vwFULLCIDADAO.siglauf, vwFULLCIDADAO.cep,"
            sql = sql & "vwFULLCIDADAO.rg , vwFULLCIDADAO.orgao, vwFULLCIDADAO.DescBairro, vwFULLCIDADAO.desccidade, vwFULLCIDADAO.Endereco "
            sql = sql & "FROM cadimob INNER JOIN proprietario ON cadimob.codreduzido = proprietario.codreduzido INNER JOIN vwFULLCIDADAO ON "
            sql = sql & "proprietario.codcidadao = vwFULLCIDADAO.codcidadao where cadimob.codreduzido=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                Do Until .EOF
                    sCPF = SubNull(!Cnpj)
                    If Trim(sCPF) = "" Then
                        sCPF = SubNull(!cpf)
                    End If
                    sTipoProp = IIf(!tipoprop = "P", "Principal", "Compromissário")
                    sql = "INSERT Partes(idCDA,Tipo,Crc,Nome,CpfCnpj,RgInscrEstadual,Cep,Endereco,Numero,Complemento,Bairro,Cidade,Estado,DtGeracao) Values("
                    sql = sql & nCDA & ",'" & sTipoProp & "'," & nCodReduz & ",'" & Mask(!nomecidadao) & "','" & sCPF & "','" & SubNull(!rg) & "','" & SubNull(!Cep) & "','"
                    sql = sql & Mask(SubNull(!Endereco)) & "'," & Val(SubNull(!NUMIMOVEL)) & ",'" & Mask(SubNull(!Complemento)) & "','" & Mask(SubNull(!DescBairro)) & "','"
                    sql = sql & Mask(SubNull(!descCidade)) & "','" & SubNull(!SiglaUF) & "','" & Format(Now, "mm/dd/yyyy") & "')"
                    cnInt.Execute sql, rdExecDirect
                   .MoveNext
                Loop
               .Close
            End With
        ElseIf nCodReduz >= 100000 And nCodReduz < 300000 Then   'cadastra os socios
            sql = "SELECT * from vwmobiliarioproprietario where codmobiliario=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                Do Until .EOF
                    sCPF = SubNull(!Cnpj)
                    If Trim(sCPF) = "" Then
                        sCPF = SubNull(!cpf)
                    End If
                    sTipoProp = "Sócio"
                    sql = "INSERT Partes(idCDA,Tipo,Crc,Nome,CpfCnpj,RgInscrEstadual,Cep,Endereco,Numero,Complemento,Bairro,Cidade,Estado,DtGeracao) Values("
                    sql = sql & nCDA & ",'" & sTipoProp & "'," & nCodReduz & ",'" & Mask(!nomecidadao) & "','" & sCPF & "','" & SubNull(!rg) & "','" & SubNull(!Cep) & "','"
                    sql = sql & Mask(SubNull(!Endereco)) & "'," & Val(SubNull(!NUMIMOVEL)) & ",'" & Mask(SubNull(!Complemento)) & "','" & Mask(SubNull(!DescBairro)) & "','"
                    sql = sql & Mask(SubNull(!descCidade)) & "','" & SubNull(!SiglaUF) & "','" & Format(Now, "mm/dd/yyyy") & "')"
                    cnInt.Execute sql, rdExecDirect
                   .MoveNext
                Loop
               .Close
            End With
        End If
        
        '**** GRAVA CDADebitos *********************
'        Sql = "SELECT debitotributo.codtributo,valortributo,abrevtributo FROM debitotributo INNER JOIN tributo ON debitotributo.codtributo = tributo.codtributo "
'        Sql = Sql & "where codreduzido=" & nCodReduz & " and anoexercicio=" & nAno & " and codlancamento=" & nLanc & " and seqlancamento=" & nSeq & " and "
'        Sql = Sql & "numparcela=" & nParc & " and codcomplemento=" & nCompl
'        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

        On Error Resume Next
        RdoJuros.Close
        On Error GoTo 0
                   
        qd.sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
        qd(0) = nCodReduz
        qd(1) = nCodReduz
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
        qd(12) = 1
        qd(13) = 99
        qd(14) = Format(Now, "mm/dd/yyyy")
        qd(15) = "Integrativa"
        Set RdoJuros = qd.OpenResultset(rdOpenKeyset)
        With RdoJuros
            Do Until .EOF
                sql = "INSERT CDADebitos(idCDA,CodTributo,Tributo,Exercicio,Lancamento,Seq,NroParcela,ComplParcela,DtVencimento,VlrOriginal,VlrMultas,VlrJuros,VlrCorrecao,DtGeracao) values("
                sql = sql & nCDA & "," & !CodTributo & ",'" & Mask(!abrevTributo) & "'," & nAno & "," & nLanc & "," & nSeq & "," & nParc & ","
                sql = sql & nCompl & ",'" & Format(dDataVencto, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(!VALORTRIBUTO)) & "," & Virg2Ponto(CStr(!ValorMulta)) & ","
                sql = sql & Virg2Ponto(CStr(!ValorJuros)) & "," & Virg2Ponto(CStr(!valorcorrecao)) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
                cnInt.Execute sql, rdExecDirect
               .MoveNext
            Loop
           .Close
        End With


Proximo2:
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With


'*****************************************

Liberado
MsgBox "fim"

End Sub


Private Sub btfix_Click()

Dim sql As String, RdoAcordo As rdoResultset, nAnoproc As Integer, nNumproc As Long, nCodReduz As Long, RdoDebito As rdoResultset, RdoAux2 As rdoResultset
Dim sTipoDivida As String, nCDA As Long, sRG As String, sCPF As String, sNumProc As String, dDataVencto As Date
Dim sNome As String, sInscricao As String, sEndereco As String, nNumero As Integer, sComplemento As String, sComplementoEntrega As String
Dim sBairro As String, sCidade As String, sUF As String, sCep As String, sEnderecoEntrega As String, nNumEntrega As Integer, sBairroEntrega As String
Dim sCidadeEntrega As String, sUFEntrega As String, sCepEntrega As String, xImovel As clsImovel, sQuadras As String, sLotes As String, nTipoEnd As Integer
Dim sTipoProp As String, nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, nLivro As Integer, nPag As Integer, nNumCertidao As Long, dDataInscricao As Date
Dim nNumExecFiscal As Long, nAnoExecFiscal As Integer, sNumExecFiscal As String

Set xImovel = New clsImovel
ConectaIntegrativa
sql = "SELECT idAcordo, anoAcordo, IdDevedor From Acordos WHERE (IdDevedor NOT IN (SELECT idDevedor FROM CDAs)) ORDER BY IdDevedor"
Set RdoAcordo = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAcordo
    Do Until .EOF
        nCodReduz = !iddevedor
        nAnoproc = !anoacordo
        nNumproc = !idacordo
        sNumProc = CStr(nNumproc) & "/" & CStr(nAnoproc)
        '***************************************************
        
        If nCodReduz < 100000 Then
            sTipoDivida = "Imobiliário"
            sql = "select * from vwfullimovel where codreduzido=" & nCodReduz
            Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                sNome = !nomecidadao
                sInscricao = !Inscricao
                sRG = SubNull(!rg)
                sCPF = SubNull(!cpf)
                If Trim(sCPF) = "" Then
                    sCPF = SubNull(!Cnpj)
                End If
                sEndereco = !Logradouro
                nNumero = !Li_Num
                sCep = RetornaNumero(RetornaCEP(!CodLogr, !Li_Num))
                sComplemento = SubNull(!Li_Compl)
                sBairro = SubNull(!DescBairro)
                sCidade = "JABOTICABAL"
                sUF = "SP"
                sQuadras = Left(SubNull(!Li_Quadras), 5)
                sLotes = Left(SubNull(!Li_Lotes), 20)
                nTipoEnd = RdoAux!Ee_TipoEnd
               .Close
            End With
            If nTipoEnd = 0 Then
                sEnderecoEntrega = sEndereco
                nNumEntrega = nNumero
                sCepEntrega = sCep
                sComplementoEntrega = sComplemento
                sBairroEntrega = sBairro
                sCidadeEntrega = sCidade
                sUFEntrega = sUF
            Else
                xImovel.RetornaEndereco nCodReduz, Imobiliario, IIf(nTipoEnd = 1, cadastrocidadao, Entrega)
                sEnderecoEntrega = xImovel.Endereco
                nNumEntrega = xImovel.Numero
                sCepEntrega = xImovel.Cep
                sComplementoEntrega = xImovel.Complemento
                sBairroEntrega = xImovel.Bairro
                sCidadeEntrega = xImovel.Cidade
                sUFEntrega = xImovel.UF
            End If
        ElseIf nCodReduz >= 100000 And nCodReduz < 300000 Then
            sTipoDivida = "Mobiliário"
            sql = "select * from vwfullempresa3 where codigomob=" & nCodReduz
            Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                sNome = !RazaoSocial
                sInscricao = nCodReduz
                sRG = SubNull(!inscestadual)
                If Trim(sRG) = "" Then
                    sRG = SubNull(!rg)
                End If
                sCPF = SubNull(!Cnpj)
                If Trim(sCPF) = "" Then
                    sCPF = SubNull(!cpf)
                End If
                sEndereco = !Logradouro
                nNumero = !Numero
                sCep = SubNull(!Cep)
                sComplemento = SubNull(!Complemento)
                sBairro = SubNull(!DescBairro)
                sCidade = SubNull(!descCidade)
                sUF = SubNull(!SiglaUF)
                sQuadras = ""
                sLotes = ""
               .Close
            End With
            sql = "SELECT MOBILIARIOENDENTREGA.CODMOBILIARIO, MOBILIARIOENDENTREGA.CODLOGRADOURO, MOBILIARIOENDENTREGA.NOMELOGRADOURO, "
            sql = sql & "MOBILIARIOENDENTREGA.NUMIMOVEL,MOBILIARIOENDENTREGA.COMPLEMENTO, MOBILIARIOENDENTREGA.UF,MOBILIARIOENDENTREGA.CODCIDADE,"
            sql = sql & "MOBILIARIOENDENTREGA.CODBAIRRO, MOBILIARIOENDENTREGA.CEP,MOBILIARIOENDENTREGA.DESCBAIRRO, MOBILIARIOENDENTREGA.DESCCIDADE, BAIRRO.DESCBAIRRO AS DESCBAIRRO2,"
            sql = sql & "CIDADE.DESCCIDADE AS DESCCIDADE2 FROM dbo.bairro INNER JOIN dbo.cidade ON dbo.bairro.siglauf = dbo.cidade.siglauf AND dbo.bairro.codcidade = dbo.cidade.codcidade RIGHT OUTER JOIN "
            sql = sql & "dbo.mobiliarioendentrega ON dbo.bairro.siglauf = dbo.mobiliarioendentrega.uf AND dbo.bairro.codcidade = dbo.mobiliarioendentrega.codcidade AND dbo.Bairro.CodBairro = dbo.mobiliarioendentrega.CodBairro "
            sql = sql & "Where CODMOBILIARIO = " & nCodReduz
            Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount > 0 Then
                    If !CodLogradouro = 0 Then
                        sEnderecoEntrega = !NomeLogradouro
                        nNumEntrega = !NUMIMOVEL
                        If !CodBairro = 0 Then
                            sBairroEntrega = SubNull(!DescBairro)
                        ElseIf !CodBairro = 999 Then
                            sBairroEntrega = ""
                        Else
                            sBairroEntrega = SubNull(!DescBairro2)
                        End If
                        sCidadeEntrega = SubNull(!descCidade)
                        If Trim(sCidadeEntrega) = "" Then
                            sCidadeEntrega = SubNull(!desccidade2)
                        End If
                        If !CodCidade = 413 Then
                            sCepEntrega = Format(!Cep, "00000000")
                        Else
                            sCepEntrega = RetornaNumero(RetornaCEP(!CodLogradouro, !NUMIMOVEL))
                        End If
                        sComplEntrega = SubNull(!Complemento)
                        sUFEntrega = !UF
                    Else
                        sql = "SELECT ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & !CodLogradouro
                        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux2
                            If .RowCount > 0 Then
                                sEnderecoEntrega = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                                nNumEntrega = SubNull(RdoAux!NUMIMOVEL)
                                sCepEntrega = RetornaNumero(RetornaCEP(RdoAux!CodLogradouro, RdoAux!NUMIMOVEL))
                            End If
                            sBairroEntrega = SubNull(RdoAux!DescBairro2)
                            sCidadeEntrega = SubNull(RdoAux!desccidade2)
                            sComplementoEntrega = SubNull(RdoAux!Complemento)
                            sUFEntrega = RdoAux!UF
                           .Close
                        End With
                    End If
                Else
                    sEnderecoEntrega = sEndereco
                    nNumEntrega = nNumero
                    sCepEntrega = sCep
                    sComplementoEntrega = sComplemento
                    sBairroEntrega = sBairro
                    sCidadeEntrega = sCidade
                    sUFEntrega = sUF
                End If
               .Close
            End With
        
        Else
            sTipoDivida = "Taxas Diversas"
        End If
        
        '***************************************************
        
        sql = "SELECT origemreparc.numprocesso, origemreparc.codreduzido, origemreparc.anoexercicio, origemreparc.codlancamento, origemreparc.numsequencia, "
        sql = sql & "origemreparc.numparcela, origemreparc.codcomplemento,debitoparcela.numexecfiscal,debitoparcela.anoexecfiscal,debitoparcela.datavencimento, debitoparcela.numerolivro, debitoparcela.paginalivro, debitoparcela.numcertidao,"
        sql = sql & "debitoparcela.datainscricao FROM origemreparc INNER JOIN debitoparcela ON origemreparc.codreduzido = debitoparcela.codreduzido AND origemreparc.anoexercicio = debitoparcela.anoexercicio AND "
        sql = sql & "origemreparc.codlancamento = debitoparcela.codlancamento AND origemreparc.numsequencia = debitoparcela.seqlancamento AND origemreparc.NumParcela = debitoparcela.NumParcela And "
        sql = sql & "origemreparc.CODCOMPLEMENTO = debitoparcela.CODCOMPLEMENTO WHERE origemreparc.numprocesso = '" & sNumProc & "'"
        Set RdoDebito = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoDebito
            Do Until .EOF
                nAno = !AnoExercicio
                nLanc = !CodLancamento
                nSeq = !numsequencia
                nParc = !NumParcela
                nCompl = !CODCOMPLEMENTO
                nLivro = Val(SubNull(!numerolivro))
                nPag = Val(SubNull(!paginalivro))
                nNumCertidao = Val(SubNull(!numcertidao))
                dDataInscricao = IIf(IsNull(!datainscricao), CDate("01/01/1900"), !datainscricao)
                dDataVencto = !DataVencimento
                nNumExecFiscal = Val(SubNull(!numexecfiscal))
                nAnoExecFiscal = Val(SubNull(!anoexecfiscal))
                sNumExecFiscal = CStr(nNumExecFiscal) & "/" & CStr(nAnoExecFiscal)
                        
                If nNumExecFiscal = 0 Then
                    sql = "INSERT CDAs(IdDevedor,SetorDevedor,DtInscricao,NroCertidao,NroLivro,NroFolha,DtGeracao) values("
                    sql = sql & nCodReduz & ",'" & sTipoDivida & "','" & Format(dDataInscricao, "mm/dd/yyyy") & "'," & nNumCertidao & ","
                    sql = sql & nLivro & "," & nPag & ",'" & Format(Now, "mm/dd/yyyy") & "')"
                Else
                    sql = "INSERT CDAs(IdDevedor,SetorDevedor,DtInscricao,NroCertidao,NroLivro,NroFolha,NroOrdem,DtGeracao) values("
                    sql = sql & nCodReduz & ",'" & sTipoDivida & "','" & Format(dDataInscricao, "mm/dd/yyyy") & "'," & nNumCertidao & ","
                    sql = sql & nLivro & "," & nPag & ",'" & sNumExecFiscal & "','" & Format(Now, "mm/dd/yyyy") & "')"
                End If
                cnInt.Execute sql, rdExecDirect
                
                sql = "select @@identity as LastKey"
                Set RdoAux2 = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                nCDA = RdoAux2!lastkey
                RdoAux2.Close
                
                If nCodReduz < 500000 Then
                    sql = "INSERT Cadastro(IdCDA,SetorDevedor,Crc,Nome,Inscricao,CPFCnpj,RgInscrEstadual,LocalCep,LocalEndereco,LocalNumero,LocalComplemento,"
                    sql = sql & "LocalBairro,LocalCidade,LocalEstado,Quadra,Lote,EntregaCep,EntregaEndereco,EntregaNumero,EntregaComplemento,EntregaBairro,"
                    sql = sql & "EntregaCidade,EntregaEstado,DtGeracao) values("
                    sql = sql & nCDA & ",'" & sTipoDivida & "'," & nCodReduz & ",'" & SubNull(Mask(sNome)) & "','" & sInscricao & "','" & sCPF & "','" & sRG & "','"
                    sql = sql & sCep & "','" & sEndereco & "'," & nNumero & ",'" & Left(sComplemento, 30) & "','" & sBairro & "','" & sCidade & "','" & sUF & "','"
                    sql = sql & sQuadras & "','" & sLotes & "','" & sCepEntrega & "','" & sEnderecoEntrega & "'," & nNumEntrega & ",'" & Left(sComplementoEntrega, 30) & "','"
                    sql = sql & sBairroEntrega & "','" & sCidadeEntrega & "','" & sUFEntrega & "','" & Format(Now, "mm/dd/yyyy") & "')"
                    cnInt.Execute sql, rdExecDirect
                Else
                    sql = "select * from vwFullCidadao where codcidadao=" & nCodReduz
                    Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        sCPF = SubNull(!Cnpj)
                        If Trim(sCPF) = "" Then
                            sCPF = SubNull(!cpf)
                        End If
                        sql = "INSERT Partes(idCDA,Tipo,Crc,Nome,CpfCnpj,RgInscrEstadual,Cep,Endereco,Numero,Complemento,Bairro,Cidade,Estado,DtGeracao) Values("
                        sql = sql & nCDA & ",'Principal'," & !CodCidadao & ",'" & Mask(!nomecidadao) & "','" & sCPF & "','" & SubNull(!rg) & "','" & SubNull(!Cep) & "','"
                        sql = sql & Mask(SubNull(!Endereco)) & "'," & Val(SubNull(!NUMIMOVEL)) & ",'" & Mask(SubNull(!Complemento)) & "','" & Mask(SubNull(!DescBairro)) & "','"
                        sql = sql & Mask(SubNull(!descCidade)) & "','" & SubNull(!SiglaUF) & "','" & Format(Now, "mm/dd/yyyy") & "')"
                        cnInt.Execute sql, rdExecDirect
                       .Close
                    End With
                 End If
                
                If nCodReduz < 100000 Then 'cadastra os proprietarios e compromissarios
                    sql = "SELECT cadimob.codreduzido, proprietario.codcidadao, proprietario.tipoprop, vwFULLCIDADAO.nomecidadao, vwFULLCIDADAO.cpf,"
                    sql = sql & "vwFULLCIDADAO.cnpj, vwFULLCIDADAO.numimovel, vwFULLCIDADAO.complemento, vwFULLCIDADAO.siglauf, vwFULLCIDADAO.cep,"
                    sql = sql & "vwFULLCIDADAO.rg , vwFULLCIDADAO.orgao, vwFULLCIDADAO.DescBairro, vwFULLCIDADAO.desccidade, vwFULLCIDADAO.Endereco "
                    sql = sql & "FROM cadimob INNER JOIN proprietario ON cadimob.codreduzido = proprietario.codreduzido INNER JOIN vwFULLCIDADAO ON "
                    sql = sql & "proprietario.codcidadao = vwFULLCIDADAO.codcidadao where cadimob.codreduzido=" & nCodReduz
                    Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        Do Until .EOF
                            sCPF = SubNull(!Cnpj)
                            If Trim(sCPF) = "" Then
                                sCPF = SubNull(!cpf)
                            End If
                            sTipoProp = IIf(!tipoprop = "P", "Principal", "Compromissário")
                            sql = "INSERT Partes(idCDA,Tipo,Crc,Nome,CpfCnpj,RgInscrEstadual,Cep,Endereco,Numero,Complemento,Bairro,Cidade,Estado,DtGeracao) Values("
                            sql = sql & nCDA & ",'" & sTipoProp & "'," & nCodReduz & ",'" & Mask(!nomecidadao) & "','" & sCPF & "','" & SubNull(!rg) & "','" & SubNull(!Cep) & "','"
                            sql = sql & Mask(SubNull(!Endereco)) & "'," & Val(SubNull(!NUMIMOVEL)) & ",'" & Mask(SubNull(!Complemento)) & "','" & Mask(SubNull(!DescBairro)) & "','"
                            sql = sql & Mask(SubNull(!descCidade)) & "','" & SubNull(!SiglaUF) & "','" & Format(Now, "mm/dd/yyyy") & "')"
                            cnInt.Execute sql, rdExecDirect
                           .MoveNext
                        Loop
                       .Close
                    End With
                ElseIf nCodReduz >= 100000 And nCodReduz < 300000 Then   'cadastra os socios
                    sql = "SELECT * from vwmobiliarioproprietario where codmobiliario=" & nCodReduz
                    Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        Do Until .EOF
                            sCPF = SubNull(!Cnpj)
                            If Trim(sCPF) = "" Then
                                sCPF = SubNull(!cpf)
                            End If
                            sTipoProp = "Sócio"
                            sql = "INSERT Partes(idCDA,Tipo,Crc,Nome,CpfCnpj,RgInscrEstadual,Cep,Endereco,Numero,Complemento,Bairro,Cidade,Estado,DtGeracao) Values("
                            sql = sql & nCDA & ",'" & sTipoProp & "'," & nCodReduz & ",'" & Mask(!nomecidadao) & "','" & sCPF & "','" & SubNull(!rg) & "','" & SubNull(!Cep) & "','"
                            sql = sql & Mask(SubNull(!Endereco)) & "'," & Val(SubNull(!NUMIMOVEL)) & ",'" & Mask(SubNull(!Complemento)) & "','" & Mask(SubNull(!DescBairro)) & "','"
                            sql = sql & Mask(SubNull(!descCidade)) & "','" & SubNull(!SiglaUF) & "','" & Format(Now, "mm/dd/yyyy") & "')"
                            cnInt.Execute sql, rdExecDirect
                           .MoveNext
                        Loop
                       .Close
                    End With
                End If
                
               '***************************************************
               'Carrega Tributos
                sql = "SELECT debitotributo.codtributo,valortributo,abrevtributo FROM debitotributo INNER JOIN tributo ON debitotributo.codtributo = tributo.codtributo "
                sql = sql & "where codreduzido=" & nCodReduz & " and anoexercicio=" & nAno & " and codlancamento=" & nLanc & " and seqlancamento=" & nSeq & " and "
                sql = sql & "numparcela=" & nParc & " and codcomplemento=" & nCompl
                Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    Do Until .EOF
                        
                        sql = "INSERT CDADebitos(idCDA,CodTributo,Tributo,Exercicio,Lancamento,Seq,NroParcela,ComplParcela,DtVencimento,VlrOriginal,DtGeracao) values("
                        sql = sql & nCDA & "," & !CodTributo & ",'" & Mask(!abrevTributo) & "'," & nAno & "," & nLanc & "," & nSeq & "," & nParc & ","
                        sql = sql & nCompl & ",'" & Format(dDataVencto, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(!VALORTRIBUTO)) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
                        cnInt.Execute sql, rdExecDirect
                        
                       .MoveNext
                    Loop
                   .Close
                End With
                
                
               '**** Grava
                
                
                
               '***************************************************
                
                
                
               .MoveNext
            Loop
           .Close
        End With
        
        '***************************************************
        DoEvents
       .MoveNext
    Loop
   .Close
End With

Set xImovel = Nothing
MsgBox "fim"

End Sub

Private Sub cmdArq_Click()
Dim fName As String, cc As cCommonDlg
Set cc = New cCommonDlg
If cc.VBGetSaveFileName(fName, "", , "Texto[*.txt]", , sPathBin, "Local para salvar o arquivo", , , 0) Then
    txtArq.Text = fName
End If
End Sub

Private Sub cmdExec_Click()
Dim nTipoRel As Integer, nCadastro As Integer

If lstTipo.ListIndex = -1 Then
    MsgBox "Selecione o tipo de relatório", vbExclamation, "Atenção"
    Exit Sub
End If

If txtArq.Text = "" Then
    MsgBox "Selecione o arquivo.", vbExclamation, "Atenção"
    Exit Sub
End If


nTipoRel = Val(Left(lstTipo.Text, 2))
nCadastro = cmbCadastro.ItemData(cmbCadastro.ListIndex)

Select Case nTipoRel
    Case 1
      '  If MsgBox("Executar operação solicitada?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmação") = vbNo Then Exit Sub
 '       ExportaDebitos True, nCadastro
        'Relatorio2018
 
        Ajuizar_New
    Case 2
'        If MsgBox("Executar operação solicitada?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmação") = vbNo Then Exit Sub
'        ExportaDebitos False, nCadastro
     '   Ajuizar
    Case 3
      '  ImportaDebitosAjuizados
    Case 4
    '    ExportaSerasa nCadastro
    Case 5
    '    ExportarCDA
    Case 6
        'Protestar
        Protestar_New
End Select

End Sub



Private Sub cmdVerificar_Click()
Dim sql As String, RdoAux As rdoResultset, nCDAIndex As Integer, aCda() As tCDADebitoCorrecao, aCdaDebito() As tCDADebito
Dim x As Integer, qd As New rdoQuery

ConectaIntegrativa
Set qd.ActiveConnection = cn
qd.QueryTimeout = 0

ReDim aCda(0)
ReDim aCdaDebito(0)

'Verifica a tabela CDADebitosACorrigir
sql = "select * from cdadebitosacorrigir where dtleitura is null"
Set RdoAux = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        nCDAIndex = !idcdadebitos
        ReDim Preserve aCda(UBound(aCda) + 1)
        aCda(UBound(aCda)).idCdaIndex = nCDAIndex
        aCda(UBound(aCda)).DataCorrecao = Format(!DtCorrecao, "dd/mm/yyyy")
       .MoveNext
    Loop
   .Close
End With

'Carrega os Débitos que precisam ser corrigidos
For x = 1 To UBound(aCda)
    sql = "SELECT CDADebitos.idCDADebitos, CDADebitos.Exercicio, CDADebitos.Lancamento, CDADebitos.Seq, CDADebitos.NroParcela,CDADebitos.ComplParcela,"
    sql = sql & "CDAs.idDevedor, dbo.CDADebitos.CodTributo FROM CDAs INNER JOIN CDADebitos ON CDAs.idCDA = CDADebitos.idCDA Where CDADebitos.idCDADebitos = " & aCda(x).idCdaIndex
    Set RdoAux = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            ReDim Preserve aCdaDebito(UBound(aCdaDebito) + 1)
            aCdaDebito(UBound(aCdaDebito)).idCdaIndex = !idcdadebitos
            aCdaDebito(UBound(aCdaDebito)).nCodReduz = !iddevedor
            aCdaDebito(UBound(aCdaDebito)).nAno = !exercicio
            aCdaDebito(UBound(aCdaDebito)).nLanc = !lancamento
            aCdaDebito(UBound(aCdaDebito)).nSeq = !Seq
            aCdaDebito(UBound(aCdaDebito)).nParc = !nroparcela
            aCdaDebito(UBound(aCdaDebito)).nCompl = !complparcela
            aCdaDebito(UBound(aCdaDebito)).nCodTributo = !CodTributo
            aCdaDebito(UBound(aCdaDebito)).dDataCorrecao = Format(aCda(x).DataCorrecao, "dd/mm/yyyy")
           .MoveNext
        Loop
       .Close
    End With
Next

'atualiza os debitos da cda e grava na tabela CDADebitosCorrecao
For x = 1 To UBound(aCdaDebito)
    With aCdaDebito(x)
        On Error Resume Next
        RdoAux.Close
        On Error GoTo 0
                   
        qd.sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
        qd(0) = .nCodReduz
        qd(1) = .nCodReduz
        qd(2) = .nAno
        qd(3) = .nAno
        qd(4) = .nLanc
        qd(5) = .nLanc
        qd(6) = .nSeq
        qd(7) = .nSeq
        qd(8) = .nParc
        qd(9) = .nParc
        qd(10) = .nCompl
        qd(11) = .nCompl
        qd(12) = 1
        qd(13) = 99
        qd(14) = Format(.dDataCorrecao, "mm/dd/yyyy")
        qd(15) = "Integrativa"
        Set RdoAux = qd.OpenResultset(rdOpenKeyset)
        With RdoAux
            Do Until .EOF
                If !CodTributo = aCdaDebito(x).nCodTributo Then
                    sql = "insert CDADebitosCorrecao(idCdaDebitos,DtCorrecao,VlrOriginal,VlrCorrecao,VlrJuros,VlrMulta,DtGeracao) values("
                    sql = sql & aCdaDebito(x).idCdaIndex & ",'" & Format(aCdaDebito(x).dDataCorrecao, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(!VALORTRIBUTO)) & ","
                    sql = sql & Virg2Ponto(CStr(!valorcorrecao)) & "," & Virg2Ponto(CStr(!ValorJuros)) & "," & Virg2Ponto(CStr(!ValorMulta)) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
                    cnInt.Execute sql, rdExecDirect
                End If
               .MoveNext
            Loop
           .Close
        End With
    End With
         
    'Insere a data de leitura na tabela CDADebitosACorrigir para não corrigir novamente
     sql = "update CDADebitosACorrigir set dtLeitura='" & Format(Now, "mm/dd/yyyy") & "' where "
     sql = sql & "idCDADebitos=" & aCda(x).idCdaIndex
     cnInt.Execute sql, rdExecDirect
Next

cnInt.Close
Set qd = Nothing
MsgBox "Fim"

End Sub

Private Sub Form_Load()

Centraliza Me
PBar.Color = vbWhite

cmbCadastro.ListIndex = 0
lstTipo.AddItem "01-Exportar débitos ajuizados"
lstTipo.AddItem "02-Exportar débitos não ajuizados"
lstTipo.AddItem "03-Importar débitos ajuizados"
lstTipo.AddItem "04-Exportar para o Serasa"
lstTipo.AddItem "05-Preencher CDA's para checagem"
lstTipo.AddItem "06-Enviar para protesto"

lstTipo.ListIndex = 0

Set xImovel = New clsImovel

End Sub


Private Sub Ajuizar()
Dim sql As String, RdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, nCodReduz As Long, RdoAux3 As rdoResultset
Dim t As Integer, s As Integer, U As Integer, v As Integer, cmdComm As ADODB.Command, clsImovel As New clsImovel
Dim RdoDebito As rdoResultset, qd As New rdoQuery, Achou As Boolean, x As Integer, RdoTrib As rdoResultset
Dim Rs As ADODB.Recordset, strQuery As String, nNumExec As Long, nAnoExec As Integer, sNumExec As String

Dim sNomeCidadao As String, sInscricao As String, sFoneRes As String, sFoneCom As String, sCelular As String, sFoneContato As String, sEmail As String, nNumeroLocal As Integer
Dim sCPFCNPJ As String, nCodLogradouroLocal As Integer, sEnderecoLocal As String, sComplementoLocal As String, sNumeroLocal As String, nCodBairroLocal As Integer, sBairroLocal As String, sCidadeLocal As String
Dim sCEPLocal As String, sQuadra As String, sLote As String, sAtividade As String, sMatricula As String, sRGIE As String, nCodLogradouro As Integer, sUFLocal As String
Dim sEndereco As String, nNumero As Integer, sComplemento As String, nCodBairro As Integer, nCodCidade As Integer, sBairro As String, sCidade As String, sCep As String, sUF As String
Dim SetorDevedor As String, DtInscricao As String, NroCertidao As Integer, NroLivro As Integer, NroFolha As Integer, NroOrdem As Integer
Dim nCDA As Long, nAno As Integer, nLanc As Integer, nSeqLanc As Integer, nParc As Integer, nCompl As Integer, nCodTrib As Integer, sDescTrib As String
Dim dtVencto As String, nValorPrincipal As Double, nCodCidadao As Long, sClassificacao As String
Dim aCda() As typeCDA, aCdaDebito() As typeCDADebito, aParte() As Reg01, nValorTotal As Double

Exit Sub
Ocupado
ConectaIntegrativa
sql = "delete from cdadebitos"
'cnInt.Execute Sql, rdExecDirect

sql = "delete from cadastro"
'cnInt.Execute Sql, rdExecDirect

sql = "delete from partes"
'cnInt.Execute Sql, rdExecDirect

sql = "delete from cdas"
'cnInt.Execute Sql, rdExecDirect

sql = "delete from cdas_new"
'cnInt.Execute Sql, rdExecDirect

sql = "delete from cdadebitos_new"
'cnInt.Execute Sql, rdExecDirect


cmdExec.Enabled = False
Set qd.ActiveConnection = cn
qd.QueryTimeout = 0

If cmbCadastro.ListIndex = 0 Then
    MsgBox "selecione um cadastro"
    Exit Sub
End If
sql = "SELECT  DISTINCT CODREDUZIDO FROM DEBITOPARCELA WHERE DATAVENCIMENTO<='" & Format(mskDataVencto.Text, "mm/dd/yyyy") & "' AND STATUSLANC in (3,38,39,42,43) AND NUMPARCELA>0 AND DATAAJUIZA IS NULL AND DATAINSCRICAO IS NOT NULL "
If cmbCadastro.ListIndex = 1 Then
    'Sql = Sql & " AND CODREDUZIDO BETWEEN 9976 AND 40000 "
    '    Sql = Sql & " AND CODREDUZIDO = 8728 "
    sql = sql & " AND CODREDUZIDO < 40000 "
    sql = sql & " AND CODLANCAMENTO<>20 "
ElseIf cmbCadastro.ListIndex = 2 Then
    sql = sql & " AND CODREDUZIDO BETWEEN 100000 AND 300000 "
    'Sql = Sql & " AND CODREDUZIDO = 118151 "
    sql = sql & " AND CODLANCAMENTO<>20 "
ElseIf cmbCadastro.ListIndex = 3 Then
    sql = sql & " AND CODREDUZIDO =538525 "
    'Sql = Sql & " AND CODREDUZIDO BETWEEN 500000 AND 699999 "
    'Sql = Sql & " AND CODLANCAMENTO in(50,65,49,16,62,27,71,48,74) "
   sql = sql & " AND CODLANCAMENTO <>20 "
End If
sql = sql & " ORDER BY CODREDUZIDO"



'***********

Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nPos = 1
    nTot = .RowCount
    Do Until .EOF
        nCodReduz = !CODREDUZIDO
        If nCodReduz < 100000 Then
            SetorDevedor = "Imobiliário"
        ElseIf nCodReduz >= 100000 And nCodReduz < 500000 Then
            SetorDevedor = "Mobiliário"
        Else
            SetorDevedor = "Cidadão"
        End If
            
        If nPos Mod 10 = 0 Then
          '  GoTo fim
            DoEvents
            CallPb nPos, nTot
        End If
        
        '*** debito ***
        ReDim aParte(0)
        ReDim aCda(0)
        ReDim aCdaDebito(0)
        
        On Error Resume Next
        RdoDebito.Close
        On Error GoTo 0
        If nCadastro = 3 Then
            qd.sql = "{ Call spEXTRATOAJUIZARTAXA(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
        Else
            qd.sql = "{ Call spEXTRATOAJUIZAR(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
        End If
        qd(0) = nCodReduz
        qd(1) = nCodReduz
        qd(2) = 2019
        qd(3) = 2023
        qd(4) = 1
        qd(5) = 999
        qd(6) = 0
        qd(7) = 999
        qd(8) = 1
        qd(9) = 999
        qd(10) = 0
        qd(11) = 99
        qd(12) = 3
        qd(13) = 3
        qd(14) = Format(Now, "mm/dd/yyyy")
        qd(15) = "Integrativa"
        qd(16) = 0
        Set RdoDebito = qd.OpenResultset(rdOpenKeyset)
        With RdoDebito
            Do Until .EOF
'                If Year(!DataVencimento) < 2015 Then
'                    GoTo proximo
'                End If
              '  If !ValorJuros = 0 Then
                    'MsgBox "teste"
               ' End If
                U = UBound(aCda)
                Achou = False
                For x = 1 To U
                    If Val(aCda(x).nAno) = !AnoExercicio And Val(aCda(x).nLanc) = !CodLancamento And Val(aCda(x).nSeq) = !SeqLancamento And _
                       Val(aCda(x).nParc) = !NumParcela And Val(aCda(x).nCompl) = !CODCOMPLEMENTO Then
                        Achou = True
                        Exit For
                    End If
                Next
                
                If Not Achou Then
                    ReDim Preserve aCda(UBound(aCda) + 1)
                    U = UBound(aCda)
                    aCda(U).nAno = !AnoExercicio
                    aCda(U).nLanc = !CodLancamento
                    aCda(U).nSeq = !SeqLancamento
                    aCda(U).nParc = !NumParcela
                    aCda(U).nCompl = !CODCOMPLEMENTO
                    On Error Resume Next
                    aCda(U).nNumCertidao = Val(SubNull(!certidao))
                    On Error GoTo 0
                    aCda(U).nNumPagina = Val(SubNull(!PAGINA))
                    aCda(U).nNumLivro = Val(SubNull(!NUMLIVRO))
                    aCda(U).dDataInscricao = !datainscricao
                End If
                
                
                ReDim Preserve aCdaDebito(UBound(aCdaDebito) + 1)
                U = UBound(aCdaDebito)
                aCdaDebito(U).nAno = !AnoExercicio
                aCdaDebito(U).nLanc = !CodLancamento
                aCdaDebito(U).nSeq = !SeqLancamento
                aCdaDebito(U).nParc = !NumParcela
                aCdaDebito(U).nCompl = !CODCOMPLEMENTO
                aCdaDebito(U).nCodTributo = !CodTributo
                aCdaDebito(U).sDescTributo = !abrevTributo
                aCdaDebito(U).nPrincipal = !VALORTRIBUTO
                aCdaDebito(U).nMulta = !ValorMulta
                aCdaDebito(U).nJuros = !ValorJuros
                aCdaDebito(U).nCorrecao = !valorcorrecao
                aCdaDebito(U).nMulta = !ValorMulta
                aCdaDebito(U).nJuros = !ValorJuros
                aCdaDebito(U).nCorrecao = !valorcorrecao
                aCdaDebito(U).nTotal = !ValorTotal
                aCdaDebito(U).dDataVencto = !DataVencimento
PROXIMODEBITO:
                DoEvents
               .MoveNext
            Loop
           .Close
        End With

        If UBound(aCda) = 0 Then GoTo Proximo
        'limita em 1301 reais
        nValorTotal = 0
        For x = 1 To UBound(aCdaDebito)
            With aCdaDebito(x)
                nValorTotal = nValorTotal + .nTotal
            End With
        Next
        If nValorTotal < 1301 Then
            GoTo Proximo
        End If
        
        
        For x = 1 To UBound(aCda)
            With aCda(x)
                sql = "insert CDAs(idDevedor, SetorDevedor, DtInscricao, NroCertidao, NroLivro, NroFolha,  DtGeracao) values("
                sql = sql & nCodReduz & ",'" & SetorDevedor & "','" & Format(.dDataInscricao, "mm/dd/yyyy") & "'," & .nNumCertidao & "," & .nNumLivro & ","
                sql = sql & .nNumPagina & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "')"
                cnInt.Execute sql, rdExecDirect
            
                sql = "select @@identity as LastKey"
                Set RdoAux3 = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                nCDA = RdoAux3!lastkey
               .nCDA = nCDA
                RdoAux3.Close
            
                For v = 1 To UBound(aCdaDebito)
                    If aCdaDebito(v).nAno = .nAno And aCdaDebito(v).nLanc = .nLanc And aCdaDebito(v).nSeq = .nSeq And _
                       aCdaDebito(v).nParc = .nParc And aCdaDebito(v).nCompl = .nCompl Then
                        aCdaDebito(v).nCDA = .nCDA
                    End If
                Next
            End With
        Next
        
        For x = 1 To UBound(aCdaDebito)
            With aCdaDebito(x)
                sql = "insert CDADebitos(idCDA, CodTributo, Tributo, Exercicio, Lancamento, Seq, NroParcela, ComplParcela, DtVencimento, vlrOriginal,vlrMultas,vlrjuros,vlrcorrecao, DtGeracao) values("
                sql = sql & .nCDA & "," & .nCodTributo & ",'" & Mask(.sDescTributo) & "'," & .nAno & "," & .nLanc & "," & .nSeq & "," & .nParc & "," & .nCompl & ",'"
                sql = sql & Format(.dDataVencto, "mm/dd/yyyy") & "'," & Virg2Ponto(Format(.nPrincipal, "#0.00")) & "," & Virg2Ponto(Format(.nMulta, "#0.00")) & "," & Virg2Ponto(Format(.nJuros, "#0.00")) & "," & Virg2Ponto(Format(.nCorrecao, "#0.00")) & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "')"
                cnInt.Execute sql, rdExecDirect
            End With
        Next
               
       '*** dados cadastrais
        If nCodReduz < 100000 Then
            sql = "select * from vwfullimovel2 where codreduzido=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If .RowCount = 0 Then GoTo Proximo
                sNome = !nomecidadao
                sInscricao = !Inscricao
                sFoneRes = Left(RetornaNumero(SubNull(!telefone)), 15)
                sFoneCom = ""
                sCelular = ""
                sFoneContato = ""
                sEmail = Left(SubNull(!Email), 100)
                If Trim(SubNull(!Cnpj)) <> "" Then
                    sCPFCNPJ = RetornaNumero(!Cnpj)
                Else
                    If Trim(SubNull(!cpf)) <> "" Then
                        sCPFCNPJ = RetornaNumero(!cpf)
                    Else
                        sCPFCNPJ = ""
                    End If
                End If
                
                xImovel.RetornaEndereco nCodReduz, Imobiliario, Localizacao
                nCodLogradouroLocal = xImovel.CodLogradouro
                
                sEnderecoLocal = xImovel.Endereco
                nNumeroLocal = xImovel.Numero
                sComplementoLocal = xImovel.Complemento
                nCodBairroLocal = xImovel.CodBairro
                sBairroLocal = xImovel.Bairro
                sCidadeLocal = xImovel.Cidade
                sUFLocal = xImovel.UF
                sCEPLocal = xImovel.Cep
                If Val(nCodBairroLocal) = 999 Then
                    nCodBairroLocal = 0
                    sBairroLocal = ""
                End If
                
                sQuadra = !Quadra
                sLote = !Lote
                sAtividade = ""
                sMatricula = Val(SubNull(!NumMat))
                
                sRGIE = RetornaNumero(Trim(SubNull(!rg)))
                If !Ee_TipoEnd = 0 Then 'imovel
                    xImovel.RetornaEndereco nCodReduz, Imobiliario, Localizacao
                    nCodLogradouro = xImovel.CodLogradouro
                    sEndereco = xImovel.Endereco
                    nNumero = xImovel.Numero
                    sComplemento = xImovel.Complemento
                    nCodBairro = xImovel.CodBairro
                    sBairro = xImovel.Bairro
                    nCodCidade = xImovel.CodCidade
                    sCidade = xImovel.Cidade
                    sCep = xImovel.Cep
                    sUF = xImovel.UF
                ElseIf !Ee_TipoEnd = 1 Then 'proprietario
                    xImovel.RetornaEndereco nCodReduz, Imobiliario, cadastrocidadao
                    nCodLogradouro = Val(xImovel.CodLogradouro)
                    sEndereco = xImovel.Endereco
                    nNumero = xImovel.Numero
                    sComplemento = xImovel.Complemento
                    nCodBairro = Val(xImovel.CodBairro)
                    sBairro = xImovel.Bairro
                    nCodCidade = xImovel.CodCidade
                    sCidade = xImovel.Cidade
                    sCep = xImovel.Cep
                    sUF = xImovel.UF
                Else 'entrega
                    xImovel.RetornaEndereco nCodReduz, Imobiliario, Entrega
                    nCodLogradouro = xImovel.CodLogradouro
                    sEndereco = xImovel.Endereco
                    nNumero = xImovel.Numero
                    sComplemento = xImovel.Complemento
                    nCodBairro = Val(xImovel.CodBairro)
                    sBairro = xImovel.Bairro
                    nCodCidade = xImovel.CodCidade
                    sCidade = xImovel.Cidade
                    sCep = xImovel.Cep
                    sUF = xImovel.UF
                End If
               .Close
            End With
        ElseIf nCodReduz >= 100000 And nCodReduz < 500000 Then
            sql = "select * from vwfullempresa3 where codigomob=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If .RowCount = 0 Then GoTo Proximo
                sInscricao = ""
                sNome = !RazaoSocial
                sFoneRes = ""
                sFoneCom = ""
                sCelular = ""
                sFoneContato = Left(RetornaNumero(SubNull(!fonecontato)), 15)
                sEmail = Left(SubNull(!emailcontato), 100)
                If Trim(SubNull(!Cnpj)) <> "" Then
                    sCPFCNPJ = RetornaNumero(!Cnpj)
                Else
                    If Trim(SubNull(!cpf)) <> "" Then
                        sCPFCNPJ = RetornaNumero(!cpf)
                    Else
                        sCPFCNPJ = ""
                    End If
                End If
                sRGIE = RetornaNumero(Trim(SubNull(!inscestadual)))
                If sRGIE = "" Then
                    sRGIE = Trim(SubNull(!rg) & " " & SubNull(!Orgao))
                End If
                
                xImovel.RetornaEndereco nCodReduz, Mobiliario, Localizacao
                nCodLogradouroLocal = xImovel.CodLogradouro
                sEnderecoLocal = xImovel.Endereco
                nNumeroLocal = xImovel.Numero
                sComplementoLocal = xImovel.Complemento
                nCodBairroLocal = Val(xImovel.CodBairro)
                sBairroLocal = xImovel.Bairro
                sCidadeLocal = xImovel.Cidade
                sUFLocal = xImovel.UF
                sCEPLocal = xImovel.Cep
                If Val(nCodBairroLocal) = 999 Then
                    nCodBairroLocal = 0
                    sBairroLocal = ""
                End If
                
                sEndereco = SubNull(!eenomelogr)
                If sEndereco = "" Then
                    'local da empresa
                    xImovel.RetornaEndereco nCodReduz, Mobiliario, Localizacao
                    nCodLogradouro = xImovel.CodLogradouro
                    sEndereco = xImovel.Endereco
                    nNumero = xImovel.Numero
                    sComplemento = xImovel.Complemento
                    nCodBairro = Val(xImovel.CodBairro)
                    sBairro = xImovel.Bairro
                    nCodCidade = xImovel.CodCidade
                    sCidade = xImovel.Cidade
                    sCep = xImovel.Cep
                    sUF = xImovel.UF
                Else
'                   'endereco entrega
                    xImovel.RetornaEndereco nCodReduz, Mobiliario, Entrega
                    nCodLogradouro = xImovel.CodLogradouro
                    sEndereco = xImovel.Endereco
                    nNumero = xImovel.Numero
                    sComplemento = xImovel.Complemento
                    nCodBairro = Val(xImovel.CodBairro)
                    sBairro = xImovel.Bairro
                    nCodCidade = xImovel.CodCidade
                    sCidade = xImovel.Cidade
                    sCep = xImovel.Cep
                    sUF = xImovel.UF
                End If
               
                sQuadra = ""
                sLote = ""
                sAtividade = Left(SubNull(!ativextenso), 80)
                sMatricula = ""
               
               .Close
            End With
        Else
            sql = "select * from vwfullcidadao where codcidadao=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If .RowCount = 0 Then GoTo Proximo
                sInscricao = ""
                sNome = !nomecidadao
                sFoneRes = Left(RetornaNumero(SubNull(!telefone)), 15)
                sFoneCom = ""
                sCelular = ""
                sFoneContato = ""
                sEmail = Left(SubNull(!Email), 100)
                If Trim(SubNull(!Cnpj)) <> "" Then
                    sCPFCNPJ = RetornaNumero(!Cnpj)
                Else
                    If Trim(SubNull(!cpf)) <> "" Then
                        sCPFCNPJ = RetornaNumero(!cpf)
                    Else
                        sCPFCNPJ = ""
                    End If
                End If
                sRGIE = RetornaNumero(Trim(Left(SubNull(!rg), 14)))
                xImovel.RetornaEndereco nCodReduz, cidadao, cadastrocidadao
                nCodLogradouro = Val(xImovel.CodLogradouro)
                sEndereco = SubNull(xImovel.Endereco)
                nNumero = SubNull(xImovel.Numero)
                sComplemento = xImovel.Complemento
                nCodBairro = Val(SubNull(xImovel.CodBairro))
                sBairro = SubNull(xImovel.Bairro)
                If nCodBairro = 999 Then
                    nCodBairro = 0
                    sBairro = ""
                End If
                nCodCidade = Val(SubNull(xImovel.CodCidade))
                sCidade = SubNull(xImovel.Cidade)
                sUFLocal = xImovel.UF
                sCep = RetornaNumero(SubNull(xImovel.Cep))
                sUF = SubNull(xImovel.UF)
               
                nCodLogradouroLocal = Val(xImovel.CodLogradouro)
                sEnderecoLocal = SubNull(xImovel.Endereco)
                nNumeroLocal = SubNull(xImovel.Numero)
                sComplementoLocal = SubNull(xImovel.Complemento)
                nCodBairroLocal = Val(SubNull(xImovel.CodBairro))
                sBairroLocal = SubNull(xImovel.Bairro)
                If Val(nCodBairroLocal) = 999 Then
                    nCodBairroLocal = 0
                    sBairroLocal = ""
                End If
                sCidadeLocal = xImovel.Cidade
                sCEPLocal = RetornaNumero(SubNull(xImovel.Cep))
                
                sQuadra = ""
                sLote = ""
                sAtividade = ""
                sMatricula = ""
               
               .Close
            End With
        End If
                
        For x = 1 To UBound(aCda)
            With aCda(x)
                sql = "INSERT Cadastro(IdCDA,SetorDevedor,Crc,Nome,Inscricao,CPFCnpj,RgInscrEstadual,LocalCep,LocalEndereco,LocalNumero,LocalComplemento,"
                sql = sql & "LocalBairro,LocalCidade,LocalEstado,Quadra,Lote,EntregaCep,EntregaEndereco,EntregaNumero,EntregaComplemento,EntregaBairro,"
                sql = sql & "EntregaCidade,EntregaEstado,DtGeracao) values("
                sql = sql & .nCDA & ",'" & SetorDevedor & "'," & nCodReduz & ",'" & Mask(Left(SubNull(sNome), 80)) & "','" & sInscricao & "','" & sCPFCNPJ & "','" & sRGIE & "','"
                sql = sql & sCep & "','" & sEnderecoLocal & "'," & nNumeroLocal & ",'" & Mask(Left(sComplementoLocal, 50)) & "','" & sBairroLocal & "','" & sCidadeLocal & "','" & sUFLocal & "','"
                sql = sql & sQuadra & "','" & sLote & "','" & sCep & "','" & sEndereco & "'," & nNumero & ",'" & Mask(Left(sComplemento, 50)) & "','"
                sql = sql & sBairro & "','" & Mask(sCidade) & "','" & sUF & "','" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "')"
                cnInt.Execute sql, rdExecDirect
            End With
        Next
        
        
        '***** PARTES **********************************

        If nCodReduz < 100000 Then
            sql = "SELECT proprietario.codcidadao, cidadao.nomecidadao, cidadao.cnpj,cidadao.cpf,cidadao.rg,cidadao.telefone, cidadao.email FROM proprietario INNER JOIN "
            sql = sql & "cidadao ON proprietario.codcidadao = cidadao.codcidadao Where codreduzido=" & nCodReduz & " and principal=0"
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                Do Until .EOF
                    ReDim Preserve aParte(UBound(aParte) + 1)
                    U = UBound(aParte)
                    xImovel.RetornaEndereco !CodCidadao, cidadao, cadastrocidadao
                    aParte(U).sNomeSocio = RdoAux2!nomecidadao
                    If Trim(SubNull(RdoAux2!Cnpj)) <> "" Then
                        aParte(U).sCPFCNPJSocio = RetornaNumero(RdoAux2!Cnpj)
                    Else
                        If Trim(SubNull(RdoAux2!cpf)) <> "" Then
                            aParte(U).sCPFCNPJSocio = RetornaNumero(RdoAux2!cpf)
                        Else
                            aParte(U).sCPFCNPJSocio = ""
                        End If
                    End If
                    aParte(U).sRGIE = RetornaNumero(Trim(Left(SubNull(RdoAux2!rg), 14)))
                    aParte(U).sCodLogradouro = xImovel.CodLogradouro
                    aParte(U).sEndereco = xImovel.Endereco
                    aParte(U).sNumero = xImovel.Numero
                    aParte(U).sComplemento = xImovel.Complemento
                    aParte(U).sCodBairro = xImovel.CodBairro
                    aParte(U).sBairro = xImovel.Bairro
                    aParte(U).sCodCidade = xImovel.CodCidade
                    aParte(U).sCidade = xImovel.Cidade
                    aParte(U).sCepSocio = RetornaNumero(xImovel.Cep)
                    aParte(U).sEstadoSocio = xImovel.UF
                    aParte(U).sClassificao = "Compromissário"
                    aParte(U).sFoneRes = Left(RetornaNumero(SubNull(RdoAux2!telefone)), 15)
                    aParte(U).sFoneCom = ""
                    aParte(U).sCelular = ""
                    aParte(U).sFoneContato = ""
                    aParte(U).sEmail = Left(SubNull(RdoAux2!Email), 100)
                   .MoveNext
                Loop
               .Close
            End With
        ElseIf nCodReduz >= 100000 And nCodReduz < 500000 Then
            sql = "SELECT mobiliarioproprietario.codcidadao, cidadao.nomecidadao, cidadao.cnpj,cidadao.cpf,cidadao.rg,cidadao.telefone, cidadao.email FROM mobiliarioproprietario INNER JOIN "
            sql = sql & "cidadao ON mobiliarioproprietario.codcidadao = cidadao.codcidadao Where codmobiliario=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                Do Until .EOF
                    ReDim Preserve aParte(UBound(aParte) + 1)
                    U = UBound(aParte)
                    xImovel.RetornaEndereco !CodCidadao, cidadao, cadastrocidadao
                    aParte(U).sNomeSocio = RdoAux2!nomecidadao
                    If Trim(SubNull(RdoAux2!Cnpj)) <> "" Then
                        aParte(U).sCPFCNPJSocio = RetornaNumero(RdoAux2!Cnpj)
                    Else
                        If Trim(SubNull(RdoAux2!cpf)) <> "" Then
                            aParte(U).sCPFCNPJSocio = RetornaNumero(RdoAux2!cpf)
                        Else
                            aParte(U).sCPFCNPJSocio = ""
                        End If
                    End If
                    aParte(U).sRGIE = RetornaNumero(Trim(Left(SubNull(RdoAux2!rg), 14)))
                    aParte(U).sCodLogradouro = xImovel.CodLogradouro
                    aParte(U).sEndereco = xImovel.Endereco
                    aParte(U).sNumero = xImovel.Numero
                    aParte(U).sComplemento = xImovel.Complemento
                    aParte(U).sCodBairro = xImovel.CodBairro
                    aParte(U).sBairro = xImovel.Bairro
                    aParte(U).sCodCidade = xImovel.CodCidade
                    aParte(U).sCidade = xImovel.Cidade
                    aParte(U).sCepSocio = RetornaNumero(xImovel.Cep)
                    aParte(U).sEstadoSocio = xImovel.UF
                    aParte(U).sClassificao = "Sócio"
                    aParte(U).sFoneRes = Left(RetornaNumero(SubNull(RdoAux2!telefone)), 15)
                    aParte(U).sFoneCom = ""
                    aParte(U).sCelular = ""
                    aParte(U).sFoneContato = ""
                    aParte(U).sEmail = Left(SubNull(RdoAux2!Email), 100)
                   .MoveNext
                Loop
                RdoAux2.Close
            End With
        End If
        
        For x = 1 To UBound(aCda)
            For U = 1 To UBound(aParte)
                With aParte(U)
                    sql = "INSERT Partes(idCDA,Tipo,Crc,Nome,CpfCnpj,RgInscrEstadual,Cep,Endereco,Numero,Complemento,Bairro,Cidade,Estado,DtGeracao) Values("
                    sql = sql & aCda(x).nCDA & ",'" & .sClassificao & "'," & nCodReduz & ",'" & Mask(.sNomeSocio) & "','" & .sCPFCNPJSocio & "','" & .sRGIE & "','" & .sCepSocio & "','"
                    sql = sql & Mask(.sEndereco) & "'," & Val(.sNumero) & ",'" & Mask(.sComplemento) & "','" & Mask(.sBairro) & "','"
                    sql = sql & Mask(.sCidade) & "','" & SubNull(.sEstadoSocio) & "','" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "')"
                    cnInt.Execute sql, rdExecDirect
                End With
            Next
        Next
        
Proximo:
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
Fim:
   .Close
End With
cmdExec.Enabled = True
cnInt.Close
Liberado
MsgBox "Exportação finalizada.", vbInformation

End Sub

Private Sub Ajuizar_New()
Dim sql As String, RdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, nCodReduz As Long, RdoAux3 As rdoResultset
Dim t As Integer, s As Integer, U As Integer, v As Integer, cmdComm As ADODB.Command, clsImovel As New clsImovel
Dim RdoDebito As rdoResultset, qd As New rdoQuery, Achou As Boolean, x As Integer, RdoTrib As rdoResultset
Dim Rs As ADODB.Recordset, strQuery As String, nNumExec As Long, nAnoExec As Integer, sNumExec As String

Dim sNomeCidadao As String, sInscricao As String, sFoneRes As String, sFoneCom As String, sCelular As String, sFoneContato As String, sEmail As String, nNumeroLocal As Integer
Dim sCPFCNPJ As String, nCodLogradouroLocal As Integer, sEnderecoLocal As String, sComplementoLocal As String, sNumeroLocal As String, nCodBairroLocal As Integer, sBairroLocal As String, sCidadeLocal As String
Dim sCEPLocal As String, sQuadra As String, sLote As String, sAtividade As String, sMatricula As String, sRGIE As String, nCodLogradouro As Integer, sUFLocal As String
Dim sEndereco As String, nNumero As Integer, sComplemento As String, nCodBairro As Integer, nCodCidade As Integer, sBairro As String, sCidade As String, sCep As String, sUF As String
Dim SetorDevedor As String, DtInscricao As String, NroCertidao As Integer, NroLivro As Integer, NroFolha As Integer, NroOrdem As Integer
Dim nCDA As Long, nAno As Integer, nLanc As Integer, nSeqLanc As Integer, nParc As Integer, nCompl As Integer, nCodTrib As Integer, sDescTrib As String
Dim dtVencto As String, nValorPrincipal As Double, nCodCidadao As Long, sClassificacao As String
Dim aCda() As typeCDANew, aCdaDebito() As typeCDADebito, aParte() As Reg01
Dim nPrincipal As Double, nMulta As Double, nJuros As Double, nCorrecao As Double, nTotal As Double, nValorTotal As Double


Ocupado
ConectaIntegrativa
sql = "delete from cdadebitos"
'cnInt.Execute Sql, rdExecDirect


sql = "delete from cadastro"
'cnInt.Execute Sql, rdExecDirect

sql = "delete from partes"
'cnInt.Execute Sql, rdExecDirect

sql = "delete from cdas"
'cnInt.Execute Sql, rdExecDirect

cmdExec.Enabled = False
Set qd.ActiveConnection = cn
qd.QueryTimeout = 0

If cmbCadastro.ListIndex = 0 Then
    MsgBox "selecione um cadastro"
    Exit Sub
End If
sql = "SELECT  DISTINCT CODREDUZIDO FROM DEBITOPARCELA WHERE DATAVENCIMENTO<='" & Format(mskDataVencto.Text, "mm/dd/yyyy") & "' AND STATUSLANC in (3,38,39,42,43) AND NUMPARCELA>0 AND DATAAJUIZA IS NULL AND DATAINSCRICAO IS NOT NULL "
If cmbCadastro.ListIndex = 1 Then
    'Sql = Sql & " AND CODREDUZIDO = 4026 "
    sql = sql & " AND CODREDUZIDO = 101061 "
    'Sql = Sql & " AND CODREDUZIDO < 5000 "
    sql = sql & " AND CODLANCAMENTO<>20 "
ElseIf cmbCadastro.ListIndex = 2 Then
    'Sql = Sql & " AND CODREDUZIDO BETWEEN 100000 AND 300000 "
    sql = sql & " AND CODREDUZIDO = 120907 "
    sql = sql & " AND CODLANCAMENTO<>20 "
ElseIf cmbCadastro.ListIndex = 3 Then
    sql = sql & " AND CODREDUZIDO =552050 "
    'Sql = Sql & " AND CODREDUZIDO BETWEEN 500000 AND 699999 "
    'Sql = Sql & " AND CODLANCAMENTO in(50,65,49,16,62,27,71,48,74) "
   sql = sql & " AND CODLANCAMENTO <>20 "
End If
sql = sql & " ORDER BY CODREDUZIDO"



'***********

Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nPos = 1
    nTot = .RowCount
    Do Until .EOF
        nCodReduz = !CODREDUZIDO
        If nCodReduz < 100000 Then
            SetorDevedor = "Imobiliário"
        ElseIf nCodReduz >= 100000 And nCodReduz < 500000 Then
            SetorDevedor = "Mobiliário"
        Else
            SetorDevedor = "Cidadão"
        End If
            
        If nPos Mod 10 = 0 Then
          '  GoTo fim
            DoEvents
            CallPb nPos, nTot
        End If
        
        '*** debito ***
        ReDim aParte(0)
        ReDim aCda(0)
        ReDim aCdaDebito(0)
        
        On Error Resume Next
        RdoDebito.Close
        On Error GoTo 0
        If nCadastro = 3 Then
            qd.sql = "{ Call spEXTRATOAJUIZARTAXA(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
        Else
            qd.sql = "{ Call spEXTRATOAJUIZAR(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
        End If
        qd(0) = nCodReduz
        qd(1) = nCodReduz
        qd(2) = 2000
        qd(3) = 2024
        qd(4) = 1
        qd(5) = 999
        qd(6) = 0
        qd(7) = 999
        qd(8) = 1
        qd(9) = 999
        qd(10) = 0
        qd(11) = 99
        qd(12) = 3
        qd(13) = 3
        qd(14) = Format(Now, "mm/dd/yyyy")
        qd(15) = "Integrativa"
        qd(16) = 0
        Set RdoDebito = qd.OpenResultset(rdOpenKeyset)
        With RdoDebito
            Do Until .EOF
'                If Year(!DataVencimento) < 2015 Then
'                    GoTo proximo
'                End If
              '  If !ValorJuros = 0 Then
                    'MsgBox "teste"
               ' End If
                U = UBound(aCda)
                Achou = False
                For x = 1 To U
'                    If Val(aCda(x).nAno) = !AnoExercicio And Val(aCda(x).nLanc) = !CodLancamento And Val(aCda(x).nSeq) = !SeqLancamento And _
'                       Val(aCda(x).nParc) = !NumParcela And Val(aCda(x).nCompl) = !CODCOMPLEMENTO And Val(aCda(x).nNumCertidao) = Val(SubNull(!certidao)) Then
                    If Val(aCda(x).nNumCertidao) = Val(SubNull(!certidao)) Then
                        Achou = True
                        Exit For
                    End If
                Next
                
                If Not Achou Then
                    ReDim Preserve aCda(UBound(aCda) + 1)
                    U = UBound(aCda)
                    aCda(U).nAno = !AnoExercicio
                    aCda(U).nLanc = !CodLancamento
                    aCda(U).nSeq = !SeqLancamento
                    aCda(U).nParc = !NumParcela
                    aCda(U).nCompl = !CODCOMPLEMENTO
                    On Error Resume Next
                    aCda(U).nNumCertidao = Val(SubNull(!certidao))
                    On Error GoTo 0
                    aCda(U).nNumPagina = Val(SubNull(!PAGINA))
                    aCda(U).nNumLivro = Val(SubNull(!NUMLIVRO))
                    aCda(U).dDataInscricao = !datainscricao
                End If
                
                
                ReDim Preserve aCdaDebito(UBound(aCdaDebito) + 1)
                U = UBound(aCdaDebito)
                aCdaDebito(U).nAno = !AnoExercicio
                aCdaDebito(U).nLanc = !CodLancamento
                aCdaDebito(U).nSeq = !SeqLancamento
                aCdaDebito(U).nParc = !NumParcela
                aCdaDebito(U).nCompl = !CODCOMPLEMENTO
                aCdaDebito(U).nCodTributo = !CodTributo
                aCdaDebito(U).sDescTributo = !abrevTributo
                aCdaDebito(U).nPrincipal = !VALORTRIBUTO
                aCdaDebito(U).nMulta = !ValorMulta
                aCdaDebito(U).nJuros = !ValorJuros
                aCdaDebito(U).nCorrecao = !valorcorrecao
                aCdaDebito(U).nMulta = !ValorMulta
                aCdaDebito(U).nJuros = !ValorJuros
                aCdaDebito(U).nCorrecao = !valorcorrecao
                aCdaDebito(U).nTotal = !ValorTotal
                aCdaDebito(U).nNumCertidao = Val(SubNull(!certidao))
                aCdaDebito(U).nFolha = Val(SubNull(!PAGINA))
                aCdaDebito(U).nLivro = Val(SubNull(!NUMLIVRO))
                aCdaDebito(U).dDataVencto = !DataVencimento
PROXIMODEBITO:
                DoEvents
               .MoveNext
            Loop
           .Close
        End With

        If UBound(aCda) = 0 Then GoTo Proximo
        'limita em 1301 reais
        nValorTotal = 0
        For x = 1 To UBound(aCdaDebito)
            With aCdaDebito(x)
                nValorTotal = nValorTotal + .nTotal
            End With
        Next
        If nValorTotal < 1301 Then
            GoTo Proximo
        End If
        
        
        For x = 1 To UBound(aCda)
        
            With aCda(x)
                sql = "insert CDAs (idDevedor, SetorDevedor, DtInscricao, NroCertidao,nrolivro, vlrOriginal, vlrjuros,vlrmultas,vlrcorrecao,vlrtotal,  DtGeracao) values("
                sql = sql & nCodReduz & ",'" & SetorDevedor & "','" & Format(.dDataInscricao, "mm/dd/yyyy") & "'," & .nNumCertidao & "," & .nNumLivro & "," & Virg2Ponto(CStr(.nPrincipal)) & ","
                sql = sql & Virg2Ponto(CStr(.nJuros)) & "," & Virg2Ponto(CStr(.nMulta)) & "," & Virg2Ponto(CStr(.nCorrecao)) & "," & Virg2Ponto(CStr(.nTotal)) & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "')"
                cnInt.Execute sql, rdExecDirect
            
                sql = "select @@identity as LastKey"
                Set RdoAux3 = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                nCDA = RdoAux3!lastkey
               .nCDA = nCDA
                RdoAux3.Close

                nPrincipal = 0: nJuros = 0: nMulta = 0: nCorrecao = 0: nTotal = 0
                For v = 1 To UBound(aCdaDebito)
'                    If aCdaDebito(v).nAno = .nAno And aCdaDebito(v).nLanc = .nLanc And aCdaDebito(v).nSeq = .nSeq And _
'                       aCdaDebito(v).nParc = .nParc And aCdaDebito(v).nCompl = .nCompl Then
                    If aCdaDebito(v).nNumCertidao = .nNumCertidao Then
                        aCdaDebito(v).nCDA = .nCDA
                        For z = 1 To UBound(aCda)
                            If aCda(z).nCDA = .nCDA Then
                                nPrincipal = nPrincipal + aCdaDebito(v).nPrincipal
                                nJuros = nJuros + aCdaDebito(v).nJuros
                                nMulta = nMulta + aCdaDebito(v).nMulta
                                nCorrecao = nCorrecao + aCdaDebito(v).nCorrecao
                            End If
                        Next
                        nTotal = nPrincipal + nJuros + nMulta + nCorrecao
                    End If
                Next
               
                sql = "update cdas set vlroriginal=" & Virg2Ponto(CStr(nPrincipal)) & ",vlrmultas = " & Virg2Ponto(CStr(nMulta))
                sql = sql & ",vlrjuros=" & Virg2Ponto(CStr(nJuros)) & ",vlrcorrecao = " & Virg2Ponto(CStr(nCorrecao)) & ",vlrtotal = " & Virg2Ponto(CStr(nTotal))
                sql = sql & " where idcda=" & .nCDA
                cnInt.Execute sql, rdExecDirect
            End With
        Next
        
        For x = 1 To UBound(aCdaDebito)
            With aCdaDebito(x)
                sql = "insert CDADebitos(idCDA, CodTributo, Tributo, Exercicio, Lancamento, Seq, NroParcela, ComplParcela, DtVencimento, vlrOriginal,vlrMultas,vlrjuros,vlrcorrecao,nroLivro,nroFolha, DtGeracao) values("
                sql = sql & .nCDA & "," & .nCodTributo & ",'" & Mask(.sDescTributo) & "'," & .nAno & "," & .nLanc & "," & .nSeq & "," & .nParc & "," & .nCompl & ",'"
                sql = sql & Format(.dDataVencto, "mm/dd/yyyy") & "'," & Virg2Ponto(Format(.nPrincipal, "#0.00")) & "," & Virg2Ponto(Format(.nMulta, "#0.00")) & "," & Virg2Ponto(Format(.nJuros, "#0.00")) & ","
                sql = sql & Virg2Ponto(Format(.nCorrecao, "#0.00")) & "," & .nLivro & "," & .nFolha & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "')"
                cnInt.Execute sql, rdExecDirect
            End With
        Next

       '*** dados cadastrais
        If nCodReduz < 100000 Then
            sql = "select * from vwfullimovel2 where codreduzido=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If .RowCount = 0 Then GoTo Proximo
                sNome = !nomecidadao
                sInscricao = !Inscricao
                sFoneRes = Left(RetornaNumero(SubNull(!telefone)), 15)
                sFoneCom = ""
                sCelular = ""
                sFoneContato = ""
                sEmail = Left(SubNull(!Email), 100)
                If Trim(SubNull(!Cnpj)) <> "" Then
                    sCPFCNPJ = RetornaNumero(!Cnpj)
                Else
                    If Trim(SubNull(!cpf)) <> "" Then
                        sCPFCNPJ = RetornaNumero(!cpf)
                    Else
                        sCPFCNPJ = ""
                    End If
                End If
                
                xImovel.RetornaEndereco nCodReduz, Imobiliario, Localizacao
                nCodLogradouroLocal = xImovel.CodLogradouro
                
                sEnderecoLocal = xImovel.Endereco
                nNumeroLocal = xImovel.Numero
                sComplementoLocal = xImovel.Complemento
                nCodBairroLocal = xImovel.CodBairro
                sBairroLocal = xImovel.Bairro
                sCidadeLocal = xImovel.Cidade
                sUFLocal = xImovel.UF
                sCEPLocal = xImovel.Cep
                If Val(nCodBairroLocal) = 999 Then
                    nCodBairroLocal = 0
                    sBairroLocal = ""
                End If
                
                sQuadra = !Quadra
                sLote = !Lote
                sAtividade = ""
                sMatricula = Val(SubNull(!NumMat))
                
                sRGIE = RetornaNumero(Trim(SubNull(!rg)))
                If !Ee_TipoEnd = 0 Then 'imovel
                    xImovel.RetornaEndereco nCodReduz, Imobiliario, Localizacao
                    nCodLogradouro = xImovel.CodLogradouro
                    sEndereco = xImovel.Endereco
                    nNumero = xImovel.Numero
                    sComplemento = xImovel.Complemento
                    nCodBairro = xImovel.CodBairro
                    sBairro = xImovel.Bairro
                    nCodCidade = xImovel.CodCidade
                    sCidade = xImovel.Cidade
                    sCep = xImovel.Cep
                    sUF = xImovel.UF
                ElseIf !Ee_TipoEnd = 1 Then 'proprietario
                    xImovel.RetornaEndereco nCodReduz, Imobiliario, cadastrocidadao
                    nCodLogradouro = Val(xImovel.CodLogradouro)
                    sEndereco = xImovel.Endereco
                    nNumero = xImovel.Numero
                    sComplemento = xImovel.Complemento
                    nCodBairro = Val(xImovel.CodBairro)
                    sBairro = xImovel.Bairro
                    nCodCidade = xImovel.CodCidade
                    sCidade = xImovel.Cidade
                    sCep = xImovel.Cep
                    sUF = xImovel.UF
                Else 'entrega
                    xImovel.RetornaEndereco nCodReduz, Imobiliario, Entrega
                    nCodLogradouro = xImovel.CodLogradouro
                    sEndereco = xImovel.Endereco
                    nNumero = xImovel.Numero
                    sComplemento = xImovel.Complemento
                    nCodBairro = Val(xImovel.CodBairro)
                    sBairro = xImovel.Bairro
                    nCodCidade = xImovel.CodCidade
                    sCidade = xImovel.Cidade
                    sCep = xImovel.Cep
                    sUF = xImovel.UF
                End If
               .Close
            End With
        ElseIf nCodReduz >= 100000 And nCodReduz < 500000 Then
            sql = "select * from vwfullempresa3 where codigomob=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If .RowCount = 0 Then GoTo Proximo
                sInscricao = ""
                sNome = !RazaoSocial
                sFoneRes = ""
                sFoneCom = ""
                sCelular = ""
                sFoneContato = Left(RetornaNumero(SubNull(!fonecontato)), 15)
                sEmail = Left(SubNull(!emailcontato), 100)
                If Trim(SubNull(!Cnpj)) <> "" Then
                    sCPFCNPJ = RetornaNumero(!Cnpj)
                Else
                    If Trim(SubNull(!cpf)) <> "" Then
                        sCPFCNPJ = RetornaNumero(!cpf)
                    Else
                        sCPFCNPJ = ""
                    End If
                End If
                sRGIE = RetornaNumero(Trim(SubNull(!inscestadual)))
                If sRGIE = "" Then
                    sRGIE = Trim(SubNull(!rg) & " " & SubNull(!Orgao))
                End If
                
                xImovel.RetornaEndereco nCodReduz, Mobiliario, Localizacao
                nCodLogradouroLocal = xImovel.CodLogradouro
                sEnderecoLocal = xImovel.Endereco
                nNumeroLocal = xImovel.Numero
                sComplementoLocal = xImovel.Complemento
                nCodBairroLocal = Val(xImovel.CodBairro)
                sBairroLocal = xImovel.Bairro
                sCidadeLocal = xImovel.Cidade
                sUFLocal = xImovel.UF
                sCEPLocal = xImovel.Cep
                If Val(nCodBairroLocal) = 999 Then
                    nCodBairroLocal = 0
                    sBairroLocal = ""
                End If
                
                sEndereco = SubNull(!eenomelogr)
                If sEndereco = "" Then
                    'local da empresa
                    xImovel.RetornaEndereco nCodReduz, Mobiliario, Localizacao
                    nCodLogradouro = xImovel.CodLogradouro
                    sEndereco = xImovel.Endereco
                    nNumero = xImovel.Numero
                    sComplemento = xImovel.Complemento
                    nCodBairro = Val(xImovel.CodBairro)
                    sBairro = xImovel.Bairro
                    nCodCidade = xImovel.CodCidade
                    sCidade = xImovel.Cidade
                    sCep = xImovel.Cep
                    sUF = xImovel.UF
                Else
'                   'endereco entrega
                    xImovel.RetornaEndereco nCodReduz, Mobiliario, Entrega
                    nCodLogradouro = xImovel.CodLogradouro
                    sEndereco = xImovel.Endereco
                    nNumero = xImovel.Numero
                    sComplemento = xImovel.Complemento
                    nCodBairro = Val(xImovel.CodBairro)
                    sBairro = xImovel.Bairro
                    nCodCidade = xImovel.CodCidade
                    sCidade = xImovel.Cidade
                    sCep = xImovel.Cep
                    sUF = xImovel.UF
                End If
               
                sQuadra = ""
                sLote = ""
                sAtividade = Left(SubNull(!ativextenso), 80)
                sMatricula = ""
               
               .Close
            End With
        Else
            sql = "select * from vwfullcidadao where codcidadao=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If .RowCount = 0 Then GoTo Proximo
                sInscricao = ""
                sNome = !nomecidadao
                sFoneRes = Left(RetornaNumero(SubNull(!telefone)), 15)
                sFoneCom = ""
                sCelular = ""
                sFoneContato = ""
                sEmail = Left(SubNull(!Email), 100)
                If Trim(SubNull(!Cnpj)) <> "" Then
                    sCPFCNPJ = RetornaNumero(!Cnpj)
                Else
                    If Trim(SubNull(!cpf)) <> "" Then
                        sCPFCNPJ = RetornaNumero(!cpf)
                    Else
                        sCPFCNPJ = ""
                    End If
                End If
                sRGIE = RetornaNumero(Trim(Left(SubNull(!rg), 14)))
                xImovel.RetornaEndereco nCodReduz, cidadao, cadastrocidadao
                nCodLogradouro = Val(xImovel.CodLogradouro)
                sEndereco = SubNull(xImovel.Endereco)
                nNumero = SubNull(xImovel.Numero)
                sComplemento = xImovel.Complemento
                nCodBairro = Val(SubNull(xImovel.CodBairro))
                sBairro = SubNull(xImovel.Bairro)
                If nCodBairro = 999 Then
                    nCodBairro = 0
                    sBairro = ""
                End If
                nCodCidade = Val(SubNull(xImovel.CodCidade))
                sCidade = SubNull(xImovel.Cidade)
                sUFLocal = xImovel.UF
                sCep = RetornaNumero(SubNull(xImovel.Cep))
                sUF = SubNull(xImovel.UF)
               
                nCodLogradouroLocal = Val(xImovel.CodLogradouro)
                sEnderecoLocal = SubNull(xImovel.Endereco)
                nNumeroLocal = SubNull(xImovel.Numero)
                sComplementoLocal = SubNull(xImovel.Complemento)
                nCodBairroLocal = Val(SubNull(xImovel.CodBairro))
                sBairroLocal = SubNull(xImovel.Bairro)
                If Val(nCodBairroLocal) = 999 Then
                    nCodBairroLocal = 0
                    sBairroLocal = ""
                End If
                sCidadeLocal = xImovel.Cidade
                sCEPLocal = RetornaCEP(CLng(nCodLogradouroLocal), nNumeroLocal)
                
                sQuadra = ""
                sLote = ""
                sAtividade = ""
                sMatricula = ""
               
               .Close
            End With
        End If
                
        For x = 1 To UBound(aCda)
            With aCda(x)
                sql = "INSERT Cadastro(IdCDA,SetorDevedor,Crc,Nome,Inscricao,CPFCnpj,RgInscrEstadual,LocalCep,LocalEndereco,LocalNumero,LocalComplemento,"
                sql = sql & "LocalBairro,LocalCidade,LocalEstado,Quadra,Lote,EntregaCep,EntregaEndereco,EntregaNumero,EntregaComplemento,EntregaBairro,"
                sql = sql & "EntregaCidade,EntregaEstado,DtGeracao) values("
                sql = sql & .nCDA & ",'" & SetorDevedor & "'," & nCodReduz & ",'" & Mask(Left(SubNull(sNome), 80)) & "','" & sInscricao & "','" & sCPFCNPJ & "','" & sRGIE & "','"
                sql = sql & sCEPLocal & "','" & sEnderecoLocal & "'," & nNumeroLocal & ",'" & Mask(Left(sComplementoLocal, 50)) & "','" & sBairroLocal & "','" & sCidadeLocal & "','" & sUFLocal & "','"
                sql = sql & sQuadra & "','" & sLote & "','" & sCep & "','" & sEndereco & "'," & nNumero & ",'" & Mask(Left(sComplemento, 50)) & "','"
                sql = sql & sBairro & "','" & Mask(sCidade) & "','" & sUF & "','" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "')"
                cnInt.Execute sql, rdExecDirect
            End With
        Next
        
        
        '***** PARTES **********************************

        If nCodReduz < 100000 Then
            sql = "SELECT proprietario.codcidadao, cidadao.nomecidadao, cidadao.cnpj,cidadao.cpf,cidadao.rg,cidadao.telefone, cidadao.email FROM proprietario INNER JOIN "
            sql = sql & "cidadao ON proprietario.codcidadao = cidadao.codcidadao Where codreduzido=" & nCodReduz & " and principal=0"
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                Do Until .EOF
                    ReDim Preserve aParte(UBound(aParte) + 1)
                    U = UBound(aParte)
                    xImovel.RetornaEndereco !CodCidadao, cidadao, cadastrocidadao
                    aParte(U).sNomeSocio = RdoAux2!nomecidadao
                    If Trim(SubNull(RdoAux2!Cnpj)) <> "" Then
                        aParte(U).sCPFCNPJSocio = RetornaNumero(RdoAux2!Cnpj)
                    Else
                        If Trim(SubNull(RdoAux2!cpf)) <> "" Then
                            aParte(U).sCPFCNPJSocio = RetornaNumero(RdoAux2!cpf)
                        Else
                            aParte(U).sCPFCNPJSocio = ""
                        End If
                    End If
                    aParte(U).sRGIE = RetornaNumero(Trim(Left(SubNull(RdoAux2!rg), 14)))
                    aParte(U).sCodLogradouro = xImovel.CodLogradouro
                    aParte(U).sEndereco = xImovel.Endereco
                    aParte(U).sNumero = xImovel.Numero
                    aParte(U).sComplemento = xImovel.Complemento
                    aParte(U).sCodBairro = xImovel.CodBairro
                    aParte(U).sBairro = xImovel.Bairro
                    aParte(U).sCodCidade = xImovel.CodCidade
                    aParte(U).sCidade = xImovel.Cidade
                    aParte(U).sCepSocio = RetornaNumero(xImovel.Cep)
                    aParte(U).sEstadoSocio = xImovel.UF
                    aParte(U).sClassificao = "Compromissário"
                    aParte(U).sFoneRes = Left(RetornaNumero(SubNull(RdoAux2!telefone)), 15)
                    aParte(U).sFoneCom = ""
                    aParte(U).sCelular = ""
                    aParte(U).sFoneContato = ""
                    aParte(U).sEmail = Left(SubNull(RdoAux2!Email), 100)
                   .MoveNext
                Loop
               .Close
            End With
        ElseIf nCodReduz >= 100000 And nCodReduz < 500000 Then
            sql = "SELECT mobiliarioproprietario.codcidadao, cidadao.nomecidadao, cidadao.cnpj,cidadao.cpf,cidadao.rg,cidadao.telefone, cidadao.email FROM mobiliarioproprietario INNER JOIN "
            sql = sql & "cidadao ON mobiliarioproprietario.codcidadao = cidadao.codcidadao Where codmobiliario=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                Do Until .EOF
                    ReDim Preserve aParte(UBound(aParte) + 1)
                    U = UBound(aParte)
                    xImovel.RetornaEndereco !CodCidadao, cidadao, cadastrocidadao
                    aParte(U).sNomeSocio = RdoAux2!nomecidadao
                    If Trim(SubNull(RdoAux2!Cnpj)) <> "" Then
                        aParte(U).sCPFCNPJSocio = RetornaNumero(RdoAux2!Cnpj)
                    Else
                        If Trim(SubNull(RdoAux2!cpf)) <> "" Then
                            aParte(U).sCPFCNPJSocio = RetornaNumero(RdoAux2!cpf)
                        Else
                            aParte(U).sCPFCNPJSocio = ""
                        End If
                    End If
                    aParte(U).sRGIE = RetornaNumero(Trim(Left(SubNull(RdoAux2!rg), 14)))
                    aParte(U).sCodLogradouro = xImovel.CodLogradouro
                    aParte(U).sEndereco = xImovel.Endereco
                    aParte(U).sNumero = xImovel.Numero
                    aParte(U).sComplemento = xImovel.Complemento
                    aParte(U).sCodBairro = xImovel.CodBairro
                    aParte(U).sBairro = xImovel.Bairro
                    aParte(U).sCodCidade = xImovel.CodCidade
                    aParte(U).sCidade = xImovel.Cidade
                    aParte(U).sCepSocio = RetornaNumero(xImovel.Cep)
                    aParte(U).sEstadoSocio = xImovel.UF
                    aParte(U).sClassificao = "Sócio"
                    aParte(U).sFoneRes = Left(RetornaNumero(SubNull(RdoAux2!telefone)), 15)
                    aParte(U).sFoneCom = ""
                    aParte(U).sCelular = ""
                    aParte(U).sFoneContato = ""
                    aParte(U).sEmail = Left(SubNull(RdoAux2!Email), 100)
                   .MoveNext
                Loop
                RdoAux2.Close
            End With
        End If
        
        For x = 1 To UBound(aCda)
            For U = 1 To UBound(aParte)
                With aParte(U)
                    sql = "INSERT Partes(idCDA,Tipo,Crc,Nome,CpfCnpj,RgInscrEstadual,Cep,Endereco,Numero,Complemento,Bairro,Cidade,Estado,DtGeracao) Values("
                    sql = sql & aCda(x).nCDA & ",'" & .sClassificao & "'," & nCodReduz & ",'" & Mask(.sNomeSocio) & "','" & .sCPFCNPJSocio & "','" & .sRGIE & "','" & .sCepSocio & "','"
                    sql = sql & Mask(.sEndereco) & "'," & Val(.sNumero) & ",'" & Mask(.sComplemento) & "','" & Mask(.sBairro) & "','"
                    sql = sql & Mask(.sCidade) & "','" & SubNull(.sEstadoSocio) & "','" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "')"
                    cnInt.Execute sql, rdExecDirect
                End With
            Next
        Next
        
Proximo:
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
Fim:
   .Close
End With
cmdExec.Enabled = True
cnInt.Close
Liberado
MsgBox "Exportação finalizada.", vbInformation

End Sub


Private Sub ExportaDebitos(bAjuizado As Boolean, nCadastro As Integer)
Dim sql As String, RdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, nCodReduz As Long, RdoAux3 As rdoResultset, ax As String
Dim t As Integer, s As Integer, U As Integer, v As Integer, aReg00() As Reg00, aReg01() As Reg01, nCodCidadao As Long, aReg10() As Reg10, aReg20() As Reg20
Dim RdoDebito As rdoResultset, qd As New rdoQuery, Achou As Boolean, x As Integer, RdoTrib As rdoResultset
Dim Rs As ADODB.Recordset, strQuery As String, sArq As String, nNumExec As Long, nAnoExec As Integer, sNumExec As String
Dim cmdComm As ADODB.Command
Dim clsImovel As New clsImovel

'MsgBox "Desativado"
'Exit Sub

Ocupado
Set qd.ActiveConnection = cn
qd.QueryTimeout = 0

'sArq = sPathBin & "\integracao.mdb"
'On Error GoTo Erro
'Set cnn = New ADODB.Connection
'cnn.CursorLocation = adUseClient
'cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sArq & ";Persist Security Info=False;Jet OLEDB:Database Password="
'cnn.Open
'If cnn.State = 0 Then
'    cnn.Close
'    GoTo Erro
'End If

'Sql = "delete from obsgeral"
'Set cmdComm = New ADODB.Command
'Set cmdComm.ActiveConnection = cnn
'cmdComm.CommandText = Sql
'cnn.Execute Sql

'Sql = "delete from obsparcela"
'Set cmdComm = New ADODB.Command
'Set cmdComm.ActiveConnection = cnn
'cmdComm.CommandText = Sql
'cnn.Execute Sql

Open txtArq.Text For Output As #1

'Carrega códigos para arquivo de débitos
sql = "SELECT  DISTINCT CODREDUZIDO FROM DEBITOPARCELA WHERE 1=1 "
If nCadastro = 1 Then
    sql = sql & " AND CODREDUZIDO BETWEEN 1 AND 39999 "
    sql = sql & " AND CODLANCAMENTO <>5 AND CODLANCAMENTO<>20 AND CODLANCAMENTO<>11 "
ElseIf nCadastro = 2 Then
    sql = sql & " AND CODREDUZIDO BETWEEN 100000 AND 300000 "
    sql = sql & " AND CODLANCAMENTO <>5 AND CODLANCAMENTO<>20 AND CODLANCAMENTO<>11 "
ElseIf nCadastro = 3 Then
    sql = sql & " AND CODREDUZIDO BETWEEN 500000 AND 699999 "
    sql = sql & " AND CODLANCAMENTO in(50,65,49,16,62,27,71,48) "
End If

If bAjuizado Then
    sql = sql & " AND DATAAJUIZA IS NOT NULL "
    sql = sql & " AND NUMPARCELA>0 AND STATUSLANC=3 "
Else
    sql = sql & " AND DATAVENCIMENTO<='12/31/2014' AND STATUSLANC =3 AND NUMPARCELA>0  "
    sql = sql & " AND DATAAJUIZA IS NULL AND DATAINSCRICAO IS NOT NULL"
End If
sql = sql & " ORDER BY CODREDUZIDO"

Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nPos = 1
    nTot = .RowCount
    Do Until .EOF
        nCodReduz = !CODREDUZIDO
        'If nCodReduz = 15 Then MsgBox "teste"
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        
        '***** OBSERVAÇÃO GERAL *****************************
    '    Sql = "select * from debitoobservacao where codreduzido=" & nCodReduz & " and usuario in('ROSE','JOSEANE','MARAB','SANTANA','ANALU','PAMELA','ADMINISTRADOR')"
    '    Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    '    If RdoAux3.RowCount > 0 Then
    '        Do Until RdoAux3.EOF
    '            Sql = "insert into obsgeral(codigo,seq,obs) values(" & nCodReduz & "," & RdoAux3!Seq & ",'" & Mask(RdoAux3!obs) & "')"
    '            Set cmdComm = New ADODB.Command
    '            Set cmdComm.ActiveConnection = cnn
    '            cmdComm.CommandText = Sql
    '            cnn.Execute Sql
    '            RdoAux3.MoveNext
    '        Loop
    '    End If
    '    RdoAux3.Close
        
        '***** REGISTRO 00 **********************************
        ReDim aReg00(0)
        t = 0
        
        aReg00(t).sTipoReg00 = "00"

        If nCodReduz < 100000 Then
            
            aReg00(t).sDescCadastro = "Imobiliário"
            sql = "select * from vwfullimovel2 where codreduzido=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If .RowCount = 0 Then GoTo Proximo
                aReg00(t).sCrc = !CodCidadao
                aReg00(t).sNome = !nomecidadao
                aReg00(t).sCadastroImob = !Inscricao
                aReg00(t).sFoneRes = Left(RetornaNumero(SubNull(!telefone)), 15)
                aReg00(t).sFoneCom = ""
                aReg00(t).sCelular = ""
                aReg00(t).sFoneContato = ""
                aReg00(t).sEmail = Left(SubNull(!Email), 100)
                If Trim(SubNull(!Cnpj)) <> "" Then
                    aReg00(t).sCPFCNPJ = RetornaNumero(!Cnpj)
                Else
                    If Trim(SubNull(!cpf)) <> "" Then
                        aReg00(t).sCPFCNPJ = RetornaNumero(!cpf)
                    Else
                        aReg00(t).sCPFCNPJ = ""
                    End If
                End If
                
                xImovel.RetornaEndereco nCodReduz, Imobiliario, Localizacao
                aReg00(t).sCodLogradouroLocal = xImovel.CodLogradouro
                aReg00(t).sEnderecoLocal = xImovel.Endereco
                aReg00(t).sNumeroLocal = xImovel.Numero
                aReg00(t).sCodBairroLocal = xImovel.CodBairro
                aReg00(t).sBairroLocal = xImovel.Bairro
                aReg00(t).sCEPLocal = xImovel.Cep
                If Val(aReg00(t).sCodBairroLocal) = 999 Then
                    aReg00(t).sCodBairroLocal = ""
                    aReg00(t).sBairroLocal = ""
                End If
                
                aReg00(t).sQuadra = !Quadra
                aReg00(t).sLote = !Lote
                aReg00(t).sAtividade = ""
                aReg00(t).sMatricula = Val(SubNull(!NumMat))
                
                aReg00(t).sRGIE = RetornaNumero(Trim(SubNull(!rg)))
                If !Ee_TipoEnd = 0 Then 'imovel
                    xImovel.RetornaEndereco nCodReduz, Imobiliario, Localizacao
                    aReg00(t).sCodLogradouro = xImovel.CodLogradouro
                    aReg00(t).sEndereco = xImovel.Endereco
                    aReg00(t).sNumero = xImovel.Numero
                    aReg00(t).sComplemento = xImovel.Complemento
                    aReg00(t).sCodBairro = xImovel.CodBairro
                    aReg00(t).sBairro = xImovel.Bairro
                    aReg00(t).sCodCidade = xImovel.CodCidade
                    aReg00(t).sCidade = xImovel.Cidade
                    aReg00(t).sCep = xImovel.Cep
                    aReg00(t).sEstado = xImovel.UF
                ElseIf !Ee_TipoEnd = 1 Then 'proprietario
                    xImovel.RetornaEndereco nCodReduz, Imobiliario, cadastrocidadao
                    aReg00(t).sCodLogradouro = xImovel.CodLogradouro
                    aReg00(t).sEndereco = xImovel.Endereco
                    aReg00(t).sNumero = xImovel.Numero
                    aReg00(t).sComplemento = xImovel.Complemento
                    aReg00(t).sCodBairro = xImovel.CodBairro
                    aReg00(t).sBairro = xImovel.Bairro
                    aReg00(t).sCodCidade = xImovel.CodCidade
                    aReg00(t).sCidade = xImovel.Cidade
                    aReg00(t).sCep = xImovel.Cep
                    aReg00(t).sEstado = xImovel.UF
                Else 'entrega
                    xImovel.RetornaEndereco nCodReduz, Imobiliario, Entrega
                    aReg00(t).sCodLogradouro = xImovel.CodLogradouro
                    aReg00(t).sEndereco = xImovel.Endereco
                    aReg00(t).sNumero = xImovel.Numero
                    aReg00(t).sComplemento = xImovel.Complemento
                    aReg00(t).sCodBairro = xImovel.CodBairro
                    aReg00(t).sBairro = xImovel.Bairro
                    aReg00(t).sCodCidade = xImovel.CodCidade
                    aReg00(t).sCidade = xImovel.Cidade
                    aReg00(t).sCep = xImovel.Cep
                    aReg00(t).sEstado = xImovel.UF
                End If
               .Close
            End With
        ElseIf nCodReduz >= 100000 And nCodReduz < 500000 Then
            aReg00(t).sCrc = nCodReduz
            aReg00(t).sDescCadastro = "Mobiliário"
            sql = "select * from vwfullempresa3 where codigomob=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If .RowCount = 0 Then GoTo Proximo
                aReg00(t).sCadastroImob = ""
                aReg00(t).sNome = !RazaoSocial
                aReg00(t).sFoneRes = ""
                aReg00(t).sFoneCom = ""
                aReg00(t).sCelular = ""
                aReg00(t).sFoneContato = Left(RetornaNumero(SubNull(!fonecontato)), 15)
                aReg00(t).sEmail = Left(SubNull(!emailcontato), 100)
                If Trim(SubNull(!Cnpj)) <> "" Then
                    aReg00(t).sCPFCNPJ = RetornaNumero(!Cnpj)
                Else
                    If Trim(SubNull(!cpf)) <> "" Then
                        aReg00(t).sCPFCNPJ = RetornaNumero(!cpf)
                    Else
                        aReg00(t).sCPFCNPJ = ""
                    End If
                End If
                aReg00(t).sRGIE = RetornaNumero(Trim(SubNull(!inscestadual)))
                
                xImovel.RetornaEndereco nCodReduz, Mobiliario, Localizacao
                aReg00(t).sCodLogradouroLocal = xImovel.CodLogradouro
                aReg00(t).sEnderecoLocal = xImovel.Endereco
                aReg00(t).sNumeroLocal = xImovel.Numero
                aReg00(t).sCodBairroLocal = xImovel.CodBairro
                aReg00(t).sBairroLocal = xImovel.Bairro
                aReg00(t).sCEPLocal = xImovel.Cep
                If Val(aReg00(t).sCodBairroLocal) = 999 Then
                    aReg00(t).sCodBairroLocal = ""
                    aReg00(t).sBairroLocal = ""
                End If
                
                aReg00(t).sEndereco = SubNull(!eenomelogr)
                If aReg00(t).sEndereco = "" Then
                    'local da empresa
                    xImovel.RetornaEndereco nCodReduz, Mobiliario, Localizacao
                    aReg00(t).sCodLogradouro = xImovel.CodLogradouro
                    aReg00(t).sEndereco = xImovel.Endereco
                    aReg00(t).sNumero = xImovel.Numero
                    aReg00(t).sComplemento = xImovel.Complemento
                    aReg00(t).sCodBairro = xImovel.CodBairro
                    aReg00(t).sBairro = xImovel.Bairro
                    aReg00(t).sCodCidade = xImovel.CodCidade
                    aReg00(t).sCidade = xImovel.Cidade
                    aReg00(t).sCep = xImovel.Cep
                    aReg00(t).sEstado = xImovel.UF
                Else
'                   'endereco entrega
                    xImovel.RetornaEndereco nCodReduz, Mobiliario, Entrega
                    aReg00(t).sCodLogradouro = xImovel.CodLogradouro
                    aReg00(t).sEndereco = xImovel.Endereco
                    aReg00(t).sNumero = xImovel.Numero
                    aReg00(t).sComplemento = xImovel.Complemento
                    aReg00(t).sCodBairro = xImovel.CodBairro
                    aReg00(t).sBairro = xImovel.Bairro
                    aReg00(t).sCodCidade = xImovel.CodCidade
                    aReg00(t).sCidade = xImovel.Cidade
                    aReg00(t).sCep = xImovel.Cep
                    aReg00(t).sEstado = xImovel.UF
                End If
               
                aReg00(t).sQuadra = ""
                aReg00(t).sLote = ""
                aReg00(t).sAtividade = Left(SubNull(!ativextenso), 80)
                aReg00(t).sMatricula = ""
               
               .Close
            End With
        Else
            aReg00(t).sCrc = nCodReduz
            aReg00(t).sDescCadastro = "Cidadão"
            sql = "select * from vwfullcidadao where codcidadao=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If .RowCount = 0 Then GoTo Proximo
                aReg00(t).sCadastroImob = ""
                aReg00(t).sNome = !nomecidadao
                aReg00(t).sFoneRes = Left(RetornaNumero(SubNull(!telefone)), 15)
                aReg00(t).sFoneCom = ""
                aReg00(t).sCelular = ""
                aReg00(t).sFoneContato = ""
                aReg00(t).sEmail = Left(SubNull(!Email), 100)
                If Trim(SubNull(!Cnpj)) <> "" Then
                    aReg00(t).sCPFCNPJ = RetornaNumero(!Cnpj)
                Else
                    If Trim(SubNull(!cpf)) <> "" Then
                        aReg00(t).sCPFCNPJ = RetornaNumero(!cpf)
                    Else
                        aReg00(t).sCPFCNPJ = ""
                    End If
                End If
                aReg00(t).sRGIE = RetornaNumero(Trim(Left(SubNull(!rg), 14)))
                xImovel.RetornaEndereco nCodReduz, cidadao, cadastrocidadao
                aReg00(t).sCodLogradouro = xImovel.CodLogradouro
                aReg00(t).sEndereco = SubNull(!Endereco)
                aReg00(t).sNumero = SubNull(!NUMIMOVEL)
                aReg00(t).sComplemento = xImovel.Complemento
                aReg00(t).sCodBairro = SubNull(!CodBairro)
                aReg00(t).sBairro = SubNull(!DescBairro)
                If Val(aReg00(t).sCodBairro) = 999 Then
                    aReg00(t).sCodBairro = ""
                    aReg00(t).sBairro = ""
                End If
                aReg00(t).sCodCidade = SubNull(!CodCidade)
                aReg00(t).sCidade = SubNull(!descCidade)
                aReg00(t).sCep = RetornaNumero(SubNull(!Cep))
                aReg00(t).sEstado = SubNull(!SiglaUF)
               
                aReg00(t).sCodLogradouroLocal = xImovel.CodLogradouro
                aReg00(t).sEnderecoLocal = SubNull(!Endereco)
                aReg00(t).sNumeroLocal = SubNull(!NUMIMOVEL)
                aReg00(t).sCodBairroLocal = SubNull(!CodBairro)
                aReg00(t).sBairroLocal = SubNull(!DescBairro)
                If Val(aReg00(t).sCodBairroLocal) = 999 Then
                    aReg00(t).sCodBairroLocal = ""
                    aReg00(t).sBairroLocal = ""
                End If
                aReg00(t).sCEPLocal = RetornaNumero(SubNull(!Cep))
                
                aReg00(t).sQuadra = ""
                aReg00(t).sLote = ""
                aReg00(t).sAtividade = ""
                aReg00(t).sMatricula = ""
               
               .Close
            End With
        End If
                
        'aReg00(T).nValorTotal = 0
        aReg00(t).sDataAtualizacao = Format(Now, "dd/mm/yyyy")
        aReg00(t).sIdAjuizamento = "0"
        aReg00(t).sCodCadastro = nCodReduz
        
        
        '***** REGISTRO 01 **********************************
        ReDim aReg01(0)

        If nCodReduz < 100000 Then
            sql = "SELECT proprietario.codcidadao, cidadao.nomecidadao, cidadao.cnpj,cidadao.cpf,cidadao.rg,cidadao.telefone, cidadao.email FROM proprietario INNER JOIN "
            sql = sql & "cidadao ON proprietario.codcidadao = cidadao.codcidadao Where codreduzido=" & nCodReduz & " and principal=0"
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount > 0 Then
                nCodCidadao = RdoAux2!CodCidadao
                
                ReDim Preserve aReg01(UBound(aReg01) + 1)
                s = UBound(aReg01)
                
                aReg01(s).sTipoReg01 = "01"
                aReg01(s).sCrcSocio = nCodCidadao
                        
                xImovel.RetornaEndereco nCodCidadao, cidadao, cadastrocidadao
                aReg01(s).sNomeSocio = RdoAux2!nomecidadao
                If Trim(SubNull(RdoAux2!Cnpj)) <> "" Then
                    aReg01(s).sCPFCNPJSocio = RetornaNumero(RdoAux2!Cnpj)
                Else
                    If Trim(SubNull(RdoAux2!cpf)) <> "" Then
                        aReg01(s).sCPFCNPJSocio = RetornaNumero(RdoAux2!cpf)
                    Else
                        aReg01(s).sCPFCNPJSocio = ""
                    End If
                End If
                aReg01(s).sRGIE = RetornaNumero(Trim(Left(SubNull(RdoAux2!rg), 14)))
                aReg01(s).sCodLogradouro = xImovel.CodLogradouro
                aReg01(s).sEndereco = xImovel.Endereco
                aReg01(s).sNumero = xImovel.Numero
                aReg01(s).sComplemento = xImovel.Complemento
                aReg01(s).sCodBairro = xImovel.CodBairro
                aReg01(s).sBairro = xImovel.Bairro
                aReg01(s).sCodCidade = xImovel.CodCidade
                aReg01(s).sCidade = xImovel.Cidade
                aReg01(s).sCepSocio = RetornaNumero(xImovel.Cep)
                aReg01(s).sEstadoSocio = xImovel.UF
                aReg01(s).sClassificao = "Compromissário"
                aReg01(s).sFoneRes = Left(RetornaNumero(SubNull(RdoAux2!telefone)), 15)
                aReg01(s).sFoneCom = ""
                aReg01(s).sCelular = ""
                aReg01(s).sFoneContato = ""
                aReg01(s).sEmail = Left(SubNull(RdoAux2!Email), 100)
                RdoAux2.Close
            End If
        ElseIf nCodReduz >= 100000 And nCodReduz < 500000 Then
            sql = "SELECT mobiliarioproprietario.codcidadao, cidadao.nomecidadao, cidadao.cnpj,cidadao.cpf,cidadao.rg,cidadao.telefone, cidadao.email FROM mobiliarioproprietario INNER JOIN "
            sql = sql & "cidadao ON mobiliarioproprietario.codcidadao = cidadao.codcidadao Where codmobiliario=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount > 0 Then
                nCodCidadao = RdoAux2!CodCidadao
                
                ReDim Preserve aReg01(UBound(aReg01) + 1)
                s = UBound(aReg01)
                
                aReg01(s).sTipoReg01 = "01"
                aReg01(s).sCrcSocio = nCodCidadao
                        
                xImovel.RetornaEndereco nCodCidadao, cidadao, cadastrocidadao
                aReg01(s).sNomeSocio = RdoAux2!nomecidadao
                If Trim(SubNull(RdoAux2!Cnpj)) <> "" Then
                    aReg01(s).sCPFCNPJSocio = RetornaNumero(RdoAux2!Cnpj)
                Else
                    If Trim(SubNull(RdoAux2!cpf)) <> "" Then
                        aReg01(s).sCPFCNPJSocio = RetornaNumero(RdoAux2!cpf)
                    Else
                        aReg01(s).sCPFCNPJSocio = ""
                    End If
                End If
                aReg01(s).sRGIE = RetornaNumero(Trim(Left(SubNull(RdoAux2!rg), 14)))
                aReg01(s).sCodLogradouro = xImovel.CodLogradouro
                aReg01(s).sEndereco = xImovel.Endereco
                aReg01(s).sNumero = xImovel.Numero
                aReg01(s).sComplemento = xImovel.Complemento
                aReg01(s).sCodBairro = xImovel.CodBairro
                aReg01(s).sBairro = xImovel.Bairro
                aReg01(s).sCodCidade = xImovel.CodCidade
                aReg01(s).sCidade = xImovel.Cidade
                aReg01(s).sCepSocio = RetornaNumero(xImovel.Cep)
                aReg01(s).sEstadoSocio = xImovel.UF
                aReg01(s).sClassificao = "Sócio"
                aReg01(s).sFoneRes = Left(RetornaNumero(SubNull(RdoAux2!telefone)), 15)
                aReg01(s).sFoneCom = ""
                aReg01(s).sCelular = ""
                aReg01(s).sFoneContato = ""
                aReg01(s).sEmail = Left(SubNull(RdoAux2!Email), 100)
                RdoAux2.Close
            End If
        End If
        
        '***** REGISTRO 10 **********************************
        ReDim aReg10(0)
        ReDim aReg20(0)
        
        On Error Resume Next
        RdoDebito.Close
        On Error GoTo 0
        If Not bAjuizado Then
            If nCadastro = 3 Then
                qd.sql = "{ Call spEXTRATOAJUIZARTAXA(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
            Else
                qd.sql = "{ Call spEXTRATOAJUIZAR(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
            End If
        Else
            qd.sql = "{ Call spEXTRATOAJUIZADO(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
        End If
        qd(0) = nCodReduz
        qd(1) = nCodReduz
        qd(2) = 1990
        qd(3) = Year(Now)
        qd(4) = 1
        qd(5) = 999
        qd(6) = 0
        qd(7) = 999
        qd(8) = 1
        qd(9) = 999
        qd(10) = 0
        qd(11) = 99
        qd(12) = 3
        qd(13) = 3
        qd(14) = Format(Now, "mm/dd/yyyy")
        qd(15) = "Integrativa"
        If bAjuizado Then
            qd(16) = 1
        End If
        Set RdoDebito = qd.OpenResultset(rdOpenKeyset)
        With RdoDebito
            Do Until .EOF
            
                U = UBound(aReg10)
                Achou = False
                For x = 1 To U
                    If Val(aReg10(x).sExercicio) = !AnoExercicio And Val(aReg10(x).sCodigoDivida) = !CodLancamento And Val(aReg10(x).sSubCodDivida) = !SeqLancamento And _
                       Val(aReg10(x).sNumParcela) = !NumParcela And Val(aReg10(x).sSubParcela) = !CODCOMPLEMENTO Then
                        nNumExec = Val(SubNull(!numexecfiscal))
                        nAnoExec = Val(SubNull(!anoexecfiscal))
                        If nAnoExec > 0 Then
                            aReg10(x).sNumExecFiscal = Format(nNumExec, "00000")
                            aReg10(x).sAnoExecFiscal = Format(nAnoExec, "0000")
                        End If
                        Achou = True
                        Exit For
                    End If
                Next
                
  '              nNumExec = Val(SubNull(!numexecfiscal))
  '              nAnoExec = Val(SubNull(!anoexecfiscal))
  '              If nAnoExec > 0 Then
   '                 aReg00(t).sNumExecFiscal = Format(nNumExec, "00000")
   '                 aReg00(t).sAnoExecFiscal = Format(nAnoExec, "0000")
   '             End If
                
                'SE NÃO ENCONTRAR O LANCAMENTO NA MATRIZ, ADICIONAR ELE
                If Not Achou Then
                                                        
                    '***** OBSERVAÇÃO DA PARCELA ************************
'                    On Error Resume Next
'                    Sql = "select * from obsparcela where codreduzido=" & nCodReduz & " and anoexercicio=" & !AnoExercicio & " and seqlancamento=" & !SeqLancamento & " and "
'                    Sql = Sql & "numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO & " and usuario in('ROSE','JOSEANE','MARAB','SANTANA','ANALU','PAMELA','ADMINISTRADOR')"
'                    Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'                    If RdoAux3.RowCount > 0 Then
'                        Do Until RdoAux3.EOF
'                            Sql = "insert into obsparcela(codigo,exercicio,coddivida,subcoddivida,numparcela,subnumparcela,seqobs,obs) values(" & nCodReduz & "," & RdoAux3!AnoExercicio & "," & RdoAux3!CodLancamento & ","
'                            Sql = Sql & RdoAux3!SeqLancamento & "," & RdoAux3!NumParcela & "," & RdoAux3!CODCOMPLEMENTO & "," & RdoAux3!Seq & ",'" & Mask(RdoAux3!obs) & "')"
'                            Set cmdComm = New ADODB.Command
'                            Set cmdComm.ActiveConnection = cnn
'                            cmdComm.CommandText = Sql
'                            cnn.Execute Sql
'                            RdoAux3.MoveNext
 '                       Loop
 '                   End If
 '                   RdoAux3.Close
 '                   On Error GoTo Erro
                    '****************************************************
                                                                            
                    
                                                        
                    ReDim Preserve aReg10(UBound(aReg10) + 1)
                    U = UBound(aReg10)
                                        
                    aReg10(U).sTipoReg = "10"
                    aReg10(U).sTributo = !DESCLANCAMENTO
                    aReg10(U).sExercicio = !AnoExercicio
                    aReg10(U).sCodigoDivida = !CodLancamento
                    aReg10(U).sSubCodDivida = !SeqLancamento
                    aReg10(U).sNumParcela = !NumParcela
                    aReg10(U).sSubParcela = !CODCOMPLEMENTO
                    aReg10(U).sSeqInscricaoDA = 0
                    aReg10(U).sNroCDA = Val(SubNull(!certidao))
                    aReg10(U).sFolha = Val(SubNull(!PAGINA))
                    aReg10(U).sLivro = Val(SubNull(!NUMLIVRO))
                    aReg10(U).sDtVencimento = Format(!DataVencimento, "dd/mm/yyyy")
                    If IsDate(!datainscricao) Then
                        aReg10(U).sDtInscricao = Format(!datainscricao, "dd/mm/yyyy")
                    Else
                        aReg10(U).sDtInscricao = ""
                    End If
                    aReg10(U).nPrincipal = RetornaNumero(FormatNumber(!VALORTRIBUTO, 2))
                    aReg10(U).nJuros = RetornaNumero(FormatNumber(!ValorJuros, 2))
                    aReg10(U).nMulta = RetornaNumero(FormatNumber(!ValorMulta, 2))
                    aReg10(U).nCorrecao = RetornaNumero(FormatNumber(!valorcorrecao, 2))
                    aReg10(U).sHonorarios = 0
                    aReg10(U).nTotal = !ValorTotal
                    aReg10(U).nTotalAcumulado = !ValorTotal
'                    aReg10(U).sExecFiscal = sNumExec
                    nNumExec = Val(SubNull(!numexecfiscal))
                    nAnoExec = Val(SubNull(!anoexecfiscal))
                    If nAnoExec > 0 Then
                        aReg10(U).sNumExecFiscal = Format(nNumExec, "00000")
                        aReg10(U).sAnoExecFiscal = Format(nAnoExec, "0000")
                    End If
                Else
                    aReg10(x).nTotal = aReg10(U).nTotal + !ValorTotal
                    aReg10(x).nPrincipal = aReg10(x).nPrincipal + RetornaNumero(FormatNumber(!VALORTRIBUTO, 2))
                    aReg10(x).nJuros = aReg10(x).nJuros + RetornaNumero(FormatNumber(!ValorJuros, 2))
                    aReg10(x).nMulta = aReg10(x).nMulta + RetornaNumero(FormatNumber(!ValorMulta, 2))
                    aReg10(x).nCorrecao = aReg10(x).nCorrecao + RetornaNumero(FormatNumber(!valorcorrecao, 2))
                    aReg10(x).nTotalAcumulado = aReg10(x).nTotalAcumulado + !ValorTotal
                    
                End If
                
                '***** FUNDAMENTO/ARTIGO ****************************
                sql = "SELECT tributo.codtributo, tributo.desctributo, tributo.livro, tributoartigo.artigo FROM tributo LEFT OUTER JOIN "
                sql = sql & "tributoartigo ON tributo.codtributo = tributoartigo.codtributo where tributo.codtributo=" & !CodTributo
                Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                If Not IsNull(RdoAux3!ARTIGO) Then
                    aReg10(U).sFundamento = aReg10(U).sFundamento & Replace(Replace(RdoAux3!ARTIGO, vbLf, ""), vbCr, "") & "|"
                End If
               ' RdoAux3.Close
'                Sql = "select livro from lancamento where codlancamento=" & !CodLancamento
'                Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                
                If !CodLancamento = 6 Then
                    sql = "select codtributo from debitotributo where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and codlancamento=" & !CodLancamento & " and "
                    sql = sql & "seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO & " and codtributo=14"
                    Set RdoTrib = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                    If RdoTrib.RowCount > 0 Then
                        aReg10(U).sNatureza = "Tributária"
                    Else
                        aReg10(U).sNatureza = "Não Tributária"
                    End If
                    RdoTrib.Close
                ElseIf !CodLancamento = 13 Then
                    aReg10(U).sNatureza = "Tributária"
                Else
                    If RdoAux3!livro = 3 Or RdoAux3!livro = 4 Or RdoAux3!livro = 5 Then
                        aReg10(U).sNatureza = "Não Tributária"
                    Else
                        aReg10(U).sNatureza = "Tributária"
                    End If
                End If
                RdoAux3.Close
                If aReg10(U).sFundamento <> "" Then
                    aReg10(U).sFundamento = Left(aReg10(U).sFundamento, Len(aReg10(U).sFundamento) - 1)
                End If
                
                
                aReg10(U).sCartorio = ""
                aReg10(U).sProcessoAdm = ""
                aReg10(U).sAutoInfracao = ""
                aReg10(U).sNroParcelamento = ""
                
                aReg10(U).sCodLogradouro = aReg00(t).sCodLogradouroLocal
                aReg10(U).sEndereco = aReg00(t).sEnderecoLocal
                aReg10(U).sNumero = aReg00(t).sNumeroLocal
                aReg10(U).sQuadra = aReg00(t).sQuadra
                aReg10(U).sLote = aReg00(t).sLote
                aReg10(U).sCep = aReg00(t).sCEPLocal
                aReg10(U).sCodBairro = aReg00(t).sCodBairroLocal
                aReg10(U).sBairro = aReg00(t).sBairroLocal
                aReg10(U).sAtividade = aReg00(t).sAtividade
                aReg10(U).sMatricula = aReg00(t).sMatricula
                
                
                
                ReDim Preserve aReg20(UBound(aReg20) + 1)
                v = UBound(aReg20)
                                    
                aReg20(v).sTipoReg = "20"
                aReg20(v).sExercicio = !AnoExercicio
                aReg20(v).sCodigoDivida = !CodLancamento
                aReg20(v).sSubCodDivida = !SeqLancamento
                aReg20(v).sNumParcela = !NumParcela
                aReg20(v).sSubParcela = !CODCOMPLEMENTO
                aReg20(v).sCodReceita = !CodTributo
                aReg20(v).sDescricao = !abrevTributo
                aReg20(v).sValorOriginal = RetornaNumero(FormatNumber(!VALORTRIBUTO, 2))
PROXIMODEBITO:
                DoEvents
               .MoveNext
            Loop
           .Close
        End With

        
        '****************************************************
        If UBound(aReg10) = 0 Then GoTo Proximo
        With aReg00(t)
            ax = cf(.sTipoReg00, Character, 2) & cf(.sCrc, Numeric, 20) & cf(.sNome, Character, 50) & cf(.sCPFCNPJ, Numeric, 14) & cf(Left(.sRGIE, 14), Numeric, 14) & cf(.sCodLogradouro, Numeric, 4) & cf(.sEndereco, Character, 60) & cf(.sNumero, Numeric, 5) & cf(.sComplemento, Character, 25)
            ax = ax & cf(.sCodBairro, Numeric, 4) & cf(.sBairro, Character, 30) & cf(.sCodCidade, Numeric, 4) & cf(.sCidade, Character, 30) & cf(.sCep, Numeric, 8) & cf(.sEstado, Character, 2) & cf(RetornaNumero(FormatNumber(.nValorTotal, 2)), Numeric, 14) & cf(.sDataAtualizacao, Character, 10)
            ax = ax & cf(.sIdAjuizamento, Numeric, 8) & cf(.sDescCadastro, Character, 20) & cf(.sCodCadastro, Numeric, 20) & cf(.sCadastroImob, Character, 25) & cf(.sFoneRes, Numeric, 15) & cf(.sFoneCom, Numeric, 15) & cf(.sCelular, Numeric, 15) & cf(.sFoneContato, Numeric, 15) & cf(.sEmail, Character, 100)
            Print #1, ax
        End With
        

        For s = 1 To UBound(aReg01)
            With aReg01(s)
                ax = cf(.sTipoReg01, Character, 2) & cf(.sCrcSocio, Numeric, 20) & cf(.sNomeSocio, Character, 50) & cf(.sCPFCNPJSocio, Numeric, 14) & cf(.sRGIE, Numeric, 14) & cf(.sCodLogradouro, Numeric, 4) & cf(.sEndereco, Character, 60) & cf(.sNumero, Numeric, 5) & cf(.sComplemento, Character, 25)
                ax = ax & cf(.sCodBairro, Numeric, 4) & cf(.sBairro, Character, 30) & cf(.sCodCidade, Numeric, 4) & cf(.sCidade, Character, 30) & cf(.sCepSocio, Numeric, 8) & cf(.sEstadoSocio, Character, 2) & cf(.sClassificao, Character, 15) & cf(.sFoneRes, Numeric, 15) & cf(.sFoneCom, Numeric, 15)
                ax = ax & cf(.sCelular, Numeric, 15) & cf(.sFoneContato, Numeric, 15) & cf(.sEmail, Character, 100)
                Print #1, ax
            End With
        Next

        For U = 1 To UBound(aReg10)
            With aReg10(U)
                ax = cf(.sTipoReg, Character, 2) & cf(.sTributo, Character, 30) & cf(.sExercicio, Numeric, 4) & cf(.sCodigoDivida, Numeric, 4) & cf(.sSubCodDivida, Numeric, 4) & cf(.sNumParcela, Numeric, 2) & cf(.sSubParcela, Numeric, 2) & cf(.sSeqInscricaoDA, Numeric, 8) & cf(.sNroCDA, Character, 10)
                ax = ax & cf(.sFolha, Numeric, 5) & cf(.sLivro, Numeric, 4) & cf(.sDtVencimento, Character, 10) & cf(.sDtInscricao, Character, 10) & cf(.nPrincipal, Numeric, 14) & cf(.nJuros, Numeric, 14) & cf(.nMulta, Numeric, 14) & cf(.nCorrecao, Numeric, 14) & cf(.sHonorarios, Numeric, 14)
                ax = ax & cf(RetornaNumero(FormatNumber(.nTotalAcumulado, 2)), Numeric, 14) & cf(.sCodLogradouro, Numeric, 4) & cf(.sEndereco, Character, 60) & cf(.sNumero, Numeric, 5) & cf(.sQuadra, Character, 10) & cf(.sLote, Character, 20) & cf(.sCep, Numeric, 8) & cf(.sCodBairro, Numeric, 4) & cf(.sBairro, Character, 30) & cf(.sAtividade, Character, 80)
                ax = ax & cf(.sMatricula, Character, 20) & cf(.sCartorio, Character, 50) & cf(.sProcessoAdm, Character, 20) & cf(.sAutoInfracao, Character, 20) & cf(.sNatureza, Character, 20) & cf(.sNroParcelamento, Character, 20) & cf(Left(.sFundamento, 1000), Character, 1000) & cf(.sNumExecFiscal, Character, 5) & cf(.sAnoExecFiscal, Character, 4)
                Print #1, ax
                For v = 1 To UBound(aReg20)
                    If .sExercicio = aReg20(v).sExercicio And .sCodigoDivida = aReg20(v).sCodigoDivida And .sSubCodDivida = aReg20(v).sSubCodDivida And .sNumParcela = aReg20(v).sNumParcela And .sSubParcela = aReg20(v).sSubParcela Then
                        With aReg20(v)
                            ax = cf(.sTipoReg, Character, 2) & cf(.sExercicio, Numeric, 4) & cf(.sCodReceita, Numeric, 4) & cf(.sDescricao, Character, 20) & cf(.sValorOriginal, Numeric, 14)
                            Print #1, ax
                        End With
                    End If
                Next
            End With
        Next



Proximo:
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

Close #1
Liberado

'cnn.Close
PBar.value = 0
PBar.Color = vbWhite
MsgBox "Arquivo finalizado.", vbInformation, "Informação"

Exit Sub
Erro:

'Close #1
Liberado

'cnn.Close
MsgBox Err.Description
Resume Next
End Sub

Private Function cf(vParametro As Variant, sTipo As FieldType, nTamanho As Integer) As String
Dim sRet As String

If sTipo = Character Then
    sRet = FillSpace(CStr(vParametro), nTamanho)
Else
    sRet = FillLeft(CStr(vParametro), nTamanho)
End If

sRet = "[" & sRet & "]"
cf = sRet

End Function

Private Function FillSpace(sPalavra As String, nTamanho As Integer) As String
Dim sTmp As String

If Len(sPalavra) > nTamanho Then sPalavra = Left(sPalavra, nTamanho)
sTmp = sPalavra & Space(nTamanho - Len(sPalavra))
FillSpace = sTmp

End Function

Private Function FillLeft(sTexto As String, nTamanho As Integer) As String

FillLeft = String$(nTamanho - Len(sTexto), "0") & sTexto

End Function

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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set xImovel = Nothing
End Sub

Private Sub ImportaDebitosAjuizados()
Dim sLinha As String, cc As cCommonDlg, fName As String, nPos As Long, nTot As Long
Dim nCDA As Long, sDataAjuiza As String, nCodReduz As Long, sNumProc As String, sAnoProc As String


If MsgBox("Deseja importar o arquivo de débitos ajuizados ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub

Set cc = New cCommonDlg
cc.VBGetOpenFileName fName, , , False, , , "Documento de Texto|*.txt", , App.Path & "\Bin", "Selecione o arquivo do simples nacional", , Me.HWND, OFN_HIDEREADONLY, False


If fName = "" Then Exit Sub
Ocupado
nTot = 0
Open fName For Input As #1
    Do While Not EOF(1)
        Line Input #1, sLinha
        nTot = nTot + 1
    Loop
Close #1

nPos = 1
Open fName For Input As #1
   Do While Not EOF(1)
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
            DoEvents
        End If
        Line Input #1, sLinha
        
        nCodReduz = Val(Mid(sLinha, 2, 25))
        nCDA = Val(Mid(sLinha, 63, 6))
        sDataAjuiza = Mid(sLinha, 112, 10)
        sNumProc = Mid(sLinha, 94, 6)
        sAnoProc = Mid(sLinha, 102, 4)
        
        sql = "update debitoparcela set dataajuiza='" & Format(sDataAjuiza, "mm/dd/yyyy") & "',numexecfiscal=" & Val(sNumProc) & ",anoexecfiscal=" & Val(sAnoProc)
        sql = sql & " where codreduzido=" & nCodReduz & " and numcertidao=" & nCDA & " and statuslanc in (2,3,4,19)"
        cn.Execute sql, rdExecDirect
        nPos = nPos + 1
   Loop
Close #1
Liberado
MsgBox "Importação concluída", vbInformation, "Informação"

End Sub

Private Sub ExportaSerasa(nCadastro As Integer)
Dim sql As String, RdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, nCodReduz As Long, RdoAux3 As rdoResultset, ax As String
Dim t As Integer, s As Integer, U As Integer, v As Integer, aReg00() As Reg00, aReg01() As Reg01, nCodCidadao As Long, aReg10() As Reg10, aReg20() As Reg20
Dim RdoDebito As rdoResultset, qd As New rdoQuery, Achou As Boolean, x As Integer
Dim Rs As ADODB.Recordset, strQuery As String, sArq As String, nNumExec As Long, nAnoExec As Integer, sNumExec As String
Dim cmdComm As ADODB.Command
Dim clsImovel As New clsImovel

Ocupado
Set qd.ActiveConnection = cn
qd.QueryTimeout = 0

Open txtArq.Text For Output As #1

'Carrega códigos para arquivo de débitos
sql = "SELECT  DISTINCT CODREDUZIDO FROM DEBITOPARCELA WHERE 1=1 "
If nCadastro = 1 Then
    sql = sql & " AND CODREDUZIDO BETWEEN 1 AND 39999 "
    'Sql = Sql & " AND CODREDUZIDO BETWEEN 10001 AND 20000 "
    'Sql = Sql & " AND CODREDUZIDO =18760 "
    sql = sql & " AND CODLANCAMENTO <>5 AND CODLANCAMENTO<>20 AND CODLANCAMENTO<>11 "
ElseIf nCadastro = 2 Then
    sql = sql & " AND CODREDUZIDO BETWEEN 100000 AND 300000 "
    sql = sql & " AND CODLANCAMENTO <>5 AND CODLANCAMENTO<>20 AND CODLANCAMENTO<>11 "
ElseIf nCadastro = 3 Then
    sql = sql & " AND CODREDUZIDO BETWEEN 500000 AND 699999 "
    sql = sql & " AND CODLANCAMENTO in(50,65,49,16,62,27,71,48) "
End If

sql = sql & " AND DATAVENCIMENTO<'01/01/2015' AND STATUSLANC in (3) AND NUMPARCELA>0  "
'Sql = Sql & " AND DATAINSCRICAO IS NOT NULL AND DATAAJUIZA IS NULL"
'Sql = Sql & " and DATAAJUIZA IS NULL "
sql = sql & " ORDER BY CODREDUZIDO"

Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nPos = 1
    nTot = .RowCount
    Do Until .EOF
        nCodReduz = !CODREDUZIDO
        'If nCodReduz = 15 Then MsgBox "teste"
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        
        '***** REGISTRO 00 **********************************
        ReDim aReg00(0)
        t = 0
        
        aReg00(t).sTipoReg00 = "00"

        If nCodReduz < 100000 Then
            
            aReg00(t).sDescCadastro = "Imobiliário"
            sql = "select * from vwfullimovel2 where codreduzido=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If .RowCount = 0 Then GoTo Proximo
                aReg00(t).sCrc = !CodCidadao
                aReg00(t).sNome = !nomecidadao
                aReg00(t).sCadastroImob = !Inscricao
                aReg00(t).sFoneRes = Left(RetornaNumero(SubNull(!telefone)), 15)
                aReg00(t).sFoneCom = ""
                aReg00(t).sCelular = ""
                aReg00(t).sFoneContato = ""
                aReg00(t).sEmail = Left(SubNull(!Email), 100)
                If Trim(SubNull(!Cnpj)) <> "" Then
                    aReg00(t).sCPFCNPJ = RetornaNumero(!Cnpj)
                Else
                    If Trim(SubNull(!cpf)) <> "" Then
                        aReg00(t).sCPFCNPJ = RetornaNumero(!cpf)
                    Else
                        aReg00(t).sCPFCNPJ = ""
                    End If
                End If
                
                xImovel.RetornaEndereco nCodReduz, Imobiliario, Localizacao
                aReg00(t).sCodLogradouroLocal = xImovel.CodLogradouro
                aReg00(t).sEnderecoLocal = xImovel.Endereco
                aReg00(t).sNumeroLocal = xImovel.Numero
                aReg00(t).sCodBairroLocal = xImovel.CodBairro
                aReg00(t).sBairroLocal = xImovel.Bairro
                aReg00(t).sCEPLocal = xImovel.Cep
                If Val(aReg00(t).sCodBairroLocal) = 999 Then
                    aReg00(t).sCodBairroLocal = ""
                    aReg00(t).sBairroLocal = ""
                End If
                
                aReg00(t).sQuadra = !Quadra
                aReg00(t).sLote = !Lote
                aReg00(t).sAtividade = ""
                aReg00(t).sMatricula = Val(SubNull(!NumMat))
                
                aReg00(t).sRGIE = RetornaNumero(Trim(SubNull(!rg)))
                If !Ee_TipoEnd = 0 Then 'imovel
                    xImovel.RetornaEndereco nCodReduz, Imobiliario, Localizacao
                    aReg00(t).sCodLogradouro = xImovel.CodLogradouro
                    aReg00(t).sEndereco = xImovel.Endereco
                    aReg00(t).sNumero = xImovel.Numero
                    aReg00(t).sComplemento = xImovel.Complemento
                    aReg00(t).sCodBairro = xImovel.CodBairro
                    aReg00(t).sBairro = xImovel.Bairro
                    aReg00(t).sCodCidade = xImovel.CodCidade
                    aReg00(t).sCidade = xImovel.Cidade
                    aReg00(t).sCep = xImovel.Cep
                    aReg00(t).sEstado = xImovel.UF
                ElseIf !Ee_TipoEnd = 1 Then 'proprietario
                    xImovel.RetornaEndereco nCodReduz, Imobiliario, cadastrocidadao
                    aReg00(t).sCodLogradouro = xImovel.CodLogradouro
                    aReg00(t).sEndereco = xImovel.Endereco
                    aReg00(t).sNumero = xImovel.Numero
                    aReg00(t).sComplemento = xImovel.Complemento
                    aReg00(t).sCodBairro = xImovel.CodBairro
                    aReg00(t).sBairro = xImovel.Bairro
                    aReg00(t).sCodCidade = xImovel.CodCidade
                    aReg00(t).sCidade = xImovel.Cidade
                    aReg00(t).sCep = xImovel.Cep
                    aReg00(t).sEstado = xImovel.UF
                Else 'entrega
                    xImovel.RetornaEndereco nCodReduz, Imobiliario, Entrega
                    aReg00(t).sCodLogradouro = xImovel.CodLogradouro
                    aReg00(t).sEndereco = xImovel.Endereco
                    aReg00(t).sNumero = xImovel.Numero
                    aReg00(t).sComplemento = xImovel.Complemento
                    aReg00(t).sCodBairro = xImovel.CodBairro
                    aReg00(t).sBairro = xImovel.Bairro
                    aReg00(t).sCodCidade = xImovel.CodCidade
                    aReg00(t).sCidade = xImovel.Cidade
                    aReg00(t).sCep = xImovel.Cep
                    aReg00(t).sEstado = xImovel.UF
                End If
               .Close
            End With
        ElseIf nCodReduz >= 100000 And nCodReduz < 500000 Then
            aReg00(t).sCrc = nCodReduz
            aReg00(t).sDescCadastro = "Mobiliário"
            sql = "select * from vwfullempresa3 where codigomob=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If .RowCount = 0 Then GoTo Proximo
                aReg00(t).sCadastroImob = ""
                aReg00(t).sNome = !RazaoSocial
                aReg00(t).sFoneRes = ""
                aReg00(t).sFoneCom = ""
                aReg00(t).sCelular = ""
                aReg00(t).sFoneContato = Left(RetornaNumero(SubNull(!fonecontato)), 15)
                aReg00(t).sEmail = Left(SubNull(!emailcontato), 100)
                If Trim(SubNull(!Cnpj)) <> "" Then
                    aReg00(t).sCPFCNPJ = RetornaNumero(!Cnpj)
                Else
                    If Trim(SubNull(!cpf)) <> "" Then
                        aReg00(t).sCPFCNPJ = RetornaNumero(!cpf)
                    Else
                        aReg00(t).sCPFCNPJ = ""
                    End If
                End If
                aReg00(t).sRGIE = RetornaNumero(Trim(SubNull(!inscestadual)))
                
                xImovel.RetornaEndereco nCodReduz, Mobiliario, Localizacao
                aReg00(t).sCodLogradouroLocal = xImovel.CodLogradouro
                aReg00(t).sEnderecoLocal = xImovel.Endereco
                aReg00(t).sNumeroLocal = xImovel.Numero
                aReg00(t).sCodBairroLocal = xImovel.CodBairro
                aReg00(t).sBairroLocal = xImovel.Bairro
                aReg00(t).sCEPLocal = xImovel.Cep
                If Val(aReg00(t).sCodBairroLocal) = 999 Then
                    aReg00(t).sCodBairroLocal = ""
                    aReg00(t).sBairroLocal = ""
                End If
                
                aReg00(t).sEndereco = SubNull(!eenomelogr)
                If aReg00(t).sEndereco = "" Then
                    'local da empresa
                    xImovel.RetornaEndereco nCodReduz, Mobiliario, Localizacao
                    aReg00(t).sCodLogradouro = xImovel.CodLogradouro
                    aReg00(t).sEndereco = xImovel.Endereco
                    aReg00(t).sNumero = xImovel.Numero
                    aReg00(t).sComplemento = xImovel.Complemento
                    aReg00(t).sCodBairro = xImovel.CodBairro
                    aReg00(t).sBairro = xImovel.Bairro
                    aReg00(t).sCodCidade = xImovel.CodCidade
                    aReg00(t).sCidade = xImovel.Cidade
                    aReg00(t).sCep = xImovel.Cep
                    aReg00(t).sEstado = xImovel.UF
                Else
'                   'endereco entrega
                    xImovel.RetornaEndereco nCodReduz, Mobiliario, Entrega
                    aReg00(t).sCodLogradouro = xImovel.CodLogradouro
                    aReg00(t).sEndereco = xImovel.Endereco
                    aReg00(t).sNumero = xImovel.Numero
                    aReg00(t).sComplemento = xImovel.Complemento
                    aReg00(t).sCodBairro = xImovel.CodBairro
                    aReg00(t).sBairro = xImovel.Bairro
                    aReg00(t).sCodCidade = xImovel.CodCidade
                    aReg00(t).sCidade = xImovel.Cidade
                    aReg00(t).sCep = xImovel.Cep
                    aReg00(t).sEstado = xImovel.UF
                End If
               
                aReg00(t).sQuadra = ""
                aReg00(t).sLote = ""
                aReg00(t).sAtividade = Left(SubNull(!ativextenso), 80)
                aReg00(t).sMatricula = ""
               
               .Close
            End With
        Else
            aReg00(t).sCrc = nCodReduz
            aReg00(t).sDescCadastro = "Cidadão"
            sql = "select * from vwfullcidadao where codcidadao=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If .RowCount = 0 Then GoTo Proximo
                aReg00(t).sCadastroImob = ""
                aReg00(t).sNome = !nomecidadao
                aReg00(t).sFoneRes = Left(RetornaNumero(SubNull(!telefone)), 15)
                aReg00(t).sFoneCom = ""
                aReg00(t).sCelular = ""
                aReg00(t).sFoneContato = ""
                aReg00(t).sEmail = Left(SubNull(!Email), 100)
                If Trim(SubNull(!Cnpj)) <> "" Then
                    aReg00(t).sCPFCNPJ = RetornaNumero(!Cnpj)
                Else
                    If Trim(SubNull(!cpf)) <> "" Then
                        aReg00(t).sCPFCNPJ = RetornaNumero(!cpf)
                    Else
                        aReg00(t).sCPFCNPJ = ""
                    End If
                End If
                aReg00(t).sRGIE = RetornaNumero(Trim(Left(SubNull(!rg), 14)))
                xImovel.RetornaEndereco nCodReduz, cidadao, cadastrocidadao
                aReg00(t).sCodLogradouro = xImovel.CodLogradouro
                aReg00(t).sEndereco = SubNull(!Endereco)
                aReg00(t).sNumero = SubNull(!NUMIMOVEL)
                aReg00(t).sComplemento = xImovel.Complemento
                aReg00(t).sCodBairro = SubNull(!CodBairro)
                aReg00(t).sBairro = SubNull(!DescBairro)
                If Val(aReg00(t).sCodBairro) = 999 Then
                    aReg00(t).sCodBairro = ""
                    aReg00(t).sBairro = ""
                End If
                aReg00(t).sCodCidade = SubNull(!CodCidade)
                aReg00(t).sCidade = SubNull(!descCidade)
                aReg00(t).sCep = RetornaNumero(SubNull(!Cep))
                aReg00(t).sEstado = SubNull(!SiglaUF)
               
                aReg00(t).sCodLogradouroLocal = xImovel.CodLogradouro
                aReg00(t).sEnderecoLocal = SubNull(!Endereco)
                aReg00(t).sNumeroLocal = SubNull(!NUMIMOVEL)
                aReg00(t).sCodBairroLocal = SubNull(!CodBairro)
                aReg00(t).sBairroLocal = SubNull(!DescBairro)
                If Val(aReg00(t).sCodBairroLocal) = 999 Then
                    aReg00(t).sCodBairroLocal = ""
                    aReg00(t).sBairroLocal = ""
                End If
                aReg00(t).sCEPLocal = RetornaNumero(SubNull(!Cep))
                
                aReg00(t).sQuadra = ""
                aReg00(t).sLote = ""
                aReg00(t).sAtividade = ""
                aReg00(t).sMatricula = ""
               
               .Close
            End With
        End If
                
        'aReg00(T).nValorTotal = 0
        aReg00(t).sDataAtualizacao = Format(Now, "dd/mm/yyyy")
        aReg00(t).sIdAjuizamento = "0"
        aReg00(t).sCodCadastro = nCodReduz
        
        
        '***** REGISTRO 01 **********************************
        ReDim aReg01(0)
'
'        If nCodReduz < 100000 Then
'            Sql = "SELECT proprietario.codcidadao, cidadao.nomecidadao, cidadao.cnpj,cidadao.cpf,cidadao.rg,cidadao.telefone, cidadao.email FROM proprietario INNER JOIN "
'            Sql = Sql & "cidadao ON proprietario.codcidadao = cidadao.codcidadao Where codreduzido=" & nCodReduz & " and principal=0"
'            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'            If RdoAux2.RowCount > 0 Then
'                nCodCidadao = RdoAux2!CodCidadao
'
'                ReDim Preserve aReg01(UBound(aReg01) + 1)
'                s = UBound(aReg01)
'
'                aReg01(s).sTipoReg01 = "01"
'                aReg01(s).sCrcSocio = nCodCidadao
'
'                xImovel.RetornaEndereco nCodCidadao, cidadao, cadastrocidadao
'                aReg01(s).sNomeSocio = RdoAux2!nomecidadao
'                If Trim(SubNull(RdoAux2!Cnpj)) <> "" Then
'                    aReg01(s).sCPFCNPJSocio = RetornaNumero(RdoAux2!Cnpj)
'                Else
'                    If Trim(SubNull(RdoAux2!CPF)) <> "" Then
'                        aReg01(s).sCPFCNPJSocio = RetornaNumero(RdoAux2!CPF)
'                    Else
'                        aReg01(s).sCPFCNPJSocio = ""
'                    End If
'                End If
'                aReg01(s).sRGIE = RetornaNumero(Trim(Left(SubNull(RdoAux2!rg), 14)))
'                aReg01(s).sCodLogradouro = xImovel.CodLogradouro
'                aReg01(s).sEndereco = xImovel.Endereco
'                aReg01(s).sNumero = xImovel.Numero
'                aReg01(s).sComplemento = xImovel.Complemento
'                aReg01(s).sCodBairro = xImovel.CodBairro
'                aReg01(s).sBairro = xImovel.Bairro
'                aReg01(s).sCodCidade = xImovel.CodCidade
'                aReg01(s).sCidade = xImovel.Cidade
'                aReg01(s).sCepSocio = RetornaNumero(xImovel.Cep)
'                aReg01(s).sEstadoSocio = xImovel.UF
'                aReg01(s).sClassificao = "Compromissário"
'                aReg01(s).sFoneRes = Left(RetornaNumero(SubNull(RdoAux2!TELEFONE)), 15)
'                aReg01(s).sFoneCom = ""
'                aReg01(s).sCelular = ""
'                aReg01(s).sFoneContato = ""
'                aReg01(s).sEmail = Left(SubNull(RdoAux2!EMAIL), 100)
'                RdoAux2.Close
'            End If
'        ElseIf nCodReduz >= 100000 And nCodReduz < 500000 Then
'            Sql = "SELECT mobiliarioproprietario.codcidadao, cidadao.nomecidadao, cidadao.cnpj,cidadao.cpf,cidadao.rg,cidadao.telefone, cidadao.email FROM mobiliarioproprietario INNER JOIN "
'            Sql = Sql & "cidadao ON mobiliarioproprietario.codcidadao = cidadao.codcidadao Where codmobiliario=" & nCodReduz
'            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'            If RdoAux2.RowCount > 0 Then
'                nCodCidadao = RdoAux2!CodCidadao
'
'                ReDim Preserve aReg01(UBound(aReg01) + 1)
'                s = UBound(aReg01)
'
'                aReg01(s).sTipoReg01 = "01"
'                aReg01(s).sCrcSocio = nCodCidadao
'
'                xImovel.RetornaEndereco nCodCidadao, cidadao, cadastrocidadao
'                aReg01(s).sNomeSocio = RdoAux2!nomecidadao
'                If Trim(SubNull(RdoAux2!Cnpj)) <> "" Then
'                    aReg01(s).sCPFCNPJSocio = RetornaNumero(RdoAux2!Cnpj)
'                Else
'                    If Trim(SubNull(RdoAux2!CPF)) <> "" Then
'                        aReg01(s).sCPFCNPJSocio = RetornaNumero(RdoAux2!CPF)
'                    Else
'                        aReg01(s).sCPFCNPJSocio = ""
'                    End If
'                End If
'                aReg01(s).sRGIE = RetornaNumero(Trim(Left(SubNull(RdoAux2!rg), 14)))
'                aReg01(s).sCodLogradouro = xImovel.CodLogradouro
'                aReg01(s).sEndereco = xImovel.Endereco
'                aReg01(s).sNumero = xImovel.Numero
'                aReg01(s).sComplemento = xImovel.Complemento
'                aReg01(s).sCodBairro = xImovel.CodBairro
'                aReg01(s).sBairro = xImovel.Bairro
'                aReg01(s).sCodCidade = xImovel.CodCidade
'                aReg01(s).sCidade = xImovel.Cidade
'                aReg01(s).sCepSocio = RetornaNumero(xImovel.Cep)
'                aReg01(s).sEstadoSocio = xImovel.UF
'                aReg01(s).sClassificao = "Sócio"
'                aReg01(s).sFoneRes = Left(RetornaNumero(SubNull(RdoAux2!TELEFONE)), 15)
'                aReg01(s).sFoneCom = ""
'                aReg01(s).sCelular = ""
'                aReg01(s).sFoneContato = ""
'                aReg01(s).sEmail = Left(SubNull(RdoAux2!EMAIL), 100)
'                RdoAux2.Close
'            End If
'        End If
        
        '***** REGISTRO 10 **********************************
        ReDim aReg10(0)
        ReDim aReg20(0)
        
        On Error Resume Next
        RdoDebito.Close
        On Error GoTo 0
        If nCadastro = 3 Then
            qd.sql = "{ Call spEXTRATOSERASATAXAS(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
        Else
            qd.sql = "{ Call spEXTRATOSERASA(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
        End If
        qd(0) = nCodReduz
        qd(1) = nCodReduz
        qd(2) = 1990
        qd(3) = Year(Now)
        qd(4) = 1
        qd(5) = 999
        qd(6) = 0
        qd(7) = 999
        qd(8) = 1
        qd(9) = 999
        qd(10) = 0
        qd(11) = 99
        qd(12) = 3
        qd(13) = 3
        qd(14) = Format(Now, "mm/dd/yyyy")
        qd(15) = "Integrativa"
        If bAjuizado Then
            qd(16) = 1
        End If
        Set RdoDebito = qd.OpenResultset(rdOpenKeyset)
        With RdoDebito
            Do Until .EOF
            
                U = UBound(aReg10)
                Achou = False
                For x = 1 To U
                    If Val(aReg10(x).sExercicio) = !AnoExercicio And Val(aReg10(x).sCodigoDivida) = !CodLancamento And Val(aReg10(x).sSubCodDivida) = !SeqLancamento And _
                       Val(aReg10(x).sNumParcela) = !NumParcela And Val(aReg10(x).sSubParcela) = !CODCOMPLEMENTO Then
                        nNumExec = Val(SubNull(!numexecfiscal))
                        nAnoExec = Val(SubNull(!anoexecfiscal))
                        If nAnoExec > 0 Then
                            aReg10(x).sNumExecFiscal = Format(nNumExec, "00000")
                            aReg10(x).sAnoExecFiscal = Format(nAnoExec, "0000")
                        End If
                        Achou = True
                        Exit For
                    End If
                Next
                
                'SE NÃO ENCONTRAR O LANCAMENTO NA MATRIZ, ADICIONAR ELE
                If Not Achou Then
                                                        
                    ReDim Preserve aReg10(UBound(aReg10) + 1)
                    U = UBound(aReg10)
                                        
                    aReg10(U).sTipoReg = "10"
                    aReg10(U).sTributo = !DESCLANCAMENTO
                    aReg10(U).sExercicio = !AnoExercicio
                    aReg10(U).sCodigoDivida = !CodLancamento
                    aReg10(U).sSubCodDivida = !SeqLancamento
                    aReg10(U).sNumParcela = !NumParcela
                    aReg10(U).sSubParcela = !CODCOMPLEMENTO
                    aReg10(U).sSeqInscricaoDA = 0
                    aReg10(U).sNroCDA = Val(SubNull(!certidao))
                    aReg10(U).sFolha = Val(SubNull(!PAGINA))
                    aReg10(U).sLivro = Val(SubNull(!NUMLIVRO))
                    aReg10(U).sDtVencimento = Format(!DataVencimento, "dd/mm/yyyy")
                    If IsDate(!datainscricao) Then
                        aReg10(U).sDtInscricao = Format(!datainscricao, "dd/mm/yyyy")
                    Else
                        aReg10(U).sDtInscricao = ""
                    End If
                    aReg10(U).nPrincipal = RetornaNumero(FormatNumber(!VALORTRIBUTO, 2))
                    aReg10(U).nJuros = RetornaNumero(FormatNumber(!ValorJuros, 2))
                    aReg10(U).nMulta = RetornaNumero(FormatNumber(!ValorMulta, 2))
                    aReg10(U).nCorrecao = RetornaNumero(FormatNumber(!valorcorrecao, 2))
                    aReg10(U).sHonorarios = 0
                    aReg10(U).nTotal = !ValorTotal
                    aReg10(U).nTotalAcumulado = !ValorTotal
'                    aReg10(U).sExecFiscal = sNumExec
                    nNumExec = Val(SubNull(!numexecfiscal))
                    nAnoExec = Val(SubNull(!anoexecfiscal))
                    If nAnoExec > 0 Then
                        aReg10(U).sNumExecFiscal = Format(nNumExec, "00000")
                        aReg10(U).sAnoExecFiscal = Format(nAnoExec, "0000")
                    End If
                Else
                    aReg10(x).nTotal = aReg10(U).nTotal + !ValorTotal
                    aReg10(x).nPrincipal = aReg10(x).nPrincipal + RetornaNumero(FormatNumber(!VALORTRIBUTO, 2))
                    aReg10(x).nJuros = aReg10(x).nJuros + RetornaNumero(FormatNumber(!ValorJuros, 2))
                    aReg10(x).nMulta = aReg10(x).nMulta + RetornaNumero(FormatNumber(!ValorMulta, 2))
                    aReg10(x).nCorrecao = aReg10(x).nCorrecao + RetornaNumero(FormatNumber(!valorcorrecao, 2))
                    aReg10(x).nTotalAcumulado = aReg10(x).nTotalAcumulado + !ValorTotal
                    
                End If
                
                '***** FUNDAMENTO/ARTIGO ****************************
                sql = "SELECT tributo.codtributo, tributo.desctributo, tributo.livro, tributoartigo.artigo FROM tributo LEFT OUTER JOIN "
                sql = sql & "tributoartigo ON tributo.codtributo = tributoartigo.codtributo where tributo.codtributo=" & !CodTributo
                Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                If Not IsNull(RdoAux3!ARTIGO) Then
                    aReg10(U).sFundamento = aReg10(U).sFundamento & Replace(Replace(RdoAux3!ARTIGO, vbLf, ""), vbCr, "") & "|"
                End If
                aReg10(U).sNatureza = IIf(RdoAux3!livro = 3, "Não Tributária", "Tributária")
                RdoAux3.Close
                If aReg10(U).sFundamento <> "" Then
                    aReg10(U).sFundamento = Left(aReg10(U).sFundamento, Len(aReg10(U).sFundamento) - 1)
                End If
                
                aReg10(U).sCartorio = ""
                aReg10(U).sProcessoAdm = ""
                aReg10(U).sAutoInfracao = ""
                aReg10(U).sNroParcelamento = ""
                
                aReg10(U).sCodLogradouro = aReg00(t).sCodLogradouroLocal
                aReg10(U).sEndereco = aReg00(t).sEnderecoLocal
                aReg10(U).sNumero = aReg00(t).sNumeroLocal
                aReg10(U).sQuadra = aReg00(t).sQuadra
                aReg10(U).sLote = aReg00(t).sLote
                aReg10(U).sCep = aReg00(t).sCEPLocal
                aReg10(U).sCodBairro = aReg00(t).sCodBairroLocal
                aReg10(U).sBairro = aReg00(t).sBairroLocal
                aReg10(U).sAtividade = aReg00(t).sAtividade
                aReg10(U).sMatricula = aReg00(t).sMatricula
                
                ReDim Preserve aReg20(UBound(aReg20) + 1)
                v = UBound(aReg20)
                                    
                aReg20(v).sTipoReg = "20"
                aReg20(v).sExercicio = !AnoExercicio
                aReg20(v).sCodigoDivida = !CodLancamento
                aReg20(v).sSubCodDivida = !SeqLancamento
                aReg20(v).sNumParcela = !NumParcela
                aReg20(v).sSubParcela = !CODCOMPLEMENTO
                aReg20(v).sCodReceita = !CodTributo
                aReg20(v).sDescricao = !abrevTributo
                aReg20(v).sValorOriginal = RetornaNumero(FormatNumber(!VALORTRIBUTO, 2))
PROXIMODEBITO:
                DoEvents
               .MoveNext
            Loop
           .Close
        End With

        
        '****************************************************
        If UBound(aReg10) = 0 Then GoTo Proximo
        With aReg00(t)
            ax = cf(.sTipoReg00, Character, 2) & cf(.sCrc, Numeric, 20) & cf(.sNome, Character, 50) & cf(.sCPFCNPJ, Numeric, 14) & cf(Left(.sRGIE, 14), Numeric, 14) & cf(.sCodLogradouro, Numeric, 4) & cf(.sEndereco, Character, 60) & cf(.sNumero, Numeric, 5) & cf(.sComplemento, Character, 25)
            ax = ax & cf(.sCodBairro, Numeric, 4) & cf(.sBairro, Character, 30) & cf(.sCodCidade, Numeric, 4) & cf(.sCidade, Character, 30) & cf(.sCep, Numeric, 8) & cf(.sEstado, Character, 2) & cf(RetornaNumero(FormatNumber(.nValorTotal, 2)), Numeric, 14) & cf(.sDataAtualizacao, Character, 10)
            ax = ax & cf(.sIdAjuizamento, Numeric, 8) & cf(.sDescCadastro, Character, 20) & cf(.sCodCadastro, Numeric, 20) & cf(.sCadastroImob, Character, 25) & cf(.sFoneRes, Numeric, 15) & cf(.sFoneCom, Numeric, 15) & cf(.sCelular, Numeric, 15) & cf(.sFoneContato, Numeric, 15) & cf(.sEmail, Character, 100)
            Print #1, ax
        End With
        

'        For s = 1 To UBound(aReg01)
'            With aReg01(s)
'                ax = cf(.sTipoReg01, Character, 2) & cf(.sCrcSocio, Numeric, 20) & cf(.sNomeSocio, Character, 50) & cf(.sCPFCNPJSocio, Numeric, 14) & cf(.sRGIE, Numeric, 14) & cf(.sCodLogradouro, Numeric, 4) & cf(.sEndereco, Character, 60) & cf(.sNumero, Numeric, 5) & cf(.sComplemento, Character, 25)
'                ax = ax & cf(.sCodBairro, Numeric, 4) & cf(.sBairro, Character, 30) & cf(.sCodCidade, Numeric, 4) & cf(.sCidade, Character, 30) & cf(.sCepSocio, Numeric, 8) & cf(.sEstadoSocio, Character, 2) & cf(.sClassificao, Character, 15) & cf(.sFoneRes, Numeric, 15) & cf(.sFoneCom, Numeric, 15)
'                ax = ax & cf(.sCelular, Numeric, 15) & cf(.sFoneContato, Numeric, 15) & cf(.sEmail, Character, 100)
'                Print #1, ax
'            End With
'        Next

        For U = 1 To UBound(aReg10)
            With aReg10(U)
                ax = cf(.sTipoReg, Character, 2) & cf(.sTributo, Character, 30) & cf(.sExercicio, Numeric, 4) & cf(.sCodigoDivida, Numeric, 4) & cf(.sSubCodDivida, Numeric, 4) & cf(.sNumParcela, Numeric, 2) & cf(.sSubParcela, Numeric, 2) & cf(.sSeqInscricaoDA, Numeric, 8) & cf(.sNroCDA, Character, 10)
                ax = ax & cf(.sFolha, Numeric, 5) & cf(.sLivro, Numeric, 4) & cf(.sDtVencimento, Character, 10) & cf(.sDtInscricao, Character, 10) & cf(.nPrincipal, Numeric, 14) & cf(.nJuros, Numeric, 14) & cf(.nMulta, Numeric, 14) & cf(.nCorrecao, Numeric, 14) & cf(.sHonorarios, Numeric, 14)
                ax = ax & cf(RetornaNumero(FormatNumber(.nTotalAcumulado, 2)), Numeric, 14) & cf(.sCodLogradouro, Numeric, 4) & cf(.sEndereco, Character, 60) & cf(.sNumero, Numeric, 5) & cf(.sQuadra, Character, 10) & cf(.sLote, Character, 20) & cf(.sCep, Numeric, 8) & cf(.sCodBairro, Numeric, 4) & cf(.sBairro, Character, 30) & cf(.sAtividade, Character, 80)
                ax = ax & cf(.sMatricula, Character, 20) & cf(.sCartorio, Character, 50) & cf(.sProcessoAdm, Character, 20) & cf(.sAutoInfracao, Character, 20) & cf(.sNatureza, Character, 20) & cf(.sNroParcelamento, Character, 20) & cf(Left(.sFundamento, 1000), Character, 1000) & cf(.sNumExecFiscal, Character, 5) & cf(.sAnoExecFiscal, Character, 4)
                Print #1, ax
                For v = 1 To UBound(aReg20)
                    If .sExercicio = aReg20(v).sExercicio And .sCodigoDivida = aReg20(v).sCodigoDivida And .sSubCodDivida = aReg20(v).sSubCodDivida And .sNumParcela = aReg20(v).sNumParcela And .sSubParcela = aReg20(v).sSubParcela Then
                        With aReg20(v)
                            ax = cf(.sTipoReg, Character, 2) & cf(.sExercicio, Numeric, 4) & cf(.sCodReceita, Numeric, 4) & cf(.sDescricao, Character, 20) & cf(.sValorOriginal, Numeric, 14)
                            Print #1, ax
                        End With
                    End If
                Next
            End With
        Next

Proximo:
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

Close #1
Liberado

PBar.value = 0
PBar.Color = vbWhite
MsgBox "Arquivo finalizado.", vbInformation, "Informação"

Exit Sub
Erro:

Liberado

MsgBox Err.Description
Resume Next
End Sub

Private Sub ExportarCDA()
Dim sql As String, RdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, nCodReduz As Long, sSetor As String, sNome As String
Dim RdoAux3 As rdoResultset, nCDA As Long
Exit Sub
ConectaIntegrativa

sql = "delete from CDADebitos"
cnInt.Execute sql, rdExecDirect
sql = "delete from CDAs"
cnInt.Execute sql, rdExecDirect

nPos = 0

sql = "select distinct codreduzido from debitoparcela where numexecfiscal is not null order by codreduzido"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
nTot = RdoAux.RowCount
Do Until RdoAux.EOF
    If nPos Mod 10 = 0 Then
        CallPb nPos, nTot
    End If

    If RdoAux!CODREDUZIDO < 100000 Then
        sSetor = "IMOBILIÁRIO"
'        Sql = "select nomecidadao as nome from vwfullimovel2 where codreduzido=" & RdoAux!CODREDUZIDO
    ElseIf RdoAux!CODREDUZIDO >= 100000 And RdoAux!CODREDUZIDO < 500000 Then
        sSetor = "MOBILIÁRIO"
 '       Sql = "select razaosocial as nome from mobiliario where codigomob=" & RdoAux!CODREDUZIDO
    Else
        sSetor = "TAXAS"
  '      Sql = "select nomecidadao as nome from cidadao where codcidadao=" & RdoAux!CODREDUZIDO
    End If
   ' Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   ' sNome = RdoAux2!nome
   ' RdoAux2.Close

    sql = "select * from debitoparcela where codreduzido=" & RdoAux!CODREDUZIDO & " and numexecfiscal is not null"
    Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        
        Do Until .EOF
            If IsNull(!datainscricao) Then GoTo Proximo
            
            sql = "insert CDAs(idDevedor, SetorDevedor, DtInscricao, NroCertidao, NroLivro, NroFolha, NroOrdem, DtGeracao) values("
            sql = sql & !CODREDUZIDO & ",'" & sSetor & "','" & Format(!datainscricao, "mm/dd/yyyy") & "'," & Val(SubNull(!numcertidao)) & "," & !numerolivro & ","
            sql = sql & Val(SubNull(!paginalivro)) & ",'" & CStr(!numexecfiscal & "/" & !anoexecfiscal) & "','" & Format(Now, "mm/dd/yyyy") & "')"
            cnInt.Execute sql, rdExecDirect
            
            sql = "select @@identity as LastKey"
            Set RdoAux3 = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            nCDA = RdoAux3!lastkey
            RdoAux3.Close
            
            sql = "select debitotributo.*,tributo.abrevtributo from debitotributo INNER JOIN tributo ON debitotributo.codtributo = tributo.codtributo where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and codlancamento=" & !CodLancamento & " and "
            sql = sql & "seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO & " and debitotributo.codtributo<>3"
            Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux3
                Do Until .EOF
                    sql = "insert CDADebitos(idCDA, CodTributo, Tributo, Exercicio, Lancamento, Seq, NroParcela, ComplParcela, DtVencimento, vlrOriginal, DtGeracao) values("
                    sql = sql & nCDA & "," & !CodTributo & ",'" & Mask(!abrevTributo) & "'," & !AnoExercicio & "," & !CodLancamento & "," & !SeqLancamento & "," & !NumParcela & ","
                    sql = sql & !CODCOMPLEMENTO & ",'" & Format(RdoAux2!DataVencimento, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(!VALORTRIBUTO)) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
                    cnInt.Execute sql, rdExecDirect
                   .MoveNext
                Loop
               .Close
            End With
            
Proximo:
            
           .MoveNext
        Loop
       .Close
    End With
    nPos = nPos + 1
    RdoAux.MoveNext
    DoEvents
Loop

End Sub


Private Sub Protestar()
'Dim sql As String, RdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, nCodReduz As Long, RdoAux3 As rdoResultset
'Dim t As Integer, s As Integer, U As Integer, v As Integer, cmdComm As ADODB.Command, clsImovel As New clsImovel
'Dim RdoDebito As rdoResultset, qd As New rdoQuery, Achou As Boolean, x As Integer, RdoTrib As rdoResultset
'Dim Rs As ADODB.Recordset, strQuery As String, nNumExec As Long, nAnoExec As Integer, sNumExec As String
'
'Dim sNomeCidadao As String, sInscricao As String, sFoneRes As String, sFoneCom As String, sCelular As String, sFoneContato As String, sEmail As String, nNumeroLocal As Integer
'Dim sCPFCNPJ As String, nCodLogradouroLocal As Integer, sEnderecoLocal As String, sComplementoLocal As String, sNumeroLocal As String, nCodBairroLocal As Integer, sBairroLocal As String, sCidadeLocal As String
'Dim sCEPLocal As String, sQuadra As String, sLote As String, sAtividade As String, sMatricula As String, sRGIE As String, nCodLogradouro As Integer, sUFLocal As String
'Dim sEndereco As String, nNumero As Integer, sComplemento As String, nCodBairro As Integer, nCodCidade As Integer, sBairro As String, sCidade As String, sCep As String, sUF As String
'Dim SetorDevedor As String, DtInscricao As String, NroCertidao As Integer, NroLivro As Integer, NroFolha As Integer, NroOrdem As Integer
'Dim nCDA As Long, nAno As Integer, nLanc As Integer, nSeqLanc As Integer, nParc As Integer, nCompl As Integer, nCodTrib As Integer, sDescTrib As String
'Dim dtVencto As String, nValorPrincipal As Double, nCodCidadao As Long, sClassificacao As String
'Dim aCda() As typeCDA, aCdaDebito() As typeCDADebito, aParte() As Reg01
'
Exit Sub
'Ocupado
'ConectaIntegrativa
'sql = "delete from cdadebitos_protesto"
''cnInt.Execute Sql, rdExecDirect
'
'sql = "delete from cadastro_protesto"
''cnInt.Execute Sql, rdExecDirect
'
'sql = "delete from partes_protesto"
''cnInt.Execute Sql, rdExecDirect
'
'sql = "delete from cdas_protesto"
''cnInt.Execute Sql, rdExecDirect
'
'cmdExec.Enabled = False
'Set qd.ActiveConnection = cn
'qd.QueryTimeout = 0
'
'If cmbCadastro.ListIndex = 0 Then
'    MsgBox "selecione um cadastro"
'    Exit Sub
'End If
''Sql = "SELECT  DISTINCT CODREDUZIDO FROM DEBITOPARCELA WHERE DATAVENCIMENTO<='" & Format(mskDataVencto.Text, "mm/dd/yyyy") & "' AND STATUSLANC =3 AND NUMPARCELA>0 AND DATAAJUIZA IS NULL AND DATAINSCRICAO IS NOT NULL AND PROTESTO_DATA_REMESSA IS NULL"
'sql = "SELECT  DISTINCT CODREDUZIDO FROM DEBITOPARCELA WHERE (anoexercicio = 2020) and DATAVENCIMENTO<='" & Format(mskDataVencto.Text, "mm/dd/yyyy") & "' AND (STATUSLANC =3 or STATUSLANC =42 or STATUSLANC =43 or STATUSLANC =38 or STATUSLANC =39  ) AND NUMPARCELA>0 AND DATAAJUIZA IS NULL AND DATAINSCRICAO IS NOT NULL AND PROTESTO_DATA_REMESSA IS NULL"
'If cmbCadastro.ListIndex = 1 Then
'    'Sql = Sql & " AND CODREDUZIDO BETWEEN 1 AND 40000 "
'    sql = sql & " AND CODREDUZIDO = 9708 "
''    sql = sql & " AND CODLANCAMENTO <>5 AND CODLANCAMENTO<>20 AND CODLANCAMENTO<>11 "
'    sql = sql & " AND CODLANCAMENTO =1 "
'ElseIf cmbCadastro.ListIndex = 2 Then
'    'sql = sql & " AND CODREDUZIDO BETWEEN 100000 AND 300000 "
'    'Sql = Sql & " AND CODREDUZIDO in (123098,123126,123130,123136,123095 ) "
'    sql = sql & " AND CODLANCAMENTO not in (11,20)  "
'ElseIf cmbCadastro.ListIndex = 3 Then
'    'Sql = Sql & " AND CODREDUZIDO BETWEEN 500000 AND 699999 "
'    sql = sql & " AND CODREDUZIDO in (547663,614016) "
'    sql = sql & " AND CODLANCAMENTO not in(20) "
'   ' Sql = Sql & " AND CODLANCAMENTO in(74) "
'End If
''sql = sql & " AND CODREDUZIDO in (12765, 626025, 112620, 114743, 31219, 30919, 116520, 10175, 121748, 112360, 13254, 631167, 564864, 625075, 574703, 534021, 566119, 17538, 38754, 539125, 562433, 608855, 552240, 619284, 617541,124420, 630231, 573929,630072,507527, 111036, 109379, 116235, 119892, 117975, 117221, 117412, 117976, 113334, 118020,118977, 116134, 123688,119538, 25692, 15580, 626937, 36604, 547441, 35712, 591741, 33053, 33158, 32915, 112064, 106396, 123576, 124080, 124086, 124081, 124082, 33083, 120732, 12026, 627717, 118531, 121579, 116290, 123653, 123275, 123655, 121371, 123640, 122938, 121374, 120141, 119599, 119664, 120435, 123234, 118151, 117193, 121708, 121284,122127,16674,8014,124750, 121155, 120831, 122617, 118272) "
'sql = sql & " ORDER BY CODREDUZIDO"
'
''***********
'
'Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    nPos = 1
'    nTot = .RowCount
'    Do Until .EOF
'        nCodReduz = !CODREDUZIDO
'        If nCodReduz < 100000 Then
'            SetorDevedor = "Imobiliário"
'        ElseIf nCodReduz >= 100000 And nCodReduz < 500000 Then
'            SetorDevedor = "Mobiliário"
'        Else
'            SetorDevedor = "Cidadão"
'        End If
'
'        If nPos Mod 10 = 0 Then
'          '  GoTo fim
'            DoEvents
'            CallPb nPos, nTot
'        End If
'
'        '*** debito ***
'        ReDim aParte(0)
'        ReDim aCda(0)
'        ReDim aCdaDebito(0)
'
'        On Error Resume Next
'        RdoDebito.Close
'        On Error GoTo 0
'        If cmbCadastro.ListIndex = 3 Then
'            qd.sql = "{ Call spEXTRATOAJUIZARTAXA(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
'        Else
'            qd.sql = "{ Call spEXTRATOAJUIZAR(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
'        End If
'        qd(0) = nCodReduz
'        qd(1) = nCodReduz
'        qd(2) = 2019
'        qd(3) = 2019
'        qd(4) = 0
'        qd(5) = 99
'        qd(6) = 0
'        qd(7) = 999
'        qd(8) = 1
'        qd(9) = 999
'        qd(10) = 0
'        qd(11) = 99
'        qd(12) = 3
'        qd(13) = 3
'        qd(14) = Format("27/12/2024", "mm/dd/yyyy")
'        qd(15) = "Protesto"
'        qd(16) = 0
'        Set RdoDebito = qd.OpenResultset(rdOpenKeyset)
'        With RdoDebito
'            Do Until .EOF
''                If Year(!DataVencimento) < 2015 Then
''                    GoTo PROXIMODEBITO
''                End If
'                'If !CodLancamento <> 79 Then GoTo PROXIMODEBITO
''                If !CODREDUZIDO = 23715 Then
''                    MsgBox "tsted"
''                End If
'                U = UBound(aCda)
'                Achou = False
'                For x = 1 To U
'                    If Val(aCda(x).nAno) = !AnoExercicio And Val(aCda(x).nLanc) = !CodLancamento And Val(aCda(x).nSeq) = !SeqLancamento And _
'                       Val(aCda(x).nParc) = !NumParcela And Val(aCda(x).nCompl) = !CODCOMPLEMENTO Then
'                        Achou = True
'                        Exit For
'                    End If
'                Next
'
'                If Not Achou Then
'                    ReDim Preserve aCda(UBound(aCda) + 1)
'                    U = UBound(aCda)
'                    aCda(U).nAno = !AnoExercicio
'                    aCda(U).nLanc = !CodLancamento
'                    aCda(U).nSeq = !SeqLancamento
'                    aCda(U).nParc = !NumParcela
'                    aCda(U).nCompl = !CODCOMPLEMENTO
'                    On Error Resume Next
'                    aCda(U).nNumCertidao = Val(SubNull(!CERTIDAO))
'                    On Error GoTo 0
'                    aCda(U).nNumPagina = Val(SubNull(!PAGINA))
'                    aCda(U).nNumLivro = Val(SubNull(!NUMLIVRO))
'                    aCda(U).dDataInscricao = !datainscricao
'                End If
'
'
'                ReDim Preserve aCdaDebito(UBound(aCdaDebito) + 1)
'                U = UBound(aCdaDebito)
'                aCdaDebito(U).nAno = !AnoExercicio
'                aCdaDebito(U).nLanc = !CodLancamento
'                aCdaDebito(U).nSeq = !SeqLancamento
'                aCdaDebito(U).nParc = !NumParcela
'                aCdaDebito(U).nCompl = !CODCOMPLEMENTO
'                aCdaDebito(U).nCodTributo = !CodTributo
'                aCdaDebito(U).sDescTributo = !abrevTributo
'                aCdaDebito(U).nPrincipal = !VALORTRIBUTO
'                aCdaDebito(U).nMulta = !ValorMulta
'                aCdaDebito(U).nJuros = !ValorJuros
'                aCdaDebito(U).nCorrecao = !valorcorrecao
'                aCdaDebito(U).nMulta = !ValorMulta
'                aCdaDebito(U).nJuros = !ValorJuros
'                aCdaDebito(U).nCorrecao = !valorcorrecao
'                aCdaDebito(U).nTotal = !ValorTotal
'                aCdaDebito(U).dDataVencto = !DataVencimento
'PROXIMODEBITO:
'                DoEvents
'               .MoveNext
'            Loop
'           .Close
'        End With
'
'        If UBound(aCda) = 0 Then GoTo Proximo
'
'        For x = 1 To UBound(aCda)
'            With aCda(x)
'                sql = "insert CDAs_protesto(idDevedor, SetorDevedor, DtInscricao, NroCertidao, NroLivro, NroFolha,  DtGeracao) values("
'                sql = sql & nCodReduz & ",'" & SetorDevedor & "','" & Format(.dDataInscricao, "mm/dd/yyyy") & "'," & .nNumCertidao & "," & .nNumLivro & ","
'                sql = sql & .nNumPagina & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "')"
'                cnInt.Execute sql, rdExecDirect
'
'                sql = "select @@identity as LastKey"
'                Set RdoAux3 = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
'                nCDA = RdoAux3!lastkey
'               .nCDA = nCDA
'                RdoAux3.Close
'
'                For v = 1 To UBound(aCdaDebito)
'                    If aCdaDebito(v).nAno = .nAno And aCdaDebito(v).nLanc = .nLanc And aCdaDebito(v).nSeq = .nSeq And _
'                       aCdaDebito(v).nParc = .nParc And aCdaDebito(v).nCompl = .nCompl Then
'                        aCdaDebito(v).nCDA = .nCDA
'                    End If
'                Next
'            End With
'        Next
'
'        For x = 1 To UBound(aCdaDebito)
'            With aCdaDebito(x)
'                sql = "insert CDADebitos_protesto(idCDA, CodTributo, Tributo, Exercicio, Lancamento, Seq, NroParcela, ComplParcela, DtVencimento, vlrOriginal,vlrMultas,vlrjuros,vlrcorrecao, DtGeracao) values("
'                sql = sql & .nCDA & "," & .nCodTributo & ",'" & Mask(.sDescTributo) & "'," & .nAno & "," & .nLanc & "," & .nSeq & "," & .nParc & "," & .nCompl & ",'"
'                sql = sql & Format(.dDataVencto, "mm/dd/yyyy") & "'," & Virg2Ponto(Format(.nPrincipal, "#0.00")) & "," & Virg2Ponto(Format(.nMulta, "#0.00")) & "," & Virg2Ponto(Format(.nJuros, "#0.00")) & "," & Virg2Ponto(Format(.nCorrecao, "#0.00")) & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "')"
'                cnInt.Execute sql, rdExecDirect
'            End With
'        Next
'
'       '*** dados cadastrais
'        If nCodReduz < 100000 Then
'            sql = "select * from vwfullimovel2 where codreduzido=" & nCodReduz
'            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
'            With RdoAux2
'                If .RowCount = 0 Then GoTo Proximo
'                sNome = !nomecidadao
'                sInscricao = !Inscricao
'                sFoneRes = Left(RetornaNumero(SubNull(!telefone)), 15)
'                sFoneCom = ""
'                sCelular = ""
'                sFoneContato = ""
'                sEmail = Left(SubNull(!Email), 100)
'                If Trim(SubNull(!Cnpj)) <> "" Then
'                    sCPFCNPJ = RetornaNumero(!Cnpj)
'                Else
'                    If Trim(SubNull(!cpf)) <> "" Then
'                        sCPFCNPJ = RetornaNumero(!cpf)
'                    Else
'                        sCPFCNPJ = ""
'                    End If
'                End If
'
'                xImovel.RetornaEndereco nCodReduz, Imobiliario, Localizacao
'                nCodLogradouroLocal = xImovel.CodLogradouro
'
'                sEnderecoLocal = xImovel.Endereco
'                nNumeroLocal = xImovel.Numero
'                sComplementoLocal = xImovel.Complemento
'                nCodBairroLocal = xImovel.CodBairro
'                sBairroLocal = xImovel.Bairro
'                sCidadeLocal = xImovel.Cidade
'                sUFLocal = xImovel.UF
'                sCEPLocal = xImovel.Cep
'                If Val(nCodBairroLocal) = 999 Then
'                    nCodBairroLocal = 0
'                    sBairroLocal = ""
'                End If
'
'                sQuadra = !Quadra
'                sLote = !Lote
'                sAtividade = ""
'                sMatricula = Val(SubNull(!NumMat))
'
'                sRGIE = RetornaNumero(Trim(SubNull(!rg)))
'                If !Ee_TipoEnd = 0 Then 'imovel
'                    xImovel.RetornaEndereco nCodReduz, Imobiliario, Localizacao
'                    nCodLogradouro = xImovel.CodLogradouro
'                    sEndereco = xImovel.Endereco
'                    nNumero = xImovel.Numero
'                    sComplemento = xImovel.Complemento
'                    nCodBairro = xImovel.CodBairro
'                    sBairro = xImovel.Bairro
'                    nCodCidade = xImovel.CodCidade
'                    sCidade = xImovel.Cidade
'                    sCep = xImovel.Cep
'                    sUF = xImovel.UF
'                ElseIf !Ee_TipoEnd = 1 Then 'proprietario
'                    xImovel.RetornaEndereco nCodReduz, Imobiliario, cadastrocidadao
'                    nCodLogradouro = Val(xImovel.CodLogradouro)
'                    sEndereco = xImovel.Endereco
'                    nNumero = xImovel.Numero
'                    sComplemento = xImovel.Complemento
'                    nCodBairro = Val(xImovel.CodBairro)
'                    sBairro = xImovel.Bairro
'                    nCodCidade = xImovel.CodCidade
'                    sCidade = xImovel.Cidade
'                    sCep = xImovel.Cep
'                    sUF = xImovel.UF
'                Else 'entrega
'                    xImovel.RetornaEndereco nCodReduz, Imobiliario, Entrega
'                    nCodLogradouro = xImovel.CodLogradouro
'                    sEndereco = xImovel.Endereco
'                    nNumero = xImovel.Numero
'                    sComplemento = xImovel.Complemento
'                    nCodBairro = Val(xImovel.CodBairro)
'                    sBairro = xImovel.Bairro
'                    nCodCidade = xImovel.CodCidade
'                    sCidade = xImovel.Cidade
'                    sCep = xImovel.Cep
'                    sUF = xImovel.UF
'                End If
'               .Close
'            End With
'        ElseIf nCodReduz >= 100000 And nCodReduz < 500000 Then
'            sql = "select * from vwfullempresa3 where codigomob=" & nCodReduz
'            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
'            With RdoAux2
'                If .RowCount = 0 Then GoTo Proximo
'                sInscricao = ""
'                sNome = !RazaoSocial
'                sFoneRes = ""
'                sFoneCom = ""
'                sCelular = ""
'                sFoneContato = Left(RetornaNumero(SubNull(!fonecontato)), 15)
'                sEmail = Left(SubNull(!emailcontato), 100)
'                If Trim(SubNull(!Cnpj)) <> "" Then
'                    sCPFCNPJ = RetornaNumero(!Cnpj)
'                Else
'                    If Trim(SubNull(!cpf)) <> "" Then
'                        sCPFCNPJ = RetornaNumero(!cpf)
'                    Else
'                        sCPFCNPJ = ""
'                    End If
'                End If
'                sRGIE = RetornaNumero(Trim(SubNull(!inscestadual)))
'                If sRGIE = "" Then
'                    sRGIE = Trim(SubNull(!rg) & " " & SubNull(!Orgao))
'                End If
'
'                xImovel.RetornaEndereco nCodReduz, Mobiliario, Localizacao
'                nCodLogradouroLocal = xImovel.CodLogradouro
'                sEnderecoLocal = xImovel.Endereco
'                nNumeroLocal = xImovel.Numero
'                sComplementoLocal = xImovel.Complemento
'                nCodBairroLocal = Val(xImovel.CodBairro)
'                sBairroLocal = xImovel.Bairro
'                sCidadeLocal = xImovel.Cidade
'                sUFLocal = xImovel.UF
'                sCEPLocal = xImovel.Cep
'                If Val(nCodBairroLocal) = 999 Then
'                    nCodBairroLocal = 0
'                    sBairroLocal = ""
'                End If
'
'                sEndereco = SubNull(!eenomelogr)
'                If sEndereco = "" Then
'                    'local da empresa
'                    xImovel.RetornaEndereco nCodReduz, Mobiliario, Localizacao
'                    nCodLogradouro = xImovel.CodLogradouro
'                    sEndereco = xImovel.Endereco
'                    nNumero = xImovel.Numero
'                    sComplemento = xImovel.Complemento
'                    nCodBairro = Val(xImovel.CodBairro)
'                    sBairro = xImovel.Bairro
'                    nCodCidade = xImovel.CodCidade
'                    sCidade = xImovel.Cidade
'                    sCep = xImovel.Cep
'                    sUF = xImovel.UF
'                Else
''                   'endereco entrega
'                    xImovel.RetornaEndereco nCodReduz, Mobiliario, Entrega
'                    nCodLogradouro = xImovel.CodLogradouro
'                    sEndereco = xImovel.Endereco
'                    nNumero = xImovel.Numero
'                    sComplemento = xImovel.Complemento
'                    nCodBairro = Val(xImovel.CodBairro)
'                    sBairro = xImovel.Bairro
'                    nCodCidade = xImovel.CodCidade
'                    sCidade = xImovel.Cidade
'                    sCep = xImovel.Cep
'                    sUF = xImovel.UF
'                End If
'
'                sQuadra = ""
'                sLote = ""
'                sAtividade = Left(SubNull(!ativextenso), 80)
'                sMatricula = ""
'
'               .Close
'            End With
'        Else
'            sql = "select * from vwfullcidadao where codcidadao=" & nCodReduz
'            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
'            With RdoAux2
'                If .RowCount = 0 Then GoTo Proximo
'                sInscricao = ""
'                sNome = !nomecidadao
'                sFoneRes = Left(RetornaNumero(SubNull(!telefone)), 15)
'                sFoneCom = ""
'                sCelular = ""
'                sFoneContato = ""
'                sEmail = Left(SubNull(!Email), 100)
'                If Trim(SubNull(!Cnpj)) <> "" Then
'                    sCPFCNPJ = RetornaNumero(!Cnpj)
'                Else
'                    If Trim(SubNull(!cpf)) <> "" Then
'                        sCPFCNPJ = RetornaNumero(!cpf)
'                    Else
'                        sCPFCNPJ = ""
'                    End If
'                End If
'                sRGIE = RetornaNumero(Trim(Left(SubNull(!rg), 14)))
'                xImovel.RetornaEndereco nCodReduz, cidadao, cadastrocidadao
'                nCodLogradouro = Val(xImovel.CodLogradouro)
'                sEndereco = SubNull(!Endereco)
'                nNumero = SubNull(!NUMIMOVEL)
'                sComplemento = xImovel.Complemento
'                nCodBairro = Val(SubNull(!CodBairro))
'                sBairro = SubNull(!DescBairro)
'                If nCodBairro = 999 Then
'                    nCodBairro = 0
'                    sBairro = ""
'                End If
'                nCodCidade = SubNull(!CodCidade)
'                sCidade = SubNull(!descCidade)
'                sUFLocal = xImovel.UF
'                sCep = RetornaNumero(SubNull(!Cep))
'                sUF = SubNull(!SiglaUF)
'
'                nCodLogradouroLocal = Val(xImovel.CodLogradouro)
'                sEnderecoLocal = SubNull(!Endereco)
'                nNumeroLocal = SubNull(!NUMIMOVEL)
'                nCodBairroLocal = Val(SubNull(!CodBairro))
'                sBairroLocal = SubNull(!DescBairro)
'                If Val(nCodBairroLocal) = 999 Then
'                    nCodBairroLocal = ""
'                    sBairroLocal = ""
'                End If
'                sCidadeLocal = xImovel.Cidade
'                sCEPLocal = RetornaNumero(SubNull(!Cep))
'
'                sQuadra = ""
'                sLote = ""
'                sAtividade = ""
'                sMatricula = ""
'
'               .Close
'            End With
'        End If
'
'        For x = 1 To UBound(aCda)
'            With aCda(x)
'                sql = "INSERT Cadastro_protesto(IdCDA,SetorDevedor,Crc,Nome,Inscricao,CPFCnpj,RgInscrEstadual,LocalCep,LocalEndereco,LocalNumero,LocalComplemento,"
'                sql = sql & "LocalBairro,LocalCidade,LocalEstado,Quadra,Lote,EntregaCep,EntregaEndereco,EntregaNumero,EntregaComplemento,EntregaBairro,"
'                sql = sql & "EntregaCidade,EntregaEstado,DtGeracao) values("
'                sql = sql & .nCDA & ",'" & SetorDevedor & "'," & nCodReduz & ",'" & Mask(Left(SubNull(sNome), 80)) & "','" & sInscricao & "','" & sCPFCNPJ & "','" & sRGIE & "','"
'                sql = sql & sCEPLocal & "','" & sEnderecoLocal & "'," & nNumeroLocal & ",'" & Mask(Left(sComplementoLocal, 50)) & "','" & sBairroLocal & "','" & sCidadeLocal & "','" & sUFLocal & "','"
'                sql = sql & sQuadra & "','" & sLote & "','" & sCep & "','" & sEndereco & "'," & nNumero & ",'" & Mask(Left(sComplemento, 50)) & "','"
'                sql = sql & sBairro & "','" & Mask(sCidade) & "','" & sUF & "','" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "')"
'                cnInt.Execute sql, rdExecDirect
'            End With
'        Next
'
'
'        '***** PARTES **********************************
'
'        If nCodReduz < 100000 Then
'            sql = "SELECT proprietario.codcidadao, cidadao.nomecidadao, cidadao.cnpj,cidadao.cpf,cidadao.rg,cidadao.telefone, cidadao.email FROM proprietario INNER JOIN "
'            sql = sql & "cidadao ON proprietario.codcidadao = cidadao.codcidadao Where codreduzido=" & nCodReduz & " and principal=0"
'            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
'            With RdoAux2
'                Do Until .EOF
'                    ReDim Preserve aParte(UBound(aParte) + 1)
'                    U = UBound(aParte)
'                    xImovel.RetornaEndereco nCodCidadao, cidadao, cadastrocidadao
'                    aParte(U).sNomeSocio = RdoAux2!nomecidadao
'                    If Trim(SubNull(RdoAux2!Cnpj)) <> "" Then
'                        aParte(U).sCPFCNPJSocio = RetornaNumero(RdoAux2!Cnpj)
'                    Else
'                        If Trim(SubNull(RdoAux2!cpf)) <> "" Then
'                            aParte(U).sCPFCNPJSocio = RetornaNumero(RdoAux2!cpf)
'                        Else
'                            aParte(U).sCPFCNPJSocio = ""
'                        End If
'                    End If
'                    aParte(U).sRGIE = RetornaNumero(Trim(Left(SubNull(RdoAux2!rg), 14)))
'                    aParte(U).sCodLogradouro = xImovel.CodLogradouro
'                    aParte(U).sEstadoSocio = xImovel.Endereco
'                    aParte(U).sNumero = xImovel.Numero
'                    aParte(U).sComplemento = xImovel.Complemento
'                    aParte(U).sCodBairro = xImovel.CodBairro
'                    aParte(U).sBairro = xImovel.Bairro
'                    aParte(U).sCodCidade = xImovel.CodCidade
'                    aParte(U).sCidade = xImovel.Cidade
'                    aParte(U).sCepSocio = RetornaNumero(xImovel.Cep)
'                    aParte(U).sEstadoSocio = xImovel.UF
'                    aParte(U).sClassificao = "Compromissário"
'                    aParte(U).sFoneRes = Left(RetornaNumero(SubNull(RdoAux2!telefone)), 15)
'                    aParte(U).sFoneCom = ""
'                    aParte(U).sCelular = ""
'                    aParte(U).sFoneContato = ""
'                    aParte(U).sEmail = Left(SubNull(RdoAux2!Email), 100)
'                   .MoveNext
'                Loop
'               .Close
'            End With
'        ElseIf nCodReduz >= 100000 And nCodReduz < 500000 Then
'            sql = "SELECT mobiliarioproprietario.codcidadao, cidadao.nomecidadao, cidadao.cnpj,cidadao.cpf,cidadao.rg,cidadao.telefone, cidadao.email FROM mobiliarioproprietario INNER JOIN "
'            sql = sql & "cidadao ON mobiliarioproprietario.codcidadao = cidadao.codcidadao Where codmobiliario=" & nCodReduz
'            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
'            With RdoAux2
'                Do Until .EOF
'                    ReDim Preserve aParte(UBound(aParte) + 1)
'                    U = UBound(aParte)
'                    xImovel.RetornaEndereco nCodCidadao, cidadao, cadastrocidadao
'                    aParte(U).sNomeSocio = RdoAux2!nomecidadao
'                    If Trim(SubNull(RdoAux2!Cnpj)) <> "" Then
'                        aParte(U).sCPFCNPJSocio = RetornaNumero(RdoAux2!Cnpj)
'                    Else
'                        If Trim(SubNull(RdoAux2!cpf)) <> "" Then
'                            aParte(U).sCPFCNPJSocio = RetornaNumero(RdoAux2!cpf)
'                        Else
'                            aParte(U).sCPFCNPJSocio = ""
'                        End If
'                    End If
'                    aParte(U).sRGIE = RetornaNumero(Trim(Left(SubNull(RdoAux2!rg), 14)))
'                    aParte(U).sCodLogradouro = xImovel.CodLogradouro
'                    aParte(U).sEndereco = xImovel.Endereco
'                    aParte(U).sNumero = xImovel.Numero
'                    aParte(U).sComplemento = xImovel.Complemento
'                    aParte(U).sCodBairro = xImovel.CodBairro
'                    aParte(U).sBairro = xImovel.Bairro
'                    aParte(U).sCodCidade = xImovel.CodCidade
'                    aParte(U).sCidade = xImovel.Cidade
'                    aParte(U).sCepSocio = RetornaNumero(xImovel.Cep)
'                    aParte(U).sEstadoSocio = xImovel.UF
'                    aParte(U).sClassificao = "Sócio"
'                    aParte(U).sFoneRes = Left(RetornaNumero(SubNull(RdoAux2!telefone)), 15)
'                    aParte(U).sFoneCom = ""
'                    aParte(U).sCelular = ""
'                    aParte(U).sFoneContato = ""
'                    aParte(U).sEmail = Left(SubNull(RdoAux2!Email), 100)
'                   .MoveNext
'                Loop
'                RdoAux2.Close
'            End With
'        End If
'
'        For x = 1 To UBound(aCda)
'            For U = 1 To UBound(aParte)
'                With aParte(U)
'                    sql = "INSERT Partes_protesto(idCDA,Tipo,Crc,Nome,CpfCnpj,RgInscrEstadual,Cep,Endereco,Numero,Complemento,Bairro,Cidade,Estado,DtGeracao) Values("
'                    sql = sql & aCda(x).nCDA & ",'" & .sClassificao & "'," & nCodReduz & ",'" & Mask(.sNomeSocio) & "','" & .sCPFCNPJSocio & "','" & .sRGIE & "','" & .sCepSocio & "','"
'                    sql = sql & Mask(.sEndereco) & "'," & Val(.sNumero) & ",'" & Mask(.sComplemento) & "','" & Mask(.sBairro) & "','"
'                    sql = sql & Mask(.sCidade) & "','" & SubNull(.sEstadoSocio) & "','" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "')"
'                    cnInt.Execute sql, rdExecDirect
'                End With
'            Next
'        Next
'
'Proximo:
'        nPos = nPos + 1
'        DoEvents
'       .MoveNext
'    Loop
'Fim:
'   .Close
'End With
'cmdExec.Enabled = True
'cnInt.Close
'Liberado
'MsgBox "Exportação finalizada.", vbInformation
'
End Sub

Private Sub Relatorio2018()
Dim sql As String, RdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, nCodReduz As Long, RdoAux3 As rdoResultset
Dim t As Integer, s As Integer, U As Integer, v As Integer, cmdComm As ADODB.Command, clsImovel As New clsImovel
Dim RdoDebito As rdoResultset, qd As New rdoQuery, Achou As Boolean, x As Integer, RdoTrib As rdoResultset
Dim Rs As ADODB.Recordset, strQuery As String, nNumExec As Long, nAnoExec As Integer, sNumExec As String

Dim sNomeCidadao As String, sInscricao As String, sFoneRes As String, sFoneCom As String, sCelular As String, sFoneContato As String, sEmail As String, nNumeroLocal As Integer
Dim sCPFCNPJ As String, nCodLogradouroLocal As Integer, sEnderecoLocal As String, sComplementoLocal As String, sNumeroLocal As String, nCodBairroLocal As Integer, sBairroLocal As String, sCidadeLocal As String
Dim sCEPLocal As String, sQuadra As String, sLote As String, sAtividade As String, sMatricula As String, sRGIE As String, nCodLogradouro As Integer, sUFLocal As String
Dim sEndereco As String, nNumero As Integer, sComplemento As String, nCodBairro As Integer, nCodCidade As Integer, sBairro As String, sCidade As String, sCep As String, sUF As String
Dim SetorDevedor As String, DtInscricao As String, NroCertidao As Integer, NroLivro As Integer, NroFolha As Integer, NroOrdem As Integer
Dim nCDA As Long, nAno As Integer, nLanc As Integer, nSeqLanc As Integer, nParc As Integer, nCompl As Integer, nCodTrib As Integer, sDescTrib As String
Dim dtVencto As String, nValorPrincipal As Double, nCodCidadao As Long, sClassificacao As String
Dim aCda() As typeCDA, aCdaDebito() As typeCDADebito, aParte() As Reg01, nValorTotal As Double


Ocupado

cmdExec.Enabled = False
Set qd.ActiveConnection = cn
qd.QueryTimeout = 0

If cmbCadastro.ListIndex = 0 Then
    MsgBox "selecione um cadastro"
    Exit Sub
End If
sql = "SELECT  DISTINCT CODREDUZIDO FROM DEBITOPARCELA WHERE DATAVENCIMENTO<='" & Format(mskDataVencto.Text, "mm/dd/yyyy") & "' AND STATUSLANC in (3,38,39,42,43) AND NUMPARCELA>0 AND DATAAJUIZA IS NULL AND DATAINSCRICAO IS NOT NULL "
If cmbCadastro.ListIndex = 1 Then
    sql = sql & " AND CODREDUZIDO BETWEEN 9976 AND 40000 "
    '    Sql = Sql & " AND CODREDUZIDO = 10508 "
    'Sql = Sql & " AND CODREDUZIDO < 40000 "
    sql = sql & " AND CODLANCAMENTO<>20 "
ElseIf cmbCadastro.ListIndex = 2 Then
    sql = sql & " AND CODREDUZIDO BETWEEN 100000 AND 300000 "
    'Sql = Sql & " AND CODREDUZIDO = 116679 "
    sql = sql & " AND CODLANCAMENTO<>20 "
ElseIf cmbCadastro.ListIndex = 3 Then
    'Sql = Sql & " AND CODREDUZIDO =559562 "
    sql = sql & " AND CODREDUZIDO BETWEEN 500000 AND 699999 "
    'Sql = Sql & " AND CODLANCAMENTO in(50,65,49,16,62,27,71,48,74) "
   sql = sql & " AND CODLANCAMENTO <>20 "
End If
sql = sql & " ORDER BY CODREDUZIDO"



'***********

Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nPos = 1
    nTot = .RowCount
    Do Until .EOF
        nCodReduz = !CODREDUZIDO
        If nCodReduz < 100000 Then
            SetorDevedor = "Imobiliário"
        ElseIf nCodReduz >= 100000 And nCodReduz < 500000 Then
            SetorDevedor = "Mobiliário"
        Else
            SetorDevedor = "Cidadão"
        End If
            
        If nPos Mod 10 = 0 Then
          '  GoTo fim
            DoEvents
            CallPb nPos, nTot
        End If
        
        '*** debito ***
        ReDim aParte(0)
        ReDim aCda(0)
        ReDim aCdaDebito(0)
        
        On Error Resume Next
        RdoDebito.Close
        On Error GoTo 0
        If nCadastro = 3 Then
            qd.sql = "{ Call spEXTRATOAJUIZARTAXA(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
        Else
            qd.sql = "{ Call spEXTRATOAJUIZAR(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
        End If
        qd(0) = nCodReduz
        qd(1) = nCodReduz
        qd(2) = 1950
        qd(3) = 2018
        qd(4) = 1
        qd(5) = 999
        qd(6) = 0
        qd(7) = 999
        qd(8) = 1
        qd(9) = 999
        qd(10) = 0
        qd(11) = 99
        qd(12) = 3
        qd(13) = 3
        qd(14) = Format("22/12/2025", "mm/dd/yyyy")
        qd(15) = "Integrativa"
        qd(16) = 0
        Set RdoDebito = qd.OpenResultset(rdOpenKeyset)
        With RdoDebito
            Do Until .EOF
'                If Year(!DataVencimento) < 2015 Then
'                    GoTo proximo
'                End If
              '  If !ValorJuros = 0 Then
                    'MsgBox "teste"
               ' End If
                U = UBound(aCda)
                Achou = False
                For x = 1 To U
                    If Val(aCda(x).nAno) = !AnoExercicio And Val(aCda(x).nLanc) = !CodLancamento And Val(aCda(x).nSeq) = !SeqLancamento And _
                       Val(aCda(x).nParc) = !NumParcela And Val(aCda(x).nCompl) = !CODCOMPLEMENTO Then
                        Achou = True
                        Exit For
                    End If
                Next
                
                If Not Achou Then
                    ReDim Preserve aCda(UBound(aCda) + 1)
                    U = UBound(aCda)
                    aCda(U).nAno = !AnoExercicio
                    aCda(U).nLanc = !CodLancamento
                    aCda(U).nSeq = !SeqLancamento
                    aCda(U).nParc = !NumParcela
                    aCda(U).nCompl = !CODCOMPLEMENTO
                    On Error Resume Next
                    aCda(U).nNumCertidao = Val(SubNull(!certidao))
                    On Error GoTo 0
                    aCda(U).nNumPagina = Val(SubNull(!PAGINA))
                    aCda(U).nNumLivro = Val(SubNull(!NUMLIVRO))
                    aCda(U).dDataInscricao = !datainscricao
                End If
                
                
                ReDim Preserve aCdaDebito(UBound(aCdaDebito) + 1)
                U = UBound(aCdaDebito)
                aCdaDebito(U).nAno = !AnoExercicio
                aCdaDebito(U).nLanc = !CodLancamento
                aCdaDebito(U).nSeq = !SeqLancamento
                aCdaDebito(U).nParc = !NumParcela
                aCdaDebito(U).nCompl = !CODCOMPLEMENTO
                aCdaDebito(U).nCodTributo = !CodTributo
                aCdaDebito(U).sDescTributo = !abrevTributo
                aCdaDebito(U).nPrincipal = !VALORTRIBUTO
                aCdaDebito(U).nMulta = !ValorMulta
                aCdaDebito(U).nJuros = !ValorJuros
                aCdaDebito(U).nCorrecao = !valorcorrecao
                aCdaDebito(U).nMulta = !ValorMulta
                aCdaDebito(U).nJuros = !ValorJuros
                aCdaDebito(U).nCorrecao = !valorcorrecao
                aCdaDebito(U).nTotal = !ValorTotal
                aCdaDebito(U).dDataVencto = !DataVencimento
PROXIMODEBITO:
                DoEvents
               .MoveNext
            Loop
           .Close
        End With

        If UBound(aCda) = 0 Then GoTo Proximo
        'limita em 1301 reais
        nValorTotal = 0
        For x = 1 To UBound(aCdaDebito)
            With aCdaDebito(x)
                nValorTotal = nValorTotal + .nTotal
            End With
        Next
        If nValorTotal > 200 Then
            GoTo Proximo
        End If
        
        sql = "insert rel2018(codigo,valor) values(" & nCodReduz & "," & Virg2Ponto(CStr(nValorTotal)) & ")"
        cn.Execute sql, rdExecDirect
        
        
'        On Error Resume Next
'        For x = 1 To UBound(aCda)
'            With aCda(x)
'                Sql = "insert rel2018(codigo,valor) values(" & nCodReduz & "," & Virg2Ponto(CStr(nValorTotal)) & ")"
'                cn.Execute Sql, rdExecDirect
 '           End With
  '      Next
        
Proximo:
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
Fim:
   .Close
End With
cmdExec.Enabled = True
'cnInt.Close
Liberado
MsgBox "Exportação finalizada.", vbInformation

End Sub

Private Sub Protestar_New()

Dim sql As String, RdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, nCodReduz As Long, RdoAux3 As rdoResultset
Dim t As Integer, s As Integer, U As Integer, v As Integer, cmdComm As ADODB.Command, clsImovel As New clsImovel
Dim RdoDebito As rdoResultset, qd As New rdoQuery, Achou As Boolean, x As Integer, RdoTrib As rdoResultset
Dim Rs As ADODB.Recordset, strQuery As String, nNumExec As Long, nAnoExec As Integer, sNumExec As String

Dim sNomeCidadao As String, sInscricao As String, sFoneRes As String, sFoneCom As String, sCelular As String, sFoneContato As String, sEmail As String, nNumeroLocal As Integer
Dim sCPFCNPJ As String, nCodLogradouroLocal As Integer, sEnderecoLocal As String, sComplementoLocal As String, sNumeroLocal As String, nCodBairroLocal As Integer, sBairroLocal As String, sCidadeLocal As String
Dim sCEPLocal As String, sQuadra As String, sLote As String, sAtividade As String, sMatricula As String, sRGIE As String, nCodLogradouro As Integer, sUFLocal As String
Dim sEndereco As String, nNumero As Integer, sComplemento As String, nCodBairro As Integer, nCodCidade As Integer, sBairro As String, sCidade As String, sCep As String, sUF As String
Dim SetorDevedor As String, DtInscricao As String, NroCertidao As Integer, NroLivro As Integer, NroFolha As Integer, NroOrdem As Integer
Dim nCDA As Long, nAno As Integer, nLanc As Integer, nSeqLanc As Integer, nParc As Integer, nCompl As Integer, nCodTrib As Integer, sDescTrib As String
Dim dtVencto As String, nValorPrincipal As Double, nCodCidadao As Long, sClassificacao As String
Dim aCda() As typeCDANew, aCdaDebito() As typeCDADebito, aParte() As Reg01
Dim nPrincipal As Double, nMulta As Double, nJuros As Double, nCorrecao As Double, nTotal As Double, nValorTotal As Double


Ocupado
ConectaIntegrativa
sql = "delete from cdadebitos_protesto"
'cnInt.Execute Sql, rdExecDirect


sql = "delete from cadastro_protesto"
'cnInt.Execute Sql, rdExecDirect

sql = "delete from partes_protesto"
'cnInt.Execute Sql, rdExecDirect

sql = "delete from cdas_protesto"
'cnInt.Execute Sql, rdExecDirect

cmdExec.Enabled = False
Set qd.ActiveConnection = cn
qd.QueryTimeout = 0

If cmbCadastro.ListIndex = 0 Then
    MsgBox "selecione um cadastro"
    Exit Sub
End If
'Sql = "SELECT  DISTINCT CODREDUZIDO FROM DEBITOPARCELA WHERE DATAVENCIMENTO<='" & Format(mskDataVencto.Text, "mm/dd/yyyy") & "' AND STATUSLANC =3 AND NUMPARCELA>0 AND DATAAJUIZA IS NULL AND DATAINSCRICAO IS NOT NULL AND PROTESTO_DATA_REMESSA IS NULL"
sql = "SELECT  DISTINCT CODREDUZIDO FROM DEBITOPARCELA WHERE (anoexercicio between 2020 and 2022) and DATAVENCIMENTO<='" & Format(mskDataVencto.Text, "mm/dd/yyyy") & "' AND (STATUSLANC =3 or STATUSLANC =42 or STATUSLANC =43 or STATUSLANC =38 or STATUSLANC =39  ) AND NUMPARCELA>0 AND DATAAJUIZA IS NULL AND DATAINSCRICAO IS NOT NULL AND PROTESTO_DATA_REMESSA IS NULL"
If cmbCadastro.ListIndex = 1 Then
    sql = sql & " AND CODREDUZIDO <50000 "
    'sql = sql & " AND CODREDUZIDO IN (4556,39466,37032,37050,25814,27859,24468,39652,39573,39598,39553,38888,31219,39661,39774,19530,16934,23063,39283,29649,34030,31919,39495,18981,33236,33192,33243,33053,32151,27642,28115,29893,29145,23038,17560,12713,32082,29082,23038,39506,39507,37064,1835,23246,39502) "
    'sql = sql & " AND CODLANCAMENTO <>5 AND CODLANCAMENTO<>20 AND CODLANCAMENTO<>11 "
    sql = sql & " AND CODLANCAMENTO =1 "
ElseIf cmbCadastro.ListIndex = 2 Then
    sql = sql & " AND CODREDUZIDO BETWEEN 100000 AND 300000 "
 '   Sql = Sql & " AND CODREDUZIDO in (123098,123126,123130,123136,123095 ) "
    sql = sql & " AND CODLANCAMENTO not in (11,20)  "
ElseIf cmbCadastro.ListIndex = 3 Then
    sql = sql & " AND CODREDUZIDO BETWEEN 500001 AND 699999 "
    'sql = sql & " AND CODREDUZIDO in (547663,614016) "
    sql = sql & " AND CODLANCAMENTO not in(20) "
   ' Sql = Sql & " AND CODLANCAMENTO in(74) "
End If
'sql = sql & " AND CODREDUZIDO in (100408,10175,102588,110006,112360,112620,112942,114743,116240,116520,117792,118550,119049,120752,121182,121748,122521,122867,122891,122948,123091,123232,123234,123502,12642,12664,12710,12765,16593,16735,1742,19691,19733,20165,23770,24921,25446,25786,26191,26421,26594,26652,27114,27909,30854,30919,31219,32692,33381,35342,35412,35715,35835,36426,36589,36596,36742,4593,519644,533869,541819,543033,552322,559350,560383,560783,562006,574112,613595,613702,614049,615786,618597,620998,622695,625389,626025,626744,627507,627717,629047,629150,629372,629659,629743,8951) "
sql = sql & " ORDER BY CODREDUZIDO"



'***********

Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nPos = 1
    nTot = .RowCount
    Do Until .EOF
        nCodReduz = !CODREDUZIDO
        If nCodReduz < 100000 Then
            SetorDevedor = "Imobiliário"
        ElseIf nCodReduz >= 100000 And nCodReduz < 500000 Then
            SetorDevedor = "Mobiliário"
        Else
            SetorDevedor = "Cidadão"
        End If
            
        If nPos Mod 10 = 0 Then
          '  GoTo fim
            DoEvents
            CallPb nPos, nTot
        End If
        
        '*** debito ***
        ReDim aParte(0)
        ReDim aCda(0)
        ReDim aCdaDebito(0)
        
        On Error Resume Next
        RdoDebito.Close
        On Error GoTo 0
        If cmbCadastro.ListIndex = 3 Then
            qd.sql = "{ Call spEXTRATOAJUIZARTAXA(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
        Else
            qd.sql = "{ Call spEXTRATOAJUIZAR(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
        End If
        qd(0) = nCodReduz
        qd(1) = nCodReduz
        qd(2) = 2020
        qd(3) = 2022
        qd(4) = 0
        qd(5) = 999
        qd(6) = 0
        qd(7) = 999
        qd(8) = 1
        qd(9) = 999
        qd(10) = 0
        qd(11) = 99
        qd(12) = 3
        qd(13) = 3
        qd(14) = Format("20/12/2025", "mm/dd/yyyy")
        qd(15) = "Protesto"
        qd(16) = 0
        Set RdoDebito = qd.OpenResultset(rdOpenKeyset)
        With RdoDebito
            Do Until .EOF
'                If Year(!DataVencimento) < 2015 Then
'                    GoTo proximo
'                End If
              '  If !ValorJuros = 0 Then
                    'MsgBox "teste"
               ' End If
                U = UBound(aCda)
                Achou = False
                For x = 1 To U
'                    If Val(aCda(x).nAno) = !AnoExercicio And Val(aCda(x).nLanc) = !CodLancamento And Val(aCda(x).nSeq) = !SeqLancamento And _
'                       Val(aCda(x).nParc) = !NumParcela And Val(aCda(x).nCompl) = !CODCOMPLEMENTO And Val(aCda(x).nNumCertidao) = Val(SubNull(!certidao)) Then
                    If Val(aCda(x).nNumCertidao) = Val(SubNull(!certidao)) Then
                        Achou = True
                        Exit For
                    End If
                Next
                
                If Not Achou Then
                    ReDim Preserve aCda(UBound(aCda) + 1)
                    U = UBound(aCda)
                    aCda(U).nAno = !AnoExercicio
                    aCda(U).nLanc = !CodLancamento
                    aCda(U).nSeq = !SeqLancamento
                    aCda(U).nParc = !NumParcela
                    aCda(U).nCompl = !CODCOMPLEMENTO
                    On Error Resume Next
                    aCda(U).nNumCertidao = Val(SubNull(!certidao))
                    On Error GoTo 0
                    aCda(U).nNumPagina = Val(SubNull(!PAGINA))
                    aCda(U).nNumLivro = Val(SubNull(!NUMLIVRO))
                    aCda(U).dDataInscricao = !datainscricao
                End If
                
                
                ReDim Preserve aCdaDebito(UBound(aCdaDebito) + 1)
                U = UBound(aCdaDebito)
                aCdaDebito(U).nAno = !AnoExercicio
                aCdaDebito(U).nLanc = !CodLancamento
                aCdaDebito(U).nSeq = !SeqLancamento
                aCdaDebito(U).nParc = !NumParcela
                aCdaDebito(U).nCompl = !CODCOMPLEMENTO
                aCdaDebito(U).nCodTributo = !CodTributo
                aCdaDebito(U).sDescTributo = !abrevTributo
                aCdaDebito(U).nPrincipal = !VALORTRIBUTO
                aCdaDebito(U).nMulta = !ValorMulta
                aCdaDebito(U).nJuros = !ValorJuros
                aCdaDebito(U).nCorrecao = !valorcorrecao
                aCdaDebito(U).nMulta = !ValorMulta
                aCdaDebito(U).nJuros = !ValorJuros
                aCdaDebito(U).nCorrecao = !valorcorrecao
                aCdaDebito(U).nTotal = !ValorTotal
                aCdaDebito(U).nNumCertidao = Val(SubNull(!certidao))
                aCdaDebito(U).nFolha = Val(SubNull(!PAGINA))
                aCdaDebito(U).nLivro = Val(SubNull(!NUMLIVRO))
                aCdaDebito(U).dDataVencto = !DataVencimento
PROXIMODEBITO:
                DoEvents
               .MoveNext
            Loop
           .Close
        End With

        If UBound(aCda) = 0 Then GoTo Proximo
        'limita em 1301 reais
        nValorTotal = 0
        For x = 1 To UBound(aCdaDebito)
            With aCdaDebito(x)
                nValorTotal = nValorTotal + .nTotal
            End With
        Next
 '       If nValorTotal < 1301 Then
'            GoTo Proximo
 '       End If
        
        
        For x = 1 To UBound(aCda)
        
            With aCda(x)
                sql = "insert CDAs_protesto (idDevedor, SetorDevedor, DtInscricao, NroCertidao,nrolivro, vlrOriginal, vlrjuros,vlrmultas,vlrcorrecao,vlrtotal,  DtGeracao) values("
                sql = sql & nCodReduz & ",'" & SetorDevedor & "','" & Format(.dDataInscricao, "mm/dd/yyyy") & "'," & .nNumCertidao & "," & .nNumLivro & "," & Virg2Ponto(CStr(.nPrincipal)) & ","
                sql = sql & Virg2Ponto(CStr(.nJuros)) & "," & Virg2Ponto(CStr(.nMulta)) & "," & Virg2Ponto(CStr(.nCorrecao)) & "," & Virg2Ponto(CStr(.nTotal)) & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "')"
                cnInt.Execute sql, rdExecDirect
            
                sql = "select @@identity as LastKey"
                Set RdoAux3 = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                nCDA = RdoAux3!lastkey
               .nCDA = nCDA
                RdoAux3.Close

                nPrincipal = 0: nJuros = 0: nMulta = 0: nCorrecao = 0: nTotal = 0
                For v = 1 To UBound(aCdaDebito)
'                    If aCdaDebito(v).nAno = .nAno And aCdaDebito(v).nLanc = .nLanc And aCdaDebito(v).nSeq = .nSeq And _
'                       aCdaDebito(v).nParc = .nParc And aCdaDebito(v).nCompl = .nCompl Then
                    If aCdaDebito(v).nNumCertidao = .nNumCertidao Then
                        aCdaDebito(v).nCDA = .nCDA
                        For z = 1 To UBound(aCda)
                            If aCda(z).nCDA = .nCDA Then
                                nPrincipal = nPrincipal + aCdaDebito(v).nPrincipal
                                nJuros = nJuros + aCdaDebito(v).nJuros
                                nMulta = nMulta + aCdaDebito(v).nMulta
                                nCorrecao = nCorrecao + aCdaDebito(v).nCorrecao
                            End If
                        Next
                        nTotal = nPrincipal + nJuros + nMulta + nCorrecao
                    End If
                Next
               
                sql = "update CDAs_protesto set vlroriginal=" & Virg2Ponto(CStr(nPrincipal)) & ",vlrmultas = " & Virg2Ponto(CStr(nMulta))
                sql = sql & ",vlrjuros=" & Virg2Ponto(CStr(nJuros)) & ",vlrcorrecao = " & Virg2Ponto(CStr(nCorrecao)) & ",vlrtotal = " & Virg2Ponto(CStr(nTotal))
                sql = sql & " where idcda=" & .nCDA
                cnInt.Execute sql, rdExecDirect
            End With
        Next
        
        For x = 1 To UBound(aCdaDebito)
            With aCdaDebito(x)
                sql = "insert CDADebitos_protesto(idCDA, CodTributo, Tributo, Exercicio, Lancamento, Seq, NroParcela, ComplParcela, DtVencimento, vlrOriginal,vlrMultas,vlrjuros,vlrcorrecao,nroLivro,nroFolha, DtGeracao) values("
                sql = sql & .nCDA & "," & .nCodTributo & ",'" & Mask(.sDescTributo) & "'," & .nAno & "," & .nLanc & "," & .nSeq & "," & .nParc & "," & .nCompl & ",'"
                sql = sql & Format(.dDataVencto, "mm/dd/yyyy") & "'," & Virg2Ponto(Format(.nPrincipal, "#0.00")) & "," & Virg2Ponto(Format(.nMulta, "#0.00")) & "," & Virg2Ponto(Format(.nJuros, "#0.00")) & ","
                sql = sql & Virg2Ponto(Format(.nCorrecao, "#0.00")) & "," & .nLivro & "," & .nFolha & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "')"
                cnInt.Execute sql, rdExecDirect
            End With
        Next

       '*** dados cadastrais
        If nCodReduz < 100000 Then
            sql = "select * from vwfullimovel2 where codreduzido=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If .RowCount = 0 Then GoTo Proximo
                sNome = !nomecidadao
                sInscricao = !Inscricao
                sFoneRes = Left(RetornaNumero(SubNull(!telefone)), 15)
                sFoneCom = ""
                sCelular = ""
                sFoneContato = ""
                sEmail = Left(SubNull(!Email), 100)
                If Trim(SubNull(!Cnpj)) <> "" Then
                    sCPFCNPJ = RetornaNumero(!Cnpj)
                Else
                    If Trim(SubNull(!cpf)) <> "" Then
                        sCPFCNPJ = RetornaNumero(!cpf)
                    Else
                        sCPFCNPJ = ""
                    End If
                End If
                
                xImovel.RetornaEndereco nCodReduz, Imobiliario, Localizacao
                nCodLogradouroLocal = xImovel.CodLogradouro
                
                sEnderecoLocal = xImovel.Endereco
                nNumeroLocal = xImovel.Numero
                sComplementoLocal = xImovel.Complemento
                nCodBairroLocal = xImovel.CodBairro
                sBairroLocal = xImovel.Bairro
                sCidadeLocal = xImovel.Cidade
                sUFLocal = xImovel.UF
                sCEPLocal = xImovel.Cep
                If Val(nCodBairroLocal) = 999 Then
                    nCodBairroLocal = 0
                    sBairroLocal = ""
                End If
                
                sQuadra = !Quadra
                sLote = !Lote
                sAtividade = ""
                sMatricula = Val(SubNull(!NumMat))
                
                sRGIE = RetornaNumero(Trim(SubNull(!rg)))
                If !Ee_TipoEnd = 0 Then 'imovel
                    xImovel.RetornaEndereco nCodReduz, Imobiliario, Localizacao
                    nCodLogradouro = xImovel.CodLogradouro
                    sEndereco = xImovel.Endereco
                    nNumero = xImovel.Numero
                    sComplemento = xImovel.Complemento
                    nCodBairro = xImovel.CodBairro
                    sBairro = xImovel.Bairro
                    nCodCidade = xImovel.CodCidade
                    sCidade = xImovel.Cidade
                    sCep = xImovel.Cep
                    sUF = xImovel.UF
                ElseIf !Ee_TipoEnd = 1 Then 'proprietario
                    xImovel.RetornaEndereco nCodReduz, Imobiliario, cadastrocidadao
                    nCodLogradouro = Val(xImovel.CodLogradouro)
                    sEndereco = xImovel.Endereco
                    nNumero = xImovel.Numero
                    sComplemento = xImovel.Complemento
                    nCodBairro = Val(xImovel.CodBairro)
                    sBairro = xImovel.Bairro
                    nCodCidade = Val(xImovel.CodCidade)
                    sCidade = xImovel.Cidade
                    sCep = xImovel.Cep
                    sUF = xImovel.UF
                Else 'entrega
                    xImovel.RetornaEndereco nCodReduz, Imobiliario, Entrega
                    nCodLogradouro = xImovel.CodLogradouro
                    sEndereco = xImovel.Endereco
                    nNumero = xImovel.Numero
                    sComplemento = xImovel.Complemento
                    nCodBairro = Val(xImovel.CodBairro)
                    sBairro = xImovel.Bairro
                    nCodCidade = xImovel.CodCidade
                    sCidade = xImovel.Cidade
                    sCep = xImovel.Cep
                    sUF = xImovel.UF
                End If
               .Close
            End With
        ElseIf nCodReduz >= 100000 And nCodReduz < 500000 Then
            sql = "select * from vwfullempresa3 where codigomob=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If .RowCount = 0 Then GoTo Proximo
                sInscricao = ""
                sNome = !RazaoSocial
                sFoneRes = ""
                sFoneCom = ""
                sCelular = ""
                sFoneContato = Left(RetornaNumero(SubNull(!fonecontato)), 15)
                sEmail = Left(SubNull(!emailcontato), 100)
                If Trim(SubNull(!Cnpj)) <> "" Then
                    sCPFCNPJ = RetornaNumero(!Cnpj)
                Else
                    If Trim(SubNull(!cpf)) <> "" Then
                        sCPFCNPJ = RetornaNumero(!cpf)
                    Else
                        sCPFCNPJ = ""
                    End If
                End If
                sRGIE = RetornaNumero(Trim(SubNull(!inscestadual)))
                If sRGIE = "" Then
                    sRGIE = Trim(SubNull(!rg) & " " & SubNull(!Orgao))
                End If
                
                xImovel.RetornaEndereco nCodReduz, Mobiliario, Localizacao
                nCodLogradouroLocal = xImovel.CodLogradouro
                sEnderecoLocal = xImovel.Endereco
                nNumeroLocal = xImovel.Numero
                sComplementoLocal = xImovel.Complemento
                nCodBairroLocal = Val(xImovel.CodBairro)
                sBairroLocal = xImovel.Bairro
                sCidadeLocal = xImovel.Cidade
                sUFLocal = xImovel.UF
                sCEPLocal = xImovel.Cep
                If Val(nCodBairroLocal) = 999 Then
                    nCodBairroLocal = 0
                    sBairroLocal = ""
                End If
                
                sEndereco = SubNull(!eenomelogr)
                If sEndereco = "" Then
                    'local da empresa
                    xImovel.RetornaEndereco nCodReduz, Mobiliario, Localizacao
                    nCodLogradouro = xImovel.CodLogradouro
                    sEndereco = xImovel.Endereco
                    nNumero = xImovel.Numero
                    sComplemento = xImovel.Complemento
                    nCodBairro = Val(xImovel.CodBairro)
                    sBairro = xImovel.Bairro
                    nCodCidade = xImovel.CodCidade
                    sCidade = xImovel.Cidade
                    sCep = xImovel.Cep
                    sUF = xImovel.UF
                Else
'                   'endereco entrega
                    xImovel.RetornaEndereco nCodReduz, Mobiliario, Entrega
                    nCodLogradouro = xImovel.CodLogradouro
                    sEndereco = xImovel.Endereco
                    nNumero = xImovel.Numero
                    sComplemento = xImovel.Complemento
                    nCodBairro = Val(xImovel.CodBairro)
                    sBairro = xImovel.Bairro
                    nCodCidade = xImovel.CodCidade
                    sCidade = xImovel.Cidade
                    sCep = xImovel.Cep
                    sUF = xImovel.UF
                End If
               
                sQuadra = ""
                sLote = ""
                sAtividade = Left(SubNull(!ativextenso), 80)
                sMatricula = ""
               
               .Close
            End With
        Else
            sql = "select * from vwfullcidadao where codcidadao=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If .RowCount = 0 Then GoTo Proximo
                sInscricao = ""
                sNome = !nomecidadao
                sFoneRes = Left(RetornaNumero(SubNull(!telefone)), 15)
                sFoneCom = ""
                sCelular = ""
                sFoneContato = ""
                sEmail = Left(SubNull(!Email), 100)
                If Trim(SubNull(!Cnpj)) <> "" Then
                    sCPFCNPJ = RetornaNumero(!Cnpj)
                Else
                    If Trim(SubNull(!cpf)) <> "" Then
                        sCPFCNPJ = RetornaNumero(!cpf)
                    Else
                        sCPFCNPJ = ""
                    End If
                End If
                sRGIE = RetornaNumero(Trim(Left(SubNull(!rg), 14)))
                xImovel.RetornaEndereco nCodReduz, cidadao, cadastrocidadao
                nCodLogradouro = Val(xImovel.CodLogradouro)
                sEndereco = SubNull(xImovel.Endereco)
                nNumero = SubNull(xImovel.Numero)
                sComplemento = xImovel.Complemento
                nCodBairro = Val(SubNull(xImovel.CodBairro))
                sBairro = SubNull(xImovel.Bairro)
                If nCodBairro = 999 Then
                    nCodBairro = 0
                    sBairro = ""
                End If
                nCodCidade = Val(SubNull(xImovel.CodCidade))
                sCidade = SubNull(xImovel.Cidade)
                sUFLocal = xImovel.UF
                sCep = RetornaNumero(SubNull(xImovel.Cep))
                sUF = SubNull(xImovel.UF)
               
                nCodLogradouroLocal = Val(xImovel.CodLogradouro)
                sEnderecoLocal = SubNull(xImovel.Endereco)
                nNumeroLocal = SubNull(xImovel.Numero)
                sComplementoLocal = SubNull(xImovel.Complemento)
                nCodBairroLocal = Val(SubNull(xImovel.CodBairro))
                sBairroLocal = SubNull(xImovel.Bairro)
                If Val(nCodBairroLocal) = 999 Then
                    nCodBairroLocal = 0
                    sBairroLocal = ""
                End If
                sCidadeLocal = xImovel.Cidade
                sCEPLocal = RetornaCEP(CLng(nCodLogradouroLocal), nNumeroLocal)
                
                sQuadra = ""
                sLote = ""
                sAtividade = ""
                sMatricula = ""
               
               .Close
            End With
        End If
                
        For x = 1 To UBound(aCda)
            With aCda(x)
                sql = "INSERT Cadastro_protesto(IdCDA,SetorDevedor,Crc,Nome,Inscricao,CPFCnpj,RgInscrEstadual,LocalCep,LocalEndereco,LocalNumero,LocalComplemento,"
                sql = sql & "LocalBairro,LocalCidade,LocalEstado,Quadra,Lote,EntregaCep,EntregaEndereco,EntregaNumero,EntregaComplemento,EntregaBairro,"
                sql = sql & "EntregaCidade,EntregaEstado,DtGeracao) values("
                sql = sql & .nCDA & ",'" & SetorDevedor & "'," & nCodReduz & ",'" & Mask(Left(SubNull(sNome), 80)) & "','" & sInscricao & "','" & sCPFCNPJ & "','" & sRGIE & "','"
                sql = sql & sCEPLocal & "','" & sEnderecoLocal & "'," & nNumeroLocal & ",'" & Mask(Left(sComplementoLocal, 50)) & "','" & Left(sBairroLocal, 40) & "','" & sCidadeLocal & "','" & sUFLocal & "','"
                sql = sql & sQuadra & "','" & sLote & "','" & sCep & "','" & sEndereco & "'," & nNumero & ",'" & Mask(Left(sComplemento, 50)) & "','"
                sql = sql & Left(sBairro, 40) & "','" & Mask(sCidade) & "','" & sUF & "','" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "')"
                cnInt.Execute sql, rdExecDirect
            End With
        Next
        
        
        '***** PARTES **********************************

        If nCodReduz < 100000 Then
            sql = "SELECT proprietario.codcidadao, cidadao.nomecidadao, cidadao.cnpj,cidadao.cpf,cidadao.rg,cidadao.telefone, cidadao.email FROM proprietario INNER JOIN "
            sql = sql & "cidadao ON proprietario.codcidadao = cidadao.codcidadao Where codreduzido=" & nCodReduz & " and principal=0"
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                Do Until .EOF
                    ReDim Preserve aParte(UBound(aParte) + 1)
                    U = UBound(aParte)
                    xImovel.RetornaEndereco !CodCidadao, cidadao, cadastrocidadao
                    aParte(U).sNomeSocio = RdoAux2!nomecidadao
                    If Trim(SubNull(RdoAux2!Cnpj)) <> "" Then
                        aParte(U).sCPFCNPJSocio = RetornaNumero(RdoAux2!Cnpj)
                    Else
                        If Trim(SubNull(RdoAux2!cpf)) <> "" Then
                            aParte(U).sCPFCNPJSocio = RetornaNumero(RdoAux2!cpf)
                        Else
                            aParte(U).sCPFCNPJSocio = ""
                        End If
                    End If
                    aParte(U).sRGIE = RetornaNumero(Trim(Left(SubNull(RdoAux2!rg), 14)))
                    aParte(U).sCodLogradouro = xImovel.CodLogradouro
                    aParte(U).sEndereco = xImovel.Endereco
                    aParte(U).sNumero = xImovel.Numero
                    aParte(U).sComplemento = xImovel.Complemento
                    aParte(U).sCodBairro = xImovel.CodBairro
                    aParte(U).sBairro = xImovel.Bairro
                    aParte(U).sCodCidade = xImovel.CodCidade
                    aParte(U).sCidade = xImovel.Cidade
                    aParte(U).sCepSocio = RetornaNumero(xImovel.Cep)
                    aParte(U).sEstadoSocio = xImovel.UF
                    aParte(U).sClassificao = "Compromissário"
                    aParte(U).sFoneRes = Left(RetornaNumero(SubNull(RdoAux2!telefone)), 15)
                    aParte(U).sFoneCom = ""
                    aParte(U).sCelular = ""
                    aParte(U).sFoneContato = ""
                    aParte(U).sEmail = Left(SubNull(RdoAux2!Email), 100)
                   .MoveNext
                Loop
               .Close
            End With
        ElseIf nCodReduz >= 100000 And nCodReduz < 500000 Then
            sql = "SELECT mobiliarioproprietario.codcidadao, cidadao.nomecidadao, cidadao.cnpj,cidadao.cpf,cidadao.rg,cidadao.telefone, cidadao.email FROM mobiliarioproprietario INNER JOIN "
            sql = sql & "cidadao ON mobiliarioproprietario.codcidadao = cidadao.codcidadao Where codmobiliario=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                Do Until .EOF
                    ReDim Preserve aParte(UBound(aParte) + 1)
                    U = UBound(aParte)
                    xImovel.RetornaEndereco !CodCidadao, cidadao, cadastrocidadao
                    aParte(U).sNomeSocio = RdoAux2!nomecidadao
                    If Trim(SubNull(RdoAux2!Cnpj)) <> "" Then
                        aParte(U).sCPFCNPJSocio = RetornaNumero(RdoAux2!Cnpj)
                    Else
                        If Trim(SubNull(RdoAux2!cpf)) <> "" Then
                            aParte(U).sCPFCNPJSocio = RetornaNumero(RdoAux2!cpf)
                        Else
                            aParte(U).sCPFCNPJSocio = ""
                        End If
                    End If
                    aParte(U).sRGIE = RetornaNumero(Trim(Left(SubNull(RdoAux2!rg), 14)))
                    aParte(U).sCodLogradouro = xImovel.CodLogradouro
                    aParte(U).sEndereco = xImovel.Endereco
                    aParte(U).sNumero = xImovel.Numero
                    aParte(U).sComplemento = xImovel.Complemento
                    aParte(U).sCodBairro = xImovel.CodBairro
                    aParte(U).sBairro = xImovel.Bairro
                    aParte(U).sCodCidade = xImovel.CodCidade
                    aParte(U).sCidade = xImovel.Cidade
                    aParte(U).sCepSocio = RetornaNumero(xImovel.Cep)
                    aParte(U).sEstadoSocio = xImovel.UF
                    aParte(U).sClassificao = "Sócio"
                    aParte(U).sFoneRes = Left(RetornaNumero(SubNull(RdoAux2!telefone)), 15)
                    aParte(U).sFoneCom = ""
                    aParte(U).sCelular = ""
                    aParte(U).sFoneContato = ""
                    aParte(U).sEmail = Left(SubNull(RdoAux2!Email), 100)
                   .MoveNext
                Loop
                RdoAux2.Close
            End With
        End If
        
        For x = 1 To UBound(aCda)
            For U = 1 To UBound(aParte)
                With aParte(U)
                    sql = "INSERT Partes_protesto(idCDA,Tipo,Crc,Nome,CpfCnpj,RgInscrEstadual,Cep,Endereco,Numero,Complemento,Bairro,Cidade,Estado,DtGeracao) Values("
                    sql = sql & aCda(x).nCDA & ",'" & .sClassificao & "'," & nCodReduz & ",'" & Mask(.sNomeSocio) & "','" & .sCPFCNPJSocio & "','" & .sRGIE & "','" & .sCepSocio & "','"
                    sql = sql & Mask(.sEndereco) & "'," & Val(.sNumero) & ",'" & Mask(.sComplemento) & "','" & Mask(.sBairro) & "','"
                    sql = sql & Mask(.sCidade) & "','" & SubNull(.sEstadoSocio) & "','" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "')"
                    cnInt.Execute sql, rdExecDirect
                End With
            Next
        Next
        
Proximo:
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
Fim:
   .Close
End With
cmdExec.Enabled = True
cnInt.Close
Liberado
MsgBox "Exportação finalizada.", vbInformation


End Sub
