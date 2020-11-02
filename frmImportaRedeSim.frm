VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportaRedeSim 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importação de dados - RedeSim"
   ClientHeight    =   4875
   ClientLeft      =   4110
   ClientTop       =   4335
   ClientWidth     =   14190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4875
   ScaleWidth      =   14190
   Begin Tributacao.jcFrames frProgress 
      Height          =   1155
      Left            =   4455
      Top             =   1755
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2037
      FrameColor      =   255
      FillColor       =   4210688
      TextBoxColor    =   8454016
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
      ThemeColor      =   3
      ColorFrom       =   192
      ColorTo         =   8438015
      Begin Tributacao.XP_ProgressBar pBar 
         Height          =   165
         Left            =   150
         TabIndex        =   8
         Top             =   780
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   291
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
         Color           =   255
         Scrolling       =   1
      End
      Begin VB.Label lblFileNumber 
         BackStyle       =   0  'Transparent
         Caption         =   "Importando Empresa 0 de 0"
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   90
         Width           =   4305
      End
      Begin VB.Label lblProtocolo 
         BackStyle       =   0  'Transparent
         Caption         =   "ARRECADA08.ret"
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   360
         Width           =   4155
      End
   End
   Begin prjChameleon.chameleonButton cmdGravar 
      Height          =   345
      Left            =   5310
      TabIndex        =   2
      ToolTipText     =   "Importar os dados para o GTI"
      Top             =   4410
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Importar"
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
      MICON           =   "frmImportaRedeSim.frx":0000
      PICN            =   "frmImportaRedeSim.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7830
      Top             =   8100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   3885
      Left            =   90
      TabIndex        =   1
      Top             =   360
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   6853
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Arquivo"
         Object.Width           =   7409
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tipo Arquivo"
         Object.Width           =   2118
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Qtde"
         Object.Width           =   1058
      EndProperty
   End
   Begin prjChameleon.chameleonButton cmdArq 
      Height          =   345
      Left            =   4005
      TabIndex        =   0
      ToolTipText     =   "Selecionar os arquivos a serem importados"
      Top             =   4410
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Arquivos"
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
      MICON           =   "frmImportaRedeSim.frx":03C1
      PICN            =   "frmImportaRedeSim.frx":03DD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView lvVia 
      Height          =   2040
      Left            =   1575
      TabIndex        =   3
      Top             =   6300
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   3598
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Protocolo"
         Object.Width           =   2823
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Razão Social"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Situação"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvLic 
      Height          =   3885
      Left            =   6660
      TabIndex        =   7
      Top             =   360
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   6853
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Protocolo"
         Object.Width           =   2823
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Razão Social"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Situação"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Licenciamento"
      Height          =   2085
      Index           =   2
      Left            =   6705
      TabIndex        =   6
      Top             =   90
      Width           =   1860
   End
   Begin VB.Label Label1 
      Caption         =   "Viablidade"
      Height          =   240
      Index           =   1
      Left            =   1620
      TabIndex        =   5
      Top             =   6030
      Width           =   1860
   End
   Begin VB.Label Label1 
      Caption         =   "Arquivos:"
      Height          =   240
      Index           =   0
      Left            =   180
      TabIndex        =   4
      Top             =   90
      Width           =   1860
   End
End
Attribute VB_Name = "frmImportaRedeSim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type tViabilidade
    Protocolo As String
    Analise As String
    Nire As String
    Cnpj As String
    EmpresaEstabelecida As String
    Cnae As String
    AtividadeAuxiliar As String
    DataProtocolo As Date
    DataResultadoAnalise As Date
    DataResultadoViabilidade As Date
    TempoAndamento As String
    CdEvento As String
    Evento As String
    Cep As String
    TipoInscricaoImovel As String
    NumeroInscricaoImovel As String
    TipoLogradouro As String
    Logradouro As String
    Bairro As String
    Complemento As String
    TipoUnidade As String
    FormaAtuacao As String
    Municipio As String
    RazaoSocial As String
    Orgao As String
    AreaImovel As Double
    AreaEstabelecimento As Double
    cpf As String
    NomeArquivo As String
    DataImportacao As Date
End Type

Private Type tLic
    ProtocoloLicenca As String
    IdSolicitacao As String
    SituacaoSolicitacao As String
    Orgao As String
    DataSolicitacaoLicenciamento As String
    IdLicenca As Long
    ProtocoloOrgao As String
    NumeroLicenca As String
    DetalheLicenca As String
    OrgaoLicenca As String
    Risco As String
    SituacaoLicenca As String
    DataEmissaoLicenca As String
    DataValidadeLicenca As String
    DataProtocolo As String
    Cnpj As String
    RazaoSocial As String
    TipoLogradouro As String
    Logradouro As String
    NumeroLogradouro As String
    Bairro As String
    Municipio As String
    Complemento As String
    Cep As String
    TipoInscricaoImovel As String
    NumeroInscricaoImovel As String
    PorteEmpresaMei As String
    EmpresaTeraEstabelecimento As String
    Cnae As String
    AtividadesAuxiliares As String
End Type

Private Type tEvento
    Codigo As Integer
    Nome As String
End Type

Dim aVia() As tViabilidade, aLic() As tLic

Private Sub CallPb(nVal As Long, nTot As Long)
If nVal > 0 Then
    pBar.Color = &HC0C000
Else
    pBar.Color = vbWhite
End If
If ((nVal * 100) / nTot) <= 100 Then
   pBar.value = (nVal * 100) / nTot
Else
   pBar.value = 100
End If

Me.Refresh
DoEvents

End Sub

Private Sub Form_Load()
Centraliza Me
frProgress.Visible = False
End Sub

Private Sub cmdArq_Click()
Dim fName As String, cc As cCommonDlg, ff As Long, sReg As String, aName() As String, nLinhas As Integer
Dim sTipo As String, nFile As Integer, itmX As ListItem
Dim vFiles As Variant
Dim lFile As Long
lvMain.ListItems.Clear


With CommonDialog1
    .FileName = "" 'Clear the filename
    .CancelError = False 'Gives an error if cancel is pressed
    .DialogTitle = "Select File(s)..."
    .flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNHideReadOnly 'Falgs, allows Multi select, Explorer style and hide the Read only tag
    .Filter = "All files (*.*)|*.*"
    .MaxFileSize = "32767"
    .ShowOpen
    vFiles = Split(.FileName, Chr(0)) 'Splits the filename up in segments
    If UBound(vFiles) = 0 Then ' If there is only 1 file then do this
        Open .FileName For Input As 37
        Do While Not EOF(37)
            Line Input #37, sReg
            Exit Do
        Loop
        Close 37
        sTipo = "Inválido"
'        If Left(sReg, 5) = "Cnpj;" Then sTipo = "Licenciamento"
        If Left(sReg, 3) = "Pro" Then sTipo = "Licenciamento"
        
        nLinhas = 0
        If sTipo <> "Inválido" Then
            Open .FileName For Input As 38
            Do While Not EOF(38)
                Line Input #38, sReg
                nLinhas = nLinhas + 1
            Loop
            Close #38
        End If
        
        Set itmX = lvMain.ListItems.Add(, , .FileName)
        itmX.SubItems(1) = sTipo
        itmX.SubItems(2) = nLinhas
    Else
        ff = FreeFile
        For lFile = 1 To UBound(vFiles) ' More than 1 file then do this until there are no more files
            Open vFiles(0) + "\" & vFiles(lFile) For Input As #ff
            Do While Not EOF(1)
                Line Input #ff, sReg
                Exit Do
            Loop
            Close #ff
            sTipo = "Inválido"
 '           If Left(sReg, 5) = "Cnpj;" Then sTipo = "Licenciamento"
            If Left(sReg, 3) = "Pro" Then sTipo = "Licenciamento"
            
            nLinhas = 0
            If sTipo <> "Inválido" Then
                Open vFiles(0) + "\" & vFiles(lFile) For Input As #ff
                Do While Not EOF(1)
                    Line Input #ff, sReg
                    nLinhas = nLinhas + 1
                Loop
                Close #ff
            End If
            Set itmX = lvMain.ListItems.Add(, , vFiles(0) + "\" & vFiles(lFile))
            itmX.SubItems(1) = sTipo
            itmX.SubItems(2) = nLinhas
        Next
    End If
End With

End Sub

Private Sub cmdGravar_Click()
Dim x As Integer, sTipo As String, sFileName As String, z As Integer, itmX As ListItem, ListaEvento() As tEvento, y As Integer, bFind As Boolean, sResultado As String
Dim Sql As String, RdoAux As rdoResultset, nPos As Long, nTot As Long, sDataResultado As String, nCodArquivo As Integer

ReDim aVia(0): ReDim aLic(0)
lvVia.ListItems.Clear
lvLic.ListItems.Clear

If lvMain.ListItems.Count = 0 Then
    MsgBox "Selecione os arquivos à importar", vbCritical, "Atenção"
    Exit Sub
End If


frProgress.Visible = True
For x = 1 To lvMain.ListItems.Count
    nPos = 1
    nTot = lvMain.ListItems(x).SubItems(2)
    lblProtocolo.Caption = ""
    
    sFileName = lvMain.ListItems(x).Text
    nCodArquivo = Grava_Arquivo(sFileName)
    sTipo = lvMain.ListItems(x).SubItems(1)
    If sTipo = "Inválido" Then
        '############
    ElseIf sTipo = "Viabilidade" Then
        ff = FreeFile
        Open sFileName For Input As #ff
        Debug.Print sTipo & vbCrLf
        Do While Not EOF(1)
            Line Input #ff, sReg
            If sReg = "" Then Exit Do
            aName = Split(sReg, ";")
            If aName(0) = "Protocolo" Then GoTo ProximoV
            lblFileNumber.Caption = "Importando Empresa " & nPos & " de " & nTot
            CallPb nPos, nTot
            ReDim Preserve aVia(UBound(aVia) + 1)
            z = UBound(aVia)
            aVia(z).Protocolo = aName(0)
            aVia(z).Analise = aName(1)
            aVia(z).Nire = aName(2)
            aVia(z).Cnpj = aName(3)
            aVia(z).EmpresaEstabelecida = aName(4)
            aVia(z).Cnae = aName(5)
            aVia(z).AtividadeAuxiliar = aName(6)
            aVia(z).DataProtocolo = aName(7)
            aVia(z).DataResultadoAnalise = aName(8)
            If IsDate(aName(9)) Then
                aVia(z).DataResultadoViabilidade = aName(9)
            End If
            aVia(z).TempoAndamento = aName(10)
            aVia(z).CdEvento = aName(11)
            aVia(z).Evento = aName(12)
            aVia(z).Cep = aName(13)
            aVia(z).TipoInscricaoImovel = aName(14)
            aVia(z).NumeroInscricaoImovel = aName(15)
            aVia(z).TipoLogradouro = aName(16)
            aVia(z).Logradouro = aName(17)
            aVia(z).Bairro = aName(18)
            aVia(z).Complemento = aName(19)
            aVia(z).TipoUnidade = aName(20)
            aVia(z).FormaAtuacao = aName(21)
            aVia(z).Municipio = aName(22)
            aVia(z).RazaoSocial = aName(23)
            aVia(z).Orgao = aName(24)
            aVia(z).AreaImovel = aName(25)
            aVia(z).AreaEstabelecimento = aName(26)
            aVia(z).cpf = aName(27)
            aVia(z).DataImportacao = Now
            lblProtocolo.Caption = aVia(z).Protocolo
            lblFileNumber.Caption = "Importando Empresa " & nPos & " de " & nTot
            Me.Refresh
            sResultado = "INCLUIDO"
            Sql = "select * from redesim_viabilidade where protocolo='" & aVia(z).Protocolo & "'"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux.RowCount = 0 Then
                'grava viabilidade
                With aVia(z)
                    If IsDate(.DataResultadoViabilidade) Then
                        If Year(.DataResultadoViabilidade) > 2019 Then
                            sDataResultado = "'" & Format(.DataResultadoViabilidade, "mm/dd/yyyy") & "'"
                        Else
                            sDataResultado = "Null"
                        End If
                    Else
                        sDataResultado = "Null"
                    End If
                    Sql = "INSERT INTO redesim_viabilidade(Protocolo,analise,nire,cnpj,empresaestabelecida,dataprotocolo,dataresultadoanalise,dataresultadoviabilidade,tempoandamento,"
                    Sql = Sql & "cep,tipoinscricaoimovel,numeroinscricaoimovel,tipologradouro,logradouro,bairro,complemento,tipounidade,formaatuacao,municipio,razaosocial,orgao,areaimovel,areaestabelecimento,"
                    Sql = Sql & "cpf,arquivo,data_importacao) values('" & .Protocolo & "','" & .Analise & "','" & .Nire & "','" & .Cnpj & "','" & .EmpresaEstabelecida & "','"
                    Sql = Sql & Format(.DataProtocolo, "mm/dd/yyyy") & "'," & IIf(IsDate(.DataResultadoAnalise), "Null", "'" & Format(.DataResultadoAnalise, "mm/dd/yyyy") & "'") & ","
                    Sql = Sql & sDataResultado & ",'" & .TempoAndamento & "','" & .Cep & "','" & .TipoInscricaoImovel & "','"
                    Sql = Sql & .NumeroInscricaoImovel & "','" & .TipoLogradouro & "','" & Mask(.Logradouro) & "','" & Mask(.Bairro) & "','" & Mask(.Complemento) & "','" & .TipoUnidade & "','" & .FormaAtuacao & "','"
                    Sql = Sql & .Municipio & "','" & Mask(.RazaoSocial) & "','" & .Orgao & "','" & Virg2Ponto(CStr(.AreaImovel)) & "','" & Virg2Ponto(CStr(.AreaEstabelecimento)) & "','" & .cpf & "'," & nCodArquivo & ",'" & Format(.DataImportacao, "mm/dd/yyyy hh:mm") & "')"
                    cn.Execute Sql, rdExecDirect
                End With
                
                'grava evento_processo
                ListaEvento = Grava_Evento(aVia(z).CdEvento, aVia(z).Evento)
                For y = 1 To UBound(ListaEvento)
                    Sql = "insert redesim_viabilidade_evento(protocolo,evento) values('" & aVia(z).Protocolo & "'," & ListaEvento(y).Codigo & ")"
                    cn.Execute Sql, rdExecDirect
                Next
                
                'grava atividade_auxiliar
                ListaEvento = Grava_Atividade_Auxiliar(aVia(z).AtividadeAuxiliar)
                For y = 1 To UBound(ListaEvento)
                    Sql = "insert redesim_viabilidade_atividade(protocolo,atividade) values('" & aVia(z).Protocolo & "'," & ListaEvento(y).Codigo & ")"
                    cn.Execute Sql, rdExecDirect
                Next
                
                'grava forma_atuacao
                ListaEvento = Grava_Forma_Atuacao(aVia(z).FormaAtuacao)
                For y = 1 To UBound(ListaEvento)
                    Sql = "insert redesim_viabilidade_forma_atuacao(protocolo,forma_atuacao) values('" & aVia(z).Protocolo & "'," & ListaEvento(y).Codigo & ")"
                    cn.Execute Sql, rdExecDirect
                Next
                
                'grava cnae
                aName = Split(aVia(z).Cnae, ",")
                For y = 0 To UBound(aName)
                    Sql = "insert redesim_viabilidade_cnae(protocolo,cnae) values('" & aVia(z).Protocolo & "','" & aName(y) & "')"
                    cn.Execute Sql, rdExecDirect
                Next
                
            Else
                sResultado = "JÁ INCLUSO"
            End If
            
            Set itmX = lvVia.ListItems.Add(, , aVia(z).Protocolo)
            itmX.SubItems(1) = aVia(z).RazaoSocial
            itmX.SubItems(2) = sResultado
ProximoV:
            nPos = nPos + 1
        Loop
        Close #ff
    ElseIf sTipo = "Licenciamento" Then
        ff = FreeFile
        Open sFileName For Input As #ff
        Debug.Print sTipo & vbCrLf
        Do While Not EOF(1)
            Line Input #ff, sReg
            If sReg = "" Then Exit Do
            aName = Split(sReg, ";")
            If aName(0) = "ProtocoloLicenca" Then GoTo ProximoL
            lblFileNumber.Caption = "Importando Empresa " & nPos & " de " & nTot
            CallPb nPos, nTot
            ReDim Preserve aLic(UBound(aLic) + 1)
            z = UBound(aLic)
            aLic(z).ProtocoloLicenca = aName(0)
            aLic(z).IdSolicitacao = aName(1)
            aLic(z).SituacaoSolicitacao = aName(2)
            aLic(z).Orgao = aName(3)
            aLic(z).DataSolicitacaoLicenciamento = aName(4)
            aLic(z).IdLicenca = aName(5)
            aLic(z).ProtocoloOrgao = aName(6)
            aLic(z).NumeroLicenca = aName(7)
            aLic(z).DetalheLicenca = aName(8)
            aLic(z).OrgaoLicenca = aName(9)
            aLic(z).Risco = aName(10)
            aLic(z).SituacaoLicenca = aName(11)
            aLic(z).DataEmissaoLicenca = aName(12)
            aLic(z).DataValidadeLicenca = aName(13)
            aLic(z).DataProtocolo = aName(14)
            aLic(z).Cnpj = aName(15)
            aLic(z).RazaoSocial = aName(16)
            aLic(z).TipoLogradouro = aName(17)
            aLic(z).Logradouro = aName(18)
            aLic(z).NumeroLogradouro = aName(19)
            aLic(z).Bairro = aName(20)
            aLic(z).Municipio = aName(21)
            aLic(z).Complemento = aName(22)
            aLic(z).Cep = aName(23)
            aLic(z).TipoInscricaoImovel = aName(24)
            If aName(24) = "Número IPTU" Then
                aLic(z).NumeroInscricaoImovel = aName(25)
            Else
                aLic(z).NumeroInscricaoImovel = 0
            End If
            aLic(z).PorteEmpresaMei = aName(26)
            aLic(z).EmpresaTeraEstabelecimento = aName(27)
            aLic(z).Cnae = aName(28)
            aLic(z).AtividadesAuxiliares = aName(29)
            lblFileNumber.Caption = "Importando Empresa " & nPos & " de " & nTot
            Me.Refresh
            sResultado = "INCLUIDO"
            Sql = "select * from redesim_licenciamento where protocololicenca='" & aLic(z).ProtocoloLicenca & "'"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux.RowCount = 0 Then
                'grava licenciamento
                With aLic(z)
                    Sql = "INSERT INTO redesim_licenciamento(ProtocoloLicenca,IdSolicitacao,DataSolicitacao,Cnpj,RazaoSocial,TipoLogradouro,Logradouro,NumeroLogradouro,Complemento,Cep,Bairro,Municipio,Cnae,CodigoImovel) "
                    Sql = Sql & "VALUES('" & .ProtocoloLicenca & "'," & .IdSolicitacao & ",'" & Format(.DataSolicitacaoLicenciamento, "mm/dd/yyyy") & "','" & .Cnpj & "','" & .RazaoSocial & "','" & .TipoLogradouro & "','"
                    Sql = Sql & .Logradouro & "'," & Val(RetornaNumero(.NumeroLogradouro)) & ",'" & .Complemento & "','" & .Cep & "','" & .Bairro & "','" & .Municipio & "','" & .Cnae & "'," & .NumeroInscricaoImovel & ")"
                    cn.Execute Sql, rdExecDirect
                End With
                sResultado = "IMPORTADO"
            Else
                sResultado = "JÁ INCLUSO"
            End If
            
            Set itmX = lvLic.ListItems.Add(, , aLic(z).ProtocoloLicenca)
            itmX.SubItems(1) = aLic(z).RazaoSocial
            itmX.SubItems(2) = sResultado
ProximoL:
            nPos = nPos + 1
        Loop
        Close #ff
    End If
    
Next
frProgress.Visible = False

End Sub

Private Function Grava_Evento(Codigo As String, Descricao As String) As tEvento()
Dim aLista() As tEvento, x As Integer, codevento As Integer, RdoAux As rdoResultset, Sql As String, y As Integer, bFind As Boolean

ReDim aLista(0)
acod = Split(Codigo, ",")
aName = Split(Descricao, ",")
For x = 0 To UBound(acod)
    bFind = False
    For y = 0 To UBound(aLista)
        If aLista(y).Codigo = acod(x) Then
            bFind = True
            Exit For
        End If
    Next
    If Not bFind Then
        ReDim Preserve aLista(UBound(aLista) + 1)
        aLista(UBound(aLista)).Codigo = acod(x)
        aLista(UBound(aLista)).Nome = aName(x)
    End If
    
    Sql = "select * from redesim_evento where codigo=" & acod(x)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount = 0 Then
        Sql = "insert redesim_evento(codigo,nome) values(" & acod(x) & ",'" & aName(x) & "')"
        cn.Execute Sql, rdExecDirect
    End If
    RdoAux.Close
Next

Grava_Evento = aLista

End Function

Private Function Grava_Atividade_Auxiliar(Descricao As String) As tEvento()
Dim aLista() As tEvento, x As Integer, nCodigo As Integer, RdoAux As rdoResultset, Sql As String, y As Integer, bFind As Boolean

ReDim aLista(0)
aName = Split(Descricao, ",")
For x = 0 To UBound(aName)
    bFind = False
    For y = 0 To UBound(aLista)
        If aLista(y).Nome = aName(x) Then
            bFind = True
            Exit For
        End If
    Next
    If Not bFind Then
        ReDim Preserve aLista(UBound(aLista) + 1)
        aLista(UBound(aLista)).Nome = aName(x)
    End If
    
    Sql = "select * from redesim_atividade_auxiliar where nome='" & aName(x) & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount = 0 Then
        Sql = "select max(codigo) as maximo from redesim_atividade_auxiliar"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux!maximo) Then
            nCodigo = 1
        Else
            nCodigo = RdoAux!maximo + 1
        End If
        aLista(UBound(aLista)).Codigo = nCodigo
        Sql = "insert redesim_atividade_auxiliar(codigo,nome) values(" & nCodigo & ",'" & aName(x) & "')"
        cn.Execute Sql, rdExecDirect
    Else
        aLista(UBound(aLista)).Codigo = RdoAux!Codigo
    End If
    RdoAux.Close
Next

Grava_Atividade_Auxiliar = aLista

End Function

Private Function Grava_Forma_Atuacao(Descricao As String) As tEvento()
Dim aLista() As tEvento, x As Integer, nCodigo As Integer, RdoAux As rdoResultset, Sql As String, y As Integer, bFind As Boolean

ReDim aLista(0)
aName = Split(Descricao, ",")
For x = 0 To UBound(aName)
    bFind = False
    For y = 0 To UBound(aLista)
        If aLista(y).Nome = aName(x) Then
            bFind = True
            Exit For
        End If
    Next
    If Not bFind Then
        ReDim Preserve aLista(UBound(aLista) + 1)
        aLista(UBound(aLista)).Nome = aName(x)
    End If
    
    Sql = "select * from redesim_forma_atuacao where nome='" & aName(x) & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount = 0 Then
        Sql = "select max(codigo) as maximo from redesim_forma_atuacao"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux!maximo) Then
            nCodigo = 1
        Else
            nCodigo = RdoAux!maximo + 1
        End If
        aLista(UBound(aLista)).Codigo = nCodigo
        Sql = "insert redesim_forma_atuacao(codigo,nome) values(" & nCodigo & ",'" & aName(x) & "')"
        cn.Execute Sql, rdExecDirect
    Else
        aLista(UBound(aLista)).Codigo = RdoAux!Codigo
    End If
    RdoAux.Close
Next

Grava_Forma_Atuacao = aLista

End Function

Private Function Grava_Arquivo(Descricao As String) As Integer
Dim aLista() As tEvento, x As Integer, nCodigo As Integer, RdoAux As rdoResultset, Sql As String, y As Integer, bFind As Boolean
    
Sql = "select * from redesim_arquivo where nome='" & Descricao & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount = 0 Then
    Sql = "select max(codigo) as maximo from redesim_arquivo"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If IsNull(RdoAux!maximo) Then
        nCodigo = 1
    Else
        nCodigo = RdoAux!maximo + 1
    End If
    Sql = "insert redesim_arquivo(codigo,nome) values(" & nCodigo & ",'" & Descricao & "')"
    cn.Execute Sql, rdExecDirect
Else
    nCodigo = RdoAux!Codigo
End If
RdoAux.Close

Grava_Arquivo = nCodigo

End Function

