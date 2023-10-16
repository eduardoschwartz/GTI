VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmsc_arquivo_saaej 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importação de arquivo do SAAEJ"
   ClientHeight    =   840
   ClientLeft      =   7155
   ClientTop       =   5130
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   840
   ScaleWidth      =   6675
   Begin Tributacao.XP_ProgressBar PBar 
      Height          =   240
      Left            =   270
      TabIndex        =   0
      Top             =   360
      Width           =   4380
      _ExtentX        =   7726
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
   Begin prjChameleon.chameleonButton btLoad 
      Height          =   375
      Left            =   4860
      TabIndex        =   1
      ToolTipText     =   "Carregar Registro(s)"
      Top             =   270
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Abrir e Importar"
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
      MICON           =   "frmsc_arquivo_saaej.frx":0000
      PICN            =   "frmsc_arquivo_saaej.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   1170
      Top             =   5625
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Color           =   8388608
      DialogTitle     =   "Selecione o arquivo de GIA"
   End
End
Attribute VB_Name = "frmsc_arquivo_saaej"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btLoad_Click()
Dim fso As FileSystemObject, TS As TextStream, Row, nPos As Long, nTot As Long, vMes As Variant, vAno As Variant, sql, RdoAux As rdoResultset

On Error GoTo Erro:

With cDialog
    .FileName = "" 'Clear the filename
    .CancelError = True
    .MaxFileSize = 30000
    .DialogTitle = "Select File"
    .flags = cdlOFNExplorer Or cdlOFNHideReadOnly 'Flags, allows Multi select, Explorer style and hide the Read only tag
    .Filter = "Text files (*.txt)|*.txt"
    .ShowOpen
    
    PBar.value = 0
    nPos = 1
    nTot = FileRowCount(.FileName)
    '#############ImportaArquivo ###################
    Set fso = New FileSystemObject

    Set TS = fso.OpenTextFile(.FileName, ForReading)
    Row = TS.ReadLine
    vMes = Mid(Row, 41, 2)
    vAno = Mid(Row, 46, 4)
    
    'Verifica se o arquivo é válido
    If Not IsNumeric(vMes) Or Not IsNumeric(vAno) Then
        TS.Close
        Liberado
        MsgBox "Arquivo inválido!", vbCritical, "Erro"
        Exit Sub
    Else
        If (Val(vMes) < 1 Or Val(vMes) > 12) Or (Val(vAno < 2020) Or Val(vAno > 2030)) Then
            TS.Close
            Liberado
            MsgBox "Arquivo inválido!", vbCritical, "Erro"
            Exit Sub
        End If
    End If
    TS.Close
    
    If MsgBox("Deseja importar este arquivo?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
        TS.Close
        Liberado
        Exit Sub
    End If
    
    'Verifica se o arquivo já existe
    sql = "select top(1) * from sc_arquivo_saaej where mesref=" & Val(vMes) & " and anoref=" & Val(vAno)
    Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount > 0 Then
        If MsgBox("Este arquivo já foi importado, deseja sobrescreve-lo?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
            TS.Close
            Liberado
            Exit Sub
        Else
            sql = "delete from sc_arquivo_saaej where mesref=" & Val(vMes) & " and anoref=" & Val(vAno)
            cn.Execute sql, rdExecDirect
            sql = "delete from sc_ligacao_agua_resumo where mes=" & Val(vMes) & " and ano=" & Val(vAno)
            cn.Execute sql, rdExecDirect
        End If
    End If
    RdoAux.Close
    
    'Arquivo válido
    Ocupado
    ImportaArquivo .FileName, nTot
    Liberado
End With

PBar.value = 0
MsgBox "Importação concluída!", vbInformation, "Atenção"

Exit Sub
Erro:
If Err.Number = 32755 Then
Else
    MsgBox Err.Description
End If

End Sub

Private Sub Form_Load()
Centraliza Me
End Sub

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

Private Function FileRowCount(sFile)
    Const BufSize As Long = 100000
    Dim T0 As Single
    Dim LfAnsi
    Dim f
    Dim FileBytes As Long
    Dim BytesLeft As Long
    Dim Buffer() As Byte
    Dim strBuffer
    Dim BufPos As Long
    Dim LineCount As Long

    T0 = Timer()
    LfAnsi = StrConv(vbLf, vbFromUnicode)
    f = FreeFile(0)
    Open sFile For Binary Access Read As #f
    FileBytes = LOF(f)
    ReDim Buffer(BufSize - 1)
    BytesLeft = FileBytes
    Do Until BytesLeft = 0
        If BufPos = 0 Then
            If BytesLeft < BufSize Then ReDim Buffer(BytesLeft - 1)
            Get #f, , Buffer
            strBuffer = Buffer 'Binary copy of bytes.
            BytesLeft = BytesLeft - LenB(strBuffer)
            BufPos = 1
        End If
        Do Until BufPos = 0
            BufPos = InStrB(BufPos, strBuffer, LfAnsi)
            If BufPos > 0 Then
                LineCount = LineCount + 1
                BufPos = BufPos + 1
            End If
        Loop
    Loop
    Close #f
    FileRowCount = LineCount
End Function

Private Sub ImportaArquivo(sFileName, nTotal As Long)
Dim fso As FileSystemObject, TS As TextStream, Row, nPos As Long, nTot As Long, sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset
Dim idread As Long, idfatura As Long, idcodlig As Long, idparcela As Long, idgrupoentrega As Long, MesRef As Integer, AnoRef As Integer, idleitura As Long, imovelnumero As Long
Dim entreganumero As Long, codloteamento As Long, Categoria As Long, cdcodgrupo As Long, cdcodsetor As Long, leitura As Long, qtleituracalc As Long, qtmediames As Long, qtconsumomedido As Long
Dim qtconsumocalcul As Long, qtmediaano As Long, qtleituraant As Long, imovelidlogradouro As Long, entregaidlogradouro As Long, dvcodligant As Long, nrodias As Integer, qtdeconomia As Long, idordem As Long, Inscricao As String
Dim Situacao As String, imovellogradouro As String, imovelcomplemento As String, imovelbairro As String, imovelcep As String, imovelcidade As String, imovelestado As String
Dim Proprietario As String, USUARIO As String, Quadra As String, Lote As String, Unidade As String, descrcategoria As String, descrincidencia As String, imovelendereco As String, hidrometro As String
Dim dcdescnormitem As String, digitocodigo As String, dataleitura As String, dtvencimento As String, dataleituraant  As String, dtcorte  As String, vlrtotal As Double, referencia1 As String, consumoanterior1 As Long
Dim referencia2 As String, consumoanterior2 As Long, referencia3 As String, consumoanterior3 As Long, referencia4 As String, consumoanterior4 As Long, referencia5 As String, consumoanterior5 As Long
Dim referencia6 As String, consumoanterior6 As Long, referencia7 As String, consumoanterior7 As Long, referencia8 As String, consumoanterior8 As Long, referencia9 As String, consumoanterior9 As Long
Dim referencia10 As String, consumoanterior10 As Long, referencia11 As String, consumoanterior11 As Long, referencia12 As String, consumoanterior12 As Long, MensagemLin1 As String, MensagemLin2_1 As String
Dim MensagemLin2_2  As String, MensagemLin2_3 As String, MensagemLin3 As String, MensagemLin4  As String, dias As Integer, bAgua As Boolean, bEsgoto As Boolean
Dim codigoLig As Integer, codigo_saaej As Long, inscricao_saaej As String, imovel_logradouro As String, imovel_numero As Integer, imovel_complemento As String, imovel_bairro As String, situacao_ligacao As Integer

On Error GoTo Erro

nTot = nTotal
Set fso = New FileSystemObject
Set TS = fso.OpenTextFile(sFileName, ForReading)
Do Until TS.AtEndOfStream
    If nPos Mod 10 = 0 Then CallPb nPos, nTot
    Row = TS.ReadLine
    idread = Val(Mid(Row, 1, 7))
    idfatura = Val(Mid(Row, 8, 7))
    idcodlig = Val(Mid(Row, 15, 7))
    idparcela = Val(Mid(Row, 22, 7))
    idgrupoentrega = Val(Mid(Row, 29, 7))
    MesRef = Val(Mid(Row, 36, 7))
    AnoRef = Val(Mid(Row, 43, 7))
    idleitura = Val(Mid(Row, 50, 7))
    imovelnumero = Val(Mid(Row, 57, 7))
    entreganumero = Val(Mid(Row, 64, 7))
    codloteamento = Val(Mid(Row, 71, 7))
    Categoria = Val(Mid(Row, 78, 7))
    cdcodgrupo = Val(Mid(Row, 85, 7))
    cdcodsetor = Val(Mid(Row, 92, 7))
    leitura = Val(Mid(Row, 99, 7))
    qtleituracalc = Val(Mid(Row, 106, 7))
    qtmediames = Val(Mid(Row, 113, 7))
    qtconsumomedido = Val(Mid(Row, 120, 7))
    qtconsumocalcul = Val(Mid(Row, 127, 7))
    qtmediaano = Val(Mid(Row, 134, 7))
    qtleituraant = Val(Mid(Row, 141, 7))
    imovelidlogradouro = Val(Mid(Row, 148, 7))
    entregaidlogradouro = Val(Mid(Row, 155, 7))
    dvcodligant = Val(Mid(Row, 162, 7))
    nrodias = Val(Mid(Row, 169, 7))
    qtdeconomia = Val(Mid(Row, 176, 7))
    idordem = Val(Mid(Row, 183, 7))
    Inscricao = Trim(Mid(Row, 190, 40))
    Situacao = Trim(Mid(Row, 230, 10))
    hidrometro = Trim(Mid(Row, 240, 30))
    imovellogradouro = Trim(Mid(Row, 270, 120))
    imovelcomplemento = Trim(Mid(Row, 390, 40))
    imovelbairro = Trim(Mid(Row, 430, 30))
    imovelcep = Trim(Mid(Row, 460, 9))
    imovelcidade = Trim(Mid(Row, 469, 40))
    imovelestado = Trim(Mid(Row, 509, 2))
    Proprietario = Trim(Mid(Row, 828, 50))
    USUARIO = Trim(Mid(Row, 878, 50))
    Quadra = Trim(Mid(Row, 928, 15))
    Lote = Trim(Mid(Row, 943, 15))
    Unidade = Trim(Mid(Row, 958, 5))
    descrcategoria = Trim(Mid(Row, 965, 50))
    descrincidencia = Trim(Mid(Row, 1015, 15))
    imovelendereco = Trim(Mid(Row, 1030, 120))
    dcdescnormitem = Trim(Mid(Row, 1270, 60))
    digitocodigo = Trim(Mid(Row, 1330, 15))
    dataleitura = Trim(Mid(Row, 1345, 10))
    dtvencimento = Trim(Mid(Row, 1355, 10))
    dataleituraant = Trim(Mid(Row, 1365, 10))
    dtcorte = Trim(Mid(Row, 1375, 10))
    vlrtotal = CDbl(Mid(Row, 1385, 16))
    referencia1 = Trim(Mid(Row, 1401, 7))
    consumoanterior1 = Val(Mid(Row, 1408, 7))
    referencia2 = Trim(Mid(Row, 1415, 7))
    consumoanterior2 = Val(Mid(Row, 1422, 7))
    referencia3 = Trim(Mid(Row, 1429, 7))
    consumoanterior3 = Val(Mid(Row, 1436, 7))
    referencia4 = Trim(Mid(Row, 1443, 7))
    consumoanterior4 = Val(Mid(Row, 1450, 7))
    referencia5 = Trim(Mid(Row, 1457, 7))
    consumoanterior5 = Val(Mid(Row, 1464, 7))
    referencia6 = Trim(Mid(Row, 1471, 7))
    consumoanterior6 = Val(Mid(Row, 1478, 7))
    referencia7 = Trim(Mid(Row, 1485, 7))
    consumoanterior7 = Val(Mid(Row, 1492, 7))
    referencia8 = Trim(Mid(Row, 1499, 7))
    consumoanterior8 = Val(Mid(Row, 1506, 7))
    referencia9 = Trim(Mid(Row, 1513, 7))
    consumoanterior9 = Val(Mid(Row, 1520, 7))
    referencia10 = Trim(Mid(Row, 1527, 7))
    consumoanterior10 = Val(Mid(Row, 1534, 7))
    referencia11 = Trim(Mid(Row, 1541, 7))
    consumoanterior11 = Val(Mid(Row, 1548, 7))
    referencia12 = Trim(Mid(Row, 1555, 7))
    consumoanterior12 = Val(Mid(Row, 1562, 7))
    MensagemLin1 = Trim(Mid(Row, 5893, 160))
    MensagemLin2_1 = Trim(Mid(Row, 6053, 250))
    MensagemLin2_2 = Trim(Mid(Row, 6303, 250))
    MensagemLin2_3 = Trim(Mid(Row, 6553, 220))
    MensagemLin3 = Trim(Mid(Row, 6773, 250))
    MensagemLin4 = Trim(Mid(Row, 7023, 50))
        
    Select Case UCase(Situacao)
        Case "LIGADA"
            situacao_ligacao = 1
        Case "CORTADA"
            situacao_ligacao = 2
        Case "RELIGADA"
            situacao_ligacao = 3
        Case Else
            situacao_ligacao = 1
    End Select
    
    bAgua = False: bEsgoto = False
    Select Case UCase(dcdescnormitem)
        Case "AGUA"
            bAgua = True
        Case "ESGOTO"
            bEsgoto = True
        Case "AGUA E ESGOTO"
            bAgua = True: bEsgoto = True
    End Select
    
    
    sql = "insert sc_arquivo_saaej(idread,idfatura,idcodlig,idparcela,idgrupoentrega,mesref,anoref,idleitura,imovelnumero,entreganumero,codloteamento,categoria,cdcodgrupo,cdcodsetor,leitura,qtleituracalc,qtmediames,"
    sql = sql & "qtconsumomedido,qtconsumocalcul,qtmediaano,qtleituraant,imovelidlogradouro,entregaidlogradouro,dvcodligant,nrodias,qtdeconomia,idordem,inscricao,situacao,hidrometro,imovellogradouro,imovelcomplemento,"
    sql = sql & "imovelbairro,imovelcep,imovelcidade,imovelestado,proprietario,usuario,quadra,lote,unidade,descrcategoria,descrincidencia,imovelendereco,dcdescnormitem,digitocodigo,dataleitura,dtvencimento,dataleituraant,"
    sql = sql & "dtcorte, vlrtotal,referencia1,consumoanterior1,referencia2,consumoanterior2,referencia3,consumoanterior3,referencia4,consumoanterior4,referencia5,consumoanterior5,referencia6,consumoanterior6,"
    sql = sql & "referencia7,consumoanterior7,referencia8,consumoanterior8,referencia9,consumoanterior9,referencia10,consumoanterior10,referencia11,consumoanterior11,referencia12,consumoanterior12,MensagemLin1,"
    sql = sql & "MensagemLin2_1,MensagemLin2_2,MensagemLin2_3,MensagemLin3,MensagemLin4) values("
    sql = sql & idread & "," & idfatura & "," & idcodlig & "," & idparcela & "," & idgrupoentrega & "," & MesRef & "," & AnoRef & "," & idleitura & "," & imovelnumero & "," & entreganumero & "," & codloteamento & ","
    sql = sql & Categoria & "," & cdcodgrupo & "," & cdcodsetor & "," & leitura & "," & qtleituracalc & "," & qtmediames & "," & qtconsumomedido & "," & qtconsumocalcul & "," & qtmediaano & "," & qtleituraant & ","
    sql = sql & imovelidlogradouro & "," & entregaidlogradouro & "," & dvcodligant & "," & nrodias & "," & qtdeconomia & "," & idordem & "," & sNull(Inscricao) & "," & sNull(Situacao) & "," & sNull(hidrometro) & ","
    sql = sql & sNull(imovellogradouro) & "," & sNull(imovelcomplemento) & "," & sNull(imovelbairro) & "," & sNull(imovelcep) & "," & sNull(imovelcidade) & "," & sNull(imovelestado) & "," & sNull(Proprietario) & ","
    sql = sql & sNull(USUARIO) & "," & sNull(Quadra) & "," & sNull(Lote) & "," & sNull(Unidade) & "," & sNull(descrcategoria) & "," & sNull(descrincidencia) & "," & sNull(imovelendereco) & "," & sNull(dcdescnormitem) & ","
    sql = sql & sNull(digitocodigo) & "," & sNullData(dataleitura) & "," & sNullData(dtvencimento) & "," & sNullData(dataleituraant) & "," & sNullData(dtcorte) & "," & Virg2Ponto(CStr(vlrtotal)) & "," & sNull(referencia1) & ","
    sql = sql & consumoanterior1 & "," & sNull(referencia2) & "," & consumoanterior2 & "," & sNull(referencia3) & "," & consumoanterior3 & "," & sNull(referencia4) & "," & consumoanterior4 & "," & sNull(referencia5) & ","
    sql = sql & consumoanterior5 & "," & sNull(referencia6) & "," & consumoanterior6 & "," & sNull(referencia7) & "," & consumoanterior7 & "," & sNull(referencia8) & "," & consumoanterior8 & "," & sNull(referencia9) & ","
    sql = sql & consumoanterior9 & "," & sNull(referencia10) & "," & consumoanterior10 & "," & sNull(referencia11) & "," & consumoanterior11 & "," & sNull(referencia12) & "," & consumoanterior12 & "," & sNull(MensagemLin1) & ","
    sql = sql & sNull(MensagemLin2_1) & "," & sNull(MensagemLin2_2) & "," & sNull(MensagemLin2_3) & "," & sNull(MensagemLin3) & "," & sNull(MensagemLin4) & ")"
    cn.Execute sql, rdExecDirect
    
    'grava sc_ligacao_agua
    sql = "select codigo,inscricao_saaej from sc_ligacao_agua where codigo_saaej=" & idcodlig
    Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount = 0 Then
        sql = "select max(codigo) as maximo from sc_ligacao_agua"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux2!maximo) Then
            codigoLig = 1
        Else
            codigoLig = RdoAux2!maximo + 1
        End If
        RdoAux2.Close
        
        codigo_saaej = idcodlig
        inscricao_saaej = Inscricao
        imovel_logradouro = imovellogradouro
        imovel_numero = imovelnumero
        imovel_complemento = imovelcomplemento
        imovel_bairro = imovelbairro
        situacao_ligacao = 1
        
        sql = "insert sc_ligacao_agua(codigo,secretaria,codigo_saaej,inscricao_saaej,proprietario,usuario,imovel_logradouro,imovel_numero,imovel_complemento,imovel_bairro) values(" & codigoLig & "," & 0 & "," & codigo_saaej & ","
        sql = sql & sNull(inscricao_saaej) & "," & sNull(Proprietario) & "," & sNull(USUARIO) & "," & sNull(imovel_logradouro) & "," & imovel_numero & "," & sNull(imovel_complemento) & "," & sNull(imovel_bairro) & ")"
    Else
        codigoLig = RdoAux!codigo
        sql = "update sc_ligacao_agua set inscricao_saaej=" & sNull(Inscricao) & ",proprietario=" & sNull(Proprietario) & ",usuario=" & sNull(USUARIO) & ",imovel_logradouro=" & sNull(imovellogradouro) & ","
        sql = sql & "imovel_numero=" & imovelnumero & ",imovel_complemento=" & sNull(imovelcomplemento) & ",imovel_bairro=" & sNull(imovelbairro) & " where codigo_saaej=" & idcodlig
    End If
    cn.Execute sql, rdExecDirect
    RdoAux.Close
    
    'grava sc_ligacao_agua_resumo
    sql = "insert sc_ligacao_agua_resumo(codigo,ano,mes,hidrometro,situacao,data_leitura,data_vencimento,data_corte,data_leitura_anterior,mensagem1,mensagem2,mensagem3,mensagem4,mensagem5,mensagem6,agua,esgoto) values("
    sql = sql & codigoLig & "," & AnoRef & "," & MesRef & "," & sNull(hidrometro) & "," & situacao_ligacao & "," & sNullData(dataleitura) & "," & sNullData(dtvencimento) & "," & sNullData(dtcorte) & "," & sNullData(dataleituraant) & ","
    sql = sql & sNull(MensagemLin1) & "," & sNull(MensagemLin2_1) & "," & sNull(MensagemLin2_2) & "," & sNull(MensagemLin2_3) & "," & sNull(MensagemLin3) & "," & sNull(MensagemLin4) & "," & IIf(bAgua, 1, 0) & "," & IIf(bEsgoto, 1, 0) & ")"
    cn.Execute sql, rdExecDirect
    
    'grava sc_ligacao_agua_consumo
    Grava_Consumo codigoLig, AnoRef, MesRef, leitura, qtleituracalc, qtleituraant, qtmediames, qtmediaano, qtconsumomedido, qtconsumocalcul, vlrtotal, nrodias
    
    AnoRef = Right(referencia1, 4): MesRef = Left(referencia1, 2)
    Grava_Consumo_Anterior codigoLig, AnoRef, MesRef, consumoanterior1
    
    AnoRef = Right(referencia2, 4): MesRef = Left(referencia2, 2)
    Grava_Consumo_Anterior codigoLig, AnoRef, MesRef, consumoanterior2
    
    AnoRef = Right(referencia3, 4): MesRef = Left(referencia3, 2)
    Grava_Consumo_Anterior codigoLig, AnoRef, MesRef, consumoanterior3
    
    AnoRef = Right(referencia4, 4): MesRef = Left(referencia4, 2)
    Grava_Consumo_Anterior codigoLig, AnoRef, MesRef, consumoanterior4
    
    AnoRef = Right(referencia5, 4): MesRef = Left(referencia5, 2)
    Grava_Consumo_Anterior codigoLig, AnoRef, MesRef, consumoanterior5
    
    AnoRef = Right(referencia6, 4): MesRef = Left(referencia6, 2)
    Grava_Consumo_Anterior codigoLig, AnoRef, MesRef, consumoanterior6
    
    AnoRef = Right(referencia7, 4): MesRef = Left(referencia7, 2)
    Grava_Consumo_Anterior codigoLig, AnoRef, MesRef, consumoanterior7
    
    AnoRef = Right(referencia8, 4): MesRef = Left(referencia8, 2)
    Grava_Consumo_Anterior codigoLig, AnoRef, MesRef, consumoanterior8
    
    AnoRef = Right(referencia9, 4): MesRef = Left(referencia9, 2)
    Grava_Consumo_Anterior codigoLig, AnoRef, MesRef, consumoanterior9
    
    AnoRef = Right(referencia10, 4): MesRef = Left(referencia10, 2)
    Grava_Consumo_Anterior codigoLig, AnoRef, MesRef, consumoanterior10
    
    AnoRef = Right(referencia11, 4): MesRef = Left(referencia11, 2)
    Grava_Consumo_Anterior codigoLig, AnoRef, MesRef, consumoanterior11
    
    AnoRef = Right(referencia12, 4): MesRef = Left(referencia12, 2)
    Grava_Consumo_Anterior codigoLig, AnoRef, MesRef, consumoanterior12

    
    nPos = nPos + 1
Loop

Exit Sub
Erro:
MsgBox Err.Description
Resume Next
End Sub

Private Sub Grava_Consumo(codigo As Integer, ano As Integer, mes As Integer, leitura As Long, leitura_calc As Long, leitura_anterior As Long, media_mes As Long, media_ano As Long, consumo_medido As Long, consumo_calc As Long, valor As Double, dias As Integer)

sql = "select * from sc_ligacao_agua_consumo where codigo=" & codigo & " and ano=" & ano & " and mes=" & mes
Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
If RdoAux2.RowCount = 0 Then
    sql = "insert sc_ligacao_agua_consumo(codigo,ano,mes,leitura,leitura_calc,leitura_anterior,media_mes,media_ano,consumo_medido,consumo_calc,valor,dias) values(" & codigo & "," & ano & "," & mes & "," & leitura & "," & leitura & ","
    sql = sql & leitura_anterior & "," & media_mes & "," & media_ano & "," & consumo_medido & "," & consumo_calc & "," & Virg2Ponto(CStr(valor)) & "," & dias & ")"
    cn.Execute sql, rdExecDirect
Else
    If Val(SubNull(RdoAux2!leitura)) <> leitura And leitura > 0 Then
        sql = "update sc_ligacao_agua_consumo set leitura=" & leitura & ",leitura_calc=" & leitura_calc & ",leitura_anterior=" & leitura_anterior & ",media_mes=" & media_mes & ",media_ano=" & media_ano & ","
        sql = sql & "consumo_medido=" & consumo_medido & ",consumo_calc=" & consumo_calc & ",valor=" & Virg2Ponto(CStr(valor)) & ",dias=" & dias & " where codigo=" & codigo & " and ano=" & ano & " and mes=" & mes
        cn.Execute sql, rdExecDirect
    End If
End If

End Sub

Private Sub Grava_Consumo_Anterior(codigo As Integer, ano As Integer, mes As Integer, consumo As Long)

sql = "select * from sc_ligacao_agua_consumo where codigo=" & codigo & " and ano=" & ano & " and mes=" & mes
Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
If RdoAux2.RowCount = 0 Then
    sql = "insert sc_ligacao_agua_consumo(codigo,ano,mes,consumo_calc) values(" & codigo & "," & ano & "," & mes & "," & consumo & ")"
    cn.Execute sql, rdExecDirect
Else
    If RdoAux2!leitura <> leitura And leitura > 0 Then
        sql = "update sc_ligacao_agua_consumo set consumo_calc=" & consumo & " where codigo=" & codigo & " and ano=" & ano & " and mes=" & mes
        cn.Execute sql, rdExecDirect
    End If
End If

End Sub

