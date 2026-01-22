VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmDataBase 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Base"
   ClientHeight    =   3690
   ClientLeft      =   4965
   ClientTop       =   2835
   ClientWidth     =   3855
   Icon            =   "frmDataBase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3690
   ScaleWidth      =   3855
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   2760
      TabIndex        =   1
      ToolTipText     =   "Sair da Tela"
      Top             =   3300
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Sair"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmDataBase.frx":08CA
      PICN            =   "frmDataBase.frx":08E6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.MonthView Mv 
      Height          =   3210
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   5662
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   16777215
      ShowToday       =   0   'False
      StartOfWeek     =   138412033
      TitleBackColor  =   192
      TitleForeColor  =   16777215
      TrailingForeColor=   12632256
      CurrentDate     =   44197
      MaxDate         =   46387
      MinDate         =   43831
   End
   Begin VB.Label lblDB 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   2
      Top             =   4080
      Width           =   2955
   End
End
Attribute VB_Name = "frmDataBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sData As String
Dim sOldData As String
Dim sql As String

Private Sub cmdSair_Click()

If sOldData <> sData Then
    If MsgBox("Deseja alterar a Data Base para " & sData, vbQuestion + vbYesNo, "Confirmação") = vbYes Then
       sql = "UPDATE PARAMETROS SET VALPARAM='" & sData & "' WHERE NOMEPARAM='DATABASE'"
       cn.Execute sql, rdExecDirect
       frmMdi.Sbar.Panels(6).Text = "Data Base: " & sData
    End If
    Ocupado
    sql = "UPDATE PROCESSOGTI SET DATAARQUIVA='" & Format(Now, "mm/dd/yyyy") & "' "
    sql = sql & " Where (FISICO = 0) And (DATAARQUIVA Is Null) And (DATACANCEL Is Null) And (DATASUSPENSO Is Null) AND (CODASSUNTO NOT IN (567,1109,1110))"
    cn.Execute sql, rdExecDirect
    Liberado
End If
If frmMdi.frTeste.Visible = False Then
    sql = "update debitoparcela SET statuslanc=5 WHERE codlancamento=41 AND statuslanc=3 AND datavencimento<'" & Format(Now, "mm/dd/yyyy") & "'"
    cn.Execute sql, rdExecDirect

    AtualizaITBI
    Cancela_ITBIs_Vencidos
    AtualizaNotificacaoISS
    AtualizaUsoPlataforma
    AtualizaIntegrativa
    Corrige_Terreno_sem_Endereco
    AtualizaSN
    Atualiza_Giss_Guia
End If

Unload Me
End Sub

Public Sub AtualizaITBI()
Dim sql As String, RdoAux As rdoResultset, nNumDoc As Long, ibti As String, RdoAux2 As rdoResultset, bPago As Boolean, sObs As String, nCodReduz As Long

'Sql = "SELECT * FROM itbi_main WHERE situacao_itbi=2"
sql = "SELECT * FROM itbi_main WHERE situacao_itbi=2 AND itbi_ano>=" & (Year(Now) - 1)
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        bPago = False
        nNumDoc = !numero_guia
        'If nNumDoc = 21380663 Then MsgBox "teste"
        itbi = Format(!itbi_numero, "00000") & "/" & Format(!itbi_ano, "0000")
        sql = "select * from debitopago where numdocumento=" & nNumDoc
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            bPago = True
        End If
        RdoAux2.Close
        If bPago Then
            sql = "update itbi_main set situacao_itbi=3 where guid='" & !guid & "'"
            cn.Execute sql, rdExecDirect
            
            nCodReduz = !imovel_codigo
            If nCodReduz > 0 Then
                sObs = "Pagamento do ITBI nº:" & itbi
                sql = "SELECT MAX(SEQ) AS MAXIMO FROM HISTORICO WHERE CODREDUZIDO=" & nCodReduz
                Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                    If IsNull(!maximo) Then
                        nSeq = 1
                    Else
                        nSeq = !maximo + 1
                    End If
                   .Close
                End With
                'Sql = "INSERT HISTORICO(CODREDUZIDO,SEQ,DATAHIST,DESCHIST,USERID,DATAHIST2) VALUES(" & nCodReduz & "," & nSeq & ",'" & Format(Now, "dd/mm/yyyy") & "','" & sObs & "'," & !liberado_por & ",'" & Format(Now, "mm/dd/yyyy") & "')"
                sql = "INSERT HISTORICO(CODREDUZIDO,SEQ,DATAHIST,DESCHIST,USERID,DATAHIST2) VALUES(" & nCodReduz & "," & nSeq & ",'" & Format(Now, "dd/mm/yyyy") & "','" & sObs & "'," & 236 & ",'" & Format(Now, "mm/dd/yyyy") & "')"
                cn.Execute sql, rdExecDirect
            End If
        End If
       .MoveNext
    Loop
   .Close
End With


End Sub

Public Sub AtualizaNotificacaoISS()
Dim sql As String, RdoAux As rdoResultset, nNumDoc As Long, ibti As String, RdoAux2 As rdoResultset, bPago As Boolean, sObs As String, nCodReduz As Long

sql = "SELECT DISTINCT numero_guia FROM notificacao_iss_web where situacao=2"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        bPago = False
        nNumDoc = !numero_guia
        sql = "select * from debitopago where numdocumento=" & nNumDoc
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            bPago = True
        End If
        RdoAux2.Close
        If bPago Then
            sql = "update notificacao_iss_web set situacao=6 where numero_guia=" & nNumDoc
            cn.Execute sql, rdExecDirect
        End If
       .MoveNext
    Loop
   .Close
End With


End Sub

Public Sub AtualizaUsoPlataforma()
Dim sql As String, RdoAux As rdoResultset, nNumDoc As Long, ibti As String, RdoAux2 As rdoResultset, bPago As Boolean, sObs As String, nCodReduz As Long

sql = "SELECT DISTINCT numero_guia FROM rodo_uso_plataforma where situacao=2"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        bPago = False
        nNumDoc = !numero_guia
        sql = "select * from debitopago where numdocumento=" & nNumDoc
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            bPago = True
        End If
        RdoAux2.Close
        If bPago Then
            sql = "update rodo_uso_plataforma set situacao=6 where numero_guia=" & nNumDoc
            cn.Execute sql, rdExecDirect
        End If
       .MoveNext
    Loop
   .Close
End With


End Sub

Private Sub Form_Load()
Centraliza Me
sOldData = Mid$(frmMdi.Sbar.Panels(6).Text, 12, 2) & "/" & Mid$(frmMdi.Sbar.Panels(6).Text, 15, 2) & "/" & Right$(frmMdi.Sbar.Panels(6).Text, 4)
sData = sOldData
End Sub

Private Sub Mv_DateClick(ByVal DateClicked As Date)

sData = CStr(Format(Mv.Day, "00") & "/" & Format(Mv.Month, "00") & "/" & Mv.Year)
lblDB.Caption = "Data Base: " & sData
End Sub

Public Sub AtualizaIntegrativa()

Dim nValorTaxa As Double, x As Integer, nSituacao As Integer, dDataProc As Date, sDescImposto As String, RdoAux2 As rdoResultset, y As Integer, nPercTrib As Double
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, sDataVencto As String, nCodTrib As Integer, nValorTributo As Double
Dim sCodReduz As String, sNomeResp As String, sTipoImposto As String, sEndImovel As String, nNumImovel As Integer, sComplImovel As String, sBairroImovel As String
Dim nCodLogr As Long, sEndEntrega As String, nNumEntrega As Integer, sBairroEntrega As String, sComplEntrega As String, sCepEntrega As String, sCidadeEntrega As String
Dim sUFEntrega As String, sNumInsc As String, sValorParc As String, nNumDoc As Long, sBarra As String, sDigitavel2 As String, nValorGuia As Double, nNumGuia As Long
Dim sNumDoc2 As String, sNumDoc3 As String, nFatorVencto As Long, sNumDoc As String, nSid As Long, sDigitavel As String, sNossoNumero As String, sCPF As String, sObs As String
Dim clsImovel As New clsImovel, nCodReduz As Long, sSetor As String, sRG As String, dDataPrimeiraParc As String, nValorTotalHon As Double, RdoAux3 As rdoResultset
Dim nPagina As Integer, nLivro As Integer, xImovel As clsImovel, nQtdeParc As Integer, RdoAux4 As rdoResultset, bCancelado As Boolean, sStatus As String, nTotalRec As Long
Dim Data1 As Date, Data2 As Date, nNumCertidao As Long, dDataInscricao As Date, sNome As String, sCepImovel As String, sCidadeImovel As String, sUFImovel As String
Dim sQuadra As String, sLote As String, nId As Long, sFone As String, sEmail As String
Dim sql As String, RdoAux As rdoResultset, RdoAcordo As rdoResultset, RdoAux5 As rdoResultset
Dim nPos As Integer, sNumProc As String, nNumproc As Long, nAnoproc As Integer, sOcorrencia As String

Data1 = Now - 6
Data2 = Now

If frmMdi.frTeste.Visible = True Then Exit Sub
Set xImovel = New clsImovel
Ocupado
ConectaIntegrativa
sql = "select * from processoreparc wHERE (dataprocesso between '" & Format(Data1, "mm/dd/yyyy") & "' and '" & Format(Data2, "mm/dd/yyyy") & "') ORDER BY numprocesso"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTotalRec = .RowCount
    Do Until .EOF
        nCodReduz = !CODIGORESP
        sNumProc = !NumProcesso
      '  If sNumProc = "12673/2016" Then MsgBox "teste"
        bCancelado = !Cancelado
        nQtdeParc = !qtdeparcela
        dDataProc = !DATAPROCESSO
        nNumproc = Val(Left$(sNumProc, InStr(1, sNumProc, "/", vbBinaryCompare) - 1))
        nAnoproc = Val(Right$(sNumProc, 4))
        
        
        sql = "SELECT origemreparc.numprocesso, debitoparcela.codreduzido, debitoparcela.dataajuiza "
        sql = sql & "FROM debitoparcela INNER JOIN origemreparc ON debitoparcela.codreduzido = origemreparc.codreduzido AND debitoparcela.anoexercicio = origemreparc.anoexercicio AND "
        sql = sql & "debitoparcela.codlancamento = origemreparc.codlancamento AND debitoparcela.seqlancamento = origemreparc.numsequencia AND "
        sql = sql & "debitoparcela.NumParcela = origemreparc.NumParcela And debitoparcela.CODCOMPLEMENTO = origemreparc.CODCOMPLEMENTO "
        sql = sql & "WHERE origemreparc.numprocesso = '" & sNumProc & "' AND debitoparcela.codreduzido = " & nCodReduz
        Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux3!dataajuiza) Then
            RdoAux3.Close
            GoTo Proximo
        Else
            RdoAux3.Close
        End If
        
        sql = "select * from acordos where idacordo=" & nNumproc & " and anoacordo=" & nAnoproc
        Set RdoAux3 = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux3
            If RdoAux3.RowCount = 0 Then
                
            
                'ENDEREÇO DO CONTRIBUINTE
                Select Case nCodReduz
                    Case 1 To 99999
                        sTipoImposto = "REPARCEL."
                        sSetor = "IMOBILIÁRIO"
                        xImovel.CarregaImovel nCodReduz
                        sNumInsc = xImovel.Inscricao
                        sCodReduz = nCodReduz
                        sNomeResp = xImovel.NomePropPrincipal
                        sNome = sNomeResp
                        sQuadra = xImovel.Li_Quadras
                        sLote = xImovel.Li_Lotes
                        xImovel.RetornaEndereco nCodReduz, Imobiliario, Localizacao
                        sEndImovel = xImovel.Endereco
                        nNumImovel = xImovel.Numero
                        sComplImovel = xImovel.Complemento
                        sBairroImovel = xImovel.Bairro
                        sCepImovel = xImovel.Cep
                        sCidadeImovel = xImovel.Cidade
                        sUFImovel = xImovel.UF
                        
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
                
                    Case 100000 To 500000
                        sSetor = "MOBILIÁRIO"
                        sTipoImposto = "REPARCEL."
                        sNomeResp = xImovel.NomePropPrincipal
                        sNumInsc = nCodReduz
                        sCodReduz = nCodReduz
                        sLote = ""
                        sQuadra = ""
                        
                        xImovel.RetornaEndereco nCodReduz, Mobiliario, Localizacao
                        sEndImovel = xImovel.Endereco
                        nNumImovel = xImovel.Numero
                        sComplImovel = xImovel.Complemento
                        sBairroImovel = xImovel.Bairro
                        sCepImovel = xImovel.Cep
                        sCidadeImovel = xImovel.Cidade
                        sUFImovel = xImovel.UF
                        
                        sEndEntrega = xImovel.Ee_NomeLog
                        nNumEntrega = xImovel.Ee_NumImovel
                        sComplEntrega = xImovel.Ee_Complemento
                        sBairroEntrega = xImovel.Bairro
                        sCidadeEntrega = xImovel.Cidade
                        sUFEntrega = xImovel.UF
                        sCepEntrega = xImovel.Ee_Cep
                        sql = "SELECT codigomob, inscestadual,razaosocial, cnpj, cpf From mobiliario WHERE codigomob = " & nCodReduz
                        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux2
                            If .RowCount > 0 Then
                                sNomeResp = !RazaoSocial
                                sCPF = SubNull(!cpf)
                                If Trim(sCPF) = "" Then
                                   sCPF = SubNull(!Cnpj)
                                End If
                             Else
                                sCPF = ""
                             End If
                             sRG = SubNull(!inscestadual)
                            .Close
                        End With
                        
                    Case 500000 To 800000
                        sSetor = "TAXAS DIVERSAS"
                        sTipoImposto = "REPARCEL."
                        sNomeResp = xImovel.NomePropPrincipal
                        sNome = sNomeResp
                        sNumInsc = nCodReduz
                        sCodReduz = nCodReduz
                        sLote = ""
                        sQuadra = ""
                        
                        xImovel.RetornaEndereco nCodReduz, cidadao, cadastrocidadao
                        sEndImovel = xImovel.Endereco
                        nNumImovel = Val(xImovel.Numero)
                        sComplImovel = xImovel.Complemento
                        sBairroImovel = xImovel.Bairro
                        sCepImovel = xImovel.Cep
                        sCidadeImovel = xImovel.Cidade
                        sUFImovel = xImovel.UF
                        
                        sEndEntrega = sEndImovel
                        nNumEntrega = nNumImovel
                        sComplEntrega = sComplImovel
                        sBairroEntrega = sBairroImovel
                        sCidadeEntrega = xImovel.Cidade
                        sUFEntrega = xImovel.UF
                        sCepEntrega = xImovel.Cep
                        
                        sql = "SELECT codcidadao,nomecidadao,cpf,cnpj,rg from cidadao WHERE CODCIDADAO=" & nCodReduz
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
                End Select
                
                
                sql = "SELECT * FROM debitoparcela WHERE (debitoparcela.codreduzido = " & nCodReduz & ") AND "
                sql = sql & "(debitoparcela.numparcela = 1) AND (debitoparcela.numprocesso = '" & sNumProc & "')"
                Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                If .RowCount > 0 Then
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
                            nValorTotalHon = !VALORTRIBUTO * nQtdeParc
                        Else
                            nValorTotalHon = 0
                        End If
                       .Close
                    End With
                   .Close
                End If
                End With
   
                '*** VERIFICA SE O PARCELAMENTO JÁ EXISTE NA TABELA ACORDOS **
                sql = "select * from acordos where idacordo=" & nNumproc & " and anoacordo=" & nAnoproc
                Set RdoAux4 = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                If RdoAux4.RowCount = 0 Then
                    
                    '***** fone e email ****
                    If sSetor = "IMOBILIÁRIO" Then
                        sql = "SELECT proprietario.codcidadao, cidadao.telefone as fone, cidadao.email  FROM proprietario INNER JOIN "
                        sql = sql & "cidadao ON proprietario.codcidadao = cidadao.codcidadao WHERE (proprietario.tipoprop = 'P') AND (proprietario.principal = 1) AND proprietario.codreduzido = " & nCodReduz
                    ElseIf sSetor = "MOBILIÁRIO" Then
                        sql = "SELECT codigomob, fonecontato as fone, emailcontato as email FROM mobiliario WHERE  codigomob = " & nCodReduz
                    Else
                        sql = "SELECT  codcidadao, telefone as fone, email FROM cidadao WHERE codcidadao = " & nCodReduz
                    End If
                    Set RdoAux5 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                    If RdoAux5.RowCount > 0 Then
                        sFone = SubNull(RdoAux5!fone)
                        sEmail = SubNull(RdoAux5!Email)
                    End If
                    RdoAux5.Close
                   '***********************
                    
                    
                    'GRAVA O ACORDO
                    sql = "insert acordos(idacordo,anoacordo,dtparcelamento,setordevedor,iddevedor,nroprocessoadm,crcacordante,nomeacordante,cpfcnpj,rginscrestadual,"
                    sql = sql & "cep,endereco,numero,complemento,bairro, cidade,estado,vlrtotal,qtdparcelas,primeirovencimento,vlrtotalhonorarios,qtdparcelashonorarios,"
                    sql = sql & "vlrparcelahonorarios,dtvenctohonorarios,VlrTotalDespesas, QtdParcelasDespesas, VlrParcelaDespesas, DtVenctoDespesas, DtGeracao,telefone,email) values ("
                    sql = sql & nNumproc & "," & nAnoproc & ",'" & Format(dDataProc, "mm/dd/yyyy") & "','" & sSetor & "',"
                    sql = sql & nCodReduz & ",'" & nNumproc & RetornaDVProcesso(nNumproc) & "/" & nAnoproc & "'," & nCodReduz & ",'" & Mask(sNomeResp) & "','" & sCPF & "','" & Left(sRG, 20) & "','" & sCep & "','" & sEndImovel & "',"
                    sql = sql & nNumImovel & ",'" & Left(sComplImovel, 40) & "','" & sBairroImovel & "','" & Mask(sCidadeEntrega) & "','" & sUFEntrega & "'," & Virg2Ponto(Round((nValorParc * nQtdeParc), 2)) & "," & nQtdeParc & ",'"
                    sql = sql & Format(dDataPrimeiraParc, "mm/dd/yyyy") & "'," & Virg2Ponto(Round(nValorTotalHon, 2)) & "," & IIf(nValorTotalHon = 0, 0, nQtdeParc) & "," & Virg2Ponto(Round((nValorTotalHon / nQtdeParc), 2)) & ","
                    sql = sql & IIf(nValorTotalHon = 0, "Null", "'" & Format(dDataPrimeiraParc, "mm/dd/yyyy") & "'") & "," & "0,0,0," & "Null" & ",'" & Format(Now, "mm/dd/yyyy") & "','" & sFone & "','" & sEmail & "')"
                    cnInt.Execute sql, rdExecDirect
                    
                   'GRAVA NA TABELA ACORDOSTATUS
                    sStatus = IIf(bCancelado = False, "PARCELAMENTO EM DIA", "PARCEL.CANCELADO")
                    sql = "insert acordostatus(idacordo,anoacordo,dtocorrencia,ocorrencia,dtgeracao) values("
                    sql = sql & nNumproc & "," & nAnoproc & ",'" & Format(Now, "mm/dd/yyyy") & "','" & sStatus & "','" & Format(Now, "mm/dd/yyyy") & "')"
                    cnInt.Execute sql, rdExecDirect
                    
                   'GRAVA OS DÉBITOS DO ACORDO
                    sql = "SELECT DISTINCT origemreparc.numprocesso, origemreparc.codreduzido, origemreparc.anoexercicio, origemreparc.codlancamento, origemreparc.numsequencia,"
                    sql = sql & "origemreparc.numparcela, origemreparc.codcomplemento, SUM(debitotributo.valortributo) AS Total, debitoparcela.numerolivro, debitoparcela.paginalivro,debitoparcela.dataajuiza, debitoparcela.numcertidao, debitoparcela.datainscricao, debitoparcela.datavencimento "
                    sql = sql & "FROM origemreparc INNER JOIN debitotributo ON origemreparc.codreduzido = debitotributo.codreduzido AND origemreparc.anoexercicio = debitotributo.anoexercicio AND "
                    sql = sql & "origemreparc.codlancamento = debitotributo.codlancamento AND origemreparc.numsequencia = debitotributo.seqlancamento AND "
                    sql = sql & "origemreparc.numparcela = debitotributo.numparcela AND origemreparc.codcomplemento = debitotributo.codcomplemento INNER JOIN debitoparcela ON origemreparc.codreduzido = debitoparcela.codreduzido AND "
                    sql = sql & "origemreparc.anoexercicio = debitoparcela.anoexercicio AND origemreparc.codlancamento = debitoparcela.codlancamento AND origemreparc.numsequencia = debitoparcela.seqlancamento AND "
                    sql = sql & "origemreparc.NumParcela = debitoparcela.NumParcela And origemreparc.CODCOMPLEMENTO = debitoparcela.CODCOMPLEMENTO GROUP BY origemreparc.numprocesso, origemreparc.codreduzido, origemreparc.anoexercicio, origemreparc.codlancamento, origemreparc.numsequencia,"
                    sql = sql & "origemreparc.NumParcela , origemreparc.CODCOMPLEMENTO, debitoparcela.numerolivro, debitoparcela.paginalivro,debitoparcela.dataajuiza, debitoparcela.numcertidao, debitoparcela.datainscricao, debitoparcela.datavencimento "
                    sql = sql & "HAVING origemreparc.numprocesso = '" & sNumProc & "' AND origemreparc.codreduzido =" & nCodReduz
                    sql = sql & " ORDER BY origemreparc.anoexercicio, origemreparc.numparcela"
                    Set RdoAcordo = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                    With RdoAcordo
                        Do Until .EOF
                            
                            nAno = !AnoExercicio
                            nLanc = !CodLancamento
                            nSeq = !numsequencia
                            nParc = !NumParcela
                            nCompl = !CODCOMPLEMENTO
                            If IsNull(!numcertidao) Then GoTo ProximoCDA
                            nPagina = Val(SubNull(!paginalivro))
                            nLivro = Val(SubNull(!numerolivro))
                            nNumCertidao = Val(SubNull(!numcertidao))
                            dDataInscricao = !datainscricao
                            dDataVencto = !DataVencimento
                           'GRAVA NA TABELA ACORDODEBITO
                            sql = "insert acordodebitos(idacordo,anoacordo,nrolivro,nrofolha,seq,lancamento,exercicio,vlroriginal,vlrcorrecao,vlrjuros,vlrmulta,vlrtotal,nroparcela,complparcela,ajuizado,dtgeracao) values("
                            sql = sql & nNumproc & "," & nAnoproc & "," & nLivro & "," & nPagina & ","
                            sql = sql & RdoAcordo!numsequencia & "," & RdoAcordo!CodLancamento & "," & RdoAcordo!AnoExercicio & "," & Virg2Ponto(Format(RdoAcordo!Total, "#0.##")) & ",0,0,0," & Virg2Ponto(Format(RdoAcordo!Total, "#0.##")) & ","
                            sql = sql & RdoAcordo!NumParcela & "," & RdoAcordo!CODCOMPLEMENTO & "," & IIf(IsNull(!dataajuiza), 0, 1) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
                            cnInt.Execute sql, rdExecDirect
                           
                           
                            '*** GRAVA AS CDA, CDADebito, PARTES E CADASTRO ***
                                
                            sql = "INSERT CDAs(IdDevedor,SetorDevedor,DtInscricao,NroCertidao,NroLivro,NroFolha,DtGeracao) values("
                            sql = sql & nCodReduz & ",'" & sSetor & "','" & Format(dDataInscricao, "mm/dd/yyyy") & "'," & nNumCertidao & ","
                            sql = sql & nLivro & "," & nPagina & ",'" & Format(Now, "mm/dd/yyyy") & "')"
                            cnInt.Execute sql, rdExecDirect
                            
                            sql = "select @@identity as LastKey"
                            Set RdoAux2 = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                            nCDA = RdoAux2!lastkey
                            RdoAux2.Close
                                
                                
                            If nCodReduz < 500000 Then
                                sql = "INSERT Cadastro(IdCDA,SetorDevedor,Crc,Nome,Inscricao,CPFCnpj,RgInscrEstadual,LocalCep,LocalEndereco,LocalNumero,LocalComplemento,"
                                sql = sql & "LocalBairro,LocalCidade,LocalEstado,Quadra,Lote,EntregaCep,EntregaEndereco,EntregaNumero,EntregaComplemento,EntregaBairro,"
                                sql = sql & "EntregaCidade,EntregaEstado,DtGeracao) values("
                                sql = sql & nCDA & ",'" & sSetor & "'," & nCodReduz & ",'" & SubNull(sNome) & "','" & sNumInsc & "','" & sCPF & "','" & sRG & "','"
                                sql = sql & sCepImovel & "','" & sEndImovel & "'," & nNumImovel & ",'" & sComplImovel & "','" & sBairroImovel & "','" & sCidadeImovel & "','" & sUFImovel & "','"
                                sql = sql & sQuadra & "','" & sLote & "','" & sCepEntrega & "','" & sEndEntrega & "'," & nNumEntrega & ",'" & sComplEntrega & "','"
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
                            
                            '***************************************************
                                
                            '**************************************************
ProximoCDA:
                           .MoveNext
                        Loop
                       .Close
                    End With
                    
                End If
                
                RdoAux4.Close
                
            End If
        End With
Proximo:
        DoEvents
       .MoveNext
    Loop
   .Close
End With


'*** PAGAMENTOS *****
sql = "SELECT debitopago.codreduzido, debitopago.anoexercicio, debitopago.codlancamento, debitopago.seqlancamento, debitopago.numparcela, debitopago.codcomplemento, "
sql = sql & "debitopago.seqpag, debitopago.datapagamento, debitopago.datarecebimento, debitopago.valorpago, debitopago.codbanco, debitopago.codagencia,"
sql = sql & "debitopago.restituido, debitopago.numdocumento, debitopago.valorpagoreal, debitopago.intacto, debitopago.valortarifa, debitopago.arquivobanco,"
sql = sql & "debitopago.valordif , debitopago.datapagamentocalc, debitopago.dataintegracao, debitoparcela.numprocesso FROM debitopago INNER JOIN "
sql = sql & "debitoparcela ON debitopago.codreduzido = debitoparcela.codreduzido AND debitopago.anoexercicio = debitoparcela.anoexercicio AND debitopago.codlancamento = debitoparcela.codlancamento AND debitopago.seqlancamento = debitoparcela.seqlancamento AND "
sql = sql & "debitopago.NumParcela = debitoparcela.NumParcela And debitopago.CODCOMPLEMENTO = debitoparcela.CODCOMPLEMENTO "
sql = sql & "WHERE debitopago.datarecebimento>='" & Format(Now - 30, "mm/dd/yyyy") & "' AND (debitopago.codlancamento = 20) AND (debitopago.restituido IS NULL) AND (debitopago.dataintegracao IS NULL)"
'Sql = Sql & "WHERE debitopago.datarecebimento>='08/01/2015' AND (debitopago.codlancamento = 20) AND (debitopago.restituido IS NULL) AND (debitopago.dataintegracao IS NULL)"

Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux3
    Do Until .EOF
        sNumProc = !NumProcesso
        nNumproc = Val(Left$(sNumProc, InStr(1, sNumProc, "/", vbBinaryCompare) - 1))
        nAnoproc = Val(Right$(sNumProc, 4))
        
        'GRAVA NA TABELA ACORDOBAIXAS
        sql = "insert acordobaixas(idAcordo, anoAcordo, DtBaixa, TipoBaixa, NroParcela, VlrOriginal, VlrCorrecao, VlrJuros, VlrMulta, VlrTotal, DtGeracao) values("
        sql = sql & nNumproc & "," & nAnoproc & ",'" & Format(!datarecebimento, "mm/dd/yyyy") & "','PAGAMENTO'," & !NumParcela & "," & Virg2Ponto(CStr(!ValorPagoreal)) & ","
        sql = sql & 0 & "," & 0 & "," & 0 & "," & Virg2Ponto(CStr(!ValorPagoreal)) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
        cnInt.Execute sql, rdExecDirect
        
        sql = "update debitopago set dataintegracao='" & Format(Now, "mm/dd/yyyy") & "' where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and "
        sql = sql & "codlancamento=" & !CodLancamento & " and seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO
        cn.Execute sql, rdExecDirect
       .MoveNext
    Loop
   .Close
End With


'Serasa
sql = "SELECT DISTINCT IDDevedor, CONVERT(char(10), DATEADD(dd, DATEDIFF(DD, 0, DtGeracao), 0), 126) AS dtgeracao from negativacao where dtleitura is null"
Set RdoAux = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        sql = "insert serasa(codigo,dtentrada) values(" & !iddevedor & ",'" & Format(!DtGeracao, "mm/dd/yyyy") & "')"
        cn.Execute sql, rdExecDirect
        sql = "update negativacao set dtleitura='" & Format(Now, "mm/dd/yyyy") & "' where iddevedor=" & !iddevedor & " and datepart(dd,dtgeracao)=" & Day(!DtGeracao) & " and "
        sql = sql & "datepart(mm,dtgeracao)=" & Month(!DtGeracao) & " and datepart(yy,dtgeracao)=" & Year(!DtGeracao)
        cnInt.Execute sql, rdExecDirect
       .MoveNext
    Loop
   .Close
End With


'Parcelamentos Ajuizados
sql = "SELECT AjuizamentoCdas.IdAjuizamentoCdas, AjuizamentoCdas.CodProcesso, AjuizamentoCdas.IdDevedor, AjuizamentoCdas.SetorDevedor, AjuizamentoCdas.NroCertidao, "
sql = sql & "AjuizamentoCdas.NroLivro, AjuizamentoCdas.NroFolha, AjuizamentoCdas.Seq, AjuizamentoCdas.Lancamento, AjuizamentoCdas.ComplParcela,"
sql = sql & "AjuizamentoCdas.Exercicio, AjuizamentoCdas.DtGeracao, AjuizamentoProcessos.DtLeitura, AjuizamentoCdas.NroParcela, AjuizamentoCdas.CodTributo,"
sql = sql & "AjuizamentoCdas.DtVencimento, AjuizamentoCdas.VlrOriginal, AjuizamentoProcessos.DtAjuizamento, AjuizamentoProcessos.Protocolo,"
sql = sql & "AjuizamentoProcessos.IdAjuizamentoProcesso , AjuizamentoProcessos.AnoProtocolo, AjuizamentoProcessos.ProcessoCNJ "
sql = sql & "FROM AjuizamentoCdas INNER JOIN AjuizamentoProcessos ON AjuizamentoCdas.CodProcesso = AjuizamentoProcessos.CodProcesso "
sql = sql & "WHERE AjuizamentoProcessos.DtLeitura IS NULL or  AjuizamentoCdas.DtLeitura IS NULL "
sql = sql & "ORDER BY AjuizamentoCdas.CodProcesso"

Set RdoAux = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
'            Sql = "update debitoparcela set dataajuiza='" & Format(!DtAjuizamento, "mm/dd/yyyy") & "',processocnj='" & !processocnj & "' where "
'            Sql = Sql & "codreduzido=" & !iddevedor & " and anoexercicio=" & !exercicio & " and codlancamento=" & !lancamento & " and "
'            Sql = Sql & "seqlancamento=" & !Seq & " and numparcela=" & !nroparcela

            'voltar o sql acima  para as próximas integrações
            sql = "update debitoparcela set dataajuiza='" & Format(!DtAjuizamento, "mm/dd/yyyy") & "',processocnj='" & !processocnj & "' where "
            sql = sql & "codreduzido=" & Val(!iddevedor) & " and anoexercicio=" & !exercicio & " and "
            sql = sql & "seqlancamento=" & !Seq & " and numparcela=" & !nroparcela & " and numcertidao=" & !NroCertidao
            cn.Execute sql, rdExecDirect
            
            
            sql = " update ajuizamentoprocessos set dtleitura='" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "' where idajuizamentoprocesso=" & !IdAjuizamentoProcesso & " and dtleitura is null"
            cnInt.Execute sql, rdExecDirect
            
            sql = " update ajuizamentocdas set dtleitura='" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "' where idajuizamentocdas=" & !idajuizamentocdas & " and dtleitura is null"
            cnInt.Execute sql, rdExecDirect

DoEvents
       .MoveNext
    Loop
   .Close
End With

'A.R. DIGITAL
sql = "SELECT idajuizamentodespesa,CodProcesso, DtDespesa, ValorDespesa, cnj From AjuizamentoDespesas WHERE (TipoDespesa = 'AR DIGITAL') AND (DtLeitura IS NULL)"
Set RdoAux = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        nId = !idajuizamentodespesa
        sql = " SELECT DISTINCT IdDevedor From AjuizamentoCdas Where CodProcesso=" & !CodProcesso
        Set RdoAux2 = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        nCodReduz = RdoAux2!iddevedor
        RdoAux2.Close
        
        sql = "SELECT MAX(SEQLANCAMENTO) AS SEQMAXIMA FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND CODLANCAMENTO=" & 78 & " AND ANOEXERCICIO=" & Year(!DTDESPESA)
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux2!SEQMAXIMA) Then
           nSeq = 0
        Else
           nSeq = RdoAux2!SEQMAXIMA + 1
        End If
        RdoAux2.Close
        
'        Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
'        Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,dataajuiza,DATADEBASE,PROCESSOCNJ,USUARIO) VALUES(" & nCodReduz & "," & Year(!DTDESPESA) & ","
'        Sql = Sql & 78 & "," & nSeq & "," & 1 & "," & 0 & "," & 3 & ",'" & Format(!DTDESPESA, "mm/dd/yyyy") & "','" & Format(!DTDESPESA, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "','"
'        Sql = Sql & !CNJ & "','INTEGRAÇÃO INTERLIS')"
        sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
        sql = sql & "NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,dataajuiza,DATADEBASE,PROCESSOCNJ,USERID) VALUES(" & nCodReduz & "," & Year(!DTDESPESA) & ","
        sql = sql & 78 & "," & nSeq & "," & 1 & "," & 0 & "," & 3 & ",'" & Format(!DTDESPESA, "mm/dd/yyyy") & "','" & Format(!DTDESPESA, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "','"
        sql = sql & !CNJ & "'," & RetornaUsuarioID(NomeDeLogin) & ")"
        cn.Execute sql, rdExecDirect
        
        nValorTributo = Ponto2Virg(!valordespesa)
        sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
        sql = sql & nCodReduz & "," & Year(!DTDESPESA) & "," & 78 & "," & nSeq & "," & 1 & "," & 0 & "," & 667 & "," & Virg2Ponto(Format(nValorTributo, "#0.00")) & ")"
        cn.Execute sql, rdExecDirect
        
        sql = "update AjuizamentoDespesas set dtleitura='" & Format(Now, "mm/dd/yyyy hh:mm") & "' where idajuizamentodespesa=" & nId
        cnInt.Execute sql, rdExecDirect
        
        DoEvents
       .MoveNext
    Loop
   .Close
End With

'RETORNO PROTESTO
GoTo Fim
sql = "SELECT Protesto_remessa.Codigo, Protesto_remessa.idDevedor, Protesto_Debitos.Exercicio, Protesto_Debitos.Lancamento, Protesto_Debitos.Seq,Protesto_Debitos.nroParcela, Protesto_Debitos.ComplParcela, "
sql = sql & "Protesto_Debitos.dtaVencimento, Protesto_remessa.cod_protesto, Protesto_remessa.nroTitulo,Protesto_remessa.data_remessa , Protesto_Ocorrencia.Ocorrencia,Irregularidade FROM Protesto_remessa INNER JOIN "
sql = sql & "Protesto_Debitos ON Protesto_remessa.cod_protesto = Protesto_Debitos.Cod_protesto LEFT OUTER JOIN Protesto_Ocorrencia ON Protesto_remessa.cod_protesto = Protesto_Ocorrencia.cod_protesto "
sql = sql & "Where (Protesto_remessa.dtLeitura Is Null or Protesto_Ocorrencia.dtLeitura is null) ORDER BY Protesto_remessa.idDevedor,Protesto_remessa.data_remessa, Protesto_Debitos.Exercicio, Protesto_Debitos.Lancamento, Protesto_Debitos.Seq, Protesto_Debitos.nroParcela,Protesto_Debitos.ComplParcela "

'Sql = "SELECT Protesto_remessa.cod_protesto, Protesto_remessa.nroTitulo,Protesto_remessa.data_remessa,idDevedor, Protesto_remessa.nroTitulo,Protesto_remessa.data_remessa,nroCertidao,anoexercicio as exercicio,codlancamento as lancamento,seqlancamento as seq,numparcela as nroparcela,codcomplemento as complparcela,Protesto_Ocorrencia.Ocorrencia,Irregularidade FROM Protesto_remessa "
'Sql = Sql & "INNER JOIN tributacao..debitoparcela ON codreduzido = idDevedor AND numcertidao=nroCertidao LEFT OUTER JOIN Protesto_Ocorrencia ON Protesto_remessa.cod_protesto = Protesto_Ocorrencia.cod_protesto  WHERE Protesto_remessa.dtleitura IS NULL AND nroCertidao IS NOT  null"
Set RdoAux = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If Val(Trim(!iddevedor)) = 15302 Then
     '        MsgBox "teste"
        End If
    
    
        If Trim(SubNull(!Ocorrencia)) = "" Then
            sql = "update debitoparcela set STATUSLANC=39,protesto_nro_titulo=" & Val(Trim(!nroTitulo)) & ",protesto_data_remessa='" & Format(!data_remessa, "mm/dd/yyyy") & "' where codreduzido=" & Val(Trim(!iddevedor)) & " and anoexercicio="
            sql = sql & !exercicio & " and codlancamento=" & !lancamento & " and seqlancamento=" & !Seq & " and numparcela=" & !nroparcela & " and codcomplemento=" & !complparcela & " and statuslanc in (3,38,39,42,42)"
        ElseIf Trim(SubNull(!Ocorrencia)) = "PROTESTADO" Then
            sql = "update debitoparcela set STATUSLANC=38,protesto_nro_titulo=" & Val(Trim(!nroTitulo)) & ",protesto_data_remessa='" & Format(!data_remessa, "mm/dd/yyyy") & "' where codreduzido=" & Val(Trim(!iddevedor)) & " and anoexercicio="
            sql = sql & !exercicio & " and codlancamento=" & !lancamento & " and seqlancamento=" & !Seq & " and numparcela=" & !nroparcela & " and codcomplemento=" & !complparcela & " and statuslanc in (3,38,39,42,43)"
        ElseIf Trim(SubNull(!Ocorrencia)) = "PAGO" Then
            sql = "update debitoparcela set STATUSLANC=41,protesto_nro_titulo=" & Val(Trim(!nroTitulo)) & ",protesto_data_remessa='" & Format(!data_remessa, "mm/dd/yyyy") & "' where codreduzido=" & Val(Trim(!iddevedor)) & " and anoexercicio="
            sql = sql & !exercicio & " and codlancamento=" & !lancamento & " and seqlancamento=" & !Seq & " and numparcela=" & !nroparcela & " and codcomplemento=" & !complparcela & " and statuslanc in (3,38,39,42,43)"
        Else
            sql = "update debitoparcela set STATUSLANC=3,protesto_nro_titulo=" & Val(Trim(!nroTitulo)) & ",protesto_data_remessa='" & Format(!data_remessa, "mm/dd/yyyy") & "' where codreduzido=" & Val(Trim(!iddevedor)) & " and anoexercicio="
            sql = sql & !exercicio & " and codlancamento=" & !lancamento & " and seqlancamento=" & !Seq & " and numparcela=" & !nroparcela & " and codcomplemento=" & !complparcela & " and statuslanc not in (4,38,39)"
            cn.Execute sql, rdExecDirect
        End If
        cn.Execute sql, rdExecDirect
        sql = "update debitoparcela set protesto_nro_titulo=" & Val(Trim(!nroTitulo)) & ",protesto_data_remessa='" & Format(!data_remessa, "mm/dd/yyyy") & "' where codreduzido=" & Val(Trim(!iddevedor)) & " and anoexercicio="
        sql = sql & !exercicio & " and codlancamento=" & !lancamento & " and seqlancamento=" & !Seq & " and numparcela=" & !nroparcela & " and codcomplemento=" & !complparcela
        cn.Execute sql, rdExecDirect
        If Trim(SubNull(!irregularidade)) <> "" Then
            sql = "update debitoparcela set STATUSLANC=39 where codreduzido=" & Val(Trim(!iddevedor)) & " and anoexercicio="
            sql = sql & !exercicio & " and codlancamento=" & !lancamento & " and seqlancamento=" & !Seq & " and numparcela=" & !nroparcela & " and codcomplemento=" & !complparcela & " and statuslanc not in (2,4,41)"
            cn.Execute sql, rdExecDirect
        End If
        
        sql = "update protesto_remessa set dtleitura='" & Format(Now, "mm/dd/yyyy hh:mm") & "' where iddevedor=" & !iddevedor & " and dtleitura is null"
        cnInt.Execute sql, rdExecDirect
        
        sql = "update protesto_ocorrencia set dtleitura='" & Format(Now, "mm/dd/yyyy hh:mm") & "' where cod_protesto=" & !Cod_protesto
        cnInt.Execute sql, rdExecDirect
        
        sql = "update protesto_debitos set dtleitura='" & Format(Now, "mm/dd/yyyy hh:mm") & "' where cod_protesto=" & !Cod_protesto & " and Exercicio=" & !exercicio & " and lancamento=" & !lancamento & " and "
        sql = sql & "seq=" & !Seq & " and nroparcela=" & !nroparcela & " and complparcela=" & !complparcela
        cnInt.Execute sql, rdExecDirect
        
        DoEvents
       .MoveNext
    Loop
   .Close
End With

sql = "update Protesto_Ocorrencia set dtleitura='" & Format(Now, "mm/dd/yyyy hh:mm") & "' where dtleitura is null"
cnInt.Execute sql, rdExecDirect



sql = "select * from CDADebitosNaoProtestados where DtLeitura is null"
Set RdoAux = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        'Sql = "update debitoparcela set STATUSLANC=3,protesto_nro_titulo=Null,protesto_data_remessa=Null where codreduzido=" & Val(Trim(!iddevedor)) & " and anoexercicio="
        sql = "update debitoparcela set STATUSLANC=3 where codreduzido=" & Val(Trim(!iddevedor)) & " and anoexercicio="
        sql = sql & !exercicio & " and codlancamento=" & !lancamento & " and seqlancamento=" & !Seq & " and numparcela=" & Val(SubNull(!nrparcela)) & " and codcomplemento=" & Val(SubNull(!complparcela)) & " and statuslanc=38"
        cn.Execute sql, rdExecDirect
        
        sql = "update CDADebitosNaoProtestados set dtleitura='" & Format(Now, "mm/dd/yyyy hh:mm") & "' where idCDADebitosNaoProtestados=" & !idCDADebitosNaoProtestados
        cnInt.Execute sql, rdExecDirect
        
        
        DoEvents
       .MoveNext
    Loop
   .Close
End With

'Atualiza Ocorrencias
sql = "SELECT * FROM Protesto_Ocorrencia WHERE dtLeitura IS null AND YEAR(data_ocorrencia) = " & Year(Now)
Set RdoAux = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        sql = "SELECT * FROM Protesto_remessa WHERE cod_protesto=" & RdoAux!Cod_protesto
        Set RdoAux2 = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        nCodReduz = RdoAux2!iddevedor
        nNumCertidao = RdoAux2!NroCertidao
        sOcorrencia = Trim(!cod_ocorrencia)
      
        If Trim(SubNull(!Ocorrencia)) = "PROTESTADO" Then
            sql = "update debitoparcela set STATUSLANC=38,protesto_nro_titulo=" & Val(Trim(!nroTitulo)) & ",protesto_data_remessa='" & Format(!data_remessa, "mm/dd/yyyy") & "' where codreduzido=" & nCodReduz & " and numcertidao=" & nNumCertidao & " and statuslanc in (3,38,39)"
        ElseIf Trim(SubNull(!Ocorrencia)) = "PAGO" Then
            sql = "update debitoparcela set STATUSLANC=41,protesto_nro_titulo=" & Val(Trim(!nroTitulo)) & ",protesto_data_remessa='" & Format(!data_remessa, "mm/dd/yyyy") & "' where codreduzido=" & nCodReduz & " and numcertidao=" & nNumCertidao & " and statuslanc in (3,38,39)"

        End If
       
        sql = "update protesto_ocorrencia set dtleitura='" & Format(Now, "mm/dd/yyyy hh:mm") & "' where cod_protesto=" & !Cod_protesto
       ' cnInt.Execute Sql, rdExecDirect
      
       .MoveNext
    Loop
   .Close
End With

Fim:
cnInt.Close
Liberado
End Sub

Private Sub Cancela_ITBIs_Vencidos()

Dim sql As String, RdoAux As rdoResultset, nNumDoc As Long, guid As String, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset
Dim nCodigo As Long, nAno As Integer, nSeq As Integer, nSeq2 As Integer, sItbi As String

sql = "SELECT * FROM debitoparcela WHERE codlancamento=36 AND statuslanc=3 AND datavencimento<'" & Format(Now, "mm/dd/yyyy") & "'"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        nCodigo = !CODREDUZIDO
        nAno = !AnoExercicio
        nSeq = !SeqLancamento
        
        sql = "SELECT numdocumento FROM parceladocumento WHERE codreduzido=" & nCodigo & " AND anoexercicio=" & nAno & " AND codlancamento=36 AND seqlancamento=" & nSeq & " AND numparcela=1 ORDER BY numdocumento"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            nNumDoc = RdoAux2!NumDocumento
        End If
        RdoAux2.Close

       ' Sql = "SELECT guid FROM itbi_main WHERE numero_guia=" & nNumDoc
       ' Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       ' If RdoAux2.RowCount > 0 Then
       '     Sql = "update itbi_main set situacao_itbi=4 where guid='" & guid & "'"
        '    cn.Execute Sql, rdExecDirect
        'End If
        'RdoAux2.Close
        
        sql = "update debitoparcela set statuslanc=5 WHERE codreduzido=" & nCodigo & " AND anoexercicio=" & nAno & " AND codlancamento=36 AND seqlancamento=" & nSeq & " AND numparcela=1 "
        cn.Execute sql, rdExecDirect
        
        sql = "SELECT * FROM itbi_main WHERE numero_guia=" & nNumDoc
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            guid = RdoAux2!guid
            itbi = RdoAux2!itbi_numero & "/" & RdoAux2!itbi_ano
            
            sql = "update itbi_main set situacao_itbi=4 where guid='" & guid & "'"
            cn.Execute sql, rdExecDirect
        End If
        RdoAux2.Close
        
        sObs = "Lançamento do ITBI nº " & itbi & " cancelado devido a falta de pagamento"
        sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA WHERE CODREDUZIDO=" & nCodigo & " AND ANOEXERCICIO=" & nAno
        sql = sql & " AND CODLANCAMENTO=" & 36 & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & 1
        sql = sql & " AND CODCOMPLEMENTO=" & 0
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If IsNull(!maximo) Then
                nSeq2 = 1
            Else
                nSeq2 = !maximo + 1
            End If
           .Close
        End With
        sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & nCodigo & "," & nAno & ","
        sql = sql & 36 & "," & nSeq & "," & 1 & "," & 0 & "," & nSeq2 & ",'" & sObs & "'," & 236 & ",'" & Format(Now, sDataFormat) & "')"
        cn.Execute sql, rdExecDirect
        
       .MoveNext
    Loop
   .Close
End With


End Sub

Private Sub Corrige_Terreno_sem_Endereco()

Dim sql As String, RdoAux As rdoResultset

sql = "SELECT * FROM vwTerrenosComEnderecoImovel ORDER BY codreduzido"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        sql = "update cadimob set ee_tipoend=1 where codreduzido=" & !CODREDUZIDO
        cn.Execute sql, rdExecDirect
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub AtualizaSN()
Dim sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, nCodReduz As Long, sNome As String, sCnpj As String

sql = "delete from empresa_sn"
cn.Execute sql, rdExecDirect

sql = "SELECT DISTINCT codigo,mobiliario.dataencerramento,razaosocial,cnpj FROM periodosn INNER JOIN mobiliario ON codigomob=codigo Where dataencerramento Is Null ORDER BY codigo"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If Not IsNull(!Cnpj) Then
        nCodReduz = !Codigo
        sNome = !RazaoSocial
        sCnpj = !Cnpj
        
        sql = "insert empresa_sn(codigo,nome,cpfcnpj,data) values(" & nCodReduz & ",'" & Mask(sNome) & "','" & sCnpj & "','" & Format(Now, "mm/dd/yyyy") & "')"
        cn.Execute sql, rdExecDirect
        End If
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub Atualiza_Giss_Guia()

Dim sql As String, RdoAux As rdoResultset, nCodigo As Long, nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer
Dim nValorPago As Double, sDataPago As String, nNumDoc As Long, nPos As Integer, sNumProc As String, RdoAux3 As rdoResultset

ConectaEicon
nPos = 0

GoTo PARCELADOS
'PAGOS
sql = "SELECT codigo,ano,lancamento,seq,parcela,complemento,documento,valorpagoreal,datapagamento FROM giss_guia gg INNER JOIN debitopago dp ON "
sql = sql & "gg.codigo=dp.codreduzido and gg.ano=dp.anoexercicio AND gg.lancamento=dp.codlancamento AND gg.seq=dp.seqlancamento AND "
sql = sql & "gg.parcela=dp.numparcela AND gg.complemento=dp.codcomplemento WHERE enviado=0 AND situacao=2 ORDER BY documento"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        nCodigo = !Codigo
        nAno = !ano
        nLanc = !lancamento
        nSeq = !Seq
        nParc = !Parcela
        nCompl = !Complemento
        nValorPago = !ValorPagoreal
        sDataPago = Format(!DataPagamento, "mm/dd/yyyy")
        nNumDoc = !Documento
'baixa ok
        sql = "SELECT * FROM tb_inter_baixa_detalhe WHERE num_documento=" & nNumDoc
        Set RdoAux3 = cnEicon.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux3.RowCount = 0 Then
    
            sql = "insert tb_inter_baixa_detalhe(cod_cliente,cod_banco,num_sequencia,num_documento,linha,timestamp,valor_titulo,valor_pago,data_pagamento,valor_encargos,"
            sql = sql & "descricao_linha_t,descricao_linha_u) values(2177,1,0" & "," & nNumDoc & "," & nPos & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "',"
            sql = sql & Virg2Ponto(CStr(nValorPago)) & "," & Virg2Ponto(CStr(nValorPago)) & ",'" & sDataPago & "',0,'','')"
            cnEicon.Execute sql, rdExecDirect
        End If
        sql = "update giss_guia set enviado=1 where documento=" & nNumDoc
        cn.Execute sql, rdExecDirect
       .MoveNext
    Loop
   .Close
End With

PARCELADOS:
'PARCELADOS
sql = "SELECT codigo,ano,lancamento,seq,parcela,complemento,documento,numprocesso FROM giss_guia g INNER JOIN origemreparc o ON g.codigo=o.codreduzido AND g.ano=o.anoexercicio AND "
sql = sql & "g.lancamento=o.codlancamento AND g.seq=o.numsequencia AND g.parcela=o.numparcela AND g.complemento = o.codcomplemento "
sql = sql & "WHERE enviado=0 AND situacao=4 ORDER BY documento"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        nCodigo = !Codigo
        nAno = !ano
        nLanc = !lancamento
        nSeq = !Seq
        nParc = !Parcela
        nCompl = !Complemento
        nNumDoc = !Documento
        sNumProc = !NumProcesso
        
        
        
        
        
        
        
        
        
        
        
        sql = "insert tb_inter_baixa_detalhe(cod_cliente,cod_banco,num_sequencia,num_documento,linha,timestamp,valor_titulo,valor_pago,data_pagamento,valor_encargos,"
        sql = sql & "descricao_linha_t,descricao_linha_u) values(2177,1,0" & "," & nNumDoc & "," & nPos & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "',"
        sql = sql & Virg2Ponto(CStr(nValorPago)) & "," & Virg2Ponto(CStr(nValorPago)) & ",'" & sDataPago & "',0,'','')"
'        cnEicon.Execute Sql, rdExecDirect
        
        sql = "update giss_guia set enviado=1 where documento=" & nNumDoc
'        cn.Execute Sql, rdExecDirect
       .MoveNext
    Loop
   .Close
End With



cnEicon.Close





End Sub








