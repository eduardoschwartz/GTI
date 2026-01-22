VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmHonorario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de honorários pagos"
   ClientHeight    =   2925
   ClientLeft      =   16005
   ClientTop       =   5640
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2925
   ScaleWidth      =   6795
   Begin Tributacao.jcFrames jcFrames3 
      Height          =   2760
      Left            =   90
      Top             =   90
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   4868
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
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.ListBox lstLog 
         Appearance      =   0  'Flat
         Height          =   1785
         ItemData        =   "frmHonorario.frx":0000
         Left            =   180
         List            =   "frmHonorario.frx":0007
         TabIndex        =   0
         Top             =   810
         Width           =   6225
      End
      Begin MSComCtl2.DTPicker dtDataDe 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   270
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   193462273
         CurrentDate     =   43831
         MaxDate         =   46387
         MinDate         =   43831
      End
      Begin MSComCtl2.DTPicker dtDataAte 
         Height          =   315
         Left            =   3390
         TabIndex        =   2
         Top             =   270
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   193462273
         CurrentDate     =   43831
         MaxDate         =   46387
         MinDate         =   43831
      End
      Begin prjChameleon.chameleonButton btExec 
         Height          =   360
         Left            =   4950
         TabIndex        =   3
         Top             =   225
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   635
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
         MCOL            =   0
         MPTR            =   1
         MICON           =   "frmHonorario.frx":0013
         PICN            =   "frmHonorario.frx":002F
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
         Caption         =   "Data até.:"
         Height          =   195
         Index           =   4
         Left            =   2610
         TabIndex        =   5
         Top             =   330
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Data de.:"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   4
         Top             =   330
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmHonorario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Ficha
    CodTributo As Integer
    F1 As Long 'Principal
    F2 As Long 'Juros e Multa
    F3 As Long 'Principal DA
    F4 As Long 'Juros e Multa DA
    F5 As Long 'Correcao DA
    F6 As Long 'Principal Aj
    F7 As Long 'Juros e Multa Aj
    F8 As Long 'Correcao Aj
End Type
Private Type FichaDetalhe
    Banco As Integer
    Data As Date
    Ficha As Long
    Descricao As String
    Seq As Integer
    Natureza As String
    Vinculo As String
    Perc As Double
    Total As Double
End Type

Private Type Reg
    nCodBanco As Integer
    sNomeBanco As String
    sArquivo As String
    nCodReduz As Long
    nAno As Integer
    nLanc As Integer
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
    nCodTrib As Integer
    sDescTributo As String
    nValorTrib As Double
    sDataVencto As String
    nValorJuros As Double
    nValorMulta As Double
    nValorCorrecao As Double
    nValorTotal As Double
    sDataInscricao As String
    sDataAjuiza As String
    ValorPago As Double
    ValorPr As Double
    ValorJM As Double
    ValorJr As Double
    ValorMl As Double
    ValorCr As Double
    ValorT As Double
    F1 As Long
    F2 As Long
    F3 As Long
    NumDocumento As Long
    sDataRecebimento As String
    nPercP As Double
    nPercJM As Double
    nPercJ As Double
    nPercM As Double
    nPercC As Double
    nPercT As Double
End Type

Private Type Registros
    sDataRecebimento As String
    nNumDocumento As Long
    nCodReduz As Long
    nAno As Integer
    nLanc As Integer
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
    nCodFicha As Long
    sDescFicha As String
    nValor As Double
    nValorPago As Double
    nCodTributo As Integer
    nCodBanco As Integer
    sNomeBanco As String
    sDescTributo As String
    sArquivo As String
    sNatureza As String
    sVinculo As String
    nPerc As Double
    nPercP As Double
    nPercM As Double
    nPercJ As Double
    nPercC As Double
    nValorP As Double
    nValorM As Double
    nValorJ As Double
    nValorC As Double
End Type

Private Type FichaValor
    sDataCredito As String
    nNumDocumento As Long
    nValorFicha As Double
    nFicha As Integer
    nId As Integer
    nPerc As Integer
    nCodTributo As Integer
    nAno As Integer
    nLanc As Integer
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
End Type

Private Type tDoc
    Documento As Long
    ValorPago As Double
    TotalTributos As Double
    DataReceita As String
End Type


Dim aFicha() As Ficha, aCodFicha() As Long, aFichaDetalhe() As FichaDetalhe, aFichas() As FichaDetalhe

Private Sub btExec_Click()
Dim bAnalise As Boolean, bBanco As Boolean, bFicha As Boolean

If dtDataDe.value > dtDataAte.value Then
    MsgBox "Data inicial não pode ser maior que data final.", vbCritical, "Erro"
Else
    Executar_Analise
End If

End Sub

Private Sub Form_Load()
Centraliza Me
dtDataDe.value = Now
dtDataAte.value = Now
lstLog.Clear

End Sub

Private Sub Executar_Analise()

Dim sql As String, RdoAux As rdoResultset, sDataLote As String, qd As New rdoQuery, RdoDeb As rdoResultset, xId As Long, nNumRec As Long, nSomaTmp As Double
Dim nValorPago As Double, nValorTotalPago As Double, aReg() As Reg, x As Integer, nCodTrib As Integer, bDA As Boolean, bAj As Boolean, nValorDif As Double, bFind As Boolean
Dim z As Long, nCodFicha As Integer, bJuros As Boolean, bMulta As Boolean, bCorrecao As Boolean, nCodFichaP As Long, nCodFichaJM As Long, nCodFichaC As Long
Dim nIndex As Long, nIndex2 As Long, aRegDoc() As Registros, RdoAux2 As rdoResultset, nValorTmp As Double, v As Long, nUserId As Integer, aFichaValor() As FichaValor
Dim aDoc() As tDoc, nSomaPago As Double, nSomaFicha As Double, nSomaDif As Double, nSomaPMJC As Double, nPos As Long, nSomaDebitoOriginal As Double, dDataRecebimento As Date
Dim nNumDoc As Long, dDatareceita As Date, bProtestado As Boolean, nFichaProtestado As Long

lstLog.Clear
sDataLote = Format(Now, "ddmmhhmm")
Set qd.ActiveConnection = cn
qd.QueryTimeout = 0
lstLog.AddItem "Iniciando análise às " & Format(Now, "hh:mm:ss")

btExec.Enabled = False

lstLog.AddItem "Preparando tabelas"
Me.Refresh
nUserId = RetornaUsuarioID(NomeDeLogin)
sql = "delete from resumo_honorario where userid=" & RetornaUsuarioID(NomeDeLogin)
cn.Execute sql, rdExecDirect

CarregaTributo

Ocupado

For dDatareceita = dtDataDe.value To dtDataAte.value
    lstLog.AddItem "Iniciando análise do dia " & Format(dDatareceita, "dd/mm/yyyy")
    Me.Refresh
    ReDim aReg(0): ReDim aRegDoc(0): ReDim aDoc(0)
    nCodBanco = 0
    
    sql = "SELECT SUM(debitopago.valorpagoreal) AS soma FROM debitoparcela INNER JOIN debitopago ON debitoparcela.codreduzido = debitopago.codreduzido AND "
    sql = sql & "debitoparcela.anoexercicio = debitopago.anoexercicio AND debitoparcela.codlancamento = debitopago.codlancamento AND debitoparcela.seqlancamento = debitopago.seqlancamento AND "
    sql = sql & "debitoparcela.numparcela = debitopago.numparcela AND debitoparcela.codcomplemento = debitopago.codcomplemento INNER JOIN numdocumento ON debitopago.numdocumento = numdocumento.numdocumento "
    sql = sql & "WHERE RESTITUIDO IS NULL AND DATARECEBIMENTO = '" & Format(dDatareceita, "mm/dd/yyyy") & "' "
    sql = sql & " AND (DEBITOPAGO.CODBANCO not in(90,91,92,93,94,95,96,97,98))"
'    sql = sql & " and (debitopago.codlancamento=41 or debitopago.codlancamento=20)"
    
    Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
    If IsNull(RdoAux!soma) Then
        lstLog.AddItem "Valor total pago no dia R$ 0,00"
        Me.Refresh
        GoTo nextday
    Else
        lstLog.AddItem "Valor total pago no dia R$ " & Format(RdoAux!soma, "#0.00")
        Me.Refresh
        nValorTotalPago = RdoAux!soma
    End If
    RdoAux.Close
    lstLog.ListIndex = (lstLog.ListCount - 1)
    
    sql = "SELECT DISTINCT  debitopago.numdocumento FROM debitopago WHERE RESTITUIDO IS NULL AND DATARECEBIMENTO = '" & Format(dDatareceita, "mm/dd/yyyy") & "' "
    sql = sql & " AND (DEBITOPAGO.CODBANCO not in(90,91,92,93,94,95,96,97,98))  "
    'sql = sql & "and numdocumento=22923930"
    'and (debitopago.codlancamento=41 or debitopago.codlancamento=20)"
    Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        nNumRec = .RowCount
        nPos = 1
        lstLog.AddItem ""
        lstLog.AddItem "Analisando " & nNumRec & " documentos."
        lstLog.AddItem ""
        Me.Refresh
        Do Until .EOF
            If nPos Mod 10 = 0 Then
               lstLog.List(lstLog.ListCount - 1) = "Carregando débitos: " & FormatNumber((nPos * 100) / nNumRec, 2) & "%"
               lstLog.ListIndex = lstLog.ListCount - 1
               lstLog.Refresh
            End If
            nNumDoc = !NumDocumento
            If nNumDoc = 22925460 Then MsgBox "teste"
            nSomaDebitoOriginal = 0
            sql = "SELECT distinct p.codreduzido,p.anoexercicio,p.codlancamento,p.seqlancamento,p.numparcela,p.codcomplemento,datapagamento,p.plano,n.desconto,arquivobanco,b.nomebanco,d.valorpago,datarecebimento,valorpagoreal FROM parceladocumento p INNER JOIN "
            sql = sql & "debitopago g ON p.codreduzido = g.codreduzido AND p.anoexercicio = g.anoexercicio AND p.codlancamento = g.codlancamento AND p.seqlancamento = g.seqlancamento AND p.numparcela = g.numparcela AND "
            sql = sql & "p.codcomplemento = g.codcomplemento AND p.numdocumento = g.numdocumento LEFT OUTER JOIN plano n ON n.codigo=p.plano LEFT OUTER JOIN banco b ON g.codbanco = b.codbanco INNER JOIN numdocumento d ON p.numdocumento = d.numdocumento "
            sql = sql & "where p.numdocumento=" & nNumDoc & " and g.datarecebimento='" & Format(dDatareceita, "mm/dd/yyyy") & "'"
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If RdoAux2.RowCount = 0 Then GoTo NextLanc2
                nValorTotalPago = 0
                dDataRecebimento = !datarecebimento
                Do Until .EOF
                    If Left(!arquivobanco, 3) = "DAF" Then
                        GoTo NextLanc
                    End If
                    DoEvents
                    nValorTotalPago = nValorTotalPago + !ValorPagoreal
                    On Error Resume Next
                    RdoDeb.Close
                    On Error GoTo 0
                    qd.sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
                    qd(0) = !CODREDUZIDO
                    qd(1) = !CODREDUZIDO
                    qd(2) = !AnoExercicio
                    qd(3) = !AnoExercicio
                    qd(4) = !CodLancamento
                    qd(5) = !CodLancamento
                    qd(6) = !SeqLancamento
                    qd(7) = !SeqLancamento
                    qd(8) = !NumParcela
                    qd(9) = !NumParcela
                    qd(10) = !CODCOMPLEMENTO
                    qd(11) = !CODCOMPLEMENTO
                    qd(12) = 0
                    qd(13) = 99
                    qd(14) = Format(!DataPagamento, "mm/dd/yyyy")
                    qd(15) = NomeDeLogin
                    Set RdoDeb = qd.OpenResultset(rdOpenKeyset)
                    With RdoDeb
                        Do Until .EOF
                            bProtestado = False
                            If !NumDocumento = 22967887 Then
'                                MsgBox "teste"
                            End If
                            
                         '   If Not IsNull(!prot_certidao) Then
                         '       bProtestado = True
                         '   End If
                            nCodBanco = !CodBanco
                            If nCodBanco = 0 Then nCodBanco = 1
                            If nCodBanco >= 90 And nCodBanco < 99 Then
                                If nCodBanco = 91 Then
                                    nCodBanco = 1
                                ElseIf nCodBanco = 90 Then nCodBanco = 90
                                ElseIf nCodBanco = 92 Then nCodBanco = 33
                                ElseIf nCodBanco = 93 Then nCodBanco = 237
                                ElseIf nCodBanco = 94 Then nCodBanco = 341
                                ElseIf nCodBanco = 95 Then nCodBanco = 409
                                ElseIf nCodBanco = 96 Then nCodBanco = 151
                                ElseIf nCodBanco = 97 Then nCodBanco = 104
                                ElseIf nCodBanco = 98 Then nCodBanco = 399
                                End If
                            End If
                            If IsNull(RdoAux2!desconto) Then
                                nPercDesconto = 0
                            Else
                                nPercDesconto = RdoAux2!desconto
                            End If
                            
                            If Format(!datarecebimento, "mm/dd/yyyy") = Format(dDatareceita, "mm/dd/yyyy") Then
                                nIndex = UBound(aReg) + 1
                                ReDim Preserve aReg(nIndex)
                                aReg(nIndex).nCodBanco = nCodBanco
                                aReg(nIndex).sNomeBanco = RdoAux2!NomeBanco
                                aReg(nIndex).sArquivo = SubNull(RdoAux2!arquivobanco)
                                aReg(nIndex).nCodReduz = !CODREDUZIDO
                                aReg(nIndex).nAno = !AnoExercicio
                                aReg(nIndex).nLanc = !CodLancamento
                                aReg(nIndex).nSeq = !SeqLancamento
                                aReg(nIndex).nParc = !NumParcela
                                aReg(nIndex).nCompl = !CODCOMPLEMENTO
                                aReg(nIndex).nCodTrib = !CodTributo
                                aReg(nIndex).sDescTributo = !abrevTributo
                                aReg(nIndex).nValorTrib = !VALORTRIBUTO
                                aReg(nIndex).ValorPago = nValorPago
                                aReg(nIndex).NumDocumento = RdoAux!NumDocumento
                                If !datarecebimento <= !DataVencimentoCalc Then
                                    aReg(nIndex).nValorJuros = 0
                                    aReg(nIndex).nValorMulta = 0
                                    aReg(nIndex).nValorCorrecao = 0
                                    aReg(nIndex).nValorTotal = !VALORTRIBUTO
                                Else
                                    aReg(nIndex).nValorJuros = !ValorJuros - (!ValorJuros * nPercDesconto / 100)
                                    aReg(nIndex).nValorMulta = !ValorMulta - (!ValorMulta * nPercDesconto / 100)
                                    aReg(nIndex).nValorCorrecao = !valorcorrecao
                                    aReg(nIndex).nValorTotal = aReg(nIndex).nValorTrib + aReg(nIndex).nValorJuros + aReg(nIndex).nValorMulta + aReg(nIndex).nValorCorrecao
                                End If
                                aReg(nIndex).sDataInscricao = IIf(IsDate(!datainscricao), "S", "N")
                                aReg(nIndex).sDataRecebimento = Format(!datarecebimento, "dd/mm/yyyy")
                                aReg(nIndex).sDataAjuiza = IIf(IsDate(!dataajuiza), "S", "N")
                                
                                nSomaDebitoOriginal = nSomaDebitoOriginal + aReg(nIndex).nValorTrib + aReg(nIndex).nValorJuros + aReg(nIndex).nValorMulta + aReg(nIndex).nValorCorrecao
                            End If
                           .MoveNext
                        Loop
                       .Close
                    End With
NextLanc:
                   .MoveNext
                Loop
               .Close
            End With
NextLanc2:
            'Documento carregado
            'Calcula a proporção de cada tributo
            For nIndex = 1 To UBound(aReg)
                If aReg(nIndex).NumDocumento = nNumDoc Then
                    aReg(nIndex).nPercP = aReg(nIndex).nValorTrib * 100 / nSomaDebitoOriginal
                    aReg(nIndex).nPercJ = aReg(nIndex).nValorJuros * 100 / nSomaDebitoOriginal
                    aReg(nIndex).nPercM = aReg(nIndex).nValorMulta * 100 / nSomaDebitoOriginal
                    aReg(nIndex).nPercC = aReg(nIndex).nValorCorrecao * 100 / nSomaDebitoOriginal
                    
                    aReg(nIndex).nValorTrib = aReg(nIndex).nPercP * nValorTotalPago / 100
                    aReg(nIndex).nValorJuros = aReg(nIndex).nPercJ * nValorTotalPago / 100
                    aReg(nIndex).nValorMulta = aReg(nIndex).nPercM * nValorTotalPago / 100
                    aReg(nIndex).nValorCorrecao = aReg(nIndex).nPercC * nValorTotalPago / 100
                    aReg(nIndex).nValorTotal = aReg(nIndex).nValorTrib + aReg(nIndex).nValorJuros + aReg(nIndex).nValorMulta + aReg(nIndex).nValorCorrecao
                End If
            Next
            
            ReDim Preserve aDoc(UBound(aDoc) + 1)
            aDoc(UBound(aDoc)).Documento = nNumDoc
            aDoc(UBound(aDoc)).DataReceita = dDataRecebimento
            aDoc(UBound(aDoc)).ValorPago = nValorTotalPago
            aDoc(UBound(aDoc)).TotalTributos = nSomaDebitoOriginal
            
NextDoc:
            nPos = nPos + 1
           .MoveNext
        Loop
       .Close
    End With
    lstLog.List(lstLog.ListCount - 1) = "Carregando débitos: 100%"
    lstLog.AddItem ""
    Me.Refresh
    DoEvents
    For nIndex = 1 To UBound(aDoc)
        If nIndex Mod 10 = 0 Then
            lstLog.List(lstLog.ListCount - 1) = "Separando em fichas: " & FormatNumber((nIndex * 100) / UBound(aDoc), 2) & "%"
            lstLog.ListIndex = lstLog.ListCount - 1
            lstLog.Refresh
        End If
        For v = 1 To UBound(aReg)
            If aReg(v).NumDocumento = aDoc(nIndex).Documento Then
                If aReg(v).NumDocumento = 22971004 Then
                      'MsgBox "teste"
                End If
                'carrega as fichas
                nCodTrib = aReg(v).nCodTrib
                z = -1
                z = BinarySearchLong(aCodFicha(), CLng(nCodTrib))
                If z < 1 Then MsgBox "teste"
                bDA = IIf(aReg(v).sDataInscricao = "S", True, False)
                bAj = IIf(aReg(v).sDataAjuiza = "S", True, False)
                If aReg(v).nValorJuros > 0 Then
                    bJuros = True
                Else
                    bJuros = False
                End If
                If aReg(v).nValorMulta > 0 Then
                    bMulta = True
                Else
                    bMulta = False
                End If
                If aReg(v).nValorCorrecao > 0 Then
                    bCorrecao = True
                Else
                    bCorrecao = False
                End If
                
                'Nova Ficha de Honorário Protestado
                'If aReg(v).nLanc = 41 Then
                '    If bProtestado Then
                '        nCodFichaP = nFichaProtestado
                '    End If
                'End If
                
                
                If Not bDA And Not bAj Then
                    If aReg(v).nLanc = 1 And aReg(v).nAno = 2025 And Year(CDate(aReg(v).sDataRecebimento)) = 2024 Then
                        nCodFichaP = 50524
                    Else
                        nCodFichaP = aFicha(z).F1 'Principal
                    End If
                
                
                
                    
                    If bJuros Or bMulta Then
                        nCodFichaJM = aFicha(z).F2 'Juros e Multa normal
                    End If
                    If bCorrecao Then
                        nCodFichaC = aFicha(z).F5 'Correção DA
                    End If
                End If
                If bDA And Not bAj Then
                    nCodFichaP = aFicha(z).F3 'Principal DA
                    If bJuros Or bMulta Then
                        nCodFichaJM = aFicha(z).F4 'Juros e Multa DA
                    End If
                    If bCorrecao Then
                        nCodFichaC = aFicha(z).F5 'Correção DA
                    End If
                End If
                If bDA And bAj Then
                    nCodFichaP = aFicha(z).F6 ' Principal Aj
                    If bJuros Or bMulta Then
                        nCodFichaJM = aFicha(z).F7 'Juros e Multa AJ
                    End If
                    If bCorrecao Then
                        nCodFichaC = aFicha(z).F8 'Correção AJ
                    End If
                End If
                
            '    *** PRINCIPAL ****
           ' If aReg(v).NumDocumento = 17969949 Then MsgBox "teste"
                If nCodFichaP > 0 Then
                    If nCodFichaP = 144 Then
                        'MsgBox "teste"
                    End If
                    sql = "SELECT NATUREZA,VINCULO,PERC,DESCTA FROM FICHACONTABIL WHERE FICHA=" & nCodFichaP
                    Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        If RdoAux2.RowCount = 0 Then
                            MsgBox "Erro!! Ficha " & nCodFichaP & " não cadastrada. (Documento: " & aReg(v).NumDocumento & ")", vbCritical, "Erro"
                        End If
                        Do Until .EOF
                            'If aReg(v).NumDocumento = 19294595 Then MsgBox "teste"

                            ReDim Preserve aRegDoc(UBound(aRegDoc) + 1)
                            aRegDoc(UBound(aRegDoc)).nNumDocumento = aReg(v).NumDocumento
                            aRegDoc(UBound(aRegDoc)).sDataRecebimento = aReg(v).sDataRecebimento
                            aRegDoc(UBound(aRegDoc)).nCodReduz = aReg(v).nCodReduz
                            aRegDoc(UBound(aRegDoc)).nCodBanco = aReg(v).nCodBanco
                            aRegDoc(UBound(aRegDoc)).sNomeBanco = aReg(v).sNomeBanco
                            aRegDoc(UBound(aRegDoc)).nAno = aReg(v).nAno
                            aRegDoc(UBound(aRegDoc)).nLanc = aReg(v).nLanc
                            aRegDoc(UBound(aRegDoc)).nSeq = aReg(v).nSeq
                            aRegDoc(UBound(aRegDoc)).nParc = aReg(v).nParc
                            aRegDoc(UBound(aRegDoc)).nCompl = aReg(v).nCompl
                            aRegDoc(UBound(aRegDoc)).nCodFicha = nCodFichaP
                            aRegDoc(UBound(aRegDoc)).nCodTributo = aReg(v).nCodTrib
                            aRegDoc(UBound(aRegDoc)).sDescFicha = !DESCTA
                            aRegDoc(UBound(aRegDoc)).sDescTributo = aReg(v).sDescTributo
                            aRegDoc(UBound(aRegDoc)).sArquivo = aReg(v).sArquivo
                            aRegDoc(UBound(aRegDoc)).sNatureza = !Natureza
                            aRegDoc(UBound(aRegDoc)).sVinculo = !Vinculo
                            aRegDoc(UBound(aRegDoc)).nPerc = !Perc
                            aRegDoc(UBound(aRegDoc)).nValorP = aReg(v).nValorTrib * !Perc / 100
                            aRegDoc(UBound(aRegDoc)).nValorJ = aReg(v).nValorJuros * !Perc / 100
                            aRegDoc(UBound(aRegDoc)).nValorM = aReg(v).nValorMulta * !Perc / 100

                           .MoveNext
                        Loop
                       .Close
                    End With
                    
                End If
            '   *******************
            '    *** juros e multa ****
                If nCodFichaJM > 0 And (aReg(v).nValorJuros Or aReg(v).nValorMulta > 0) Then
                    sql = "SELECT NATUREZA,VINCULO,PERC,DESCTA FROM FICHACONTABIL WHERE FICHA=" & nCodFichaJM
                    Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        If RdoAux2.RowCount = 0 Then
                            MsgBox "Erro!! Ficha " & nCodFichaJM & " não cadastrada. (Documento: " & aReg(v).NumDocumento & ")", vbCritical, "Erro"
                        End If
                        Do Until .EOF
                            ReDim Preserve aRegDoc(UBound(aRegDoc) + 1)
                            aRegDoc(UBound(aRegDoc)).nNumDocumento = aReg(v).NumDocumento
                            aRegDoc(UBound(aRegDoc)).sDataRecebimento = aReg(v).sDataRecebimento
                            aRegDoc(UBound(aRegDoc)).nCodReduz = aReg(v).nCodReduz
                            aRegDoc(UBound(aRegDoc)).nCodBanco = aReg(v).nCodBanco
                            aRegDoc(UBound(aRegDoc)).sNomeBanco = aReg(v).sNomeBanco
                            aRegDoc(UBound(aRegDoc)).nAno = aReg(v).nAno
                            aRegDoc(UBound(aRegDoc)).nLanc = aReg(v).nLanc
                            aRegDoc(UBound(aRegDoc)).nSeq = aReg(v).nSeq
                            aRegDoc(UBound(aRegDoc)).nParc = aReg(v).nParc
                            aRegDoc(UBound(aRegDoc)).nCompl = aReg(v).nCompl
                            aRegDoc(UBound(aRegDoc)).nCodFicha = nCodFichaJM
                            aRegDoc(UBound(aRegDoc)).nCodTributo = aReg(v).nCodTrib
                            aRegDoc(UBound(aRegDoc)).sDescFicha = !DESCTA
                            aRegDoc(UBound(aRegDoc)).sDescTributo = aReg(v).sDescTributo
                            aRegDoc(UBound(aRegDoc)).sArquivo = aReg(v).sArquivo
                            aRegDoc(UBound(aRegDoc)).sNatureza = !Natureza
                            aRegDoc(UBound(aRegDoc)).sVinculo = !Vinculo
                            aRegDoc(UBound(aRegDoc)).nPerc = !Perc
                            aRegDoc(UBound(aRegDoc)).nValorJ = aReg(v).nValorJuros * !Perc / 100
                            aRegDoc(UBound(aRegDoc)).nValorM = aReg(v).nValorMulta * !Perc / 100
                           .MoveNext
                        Loop
                       .Close
                    End With

                End If
            '   *******************
            '    *** correção ****
                If nCodFichaC > 0 And aReg(v).nValorCorrecao > 0 Then
                    sql = "SELECT NATUREZA,VINCULO,PERC,DESCTA FROM FICHACONTABIL WHERE FICHA=" & nCodFichaC
                    Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        If RdoAux2.RowCount = 0 Then
                            MsgBox "Erro!! Ficha " & nCodFichaC & " não cadastrada. (Documento: " & aReg(v).NumDocumento & ")", vbCritical, "Erro"
                        End If
                        Do Until .EOF
                            ReDim Preserve aRegDoc(UBound(aRegDoc) + 1)
                            aRegDoc(UBound(aRegDoc)).nNumDocumento = aReg(v).NumDocumento
                            aRegDoc(UBound(aRegDoc)).sDataRecebimento = aReg(v).sDataRecebimento
                            aRegDoc(UBound(aRegDoc)).nCodReduz = aReg(v).nCodReduz
                            aRegDoc(UBound(aRegDoc)).nCodBanco = aReg(v).nCodBanco
                            aRegDoc(UBound(aRegDoc)).sNomeBanco = aReg(v).sNomeBanco
                            aRegDoc(UBound(aRegDoc)).nAno = aReg(v).nAno
                            aRegDoc(UBound(aRegDoc)).nLanc = aReg(v).nLanc
                            aRegDoc(UBound(aRegDoc)).nSeq = aReg(v).nSeq
                            aRegDoc(UBound(aRegDoc)).nParc = aReg(v).nParc
                            aRegDoc(UBound(aRegDoc)).nCompl = aReg(v).nCompl
                            aRegDoc(UBound(aRegDoc)).nCodFicha = nCodFichaC
                            aRegDoc(UBound(aRegDoc)).nCodTributo = aReg(v).nCodTrib
                            aRegDoc(UBound(aRegDoc)).sDescFicha = !DESCTA
                            aRegDoc(UBound(aRegDoc)).sDescTributo = aReg(v).sDescTributo
                            aRegDoc(UBound(aRegDoc)).sArquivo = aReg(v).sArquivo
                            aRegDoc(UBound(aRegDoc)).sNatureza = !Natureza
                            aRegDoc(UBound(aRegDoc)).sVinculo = !Vinculo
                            aRegDoc(UBound(aRegDoc)).nPerc = !Perc
                            aRegDoc(UBound(aRegDoc)).nValorC = aReg(v).nValorCorrecao * !Perc / 100
                           .MoveNext
                        Loop
                       .Close
                    End With
                    
                End If
                
            End If
            
        Next
    Next
    
    lstLog.List(lstLog.ListCount - 1) = "Separando em fichas: 100%"
    lstLog.AddItem ""
    Me.Refresh
    
    sql = "select count(*) as contador from resumo_pagto_banco_ficha where userid=" & nUserId
    Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
    If IsNull(RdoAux3!contador) Then
        xId = 1
    Else
        xId = RdoAux3!contador + 1
    End If
    
    For v = 1 To UBound(aRegDoc)
        aRegDoc(v).nValor = aRegDoc(v).nValorP + aRegDoc(v).nValorM + aRegDoc(v).nValorJ + aRegDoc(v).nValorC
       ' If aRegDoc(v).nNumDocumento = 22961011 And aRegDoc(v).nLanc > 5 Then MsgBox "teste"
        If v Mod 20 = 0 Then
            lstLog.List(lstLog.ListCount - 1) = "Gravando análise: " & FormatNumber((v * 100) / UBound(aRegDoc), 2) & "%"
            lstLog.ListIndex = lstLog.ListCount - 1
            lstLog.Refresh
        End If
        If Format(aRegDoc(v).sDataRecebimento, "dd/mm/yyyy") = Format(dDatareceita, "dd/mm/yyyy") Then
        If aRegDoc(v).nCodBanco = 0 Then aRegDoc(v).nCodBanco = 1
           If aRegDoc(v).nCodFicha = 144 Or aRegDoc(v).nCodFicha = 277 Then
               sql = "insert resumo_honorario(userid,datacredito,documento,codigo,ano,lanc,seq,parc,compl,codtributo,desctributo,descficha,ficha,arquivo,natureza,vinculo,perc,valor,codbanco,id,valorp,valorj,valorm,"
               sql = sql & "valorc,nomebanco) values(" & nUserId & ",'" & Format(aRegDoc(v).sDataRecebimento, "mm/dd/yyyy") & "'," & aRegDoc(v).nNumDocumento & "," & aRegDoc(v).nCodReduz & "," & aRegDoc(v).nAno & ","
               sql = sql & aRegDoc(v).nLanc & "," & aRegDoc(v).nSeq & "," & aRegDoc(v).nParc & "," & aRegDoc(v).nCompl & "," & aRegDoc(v).nCodTributo & ",'" & aRegDoc(v).sDescTributo & "','" & aRegDoc(v).sDescFicha & "',"
               sql = sql & aRegDoc(v).nCodFicha & ",'" & aRegDoc(v).sArquivo & "','" & aRegDoc(v).sNatureza & "','" & aRegDoc(v).sVinculo & "'," & aRegDoc(v).nPerc & "," & Virg2Ponto(CStr(aRegDoc(v).nValor)) & ","
               sql = sql & aRegDoc(v).nCodBanco & "," & xId & "," & Virg2Ponto(CStr(aRegDoc(v).nValorP)) & "," & Virg2Ponto(CStr(aRegDoc(v).nValorJ)) & "," & Virg2Ponto(CStr(aRegDoc(v).nValorM)) & ","
               sql = sql & Virg2Ponto(CStr(aRegDoc(v).nValorC)) & ",'" & Mask(aRegDoc(v).sNomeBanco) & "')"
               cn.Execute sql, rdExecDirect
           End If
        End If
        xId = xId + 1
    Next
    lstLog.List(lstLog.ListCount - 1) = "Gravando análise: 100%"
    lstLog.AddItem ""
    lstLog.AddItem "Imprimindo relatório(s)"
    lstLog.AddItem ""
    lstLog.Refresh
    
    
nextday:
    lstLog.AddItem ""
Next 'muda de data
    
lstLog.AddItem "Análise encerrada às " & Format(Now, "hh:mm:ss")
lstLog.ListIndex = lstLog.ListCount - 1

Analise:
If bAnalise Then
    
    
    ReDim aRegDoc(0)
    sql = "select * from resumo_pagto_banco_ficha where userid=" & nUserId
    Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            bFind = False
            For v = 1 To UBound(aRegDoc)
                If aRegDoc(v).sNatureza = !Natureza And aRegDoc(v).sVinculo = !Vinculo And aRegDoc(v).nCodBanco = !CodBanco Then
                    bFind = True
                    Exit For
                End If
            Next
            If Not bFind Then
                ReDim Preserve aRegDoc(UBound(aRegDoc) + 1)
                aRegDoc(UBound(aRegDoc)).sNatureza = !Natureza
                aRegDoc(UBound(aRegDoc)).sVinculo = !Vinculo
                aRegDoc(UBound(aRegDoc)).sDataRecebimento = Format(!DataCredito, "yyyymmdd")
                aRegDoc(UBound(aRegDoc)).nCodBanco = !CodBanco
                aRegDoc(UBound(aRegDoc)).nValor = !valor
            Else
                aRegDoc(v).nValor = aRegDoc(v).nValor + !valor
            End If
           .MoveNext
        Loop
       .Close
    End With
    
    If cmbBanco.ListIndex = 0 Then
       Open sPathBin & "\REC" & Format(Day(Now), "00") & Format(Month(Now), "00") For Output Shared As #1
    Else
       Open sPathBin & "\REC" & Format(Day(Now), "00") & Format(Month(Now), "00") & Format(cmbBanco.ItemData(cmbBanco.ListIndex), "000") For Output Shared As #1
    End If
    
    For x = 1 To UBound(aRegDoc)
        With aRegDoc(x)
            ax = FillSpace(.sNatureza, 20) & FillSpace(.sVinculo, 20) & .sDataRecebimento & Format(.nCodBanco, "0000") & "00000000" & Format(RemovePonto(Replace(CStr(FormatNumber(.nValor, 2)), ",", "")), "0000000000000") & "0000000000" & sDataLote
            Print #1, ax
        End With
    Next
    Close #1
    
End If

frmReport.ShowReport3 "Honorario_Pago", frmMdi.HWND, Me.HWND

Fim:
sql = "delete from resumo_honorario where userid=" & nUserId
cn.Execute sql, rdExecDirect


Liberado
btExec.Enabled = True

End Sub

Private Sub CarregaTributo()
Dim sql As String, RdoAux As rdoResultset

ReDim aFicha(0): ReDim aCodFicha(0)

sql = "SELECT codtributo, ficha, fichajrmulta, fichadivida, fichadajrmul, fichadaenca, fichaajuiza, fichaajjrmul, fichaajenca FROM tributo order by codtributo"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aFicha(UBound(aFicha) + 1)
        ReDim Preserve aCodFicha(UBound(aCodFicha) + 1)
        aCodFicha(UBound(aCodFicha)) = !CodTributo
        aFicha(UBound(aFicha)).CodTributo = !CodTributo
        aFicha(UBound(aFicha)).F1 = !Ficha
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Function FillSpace(sPalavra As String, nTamanho As Integer) As String
Dim sTmp As String

If Len(sPalavra) > nTamanho Then sPalavra = Left(sPalavra, nTamanho)
sTmp = sPalavra & Space(nTamanho - Len(sPalavra))
FillSpace = sTmp

End Function

