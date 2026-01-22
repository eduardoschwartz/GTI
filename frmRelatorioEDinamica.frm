VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmRelatorioEDinamica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatórios Estrutura Dinâmica"
   ClientHeight    =   2625
   ClientLeft      =   7170
   ClientTop       =   6540
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2625
   ScaleWidth      =   4965
   Begin VB.ComboBox cmbTipo 
      Height          =   315
      ItemData        =   "frmRelatorioEDinamica.frx":0000
      Left            =   135
      List            =   "frmRelatorioEDinamica.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   405
      Width           =   4515
   End
   Begin VB.ComboBox cmbOperador 
      Height          =   315
      ItemData        =   "frmRelatorioEDinamica.frx":0004
      Left            =   135
      List            =   "frmRelatorioEDinamica.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1080
      Width           =   4515
   End
   Begin esMaskEdit.esMaskedEdit mskDataIni 
      Height          =   285
      Left            =   1215
      TabIndex        =   2
      Top             =   1620
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   503
      MouseIcon       =   "frmRelatorioEDinamica.frx":0008
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
      Text            =   "__/__/____"
      HideSelection   =   -1  'True
   End
   Begin esMaskEdit.esMaskedEdit mskDataFim 
      Height          =   285
      Left            =   3615
      TabIndex        =   3
      Top             =   1635
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   503
      MouseIcon       =   "frmRelatorioEDinamica.frx":0024
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
      Text            =   "__/__/____"
      HideSelection   =   -1  'True
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Default         =   -1  'True
      Height          =   360
      Left            =   3555
      TabIndex        =   4
      ToolTipText     =   "Imprimir o relatório"
      Top             =   2160
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   635
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
      MICON           =   "frmRelatorioEDinamica.frx":0040
      PICN            =   "frmRelatorioEDinamica.frx":005C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblEquipe 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de relatório"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   8
      Top             =   135
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Fim.....:"
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   2595
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Início..:"
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   1
      Left            =   180
      TabIndex        =   6
      Top             =   1665
      Width           =   1035
   End
   Begin VB.Label lblEquipe 
      BackStyle       =   0  'Transparent
      Caption         =   "Operador"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   5
      Top             =   855
      Width           =   870
   End
End
Attribute VB_Name = "frmRelatorioEDinamica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type tCodigo
    nCodigo As Long
    sNome As String
End Type

Private Type tBoleto
    Documento As Long
    Codigo As Long
    Nome As String
    Emissor As String
    ValorGuia As Double
    DataDoc As Date
    ValorPago As Double
    DataPago As Date
    UserId As Integer
End Type

Private Type tParcelamento
    NumProcesso As String
    AnoProc As Integer
    NumProc As Long
    Seq As Integer
    Codigo As Long
    Nome As String
    User_Id As Integer
    User_Name As String
    Data As Date
    QtdeParc As Integer
    QtdePago As Integer
    ValorGerado As Double
    ValorPago As Double
End Type

Dim sData1 As String, sData2 As String, sOperador As String, bTodos As Boolean

Private Sub cmdPrint_Click()

sData1 = mskDataIni.Text
sData2 = mskDataFim.Text

If Not IsDate(sData1) Or Not IsDate(sData2) Then
    MsgBox "Data inicial e/ou data final inválida", vbCritical, "Erro"
    Exit Sub
End If

If CDate(sData1) > CDate(sData2) Then
    MsgBox "Data inicial maior que data final", vbCritical, "Erro"
    Exit Sub
End If

sOperador = cmbOperador.Text
bTodos = cmbOperador.ListIndex = 0

If MsgBox("Deseja gerar o relatório?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub

Select Case cmbTipo.ListIndex
    Case 0
        RelBoletoAvista
    Case 1
        RelParcelamento
End Select


End Sub

Private Sub Form_Load()
Dim RdoAux As rdoResultset, Sql As String

Centraliza Me

cmbTipo.AddItem ("01-Lista dos Boletos à vista emitidos")
cmbTipo.AddItem ("02-Lista dos Parcelamentos feitos no GTI")
cmbTipo.ListIndex = 0

cmbOperador.AddItem ("(Todos)")
cmbOperador.ItemData(cmbOperador.NewIndex) = 999
cmbOperador.AddItem ("SATYLA.SOUZA")
cmbOperador.ItemData(cmbOperador.NewIndex) = 666
cmbOperador.AddItem ("ARIANA.SILVA")
cmbOperador.ItemData(cmbOperador.NewIndex) = 664
cmbOperador.AddItem ("LUANA.OLIVEIRA")
cmbOperador.ItemData(cmbOperador.NewIndex) = 665
cmbOperador.AddItem ("JAINE.JESUS")
cmbOperador.ItemData(cmbOperador.NewIndex) = 660
cmbOperador.ListIndex = 0

End Sub

Private Sub mskDataFim_Click()
mskDataFim.SelStart = 0
mskDataFim.SelLength = Len(mskDataFim.Text)

End Sub

Private Sub mskDataFim_GotFocus()
mskDataFim.SelStart = 0
mskDataFim.SelLength = Len(mskDataFim.Text)

End Sub

Private Sub mskDataIni_Click()
mskDataIni.SetFocus
mskDataIni.SelStart = 0
mskDataIni.SelLength = Len(mskDataIni.Text)

End Sub

Private Sub mskDataIni_GotFocus()
mskDataIni.SelStart = 0
mskDataIni.SelLength = Len(mskDataIni.Text)
End Sub

Private Sub RelBoletoAvistaOld()
Dim RdoAux As rdoResultset, Sql As String, aCodigo() As tCodigo, x As Integer, nCodigo As Long
Dim RdoAux2 As rdoResultset, sNome As String, aBoleto() As tBoleto, y As Integer, nDoc As Long
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer

ReDim aCodigo(0): ReDim aBoleto(0)
Ocupado

Sql = "select distinct(codreduzido) from vwGuias_Emitidas_EDinamica where datadocumento>='" & Format(sData1, "mm/dd/yyyy") & "' and datadocumento<='" & Format(sData2, "mm/dd/yyyy") & "' "
If Not bTodos Then
    Sql = Sql & " and emissor='" & sOperador & "'"
End If
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aCodigo(UBound(aCodigo) + 1)
        aCodigo(UBound(aCodigo)).nCodigo = !CODREDUZIDO
       .MoveNext
    Loop
   .Close
End With

For x = 1 To UBound(aCodigo)
    nCodigo = aCodigo(x).nCodigo
    
    If nCodigo < 100000 Then
        Sql = "SELECT NOMECIDADAO AS NOME FROM vwFULLIMOVEL WHERE CODREDUZIDO=" & nCodigo
    ElseIf nCodigo >= 100000 And nCodigo < 500000 Then
        Sql = "SELECT RAZAOSOCIAL AS NOME FROM MOBILIARIO WHERE CODIGOMOB=" & nCodigo
    Else
        Sql = "SELECT NOMECIDADAO AS NOME FROM CIDADAO WHERE CODCIDADAO=" & nCodigo
    End If
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    sNome = RdoAux2!Nome
    RdoAux2.Close
    aCodigo(x).sNome = sNome
Next

Sql = "delete from Relatorio_EDinamica1 where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect


Sql = "select * from vwGuias_Emitidas_EDinamica where datadocumento>='" & Format(sData1, "mm/dd/yyyy") & "' and datadocumento<='" & Format(sData2, "mm/dd/yyyy") & "' "
If Not bTodos Then
    Sql = Sql & " and userid=" & cmbOperador.ItemData(cmbOperador.ListIndex)
End If
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        For x = 1 To UBound(aCodigo)
            If aCodigo(x).nCodigo = RdoAux!CODREDUZIDO Then
                sNome = aCodigo(x).sNome
                Exit For
            End If
        Next
    
        ReDim Preserve aBoleto(UBound(aBoleto) + 1)
        y = UBound(aBoleto)
        aBoleto(y).Documento = !NumDocumento
        aBoleto(y).Emissor = !Emissor
        aBoleto(y).Codigo = !CODREDUZIDO
        aBoleto(y).Nome = sNome
        If IsNull(!ValorGuia) Then
            aBoleto(y).ValorGuia = 0
        Else
            aBoleto(y).ValorGuia = !ValorGuia
        End If
        aBoleto(y).DataDoc = !Datadocumento
        If IsNull(!ValorPago) Then
            aBoleto(y).ValorPago = 0
            aBoleto(y).DataPago = Date
        Else
            aBoleto(y).ValorPago = !ValorPago
            aBoleto(y).DataPago = !DataPagamento
        End If
    
       .MoveNext
    Loop
   .Close
End With

For x = 1 To UBound(aBoleto)
    With aBoleto(x)
        If .ValorPago = 0 Then
            Sql = "select * from parceladocumento where numdocumento=" & .Documento
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset)
            With RdoAux
                If .RowCount > 0 Then
                    nAno = !AnoExercicio
                    nLanc = !CodLancamento
                    nSeq = !SeqLancamento
                    nParc = !NumParcela
                    nCompl = !CODCOMPLEMENTO
                    Sql = "select numdocumento,datarecebimento from debitopago where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & nAno & " and codlancamento=" & nLanc & " and "
                    Sql = Sql & "seqlancamento=" & nSeq & " and numparcela=" & nParc & " and codcomplemento=" & nCompl
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    If RdoAux2.RowCount > 0 Then
                        nDoc = RdoAux2!NumDocumento
                        aBoleto(x).DataPago = RdoAux2!datarecebimento
                        RdoAux2.Close
                        Sql = "select sum(valorpagoreal) as soma from debitopago where numdocumento=" & nDoc
                        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        aBoleto(x).ValorPago = RdoAux2!soma
                    End If
                    RdoAux2.Close
                End If
               .Close
            End With
        End If
    End With
Next

For x = 1 To UBound(aBoleto)
    With aBoleto(x)
        Sql = "insert Relatorio_EDinamica1(usuario,documento,emissor,codigo,nome,valorguia,datadoc,valorpago,datapago) values('"
        Sql = Sql & NomeDeLogin & "'," & .Documento & ",'" & .Emissor & "'," & .Codigo & ",'" & .Nome & "'," & Virg2Ponto(CStr(.ValorGuia)) & ",'"
        Sql = Sql & Format(.DataDoc, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(.ValorPago)) & ",'" & Format(.DataPago, "mm/dd/yyyy") & "')"
        cn.Execute Sql, rdExecDirect
    End With
Next

Sql = "update Relatorio_EDinamica1 set datapago=null where usuario='" & NomeDeLogin & "' and valorpago=0"
cn.Execute Sql, rdExecDirect


frmReport.ShowReport3 "ESTRUTURA_DINAMICA1", frmMdi.HWND, Me.HWND

Liberado

Sql = "delete from Relatorio_EDinamica1 where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect


End Sub

Private Sub RelParcelamento()
Dim RdoAux As rdoResultset, Sql As String, aCodigo() As tCodigo, x As Integer, nCodigo As Long
Dim RdoAux2 As rdoResultset, sNome As String, y As Integer, aParc() As tParcelamento, nUser As Integer
Dim nValorGerado As Double, nValorPago As Double, nQtdePago As Integer, sNumProc As String

ReDim aParc(0)
nUser = cmbOperador.ItemData(cmbOperador.ListIndex)
nValorGerado = 0: nValorPago = 0: nQtdePago = 0

Ocupado


Sql = "SELECT numprocesso,numproc,anoproc,qtdeparcela,datareparc,userid,codigoresp FROM processoreparc "
Sql = Sql & "where DATEADD(dd, 0, DATEDIFF(dd, 0, datareparc)) between '" & Format(mskDataIni.Text, "mm/dd/yyyy") & "' and '" & Format(mskDataFim.Text, "mm/dd/yyyy") & "' and "
If nUser = 999 Then
    Sql = Sql & " userid BETWEEN 660 AND 668"
Else
    Sql = Sql & " userid=" & nUser
End If

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aParc(UBound(aParc) + 1)
        y = UBound(aParc)
        sNumProc = !NumProcesso
        aParc(y).NumProcesso = Format(!NumProc, "00000") & "-" & RetornaDVProcesso(!NumProc) & "/" & !AnoProc
        aParc(y).AnoProc = !AnoProc
        aParc(y).NumProc = !NumProc
        aParc(y).QtdeParc = !qtdeparcela
        aParc(y).Data = !datareparc
        aParc(y).Codigo = !CODIGORESP
        aParc(y).Nome = RetornaNome(!CODIGORESP)
        aParc(y).User_Id = !UserId
        aParc(y).User_Name = RetornaUsuarioFullName3(!UserId)
        aParc(y).ValorGerado = nValorGerado
        aParc(y).ValorPago = nValorPago
        aParc(y).QtdePago = nQtdePago
       .MoveNext
    Loop
   .Close
End With

Sql = "delete from relatorio_parcelamentoed where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

For y = 1 To UBound(aParc)
    With aParc(y)
        sNumProc = .NumProc & "/" & .AnoProc
        Sql = "SELECT DISTINCT numsequencia FROM destinoreparc WHERE numprocesso='" & sNumProc & "'"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        aParc(y).Seq = RdoAux!numsequencia
        RdoAux.Close
        Sql = "SELECT COUNT(*) AS contador FROM debitopago WHERE codreduzido=" & .Codigo & " AND CODLANCAMENTO=20 AND seqlancamento=" & .Seq & " AND "
        Sql = Sql & "DATAPAGAMENTO BETWEEN '" & Format(mskDataIni.Text, "mm/dd/yyyy") & "' and '" & Format(mskDataFim.Text, "mm/dd/yyyy") & "'"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux!contador) Then
            aParc(y).QtdePago = 0
        Else
            aParc(y).QtdePago = RdoAux!contador
        End If
        RdoAux.Close
        
        Sql = "SELECT sum(valorpagoreal) AS contador FROM debitopago WHERE codreduzido=" & .Codigo & " AND CODLANCAMENTO=20 AND seqlancamento=" & .Seq & " AND "
        Sql = Sql & "DATAPAGAMENTO BETWEEN '" & Format(mskDataIni.Text, "mm/dd/yyyy") & "' and '" & Format(mskDataFim.Text, "mm/dd/yyyy") & "'"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux!contador) Then
            aParc(y).ValorPago = 0
        Else
            aParc(y).ValorPago = RdoAux!contador
        End If
        RdoAux.Close
        
        Sql = "SELECT SUM(valortributo) AS soma FROM debitotributo WHERE codreduzido=" & .Codigo & " AND CODLANCAMENTO=20 AND seqlancamento=" & .Seq
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux!soma) Then
            aParc(y).ValorGerado = 0
        Else
            aParc(y).ValorGerado = RdoAux!soma
        End If
        RdoAux.Close
        
        
        Sql = "insert relatorio_parcelamentoed(usuario,anoproc,numproc,numprocesso,data,qtdeparcelado,qtdepago,valorparcelado,valorpago,codigo,nome,userid,atendente,seq) values('"
        Sql = Sql & NomeDeLogin & "'," & .AnoProc & "," & .NumProc & ",'" & .NumProcesso & "','" & Format(.Data, "mm/dd/yyyy") & "'," & .QtdeParc & "," & .QtdePago & ","
        Sql = Sql & Virg2Ponto(CStr(.ValorGerado)) & "," & Virg2Ponto(CStr(.ValorPago)) & "," & .Codigo & ",'" & Mask(.Nome) & "'," & .User_Id & ",'" & .User_Name & "'," & .Seq & ")"
        cn.Execute Sql, rdExecDirect
    End With
Next

Liberado
frmReport.ShowReport3 "ESTRUTURA_DINAMICA2", frmMdi.HWND, Me.HWND

Sql = "delete from relatorio_parcelamentoed where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect


End Sub

Private Sub RelBoletoAvista()
Dim RdoAux As rdoResultset, Sql As String, aCodigo() As tCodigo, x As Integer, nCodigo As Long
Dim RdoAux2 As rdoResultset, sNome As String, aBoleto() As tBoleto, y As Integer, nDoc As Long, nUser As Integer
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, bFind As Boolean, sUser As String

Sql = "delete from Relatorio_EDinamica1 where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

ReDim aCodigo(0): ReDim aBoleto(0)
Ocupado

nUser = cmbOperador.ItemData(cmbOperador.ListIndex)
sUser = cmbOperador.Text & " (DAM%"

'JAINE.JESUS (DAM.%
'LUANA.OLIVEIRA (DAM.%
'ARIANA.SILVA (DAM.%
'SATYLA.SOUZA (DAM.%
'

Sql = "SELECT DISTINCT NUMDOCUMENTO.numdocumento,datadocumento,valorguia,emissor,userid,valorpagoreal,datapagamento,parceladocumento.codreduzido,nomecidadao AS NOME, "
Sql = Sql & "debitopago.codreduzido ,debitopago.anoexercicio ,debitopago.codlancamento ,debitopago.seqlancamento ,debitopago.numparcela ,debitopago.codcomplemento "
Sql = Sql & "From NumDocumento INNER JOIN parceladocumento ON numdocumento.numdocumento=parceladocumento.numdocumento "
Sql = Sql & "INNER JOIN vwFULLIMOVEL ON parceladocumento.codreduzido=vwFULLIMOVEL.codreduzido LEFT OUTER JOIN debitopago ON numdocumento.numdocumento = debitopago.numdocumento "
Sql = Sql & "WHERE datadocumento between '" & Format(mskDataIni.Text, "mm/dd/yyyy") & "' and '" & Format(mskDataFim.Text, "mm/dd/yyyy") & "' and parceladocumento.CODREDUZIDO<50000 and "
If bTodos Then
    Sql = Sql & "(EMISSOR LIKE 'JAINE.JESUS (DAM%' OR EMISSOR LIKE 'LUANA.OLIVEIRA (DAM%' OR EMISSOR LIKE 'ARIANA.SILVA (DAM%' OR EMISSOR LIKE 'SATYLA.SOUZA (DAM%')"
    Sql = Sql & "order by numdocumento"
Else
    nUser = cmbOperador.ItemData(cmbOperador.ListIndex)
    Sql = Sql & "emissor like '" & sUser & "'"
End If
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF

        For x = 1 To UBound(aBoleto)
            bFind = False
            If aBoleto(x).Documento = !NumDocumento Then
                bFind = True
                Exit For
            End If
        Next
        If Not bFind Then
          ReDim Preserve aBoleto(UBound(aBoleto) + 1)
          y = UBound(aBoleto)
           aBoleto(y).Documento = !NumDocumento
           aBoleto(y).Emissor = !Emissor
            aBoleto(y).Codigo = !CODREDUZIDO
            aBoleto(y).Nome = sNome
            If IsNull(!ValorGuia) Then
                aBoleto(y).ValorGuia = 0
            Else
                aBoleto(y).ValorGuia = !ValorGuia
            End If
            aBoleto(y).DataDoc = !Datadocumento
            If IsNull(!ValorPagoreal) Then
                aBoleto(y).ValorPago = 0
                aBoleto(y).DataPago = Date
            Else
                aBoleto(y).ValorPago = !ValorPagoreal
                aBoleto(y).DataPago = !DataPagamento
            End If
        Else
            If Not IsNull(!ValorPagoreal) Then
                aBoleto(x).ValorPago = aBoleto(x).ValorPago + !ValorPagoreal
            End If
        End If
       .MoveNext
    Loop
   .Close
End With

Sql = "SELECT DISTINCT NUMDOCUMENTO.numdocumento,datadocumento,valorguia,emissor,userid,valorpagoreal,datapagamento,parceladocumento.codreduzido,razaosocial AS NOME "
Sql = Sql & "FROM numdocumento INNER JOIN parceladocumento ON numdocumento.numdocumento=parceladocumento.numdocumento "
Sql = Sql & "INNER JOIN mobiliario  ON parceladocumento.codreduzido=mobiliario.codigomob LEFT OUTER JOIN debitopago ON numdocumento.numdocumento = debitopago.numdocumento "
Sql = Sql & "WHERE datadocumento between '" & Format(mskDataIni.Text, "mm/dd/yyyy") & "' and '" & Format(mskDataFim.Text, "mm/dd/yyyy") & "' and parceladocumento.CODREDUZIDO>=100000 AND parceladocumento.CODREDUZIDO<200000 and "
If bTodos Then
    Sql = Sql & "(EMISSOR LIKE 'JAINE.JESUS (DAM%' OR EMISSOR LIKE 'LUANA.OLIVEIRA (DAM%' OR EMISSOR LIKE 'ARIANA.SILVA (DAM%' OR EMISSOR LIKE 'SATYLA.SOUZA (DAM%')"
    Sql = Sql & "order by numdocumento"
Else
    nUser = cmbOperador.ItemData(cmbOperador.ListIndex)
    Sql = Sql & "emissor like '" & sUser & "'"
End If
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF

        
        For x = 1 To UBound(aBoleto)
            bFind = False
            If aBoleto(x).Documento = !NumDocumento Then
                bFind = True
                Exit For
            End If
        Next
        If Not bFind Then
          ReDim Preserve aBoleto(UBound(aBoleto) + 1)
          y = UBound(aBoleto)

          
           aBoleto(y).Documento = !NumDocumento
           aBoleto(y).Emissor = !Emissor
            aBoleto(y).Codigo = !CODREDUZIDO
            aBoleto(y).Nome = sNome
            If IsNull(!ValorGuia) Then
                aBoleto(y).ValorGuia = 0
            Else
                aBoleto(y).ValorGuia = !ValorGuia
            End If
            aBoleto(y).DataDoc = !Datadocumento
            If IsNull(!ValorPagoreal) Then
                aBoleto(y).ValorPago = 0
                aBoleto(y).DataPago = Date
            Else
                aBoleto(y).ValorPago = !ValorPagoreal
                aBoleto(y).DataPago = !DataPagamento
            End If
        Else
            If Not IsNull(!ValorPagoreal) Then
                aBoleto(x).ValorPago = aBoleto(x).ValorPago + !ValorPagoreal
            End If
        End If
       
       
       .MoveNext
    Loop
   .Close
End With


Sql = "SELECT DISTINCT NUMDOCUMENTO.numdocumento,datadocumento,valorguia,emissor,userid,valorpagoreal,datapagamento,parceladocumento.codreduzido,nomecidadao as nome "
Sql = Sql & "FROM numdocumento INNER JOIN parceladocumento ON numdocumento.numdocumento=parceladocumento.numdocumento "
Sql = Sql & "INNER JOIN CIDADAO  ON parceladocumento.codreduzido=cidadao.codcidadao LEFT OUTER JOIN debitopago ON numdocumento.numdocumento = debitopago.numdocumento "
Sql = Sql & "WHERE datadocumento between '" & Format(mskDataIni.Text, "mm/dd/yyyy") & "' and '" & Format(mskDataFim.Text, "mm/dd/yyyy") & "' and parceladocumento.CODREDUZIDO>500000 and "
If bTodos Then
    Sql = Sql & "(EMISSOR LIKE 'JAINE.JESUS (DAM%' OR EMISSOR LIKE 'LUANA.OLIVEIRA (DAM%' OR EMISSOR LIKE 'ARIANA.SILVA (DAM%' OR EMISSOR LIKE 'SATYLA.SOUZA (DAM%')"
    Sql = Sql & "order by numdocumento"
Else
    nUser = cmbOperador.ItemData(cmbOperador.ListIndex)
    Sql = Sql & "emissor like '" & sUser & "'"
End If
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF

         For x = 1 To UBound(aBoleto)
            bFind = False
            If aBoleto(x).Documento = !NumDocumento Then
                bFind = True
                Exit For
            End If
        Next
          If Not bFind Then
          ReDim Preserve aBoleto(UBound(aBoleto) + 1)
          y = UBound(aBoleto)
           aBoleto(y).Documento = !NumDocumento
           aBoleto(y).Emissor = !Emissor
            aBoleto(y).Codigo = !CODREDUZIDO
            aBoleto(y).Nome = sNome
            If IsNull(!ValorGuia) Then
                aBoleto(y).ValorGuia = 0
            Else
                aBoleto(y).ValorGuia = !ValorGuia
            End If
            aBoleto(y).DataDoc = !Datadocumento
            If IsNull(!ValorPagoreal) Then
                aBoleto(y).ValorPago = 0
                aBoleto(y).DataPago = Date
            Else
                aBoleto(y).ValorPago = !ValorPagoreal
                aBoleto(y).DataPago = !DataPagamento
            End If
        Else
            If Not IsNull(!ValorPagoreal) Then
                aBoleto(x).ValorPago = aBoleto(x).ValorPago + !ValorPagoreal
            End If
        End If

      .MoveNext
    Loop
   .Close
End With


For x = 1 To UBound(aBoleto)
    With aBoleto(x)
 '       If .Documento = 22118578 Then MsgBox "teste"
'        If .ValorPago = 0 Then
'            Sql = "select * from parceladocumento where numdocumento=" & .Documento
'            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset)
'            With RdoAux
'                If .RowCount > 0 Then
'                    nAno = !AnoExercicio
'                    nLanc = !CodLancamento
'                    nSeq = !SeqLancamento
'                    nParc = !NumParcela
'                    nCompl = !CODCOMPLEMENTO
'                    Sql = "select numdocumento,datarecebimento from debitopago where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & nAno & " and codlancamento=" & nLanc & " and "
'                    Sql = Sql & "seqlancamento=" & nSeq & " and numparcela=" & nParc & " and codcomplemento=" & nCompl
'                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'                    If RdoAux2.RowCount > 0 Then
'                        nDoc = RdoAux2!NumDocumento
'
'                        aBoleto(x).DataPago = RdoAux2!datarecebimento
 '                       RdoAux2.Close
 '                       Sql = "select sum(valorpagoreal) as soma from debitopago where numdocumento=" & nDoc
 '                       Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
 '                       aBoleto(x).ValorPago = RdoAux2!soma
 '                   End If
 '                   RdoAux2.Close
 '               End If
 '              .Close
 '           End With
 '       End If
       
        If .Codigo < 100000 Then
            Sql = "SELECT NOMECIDADAO AS NOME FROM vwFULLIMOVEL WHERE CODREDUZIDO=" & .Codigo
        ElseIf .Codigo >= 100000 And .Codigo < 500000 Then
            Sql = "SELECT RAZAOSOCIAL AS NOME FROM MOBILIARIO WHERE CODIGOMOB=" & .Codigo
        Else
            Sql = "SELECT NOMECIDADAO AS NOME FROM CIDADAO WHERE CODCIDADAO=" & .Codigo
        End If
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        .Nome = RdoAux2!Nome
        RdoAux2.Close
        
       
       If .ValorGuia = 0 Then .ValorGuia = .ValorPago
    End With
Next

For x = 1 To UBound(aBoleto)
    With aBoleto(x)

        Sql = "insert Relatorio_EDinamica1(usuario,documento,emissor,codigo,nome,valorguia,datadoc,valorpago,datapago) values('"
        Sql = Sql & NomeDeLogin & "'," & .Documento & ",'" & .Emissor & "'," & .Codigo & ",'" & .Nome & "'," & Virg2Ponto(CStr(.ValorGuia)) & ",'"
        Sql = Sql & Format(.DataDoc, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(.ValorPago)) & ",'" & Format(.DataPago, "mm/dd/yyyy") & "')"
        cn.Execute Sql, rdExecDirect
    End With
Next

Sql = "update Relatorio_EDinamica1 set datapago=null where usuario='" & NomeDeLogin & "' and valorpago=0"
cn.Execute Sql, rdExecDirect


frmReport.ShowReport3 "ESTRUTURA_DINAMICA1", frmMdi.HWND, Me.HWND
'ShellExecute HWND, "open", "gtiv4.jaboticabal.sp.gov.br/Tributario/relEDinamica?d1=" & mskDataIni.Text & "&d2=" & mskDataFim.Text & "&u=" & NomeDeLogin, vbNullString, vbNullString, conSwNormal

Liberado

Sql = "delete from Relatorio_EDinamica1 where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect


End Sub

