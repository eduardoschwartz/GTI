VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmDivideDebito 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Divisão de Débito"
   ClientHeight    =   7050
   ClientLeft      =   8220
   ClientTop       =   4620
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   150
      Left            =   135
      TabIndex        =   11
      Top             =   2475
      Width           =   9015
   End
   Begin VB.TextBox txtValor 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2745
      MaxLength       =   15
      TabIndex        =   1
      Top             =   2025
      Width           =   1335
   End
   Begin prjChameleon.chameleonButton cmdGerar 
      Height          =   375
      Left            =   7245
      TabIndex        =   0
      ToolTipText     =   "Efetua a divisão da parcela"
      Top             =   6570
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Dividir o Débito"
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
      MICON           =   "frmDivideDebito.frx":0000
      PICN            =   "frmDivideDebito.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid grdTemp 
      Height          =   1305
      Left            =   90
      TabIndex        =   2
      Top             =   495
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   2302
      _Version        =   393216
      Rows            =   1
      Cols            =   14
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      BackColorSel    =   128
      ForeColorSel    =   16777215
      GridColorFixed  =   16777215
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmDivideDebito.frx":00BB
   End
   Begin prjChameleon.chameleonButton cmdSimular 
      Cancel          =   -1  'True
      Height          =   330
      Left            =   4185
      TabIndex        =   4
      ToolTipText     =   "Cancelar Edição"
      Top             =   1980
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   582
      BTYPE           =   3
      TX              =   "&Simular Parcelas"
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
      MICON           =   "frmDivideDebito.frx":014F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid grdDestino1 
      Height          =   1305
      Left            =   90
      TabIndex        =   12
      Top             =   3420
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   2302
      _Version        =   393216
      Rows            =   1
      Cols            =   14
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      BackColorSel    =   128
      ForeColorSel    =   16777215
      GridColorFixed  =   16777215
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmDivideDebito.frx":016B
   End
   Begin MSFlexGridLib.MSFlexGrid grdDestino2 
      Height          =   1305
      Left            =   45
      TabIndex        =   13
      Top             =   5130
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   2302
      _Version        =   393216
      Rows            =   1
      Cols            =   14
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      BackColorSel    =   128
      ForeColorSel    =   16777215
      GridColorFixed  =   16777215
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmDivideDebito.frx":01FF
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Complemento Gerado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   5
      Left            =   180
      TabIndex        =   10
      Top             =   4860
      Width           =   2490
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Parcela Reduzida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   165
      Index           =   4
      Left            =   135
      TabIndex        =   9
      Top             =   3105
      Width           =   2490
   End
   Begin VB.Label lblPerc 
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   7200
      TabIndex        =   8
      Top             =   2070
      Width           =   510
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "% Redução.:"
      Height          =   255
      Index           =   3
      Left            =   6210
      TabIndex        =   7
      Top             =   2070
      Width           =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Parcelas a serem criadas:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   2
      Left            =   3105
      TabIndex        =   6
      Top             =   2790
      Width           =   2490
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Parcela de Origem:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   210
      Index           =   1
      Left            =   135
      TabIndex        =   5
      Top             =   225
      Width           =   2490
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor corrigido da primeira parcela.:"
      Height          =   255
      Index           =   0
      Left            =   135
      TabIndex        =   3
      Top             =   2070
      Width           =   2490
   End
End
Attribute VB_Name = "frmDivideDebito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sAno As String, sLanc As String, sSeq As String, sParc As String, sComp As String, sVencto As String, nStatus As Integer
Dim nValorTotal As Double, nCodReduz As Long, nValorPrincipal As Double, nValorCorrecao As Double, nValorJuros As Double, nValorMulta As Double
Dim nValorP1 As Double, nValorJ1 As Double, nValorM1 As Double, nValorC1 As Double, nValorT1 As Double, nSomaGeral As Double
Dim nValorP2 As Double, nValorJ2 As Double, nValorM2 As Double, nValorC2 As Double, nValorT2 As Double

Private Sub cmdSimular_Click()
Dim n As Double

If Val(txtValor.Text) = 0 Then
    MsgBox "Digite um valor válido", vbCritical, "Erro"
    Exit Sub
End If

If CDbl(txtValor.Text) >= nSomaGeral Then
    MsgBox "O valor da primeira parcela deverá der menor que o valor total da parcela de origem.", vbCritical, "Erro"
    Exit Sub
End If

CarregaDestino

End Sub

Private Sub Form_Load()
Carrega_Origem
End Sub

Private Sub Carrega_Origem()
Dim nSomaP As Double, nSomaJ As Double, nSomaM As Double, nSomaC As Double, nSomaT As Double
Dim qd As New rdoQuery, RdoAux As rdoResultset, sDA As String, sAj As String

With frmDebitoImob.grdExtrato
    nCodReduz = Val(frmDebitoImob.txtCod.Text)
    For x = 1 To .Rows
        If .CellText(x, 12) = "S" Then
            sAno = .CellText(x, 1)
            sLanc = Left$(.CellText(x, 2), 3)
            sSeq = .CellText(x, 3)
            sParc = IIf(.CellText(x, 4) = "Unica", "00", .CellText(x, 4))
            sComp = .CellText(x, 5)
            sVencto = .CellText(x, 7)
                       
            '***********************
            'CARREGA O EXTRATO
            Set qd.ActiveConnection = cn
            On Error Resume Next
            RdoAux.Close
            On Error GoTo 0
            qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
            qd(0) = nCodReduz
            qd(1) = nCodReduz
            qd(2) = Val(sAno)
            qd(3) = Val(sAno)
            qd(4) = Val(sLanc)
            qd(5) = Val(sLanc)
            qd(6) = Val(sSeq)
            qd(7) = Val(sSeq)
            qd(8) = Val(sParc)
            qd(9) = Val(sParc)
            qd(10) = Val(sComp)
            qd(11) = Val(sComp)
            qd(12) = 1
            qd(13) = 99
            qd(14) = Format(Now, "mm/dd/yyyy")
            qd(15) = NomeDoUsuario
            Set RdoAux = qd.OpenResultset(rdOpenKeyset)
            With RdoAux
                Do Until .EOF
                    nValorPrincipal = !VALORTRIBUTO
                    nValorJuros = !ValorJuros
                    nValorMulta = !ValorMulta
                    nValorCorrecao = !valorcorrecao
                    nValorTotal = !ValorTotal
                    
                    nSomaP = nSomaP + nValorPrincipal
                    nSomaJ = nSomaJ + nValorJuros
                    nSomaM = nSomaM + nValorMulta
                    nSomaC = nSomaC + nValorCorrecao
                    nSomaT = nSomaT + nValorTotal
                    
                    grdTemp.AddItem sAno & Chr(9) & sLanc & Chr(9) & sSeq & Chr(9) & sParc & Chr(9) & _
                    sComp & Chr(9) & Format(!CodTributo, "000") & Chr(9) & sVencto & Chr(9) & FormatNumber(nValorPrincipal, 2) & Chr(9) & FormatNumber(nValorCorrecao, 2) & Chr(9) & _
                    FormatNumber(nValorMulta, 2) & Chr(9) & FormatNumber(nValorJuros, 2) & Chr(9) & FormatNumber(nValorTotal, 2)
                   
                   .MoveNext
                Loop
               .Close
            End With
                       
        End If
    Next
End With

grdTemp.AddItem "Valor Total ===>" & Chr(9) & "Valor Total ===>" & Chr(9) & "Valor Total ===>" & Chr(9) & "Valor Total ===>" & Chr(9) & _
"Valor Total ===>" & Chr(9) & "Valor Total ===>" & Chr(9) & "Valor Total ===>" & Chr(9) & FormatNumber(nSomaP, 2) & Chr(9) & FormatNumber(nSomaC, 2) & Chr(9) & _
FormatNumber(nSomaM, 2) & Chr(9) & FormatNumber(nSomaJ, 2) & Chr(9) & FormatNumber(nSomaT, 2)
grdTemp.col = 0
nSomaGeral = nSomaT
grdTemp.col = 0
grdTemp.MergeCells = flexMergeFree
grdTemp.MergeRow(grdTemp.Rows - 1) = True

For x = 0 To grdTemp.Cols - 1
    grdTemp.Row = grdTemp.Rows - 1
    grdTemp.col = x
    grdTemp.CellBackColor = Vinho
    grdTemp.CellForeColor = vbWhite
Next

End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
Tweak txtValor, KeyAscii, DecimalPositive, 2
End Sub

Private Sub CarregaDestino()
Dim Sql As String, RdoAux As rdoResultset, nCompl As Integer, nPerc As Double, x As Integer, nCodTributo As Integer
Dim nSomaP1 As Double, nSomaJ1 As Double, nSomaM1 As Double, nSomaC1 As Double, nSomaT1 As Double
Dim nSomaP2 As Double, nSomaJ2 As Double, nSomaM2 As Double, nSomaC2 As Double, nSomaT2 As Double
Dim nCompl1 As Integer, nCompl2 As Integer

Sql = "select * from debitotributo where codreduzido=" & nCodReduz & " and anoexercicio=" & Val(sAno) & " and codlancamento=" & Val(sLanc) & " and "
Sql = Sql & "seqlancamento=" & Val(sSeq) & " and numparcela=" & Val(sParc) & " order by codcomplemento desc"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nCompl = RdoAux!CODCOMPLEMENTO
nCompl1 = nCompl + 1
nCompl2 = nCompl + 2

nPerc = CDbl(txtValor.Text) * 100 / nSomaGeral
lblPerc.Caption = Format(nPerc, "#0.0000")
nPerc = nPerc / 100

grdDestino1.Rows = 1: grdDestino2.Rows = 1

For x = 1 To grdTemp.Rows - 2
    With grdTemp
        nCodTributo = .TextMatrix(x, 5)
        nValorPrincipal = .TextMatrix(x, 7)
        nValorCorrecao = .TextMatrix(x, 8)
        nValorMulta = .TextMatrix(x, 9)
        nValorJuros = .TextMatrix(x, 10)
        nValorTotal = .TextMatrix(x, 11)
    End With
    nValorP1 = nValorPrincipal * nPerc
    nValorP2 = nValorPrincipal - nValorP1
    nValorJ1 = nValorJuros * nPerc
    nValorJ2 = nValorJuros - nValorJ1
    nValorM1 = nValorMulta * nPerc
    nValorM2 = nValorMulta - nValorM1
    nValorC1 = nValorCorrecao * nPerc
    nValorC2 = nValorCorrecao - nValorC1
    nValorT1 = nValorP1 + nValorM1 + nValorJ1 + nValorC1
    nValorT2 = nValorP2 + nValorM2 + nValorJ2 + nValorC2
    
    nSomaP1 = nSomaP1 + nValorP1
    nSomaJ1 = nSomaJ1 + nValorJ1
    nSomaM1 = nSomaM1 + nValorM1
    nSomaC1 = nSomaC1 + nValorC1
    nSomaT1 = nSomaT1 + nValorT1
    nSomaP2 = nSomaP2 + nValorP2
    nSomaJ2 = nSomaJ2 + nValorJ2
    nSomaM2 = nSomaM2 + nValorM2
    nSomaC2 = nSomaC2 + nValorC2
    nSomaT2 = nSomaT2 + nValorT2

    
    'Parcela 1
    grdDestino1.AddItem sAno & Chr(9) & sLanc & Chr(9) & sSeq & Chr(9) & sParc & Chr(9) & _
                nCompl1 & Chr(9) & Format(nCodTributo, "000") & Chr(9) & sVencto & Chr(9) & FormatNumber(nValorP1, 4) & Chr(9) & FormatNumber(nValorC1, 4) & Chr(9) & _
                FormatNumber(nValorM1, 4) & Chr(9) & FormatNumber(nValorJ1, 4) & Chr(9) & FormatNumber(nValorT1, 4)
    
    'Parcela 2
    grdDestino2.AddItem sAno & Chr(9) & sLanc & Chr(9) & sSeq & Chr(9) & sParc & Chr(9) & _
                nCompl2 & Chr(9) & Format(nCodTributo, "000") & Chr(9) & sVencto & Chr(9) & FormatNumber(nValorP2, 4) & Chr(9) & FormatNumber(nValorC2, 4) & Chr(9) & _
                FormatNumber(nValorM2, 4) & Chr(9) & FormatNumber(nValorJ2, 2) & Chr(9) & FormatNumber(nValorT2, 4)
Next

With grdDestino1
    .AddItem "Valor Total ===>" & Chr(9) & "Valor Total ===>" & Chr(9) & "Valor Total ===>" & Chr(9) & "Valor Total ===>" & Chr(9) & _
    "Valor Total ===>" & Chr(9) & "Valor Total ===>" & Chr(9) & "Valor Total ===>" & Chr(9) & FormatNumber(nSomaP1, 2) & Chr(9) & FormatNumber(nSomaC1, 2) & Chr(9) & _
    FormatNumber(nSomaM1, 2) & Chr(9) & FormatNumber(nSomaJ1, 2) & Chr(9) & FormatNumber(nSomaT1, 2)
    .col = 0
    .col = 0
    .MergeCells = flexMergeFree
    .MergeRow(.Rows - 1) = True
    
    For x = 0 To .Cols - 1
        .Row = .Rows - 1
        .col = x
        .CellBackColor = Vinho
        .CellForeColor = vbWhite
    Next
End With

With grdDestino2
    .AddItem "Valor Total ===>" & Chr(9) & "Valor Total ===>" & Chr(9) & "Valor Total ===>" & Chr(9) & "Valor Total ===>" & Chr(9) & _
    "Valor Total ===>" & Chr(9) & "Valor Total ===>" & Chr(9) & "Valor Total ===>" & Chr(9) & FormatNumber(nSomaP2, 2) & Chr(9) & FormatNumber(nSomaC2, 2) & Chr(9) & _
    FormatNumber(nSomaM2, 2) & Chr(9) & FormatNumber(nSomaJ2, 2) & Chr(9) & FormatNumber(nSomaT2, 2)
    .col = 0
    .col = 0
    .MergeCells = flexMergeFree
    .MergeRow(.Rows - 1) = True
    
    For x = 0 To .Cols - 1
        .Row = .Rows - 1
        .col = x
        .CellBackColor = Vinho
        .CellForeColor = vbWhite
    Next
End With


End Sub

Private Sub cmdGerar_Click()
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, nUserId As Integer, x As Integer, nCodTributo As Integer, nValorTributo As Double, sExecFiscal As String
Dim nNumCertidao  As Long, sDataInscricao As String, sDataAjuiza As String, nLivro As Integer, nPagina As Integer

nUserId = RetornaUsuarioID(NomeDeLogin)
nAno = Val(sAno)
nLanc = Val(sLanc)
nParc = Val(sParc)
nSeq = Val(sSeq)

If grdDestino1.Rows = 1 Then
    MsgBox "Nada a gerar", vbCritical, "Erro"
    Exit Sub
End If

If MsgBox("Deseja efetuar a divisão desta parcela?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub

Sql = "select * from debitoparcela  WHERE CODREDUZIDO = " & nCodReduz & " And AnoExercicio = " & nAno & " And CodLancamento = " & nLanc & " And SeqLancamento = " & nSeq & " And "
Sql = Sql & "NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & Val(sComp)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
sExecFiscal = SubNull(RdoAux!processocnj)
nNumCertidao = Val(SubNull(RdoAux!numcertidao))
If SubNull(RdoAux!datainscricao) <> "" Then
    sDataInscricao = Format(RdoAux!datainscricao, "mm/dd/yyyy")
Else
    sDataInscricao = ""
End If
If SubNull(RdoAux!datainscricao) <> "" Then
    sDataAjuiza = Format(RdoAux!dataajuiza, "mm/dd/yyyy")
End If
nNumLivro = Val(SubNull(RdoAux!numerolivro))
nPagina = Val(SubNull(RdoAux!paginalivro))

'Parcela 1
With grdDestino1
    nCompl = Val(.TextMatrix(1, 4))
    Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,USERID,PROCESSOCNJ,"
    Sql = Sql & "NUMCERTIDAO,DATAINSCRICAO,DATAAJUIZA,NUMEROLIVRO,PAGINALIVRO) VALUES("
    Sql = Sql & nCodReduz & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & "," & 3 & ",'" & Format(sVencto, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "'," & nUserId & ",'"
    Sql = Sql & sExecFiscal & "'," & nNumCertidao & ",'" & sDataInscricao & "','" & sDataAjuiza & "'," & nNumLivro & "," & nPagina & ")"
    cn.Execute Sql, rdExecDirect
    For x = 1 To .Rows - 2
        nCodTributo = Val(.TextMatrix(x, 5))
        nValorTributo = CDbl(.TextMatrix(x, 7))
        Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
        Sql = Sql & nCodReduz & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & "," & nCodTributo & "," & Virg2Ponto(CStr(nValorTributo)) & ")"
        cn.Execute Sql, rdExecDirect
    Next
End With
Sql = "INSERT OBSPARCELA SELECT CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA," & nCompl & ",SEQ,OBS,DATA,USERID FROM OBSPARCELA WHERE "
Sql = Sql & "CODREDUZIDO = " & nCodReduz & " And AnoExercicio = " & nAno & " And CodLancamento = " & nLanc & " And SeqLancamento = " & nSeq & " And "
Sql = Sql & "NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & Val(sComp)
cn.Execute Sql, rdExecDirect


'Parcela 2
With grdDestino2
    nCompl = Val(.TextMatrix(1, 4))
    Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,USERID,PROCESSOCNJ,"
    Sql = Sql & "NUMCERTIDAO,DATAINSCRICAO,DATAAJUIZA,NUMEROLIVRO,PAGINALIVRO) VALUES("
    Sql = Sql & nCodReduz & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & "," & 3 & ",'" & Format(sVencto, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "'," & nUserId & ",'"
    Sql = Sql & sExecFiscal & "'," & nNumCertidao & ",'" & sDataInscricao & "','" & sDataAjuiza & "'," & nNumLivro & "," & nPagina & ")"
    cn.Execute Sql, rdExecDirect
    For x = 1 To .Rows - 2
        nCodTributo = Val(.TextMatrix(x, 5))
        nValorTributo = CDbl(.TextMatrix(x, 7))
        Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
        Sql = Sql & nCodReduz & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & "," & nCodTributo & "," & Virg2Ponto(CStr(nValorTributo)) & ")"
        cn.Execute Sql, rdExecDirect
    Next
End With
Sql = "INSERT OBSPARCELA SELECT CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA," & nCompl & ",SEQ,OBS,DATA,USERID FROM OBSPARCELA WHERE "
Sql = Sql & "CODREDUZIDO = " & nCodReduz & " And AnoExercicio = " & nAno & " And CodLancamento = " & nLanc & " And SeqLancamento = " & nSeq & " And "
Sql = Sql & "NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & Val(sComp)
cn.Execute Sql, rdExecDirect

'Atualiza Status Otigem
Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=46 WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & Val(sComp)
cn.Execute Sql, rdExecDirect


MsgBox "O débito foi divido com sucesso, atualize o extrato para ver as alterações", vbInformation, "Informação"

Unload Me


End Sub

