VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmRelatorioDevedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatórios de Devedores"
   ClientHeight    =   2070
   ClientLeft      =   7185
   ClientTop       =   6765
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2070
   ScaleWidth      =   5610
   Begin VB.TextBox txtValor 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1515
      TabIndex        =   10
      Top             =   990
      Width           =   885
   End
   Begin VB.ComboBox cmbAj 
      Height          =   315
      ItemData        =   "frmRelatorioDevedor.frx":0000
      Left            =   3960
      List            =   "frmRelatorioDevedor.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   540
      Width           =   975
   End
   Begin VB.ComboBox cmbDA 
      Height          =   315
      ItemData        =   "frmRelatorioDevedor.frx":0024
      Left            =   1500
      List            =   "frmRelatorioDevedor.frx":0031
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   540
      Width           =   975
   End
   Begin VB.TextBox txtCod2 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   3960
      MaxLength       =   6
      TabIndex        =   3
      Top             =   135
      Width           =   885
   End
   Begin VB.TextBox txtCod1 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1530
      MaxLength       =   6
      TabIndex        =   2
      Top             =   135
      Width           =   885
   End
   Begin Tributacao.XP_ProgressBar Pb 
      Height          =   195
      Left            =   270
      TabIndex        =   0
      Top             =   1620
      Width           =   3615
      _ExtentX        =   6376
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
      Color           =   12632064
      Scrolling       =   1
   End
   Begin prjChameleon.chameleonButton cmdExec 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   4050
      TabIndex        =   1
      ToolTipText     =   "Emitir Relatório"
      Top             =   1530
      Width           =   1320
      _ExtentX        =   2328
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
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRelatorioDevedor.frx":0048
      PICN            =   "frmRelatorioDevedor.frx":0064
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin esMaskEdit.esMaskedEdit mskVenc 
      Height          =   285
      Left            =   3945
      TabIndex        =   11
      Top             =   990
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   503
      MouseIcon       =   "frmRelatorioDevedor.frx":0103
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor até.............:"
      Height          =   255
      Index           =   7
      Left            =   180
      TabIndex        =   13
      Top             =   1035
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento até...:"
      Height          =   255
      Index           =   8
      Left            =   2565
      TabIndex        =   12
      Top             =   1035
      Width           =   1395
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ajuizado..............:"
      Height          =   255
      Index           =   4
      Left            =   2610
      TabIndex        =   9
      Top             =   600
      Width           =   1395
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Divida Ativa........:"
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   8
      Top             =   600
      Width           =   1395
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Final........:"
      Height          =   255
      Left            =   2610
      TabIndex        =   5
      Top             =   165
      Width           =   1395
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Inicial......:"
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   4
      Top             =   165
      Width           =   1395
   End
End
Attribute VB_Name = "frmRelatorioDevedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Debito
    nCodReduz As Long
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
    nValorGeral As Double
    nValorHon As Double
    nValorJurApl As Double
    nSaldo As Double
    nCodBanco As Integer
    dDataPag As Date
    sNome As String
    sFullLanc As String
End Type

Private Sub Form_Load()
Centraliza Me
End Sub

Private Sub CallPb(nPosF As Long, nTotal As Long)
On Error GoTo Erro
If nPosF > 0 Then
    Pb.Color = &HC0C000
Else
    Pb.Color = vbWhite
End If
If cGetInputState() <> 0 Then DoEvents
If ((nPosF * 100) / nTotal) <= 100 Then
   Pb.value = (nPosF * 100) / nTotal
Else
   Pb.value = 100
End If
Me.Refresh
If cGetInputState() <> 0 Then DoEvents
Exit Sub
Erro:
MsgBox Err.Description
End Sub

Private Sub mskVenc_GotFocus()
mskVenc.SelStart = 0
mskVenc.SelLength = Len(mskVenc.Text)
End Sub

Private Sub txtCod1_GotFocus()
txtCod1.SelStart = o
txtCod1.SelLength = Len(txtCod1.Text)
End Sub

Private Sub txtCod2_GotFocus()
txtCod2.SelStart = 0
txtCod2.SelLength = Len(txtCod2.Text)
End Sub

Private Sub txtValor_GotFocus()
txtValor.SelStart = 0
txtValor.SelLength = Len(txtValor.Text)
End Sub

Private Sub cmdExec_Click()
Dim aCodigo() As Long, nPos As Long, Sql As String, RdoAux As rdoResultset, sDA As String, sAj As String, nSoma As Double
Dim nValorMax As Double, sDataMax As String, nCod1 As Long, nCod2 As Long, nTot As Long, nCodigo As Long, nValor As Double
Dim qd As New rdoQuery, aDebito() As Debito, Achou As Boolean, nEval As Integer, x As Integer

If MsgBox("Executar o relatório?", vbYesNo, "Confirmação") = vbNo Then Exit Sub
cmdExec.Enabled = False
Me.Refresh

ReDim aCodigo(0)

Sql = "delete from relatorio_devedor"
cn.Execute Sql, rdExecDirect

nCod1 = Val(txtCod1.Text)
nCod2 = Val(txtCod2.Text)

If cmbDA.ListIndex = 0 Then
    sDA = "T"
ElseIf cmbDA.ListIndex = 1 Then
    sDA = "S"
ElseIf cmbDA.ListIndex = 2 Then
    sDA = "N"
End If

If cmbAj.ListIndex = 0 Then
    sAj = "T"
ElseIf cmbAj.ListIndex = 1 Then
    sAj = "S"
ElseIf cmbAj.ListIndex = 2 Then
    sAj = "N"
End If

If Val(txtValor.Text) > 0 Then
    nValorMax = CDbl(txtValor.Text)
End If

If IsDate(mskVenc.Text) Then
    sDataMax = mskVenc.Text
End If

Sql = "SELECT distinct codreduzido FROM debitoparcela WHERE STATUSLANC in (3,42,43) "
If nCod2 > 0 Then
    Sql = Sql & " AND CODREDUZIDO BETWEEN " & nCod1 & " AND " & nCod2
End If
If IsDate(sDataMax) Then
    Sql = Sql & " AND DATAVENCIMENTO<'" & Format(sDataMax, "mm/dd/yyyy") & "'"
End If

If sDA = "S" Then
    Sql = Sql & " AND DATAINSCRICAO IS NOT NULL"
ElseIf sDA = "N" Then
    Sql = Sql & " AND DATAINSCRICAO IS NULL"
End If

If sAj = "S" Then
    Sql = Sql & " AND DATAAJUIZA IS NOT NULL"
ElseIf sAj = "N" Then
    Sql = Sql & " AND DATAAJUIZA IS NULL"
End If
Sql = Sql & " ORDER BY CODREDUZIDO"

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nPos = 1
    nTot = .RowCount
    Do Until .EOF
        ReDim Preserve aCodigo(UBound(aCodigo) + 1)
        aCodigo(UBound(aCodigo)) = !CODREDUZIDO
       .MoveNext
    Loop
   .Close
End With

nTot = UBound(aCodigo)
For nPos = 1 To nTot
    nCodigo = aCodigo(nPos)
    nValor = 0
    ReDim aDebito(0)
    If nPos Mod 10 = 0 Then CallPb nPos, nTot
    
    'CARREGA O EXTRATO
    Set qd.ActiveConnection = cn
    qd.QueryTimeout = 0
    On Error Resume Next
    RdoAux.Close
    On Error GoTo 0
    qd.Sql = "{ Call spEXTRATOLISTADEVEDOR(?,?) }"
    qd(0) = nCodigo
    qd(1) = Format(Now, "mm/dd/yyyy")
    Set RdoAux = qd.OpenResultset(rdOpenKeyset)
    With RdoAux
        If RdoAux.RowCount > 0 Then
            nEval = UBound(aDebito)
            Do Until .EOF
                bJuros = False: bMulta = False
                If sAj = "S" Then
                    If IsNull(!dataajuiza) Then GoTo Proximo
                End If
                If sAj = "N" Then
                    If Not IsNull(!dataajuiza) Then GoTo Proximo
                End If
                If sDA = "S" Then
                    If IsNull(!datainscricao) Then GoTo Proximo
                End If
                If sDA = "N" Then
                    If Not IsNull(!datainscricao) Then GoTo Proximo
                End If
                
                If IsDate(sDataMax) Then
                    If !DataVencimento > CDate(sDataMax) Then GoTo Proximo
                End If
    
               
                Achou = False
                For x = 1 To nEval
                    If aDebito(x).nAno = !AnoExercicio And aDebito(x).nLanc = !CodLancamento And _
                       aDebito(x).nSeq = !SeqLancamento And _
                       aDebito(x).nParc = !NumParcela And aDebito(x).nCompl = !CODCOMPLEMENTO Then
                       Achou = True
                       Exit For
                    End If
                Next
                
                If Not Achou Then
                    ReDim Preserve aDebito(UBound(aDebito) + 1)
                    nEval = UBound(aDebito)
                    aDebito(nEval).nCodReduz = !CODREDUZIDO
                    aDebito(nEval).nAno = !AnoExercicio
                    aDebito(nEval).nLanc = !CodLancamento
                    aDebito(nEval).nSeq = !SeqLancamento
                    aDebito(nEval).nParc = !NumParcela
                    aDebito(nEval).nCompl = !CODCOMPLEMENTO
                    aDebito(nEval).nSituacao = !statuslanc
                    aDebito(nEval).sSituacao = !Situacao
                    aDebito(nEval).sVencto = Format(!DataVencimento, "dd/mm/yyyy")
                    aDebito(nEval).sDA = IIf(IsNull(!datainscricao), "N", "S")
                    aDebito(nEval).sAj = IIf(IsNull(!dataajuiza), "N", "S")
                    aDebito(nEval).nCodTributo = !CodTributo
                    aDebito(nEval).nValorTributo = !ValorTributo
                    aDebito(nEval).nValorMulta = !ValorMulta
                    aDebito(nEval).nValorJuros = !ValorJuros
                    aDebito(nEval).nValorCorrecao = !ValorCorrecao
                    aDebito(nEval).nValorAtual = !ValorTotal
                    aDebito(nEval).nValorGeral = !ValorTotal
                    aDebito(nEval).sNome = sNome
                    aDebito(nEval).sFullLanc = !DESCLANCAMENTO
                Else
                    If aDebito(x).nCodTributo <> !CodTributo Then
                        aDebito(x).nValorTributo = aDebito(x).nValorTributo + !ValorTributo
                        aDebito(x).nValorAtual = aDebito(x).nValorAtual + !ValorTotal
                        aDebito(x).nValorGeral = aDebito(x).nValorGeral + !ValorTotal
                    End If
                End If
Proximo:
                            
                .MoveNext
            Loop
          End If
       .Close
    End With
    
    nValor = 0
    For x = 0 To UBound(aDebito)
        nValor = nValor + aDebito(x).nValorAtual
    Next
    
    If nValor > 0 Then
        If nValorMax > 0 Then
            If nValor > nValorMax Then
                GoTo proximo2
            End If
        End If
        Sql = "INSERT relatorio_devedor(CODIGO,VALOR) VALUES(" & nCodigo & "," & Virg2Ponto(CStr(nValor)) & ")"
        cn.Execute Sql, rdExecDirect
    End If

    
proximo2:
Next





'   .MoveFirst
'    Do Until .EOF
 '       If nPos Mod 10 = 0 Then CallPb nPos, nTot
 '
 '       nPos = nPos + 1
 '      .MoveNext
 '   Loop


cmdExec.Enabled = True
Pb.value = 100
MsgBox "Relatório finalizado!"

End Sub

