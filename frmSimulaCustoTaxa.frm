VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmSimulaCustoTaxa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simulador de custo de taxa de licença e Iss Fixo"
   ClientHeight    =   2520
   ClientLeft      =   12450
   ClientTop       =   6075
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2520
   ScaleWidth      =   8235
   Begin VB.TextBox txtArea 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   2070
      TabIndex        =   5
      Text            =   "0,00"
      Top             =   1980
      Width           =   1590
   End
   Begin MSComCtl2.DTPicker dtData 
      Height          =   330
      Left            =   135
      TabIndex        =   4
      Top             =   1980
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   582
      _Version        =   393216
      Format          =   164364289
      CurrentDate     =   44221
   End
   Begin VB.ComboBox cmbIss 
      Height          =   315
      Left            =   135
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1170
      Width           =   6720
   End
   Begin VB.ComboBox cmbAliqIss 
      Height          =   315
      Left            =   6930
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1170
      Width           =   1185
   End
   Begin VB.ComboBox cmbAliqTaxa 
      Height          =   315
      Left            =   6930
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   405
      Width           =   1185
   End
   Begin VB.ComboBox cmbTaxa 
      Height          =   315
      Left            =   135
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   405
      Width           =   6720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Área m²"
      Height          =   195
      Index           =   7
      Left            =   2115
      TabIndex        =   15
      Top             =   1665
      Width           =   1500
   End
   Begin VB.Label lblValorIss 
      Alignment       =   2  'Center
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   6075
      TabIndex        =   14
      Top             =   2025
      Width           =   1995
   End
   Begin VB.Label lblValorTaxa 
      Alignment       =   2  'Center
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   4050
      TabIndex        =   13
      Top             =   2025
      Width           =   1995
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Valor do Iss Fixo"
      Height          =   195
      Index           =   6
      Left            =   6120
      TabIndex        =   12
      Top             =   1665
      Width           =   1995
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Valor da Taxa de Licença"
      Height          =   195
      Index           =   5
      Left            =   4050
      TabIndex        =   11
      Top             =   1665
      Width           =   1995
   End
   Begin VB.Label Label1 
      Caption         =   "Data da simulação"
      Height          =   195
      Index           =   4
      Left            =   180
      TabIndex        =   10
      Top             =   1665
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Atividade de Iss Fixo"
      Height          =   195
      Index           =   3
      Left            =   180
      TabIndex        =   9
      Top             =   900
      Width           =   2355
   End
   Begin VB.Label Label1 
      Caption         =   "Aliquota"
      Height          =   195
      Index           =   2
      Left            =   6975
      TabIndex        =   8
      Top             =   900
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Aliquota"
      Height          =   195
      Index           =   1
      Left            =   6975
      TabIndex        =   7
      Top             =   135
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Atividade da Taxa de Licença"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   6
      Top             =   135
      Width           =   2355
   End
End
Attribute VB_Name = "frmSimulaCustoTaxa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type tAliquotaTx
    Codigo As Long
    Nome As String
    Aliquota1 As Double
    Aliquota2 As Double
    Aliquota3 As Double
End Type

Private Type tAliquotaIss
    Codigo As Long
    Nome As String
    Aliquota As Double
End Type

Dim bExec As Boolean, aAliqTx() As tAliquotaTx, aAliqIss() As tAliquotaIss

Private Sub cmbAliqIss_Click()
Calculo
End Sub

Private Sub cmbAliqTaxa_Click()
Calculo
End Sub

Private Sub cmbIss_Click()
If cmbIss.ListIndex = -1 Then Exit Sub
CarregaAliquotaIss
Calculo

End Sub

Private Sub cmbTaxa_Click()

If cmbTaxa.ListIndex = -1 Then Exit Sub
CarregaAliquotaTx
Calculo

End Sub

Private Sub dtData_Change()
Calculo
End Sub

Private Sub Form_Load()
Centraliza Me
ReDim aAliqTx(0): ReDim aAliqIss(0)
bExec = False
CarregaLista
bExec = True
dtData.MinDate = CDate("01/01/" & Year(Now))
dtData.MaxDate = CDate("31/12/" & Year(Now))
Calculo
End Sub

Private Sub CarregaLista()
Dim Sql As String, RdoAux As rdoResultset, nPos As Integer

Ocupado

Sql = "SELECT CODATIVIDADE,DESCATIVIDADE,VALORALIQ1,VALORALIQ2,VALORALIQ3 FROM ATIVIDADE ORDER BY DESCATIVIDADE "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aAliqTx(UBound(aAliqTx) + 1)
        nPos = UBound(aAliqTx)
        aAliqTx(nPos).Codigo = !codatividade
        aAliqTx(nPos).Nome = UCase(!DESCATIVIDADE)
        aAliqTx(nPos).Aliquota1 = !VALORALIQ1
        aAliqTx(nPos).Aliquota2 = !VALORALIQ2
        aAliqTx(nPos).Aliquota3 = !VALORALIQ3
       .MoveNext
    Loop
   .Close
End With


For nPos = 1 To UBound(aAliqTx)
    With aAliqTx(nPos)
        cmbTaxa.AddItem .Nome
        cmbTaxa.ItemData(cmbTaxa.NewIndex) = nPos
    End With
Next

Sql = "SELECT DISTINCT ATIVIDADEISS.CODATIVIDADE,ATIVIDADEISS.DESCATIVIDADE,TABELAISS.TIPOISS "
Sql = Sql & "FROM ATIVIDADEISS INNER JOIN TABELAISS ON ATIVIDADEISS.CODATIVIDADE = TABELAISS.CODIGOATIV "
Sql = Sql & "WHERE TIPOISS=11 ORDER BY ATIVIDADEISS.DESCATIVIDADE "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aAliqIss(UBound(aAliqIss) + 1)
        nPos = UBound(aAliqIss)
        aAliqIss(nPos).Codigo = !codatividade
        aAliqIss(nPos).Nome = UCase(!DESCATIVIDADE)
        aAliqIss(nPos).Aliquota = RetornaAliquotaISS(!codatividade, Now)
       .MoveNext
    Loop
   .Close
End With


For nPos = 1 To UBound(aAliqIss)
    With aAliqIss(nPos)
        cmbIss.AddItem .Nome
        cmbIss.ItemData(cmbIss.NewIndex) = nPos
    End With
Next

cmbTaxa.ListIndex = 0
cmbIss.ListIndex = 0
Liberado


End Sub

Private Sub CarregaAliquotaTx()
Dim nPos As Integer, nValor1 As Double, nValor2 As Double, nValor3 As Double

nPos = cmbTaxa.ItemData(cmbTaxa.ListIndex)
cmbAliqTaxa.Clear
nValor1 = aAliqTx(nPos).Aliquota1
nValor2 = aAliqTx(nPos).Aliquota2
nValor3 = aAliqTx(nPos).Aliquota3
cmbAliqTaxa.AddItem nValor1
If nValor2 > 0 Then
    cmbAliqTaxa.AddItem nValor2
End If
If nValor3 > 0 Then
    cmbAliqTaxa.AddItem nValor3
End If
cmbAliqTaxa.ListIndex = 0

End Sub

Private Sub CarregaAliquotaIss()
Dim nPos As Integer, nValor As Double

nPos = cmbIss.ItemData(cmbIss.ListIndex)
cmbAliqIss.Clear
nValor = aAliqIss(nPos).Aliquota
cmbAliqIss.AddItem nValor
cmbAliqIss.ListIndex = 0

End Sub

Private Sub txtArea_Change()
Calculo
End Sub

Private Sub txtArea_GotFocus()
txtArea.SelStart = 0
txtArea.SelLength = Len(txtArea.Text)
End Sub

Private Sub txtArea_KeyPress(KeyAscii As Integer)

Tweak txtArea, KeyAscii, DecimalPositive, 2
End Sub

Private Sub Calculo()
Dim nValorTaxa As Double, nMeses As Integer, nArea As Double, nValorIss As Double

If Not bExec Then Exit Sub
If IsNumeric(txtArea.Text) Then
    nArea = CDbl(txtArea.Text)
Else
    nArea = 1
End If
If nArea = 0 Then nArea = 1

nMeses = DateDiff("m", CDate(dtData.value), CDate("31/12/" & Year(Now))) + 1
nValorTaxa = CDbl(cmbAliqTaxa.Text) * RetornaUFIR(Year(Now)) * (nMeses / 12) * nArea
nValorIss = CDbl(cmbAliqIss.Text) * RetornaUFIR(Year(Now)) * (nMeses / 12)

lblValorTaxa.Caption = Format(nValorTaxa, "#0.00")
lblValorIss.Caption = Format(nValorIss, "#0.00")

End Sub

