VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmNDevedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de N Devedores"
   ClientHeight    =   2175
   ClientLeft      =   11070
   ClientTop       =   6150
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4965
   Begin VB.ComboBox cmbPessoa 
      Height          =   315
      ItemData        =   "frmNDevedor.frx":0000
      Left            =   3420
      List            =   "frmNDevedor.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1080
      Width           =   1320
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   285
      Left            =   1695
      TabIndex        =   10
      Top             =   1080
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtTopN"
      BuddyDispid     =   196610
      OrigLeft        =   2025
      OrigTop         =   945
      OrigRight       =   2280
      OrigBottom      =   1275
      Max             =   1000
      Min             =   1
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtTopN 
      Height          =   285
      Left            =   945
      TabIndex        =   9
      Text            =   "10"
      Top             =   1080
      Width           =   750
   End
   Begin MSComCtl2.DTPicker dtDataDe 
      Height          =   285
      Left            =   945
      TabIndex        =   6
      Top             =   630
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   503
      _Version        =   393216
      Format          =   77529089
      CurrentDate     =   44208
   End
   Begin VB.OptionButton optTipo 
      Caption         =   "Cidadão"
      Height          =   240
      Index           =   2
      Left            =   2925
      TabIndex        =   3
      Top             =   225
      Width           =   960
   End
   Begin VB.OptionButton optTipo 
      Caption         =   "Empresa"
      Height          =   240
      Index           =   1
      Left            =   1845
      TabIndex        =   2
      Top             =   225
      Width           =   960
   End
   Begin VB.OptionButton optTipo 
      Caption         =   "Imóvel"
      Height          =   240
      Index           =   0
      Left            =   855
      TabIndex        =   1
      Top             =   225
      Value           =   -1  'True
      Width           =   960
   End
   Begin MSComCtl2.DTPicker dtDataAte 
      Height          =   285
      Left            =   3420
      TabIndex        =   7
      Top             =   630
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   503
      _Version        =   393216
      Format          =   117637121
      CurrentDate     =   44208
   End
   Begin MSComctlLib.ProgressBar Pb 
      Height          =   165
      Left            =   180
      TabIndex        =   11
      Top             =   1710
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   291
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   345
      Left            =   3645
      TabIndex        =   13
      Top             =   1620
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Imprimir"
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
      MICON           =   "frmNDevedor.frx":002C
      PICN            =   "frmNDevedor.frx":0048
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
      Caption         =   "Pessoa:"
      Height          =   195
      Index           =   1
      Left            =   2655
      TabIndex        =   14
      Top             =   1125
      Width           =   645
   End
   Begin VB.Label lblPB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100 %"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2340
      TabIndex        =   12
      Top             =   1710
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "Top N:"
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   1125
      Width           =   510
   End
   Begin VB.Label Label2 
      Caption         =   "Data Até:"
      Height          =   195
      Index           =   1
      Left            =   2655
      TabIndex        =   5
      Top             =   675
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Data De:"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   4
      Top             =   675
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo:"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   225
      Width           =   510
   End
End
Attribute VB_Name = "frmNDevedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type tCodigo
    Codigo As Long
    Ano As Integer
    Lanc As Integer
    Seq As Integer
    Parc As Integer
    Compl As Integer
    CodTributo As Integer
    Nome As String
    Tipo As String
    CpfCnpj As String
    ValorP As Double
    ValorM As Double
    ValorJ As Double
    ValorC As Double
    ValorT As Double
End Type

Private Sub cmdPrint_Click()
Dim Sql As String, RdoAux As rdoResultset, Codigo1 As Long, Codigo2 As Long, Data1 As Date, Data2 As Date, aCodigo() As tCodigo, nTop As Integer, nPos As Long, nTot As Long, t As Integer, nUserID As Integer
Dim RdoAux2 As rdoResultset, qd As New rdoQuery

nTop = Val(txtTopN.Text)
'GoTo Report
Ocupado
Set qd.ActiveConnection = cn
qd.QueryTimeout = 0
nUserID = RetornaUsuarioID(NomeDeLogin)
Data1 = dtDataDe.value
Data2 = dtDataAte.value

Sql = "delete from devedorTopN where userid=" & nUserID
cn.Execute Sql, rdExecDirect

ReDim aCodigo(0)
If optTipo(0).value = True Then
    Codigo1 = 1
    Codigo2 = 50000
    Sql = "SELECT distinct codreduzido AS CODIGO,nomecidadao as Nome,cpf,cnpj FROM debitoparcela INNER JOIN VWFULLIMOVEL ON  debitoparcela.codreduzido = VWFULLIMOVEL.codreduzido "
    Sql = Sql & "WHERE debitoparcela.codreduzido BETWEEN " & Codigo1 & " AND " & Codigo2 & " AND "
    If cmbPessoa.ListIndex = 1 Then
        Sql = Sql & "CNPJ IS NULL "
    ElseIf cmbPessoa.ListIndex = 2 Then
        Sql = Sql & "CNPJ IS NOT NULL "
    End If
    Sql = Sql & " AND codlancamento<> 20 and datavencimento BETWEEN '" & Format(Data1, "mm/dd/yyyy") & "' AND '" & Format(Data2, "mm/dd/yyyy") & "' AND (statuslanc=3 or statuslanc=42 or statuslanc=43) "
    Sql = Sql & "ORDER BY debitoparcela.codreduzido"
ElseIf optTipo(1).value = True Then
    Codigo1 = 100000
    Codigo2 = 300000
    Sql = "SELECT distinct codreduzido as codigo,razaosocial as nome,cpf,cnpj FROM debitoparcela INNER JOIN mobiliario ON  codreduzido = codigomob "
    Sql = Sql & "WHERE debitoparcela.codreduzido BETWEEN " & Codigo1 & " AND " & Codigo2 & " AND "
    If cmbPessoa.ListIndex = 1 Then
        Sql = Sql & "CNPJ IS NULL "
    ElseIf cmbPessoa.ListIndex = 2 Then
        Sql = Sql & "CNPJ IS NOT NULL "
    End If
    Sql = Sql & " AND codlancamento<> 20 and datavencimento BETWEEN '" & Format(Data1, "mm/dd/yyyy") & "' AND '" & Format(Data2, "mm/dd/yyyy") & "' AND (statuslanc=3 or statuslanc=42 or statuslanc=43) "
    Sql = Sql & "ORDER BY debitoparcela.codreduzido"
Else
    Codigo1 = 500001
    Codigo2 = 650000
    Sql = "SELECT distinct codcidadao as codigo,nomecidadao as nome,cpf,cnpj FROM debitoparcela INNER JOIN CIDADAO ON  codreduzido = codigomob "
    Sql = Sql & "WHERE debitoparcela.codreduzido BETWEEN " & Codigo1 & " AND " & Codigo2 & " AND "
    If cmbPessoa.ListIndex = 1 Then
        Sql = Sql & "CNPJ IS NULL "
    ElseIf cmbPessoa.ListIndex = 2 Then
        Sql = Sql & "CNPJ IS NOT NULL "
    End If
    Sql = Sql & " AND codlancamento<> 20 and datavencimento BETWEEN '" & Format(Data1, "mm/dd/yyyy") & "' AND '" & Format(Data2, "mm/dd/yyyy") & "' AND (statuslanc=3 or statuslanc=42 or statuslanc=43) "
    Sql = Sql & "ORDER BY debitoparcela.codreduzido"
End If

nTop = Val(txtTopN.Text)

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If t Mod 20 = 0 Then CallPb nPos, CLng(nTot)

        On Error Resume Next
        RdoAux2.Close
        On Error GoTo 0
        
        qd.Sql = "{ Call spEXTRATODEVEDOR(?,?,?,?,?) }"
        qd(0) = !Codigo
        qd(1) = Format(Data1, "mm/dd/yyyy")
        qd(2) = Format(Data2, "mm/dd/yyyy")
        qd(3) = Format(Now, "mm/dd/yyyy")
        qd(4) = "GTI"
        Set RdoAux2 = qd.OpenResultset(rdOpenKeyset)
        With RdoAux2
            Do Until .EOF
                ReDim Preserve aCodigo(UBound(aCodigo) + 1)
                aCodigo(UBound(aCodigo)).Codigo = RdoAux!Codigo
                aCodigo(UBound(aCodigo)).Nome = RdoAux!Nome
                If IsNull(RdoAux!Cnpj) Then
                    aCodigo(UBound(aCodigo)).Tipo = "F"
                    aCodigo(UBound(aCodigo)).CpfCnpj = SubNull(RdoAux!cpf)
                Else
                    aCodigo(UBound(aCodigo)).Tipo = "J"
                    aCodigo(UBound(aCodigo)).CpfCnpj = SubNull(RdoAux!Cnpj)
                End If
                aCodigo(UBound(aCodigo)).Ano = !AnoExercicio
                aCodigo(UBound(aCodigo)).Lanc = !CodLancamento
                aCodigo(UBound(aCodigo)).Seq = !SeqLancamento
                aCodigo(UBound(aCodigo)).Parc = !NumParcela
                aCodigo(UBound(aCodigo)).Compl = !CODCOMPLEMENTO
                aCodigo(UBound(aCodigo)).CodTributo = !CodTributo
                aCodigo(UBound(aCodigo)).ValorP = !ValorTributo
                aCodigo(UBound(aCodigo)).ValorM = !ValorMulta
                aCodigo(UBound(aCodigo)).ValorJ = !ValorJuros
                aCodigo(UBound(aCodigo)).ValorC = !ValorCorrecao
                aCodigo(UBound(aCodigo)).ValorT = !ValorTotal
               .MoveNext
            Loop
           .Close
        End With
        
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

CallPb 100, 100

nPos = 1: nTot = UBound(aCodigo)
For t = 1 To UBound(aCodigo)
    If t Mod 10 = 0 Then CallPb nPos, nTot
    With aCodigo(t)
        Sql = "INSERT INTO dbo.devedorTopN(userid,codigo,ano,lanc,seq,parc,compl,codtributo,nome,tipodoc,cpfcnpj,valorP,valorM,valorJ,valorC,valorT) values("
        Sql = Sql & nUserID & "," & .Codigo & "," & .Ano & "," & .Lanc & "," & .Seq & "," & .Parc & "," & .Compl & "," & .CodTributo & ",'" & Mask(.Nome) & "','"
        Sql = Sql & .Tipo & "','" & .CpfCnpj & "'," & Virg2Ponto(CStr(.ValorP)) & "," & Virg2Ponto(CStr(.ValorM)) & "," & Virg2Ponto(CStr(.ValorJ)) & ","
        Sql = Sql & Virg2Ponto(CStr(.ValorC)) & "," & Virg2Ponto(CStr(.ValorT)) & ")"
        cn.Execute Sql, rdExecDirect
    End With
    nPos = nPos + 1
Next
Liberado
CallPb 100, 100

Report:
frmReport.ShowReport3 "DEVEDORTOPN", frmMdi.HWND, Me.HWND, CLng(nTop)




End Sub

Private Sub Form_Load()
Centraliza Me
cmbPessoa.ListIndex = 0
End Sub

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

