VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmAnalise 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "An�lise da Receita"
   ClientHeight    =   1965
   ClientLeft      =   7455
   ClientTop       =   4290
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   11100
   Begin Tributacao.jcFrames jcFrames2 
      Height          =   915
      Left            =   90
      Top             =   990
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   1614
      FrameColor      =   12829635
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Relat�rios de Pagamento"
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
      Begin MSComCtl2.DTPicker dtDataDe 
         Height          =   315
         Left            =   960
         TabIndex        =   14
         Top             =   360
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   125829121
         CurrentDate     =   42026
      End
      Begin MSComCtl2.DTPicker dtDataAte 
         Height          =   315
         Left            =   3480
         TabIndex        =   16
         Top             =   360
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   125829121
         CurrentDate     =   42026
      End
      Begin prjChameleon.chameleonButton btBanco 
         Height          =   315
         Left            =   5130
         TabIndex        =   18
         Top             =   360
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Relat�rio por Banco"
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
         MICON           =   "frmAnalise.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton btFicha 
         Cancel          =   -1  'True
         Height          =   315
         Left            =   7380
         TabIndex        =   19
         Top             =   360
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Relat�rio por Ficha"
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
         MICON           =   "frmAnalise.frx":001C
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
         Caption         =   "Data at�.:"
         Height          =   195
         Index           =   4
         Left            =   2700
         TabIndex        =   17
         Top             =   420
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Data de.:"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   15
         Top             =   420
         Width           =   780
      End
   End
   Begin Tributacao.jcFrames jcFrames1 
      Height          =   915
      Left            =   60
      Top             =   30
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   1614
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
      Begin VB.ComboBox cmbCC 
         Height          =   315
         ItemData        =   "frmAnalise.frx":0038
         Left            =   720
         List            =   "frmAnalise.frx":0048
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   510
         Width           =   2025
      End
      Begin VB.OptionButton opt1 
         Caption         =   "Banco...:"
         Height          =   255
         Index           =   1
         Left            =   4020
         TabIndex        =   5
         Top             =   180
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.OptionButton opt1 
         Caption         =   "Simples Nacional"
         Height          =   255
         Index           =   0
         Left            =   2250
         TabIndex        =   4
         Top             =   180
         Width           =   1575
      End
      Begin VB.ComboBox cmbBanco 
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   150
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker dpData 
         Height          =   315
         Left            =   720
         TabIndex        =   2
         Top             =   150
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   126222337
         CurrentDate     =   42026
      End
      Begin prjChameleon.chameleonButton cmdGerar 
         Height          =   315
         Left            =   6900
         TabIndex        =   0
         ToolTipText     =   "Executar an�lise"
         Top             =   510
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Gerar"
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
         MICON           =   "frmAnalise.frx":0087
         PICN            =   "frmAnalise.frx":00A3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ProgressBar Pb 
         Height          =   225
         Left            =   3450
         TabIndex        =   11
         Top             =   570
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 %"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   3120
         TabIndex        =   13
         Top             =   570
         Width           =   270
      End
      Begin VB.Label lblPB 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100 %"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5685
         TabIndex        =   12
         Top             =   585
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "C/C...:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   570
         Width           =   555
      End
      Begin VB.Label lblTotalDia 
         Caption         =   "R$ 0,00"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   9180
         TabIndex        =   7
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label1 
         Caption         =   "ValorTotal...:"
         Height          =   195
         Index           =   1
         Left            =   8190
         TabIndex        =   6
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Data..:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   210
         Width           =   555
      End
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   5220
      Left            =   30
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   12825
      _ExtentX        =   22622
      _ExtentY        =   9208
      SortKey         =   12
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   23
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Bco"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "C�digo"
         Object.Width           =   1377
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Ano"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Lc"
         Object.Width           =   881
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Sq"
         Object.Width           =   881
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Pc"
         Object.Width           =   881
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Cp"
         Object.Width           =   881
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Tb"
         Object.Width           =   881
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Pnc"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Jrs"
         Object.Width           =   1409
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Mlt"
         Object.Width           =   1409
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Crc"
         Object.Width           =   1409
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "Tot"
         Object.Width           =   1409
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "DA"
         Object.Width           =   776
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Aj"
         Object.Width           =   776
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Text            =   "Pago"
         Object.Width           =   1408
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Text            =   "Pr"
         Object.Width           =   1409
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   17
         Text            =   "Jr"
         Object.Width           =   1409
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   18
         Text            =   "Ml"
         Object.Width           =   1409
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   19
         Text            =   "Cr"
         Object.Width           =   1409
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   20
         Text            =   "F1"
         Object.Width           =   883
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   21
         Text            =   "F2"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   22
         Text            =   "F3"
         Object.Width           =   882
      EndProperty
   End
End
Attribute VB_Name = "frmAnalise"
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
    nCodFicha As Integer
    sDescFicha As String
    nValor As Double
    nValorPago As Double
    nCodTributo As Integer
    nCodBanco As Integer
    sDescTributo As String
    sArquivo As String
    sNatureza As String
    sVinculo As String
    nPerc As Integer
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
End Type


Dim aFicha() As Ficha, aCodFicha() As Long, aFichaDetalhe() As FichaDetalhe, aFichas() As FichaDetalhe

Private Sub CarregaTributo()
Dim Sql As String, RdoAux As rdoResultset

ReDim aFicha(0): ReDim aCodFicha(0)

Sql = "SELECT codtributo, ficha, fichajrmulta, fichadivida, fichadajrmul, fichadaenca, fichaajuiza, fichaajjrmul, fichaajenca FROM tributo order by codtributo"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aFicha(UBound(aFicha) + 1)
        ReDim Preserve aCodFicha(UBound(aCodFicha) + 1)
        aCodFicha(UBound(aCodFicha)) = !CodTributo
        aFicha(UBound(aFicha)).CodTributo = !CodTributo
        aFicha(UBound(aFicha)).F1 = !Ficha
        aFicha(UBound(aFicha)).F2 = !FichaJrMulta
        aFicha(UBound(aFicha)).F3 = !FichaDivida
        aFicha(UBound(aFicha)).F4 = !FichaDaJrMul
        aFicha(UBound(aFicha)).F5 = !FichaDaEnca
        aFicha(UBound(aFicha)).F6 = !FichaAjuiza
        aFicha(UBound(aFicha)).F7 = !FichaAjJrMul
        aFicha(UBound(aFicha)).F8 = !FichaAjEnca
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub btBanco_Click()
Dim Sql As String, RdoAux As rdoResultset, nUserID As Integer, nPos As Long, nTot As Long
Dim qd As New rdoQuery, RdoDeb As rdoResultset
Set qd.ActiveConnection = cn

qd.QueryTimeout = 0
nPos = 0
Pb.value = 0
Me.Refresh
nUserID = RetornaUsuarioID(NomeDeLogin)
Sql = "delete from resumo_pagto_banco where userid=" & nUserID
cn.Execute Sql, rdExecDirect
Ocupado
Sql = "SELECT CODREDUZIDO,anoexercicio,codlancamento,seqlancamento,numparcela,codcomplemento,datapagamento,datarecebimento,valorpagoreal,debitopago.CODBANCO,numdocumento,arquivobanco,BANCO.nomebanco FROM debitopago "
Sql = Sql & "INNER JOIN banco ON debitopago.codbanco = banco.codbanco WHERE datarecebimento BETWEEN '" & Format(dtDataDe.value, "mm/dd/yyyy") & "' AND '" & Format(dtDataAte.value, "mm/dd/yyyy") & "' "
Sql = Sql & "ORDER BY datarecebimento,codbanco,numdocumento"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    
    Do Until .EOF
        CallPb nPos, nTot
        On Error Resume Next
        RdoDeb.Close
        On Error GoTo 0
        qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
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
                On Error Resume Next
                Sql = "insert resumo_pagto_banco(userid,datacredito,codbanco,numdocumento,codigo,ano,lanc,seq,parc,compl,codtributo,desctributo,nomebanco,arquivo,valorp,valorj,valorm,valorc,valort,valorpago) values("
                Sql = Sql & nUserID & ",'" & Format(RdoAux!datarecebimento, "mm/dd/yyyy") & "'," & RdoAux!CodBanco & "," & RdoAux!NumDocumento & "," & RdoAux!CODREDUZIDO & "," & RdoAux!AnoExercicio & ","
                Sql = Sql & RdoAux!CodLancamento & "," & RdoAux!SeqLancamento & "," & RdoAux!NumParcela & "," & RdoAux!CODCOMPLEMENTO & "," & !CodTributo & ",'" & !ABREVTRIBUTO & "','" & RdoAux!NomeBanco & "','" & RdoAux!arquivobanco & "',"
                Sql = Sql & Virg2Ponto(CStr(!ValorTributo)) & "," & Virg2Ponto(CStr(!ValorJuros)) & "," & Virg2Ponto(CStr(!ValorMulta)) & "," & Virg2Ponto(CStr(!ValorCorrecao)) & ","
                Sql = Sql & Virg2Ponto(CStr(!ValorTotal)) & "," & Virg2Ponto(CStr(RdoAux!valorpagoreal)) & ")"
                cn.Execute Sql, rdExecDirect
               .MoveNext
            Loop
           .Close
        End With
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With
Liberado
Pb.value = 0
lblPB.Caption = "0 %"
Me.Refresh

frmReport.ShowReport3 "Resumo_Pagto_Banco", frmMdi.HWND, Me.HWND
Sql = "delete from resumo_pagto_banco where userid=" & nUserID
cn.Execute Sql, rdExecDirect

End Sub

Private Sub btFicha_Click()
Dim Sql As String, RdoAux As rdoResultset, nUserID As Integer, nPos As Long, nTot As Long, nCodTributo As Integer, x As Integer

Ocupado
nUserID = RetornaUsuarioID(NomeDeLogin)
Sql = "delete from resumo_pagto_banco_ficha where userid=" & RetornaUsuarioID(NomeDeLogin)
cn.Execute Sql, rdExecDirect
CarregaTributo

Sql = "SELECT distinct numdocumento,datarecebimento from debitopago WHERE datarecebimento BETWEEN '" & Format(dtDataDe.value, "mm/dd/yyyy") & "' AND '" & Format(dtDataAte.value, "mm/dd/yyyy") & "' order by numdocumento"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    Do Until .EOF
        If nPos Mod 50 = 0 Then
            CallPb nPos, nTot
        End If
        ReDim aRegDoc(0)
        'aRegDoc = Retorna_Fichas_Documento(!Numdocumento)
        
        For x = 1 To UBound(aRegDoc)
            Sql = "insert resumo_pagto_banco_ficha(userid,datacredito,documento,codigo,codtributo,desctributo,ficha,valor) values(" & RetornaUsuarioID(NomeDeLogin) & ",'"
            Sql = Sql & Format(aRegDoc(x).sDataRecebimento, "mm/dd/yyyy") & "'," & aRegDoc(x).nNumDocumento & "," & aRegDoc(x).nCodReduz & "," & aRegDoc(x).nCodTributo & ",'"
            Sql = Sql & aRegDoc(x).sDescTributo & "'," & aRegDoc(x).nCodFicha & "," & Virg2Ponto(CStr(aRegDoc(x).nValor)) & ")"
            cn.Execute Sql, rdExecDirect
        Next
        
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With
Liberado
Pb.value = 0
lblPB.Caption = "0 %"
Me.Refresh



fim:
Sql = "delete from resumo_pagto_banco_ficha where userid=" & RetornaUsuarioID(NomeDeLogin)
'cn.Execute Sql, rdExecDirect
Liberado
Pb.value = 0
lblPB.Caption = "0 %"
Me.Refresh

End Sub

Private Sub cmdGerar_Click()
If NomeDeLogin = "SCHWARTZ" Then
    CarregaGrid3
    'TESTE
Else
    CarregaGrid
End If
End Sub

Private Sub Form_Load()
Dim RdoAux As rdoResultset, Sql As String

Centraliza Me
dpData.value = Now
dtDataDe.value = Now
dtDataAte.value = Now
cmbBanco.AddItem ("(Todos os Bancos)")
Sql = "SELECT CODBANCO,NOMEBANCO FROM BANCO WHERE CODBANCO<>0"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbBanco.AddItem !NomeBanco
        cmbBanco.ItemData(cmbBanco.NewIndex) = !CodBanco
       .MoveNext
    Loop
End With
cmbBanco.ListIndex = 0
cmbCC.ListIndex = 0

End Sub

Private Sub CarregaGrid2()
Dim Sql As String, RdoAux As rdoResultset, sDataLote As String, qd As New rdoQuery, RdoDeb As rdoResultset, xId As Long, nNumRec As Long
Dim nValorPago As Double, nValorTotalPago As Double, aReg() As Reg, x As Integer, nCodTrib As Integer, bDA As Boolean, bAj As Boolean, nValorDif As Double, bFind As Boolean
Dim z As Long, nCodFicha As Integer, bJuros As Boolean, bMulta As Boolean, bCorrecao As Boolean, nCodFichaP As Integer, nCodFichaJM As Integer, nCodFichaC As Integer
Dim nIndex As Integer, nIndex2 As Integer, aRegDoc() As Registros, RdoAux2 As rdoResultset, nValorTmp As Double, v As Integer, nUserID As Integer, aFichaValor() As FichaValor

ReDim aReg(0): ReDim aRegDoc(0)
sDataLote = Format(Now, "ddmmhhmm")
Set qd.ActiveConnection = cn
qd.QueryTimeout = 0
lblTotalDia.Caption = "0,00"
dpData.Enabled = False
opt1(0).Enabled = False
opt1(1).Enabled = False
cmbBanco.Enabled = False
cmdGerar.Enabled = False

nUserID = RetornaUsuarioID(NomeDeLogin)
Sql = "delete from resumo_pagto_banco_ficha where userid=" & RetornaUsuarioID(NomeDeLogin)
cn.Execute Sql, rdExecDirect
CarregaTributo

Ocupado

dDataReceita = dpData.value
If cmbBanco.ListIndex > 0 Then
    nCodBanco = cmbBanco.ItemData(cmbBanco.ListIndex)
Else
    nCodBanco = 0
End If

Sql = "SELECT SUM(debitopago.valorpagoreal) AS soma FROM debitoparcela INNER JOIN debitopago ON debitoparcela.codreduzido = debitopago.codreduzido AND "
Sql = Sql & "debitoparcela.anoexercicio = debitopago.anoexercicio AND debitoparcela.codlancamento = debitopago.codlancamento AND debitoparcela.seqlancamento = debitopago.seqlancamento AND "
Sql = Sql & "debitoparcela.numparcela = debitopago.numparcela AND debitoparcela.codcomplemento = debitopago.codcomplemento INNER JOIN numdocumento ON debitopago.numdocumento = numdocumento.numdocumento "
Sql = Sql & "WHERE RESTITUIDO IS NULL AND DATARECEBIMENTO = '" & Format(dDataReceita, "mm/dd/yyyy") & "' "
If opt1(0).value = True Then
    Sql = Sql & " AND (DEBITOPAGO.CODBANCO=90 or DEBITOPAGO.CODBANCO=91 OR DEBITOPAGO.CODBANCO=92 OR DEBITOPAGO.CODBANCO=93 OR DEBITOPAGO.CODBANCO=94 OR DEBITOPAGO.CODBANCO=95 OR DEBITOPAGO.CODBANCO=96 OR DEBITOPAGO.CODBANCO=97 OR DEBITOPAGO.CODBANCO=98) "
Else
    If nCodBanco > 0 Then
        Sql = Sql & " AND DEBITOPAGO.CODBANCO=" & nCodBanco
    Else
        Sql = Sql & " AND (DEBITOPAGO.CODBANCO not in(90,91,92,93,94,95,96,97,98))"
    End If
End If
If cmbCC.ListIndex = 0 Then
    Sql = Sql & " AND (DEBITOPAGO.CONTACORRENTE='0740004' OR DEBITOPAGO.CONTACORRENTE=NULL  OR DEBITOPAGO.CONTACORRENTE='' OR (CONVERT(int, debitopago.contacorrente) = 0)) AND debitopago.CODLANCAMENTO<>79"
ElseIf cmbCC.ListIndex = 1 Then
    Sql = Sql & " AND (DEBITOPAGO.CONTACORRENTE='0740004' OR DEBITOPAGO.CONTACORRENTE=NULL  OR DEBITOPAGO.CONTACORRENTE='' OR (CONVERT(int, debitopago.contacorrente) = 0)) AND debitopago.CODLANCAMENTO=79 "
ElseIf cmbCC.ListIndex = 2 Then
    Sql = Sql & " AND DEBITOPAGO.CONTACORRENTE='" & cmbCC.Text & "'"
End If

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!soma) Then
    lblTotalDia.Caption = Format(0, "#0.00")
    GoTo fim
Else
    lblTotalDia.Caption = Format(RdoAux!soma, "#0.00")
    nValorTotalPago = RdoAux!soma
End If
RdoAux.Close

Sql = "SELECT DISTINCT debitopago.codreduzido, debitopago.anoexercicio, debitopago.codlancamento, debitopago.seqlancamento, debitopago.numparcela, debitopago.codcomplemento, debitopago.seqpag, debitopago.datapagamento,"
Sql = Sql & "debitopago.datarecebimento, debitopago.valorpago, debitopago.codbanco, debitopago.codagencia, debitopago.restituido, debitopago.numdocumento, debitopago.valorpagoreal, debitopago.intacto, debitopago.valortarifa,"
Sql = Sql & "debitopago.arquivobanco , debitopago.valordif, debitopago.datapagamentocalc, debitopago.dataintegracao, debitopago.contacorrente, parceladocumento.plano, plano.desconto "
Sql = Sql & "FROM debitopago INNER JOIN parceladocumento ON debitopago.codreduzido = parceladocumento.codreduzido AND debitopago.anoexercicio = parceladocumento.anoexercicio AND debitopago.codlancamento = parceladocumento.codlancamento AND "
Sql = Sql & "debitopago.seqlancamento = parceladocumento.seqlancamento AND debitopago.numparcela = parceladocumento.numparcela AND debitopago.codcomplemento = parceladocumento.codcomplemento AND "
Sql = Sql & "debitopago.numdocumento = parceladocumento.numdocumento LEFT OUTER JOIN plano ON parceladocumento.plano = plano.codigo "
Sql = Sql & "WHERE RESTITUIDO IS NULL AND DATARECEBIMENTO = '" & Format(dDataReceita, "mm/dd/yyyy") & "' "
If opt1(0).value = True Then
    Sql = Sql & " AND (DEBITOPAGO.CODBANCO=90 or DEBITOPAGO.CODBANCO=91 OR DEBITOPAGO.CODBANCO=92 OR DEBITOPAGO.CODBANCO=93 OR DEBITOPAGO.CODBANCO=94 OR DEBITOPAGO.CODBANCO=95 OR DEBITOPAGO.CODBANCO=96 OR DEBITOPAGO.CODBANCO=97 OR DEBITOPAGO.CODBANCO=98) "
Else
    If nCodBanco > 0 Then
        Sql = Sql & " AND DEBITOPAGO.CODBANCO=" & nCodBanco
    Else
        Sql = Sql & " AND (DEBITOPAGO.CODBANCO not in(90,91,92,93,94,95,96,97,98))"
    End If
End If
If cmbCC.ListIndex = 0 Then
    Sql = Sql & " AND (DEBITOPAGO.CONTACORRENTE='0740004' OR DEBITOPAGO.CONTACORRENTE=NULL  OR DEBITOPAGO.CONTACORRENTE='' OR (CONVERT(int, debitopago.contacorrente) = 0)) AND debitopago.CODLANCAMENTO<>79"
ElseIf cmbCC.ListIndex = 1 Then
    Sql = Sql & " AND (DEBITOPAGO.CONTACORRENTE='0740004' OR DEBITOPAGO.CONTACORRENTE=NULL  OR DEBITOPAGO.CONTACORRENTE='' OR (CONVERT(int, debitopago.contacorrente) = 0)) AND debitopago.CODLANCAMENTO=79 "
Else
    Sql = Sql & " AND DEBITOPAGO.CONTACORRENTE='" & cmbCC.Text & "'"
End If
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nNumRec = .RowCount
    Do Until .EOF
    
        If IsNull(!desconto) Then
            nPercDesconto = 0
        Else
            nPercDesconto = !desconto
        End If
        nCodBanco = !CodBanco
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
        
        If xId Mod 50 = 0 Then
            CallPb xId, nNumRec
        End If
        
        On Error Resume Next
        RdoDeb.Close
        On Error GoTo 0
        qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
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
            If .RowCount > 0 Then
            nValorPago = !valorpagoreal
            Do Until .EOF
                If Format(!datarecebimento, "dd/mm/yyyy") <> Format(dDataReceita, "dd/mm/yyyy") Then
                    GoTo NextDoc
                End If
                nIndex = UBound(aReg) + 1
                ReDim Preserve aReg(nIndex)
                aReg(nIndex).nCodBanco = nCodBanco
                aReg(nIndex).sArquivo = RdoAux!arquivobanco
                aReg(nIndex).nCodReduz = !CODREDUZIDO
                aReg(nIndex).nAno = !AnoExercicio
                aReg(nIndex).nLanc = !CodLancamento
                aReg(nIndex).nSeq = !SeqLancamento
                aReg(nIndex).nParc = !NumParcela
                aReg(nIndex).nCompl = !CODCOMPLEMENTO
                aReg(nIndex).nCodTrib = !CodTributo
                aReg(nIndex).sDescTributo = !ABREVTRIBUTO
                aReg(nIndex).nValorTrib = !ValorTributo
                aReg(nIndex).NumDocumento = RdoAux!NumDocumento
                If !DataPagamento <= !DataVencimentoCalc Then
                    aReg(nIndex).nValorJuros = 0
                    aReg(nIndex).nValorMulta = 0
                    aReg(nIndex).nValorCorrecao = 0
                    aReg(nIndex).nValorTotal = !ValorTributo
                Else
                    aReg(nIndex).nValorJuros = !ValorJuros - (!ValorJuros * nPercDesconto / 100)
                    aReg(nIndex).nValorMulta = !ValorMulta - (!ValorMulta * nPercDesconto / 100)
                    aReg(nIndex).nValorCorrecao = !ValorCorrecao
                    aReg(nIndex).nValorTotal = aReg(nIndex).nValorTrib + aReg(nIndex).nValorJuros + aReg(nIndex).nValorMulta + aReg(nIndex).nValorCorrecao
                End If
                aReg(nIndex).sDataInscricao = IIf(IsDate(!datainscricao), "S", "N")
                aReg(nIndex).sDataRecebimento = Format(!datarecebimento, "dd/mm/yyyy")
                
                aReg(nIndex).sDataAjuiza = IIf(IsDate(!dataajuiza), "S", "N")
NextDoc:
               .MoveNext
            Loop
            End If
           .Close
        End With
        xId = xId + 1
       .MoveNext
    Loop
   .Close
End With


For nIndex = 1 To UBound(aReg)
'    If aReg(nIndex).NumDocumento = 17949428 Then MsgBox "teste"
    If nIndex Mod 10 = 0 Then
        CallPb CLng(nIndex), CLng(UBound(aReg))
    End If
    nCodTrib = aReg(nIndex).nCodTrib
    z = -1
    z = BinarySearchLong(aCodFicha(), CLng(nCodTrib))
    bDA = IIf(aReg(x).sDataInscricao = "S", True, False)
    bAj = IIf(aReg(x).sDataAjuiza = "S", True, False)
    If aReg(nIndex).nValorJuros > 0 Then bJuros = True
    If aReg(nIndex).nValorMulta > 0 Then bMulta = True
    If aReg(nIndex).nValorCorrecao > 0 Then bCorrecao = True
    If Not bDA And Not bAj Then
        nCodFichaP = aFicha(z).F1 'Principal
        If bJuros Or bMulta Then
            nCodFichaJM = aFicha(z).F2 'Juros e Multa normal
        End If
    End If
    If bDA And Not bAj Then
        nCodFichaP = aFicha(z).F3 'Principal DA
        If bJuros Or bMulta Then
            nCodFichaJM = aFicha(z).F4 'Juros e Multa DA
        End If
        If bCorrecao Then
            nCodFichaC = aFicha(z).F5 'Corre��o DA
        End If
    End If
    If bDA And bAj Then
        nCodFichaP = aFicha(z).F6 ' Principal Aj
        If bJuros Or bMulta Then
            nCodFichaJM = aFicha(z).F7 'Juros e Multa AJ
        End If
        If bCorrecao Then
            nCodFichaC = aFicha(z).F8 'Corre��o AJ
        End If
    End If
        
'    *** PRINCIPAL ****
    If nCodFichaP > 0 Then
        Sql = "SELECT NATUREZA,VINCULO,PERC,DESCTA FROM FICHACONTABIL WHERE FICHA=" & nCodFichaP
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                ReDim Preserve aRegDoc(UBound(aRegDoc) + 1)
                aRegDoc(UBound(aRegDoc)).nNumDocumento = aReg(nIndex).NumDocumento
                aRegDoc(UBound(aRegDoc)).sDataRecebimento = aReg(nIndex).sDataRecebimento
                aRegDoc(UBound(aRegDoc)).nCodReduz = aReg(nIndex).nCodReduz
                aRegDoc(UBound(aRegDoc)).nCodBanco = aReg(nIndex).nCodBanco
                aRegDoc(UBound(aRegDoc)).nAno = aReg(nIndex).nAno
                aRegDoc(UBound(aRegDoc)).nLanc = aReg(nIndex).nLanc
                aRegDoc(UBound(aRegDoc)).nSeq = aReg(nIndex).nSeq
                aRegDoc(UBound(aRegDoc)).nParc = aReg(nIndex).nParc
                aRegDoc(UBound(aRegDoc)).nCompl = aReg(nIndex).nCompl
                aRegDoc(UBound(aRegDoc)).nCodFicha = nCodFichaP
                aRegDoc(UBound(aRegDoc)).nCodTributo = aReg(nIndex).nCodTrib
                aRegDoc(UBound(aRegDoc)).sDescFicha = !descta
                aRegDoc(UBound(aRegDoc)).sDescTributo = aReg(nIndex).sDescTributo
                aRegDoc(UBound(aRegDoc)).sArquivo = aReg(nIndex).sArquivo
                aRegDoc(UBound(aRegDoc)).sNatureza = !Natureza
                aRegDoc(UBound(aRegDoc)).sVinculo = !Vinculo
                aRegDoc(UBound(aRegDoc)).nPerc = !Perc
                aRegDoc(UBound(aRegDoc)).nValor = aReg(nIndex).nValorTrib * !Perc / 100
                
               .MoveNext
            Loop
           .Close
        End With
    End If
'   *******************
    
'    *** juros e multa ****
    If nCodFichaJM > 0 Then
        Sql = "SELECT NATUREZA,VINCULO,PERC,DESCTA FROM FICHACONTABIL WHERE FICHA=" & nCodFichaJM
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                ReDim Preserve aRegDoc(UBound(aRegDoc) + 1)
                aRegDoc(UBound(aRegDoc)).nNumDocumento = aReg(nIndex).NumDocumento
                aRegDoc(UBound(aRegDoc)).sDataRecebimento = aReg(nIndex).sDataRecebimento
                aRegDoc(UBound(aRegDoc)).nCodReduz = aReg(nIndex).nCodReduz
                aRegDoc(UBound(aRegDoc)).nCodBanco = aReg(nIndex).nCodBanco
                aRegDoc(UBound(aRegDoc)).nAno = aReg(nIndex).nAno
                aRegDoc(UBound(aRegDoc)).nLanc = aReg(nIndex).nLanc
                aRegDoc(UBound(aRegDoc)).nSeq = aReg(nIndex).nSeq
                aRegDoc(UBound(aRegDoc)).nParc = aReg(nIndex).nParc
                aRegDoc(UBound(aRegDoc)).nCompl = aReg(nIndex).nCompl
                aRegDoc(UBound(aRegDoc)).nCodFicha = nCodFichaJM
                aRegDoc(UBound(aRegDoc)).nCodTributo = aReg(nIndex).nCodTrib
                aRegDoc(UBound(aRegDoc)).sDescFicha = !descta
                aRegDoc(UBound(aRegDoc)).sDescTributo = aReg(nIndex).sDescTributo
                aRegDoc(UBound(aRegDoc)).sArquivo = aReg(nIndex).sArquivo
                aRegDoc(UBound(aRegDoc)).sNatureza = !Natureza
                aRegDoc(UBound(aRegDoc)).sVinculo = !Vinculo
                aRegDoc(UBound(aRegDoc)).nPerc = !Perc
                nValorTmp = aReg(nIndex).nValorJuros + aReg(nIndex).nValorMulta
                aRegDoc(UBound(aRegDoc)).nValor = nValorTmp * !Perc / 100
                
               .MoveNext
            Loop
           .Close
        End With
    End If
'   *******************

'    *** corre��o ****
    If nCodFichaC > 0 Then
        Sql = "SELECT NATUREZA,VINCULO,PERC,DESCTA FROM FICHACONTABIL WHERE FICHA=" & nCodFichaC
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                ReDim Preserve aRegDoc(UBound(aRegDoc) + 1)
                aRegDoc(UBound(aRegDoc)).nNumDocumento = aReg(nIndex).NumDocumento
                aRegDoc(UBound(aRegDoc)).sDataRecebimento = aReg(nIndex).sDataRecebimento
                aRegDoc(UBound(aRegDoc)).nCodReduz = aReg(nIndex).nCodReduz
                aRegDoc(UBound(aRegDoc)).nCodBanco = aReg(nIndex).nCodBanco
                aRegDoc(UBound(aRegDoc)).nAno = aReg(nIndex).nAno
                aRegDoc(UBound(aRegDoc)).nLanc = aReg(nIndex).nLanc
                aRegDoc(UBound(aRegDoc)).nSeq = aReg(nIndex).nSeq
                aRegDoc(UBound(aRegDoc)).nParc = aReg(nIndex).nParc
                aRegDoc(UBound(aRegDoc)).nCompl = aReg(nIndex).nCompl
                aRegDoc(UBound(aRegDoc)).nCodFicha = nCodFichaC
                aRegDoc(UBound(aRegDoc)).nCodTributo = aReg(nIndex).nCodTrib
                aRegDoc(UBound(aRegDoc)).sDescFicha = !descta
                aRegDoc(UBound(aRegDoc)).sDescTributo = aReg(nIndex).sDescTributo
                aRegDoc(UBound(aRegDoc)).sArquivo = aReg(nIndex).sArquivo
                aRegDoc(UBound(aRegDoc)).sNatureza = !Natureza
                aRegDoc(UBound(aRegDoc)).sVinculo = !Vinculo
                aRegDoc(UBound(aRegDoc)).nPerc = !Perc
                aRegDoc(UBound(aRegDoc)).nValor = aReg(nIndex).nValorCorrecao * !Perc / 100
                
               .MoveNext
            Loop
           .Close
        End With
    End If
'   *******************
    
Next
    
For x = 1 To UBound(aRegDoc)
    If aRegDoc(x).nValor > 0 Then
        Sql = "insert resumo_pagto_banco_ficha(userid,datacredito,documento,codigo,ano,lanc,seq,parc,compl,codtributo,desctributo,descficha,ficha,arquivo,natureza,vinculo,perc,valor,codbanco,id) values("
        Sql = Sql & nUserID & ",'" & Format(aRegDoc(x).sDataRecebimento, "mm/dd/yyyy") & "'," & aRegDoc(x).nNumDocumento & "," & aRegDoc(x).nCodReduz & "," & aRegDoc(x).nAno & ","
        Sql = Sql & aRegDoc(x).nLanc & "," & aRegDoc(x).nSeq & "," & aRegDoc(x).nParc & "," & aRegDoc(x).nCompl & "," & aRegDoc(x).nCodTributo & ",'" & aRegDoc(x).sDescTributo & "','" & aRegDoc(x).sDescFicha & "',"
        Sql = Sql & aRegDoc(x).nCodFicha & ",'" & aRegDoc(x).sArquivo & "','" & aRegDoc(x).sNatureza & "','" & aRegDoc(x).sVinculo & "'," & aRegDoc(x).nPerc & "," & Virg2Ponto(CStr(aRegDoc(x).nValor)) & "," & aRegDoc(x).nCodBanco & "," & x & ")"
        cn.Execute Sql, rdExecDirect
    End If
Next


'ARREDONDAMENTO
Dim nValorPagoDoc As Double, nValorTotalFicha As Double
Sql = "SELECT DISTINCT documento FROM resumo_pagto_banco_ficha WHERE userid=" & nUserID & " ORDER BY documento"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        'If !Documento = 17818695 Then MsgBox "teste"
        Sql = "select sum(valorpagoreal) as totalpago from debitopago where numdocumento=" & !Documento
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        nValorPagoDoc = RdoAux2!totalpago
        RdoAux2.Close
        Sql = "select sum(valor) as totalficha from resumo_pagto_banco_ficha where documento=" & !Documento
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        nValorTotalFicha = RdoAux2!Totalficha
        RdoAux2.Close
        
        If nValorPagoDoc > nValorTotalFicha Then
            nValorDif = Abs(Round(nValorPagoDoc - nValorTotalFicha, 2))
            Sql = "select top(1) * from resumo_pagto_banco_ficha where userid=" & nUserID & " and documento=" & !Documento & " order by valor desc"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            Sql = "update resumo_pagto_banco_ficha set valor=valor+" & Virg2Ponto(CStr(nValorDif)) & " where userid=" & nUserID & " and id=" & RdoAux2!id
            cn.Execute Sql, rdExecDirect
'            Sql = "update resumo_pagto_banco_ficha set valor=valor+" & Virg2Ponto(CStr(nValorDif)) & " where userid=" & nUserID & " and documento=" & !Documento & " and "
'            Sql = Sql & "ficha=" & RdoAux2!Ficha & " and perc=" & RdoAux2!Perc & " and codtributo=" & RdoAux2!CodTributo & " and ano=" & RdoAux2!Ano & " and "
'            Sql = Sql & "lanc=" & RdoAux2!Lanc & " and seq=" & RdoAux2!Seq & " and parc=" & RdoAux2!Parc & " and compl=" & RdoAux2!Compl
'            cn.Execute Sql, rdExecDirect
            GoTo NextReg
        ElseIf nValorPagoDoc < nValorTotalFicha Then
            nValorDif = Abs(Round(nValorPagoDoc - nValorTotalFicha, 2))
            ReDim aFichaValor(0)
            Sql = "select * from resumo_pagto_banco_ficha where userid=" & nUserID & " and documento=" & !Documento & " order by valor desc"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                Do Until .EOF
                    ReDim Preserve aFichaValor(UBound(aFichaValor) + 1)
                    aFichaValor(UBound(aFichaValor)).nNumDocumento = !Documento
                    aFichaValor(UBound(aFichaValor)).nValorFicha = !Valor
                    aFichaValor(UBound(aFichaValor)).sDataCredito = Format(!DataCredito, "dd/mm/yyyy")
                    aFichaValor(UBound(aFichaValor)).nFicha = !Ficha
                    aFichaValor(UBound(aFichaValor)).nId = !id
                    aFichaValor(UBound(aFichaValor)).nPerc = !Perc
                    aFichaValor(UBound(aFichaValor)).nCodTributo = !CodTributo
                    aFichaValor(UBound(aFichaValor)).nAno = !Ano
                    aFichaValor(UBound(aFichaValor)).nLanc = !Lanc
                    aFichaValor(UBound(aFichaValor)).nSeq = !Seq
                    aFichaValor(UBound(aFichaValor)).nParc = !Parc
                    aFichaValor(UBound(aFichaValor)).nCompl = !COMPL
                   .MoveNext
                Loop
               .Close
            End With
            For x = 1 To UBound(aFichaValor)
                If aFichaValor(x).nValorFicha >= nValorDif Then
                    aFichaValor(x).nValorFicha = aFichaValor(x).nValorFicha - nValorDif
                    Exit For
                Else
                    nValorDif = (nValorDif - aFichaValor(x).nValorFicha)
                    aFichaValor(x).nValorFicha = 0
                End If
            Next
            
            For x = 1 To UBound(aFichaValor)
                Sql = "update resumo_pagto_banco_ficha set valor=" & Virg2Ponto(CStr(aFichaValor(x).nValorFicha))
                Sql = Sql & " where userid=" & nUserID & " and documento=" & !Documento & " and id=" & aFichaValor(x).nId
                cn.Execute Sql, rdExecDirect
            Next
            
            Sql = "delete from resumo_pagto_banco_ficha where userid=" & nUserID & " and documento=" & !Documento & " and valor=0"
            cn.Execute Sql, rdExecDirect
            
        End If
NextReg:
       .MoveNext
    Loop
   .Close
End With


Sql = "delete from resumo_pagto_banco_ficha where userid=" & nUserID
'cn.Execute Sql, rdExecDirect

MsgBox "fim"

fim:
Liberado
Pb.value = 0
lblPB.Caption = "0%"
dpData.Enabled = True
opt1(0).Enabled = True
opt1(1).Enabled = True

If opt1(1).value = True Then
    cmbBanco.Enabled = True
End If
cmdGerar.Enabled = True

End Sub


Private Sub CarregaGrid()
Dim Sql As String, RdoAux As rdoResultset, nNumRec As Long, itmX As ListItem, z As Long, nCodBanco As Integer, xId As Long
Dim qd As New rdoQuery, RdoDeb As rdoResultset, x As Integer, nCodTrib As Integer, bDA As Boolean, bAj As Boolean
Dim nValorTotal As Double, nValorPago As Double, nValorTotalPago As Double, nValorCompensado As Double, dDataReceita As Date
Dim nMaiorValor As Double, nIndMaior As Integer, nTotalMatriz As Double, nDif As Double, aReg() As Reg, nIndex As Integer
Dim sDataLote As String, nPercDesconto As Double

sDataLote = Format(Now, "ddmmhhmm")
ReDim aReg(0)
Sql = "delete from analise2 where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

lblTotalDia.Caption = "0,00"
dpData.Enabled = False
opt1(0).Enabled = False
opt1(1).Enabled = False
cmbBanco.Enabled = False
cmdGerar.Enabled = False

CarregaTributo
ReDim aFichaDetalhe(0)

z = SendMessage(lvMain.HWND, LVM_DELETEALLITEMS, 0, 0)
Set qd.ActiveConnection = cn
qd.QueryTimeout = 0
xId = 1
dDataReceita = dpData.value
If cmbBanco.ListIndex > 0 Then
    nCodBanco = cmbBanco.ItemData(cmbBanco.ListIndex)
Else
    nCodBanco = 0
End If

Sql = "SELECT SUM(debitopago.valorpagoreal) AS soma FROM debitoparcela INNER JOIN debitopago ON debitoparcela.codreduzido = debitopago.codreduzido AND "
Sql = Sql & "debitoparcela.anoexercicio = debitopago.anoexercicio AND debitoparcela.codlancamento = debitopago.codlancamento AND debitoparcela.seqlancamento = debitopago.seqlancamento AND "
Sql = Sql & "debitoparcela.numparcela = debitopago.numparcela AND debitoparcela.codcomplemento = debitopago.codcomplemento INNER JOIN numdocumento ON debitopago.numdocumento = numdocumento.numdocumento "
Sql = Sql & "WHERE RESTITUIDO IS NULL AND DATARECEBIMENTO = '" & Format(dDataReceita, "mm/dd/yyyy") & "' "
If opt1(0).value = True Then
    Sql = Sql & " AND (DEBITOPAGO.CODBANCO=90 or DEBITOPAGO.CODBANCO=91 OR DEBITOPAGO.CODBANCO=92 OR DEBITOPAGO.CODBANCO=93 OR DEBITOPAGO.CODBANCO=94 OR DEBITOPAGO.CODBANCO=95 OR DEBITOPAGO.CODBANCO=96 OR DEBITOPAGO.CODBANCO=97 OR DEBITOPAGO.CODBANCO=98) "
Else
    If nCodBanco > 0 Then
        Sql = Sql & " AND DEBITOPAGO.CODBANCO=" & nCodBanco
    Else
        Sql = Sql & " AND (DEBITOPAGO.CODBANCO not in(90,91,92,93,94,95,96,97,98))"
    End If
End If
If cmbCC.ListIndex = 0 Then
    Sql = Sql & " AND (DEBITOPAGO.CONTACORRENTE='0740004' OR DEBITOPAGO.CONTACORRENTE=NULL  OR DEBITOPAGO.CONTACORRENTE='' OR (CONVERT(int, debitopago.contacorrente) = 0)) AND debitopago.CODLANCAMENTO<>79"
ElseIf cmbCC.ListIndex = 1 Then
    Sql = Sql & " AND (DEBITOPAGO.CONTACORRENTE='0740004' OR DEBITOPAGO.CONTACORRENTE=NULL  OR DEBITOPAGO.CONTACORRENTE='' OR (CONVERT(int, debitopago.contacorrente) = 0)) AND debitopago.CODLANCAMENTO=79 "
Else
    Sql = Sql & " AND DEBITOPAGO.CONTACORRENTE='" & cmbCC.Text & "'"
End If

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!soma) Then
    lblTotalDia.Caption = Format(0, "#0.00")
    GoTo fim
Else
    lblTotalDia.Caption = Format(RdoAux!soma, "#0.00")
    nValorTotalPago = RdoAux!soma
End If
RdoAux.Close
Ocupado

Sql = "SELECT DISTINCT debitopago.codreduzido, debitopago.anoexercicio, debitopago.codlancamento, debitopago.seqlancamento, debitopago.numparcela, debitopago.codcomplemento, debitopago.seqpag, debitopago.datapagamento,"
Sql = Sql & "debitopago.datarecebimento, debitopago.valorpago, debitopago.codbanco, debitopago.codagencia, debitopago.restituido, debitopago.numdocumento, debitopago.valorpagoreal, debitopago.intacto, debitopago.valortarifa,"
Sql = Sql & "debitopago.arquivobanco , debitopago.valordif, debitopago.datapagamentocalc, debitopago.dataintegracao, debitopago.contacorrente, parceladocumento.plano, plano.desconto "
Sql = Sql & "FROM debitopago INNER JOIN parceladocumento ON debitopago.codreduzido = parceladocumento.codreduzido AND debitopago.anoexercicio = parceladocumento.anoexercicio AND debitopago.codlancamento = parceladocumento.codlancamento AND "
Sql = Sql & "debitopago.seqlancamento = parceladocumento.seqlancamento AND debitopago.numparcela = parceladocumento.numparcela AND debitopago.codcomplemento = parceladocumento.codcomplemento AND "
Sql = Sql & "debitopago.numdocumento = parceladocumento.numdocumento LEFT OUTER JOIN plano ON parceladocumento.plano = plano.codigo "
Sql = Sql & "WHERE RESTITUIDO IS NULL AND DATARECEBIMENTO = '" & Format(dDataReceita, "mm/dd/yyyy") & "' "
If opt1(0).value = True Then
    Sql = Sql & " AND (DEBITOPAGO.CODBANCO=90 or DEBITOPAGO.CODBANCO=91 OR DEBITOPAGO.CODBANCO=92 OR DEBITOPAGO.CODBANCO=93 OR DEBITOPAGO.CODBANCO=94 OR DEBITOPAGO.CODBANCO=95 OR DEBITOPAGO.CODBANCO=96 OR DEBITOPAGO.CODBANCO=97 OR DEBITOPAGO.CODBANCO=98) "
Else
    If nCodBanco > 0 Then
        Sql = Sql & " AND DEBITOPAGO.CODBANCO=" & nCodBanco
    Else
        Sql = Sql & " AND (DEBITOPAGO.CODBANCO not in(90,91,92,93,94,95,96,97,98))"
    End If
End If
If cmbCC.ListIndex = 0 Then
    Sql = Sql & " AND (DEBITOPAGO.CONTACORRENTE='0740004' OR DEBITOPAGO.CONTACORRENTE=NULL  OR DEBITOPAGO.CONTACORRENTE='' OR (CONVERT(int, debitopago.contacorrente) = 0)) AND debitopago.CODLANCAMENTO<>79"
ElseIf cmbCC.ListIndex = 1 Then
    Sql = Sql & " AND (DEBITOPAGO.CONTACORRENTE='0740004' OR DEBITOPAGO.CONTACORRENTE=NULL  OR DEBITOPAGO.CONTACORRENTE='' OR (CONVERT(int, debitopago.contacorrente) = 0)) AND debitopago.CODLANCAMENTO=79 "
Else
    Sql = Sql & " AND DEBITOPAGO.CONTACORRENTE='" & cmbCC.Text & "'"
End If
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux

    nNumRec = .RowCount
    Do Until .EOF
        If IsNull(!desconto) Then
            nPercDesconto = 0
        Else
            nPercDesconto = !desconto
        End If
        nCodBanco = !CodBanco
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
        
        If xId Mod 50 = 0 Then
            CallPb xId, nNumRec
        End If
'        If !CODREDUZIDO = 101326 Then MsgBox "teste"
        On Error Resume Next
        RdoDeb.Close
        On Error GoTo 0
        qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
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
            If .RowCount > 0 Then
            nValorPago = !valorpagoreal
        '    If nValorPago = 50000 Then MsgBox "teste"
            Do Until .EOF
                nIndex = UBound(aReg) + 1
                'If !CODREDUZIDO = 120920 Then MsgBox "teste"
                ReDim Preserve aReg(nIndex)
                aReg(nIndex).nCodBanco = nCodBanco
                aReg(nIndex).nCodReduz = Format(!CODREDUZIDO, "000000")
                aReg(nIndex).nAno = !AnoExercicio
                aReg(nIndex).nLanc = !CodLancamento
                aReg(nIndex).nSeq = !SeqLancamento
                aReg(nIndex).nParc = !NumParcela
                aReg(nIndex).nCompl = !CODCOMPLEMENTO
                aReg(nIndex).nCodTrib = Format(!CodTributo, "000")
                aReg(nIndex).nValorTrib = Format(!ValorTributo, "#0.00")
                If !DataPagamento < !DataVencimentoCalc Then
                    aReg(nIndex).nValorJuros = Format(0, "#0.00")
                    aReg(nIndex).nValorMulta = Format(0, "#0.00")
                    aReg(nIndex).nValorCorrecao = Format(0, "#0.00")
                    aReg(nIndex).nValorTotal = Format(!ValorTributo, "#0.00")
                Else
'                   aReg(nIndex).nValorJuros = Format(!ValorJuros, "#0.00")
'                    aReg(nIndex).nValorMulta = Format(!ValorMulta, "#0.00")
                    aReg(nIndex).nValorJuros = !ValorJuros - (!ValorJuros * nPercDesconto / 100)
                    aReg(nIndex).nValorMulta = !ValorMulta - (!ValorMulta * nPercDesconto / 100)
                    aReg(nIndex).nValorCorrecao = !ValorCorrecao
                    aReg(nIndex).nValorTotal = aReg(nIndex).nValorTrib + aReg(nIndex).nValorJuros + aReg(nIndex).nValorMulta + aReg(nIndex).nValorCorrecao
                End If
                
                aReg(nIndex).sDataInscricao = IIf(IsDate(!datainscricao), "S", "N")
                aReg(nIndex).sDataAjuiza = IIf(IsDate(!dataajuiza), "S", "N")
                
               .MoveNext
            Loop
            End If
           .Close
           
        End With
        
        xId = xId + 1
       .MoveNext
    Loop
   .Close
End With
CallPb 100, 100

For x = 1 To UBound(aReg)
    nCodTrib = aReg(x).nCodTrib
'    If aReg(x).Numdocumento = 17955066 Then
'        MsgBox "teste"
'    End If
    'If aReg(x).nCodReduz = 632603 Then MsgBox "teste"
    bDA = IIf(aReg(x).sDataInscricao = "S", True, False)
    bAj = IIf(aReg(x).sDataAjuiza = "S", True, False)
        
    If x Mod 10 = 0 Then
        CallPb CLng(x), CLng(UBound(aReg))
    End If
    If nValorTotalPago >= CDbl(aReg(x).nValorTotal) Then
        aReg(x).ValorPago = aReg(x).nValorTotal
        aReg(x).ValorPr = aReg(x).nValorTrib
        aReg(x).ValorJr = aReg(x).nValorJuros
        aReg(x).ValorMl = aReg(x).nValorMulta
        aReg(x).ValorCr = aReg(x).nValorCorrecao
        GoTo Ficha
    Else
        If nValorTotalPago <= 0 Then
            aReg(x).ValorPago = 0
            GoTo Proximo
        Else
            aReg(x).ValorPago = Format(nValorTotalPago, "#0.00")
            aReg(x).ValorPr = Format(nValorTotalPago, "#0.00")
            nValorTotalPago = nValorTotalPago - aReg(x).ValorPr
'            If aReg(x).ValorJr = "" Then aReg(x).ValorJr = 0
'            If lvMain.ListItems(x).SubItems(18) = "" Then lvMain.ListItems(x).SubItems(18) = "0"
'            If lvMain.ListItems(x).SubItems(19) = "" Then lvMain.ListItems(x).SubItems(19) = "0"
            If nValorTotalPago >= CDbl(aReg(x).ValorJr) Then
                aReg(x).ValorJr = Format(nValorTotalPago, "#0.00")
                nValorTotalPago = nValorTotalPago - aReg(x).ValorJr
            Else
                aReg(x).ValorJr = Format(0, "#0.00")
            End If
            If nValorTotalPago >= CDbl(aReg(x).ValorMl) Then
                aReg(x).ValorMl = Format(nValorTotalPago, "#0.00")
                nValorTotalPago = nValorTotalPago - aReg(x).ValorMl
            Else
                aReg(x).ValorMl = Format(0, "#0.00")
            End If
            If nValorTotalPago >= 0 Then
                aReg(x).ValorCr = Format(nValorTotalPago, "#0.00")
            Else
                aReg(x).ValorCr = Format(0, "#0.00")
            End If
            GoTo Ficha
        End If
    End If
Ficha:
   ' If aReg(x).nCodReduz = 632603 Then MsgBox "teste"
    If aReg(x).ValorPago = 0 Then GoTo Proximo
    
    z = -1
    z = BinarySearchLong(aCodFicha(), CLng(nCodTrib))
    nCodBanco = aReg(x).nCodBanco
'    If aReg(x).nAno = 2020 Then
'        MsgBox "teste"
'    End If
    
'    If aReg(x).nCodTrib = 505 Then
'        MsgBox "teste"
'    End If
    If aReg(x).nValorTrib > 0 Then
        If Not bDA And Not bAj Then
            aReg(x).F1 = aFicha(z).F1 'Principal
            If (aReg(x).nCodTrib = 1 Or aReg(x).nCodTrib = 2) And aReg(x).nAno = 2020 And Year(dDataReceita) = 2019 Then
                aReg(x).F1 = 50513 'iptu
                FichaDetalhe dDataReceita, nCodBanco, 50513, aReg(x).nValorTrib
            ElseIf (aReg(x).nCodTrib = 14) And aReg(x).nAno = 2020 And Year(dDataReceita) = 2019 Then
                aReg(x).F1 = 50514 'tx lic
                FichaDetalhe dDataReceita, nCodBanco, 50514, aReg(x).nValorTrib
            ElseIf (aReg(x).nCodTrib = 11) And aReg(x).nAno = 2020 And Year(dDataReceita) = 2019 Then
                aReg(x).F1 = 50514 'iss
                FichaDetalhe dDataReceita, nCodBanco, 50514, aReg(x).nValorTrib
            ElseIf (aReg(x).nCodTrib = 25) And aReg(x).nAno = 2020 And Year(dDataReceita) = 2019 Then
                aReg(x).F1 = 50509 'vig
                FichaDetalhe dDataReceita, nCodBanco, 50509, aReg(x).nValorTrib
            Else
                FichaDetalhe dDataReceita, nCodBanco, aFicha(z).F1, aReg(x).nValorTrib
            End If
        ElseIf bDA And Not bAj Then
            aReg(x).F1 = aFicha(z).F3 'Principal DA
            FichaDetalhe dDataReceita, nCodBanco, aFicha(z).F3, aReg(x).nValorTrib
        ElseIf bDA And bAj Then
            aReg(x).F1 = aFicha(z).F6 ' Principal Aj
            FichaDetalhe dDataReceita, nCodBanco, aFicha(z).F6, aReg(x).nValorTrib
        End If
    End If
    
    If (aReg(x).nValorJuros > 0 Or aReg(x).nValorMulta > 0) Then
        If Not bDA And Not bAj Then
            aReg(x).F2 = aFicha(z).F2 'Juros e Multa
            FichaDetalhe dDataReceita, nCodBanco, aFicha(z).F2, aReg(x).nValorJuros + aReg(x).nValorMulta
        ElseIf bDA And Not bAj Then
            aReg(x).F2 = aFicha(z).F4 'Juros e Multa DA
            FichaDetalhe dDataReceita, nCodBanco, aFicha(z).F4, aReg(x).nValorJuros + aReg(x).nValorMulta
        ElseIf bDA And bAj Then
            aReg(x).F2 = aFicha(z).F7 'Juros e Multa Aj
            FichaDetalhe dDataReceita, nCodBanco, aFicha(z).F7, aReg(x).nValorJuros + aReg(x).nValorMulta
        End If
    End If
    
    If aReg(x).nValorCorrecao > 0 Then
        If Not bDA And Not bAj Then
           aReg(x).F3 = aFicha(z).F5 'Correcao
            FichaDetalhe dDataReceita, nCodBanco, aFicha(z).F5, aReg(x).nValorCorrecao
        ElseIf bDA And Not bAj Then
            aReg(x).F3 = aFicha(z).F5 'Correcao DA
            FichaDetalhe dDataReceita, nCodBanco, aFicha(z).F5, aReg(x).nValorCorrecao
        ElseIf bDA And bAj Then
            aReg(x).F3 = aFicha(z).F8 'Correcao Aj
            FichaDetalhe dDataReceita, nCodBanco, aFicha(z).F8, aReg(x).nValorCorrecao
        End If
    End If
    
    
    
    
Proximo:
    If aReg(x).nValorTotal < 0 Then MsgBox "teste"
    nValorTotalPago = nValorTotalPago - aReg(x).nValorTotal
Next
Pb.value = 100
lblPB.Caption = "100%"


'Arredondamento

nTotalMatriz = 0
nMaiorValor = 0
For x = 1 To UBound(aFichaDetalhe)
   nTotalMatriz = nTotalMatriz + aFichaDetalhe(x).Total
   If aFichaDetalhe(x).Ficha < 50000 And aFichaDetalhe(x).Total > nMaiorValor Then
      nMaiorValor = aFichaDetalhe(x).Total
      nIndMaior = x
   End If
Next

nDif = nTotalMatriz - CDbl(lblTotalDia.Caption)

If Round(nDif, 2) > 0 Then
    aFichaDetalhe(nIndMaior).Total = aFichaDetalhe(nIndMaior).Total - nDif
ElseIf Round(nDif, 2) < 0 Then
    aFichaDetalhe(nIndMaior).Total = aFichaDetalhe(nIndMaior).Total + Abs(nDif)
End If

If cmbBanco.ListIndex = 0 Then
   Open sPathBin & "\REC" & Format(Day(Now), "00") & Format(Month(Now), "00") For Output Shared As #1
Else
   Open sPathBin & "\REC" & Format(Day(Now), "00") & Format(Month(Now), "00") & Format(cmbBanco.ItemData(cmbBanco.ListIndex), "000") For Output Shared As #1
End If

For x = 1 To UBound(aFichaDetalhe)
    With aFichaDetalhe(x)
        Sql = "insert analise2 (usuario,datareceita,codbanco,valortotal,numficha,natureza,vinculo,perc,descficha) values('"
        Sql = Sql & NomeDeLogin & "','" & Format(dDataReceita, "mm/dd/yyyy") & "'," & .Banco & "," & Virg2Ponto(CStr(.Total)) & ","
        Sql = Sql & .Ficha & ",'" & .Natureza & "','" & .Vinculo & "'," & Virg2Ponto(CStr(.Perc)) & ",'" & .Descricao & "')"
        cn.Execute Sql, rdExecDirect
        
        ax = FillSpace(.Natureza, 20) & FillSpace(.Vinculo, 20) & Year(.Data) & Format(Month(.Data), "00") & Format(Day(.Data), "00") & Format(.Banco, "0000") & "00000000" & Format(RemovePonto(Replace(CStr(FormatNumber(.Total, 2)), ",", "")), "0000000000000") & "0000000000" & sDataLote
        Print #1, ax
    End With
Next
Close #1

fim:
Liberado
If UBound(aReg) = 0 Then
    MsgBox "N�o existem baixas neste per�odo para este(s) banco(s).", vbInformation, "Aten��o"
Else
    If frmMdi.frTeste.Visible = True Then
        frmReport.ShowReport "Analise2_tmp", frmMdi.HWND, Me.HWND
    Else
        frmReport.ShowReport "Analise2", frmMdi.HWND, Me.HWND
    End If
    Sql = "delete from analise2 where usuario='" & NomeDeLogin & "'"
    cn.Execute Sql, rdExecDirect
End If
dpData.Enabled = True
opt1(0).Enabled = True
opt1(1).Enabled = True

If opt1(1).value = True Then
    cmbBanco.Enabled = True
End If
cmdGerar.Enabled = True




End Sub

Private Sub CarregaGridOld()
Dim Sql As String, RdoAux As rdoResultset, nNumRec As Long, itmX As ListItem, z As Long, nCodBanco As Integer, xId As Long
Dim qd As New rdoQuery, RdoDeb As rdoResultset, x As Integer, nCodTrib As Integer, bDA As Boolean, bAj As Boolean
Dim nValorTotal As Double, nValorPago As Double, nValorTotalPago As Double, nValorCompensado As Double, dDataReceita As Date
Dim nMaiorValor As Double, nIndMaior As Integer, nTotalMatriz As Double, nDif As Double

Sql = "delete from analise2 where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

lblTotalDia.Caption = "0,00"
dpData.Enabled = False
opt1(0).Enabled = False
opt1(1).Enabled = False
cmbBanco.Enabled = False
cmdGerar.Enabled = False

CarregaTributo
ReDim aFichaDetalhe(0)

z = SendMessage(lvMain.HWND, LVM_DELETEALLITEMS, 0, 0)
Set qd.ActiveConnection = cn
qd.QueryTimeout = 0
xId = 1
dDataReceita = dpData.value
If cmbBanco.ListIndex > 0 Then
    nCodBanco = cmbBanco.ItemData(cmbBanco.ListIndex)
Else
    nCodBanco = 0
End If


Sql = "SELECT SUM(debitopago.valorpagoreal) AS soma FROM debitoparcela INNER JOIN debitopago ON debitoparcela.codreduzido = debitopago.codreduzido AND "
Sql = Sql & "debitoparcela.anoexercicio = debitopago.anoexercicio AND debitoparcela.codlancamento = debitopago.codlancamento AND debitoparcela.seqlancamento = debitopago.seqlancamento AND "
Sql = Sql & "debitoparcela.numparcela = debitopago.numparcela AND debitoparcela.codcomplemento = debitopago.codcomplemento INNER JOIN numdocumento ON debitopago.numdocumento = numdocumento.numdocumento "
Sql = Sql & "WHERE RESTITUIDO IS NULL AND DATARECEBIMENTO = '" & Format(dDataReceita, "mm/dd/yyyy") & "' "
If opt1(0).value = True Then
    Sql = Sql & " AND (DEBITOPAGO.CODBANCO=91 OR DEBITOPAGO.CODBANCO=92 OR DEBITOPAGO.CODBANCO=93 OR DEBITOPAGO.CODBANCO=94 OR DEBITOPAGO.CODBANCO=95 OR DEBITOPAGO.CODBANCO=96 OR DEBITOPAGO.CODBANCO=97 OR DEBITOPAGO.CODBANCO=98) "
Else
    If nCodBanco > 0 Then
        Sql = Sql & " AND DEBITOPAGO.CODBANCO=" & nCodBanco
    Else
        Sql = Sql & " AND (DEBITOPAGO.CODBANCO not in(91,92,93,94,95,96,97,98))"
    End If
End If
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!soma) Then
    lblTotalDia.Caption = Format(0, "#0.00")
    GoTo fim
Else
    lblTotalDia.Caption = Format(RdoAux!soma, "#0.00")
    nValorTotalPago = RdoAux!soma
End If
RdoAux.Close
Ocupado
Sql = "SELECT DISTINCT debitoparcela.codreduzido, debitoparcela.anoexercicio, debitoparcela.codlancamento, debitoparcela.seqlancamento, debitoparcela.numparcela,"
Sql = Sql & "debitoparcela.codcomplemento, debitoparcela.datainscricao, debitoparcela.dataajuiza,debitopago.datapagamento , debitopago.datarecebimento, debitopago.valorpagoreal,"
Sql = Sql & "debitopago.CodBanco , NumDocumento.valortaxadoc FROM debitoparcela INNER JOIN debitopago ON debitoparcela.codreduzido = debitopago.codreduzido AND "
Sql = Sql & "debitoparcela.anoexercicio = debitopago.anoexercicio AND debitoparcela.codlancamento = debitopago.codlancamento AND debitoparcela.seqlancamento = debitopago.seqlancamento AND "
Sql = Sql & "debitoparcela.numparcela = debitopago.numparcela AND debitoparcela.codcomplemento = debitopago.codcomplemento INNER JOIN numdocumento ON debitopago.numdocumento = numdocumento.numdocumento "
Sql = Sql & "WHERE RESTITUIDO IS NULL AND DATARECEBIMENTO = '" & Format(dDataReceita, "mm/dd/yyyy") & "' "
If opt1(0).value = True Then
    Sql = Sql & " AND (DEBITOPAGO.CODBANCO=91 OR DEBITOPAGO.CODBANCO=92 OR DEBITOPAGO.CODBANCO=93 OR DEBITOPAGO.CODBANCO=94 OR DEBITOPAGO.CODBANCO=95 OR DEBITOPAGO.CODBANCO=96 OR DEBITOPAGO.CODBANCO=97 OR DEBITOPAGO.CODBANCO=98) "
Else
    If nCodBanco > 0 Then
        Sql = Sql & " AND DEBITOPAGO.CODBANCO=" & nCodBanco
    Else
        Sql = Sql & " AND (DEBITOPAGO.CODBANCO not in(91,92,93,94,95,96,97,98))"
    End If
End If
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nNumRec = .RowCount
    Do Until .EOF
        nCodBanco = !CodBanco
        If nCodBanco > 90 And nCodBanco < 99 Then
            If nCodBanco = 91 Then
                nCodBanco = 1
            ElseIf nCodBanco = 92 Then nCodBanco = 33
            ElseIf nCodBanco = 93 Then nCodBanco = 237
            ElseIf nCodBanco = 94 Then nCodBanco = 341
            ElseIf nCodBanco = 95 Then nCodBanco = 409
            ElseIf nCodBanco = 96 Then nCodBanco = 151
            ElseIf nCodBanco = 97 Then nCodBanco = 104
            ElseIf nCodBanco = 98 Then nCodBanco = 399
            End If
        End If
        
        If xId Mod 50 = 0 Then
            CallPb xId, nNumRec
        End If
'        If !CODREDUZIDO = 7600 Then MsgBox "teste"
        On Error Resume Next
        RdoDeb.Close
        On Error GoTo 0
        qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
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
            nValorPago = !valorpagoreal
            Do Until .EOF
                Set itmX = lvMain.ListItems.Add(, , Format(nCodBanco, "000"))
                itmX.SubItems(1) = Format(!CODREDUZIDO, "000000")
                itmX.SubItems(2) = !AnoExercicio
                itmX.SubItems(3) = Format(!CodLancamento, "00")
                itmX.SubItems(4) = Format(!SeqLancamento, "000")
                itmX.SubItems(5) = Format(!NumParcela, "000")
                itmX.SubItems(6) = !CODCOMPLEMENTO
                itmX.SubItems(7) = Format(!CodTributo, "000")
                itmX.SubItems(8) = Format(!ValorTributo, "#0.00")
                If !DataPagamento < !DataVencimentoCalc Then
                    itmX.SubItems(9) = Format(0, "#0.00")
                    itmX.SubItems(10) = Format(0, "#0.00")
                    itmX.SubItems(11) = Format(0, "#0.00")
                    itmX.SubItems(12) = Format(!ValorTributo, "#0.00")
                Else
                    itmX.SubItems(9) = Format(!ValorJuros, "#0.00")
                    itmX.SubItems(10) = Format(!ValorMulta, "#0.00")
                    itmX.SubItems(11) = Format(!ValorCorrecao, "#0.00")
                    itmX.SubItems(12) = Format(!ValorTotal, "#0.00")
                End If
                itmX.SubItems(13) = IIf(IsDate(!datainscricao), "S", "N")
                itmX.SubItems(14) = IIf(IsDate(!dataajuiza), "S", "N")
               .MoveNext
            Loop
           .Close
        End With
        If !ValorTaxaDoc > 0 Then
            Set itmX = lvMain.ListItems.Add(, , Format(nCodBanco, "000"))
            itmX.SubItems(1) = Format(!CODREDUZIDO, "000000")
            itmX.SubItems(2) = !AnoExercicio
            itmX.SubItems(3) = Format(!CodLancamento, "00")
            itmX.SubItems(4) = Format(!SeqLancamento, "000")
            itmX.SubItems(5) = Format(!NumParcela, "000")
            itmX.SubItems(6) = !CODCOMPLEMENTO
            itmX.SubItems(7) = Format(3, "000")
            itmX.SubItems(8) = Format(!ValorTaxaDoc, "#0.00")
            itmX.SubItems(9) = Format(0, "#0.00")
            itmX.SubItems(10) = Format(0, "#0.00")
            itmX.SubItems(11) = Format(0, "#0.00")
            itmX.SubItems(12) = Format(!ValorTaxaDoc, "#0.00")
            itmX.SubItems(13) = IIf(IsDate(!datainscricao), "S", "N")
            itmX.SubItems(14) = IIf(IsDate(!dataajuiza), "S", "N")
        End If
        
        xId = xId + 1
       .MoveNext
    Loop
   .Close
End With
CallPb 100, 100

For x = 1 To lvMain.ListItems.Count
    nCodTrib = Val(lvMain.ListItems(x).SubItems(7))
    bDA = IIf(lvMain.ListItems(x).SubItems(13) = "S", True, False)
    bAj = IIf(lvMain.ListItems(x).SubItems(14) = "S", True, False)
        
    If x Mod 10 = 0 Then
        CallPb CLng(x), CLng(lvMain.ListItems.Count)
    End If
    If nValorTotalPago >= CDbl(lvMain.ListItems(x).SubItems(12)) Then
        lvMain.ListItems(x).SubItems(15) = lvMain.ListItems(x).SubItems(12)
        lvMain.ListItems(x).SubItems(16) = lvMain.ListItems(x).SubItems(8)
        lvMain.ListItems(x).SubItems(17) = lvMain.ListItems(x).SubItems(9)
        lvMain.ListItems(x).SubItems(18) = lvMain.ListItems(x).SubItems(10)
        lvMain.ListItems(x).SubItems(19) = lvMain.ListItems(x).SubItems(11)
        GoTo Ficha
    Else
        If nValorTotalPago <= 0 Then
            lvMain.ListItems(x).SubItems(15) = 0
            GoTo Proximo
        Else
            lvMain.ListItems(x).SubItems(15) = Format(nValorTotalPago, "#0.00")
            lvMain.ListItems(x).SubItems(16) = Format(nValorTotalPago, "#0.00")
            nValorTotalPago = nValorTotalPago - lvMain.ListItems(x).SubItems(16)
            If lvMain.ListItems(x).SubItems(17) = "" Then lvMain.ListItems(x).SubItems(17) = "0"
            If lvMain.ListItems(x).SubItems(18) = "" Then lvMain.ListItems(x).SubItems(18) = "0"
            If lvMain.ListItems(x).SubItems(19) = "" Then lvMain.ListItems(x).SubItems(19) = "0"
            If nValorTotalPago >= CDbl(lvMain.ListItems(x).SubItems(17)) Then
                lvMain.ListItems(x).SubItems(17) = Format(nValorTotalPago, "#0.00")
                nValorTotalPago = nValorTotalPago - lvMain.ListItems(x).SubItems(17)
            Else
                lvMain.ListItems(x).SubItems(17) = Format(0, "#0.00")
            End If
            If nValorTotalPago >= CDbl(lvMain.ListItems(x).SubItems(18)) Then
                lvMain.ListItems(x).SubItems(18) = Format(nValorTotalPago, "#0.00")
                nValorTotalPago = nValorTotalPago - lvMain.ListItems(x).SubItems(18)
            Else
                lvMain.ListItems(x).SubItems(18) = Format(0, "#0.00")
            End If
            If nValorTotalPago >= 0 Then
                lvMain.ListItems(x).SubItems(19) = Format(nValorTotalPago, "#0.00")
            Else
                lvMain.ListItems(x).SubItems(19) = Format(0, "#0.00")
            End If
            GoTo Ficha
        End If
    End If
Ficha:
    If lvMain.ListItems(x).SubItems(15) = 0 Then GoTo Proximo
   ' If lvMain.ListItems(x).SubItems(1) = "545068" Then MsgBox "teste"
    
    z = BinarySearchLong(aCodFicha(), CLng(nCodTrib))
    nCodBanco = Val(lvMain.ListItems(x).Text)
    If CDbl(lvMain.ListItems(x).SubItems(8)) > 0 Then
        If Not bDA And Not bAj Then
            lvMain.ListItems(x).SubItems(20) = aFicha(z).F1 'Principal
            FichaDetalhe dDataReceita, nCodBanco, aFicha(z).F1, CDbl(lvMain.ListItems(x).SubItems(8))
        ElseIf bDA And Not bAj Then
            lvMain.ListItems(x).SubItems(20) = aFicha(z).F3 'Principal DA
            FichaDetalhe dDataReceita, nCodBanco, aFicha(z).F3, CDbl(lvMain.ListItems(x).SubItems(8))
        ElseIf bDA And bAj Then
            lvMain.ListItems(x).SubItems(20) = aFicha(z).F6 ' Principal Aj
            FichaDetalhe dDataReceita, nCodBanco, aFicha(z).F6, CDbl(lvMain.ListItems(x).SubItems(8))
        End If
    End If
    
    If (CDbl(lvMain.ListItems(x).SubItems(9)) > 0 Or CDbl(lvMain.ListItems(x).SubItems(10)) > 0) Then
        If Not bDA And Not bAj Then
            lvMain.ListItems(x).SubItems(21) = aFicha(z).F2 'Juros e Multa
            FichaDetalhe dDataReceita, nCodBanco, aFicha(z).F2, CDbl(lvMain.ListItems(x).SubItems(9)) + CDbl(lvMain.ListItems(x).SubItems(10))
        ElseIf bDA And Not bAj Then
            lvMain.ListItems(x).SubItems(21) = aFicha(z).F4 'Juros e Multa DA
            FichaDetalhe dDataReceita, nCodBanco, aFicha(z).F4, CDbl(lvMain.ListItems(x).SubItems(9)) + CDbl(lvMain.ListItems(x).SubItems(10))
        ElseIf bDA And bAj Then
            lvMain.ListItems(x).SubItems(21) = aFicha(z).F7 'Juros e Multa Aj
            FichaDetalhe dDataReceita, nCodBanco, aFicha(z).F7, CDbl(lvMain.ListItems(x).SubItems(9)) + CDbl(lvMain.ListItems(x).SubItems(10))
        End If
    End If
    
    If CDbl(lvMain.ListItems(x).SubItems(11)) > 0 Then
        If Not bDA And Not bAj Then
            lvMain.ListItems(x).SubItems(22) = aFicha(z).F5 'Correcao
            FichaDetalhe dDataReceita, nCodBanco, aFicha(z).F5, CDbl(lvMain.ListItems(x).SubItems(11))
        ElseIf bDA And Not bAj Then
            lvMain.ListItems(x).SubItems(22) = aFicha(z).F5 'Correcao DA
            FichaDetalhe dDataReceita, nCodBanco, aFicha(z).F5, CDbl(lvMain.ListItems(x).SubItems(11))
        ElseIf bDA And bAj Then
            lvMain.ListItems(x).SubItems(22) = aFicha(z).F8 'Correcao Aj
            FichaDetalhe dDataReceita, nCodBanco, aFicha(z).F8, CDbl(lvMain.ListItems(x).SubItems(11))
        End If
    End If
    
Proximo:
    nValorTotalPago = nValorTotalPago - CDbl(lvMain.ListItems(x).SubItems(12))
Next
Pb.value = 100
lblPB.Caption = "100%"


'Arredondamento

nTotalMatriz = 0
nMaiorValor = 0
For x = 1 To UBound(aFichaDetalhe)
   nTotalMatriz = nTotalMatriz + aFichaDetalhe(x).Total
   If aFichaDetalhe(x).Total > nMaiorValor Then
      nMaiorValor = aFichaDetalhe(x).Total
      nIndMaior = x
   End If
Next

nDif = nTotalMatriz - CDbl(lblTotalDia.Caption)

If Round(nDif, 2) > 0 Then
    aFichaDetalhe(nIndMaior).Total = aFichaDetalhe(nIndMaior).Total - nDif
ElseIf Round(nDif, 2) < 0 Then
    aFichaDetalhe(nIndMaior).Total = aFichaDetalhe(nIndMaior).Total + Abs(nDif)
End If

If cmbBanco.ListIndex = 0 Then
   Open sPathBin & "\REC" & Format(Day(Now), "00") & Format(Month(Now), "00") For Output Shared As #1
Else
   Open sPathBin & "\REC" & Format(Day(Now), "00") & Format(Month(Now), "00") & Format(cmbBanco.ItemData(cmbBanco.ListIndex), "000") For Output Shared As #1
End If

For x = 1 To UBound(aFichaDetalhe)
    With aFichaDetalhe(x)
        Sql = "insert analise2 (usuario,datareceita,codbanco,valortotal,numficha,natureza,vinculo,perc,descficha) values('"
        Sql = Sql & NomeDeLogin & "','" & Format(dDataReceita, "mm/dd/yyyy") & "'," & .Banco & "," & Virg2Ponto(CStr(.Total)) & ","
        Sql = Sql & .Ficha & ",'" & .Natureza & "','" & .Vinculo & "'," & Virg2Ponto(CStr(.Perc)) & ",'" & .Descricao & "')"
        cn.Execute Sql, rdExecDirect
        
        ax = FillSpace(.Natureza, 20) & FillSpace(.Vinculo, 20) & Year(.Data) & Format(Month(.Data), "00") & Format(Day(.Data), "00") & Format(.Banco, "0000") & "00000000" & Format(RemovePonto(Replace(CStr(FormatNumber(.Total, 2)), ",", "")), "0000000000000")
        Print #1, ax
    End With
Next
Close #1

fim:
Liberado
If lvMain.ListItems.Count = 0 Then
    MsgBox "N�o existem baixas neste per�odo para este(s) banco(s).", vbInformation, "Aten��o"
Else
    If frmMdi.frTeste.Visible = True Then
        frmReport.ShowReport "Analise2_tmp", frmMdi.HWND, Me.HWND
    Else
        frmReport.ShowReport "Analise2", frmMdi.HWND, Me.HWND
    End If
    Sql = "delete from analise2 where usuario='" & NomeDeLogin & "'"
    cn.Execute Sql, rdExecDirect
End If
dpData.Enabled = True
opt1(0).Enabled = True
opt1(1).Enabled = True

If opt1(1).value = True Then
    cmbBanco.Enabled = True
End If
cmdGerar.Enabled = True




End Sub

Private Sub FichaDetalhe(Data, Banco, NumFicha As Long, Valor As Double)
Dim Sql As String, RdoAux As rdoResultset, x As Integer, bFind As Boolean, q As Integer

bFind = False
For q = 1 To UBound(aFichaDetalhe)
    If aFichaDetalhe(q).Ficha = NumFicha And aFichaDetalhe(q).Banco = Banco Then
        bFind = True
        Exit For
    End If
Next

If bFind Then
    For q = 1 To UBound(aFichaDetalhe)
        If aFichaDetalhe(q).Ficha = NumFicha And aFichaDetalhe(q).Banco = Banco Then
            aFichaDetalhe(q).Total = aFichaDetalhe(q).Total + (Valor * aFichaDetalhe(q).Perc / 100)
        End If
    Next
Else
    Sql = "SELECT NATUREZA,VINCULO,PERC,DESCTA FROM FICHACONTABIL WHERE FICHA=" & NumFicha
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            ReDim Preserve aFichaDetalhe(UBound(aFichaDetalhe) + 1)
            aFichaDetalhe(UBound(aFichaDetalhe)).Data = Data
            aFichaDetalhe(UBound(aFichaDetalhe)).Banco = Banco
            aFichaDetalhe(UBound(aFichaDetalhe)).Ficha = NumFicha
            aFichaDetalhe(UBound(aFichaDetalhe)).Descricao = Left(!descta, 50)
            aFichaDetalhe(UBound(aFichaDetalhe)).Seq = .AbsolutePosition
            aFichaDetalhe(UBound(aFichaDetalhe)).Natureza = !Natureza
            aFichaDetalhe(UBound(aFichaDetalhe)).Vinculo = !Vinculo
            aFichaDetalhe(UBound(aFichaDetalhe)).Perc = !Perc
            aFichaDetalhe(UBound(aFichaDetalhe)).Total = aFichaDetalhe(UBound(aFichaDetalhe)).Total + (Valor * !Perc / 100)
           .MoveNext
        Loop
       .Close
    End With
End If

End Sub

Private Sub CallPb(nPosF As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents

If ((nPosF * 100) / nTotal) <= 100 Then
   Pb.value = (nPosF * 100) / nTotal
Else
   Pb.value = 100
End If
lblPB.Caption = Int(Pb.value) & " %"

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
Resume Next
End Sub

Private Sub opt1_Click(Index As Integer)

If opt1(0).value = True Then
    cmbBanco.Enabled = False
Else
    cmbBanco.Enabled = True
End If

End Sub

Private Function FillSpace(sPalavra As String, nTamanho As Integer) As String
Dim sTmp As String

If Len(sPalavra) > nTamanho Then sPalavra = Left(sPalavra, nTamanho)
sTmp = sPalavra & Space(nTamanho - Len(sPalavra))
FillSpace = sTmp

End Function

Private Sub CarregaGrid3()
Dim Sql As String, RdoAux As rdoResultset, sDataLote As String, qd As New rdoQuery, RdoDeb As rdoResultset, xId As Long, nNumRec As Long, nSomaTmp As Double
Dim nValorPago As Double, nValorTotalPago As Double, aReg() As Reg, x As Integer, nCodTrib As Integer, bDA As Boolean, bAj As Boolean, nValorDif As Double, bFind As Boolean
Dim z As Long, nCodFicha As Integer, bJuros As Boolean, bMulta As Boolean, bCorrecao As Boolean, nCodFichaP As Integer, nCodFichaJM As Integer, nCodFichaC As Integer
Dim nIndex As Integer, nIndex2 As Integer, aRegDoc() As Registros, RdoAux2 As rdoResultset, nValorTmp As Double, v As Integer, nUserID As Integer, aFichaValor() As FichaValor
Dim aDoc() As tDoc

ReDim aReg(0): ReDim aRegDoc(0)
sDataLote = Format(Now, "ddmmhhmm")
Set qd.ActiveConnection = cn
qd.QueryTimeout = 0
lblTotalDia.Caption = "0,00"
dpData.Enabled = False
opt1(0).Enabled = False
opt1(1).Enabled = False
cmbBanco.Enabled = False
cmdGerar.Enabled = False

nUserID = RetornaUsuarioID(NomeDeLogin)
Sql = "delete from resumo_pagto_banco_ficha where userid=" & RetornaUsuarioID(NomeDeLogin)
cn.Execute Sql, rdExecDirect
CarregaTributo

Ocupado

dDataReceita = dpData.value
If cmbBanco.ListIndex > 0 Then
    nCodBanco = cmbBanco.ItemData(cmbBanco.ListIndex)
Else
    nCodBanco = 0
End If

Sql = "SELECT SUM(debitopago.valorpagoreal) AS soma FROM debitoparcela INNER JOIN debitopago ON debitoparcela.codreduzido = debitopago.codreduzido AND "
Sql = Sql & "debitoparcela.anoexercicio = debitopago.anoexercicio AND debitoparcela.codlancamento = debitopago.codlancamento AND debitoparcela.seqlancamento = debitopago.seqlancamento AND "
Sql = Sql & "debitoparcela.numparcela = debitopago.numparcela AND debitoparcela.codcomplemento = debitopago.codcomplemento INNER JOIN numdocumento ON debitopago.numdocumento = numdocumento.numdocumento "
Sql = Sql & "WHERE RESTITUIDO IS NULL AND DATARECEBIMENTO = '" & Format(dDataReceita, "mm/dd/yyyy") & "' "
If opt1(0).value = True Then
    Sql = Sql & " AND (DEBITOPAGO.CODBANCO=90 or DEBITOPAGO.CODBANCO=91 OR DEBITOPAGO.CODBANCO=92 OR DEBITOPAGO.CODBANCO=93 OR DEBITOPAGO.CODBANCO=94 OR DEBITOPAGO.CODBANCO=95 OR DEBITOPAGO.CODBANCO=96 OR DEBITOPAGO.CODBANCO=97 OR DEBITOPAGO.CODBANCO=98) "
Else
    If nCodBanco > 0 Then
        Sql = Sql & " AND DEBITOPAGO.CODBANCO=" & nCodBanco
    Else
        Sql = Sql & " AND (DEBITOPAGO.CODBANCO not in(90,91,92,93,94,95,96,97,98))"
    End If
End If
If cmbCC.ListIndex = 0 Then
    Sql = Sql & " AND (DEBITOPAGO.CONTACORRENTE='0740004' OR DEBITOPAGO.CONTACORRENTE=NULL  OR DEBITOPAGO.CONTACORRENTE='' OR (CONVERT(int, debitopago.contacorrente) = 0)) AND debitopago.CODLANCAMENTO<>79"
ElseIf cmbCC.ListIndex = 1 Then
    Sql = Sql & " AND (DEBITOPAGO.CONTACORRENTE='0740004' OR DEBITOPAGO.CONTACORRENTE=NULL  OR DEBITOPAGO.CONTACORRENTE='' OR (CONVERT(int, debitopago.contacorrente) = 0)) AND debitopago.CODLANCAMENTO=79 "
ElseIf cmbCC.ListIndex = 2 Then
    Sql = Sql & " AND DEBITOPAGO.CONTACORRENTE='" & cmbCC.Text & "'"
End If

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!soma) Then
    lblTotalDia.Caption = Format(0, "#0.00")
    GoTo fim
Else
    lblTotalDia.Caption = Format(RdoAux!soma, "#0.00")
    nValorTotalPago = RdoAux!soma
End If
RdoAux.Close

Sql = "SELECT DISTINCT debitopago.codreduzido, debitopago.anoexercicio, debitopago.codlancamento, debitopago.seqlancamento, debitopago.numparcela, debitopago.codcomplemento, debitopago.seqpag, debitopago.datapagamento,"
Sql = Sql & "debitopago.datarecebimento, debitopago.valorpago, debitopago.codbanco, debitopago.codagencia, debitopago.restituido, debitopago.numdocumento, debitopago.valorpagoreal, debitopago.intacto, debitopago.valortarifa,"
Sql = Sql & "debitopago.arquivobanco , debitopago.valordif, debitopago.datapagamentocalc, debitopago.dataintegracao, debitopago.contacorrente, parceladocumento.plano, plano.desconto "
Sql = Sql & "FROM debitopago INNER JOIN parceladocumento ON debitopago.codreduzido = parceladocumento.codreduzido AND debitopago.anoexercicio = parceladocumento.anoexercicio AND debitopago.codlancamento = parceladocumento.codlancamento AND "
Sql = Sql & "debitopago.seqlancamento = parceladocumento.seqlancamento AND debitopago.numparcela = parceladocumento.numparcela AND debitopago.codcomplemento = parceladocumento.codcomplemento AND "
Sql = Sql & "debitopago.numdocumento = parceladocumento.numdocumento LEFT OUTER JOIN plano ON parceladocumento.plano = plano.codigo "
Sql = Sql & "WHERE RESTITUIDO IS NULL AND DATARECEBIMENTO = '" & Format(dDataReceita, "mm/dd/yyyy") & "' "
If opt1(0).value = True Then
    Sql = Sql & " AND (DEBITOPAGO.CODBANCO=90 or DEBITOPAGO.CODBANCO=91 OR DEBITOPAGO.CODBANCO=92 OR DEBITOPAGO.CODBANCO=93 OR DEBITOPAGO.CODBANCO=94 OR DEBITOPAGO.CODBANCO=95 OR DEBITOPAGO.CODBANCO=96 OR DEBITOPAGO.CODBANCO=97 OR DEBITOPAGO.CODBANCO=98) "
Else
    If nCodBanco > 0 Then
        Sql = Sql & " AND DEBITOPAGO.CODBANCO=" & nCodBanco
    Else
        Sql = Sql & " AND (DEBITOPAGO.CODBANCO not in(90,91,92,93,94,95,96,97,98))"
    End If
End If
If cmbCC.ListIndex = 0 Then
    Sql = Sql & " AND (DEBITOPAGO.CONTACORRENTE='0740004' OR DEBITOPAGO.CONTACORRENTE=NULL  OR DEBITOPAGO.CONTACORRENTE='' OR (CONVERT(int, debitopago.contacorrente) = 0)) AND debitopago.CODLANCAMENTO<>79"
ElseIf cmbCC.ListIndex = 1 Then
    Sql = Sql & " AND (DEBITOPAGO.CONTACORRENTE='0740004' OR DEBITOPAGO.CONTACORRENTE=NULL  OR DEBITOPAGO.CONTACORRENTE='' OR (CONVERT(int, debitopago.contacorrente) = 0)) AND debitopago.CODLANCAMENTO=79 "
Else
    Sql = Sql & " AND DEBITOPAGO.CONTACORRENTE='" & cmbCC.Text & "'"
End If
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nNumRec = .RowCount
    Do Until .EOF
    
        If IsNull(!desconto) Then
            nPercDesconto = 0
        Else
            nPercDesconto = !desconto
        End If
        nCodBanco = !CodBanco
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
        
        If xId Mod 50 = 0 Then
            CallPb xId, nNumRec
        End If
        
        On Error Resume Next
        RdoDeb.Close
        On Error GoTo 0
        qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
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
            If .RowCount > 0 Then
            
            Do Until .EOF
                nValorPago = !valorpagoreal
                'If RdoAux!NumDocumento = 17948047 Then MsgBox "teste"
                If Format(!datarecebimento, "dd/mm/yyyy") <> Format(dDataReceita, "dd/mm/yyyy") Then
                    GoTo NextDoc
                End If
                If !ValorTributo = 0 Then
                    GoTo NextDoc
                End If
                nIndex = UBound(aReg) + 1
                ReDim Preserve aReg(nIndex)
                aReg(nIndex).nCodBanco = nCodBanco
                aReg(nIndex).sArquivo = RdoAux!arquivobanco
                aReg(nIndex).nCodReduz = !CODREDUZIDO
                aReg(nIndex).nAno = !AnoExercicio
                aReg(nIndex).nLanc = !CodLancamento
                aReg(nIndex).nSeq = !SeqLancamento
                aReg(nIndex).nParc = !NumParcela
                aReg(nIndex).nCompl = !CODCOMPLEMENTO
                aReg(nIndex).nCodTrib = !CodTributo
                aReg(nIndex).sDescTributo = !ABREVTRIBUTO
                aReg(nIndex).nValorTrib = !ValorTributo
                aReg(nIndex).ValorPago = nValorPago
                aReg(nIndex).NumDocumento = RdoAux!NumDocumento
                If !DataPagamento <= !DataVencimentoCalc Then
                    aReg(nIndex).nValorJuros = 0
                    aReg(nIndex).nValorMulta = 0
                    aReg(nIndex).nValorCorrecao = 0
                    aReg(nIndex).nValorTotal = !ValorTributo
                Else
                    aReg(nIndex).nValorJuros = !ValorJuros - (!ValorJuros * nPercDesconto / 100)
                    aReg(nIndex).nValorMulta = !ValorMulta - (!ValorMulta * nPercDesconto / 100)
                    aReg(nIndex).nValorCorrecao = !ValorCorrecao
                    aReg(nIndex).nValorTotal = aReg(nIndex).nValorTrib + aReg(nIndex).nValorJuros + aReg(nIndex).nValorMulta + aReg(nIndex).nValorCorrecao
                End If
                aReg(nIndex).sDataInscricao = IIf(IsDate(!datainscricao), "S", "N")
                aReg(nIndex).sDataRecebimento = Format(!datarecebimento, "dd/mm/yyyy")
                aReg(nIndex).sDataAjuiza = IIf(IsDate(!dataajuiza), "S", "N")
NextDoc:
               .MoveNext
            Loop
            End If
           .Close
        End With
        xId = xId + 1
       .MoveNext
    Loop
   .Close
End With

'carrega array com os documentos
ReDim aDoc(0)
For nIndex = 1 To UBound(aReg)
    bFind = False
    For v = 1 To UBound(aDoc)
        'If aDoc(v).Documento = 17316323 Then MsgBox "teste"
        If aDoc(v).Documento = aReg(nIndex).NumDocumento Then
            bFind = True
            Exit For
        End If
    Next
    If Not bFind Then
       'If aReg(nIndex).NumDocumento = 17948047 Then MsgBox "teste"
        ReDim Preserve aDoc(UBound(aDoc) + 1)
        aDoc(UBound(aDoc)).Documento = aReg(nIndex).NumDocumento
        Sql = "select sum(valorpagoreal) as soma from debitopago where numdocumento=" & aReg(nIndex).NumDocumento & " and datarecebimento='" & Format(dDataReceita, "mm/dd/yyyy") & "'"
        
        If cmbCC.ListIndex = 0 Then
            Sql = Sql & " AND (DEBITOPAGO.CONTACORRENTE='0740004' OR DEBITOPAGO.CONTACORRENTE=NULL  OR DEBITOPAGO.CONTACORRENTE='' OR (CONVERT(int, debitopago.contacorrente) = 0)) AND debitopago.CODLANCAMENTO<>79"
        ElseIf cmbCC.ListIndex = 1 Then
            Sql = Sql & " AND (DEBITOPAGO.CONTACORRENTE='0740004' OR DEBITOPAGO.CONTACORRENTE=NULL  OR DEBITOPAGO.CONTACORRENTE='' OR (CONVERT(int, debitopago.contacorrente) = 0)) AND debitopago.CODLANCAMENTO=79 "
        Else
            Sql = Sql & " AND DEBITOPAGO.CONTACORRENTE='" & cmbCC.Text & "'"
        End If

        
        
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        aDoc(UBound(aDoc)).ValorPago = RdoAux2!soma
        RdoAux2.Close
        'aDoc(UBound(aDoc)).ValorPago = aReg(nIndex).ValorPago
    End If
Next

'calcula as taxas de propor��es
For nIndex = 1 To UBound(aDoc)
    '% Total de cada linha
    For v = 1 To UBound(aReg)
'       If aDoc(nIndex).Documento = 17948047 Then MsgBox "teste"
        If aReg(v).NumDocumento = aDoc(nIndex).Documento Then
            If aReg(v).nValorTotal > aDoc(nIndex).ValorPago Then
                aReg(v).nPercT = 100
            Else
                aReg(v).nPercT = aReg(v).nValorTotal * 100 / aDoc(nIndex).ValorPago
            End If
        End If
    Next
    'If aDoc(nIndex).Documento = 17818695 Then MsgBox "teste"
    'verifica se a soma do % dodocumento � de 100%
    nSomaTmp = 0
    For v = 1 To UBound(aReg)
        If aReg(v).NumDocumento = aDoc(nIndex).Documento Then
            nSomaTmp = nSomaTmp + aReg(v).nPercT
        End If
    Next
    If nSomaTmp < 100 Then
        For v = 1 To UBound(aReg)
            If aReg(v).NumDocumento = aDoc(nIndex).Documento Then
                 aReg(v).nPercT = aReg(v).nPercT + (100 - nSomaTmp)
                 If aReg(v).nPercT <= 100 Then
                    Exit For
                 Else
                    nSomaTmp = (100 - nSomaTmp)
                 End If
            End If
        Next
    End If
    
    '% Principal, Multa, Juros e Corre��o de cada linha
    For v = 1 To UBound(aReg)
       ' If aDoc(nIndex).Documento = 17316323 Then MsgBox "teste"
        If aReg(v).NumDocumento = aDoc(nIndex).Documento Then
            aReg(v).ValorT = aDoc(nIndex).ValorPago * aReg(v).nPercT / 100
            aReg(v).nPercP = aReg(v).nValorTrib * 100 / aReg(v).ValorT
            aReg(v).ValorPr = aReg(v).ValorT * aReg(v).nPercP / 100
            If (aReg(v).nValorJuros + aReg(v).nValorMulta) > 0 Then
                aReg(v).nPercJM = (aReg(v).nValorJuros + aReg(v).nValorMulta) * 100 / aReg(v).ValorT
                If aReg(v).nPercC >= 100 Then
                    aReg(v).nPercJM = 0
                Else
                    If aReg(v).nPercP + aReg(v).nPercJM > 100 Then
                        aReg(v).nPercJM = 100 - aReg(v).nPercP
                    End If
                End If
                aReg(v).ValorJM = aReg(v).ValorT * aReg(v).nPercJM / 100
            End If
            If aReg(v).nValorCorrecao > 0 Then
                aReg(v).nPercC = aReg(v).nValorCorrecao * 100 / aReg(v).ValorT
                If aReg(v).nPercP >= 100 Then
                    aReg(v).nPercC = 0
                Else
                    If aReg(v).nPercP + aReg(v).nPercJM > 100 Then
                        aReg(v).nPercC = 0
                    Else
                        If aReg(v).nPercP + aReg(v).nPercJM + aReg(v).nPercC > 100 Then
                            aReg(v).nPercC = 100 - (aReg(v).nPercP + aReg(v).nPercJM)
                        End If
                    End If
                End If
                aReg(v).ValorCr = aReg(v).ValorT * aReg(v).nPercC / 100
            End If
        End If
    Next
Next


For nIndex = 1 To UBound(aDoc)
    For v = 1 To UBound(aReg)
       'If aDoc(nIndex).Documento = 17316323 Then MsgBox "teste"
        If aReg(v).NumDocumento = aDoc(nIndex).Documento Then
            'carrega as fichas
            nCodTrib = aReg(nIndex).nCodTrib
            z = -1
            z = BinarySearchLong(aCodFicha(), CLng(nCodTrib))
            bDA = IIf(aReg(v).sDataInscricao = "S", True, False)
            bAj = IIf(aReg(v).sDataAjuiza = "S", True, False)
            If aReg(v).nValorJuros > 0 Then bJuros = True
            If aReg(v).nValorMulta > 0 Then bMulta = True
            If aReg(v).nValorCorrecao > 0 Then bCorrecao = True
            If Not bDA And Not bAj Then
                nCodFichaP = aFicha(z).F1 'Principal
                If bJuros Or bMulta Then
                    nCodFichaJM = aFicha(z).F2 'Juros e Multa normal
                End If
            End If
            If bDA And Not bAj Then
                nCodFichaP = aFicha(z).F3 'Principal DA
                If bJuros Or bMulta Then
                    nCodFichaJM = aFicha(z).F4 'Juros e Multa DA
                End If
                If bCorrecao Then
                    nCodFichaC = aFicha(z).F5 'Corre��o DA
                End If
            End If
            If bDA And bAj Then
                nCodFichaP = aFicha(z).F6 ' Principal Aj
                If bJuros Or bMulta Then
                    nCodFichaJM = aFicha(z).F7 'Juros e Multa AJ
                End If
                If bCorrecao Then
                    nCodFichaC = aFicha(z).F8 'Corre��o AJ
                End If
            End If
            
        '    *** PRINCIPAL ****
            If nCodFichaP > 0 Then
                Sql = "SELECT NATUREZA,VINCULO,PERC,DESCTA FROM FICHACONTABIL WHERE FICHA=" & nCodFichaP
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    Do Until .EOF
                        ReDim Preserve aRegDoc(UBound(aRegDoc) + 1)
                        aRegDoc(UBound(aRegDoc)).nNumDocumento = aReg(v).NumDocumento
                        aRegDoc(UBound(aRegDoc)).sDataRecebimento = aReg(v).sDataRecebimento
                        aRegDoc(UBound(aRegDoc)).nCodReduz = aReg(v).nCodReduz
                        aRegDoc(UBound(aRegDoc)).nCodBanco = aReg(v).nCodBanco
                        aRegDoc(UBound(aRegDoc)).nAno = aReg(v).nAno
                        aRegDoc(UBound(aRegDoc)).nLanc = aReg(v).nLanc
                        aRegDoc(UBound(aRegDoc)).nSeq = aReg(v).nSeq
                        aRegDoc(UBound(aRegDoc)).nParc = aReg(v).nParc
                        aRegDoc(UBound(aRegDoc)).nCompl = aReg(v).nCompl
                        aRegDoc(UBound(aRegDoc)).nCodFicha = nCodFichaP
                        aRegDoc(UBound(aRegDoc)).nCodTributo = aReg(v).nCodTrib
                        aRegDoc(UBound(aRegDoc)).sDescFicha = !descta
                        aRegDoc(UBound(aRegDoc)).sDescTributo = aReg(v).sDescTributo
                        aRegDoc(UBound(aRegDoc)).sArquivo = aReg(v).sArquivo
                        aRegDoc(UBound(aRegDoc)).sNatureza = !Natureza
                        aRegDoc(UBound(aRegDoc)).sVinculo = !Vinculo
                        aRegDoc(UBound(aRegDoc)).nPerc = !Perc
                        aRegDoc(UBound(aRegDoc)).nValor = aReg(v).ValorPr * !Perc / 100
                        aRegDoc(UBound(aRegDoc)).nValorPago = aDoc(nIndex).ValorPago
                       .MoveNext
                    Loop
                   .Close
                End With
            End If
        '   *******************
        '    *** juros e multa ****
            If nCodFichaJM > 0 And aReg(v).ValorJM > 0 Then
                Sql = "SELECT NATUREZA,VINCULO,PERC,DESCTA FROM FICHACONTABIL WHERE FICHA=" & nCodFichaJM
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    Do Until .EOF
                        ReDim Preserve aRegDoc(UBound(aRegDoc) + 1)
                        aRegDoc(UBound(aRegDoc)).nNumDocumento = aReg(v).NumDocumento
                        aRegDoc(UBound(aRegDoc)).sDataRecebimento = aReg(v).sDataRecebimento
                        aRegDoc(UBound(aRegDoc)).nCodReduz = aReg(v).nCodReduz
                        aRegDoc(UBound(aRegDoc)).nCodBanco = aReg(v).nCodBanco
                        aRegDoc(UBound(aRegDoc)).nAno = aReg(v).nAno
                        aRegDoc(UBound(aRegDoc)).nLanc = aReg(v).nLanc
                        aRegDoc(UBound(aRegDoc)).nSeq = aReg(v).nSeq
                        aRegDoc(UBound(aRegDoc)).nParc = aReg(v).nParc
                        aRegDoc(UBound(aRegDoc)).nCompl = aReg(v).nCompl
                        aRegDoc(UBound(aRegDoc)).nCodFicha = nCodFichaJM
                        aRegDoc(UBound(aRegDoc)).nCodTributo = aReg(v).nCodTrib
                        aRegDoc(UBound(aRegDoc)).sDescFicha = !descta
                        aRegDoc(UBound(aRegDoc)).sDescTributo = aReg(v).sDescTributo
                        aRegDoc(UBound(aRegDoc)).sArquivo = aReg(v).sArquivo
                        aRegDoc(UBound(aRegDoc)).sNatureza = !Natureza
                        aRegDoc(UBound(aRegDoc)).sVinculo = !Vinculo
                        aRegDoc(UBound(aRegDoc)).nPerc = !Perc
                        aRegDoc(UBound(aRegDoc)).nValor = aReg(v).ValorJM * !Perc / 100
                        aRegDoc(UBound(aRegDoc)).nValorPago = aDoc(nIndex).ValorPago
                       .MoveNext
                    Loop
                   .Close
                End With
            End If
        '   *******************
        
        '    *** corre��o ****
            If nCodFichaC > 0 Then
                Sql = "SELECT NATUREZA,VINCULO,PERC,DESCTA FROM FICHACONTABIL WHERE FICHA=" & nCodFichaC
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    Do Until .EOF
                        ReDim Preserve aRegDoc(UBound(aRegDoc) + 1)
                        aRegDoc(UBound(aRegDoc)).nNumDocumento = aReg(v).NumDocumento
                        aRegDoc(UBound(aRegDoc)).sDataRecebimento = aReg(v).sDataRecebimento
                        aRegDoc(UBound(aRegDoc)).nCodReduz = aReg(v).nCodReduz
                        aRegDoc(UBound(aRegDoc)).nCodBanco = aReg(v).nCodBanco
                        aRegDoc(UBound(aRegDoc)).nAno = aReg(v).nAno
                        aRegDoc(UBound(aRegDoc)).nLanc = aReg(v).nLanc
                        aRegDoc(UBound(aRegDoc)).nSeq = aReg(v).nSeq
                        aRegDoc(UBound(aRegDoc)).nParc = aReg(v).nParc
                        aRegDoc(UBound(aRegDoc)).nCompl = aReg(v).nCompl
                        aRegDoc(UBound(aRegDoc)).nCodFicha = nCodFichaC
                        aRegDoc(UBound(aRegDoc)).nCodTributo = aReg(v).nCodTrib
                        aRegDoc(UBound(aRegDoc)).sDescFicha = !descta
                        aRegDoc(UBound(aRegDoc)).sDescTributo = aReg(v).sDescTributo
                        aRegDoc(UBound(aRegDoc)).sArquivo = aReg(v).sArquivo
                        aRegDoc(UBound(aRegDoc)).sNatureza = !Natureza
                        aRegDoc(UBound(aRegDoc)).sVinculo = !Vinculo
                        aRegDoc(UBound(aRegDoc)).nPerc = !Perc
                        aRegDoc(UBound(aRegDoc)).nValor = aReg(v).ValorCr * !Perc / 100
                        aRegDoc(UBound(aRegDoc)).nValorPago = aDoc(nIndex).ValorPago
                       .MoveNext
                    Loop
                   .Close
                End With
            End If
        End If
    Next
Next

'%arredonda valores dentro do documento
For nIndex = 1 To UBound(aDoc)
    nSomaTmp = 0
    xId = 0
    For v = 1 To UBound(aRegDoc)
'        If aDoc(nIndex).Documento = 17908234 Then
'            MsgBox "teste"
'        End If
        If aRegDoc(v).nNumDocumento = aDoc(nIndex).Documento Then
            If xId = 0 Then xId = v
            nSomaTmp = nSomaTmp + aRegDoc(v).nValor
        End If
    Next
    If nSomaTmp < aRegDoc(xId).nValorPago Then
        aRegDoc(xId).nValor = aRegDoc(xId).nValor + (aDoc(nIndex).ValorPago - nSomaTmp)
    End If

Next


For v = 1 To UBound(aRegDoc)
    If aRegDoc(v).nValor > 0 Then
        Sql = "insert resumo_pagto_banco_ficha(userid,datacredito,documento,codigo,ano,lanc,seq,parc,compl,codtributo,desctributo,descficha,ficha,arquivo,natureza,vinculo,perc,valor,codbanco,id) values("
        Sql = Sql & nUserID & ",'" & Format(aRegDoc(v).sDataRecebimento, "mm/dd/yyyy") & "'," & aRegDoc(v).nNumDocumento & "," & aRegDoc(v).nCodReduz & "," & aRegDoc(v).nAno & ","
        Sql = Sql & aRegDoc(v).nLanc & "," & aRegDoc(v).nSeq & "," & aRegDoc(v).nParc & "," & aRegDoc(v).nCompl & "," & aRegDoc(v).nCodTributo & ",'" & aRegDoc(v).sDescTributo & "','" & aRegDoc(v).sDescFicha & "',"
        Sql = Sql & aRegDoc(v).nCodFicha & ",'" & aRegDoc(v).sArquivo & "','" & aRegDoc(v).sNatureza & "','" & aRegDoc(v).sVinculo & "'," & aRegDoc(v).nPerc & "," & Virg2Ponto(CStr(aRegDoc(v).nValor)) & ","
        Sql = Sql & aRegDoc(v).nCodBanco & "," & v & ")"
        cn.Execute Sql, rdExecDirect
    End If
Next





fim:
Sql = "delete from resumo_pagto_banco_ficha where userid=" & nUserID
'cn.Execute Sql, rdExecDirect

MsgBox "fim"


Liberado
Pb.value = 0
lblPB.Caption = "0%"
dpData.Enabled = True
opt1(0).Enabled = True
opt1(1).Enabled = True

If opt1(1).value = True Then
    cmbBanco.Enabled = True
End If
cmdGerar.Enabled = True

End Sub





'Private Function Retorna_Fichas_Documento(Numdocumento As Long) As Ficha_Documento()
'Dim Sql As String, RdoAux As rdoResultset, nUserID As Integer, aReg() As Reg, aRegDoc() As Ficha_Documento, RdoDeb As rdoResultset, qd As New rdoQuery
'Dim nIndex As Integer, nCodBanco As Integer, nIndex2 As Integer, bFind As Boolean, z As Long, bDA As Boolean, bAj As Boolean
'Dim bJuros As Boolean, bMulta As Boolean, bCorrecao As Boolean, nCodFichaP As Integer, nCodFichaJM As Integer, nCodFichaC As Integer
'
'ReDim aReg(0): ReDim aRegDoc(0)
'Set qd.ActiveConnection = cn
'qd.QueryTimeout = 0
'
'On Error GoTo Erro
''If Numdocumento = 17954996 Then
''    MsgBox "teste"
''End If
'Sql = "SELECT CODREDUZIDO,anoexercicio,codlancamento,seqlancamento,numparcela,codcomplemento,datarecebimento From debitopago "
'Sql = Sql & "Where Numdocumento = " & Numdocumento & " and datarecebimento>='" & Format(dtDataDe.value, "mm/dd/yyyy") & "'"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    Do Until .EOF
'        On Error Resume Next
'        RdoDeb.Close
'        On Error GoTo 0
'        qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
'        qd(0) = !CODREDUZIDO
'        qd(1) = !CODREDUZIDO
'        qd(2) = !AnoExercicio
'
'        qd(3) = !AnoExercicio
'        qd(4) = !CodLancamento
'        qd(5) = !CodLancamento
'        qd(6) = !SeqLancamento
'        qd(7) = !SeqLancamento
'        qd(8) = !NumParcela
'        qd(9) = !NumParcela
'        qd(10) = !CODCOMPLEMENTO
'        qd(11) = !CODCOMPLEMENTO
'        qd(12) = 0
'        qd(13) = 99
'        qd(14) = Format(!datarecebimento, "mm/dd/yyyy")
'        qd(15) = NomeDeLogin
'        Set RdoDeb = qd.OpenResultset(rdOpenKeyset)
'        With RdoDeb
'            Do Until .EOF
'                nCodBanco = !CodBanco
'                If nCodBanco >= 90 And nCodBanco < 99 Then
'                    If nCodBanco = 91 Then
'                        nCodBanco = 1
'                    ElseIf nCodBanco = 90 Then nCodBanco = 90
'                    ElseIf nCodBanco = 92 Then nCodBanco = 33
'                    ElseIf nCodBanco = 93 Then nCodBanco = 237
'                    ElseIf nCodBanco = 94 Then nCodBanco = 341
'                    ElseIf nCodBanco = 95 Then nCodBanco = 409
'                    ElseIf nCodBanco = 96 Then nCodBanco = 151
'                    ElseIf nCodBanco = 97 Then nCodBanco = 104
'                    ElseIf nCodBanco = 98 Then nCodBanco = 399
'                    End If
'                End If
'
'                nIndex = UBound(aReg) + 1
'                ReDim Preserve aReg(nIndex)
'                aReg(nIndex).nCodBanco = nCodBanco
'                aReg(nIndex).Numdocumento = !Numdocumento
'                aReg(nIndex).sDataRecebimento = !datarecebimento
'                aReg(nIndex).nCodReduz = Format(!CODREDUZIDO, "000000")
'                aReg(nIndex).nAno = !AnoExercicio
'                aReg(nIndex).nLanc = !CodLancamento
'                aReg(nIndex).nSeq = !SeqLancamento
'                aReg(nIndex).nParc = !NumParcela
'                aReg(nIndex).nCompl = !CODCOMPLEMENTO
'                aReg(nIndex).nCodTrib = Format(!CodTributo, "000")
'                aReg(nIndex).sDescTributo = !ABREVTRIBUTO
'                aReg(nIndex).nValorTrib = Format(!ValorTributo, "#0.00")
'                If !DataPagamento < !DataVencimentoCalc Then
'                    aReg(nIndex).nValorJuros = Format(0, "#0.00")
'                    aReg(nIndex).nValorMulta = Format(0, "#0.00")
'                    aReg(nIndex).nValorCorrecao = Format(0, "#0.00")
'                    aReg(nIndex).nValorTotal = Format(!ValorTributo, "#0.00")
'                Else
'                    If Year(!DataVencimento) = 2020 Then
'                       If Month(!DataVencimento) > 3 And Month(!DataVencimento) < 7 Then
'                          aReg(nIndex).nValorJuros = 0
'                          aReg(nIndex).nValorMulta = 0
'                       Else
'                          aReg(nIndex).nValorJuros = !ValorJuros
'                          aReg(nIndex).nValorMulta = !ValorMulta
'                       End If
'                    Else
'                       aReg(nIndex).nValorJuros = !ValorJuros
'                       aReg(nIndex).nValorMulta = !ValorMulta
'                    End If
'                    aReg(nIndex).nValorCorrecao = !ValorCorrecao
'                    aReg(nIndex).nValorTotal = aReg(nIndex).nValorTrib + aReg(nIndex).nValorJuros + aReg(nIndex).nValorMulta + aReg(nIndex).nValorCorrecao
'                End If
'
'                aReg(nIndex).sDataInscricao = IIf(IsDate(!datainscricao), "S", "N")
'                aReg(nIndex).sDataAjuiza = IIf(IsDate(!dataajuiza), "S", "N")
'               .MoveNext
'            Loop
'          .Close
'        End With
'       .MoveNext
'    Loop
'   .Close
'End With
'
'For nIndex = 1 To UBound(aReg)
'    z = -1
'    z = BinarySearchLong(aCodFicha(), CLng(aReg(nIndex).nCodTrib))
'    bDA = False: bAj = False: bJuros = False: bMulta = False: bCorrecao = False
'    If aReg(nIndex).sDataInscricao = "S" Then bDA = True
'    If aReg(nIndex).sDataAjuiza = "S" Then bAj = True
'    If aReg(nIndex).nValorJuros > 0 Then bJuros = True
'    If aReg(nIndex).nValorMulta > 0 Then bMulta = True
'    If aReg(nIndex).nValorCorrecao > 0 Then bCorrecao = True
''    If aReg(nIndex).nCodTrib = 505 Then MsgBox "teste"
'    If Not bDA And Not bAj Then
'        nCodFichaP = aFicha(z).F1 'Principal
'        If bJuros Or bMulta Then
'            nCodFichaJM = aFicha(z).F2 'Juros e Multa normal
'        End If
'    End If
'    If bDA And Not bAj Then
'        nCodFichaP = aFicha(z).F3 'Principal DA
'        If bJuros Or bMulta Then
'            nCodFichaJM = aFicha(z).F4 'Juros e Multa DA
'        End If
'        If bCorrecao Then
'            nCodFichaC = aFicha(z).F5 'Corre��o DA
'        End If
'    End If
'    If bDA And bAj Then
'        nCodFichaP = aFicha(z).F6 ' Principal Aj
'        If bJuros Or bMulta Then
'            nCodFichaJM = aFicha(z).F7 'Juros e Multa AJ
'        End If
'        If bCorrecao Then
'            nCodFichaC = aFicha(z).F8 'Corre��o AJ
'        End If
'    End If
'
''    *** PRINCIPAL ****
'    If nCodFichaP > 0 Then
'        bFind = False
'        For nIndex2 = 1 To UBound(aRegDoc)
'            If nCodFichaP = aRegDoc(nIndex2).nCodFicha Then
'                bFind = True
'                Exit For
'            End If
'        Next
'        If Not bFind Then
'            ReDim Preserve aRegDoc(UBound(aRegDoc) + 1)
'            aRegDoc(UBound(aRegDoc)).nNumDocumento = aReg(nIndex).Numdocumento
'            aRegDoc(UBound(aRegDoc)).sDataRecebimento = aReg(nIndex).sDataRecebimento
'            aRegDoc(UBound(aRegDoc)).nCodReduz = aReg(nIndex).nCodReduz
'            aRegDoc(UBound(aRegDoc)).nCodFicha = nCodFichaP
'            aRegDoc(UBound(aRegDoc)).nCodTributo = aReg(nIndex).nCodTrib
'            aRegDoc(UBound(aRegDoc)).sDescTributo = aReg(nIndex).sDescTributo
'            aRegDoc(UBound(aRegDoc)).nValor = aReg(nIndex).nValorTrib
'        Else
'            aRegDoc(nIndex2).nValor = aRegDoc(nIndex2).nValor + aReg(nIndex).nValorTrib
'        End If
'    End If
''   *******************
'
''    *** juros e multa ****
'    If nCodFichaJM > 0 Then
'        bFind = False
'        For nIndex2 = 1 To UBound(aRegDoc)
'            If nCodFichaJM = aRegDoc(nIndex2).nCodFicha Then
'                bFind = True
'                Exit For
'            End If
'        Next
'        If Not bFind Then
'            ReDim Preserve aRegDoc(UBound(aRegDoc) + 1)
'            aRegDoc(UBound(aRegDoc)).nNumDocumento = aReg(nIndex).Numdocumento
'            aRegDoc(UBound(aRegDoc)).sDataRecebimento = aReg(nIndex).sDataRecebimento
'            aRegDoc(UBound(aRegDoc)).nCodReduz = aReg(nIndex).nCodReduz
'            aRegDoc(UBound(aRegDoc)).nCodFicha = nCodFichaJM
'            aRegDoc(UBound(aRegDoc)).nCodTributo = aReg(nIndex).nCodTrib
'            aRegDoc(UBound(aRegDoc)).sDescTributo = aReg(nIndex).sDescTributo
'            aRegDoc(UBound(aRegDoc)).nValor = aReg(nIndex).nValorJuros + aReg(nIndex).nValorMulta
'        Else
'            aRegDoc(nIndex2).nValor = aRegDoc(nIndex2).nValor + aReg(nIndex).nValorJuros + aReg(nIndex).nValorMulta
'        End If
'    End If
''   *******************
'
''    *** corre��o ****
'    If nCodFichaC > 0 Then
'        bFind = False
'        For nIndex2 = 1 To UBound(aRegDoc)
'            If nCodFichaC = aRegDoc(nIndex2).nCodFicha Then
'                bFind = True
'                Exit For
'            End If
'        Next
'        If Not bFind Then
'            ReDim Preserve aRegDoc(UBound(aRegDoc) + 1)
'            aRegDoc(UBound(aRegDoc)).nNumDocumento = aReg(nIndex).Numdocumento
'            aRegDoc(UBound(aRegDoc)).sDataRecebimento = aReg(nIndex).sDataRecebimento
'            aRegDoc(UBound(aRegDoc)).nCodReduz = aReg(nIndex).nCodReduz
'            aRegDoc(UBound(aRegDoc)).nCodFicha = nCodFichaC
'            aRegDoc(UBound(aRegDoc)).nCodTributo = aReg(nIndex).nCodTrib
'            aRegDoc(UBound(aRegDoc)).sDescTributo = aReg(nIndex).sDescTributo
'            aRegDoc(UBound(aRegDoc)).nValor = aReg(nIndex).nValorCorrecao
'        Else
'            aRegDoc(nIndex2).nValor = aRegDoc(nIndex2).nValor + aReg(nIndex).nValorCorrecao
'        End If
'    End If
''   *******************
'
'Next
'
'Retorna_Fichas_Documento = aRegDoc
'
'Exit Function
'Erro:
'MsgBox Err.Description
'
'End Function
'
'
Private Sub TESTE()
Dim Sql As String, RdoAux As rdoResultset, sNomeArq As String, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, nValor As Double
Dim nSomaFicha As Double, nSomaPago As Double
sNomeArq = sPathBin & "\teste.txt"
nValor = 0: nSomaFicha = 0: nSomaPago = 0
Open sNomeArq For Output As #1
Sql = "SELECT DISTINCT(numdocumento) FROM debitopago WHERE datarecebimento='" & Format(dpData.value, "mm/dd/yyyy") & "' ORDER BY numdocumento"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Sql = "select sum(valorpagoreal) as totalpago from debitopago where numdocumento=" & !NumDocumento & " and datarecebimento='" & Format(dpData.value, "mm/dd/yyyy") & "'"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        Sql = "select sum(valor) as total from resumo_pagto_banco_ficha where documento=" & !NumDocumento & " and datacredito='" & Format(dpData.value, "mm/dd/yyyy") & "'"
        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux3!Total > 0 Then
'            nValor = nValor + RdoAux2!totalpago
'            If FormatNumber(RdoAux2!totalpago, 2) <> FormatNumber(RdoAux3!Total, 2) Then
'            MsgBox !NumDocumento
'            End If
            nSomaPago = nSomaPago + RdoAux2!totalpago
            nSomaFicha = nSomaFicha + RdoAux3!Total
            Print #1, !NumDocumento & " - " & FormatNumber(RdoAux2!totalpago, 2) & " - " & FormatNumber(RdoAux3!Total, 2)
        End If
        RdoAux2.Close
        RdoAux3.Close
       .MoveNext
    Loop
   .Close
End With
Close #1
'MsgBox nSomaPago
'MsgBox nSomaFicha
ret = Shell("NOTEPAD" & " " & sNomeArq, vbNormalFocus)

End Sub
