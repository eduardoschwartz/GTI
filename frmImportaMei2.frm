VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmImportaMei2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importação de dados MEI"
   ClientHeight    =   8250
   ClientLeft      =   3690
   ClientTop       =   3945
   ClientWidth     =   10710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Tributacao.XP_ProgressBar PBar 
      Height          =   240
      Left            =   90
      TabIndex        =   3
      Top             =   7830
      Width           =   3795
      _ExtentX        =   6694
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
   Begin MSComctlLib.ListView lvMain 
      Height          =   7485
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   13203
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
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CNPJ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Razão Social"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Dt.Vencto"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Valor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Código"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Endereço"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Num"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Compl"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Bairro"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Cep"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Duplicado"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Arquivo"
         Object.Width           =   0
      EndProperty
   End
   Begin prjChameleon.chameleonButton cmdExec 
      Height          =   360
      Left            =   7155
      TabIndex        =   0
      ToolTipText     =   "Executar o Cancelamento"
      Top             =   7740
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "&Gerar os Débitos"
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
      MICON           =   "frmImportaMei2.frx":0000
      PICN            =   "frmImportaMei2.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   360
      Left            =   9225
      TabIndex        =   1
      ToolTipText     =   "Sair da Tela"
      Top             =   7740
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   635
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmImportaMei2.frx":00BB
      PICN            =   "frmImportaMei2.frx":00D7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmImportaMei2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExec_Click()
On Error GoTo Erro
Dim sNome As String, sCnpj As String, sDataVencto As String, sValor As String, Sql As String, sObs As String
Dim nCodigo As Long, x As Long, sEndereco As String, nNum As Integer, RdoAux As rdoResultset, nSeqMaxima As Integer
Dim sCompl As String, sBairro As String, sCep As String, sDuplicado As String, nAno As Integer, bFind As Boolean, sArquivo As String

If MsgBox("Deseja gerar os débitos para os registros acima?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub
cmdSair.Enabled = False
cmdExec.Enabled = False
For x = 1 To lvMain.ListItems.Count
    CallPb x, CLng(lvMain.ListItems.Count)
    sCnpj = lvMain.ListItems(x).Text
    sNome = lvMain.ListItems(x).SubItems(1)
    sDataVencto = lvMain.ListItems(x).SubItems(2)
    sValor = lvMain.ListItems(x).SubItems(3)
    nCodigo = Val(lvMain.ListItems(x).SubItems(4))
    sEndereco = lvMain.ListItems(x).SubItems(5)
    nNumero = Val(lvMain.ListItems(x).SubItems(6))
    sCompl = lvMain.ListItems(x).SubItems(7)
    sBairro = lvMain.ListItems(x).SubItems(8)
    sCep = lvMain.ListItems(x).SubItems(9)
    sDuplicado = lvMain.ListItems(x).SubItems(10)
    sArquivo = lvMain.ListItems(x).SubItems(11)
    
    If sDuplicado = "Não" Then
        Sql = "insert importacao_mei(cnpj,data_vencimento,nome,valor,codigo,endereco,numero,complemento,bairro,cep,ano,seq,arquivo) values('"
        Sql = Sql & sCnpj & "','" & Format(CDate(sDataVencto), "mm/dd/yyyy") & "','" & Mask(sNome) & "'," & Virg2Ponto(sValor) & ","
        Sql = Sql & nCodigo & ",'" & Mask(sEndereco) & "'," & nNumero & ",'" & Mask(sCompl) & "','" & Mask(sBairro) & "','" & sCep & "',"
        Sql = Sql & Year(CDate(sDataVencto)) & ",0,'" & Mask(sArquivo) & "')"
        cn.Execute Sql, rdExecDirect
    End If
    If nCodigo = 0 Then GoTo Proximo
    'Geração do débito
    nAno = Year(sDataVencto)
    Sql = "select codreduzido from debitoparcela where codreduzido=" & nCodigo & " and anoexercicio=" & nAno & " and codlancamento=5 and datavencimento='" & Format(CDate(sDataVencto), "mm/dd/yyyy") & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    bFind = False
    If RdoAux.RowCount > 0 Then
        bFind = True
    End If
    RdoAux.Close
    If Not bFind Then
       'busca próxima sequencia
        nSeqMaxima = 0
        Sql = "select max(seqlancamento) as maximo from debitoparcela where codreduzido=" & nCodigo & " and anoexercicio=" & nAno & " and codlancamento=5 and numparcela=1"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If Not IsNull(RdoAux!maximo) Then
            nSeqMaxima = RdoAux!maximo + 1
        End If
        RdoAux.Close
        
        'Atualiza seq na tabela importacao_mei
        Sql = "update importacao_mei set seq=" & nSeqMaxima & " where cnpj='" & sCnpj & "' and data_vencimento='" & Format(CDate(sDataVencto), "mm/dd/yyyy") & "'"
        cn.Execute Sql, rdExecDirect
        
        'GRAVA NA TABELA DEBITOPARCELA
        Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
        Sql = Sql & "DATAVENCIMENTO,DATADEBASE,USERID) VALUES(" & nCodigo & "," & nAno & ",5," & nSeqMaxima & "," & 1 & ",0,3,'"
        Sql = Sql & Format(CDate(sDataVencto), "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "'," & RetornaUsuarioID(NomeDeLogin) & ")"
        cn.Execute Sql, rdExecDirect
        
        'GRAVA NA TABELA DEBITOTRIBUTO
        Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,"
        Sql = Sql & "VALORTRIBUTO) VALUES(" & nCodigo & "," & nAno & ",5," & nSeqMaxima & "," & 1 & ",0,13," & Virg2Ponto(sValor) & ")"
        cn.Execute Sql, rdExecDirect
               
        'GRAVA OBS PARCELA
        sObs = "ISS Variável importado do arquivo MEI da Receita Federal"
        Sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & nCodigo & "," & nAno & ","
        Sql = Sql & 5 & "," & nSeqMaxima & ",1,0,0" & ",'" & sObs & "'," & RetornaUsuarioID(NomeDeLogin) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
        cn.Execute Sql, rdExecDirect
               
    End If
Proximo:
Next
CallPb 100, 100
Me.Refresh
cmdSair.Enabled = True
MsgBox "Todos os débitos da lista acima foram criados.", vbInformation, "Informação"
Unload Me
Exit Sub
Erro:
If rdoErrors(1).Number = 2627 Then
    Resume Next
Else
    MsgBox rdoErrors(1).Description
End If


End Sub

Private Sub cmdSair_Click()
Unload Me
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

