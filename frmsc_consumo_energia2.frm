VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmsc_consumo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consumo"
   ClientHeight    =   7485
   ClientLeft      =   10200
   ClientTop       =   4965
   ClientWidth     =   12435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   12435
   Begin VB.TextBox txtConsumo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
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
      Height          =   285
      Left            =   765
      TabIndex        =   17
      Top             =   7425
      Width           =   960
   End
   Begin VB.Frame Frame2 
      Height          =   465
      Left            =   90
      TabIndex        =   10
      Top             =   6975
      Width           =   12255
      Begin VB.Label lblCod 
         Caption         =   "0"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3375
         TabIndex        =   16
         Top             =   180
         Width           =   465
      End
      Begin VB.Label lblCol 
         Caption         =   "0"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1800
         TabIndex        =   15
         Top             =   180
         Width           =   465
      End
      Begin VB.Label lblRow 
         Caption         =   "0"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   585
         TabIndex        =   14
         Top             =   180
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Código:"
         Height          =   195
         Index           =   2
         Left            =   2700
         TabIndex        =   13
         Top             =   180
         Width           =   690
      End
      Begin VB.Label Label2 
         Caption         =   "Col:"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   12
         Top             =   180
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Row:"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   11
         Top             =   180
         Width           =   465
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdMain 
      Height          =   6045
      Left            =   45
      TabIndex        =   9
      Top             =   720
      Width           =   12345
      _ExtentX        =   21775
      _ExtentY        =   10663
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      FixedCols       =   4
      ForeColor       =   1052688
      BackColorFixed  =   12632256
      ForeColorFixed  =   0
      BackColorBkg    =   -2147483633
      AllowBigSelection=   0   'False
      HighLight       =   2
      Appearance      =   0
      FormatString    =   $"frmsc_consumo_energia2.frx":0000
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   45
      TabIndex        =   4
      Top             =   0
      Width           =   12345
      Begin VB.TextBox txtLigacao 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5085
         MaxLength       =   30
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   225
         Width           =   2805
      End
      Begin VB.TextBox txtUnidade 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8955
         MaxLength       =   30
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   225
         Width           =   3030
      End
      Begin VB.TextBox txtAno 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3105
         MaxLength       =   4
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   225
         Width           =   600
      End
      Begin VB.ComboBox cmbMes 
         Height          =   315
         ItemData        =   "frmsc_consumo_energia2.frx":00E9
         Left            =   720
         List            =   "frmsc_consumo_energia2.frx":00EB
         Style           =   2  'Dropdown List
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   225
         Width           =   1590
      End
      Begin VB.Label lblLigacao 
         Caption         =   "..:"
         Height          =   240
         Left            =   3960
         TabIndex        =   8
         Top             =   270
         Width           =   1050
      End
      Begin VB.Label Label1 
         Caption         =   "Unidade..:"
         Height          =   240
         Index           =   2
         Left            =   8145
         TabIndex        =   7
         Top             =   270
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Ano..:"
         Height          =   240
         Index           =   1
         Left            =   2565
         TabIndex        =   6
         Top             =   270
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Mês..:"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   270
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmsc_consumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FormCall As String
Private Type tLista
    ano As Integer
    mes As Integer
    Codigo As Integer
    Unidade As String
    Secretaria As String
    ligacao As String
    consumo As Double
    valor As Double
End Type

Dim aLista() As tLista, bExec As Boolean
Dim nTipo As Integer, nAno As Integer, nMes As Integer

Public Property Let sForm(sNomeForm As String)
    FormCall = sNomeForm
End Property

Private Sub cmbMes_Click()
Gravar
nAno = Val(txtAno.Text)
nMes = cmbMes.ItemData(cmbMes.ListIndex)
CarregaLista
CarregaGrid
End Sub

Private Sub cmbMes_GotFocus()
txtConsumo.Visible = False
End Sub

Private Sub Form_Activate()
txtLigacao.SetFocus
End Sub

Private Sub Form_Load()
Dim x As Integer, ligacaoLabel As String, ligacaoHeader As String
Centraliza Me
bExec = False
Me.Top = Me.Top + 1100
cmbMes.AddItem ("JANEIRO")
cmbMes.AddItem ("FEVEREIRO")
cmbMes.AddItem ("MARÇO")
cmbMes.AddItem ("ABRIL")
cmbMes.AddItem ("MAIO")
cmbMes.AddItem ("JUNHO")
cmbMes.AddItem ("JULHO")
cmbMes.AddItem ("AGOSTO")
cmbMes.AddItem ("SETEMBRO")
cmbMes.AddItem ("OUTUBRO")
cmbMes.AddItem ("NOVEMBRO")
cmbMes.AddItem ("DEZEMBRO")

For x = 1 To 12
    cmbMes.ItemData(x - 1) = x
Next

cmbMes.ListIndex = Month(Now) - 1
txtAno.Text = Year(Now)
bExec = True
nAno = Val(txtAno.Text)
nMes = cmbMes.ItemData(cmbMes.ListIndex)

Select Case FormCall
    Case "frmsc_unidade_energia"
        Me.Caption = "Consumo de Energia"
        ligacaoLabel = "Nº Relógio..:"
        ligacaoHeader = "Nº Relógio"
        nTipo = 2
    Case "frmsc_telefonia_fixa"
        Me.Caption = "Consumo de Telefonia Fixa"
        ligacaoLabel = "Nº Telefone..:"
        ligacaoHeader = "Nº Telefone"
        nTipo = 3
    Case "frmsc_telefonia_celular"
        Me.Caption = "Consumo de Telefonia Celular"
        ligacaoLabel = "Nº Telefone..:"
        ligacaoHeader = "Nº Telefone"
        nTipo = 4
    Case "frmsc_conexao_internet"
        Me.Caption = "Consumo de Intenet"
        ligacaoLabel = "Nº Telefone..:"
        ligacaoHeader = "Nº Telefone"
        nTipo = 5
End Select
FormataGrid
CarregaLista
CarregaGrid
lblLigacao.Caption = ligacaoLabel
grdMain.TextMatrix(0, 3) = ligacaoHeader

End Sub

Private Sub SetTextBox(Row As Integer, col As Integer)
If grdMain.Rows = 1 Or Row > grdMain.Rows - 1 Then Exit Sub

If Row > 0 And (col = 4 Or col = 5) Then
    grdMain.Row = Row
    grdMain.col = col
    txtConsumo.Visible = True
    txtConsumo.Text = grdMain.TextMatrix(Row, col)
    txtConsumo.Top = grdMain.Top + grdMain.CellTop
    txtConsumo.Left = grdMain.Left + grdMain.CellLeft
    txtConsumo.Height = grdMain.CellHeight
    txtConsumo.Width = grdMain.CellWidth
    txtConsumo.SetFocus
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Gravar
End Sub

Private Sub grdMain_EnterCell()
Dim Row As Integer, col As Integer

Row = grdMain.Row
col = grdMain.col

If Row = 0 Then Exit Sub

lblRow.Caption = Row
lblCol.Caption = col
lblCod.Caption = grdMain.TextMatrix(Row, 0)
SetTextBox Row, col

End Sub

Private Sub grdMain_LeaveCell()
If Val(lblCol.Caption) = 4 Then
    If Trim(txtConsumo.Text) = "" Then txtConsumo.Text = "0"
ElseIf Val(lblCol.Caption) = 5 Then
    If Val(txtConsumo.Text) = 0 Then txtConsumo.Text = "0,00"
End If
grdMain.TextMatrix(Val(lblRow.Caption), Val(lblCol.Caption)) = txtConsumo.Text
txtConsumo.Visible = False
End Sub

Private Sub txtAno_GotFocus()
txtConsumo.Visible = False
txtAno.SelStart = 0
txtAno.SelLength = Len(txtAno.Text)
End Sub

Private Sub txtAno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CarregaLista
    CarregaGrid
End If
End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
Tweak txtAno, KeyAscii, IntegerAllowNegative
End Sub


Private Sub txtAno_LostFocus()
If Val(txtAno.Text) < 2022 Or Val(txtAno.Text) > Year(Now) + 1 Then
    MsgBox "Ano inválido", vbCritical, "Erro"
    txtAno.Text = Year(Now)
    CarregaLista
    CarregaGrid
End If
End Sub

Private Sub txtConsumo_GotFocus()
txtConsumo.SelStart = 0
txtConsumo.SelLength = Len(txtConsumo.Text)
End Sub

Private Sub txtConsumo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Row As Integer, col As Integer, nIndex As Integer

'37 left, 38 top, 39 right, 40 bottom
Select Case KeyCode
    Case 37
        If Val(lblCol.Caption) = 5 Then
            Row = Val(lblRow.Caption)
            col = Val(lblCol.Caption) - 1
        End If
    Case 38
        Row = Val(lblRow.Caption) - 1
        col = Val(lblCol.Caption)
    Case 39
        If Val(lblCol.Caption) = 4 Then
            Row = Val(lblRow.Caption)
            col = Val(lblCol.Caption) + 1
        End If
    Case 40
        Row = Val(lblRow.Caption) + 1
        col = Val(lblCol.Caption)
    Case 13
        If Val(lblCol.Caption) = 4 Then
            nIndex = RetornaIndex
            Row = Val(lblRow.Caption)
            col = Val(lblCol.Caption) + 1
            aLista(nIndex).consumo = txtConsumo.Text
            If Trim(txtConsumo.Text) = "" Then txtConsumo.Text = "0"
            grdMain.TextMatrix(Val(lblRow.Caption), Val(lblCol.Caption)) = txtConsumo.Text
        ElseIf Val(lblCol.Caption) = 5 Then
            nIndex = RetornaIndex
            aLista(nIndex).valor = txtConsumo.Text
            If Trim(txtConsumo.Text) = "" Then txtConsumo.Text = "0,00"
            grdMain.TextMatrix(Val(lblRow.Caption), Val(lblCol.Caption)) = txtConsumo.Text
            txtConsumo.Visible = False
            txtLigacao.Text = ""
            CarregaGrid
            txtLigacao.SetFocus
        End If
End Select
SetTextBox Row, col

End Sub

Private Sub txtConsumo_KeyPress(KeyAscii As Integer)
Tweak txtConsumo, KeyAscii, DecimalPositive, 2
End Sub



Private Sub txtUnidade_GotFocus()
txtConsumo.Visible = False
txtUnidade.SelStart = 0
txtUnidade.SelLength = Len(txtUnidade.Text)
End Sub

Private Sub FormataGrid()
Dim x As Integer, y As Integer

grdMain.TextMatrix(0, 0) = "Cód"
grdMain.TextMatrix(0, 1) = "Unidade"
grdMain.TextMatrix(0, 2) = "Secretaria"
grdMain.TextMatrix(0, 3) = "Nº Relógio"
grdMain.TextMatrix(0, 4) = "Consumo"
grdMain.TextMatrix(0, 5) = "Valor"

For y = 0 To 5
    grdMain.Row = 0
    grdMain.col = y
    grdMain.CellBackColor = &HC00000
    grdMain.CellForeColor = &HFFFFFF
Next
grdMain.AllowBigSelection = False

End Sub


Private Sub txtUnidade_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CarregaGrid
End If
End Sub

Private Sub txtLigacao_GotFocus()
txtConsumo.Visible = False
txtLigacao.SelStart = 0
txtLigacao.SelLength = Len(txtLigacao.Text)
End Sub

Private Sub txtLigacao_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CarregaGrid
End If
End Sub

Private Sub CarregaGrid()
Dim nMes As Integer, nAno As Integer, sLigacao As String, sUnidade As String, x As Integer
If Not bExec Then Exit Sub
sLigacao = Trim(UCase(txtLigacao.Text))
sUnidade = Trim(UCase(txtUnidade.Text))
grdMain.Rows = 1
Ocupado
For x = 1 To UBound(aLista)
    If sLigacao <> "" And sUnidade <> "" Then
        With aLista(x)
            If InStr(1, .ligacao, sLigacao) > 0 And InStr(1, sUnidade, UCase(.Unidade)) > 0 Then
                grdMain.AddItem .Codigo & Chr(9) & UCase(.Unidade) & Chr(9) & .Secretaria & Chr(9) & .ligacao & Chr(9) & .consumo & Chr(9) & Format(.valor, "#0.00")
            End If
        End With
    End If
    If sLigacao <> "" And sUnidade = "" Then
        With aLista(x)
            If InStr(1, .ligacao, sLigacao) <> 0 Then
                grdMain.AddItem .Codigo & Chr(9) & UCase(.Unidade) & Chr(9) & .Secretaria & Chr(9) & .ligacao & Chr(9) & .consumo & Chr(9) & Format(.valor, "#0.00")
            End If
        End With
    End If
    If sLigacao = "" And sUnidade <> "" Then
        With aLista(x)
            If InStr(1, UCase(.Unidade), sUnidade) > 0 Then
                grdMain.AddItem .Codigo & Chr(9) & UCase(.Unidade) & Chr(9) & .Secretaria & Chr(9) & .ligacao & Chr(9) & .consumo & Chr(9) & Format(.valor, "#0.00")
            End If
        End With
    End If
    If sLigacao = "" And sUnidade = "" Then
        With aLista(x)
             grdMain.AddItem .Codigo & Chr(9) & UCase(.Unidade) & Chr(9) & .Secretaria & Chr(9) & .ligacao & Chr(9) & .consumo & Chr(9) & Format(.valor, "#0.00")
        End With
    End If
Next

If grdMain.Rows > 1 Then
    lblRow.Caption = "1"
    lblCol.Caption = "4"
End If


Fim:

On Error Resume Next
If txtLigacao.Text <> "" Or txtUnidade.Text <> "" Then
    If grdMain.Rows > 1 Then
        grdMain.SetFocus
        grdMain.Row = 1
        grdMain.col = 4
        SetTextBox 1, 4
   End If
End If
Liberado
End Sub

Private Sub CarregaLista()
Dim nMes As Integer, nAno As Integer, sLigacao As String, sUnidade As String, nPos As Integer, Sql As String, RdoAux As rdoResultset
If Not bExec Then Exit Sub
sLigacao = Trim(UCase(txtLigacao.Text))
sUnidade = Trim(UCase(txtUnidade.Text))
ReDim aLista(0)
Ocupado
If nTipo = 2 Then
    Sql = "SELECT sc_ligacao_energia.codigo,sc_ligacao_energia.nome,sc_ligacao_energia.endereco,sc_ligacao_energia.dotacao,sc_ligacao_energia.ligacao,sc_secretaria.sigla,sc_secretaria.codigo AS codigo_secretaria, sc_ligacao_energia.empenho, sc_ligacao_energia.dia "
    Sql = Sql & "FROM sc_ligacao_energia INNER JOIN sc_secretaria ON sc_ligacao_energia.secretaria = sc_secretaria.codigo where 1=1 "
    If sLigacao <> "" Then
        Sql = Sql & " and sc_ligacao_energia.ligacao like '%" & Mask(sLigacao) & "%'"
    End If
    If sUnidade <> "" Then
        Sql = Sql & " and  sc_ligacao_energia.nome like '%" & Mask(sUnidade) & "%'"
    End If
    Sql = Sql & " ORDER BY sc_ligacao_energia.nome "
ElseIf nTipo = 3 Then
    Sql = "SELECT sc_telefonia_fixa.codigo,sc_telefonia_fixa.nome,sc_telefonia_fixa.endereco,sc_telefonia_fixa.dotacao,sc_telefonia_fixa.telefone as ligacao,sc_secretaria.sigla,sc_secretaria.codigo AS codigo_secretaria "
    Sql = Sql & "FROM sc_telefonia_fixa INNER JOIN sc_secretaria ON sc_telefonia_fixa.secretaria = sc_secretaria.codigo"
    If sLigacao <> "" Then
        Sql = Sql & " and sc_telefonia_fixa.telefone like '%" & Mask(sLigacao) & "%'"
    End If
    If sUnidade <> "" Then
        Sql = Sql & " and  sc_telefonia_fixa.nome like '%" & Mask(sUnidade) & "%'"
    End If
ElseIf nTipo = 4 Then
    Sql = "SELECT sc_telefonia_celular.codigo,sc_telefonia_celular.nome,sc_telefonia_celular.endereco,sc_telefonia_celular.dotacao,sc_telefonia_celular.telefone as ligacao,sc_secretaria.sigla,sc_secretaria.codigo AS codigo_secretaria "
    Sql = Sql & "FROM sc_telefonia_celular INNER JOIN sc_secretaria ON sc_telefonia_celular.secretaria = sc_secretaria.codigo"
    If sLigacao <> "" Then
        Sql = Sql & " and sc_telefonia_celular.telefone like '%" & Mask(sLigacao) & "%'"
    End If
    If sUnidade <> "" Then
        Sql = Sql & " and  sc_telefonia_celular.nome like '%" & Mask(sUnidade) & "%'"
    End If
ElseIf nTipo = 5 Then
Sql = "SELECT sc_conexao_internet.codigo,sc_conexao_internet.nome,sc_conexao_internet.endereco,sc_conexao_internet.dotacao,sc_conexao_internet.telefone as ligacao,sc_conexao_internet.ligacao,sc_secretaria.sigla,sc_secretaria.codigo AS codigo_secretaria "
Sql = Sql & "FROM sc_conexao_internet INNER JOIN sc_secretaria ON sc_conexao_internet.secretaria = sc_secretaria.codigo"
    If sLigacao <> "" Then
        Sql = Sql & " and sc_conexao_internet.telefone like '%" & Mask(sLigacao) & "%'"
    End If
    If sUnidade <> "" Then
        Sql = Sql & " and  sc_conexao_internet.nome like '%" & Mask(sUnidade) & "%'"
    End If
End If
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aLista(UBound(aLista) + 1)
        nPos = UBound(aLista)
        aLista(nPos).ano = Val(txtAno.Text)
        aLista(nPos).mes = cmbMes.ItemData(cmbMes.ListIndex)
        aLista(nPos).Codigo = !Codigo
        aLista(nPos).Unidade = !Nome
        aLista(nPos).Secretaria = !Sigla
        aLista(nPos).ligacao = !ligacao
       .MoveNext
    Loop
   .Close
End With


Sql = "select * from sc_consumo where tipo=" & nTipo & " and ano=" & Val(txtAno.Text) & " and mes=" & cmbMes.ItemData(cmbMes.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        For nPos = 1 To UBound(aLista)
            If aLista(nPos).Codigo = !Unidade Then
                Exit For
            End If
        Next
        aLista(nPos).consumo = !consumo
        aLista(nPos).valor = !valor
       .MoveNext
    Loop
   .Close
End With
Liberado
End Sub

Private Function RetornaIndex() As Integer
Dim x As Integer
For x = 1 To UBound(aLista)
    If aLista(x).Codigo = Val(lblCod.Caption) Then
       Exit For
    End If
Next
RetornaIndex = x
End Function

Private Sub Gravar()
Dim Sql As String, x As Integer, nUnidade As Integer, nConsumo As Double, nValor As Double
If Not bExec Then Exit Sub
If nAno = 0 Or nMes = 0 Then Exit Sub
Sql = "DELETE FROM SC_CONSUMO WHERE TIPO=" & nTipo & " AND ANO=" & nAno & " AND MES=" & nMes
cn.Execute Sql, rdExecDirect

For x = 1 To UBound(aLista)
    If aLista(x).consumo > 0 Then
        nUnidade = aLista(x).Codigo
        nConsumo = aLista(x).consumo
        nValor = aLista(x).valor
        Sql = "INSERT SC_CONSUMO(TIPO,ANO,MES,UNIDADE,CONSUMO,VALOR) VALUES(" & nTipo & "," & nAno & "," & nMes & "," & nUnidade & ","
        Sql = Sql & Virg2Ponto(CStr(nConsumo)) & "," & Virg2Ponto(CStr(nValor)) & ")"
        cn.Execute Sql, rdExecDirect
    End If
Next

End Sub


