VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form frmsc_unidade_agua 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle das Unidades - Consumo de Água"
   ClientHeight    =   6975
   ClientLeft      =   14865
   ClientTop       =   5340
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   10980
   Begin prjChameleon.chameleonButton cmdAlterar 
      Height          =   360
      Left            =   8325
      TabIndex        =   27
      ToolTipText     =   "Editar Registro"
      Top             =   6525
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "&Editar"
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
      MICON           =   "frmsc_unidade_agua.frx":0000
      PICN            =   "frmsc_unidade_agua.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdConsumo 
      Height          =   360
      Left            =   7065
      TabIndex        =   26
      ToolTipText     =   "Exibir consumo da unidade"
      Top             =   6525
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "&Consumo"
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
      MCOL            =   16777152
      MPTR            =   1
      MICON           =   "frmsc_unidade_agua.frx":0176
      PICN            =   "frmsc_unidade_agua.frx":0192
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame2 
      Height          =   1545
      Left            =   45
      TabIndex        =   5
      Top             =   4860
      Width           =   10905
      Begin VB.TextBox txtDotacao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8640
         MaxLength       =   6
         TabIndex        =   14
         Top             =   540
         Width           =   1950
      End
      Begin VB.ComboBox cmbSecretaria 
         Height          =   315
         Left            =   8640
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   180
         Width           =   1995
      End
      Begin VB.Label Label2 
         Caption         =   "Dotação...:"
         Height          =   195
         Index           =   8
         Left            =   7740
         TabIndex        =   28
         Top             =   585
         Width           =   870
      End
      Begin VB.Label lblHidrometro 
         Caption         =   "0000000"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   8640
         TabIndex        =   22
         Top             =   1170
         Width           =   1905
      End
      Begin VB.Label Label2 
         Caption         =   "Hidrômetro:"
         Height          =   195
         Index           =   7
         Left            =   7740
         TabIndex        =   21
         Top             =   1170
         Width           =   870
      End
      Begin VB.Label lblEndereco 
         Caption         =   "..."
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1080
         TabIndex        =   20
         Top             =   1170
         Width           =   6585
      End
      Begin VB.Label Label2 
         Caption         =   "Endereço...:"
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   19
         Top             =   1170
         Width           =   915
      End
      Begin VB.Label lblProprietario 
         Caption         =   "..."
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1080
         TabIndex        =   18
         Top             =   855
         Width           =   9150
      End
      Begin VB.Label lblDescricao 
         Caption         =   "..."
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1080
         TabIndex        =   17
         Top             =   540
         Width           =   6585
      End
      Begin VB.Label Label2 
         Caption         =   "Proprietário.:"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   16
         Top             =   855
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "Descrição..:"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   15
         Top             =   540
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "Secretaria.:"
         Height          =   195
         Index           =   3
         Left            =   7740
         TabIndex        =   12
         Top             =   225
         Width           =   870
      End
      Begin VB.Label lblInscricao 
         Caption         =   "000-0000-00000000-00000-00"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4230
         TabIndex        =   11
         Top             =   225
         Width           =   2445
      End
      Begin VB.Label Label2 
         Caption         =   "Nº Inscrição:"
         Height          =   195
         Index           =   2
         Left            =   3195
         TabIndex        =   10
         Top             =   225
         Width           =   960
      End
      Begin VB.Label lblLigacao 
         Caption         =   "0000000"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2160
         TabIndex        =   9
         Top             =   225
         Width           =   780
      End
      Begin VB.Label Label2 
         Caption         =   "Nº Ligação:"
         Height          =   195
         Index           =   1
         Left            =   1260
         TabIndex        =   8
         Top             =   225
         Width           =   915
      End
      Begin VB.Label lblCod 
         Caption         =   "0000"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   495
         TabIndex        =   7
         Top             =   225
         Width           =   510
      End
      Begin VB.Label Label2 
         Caption         =   "Cód:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   225
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   45
      TabIndex        =   3
      Top             =   -45
      Width           =   10905
      Begin VB.TextBox txtFilter 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1350
         TabIndex        =   1
         Top             =   195
         Width           =   8160
      End
      Begin prjChameleon.chameleonButton cmdFiltrar 
         Default         =   -1  'True
         Height          =   315
         Left            =   9765
         TabIndex        =   2
         ToolTipText     =   "Filtrar unidades"
         Top             =   180
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Filtrar"
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
         MICON           =   "frmsc_unidade_agua.frx":0530
         PICN            =   "frmsc_unidade_agua.frx":054C
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
         Caption         =   "Digite para filtrar:"
         Height          =   240
         Left            =   90
         TabIndex        =   4
         Top             =   225
         Width           =   1275
      End
   End
   Begin vbAcceleratorSGrid6.vbalGrid grdMain 
      Height          =   4200
      Left            =   45
      TabIndex        =   0
      Top             =   585
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   7408
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      AlternateRowBackColor=   15531775
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderDragReorderColumns=   0   'False
      HeaderHotTrack  =   0   'False
      HeaderFlat      =   -1  'True
      BorderStyle     =   2
      ScrollBarStyle  =   1
      Editable        =   -1  'True
      DisableIcons    =   -1  'True
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   360
      Left            =   9600
      TabIndex        =   23
      ToolTipText     =   "Sair da Tela"
      Top             =   6525
      Width           =   1170
      _ExtentX        =   2064
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmsc_unidade_agua.frx":0726
      PICN            =   "frmsc_unidade_agua.frx":0742
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdGravar 
      Height          =   360
      Left            =   8325
      TabIndex        =   25
      ToolTipText     =   "Gravar os Dados"
      Top             =   6525
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "&Gravar"
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
      MICON           =   "frmsc_unidade_agua.frx":07B0
      PICN            =   "frmsc_unidade_agua.frx":07CC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   9600
      TabIndex        =   24
      ToolTipText     =   "Cancelar Edição"
      Top             =   6525
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "&Cancelar"
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
      MICON           =   "frmsc_unidade_agua.frx":0B71
      PICN            =   "frmsc_unidade_agua.frx":0B8D
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
Attribute VB_Name = "frmsc_unidade_agua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAlterar_Click()
If grdMain.Rows = 0 Then
    MsgBox "Selecione uma unidade!", vbExclamation, "Erro"
    Exit Sub
End If

Eventos "INCLUIR"
End Sub

Private Sub cmdCancel_Click()
Eventos "INICIAR"
End Sub

Private Sub cmdConsumo_Click()
frmsc_unidade_agua_consumo.show vbModal
End Sub

Private Sub cmdFiltrar_Click()
Dim s As String

s = Trim(UCase(txtFilter.Text))
If Len(s) = 0 Then
    CarregaLista ""
    Exit Sub
End If
If Len(s) < 3 Then
    MsgBox "Digite ao menos 3 caracteres.", vbExclamation, "Atenção"
    Exit Sub
End If

CarregaLista s

End Sub

Private Sub cmdGravar_Click()
Dim nSec As Integer, nCodigo As Integer, nRow As Integer

nRow = grdMain.SelectedRow
nSec = cmbSecretaria.ItemData(cmbSecretaria.ListIndex)
nCodigo = Val(lblCod.Caption)

sql = "update sc_ligacao_agua set secretaria=" & nSec & ",dotacao=" & sNullVal(txtDotacao.Text) & " where codigo=" & nCodigo
cn.Execute sql, rdExecDirect

Eventos "INICIAR"
grdMain.cell(nRow, 2).Text = nSec
grdMain.cell(nRow, 3).Text = cmbSecretaria.Text
grdMain.cell(nRow, 12).Text = txtDotacao.Text

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim sql As String, RdoAux As rdoResultset

Centraliza Me
GridHeader

sql = "select codigo,sigla from sc_secretaria order by nome"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbSecretaria.AddItem !Sigla
        cmbSecretaria.ItemData(cmbSecretaria.NewIndex) = !codigo
       .MoveNext
    Loop
   .Close
End With

CarregaLista ""
Eventos "INICIAR"
End Sub

Private Sub GridHeader()
With grdMain
    .GridFillLineColor = vbWhite
    .Editable = False
    .GridLines = True
    .HighlightBackColor = Marrom
    .HighlightForeColor = vbWhite
    .RowMode = True
    .DefaultRowHeight = 17
    .AddColumn "kCod", "Cód", ecgHdrTextALignCentre, , 40
    .AddColumn "kScd", "CS", , 10, , False
    .AddColumn "kSec", "Secret.", , 80
    .AddColumn "kUsu", "Descrição da Unidade", ecgHdrTextALignLeft, , 250
    .AddColumn "kEnd", "Endereço", ecgHdrTextALignLeft, , 250
    .AddColumn "kCpl", "Complemento", ecgHdrTextALignLeft, , 80
    .AddColumn "kBai", "Bairro", ecgHdrTextALignLeft, , 120
    .AddColumn "kLig", "Ligação", ecgHdrTextALignLeft, , 50
    .AddColumn "kIns", "Inscrição", ecgHdrTextALignLeft, , 180
    .AddColumn "kPro", "Proprietário", ecgHdrTextALignLeft, , 250
    .AddColumn "kHid", "Hidrômetro", ecgHdrTextALignLeft, , 120
    .AddColumn "kDot", "Dotação", ecgHdrTextALignLeft, , 100
    .ColumnSortType("kCod") = CCLSortNumeric
End With

End Sub

Private Sub grdMain_ColumnClick(ByVal lCol As Long)
Dim sTag As String, iSortIndex As Long
   With grdMain.SortObject
      .ClearNongrouped
      iSortIndex = .IndexOf(lCol)
      If (iSortIndex = 0) Then
         iSortIndex = .Count + 1
         .SortColumn(iSortIndex) = lCol
      End If
      sTag = grdMain.ColumnTag(lCol)
      If (sTag = "") Then
         sTag = "DESC"
         .SortOrder(iSortIndex) = CCLOrderAscending
      Else
         sTag = ""
         .SortOrder(iSortIndex) = CCLOrderDescending
      End If
      grdMain.ColumnTag(lCol) = sTag
      .SortType(iSortIndex) = grdMain.ColumnSortType(lCol)
   End With
   Screen.MousePointer = vbHourglass
   grdMain.Sort
   Screen.MousePointer = vbDefault
End Sub

Private Sub CarregaLista(Filter As String)
Dim sql As String, RdoAux As rdoResultset

grdMain.Clear

sql = "SELECT sc_ligacao_agua.codigo,sc_ligacao_agua.usuario,sc_secretaria.sigla,sc_ligacao_agua.imovel_logradouro,sc_ligacao_agua.imovel_numero,sc_ligacao_agua.imovel_complemento,"
sql = sql & "sc_ligacao_agua.codigo_saaej,sc_ligacao_agua.inscricao_saaej,sc_ligacao_agua.proprietario,sc_ligacao_agua.imovel_bairro,vwsc_hidrometro.hidrometro,sc_secretaria.codigo AS codigo_secretaria,"
sql = sql & "sc_ligacao_agua.dotacao FROM sc_ligacao_agua INNER JOIN "
sql = sql & "sc_secretaria ON sc_ligacao_agua.secretaria = sc_secretaria.codigo LEFT OUTER JOIN dbo.vwsc_hidrometro ON sc_ligacao_agua.codigo = vwsc_hidrometro.Codigo "
If Filter <> "" Then
    sql = sql & " where usuario like '%" & Mask(Filter) & "%' or proprietario like '%" & Mask(Filter) & "%' or imovel_logradouro like '%" & Mask(Filter) & "%' or "
    sql = sql & "imovel_complemento like '%" & Mask(Filter) & "%' or imovel_bairro like '%" & Mask(Filter) & "%' or inscricao_saaej like '%" & Mask(Filter) & "%' or "
    sql = sql & "codigo_saaej like '%" & Mask(Filter) & "%' or hidrometro like '%" & Mask(Filter) & "%' or sigla like '%" & Mask(Filter) & "%'"
End If
sql = sql & " ORDER BY sc_ligacao_agua.usuario "
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        grdMain.AddRow
        grdMain.CellDetails grdMain.Rows, 1, !codigo, DT_RIGHT
        grdMain.CellDetails grdMain.Rows, 2, !codigo_secretaria
        grdMain.CellDetails grdMain.Rows, 3, !Sigla
        grdMain.CellDetails grdMain.Rows, 4, !USUARIO
        grdMain.CellDetails grdMain.Rows, 5, UCase(SubNull(!imovel_logradouro) & ", " & Val(SubNull(!imovel_numero)))
        grdMain.CellDetails grdMain.Rows, 6, UCase(SubNull(!imovel_complemento))
        grdMain.CellDetails grdMain.Rows, 7, UCase(SubNull(!imovel_bairro))
        grdMain.CellDetails grdMain.Rows, 8, SubNull(!codigo_saaej)
        grdMain.CellDetails grdMain.Rows, 9, SubNull(!inscricao_saaej)
        grdMain.CellDetails grdMain.Rows, 10, SubNull(!Proprietario)
        grdMain.CellDetails grdMain.Rows, 11, SubNull(!hidrometro)
        grdMain.CellDetails grdMain.Rows, 12, SubNull(!dotacao)
       .MoveNext
    Loop
   .Close
End With

If grdMain.Rows > 0 Then
    grdMain.SelectedRow = 1
    Le
End If

End Sub

Private Sub grdMain_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
Le
End Sub


Private Sub txtDotacao_KeyPress(KeyAscii As Integer)
Tweak txtDotacao, KeyAscii, IntegerPositive
End Sub

Private Sub txtFilter_Change()
If txtFilter.Text = "" Then CarregaLista ""
End Sub

Private Sub Le()
Dim Row As Integer, sCompl As String, nSec As Integer, x As Integer

Row = grdMain.SelectedRow
If Row = 0 Then Exit Sub
sCompl = grdMain.cell(Row, 6).Text
nSec = Val(grdMain.cell(Row, 2).Text)

lblCod.Caption = Format(grdMain.cell(Row, 1).Text, "0000")
lblLigacao.Caption = Format(grdMain.cell(Row, 8).Text, "0000000")
lblInscricao.Caption = grdMain.cell(Row, 9).Text
lblDescricao.Caption = grdMain.cell(Row, 4).Text
lblProprietario.Caption = grdMain.cell(Row, 10).Text
lblEndereco.Caption = grdMain.cell(Row, 5).Text & IIf(sCompl <> "", " ", "") & " - " & grdMain.cell(Row, 7).Text
lblHidrometro.Caption = grdMain.cell(Row, 11).Text
txtDotacao.Text = grdMain.cell(Row, 12).Text
If nSec = 0 Then
    cmbSecretaria.ListIndex = 0
Else
    For x = 0 To cmbSecretaria.ListCount - 1
        If cmbSecretaria.ItemData(x) = nSec Then
            cmbSecretaria.ListIndex = x
            Exit For
        End If
    Next
End If

End Sub

Private Sub Eventos(Tipo As String)

If Tipo = "INICIAR" Then
   cmdAlterar.Visible = True
   cmdConsumo.Visible = True
   cmdSair.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   cmbSecretaria.Enabled = False
   cmbSecretaria.BackColor = Kde
   txtDotacao.Locked = True
   txtDotacao.BackColor = Kde
   grdMain.Enabled = True
ElseIf Tipo = "INCLUIR" Then
   cmdAlterar.Visible = False
   cmdConsumo.Visible = False
   cmdSair.Visible = False
   cmbSecretaria.Enabled = True
   txtDotacao.Locked = False
   txtDotacao.BackColor = Branco
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   grdMain.Enabled = False
End If

End Sub

