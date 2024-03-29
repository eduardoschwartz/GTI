VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form frmsc_unidade_energia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle das Unidades - Consumo de Energia"
   ClientHeight    =   6810
   ClientLeft      =   6420
   ClientTop       =   3165
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   11010
   Begin prjChameleon.chameleonButton cmdNovo 
      Height          =   360
      Left            =   5805
      TabIndex        =   11
      ToolTipText     =   "Novo Registro"
      Top             =   6345
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "&Novo"
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
      MICON           =   "frmsc_unidade_energia.frx":0000
      PICN            =   "frmsc_unidade_energia.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdExcluir 
      Height          =   360
      Left            =   8325
      TabIndex        =   14
      ToolTipText     =   "Excluir Registro"
      Top             =   6345
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "&Excluir"
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
      MICON           =   "frmsc_unidade_energia.frx":0176
      PICN            =   "frmsc_unidade_energia.frx":0192
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   45
      TabIndex        =   23
      Top             =   45
      Width           =   10905
      Begin VB.TextBox txtFilter 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1350
         TabIndex        =   0
         Top             =   195
         Width           =   8160
      End
      Begin prjChameleon.chameleonButton cmdFiltrar 
         Default         =   -1  'True
         Height          =   315
         Left            =   9765
         TabIndex        =   1
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
         MICON           =   "frmsc_unidade_energia.frx":0234
         PICN            =   "frmsc_unidade_energia.frx":0250
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
         TabIndex        =   24
         Top             =   225
         Width           =   1275
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1320
      Left            =   45
      TabIndex        =   12
      Top             =   4950
      Width           =   10905
      Begin VB.TextBox txtEmpenho 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5940
         MaxLength       =   6
         TabIndex        =   4
         Top             =   180
         Width           =   1545
      End
      Begin VB.TextBox txtDia 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3510
         MaxLength       =   2
         TabIndex        =   3
         Top             =   180
         Width           =   465
      End
      Begin VB.TextBox txtEndereco 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1035
         MaxLength       =   250
         TabIndex        =   8
         Top             =   900
         Width           =   6450
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1035
         MaxLength       =   250
         TabIndex        =   6
         Top             =   540
         Width           =   6450
      End
      Begin VB.TextBox txtLigacao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8640
         MaxLength       =   30
         TabIndex        =   5
         Top             =   180
         Width           =   1950
      End
      Begin VB.ComboBox cmbSecretaria 
         Height          =   315
         Left            =   8640
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   540
         Width           =   1995
      End
      Begin VB.TextBox txtDotacao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8640
         MaxLength       =   6
         TabIndex        =   9
         Top             =   900
         Width           =   1950
      End
      Begin VB.Label Label2 
         Caption         =   "N� do Emprenho..:"
         Height          =   195
         Index           =   2
         Left            =   4455
         TabIndex        =   28
         Top             =   225
         Width           =   1365
      End
      Begin VB.Label Label3 
         Caption         =   "Vencimento dia:"
         Height          =   195
         Left            =   2295
         TabIndex        =   27
         Top             =   225
         Width           =   1185
      End
      Begin VB.Label Label2 
         Caption         =   "C�d:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Top             =   225
         Width           =   375
      End
      Begin VB.Label lblCod 
         Caption         =   "0000"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   495
         TabIndex        =   21
         Top             =   225
         Width           =   510
      End
      Begin VB.Label Label2 
         Caption         =   "N� Liga��o:"
         Height          =   195
         Index           =   1
         Left            =   7740
         TabIndex        =   20
         Top             =   225
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "Secretaria.:"
         Height          =   195
         Index           =   3
         Left            =   7740
         TabIndex        =   19
         Top             =   585
         Width           =   870
      End
      Begin VB.Label Label2 
         Caption         =   "Descri��o..:"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   18
         Top             =   585
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "Endere�o...:"
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   17
         Top             =   945
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "Dota��o...:"
         Height          =   195
         Index           =   8
         Left            =   7740
         TabIndex        =   16
         Top             =   945
         Width           =   870
      End
   End
   Begin prjChameleon.chameleonButton cmdAlterar 
      Height          =   360
      Left            =   7065
      TabIndex        =   13
      ToolTipText     =   "Editar Registro"
      Top             =   6345
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
      MICON           =   "frmsc_unidade_energia.frx":042A
      PICN            =   "frmsc_unidade_energia.frx":0446
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
      Left            =   180
      TabIndex        =   10
      ToolTipText     =   "Exibir consumo da unidade"
      Top             =   6345
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
      MICON           =   "frmsc_unidade_energia.frx":05A0
      PICN            =   "frmsc_unidade_energia.frx":05BC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin vbAcceleratorSGrid6.vbalGrid grdMain 
      Height          =   4200
      Left            =   45
      TabIndex        =   2
      Top             =   675
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
   Begin prjChameleon.chameleonButton cmdGravar 
      Height          =   360
      Left            =   8325
      TabIndex        =   25
      ToolTipText     =   "Gravar os Dados"
      Top             =   6345
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
      MICON           =   "frmsc_unidade_energia.frx":095A
      PICN            =   "frmsc_unidade_energia.frx":0976
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
      TabIndex        =   26
      ToolTipText     =   "Cancelar Edi��o"
      Top             =   6345
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
      MICON           =   "frmsc_unidade_energia.frx":0D1B
      PICN            =   "frmsc_unidade_energia.frx":0D37
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
      Left            =   9600
      TabIndex        =   15
      ToolTipText     =   "Sair da Tela"
      Top             =   6345
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
      MICON           =   "frmsc_unidade_energia.frx":0E91
      PICN            =   "frmsc_unidade_energia.frx":0EAD
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
Attribute VB_Name = "frmsc_unidade_energia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Evento As String

Private Sub cmdAlterar_Click()
If grdMain.Rows = 0 Then
    MsgBox "Selecione uma unidade!", vbExclamation, "Erro"
    Exit Sub
End If

Evento = "ALTERAR"
Eventos "INCLUIR"
End Sub

Private Sub cmdCancel_Click()
Eventos "INICIAR"
If grdMain.Rows > 0 Then
    grdMain.SelectedRow = 1
    Le
End If
End Sub

Private Sub cmdConsumo_Click()
Dim nCodigo As Integer
nCodigo = Val(lblCod.Caption)
If nCodigo = 0 Then
    MsgBox "Selecione uma unidade", vbCritical, "Erro"
Else
    Set frm = frmsc_consumo
    frm.sForm = Me.Name
    frm.show vbModal
End If

End Sub

Private Sub cmdExcluir_Click()
Dim Sql As String, nCodigo As Integer

nCodigo = Val(lblCod.Caption)
If nCodigo = 0 Then
    MsgBox "Selecione uma unidade", vbCritical, "Erro"
Else
    If MsgBox("Voc� deseja excluir esta unidade e todo o seu consumo?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirma��o") = vbYes Then
        Sql = "delete from sc_ligacao_energia where codigo=" & nCodigo
        cn.Execute Sql, rdExecDirect
        Sql = "delete from sc_ligacao_energia_consumo where codigo=" & nCodigo
        cn.Execute Sql, rdExecDirect
        grdMain.RemoveRow (grdMain.SelectedRow)
    End If
End If

End Sub

Private Sub cmdFiltrar_Click()
Dim s As String

s = Trim(UCase(txtFilter.Text))
If Len(s) = 0 Then
    CarregaLista ""
    Exit Sub
End If
If Len(s) < 3 Then
    MsgBox "Digite ao menos 3 caracteres.", vbExclamation, "Aten��o"
    Exit Sub
End If

CarregaLista s

End Sub

Private Sub cmdGravar_Click()
Dim nSec As Integer, nCodigo As Integer, nRow As Integer, RdoAux As rdoResultset, x As Integer, bFind As Boolean

If Trim(txtNome.Text) = "" Then
    MsgBox "Digite a descri��o da unidade", vbCritical, "Erro"
    Exit Sub
End If

If Trim(txtLigacao.Text) = "" Then
    MsgBox "Digite o n� da liga��o", vbCritical, "Erro"
    Exit Sub
End If

If cmbSecretaria.ListIndex = 0 Then
    MsgBox "Selecione a secretaria", vbCritical, "Erro"
    Exit Sub
End If

If Val(txtDia.Text) > 31 Then
    MsgBox "Dia de vencimento inv�lido", vbCritical, "Erro"
    Exit Sub
End If

If Evento = "NOVO" Then
    bFind = False
    For x = 1 To grdMain.Rows
        If grdMain.cell(x, 6).Text = UCase(txtLigacao.Text) Then
            bFind = True
        End If
    Next
    If bFind Then
        If MsgBox("J� existe um cadastro com este n� de liga��o! Voc� deseja criar outro cadastro como o mesmo n�mero?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirma��o") = vbNo Then
            Exit Sub
        End If
    End If
End If

nRow = grdMain.SelectedRow
nSec = cmbSecretaria.ItemData(cmbSecretaria.ListIndex)
If Evento = "ALTERAR" Then
    nCodigo = Val(lblCod.Caption)
    Sql = "update sc_ligacao_energia set secretaria=" & nSec & ",dotacao=" & sNullVal(txtDotacao.Text) & ",nome=" & sNull(txtNome.Text) & ","
    Sql = Sql & "endereco=" & sNull(txtEndereco.Text) & ",empenho=" & sNullVal(txtEmpenho.Text) & ",dia=" & sNullVal(txtDia.Text) & " where codigo=" & nCodigo
    cn.Execute Sql, rdExecDirect
    grdMain.cell(nRow, 2).Text = nSec
    grdMain.cell(nRow, 3).Text = cmbSecretaria.Text
    grdMain.cell(nRow, 7).Text = txtDotacao.Text
    grdMain.cell(nRow, 8).Text = txtEmpenho.Text
    grdMain.cell(nRow, 9).Text = txtDia.Text
    grdMain.cell(nRow, 4).Text = txtNome.Text
    grdMain.cell(nRow, 5).Text = txtEndereco.Text
Else
    Sql = "select max(codigo) as maximo from sc_ligacao_energia"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If IsNull(RdoAux!maximo) Then
        nCodigo = 1
    Else
        nCodigo = RdoAux!maximo + 1
    End If
    RdoAux.Close
    lblCod.Caption = Format(nCodigo, "000")
    
    Sql = "insert sc_ligacao_energia(codigo,secretaria, ligacao,nome,endereco,dotacao,empenho,dia) values(" & nCodigo & "," & nSec & "," & sNull(txtLigacao.Text) & ","
    Sql = Sql & sNull(txtNome.Text) & "," & sNull(txtEndereco.Text) & "," & sNullVal(txtDotacao.Text) & "," & sNullVal(txtEmpenho.Text) & "," & sNullVal(txtDia.Text) & ")"
    cn.Execute Sql, rdExecDirect
    grdMain.AddRow
    grdMain.CellDetails grdMain.Rows, 1, nCodigo, DT_RIGHT
    grdMain.CellDetails grdMain.Rows, 2, nSec
    grdMain.CellDetails grdMain.Rows, 3, cmbSecretaria.Text
    grdMain.CellDetails grdMain.Rows, 4, UCase(txtNome.Text)
    grdMain.CellDetails grdMain.Rows, 5, UCase(txtEndereco.Text)
    grdMain.CellDetails grdMain.Rows, 6, UCase(txtLigacao.Text)
    grdMain.CellDetails grdMain.Rows, 7, txtDotacao.Text
    grdMain.CellDetails grdMain.Rows, 8, txtEmpenho.Text
    grdMain.CellDetails grdMain.Rows, 9, txtDia.Text
End If


Eventos "INICIAR"

End Sub

Private Sub cmdNovo_Click()
Evento = "NOVO"
Eventos "INCLUIR"
Limpa
txtLigacao.SetFocus
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Sql As String, RdoAux As rdoResultset

Centraliza Me
GridHeader

Sql = "select codigo,sigla from sc_secretaria order by nome"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbSecretaria.AddItem !Sigla
        cmbSecretaria.ItemData(cmbSecretaria.NewIndex) = !Codigo
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
    .AddColumn "kCod", "C�d", ecgHdrTextALignCentre, , 40
    .AddColumn "kScd", "CS", , 10, , False
    .AddColumn "kSec", "Secret.", , 80
    .AddColumn "kUsu", "Descri��o da Unidade", ecgHdrTextALignLeft, , 250
    .AddColumn "kEnd", "Endere�o", ecgHdrTextALignLeft, , 240
    .AddColumn "kLig", "Liga��o", ecgHdrTextALignLeft, , 90
    .AddColumn "kDot", "Dota��o", ecgHdrTextALignLeft, , 80
    .AddColumn "kEmp", "Empenho", ecgHdrTextALignLeft, , 80
    .AddColumn "kDia", "Dia", ecgHdrTextALignLeft, , 40
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

Private Sub Eventos(Tipo As String)

If Tipo = "INICIAR" Then
   cmdNovo.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdConsumo.Visible = True
   cmdSair.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   cmbSecretaria.Enabled = False
   cmbSecretaria.BackColor = Kde
   txtDotacao.Locked = True
   txtDotacao.BackColor = Kde
   txtNome.Locked = True
   txtNome.BackColor = Kde
   txtEndereco.Locked = True
   txtEndereco.BackColor = Kde
   txtLigacao.Locked = True
   txtLigacao.BackColor = Kde
   txtEmpenho.Locked = True
   txtEmpenho.BackColor = Kde
   txtDia.Locked = True
   txtDia.BackColor = Kde
   grdMain.Enabled = True
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdConsumo.Visible = False
   cmdSair.Visible = False
   cmbSecretaria.Enabled = True
   cmbSecretaria.BackColor = Branco
   txtDotacao.Locked = False
   txtDotacao.BackColor = Branco
   txtEmpenho.Locked = False
   txtEmpenho.BackColor = Branco
   txtDia.Locked = False
   txtDia.BackColor = Branco
   txtNome.Locked = False
   txtNome.BackColor = Branco
   txtEndereco.Locked = False
   txtEndereco.BackColor = Branco
   txtLigacao.Locked = False
   txtLigacao.BackColor = Branco
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   grdMain.Enabled = False
End If

End Sub

Private Sub CarregaLista(Filter As String)
Dim Sql As String, RdoAux As rdoResultset

grdMain.Clear

Sql = "SELECT sc_ligacao_energia.codigo,sc_ligacao_energia.nome,sc_ligacao_energia.endereco,sc_ligacao_energia.dotacao,sc_ligacao_energia.ligacao,sc_secretaria.sigla,sc_secretaria.codigo AS codigo_secretaria, sc_ligacao_energia.empenho, sc_ligacao_energia.dia "
Sql = Sql & "FROM sc_ligacao_energia INNER JOIN sc_secretaria ON sc_ligacao_energia.secretaria = sc_secretaria.codigo"
If Filter <> "" Then
    Sql = Sql & " where sc_ligacao_energia.nome like '%" & Mask(Filter) & "%' or endereco like '%" & Mask(Filter) & "%' or ligacao like '%" & Mask(Filter) & "%' or sigla like '%" & Mask(Filter) & "%'"
End If
Sql = Sql & " ORDER BY sc_ligacao_energia.nome "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        grdMain.AddRow
        grdMain.CellDetails grdMain.Rows, 1, !Codigo, DT_RIGHT
        grdMain.CellDetails grdMain.Rows, 2, !codigo_secretaria
        grdMain.CellDetails grdMain.Rows, 3, !Sigla
        grdMain.CellDetails grdMain.Rows, 4, !Nome
        grdMain.CellDetails grdMain.Rows, 5, UCase(SubNull(!Endereco))
        grdMain.CellDetails grdMain.Rows, 6, SubNull(!ligacao)
        grdMain.CellDetails grdMain.Rows, 7, SubNull(!dotacao)
        grdMain.CellDetails grdMain.Rows, 8, SubNull(!empenho)
        grdMain.CellDetails grdMain.Rows, 9, SubNull(!DIA)
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


Private Sub txtDia_KeyPress(KeyAscii As Integer)
Tweak txtDia, KeyAscii, IntegerPositive
End Sub

Private Sub txtDotacao_KeyPress(KeyAscii As Integer)
Tweak txtDotacao, KeyAscii, IntegerPositive
End Sub

Private Sub txtEmpenho_KeyPress(KeyAscii As Integer)
Tweak txtEmpenho, KeyAscii, IntegerPositive
End Sub

Private Sub txtFilter_Change()
If txtFilter.Text = "" Then CarregaLista ""
End Sub

Private Sub Le()
Dim Row As Integer, sCompl As String, nSec As Integer, x As Integer

Row = grdMain.SelectedRow
If Row = 0 Then Exit Sub
nSec = Val(grdMain.cell(Row, 2).Text)

lblCod.Caption = Format(grdMain.cell(Row, 1).Text, "0000")
txtLigacao.Text = Format(grdMain.cell(Row, 6).Text, "0000000")
txtNome.Text = grdMain.cell(Row, 4).Text
txtEndereco.Text = grdMain.cell(Row, 5).Text
txtDotacao.Text = grdMain.cell(Row, 7).Text
txtEmpenho.Text = grdMain.cell(Row, 8).Text
txtDia.Text = grdMain.cell(Row, 9).Text
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

Private Sub Limpa()
txtNome.Text = ""
txtEndereco.Text = ""
txtDotacao.Text = ""
txtLigacao.Text = ""
txtEmpenho.Text = ""
txtDia.Text = ""
cmbSecretaria.ListIndex = 0
lblCod.Caption = "000"

End Sub


