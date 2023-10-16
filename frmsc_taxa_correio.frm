VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form frmsc_taxa_correio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle das Unidades - Taxas do Correio"
   ClientHeight    =   5775
   ClientLeft      =   7275
   ClientTop       =   4185
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   7305
   Begin prjChameleon.chameleonButton cmdConsumo 
      Height          =   360
      Left            =   90
      TabIndex        =   14
      ToolTipText     =   "Exibir consumo da unidade"
      Top             =   5310
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
      MICON           =   "frmsc_taxa_correio.frx":0000
      PICN            =   "frmsc_taxa_correio.frx":001C
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
      BackColor       =   &H00EEEEEE&
      Height          =   1005
      Left            =   45
      TabIndex        =   1
      Top             =   4215
      Width           =   7230
      Begin VB.TextBox txtDotacao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3780
         MaxLength       =   6
         TabIndex        =   12
         Top             =   585
         Width           =   1950
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   765
         Locked          =   -1  'True
         MaxLength       =   200
         TabIndex        =   3
         Top             =   225
         Width           =   6345
      End
      Begin VB.TextBox txtSigla 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   765
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   2
         Top             =   585
         Width           =   1755
      End
      Begin VB.Label Label2 
         Caption         =   "Dotação...:"
         Height          =   195
         Index           =   8
         Left            =   2835
         TabIndex        =   13
         Top             =   630
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome...:"
         Height          =   195
         Index           =   11
         Left            =   90
         TabIndex        =   7
         Top             =   225
         Width           =   645
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sigla.....:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   585
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Código:"
         Height          =   195
         Index           =   0
         Left            =   6120
         TabIndex        =   5
         Top             =   630
         Width           =   645
      End
      Begin VB.Label lblCod 
         Alignment       =   1  'Right Justify
         Caption         =   "000"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   6705
         TabIndex        =   4
         Top             =   630
         Width           =   330
      End
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   360
      Left            =   6015
      TabIndex        =   0
      ToolTipText     =   "Sair da Tela"
      Top             =   5310
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
      MICON           =   "frmsc_taxa_correio.frx":03BA
      PICN            =   "frmsc_taxa_correio.frx":03D6
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
      Left            =   6015
      TabIndex        =   8
      ToolTipText     =   "Cancelar Edição"
      Top             =   5310
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
      MICON           =   "frmsc_taxa_correio.frx":0444
      PICN            =   "frmsc_taxa_correio.frx":0460
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdAlterar 
      Height          =   360
      Left            =   1350
      TabIndex        =   9
      ToolTipText     =   "Editar Registro"
      Top             =   5310
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
      MICON           =   "frmsc_taxa_correio.frx":05BA
      PICN            =   "frmsc_taxa_correio.frx":05D6
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
      Left            =   4740
      TabIndex        =   10
      ToolTipText     =   "Gravar os Dados"
      Top             =   5310
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
      MICON           =   "frmsc_taxa_correio.frx":0730
      PICN            =   "frmsc_taxa_correio.frx":074C
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
      Height          =   4155
      Left            =   45
      TabIndex        =   11
      Top             =   90
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   7329
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      NoFocusHighlightForeColor=   16777215
      NoFocusHighlightBackColor=   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderButtons   =   0   'False
      HeaderDragReorderColumns=   0   'False
      HeaderHotTrack  =   0   'False
      HeaderFlat      =   -1  'True
      BorderStyle     =   0
      ScrollBarStyle  =   1
      Editable        =   -1  'True
      DisableIcons    =   -1  'True
   End
End
Attribute VB_Name = "frmsc_taxa_correio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Evento As String

Private Sub cmdAlterar_Click()
    If grdMain.SelectedRow < 1 Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
    Eventos "INCLUIR"
    Evento = "Alterar"
End Sub

Private Sub cmdCancel_Click()
    Le
    Eventos "INICIAR"
    Evento = ""
End Sub

Private Sub cmdExcluir_Click()
    Dim nCodigo As Integer

    If grdMain.SelectedRow = 0 Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
    
    nCodigo = grdMain.cell(grdMain.SelectedRow, 1).Text
    
    If MsgBox("Excluir o local: " & vbCrLf & txtNome.Text & "?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
'       sql = "SELECT * FROM CADASTRORURALPRODUTO WHERE CODPRODUTO=" & Val(txtCod.Text)
'       Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
'       If RdoAux.RowCount = 0 Then
            Sql = "DELETE FROM sc_secretaria WHERE CODIGO=" & nCodigo
            cn.Execute Sql, rdExecDirect
            CarregaLista
'       Else
'            MsgBox "Este produto esta sendo utilizado e não pode ser excluido.", vbExclamation, "Atenção"
'       End If
    End If

End Sub

Private Sub cmdConsumo_Click()
frmsc_consumo_correio.show vbModal
End Sub

Private Sub cmdGravar_Click()
    If Trim(txtNome.Text) = "" Then
       MsgBox "Favor digitar o nome do local de consumo.", vbExclamation, "Atenção"
       txtNome.SetFocus
       Exit Sub
    End If
    Grava
    CarregaLista
    Eventos "INICIAR"

End Sub

Private Sub cmdNovo_Click()
    Limpa
    Eventos "INCLUIR"
    Evento = "Novo"
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
GridHeader
CarregaLista
Le
Eventos "INICIAR"

End Sub

Private Sub Eventos(Tipo As String)

If Tipo = "INICIAR" Then
   cmdConsumo.Visible = True
   cmdAlterar.Visible = True
   cmdSair.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   txtDotacao.Enabled = True
   txtDotacao.BackColor = Kde
   grdMain.Enabled = True
ElseIf Tipo = "INCLUIR" Then
   cmdConsumo.Visible = False
   cmdAlterar.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   grdMain.Enabled = False
   txtDotacao.Enabled = True
   txtDotacao.BackColor = Branco
   txtDotacao.SetFocus
End If

End Sub

Private Sub Limpa()
txtNome.Text = ""
txtSigla.Text = ""
End Sub

Private Sub CarregaLista()
Dim Sql As String, RdoAux As rdoResultset

grdMain.Clear
Sql = "SELECT sc_secretaria.codigo,sc_secretaria.nome,sc_secretaria.sigla,sc_taxa_correio.dotacao From dbo.sc_secretaria "
Sql = Sql & "LEFT OUTER JOIN dbo.sc_taxa_correio ON sc_secretaria.codigo = sc_taxa_correio.codigo where sc_secretaria.codigo>0 order by nome"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        grdMain.AddRow
        grdMain.CellDetails grdMain.Rows, 1, !codigo, DT_RIGHT
        grdMain.CellDetails grdMain.Rows, 2, !Nome
        grdMain.CellDetails grdMain.Rows, 3, !Sigla
        grdMain.CellDetails grdMain.Rows, 4, SubNull(!dotacao)
       .MoveNext
    Loop
   .Close
End With

grdMain.SelectedRow = 1

End Sub

Private Sub Le()
If grdMain.SelectedRow > 0 Then
    lblCod.Caption = grdMain.cell(grdMain.SelectedRow, 1).Text
    txtNome.Text = grdMain.cell(grdMain.SelectedRow, 2).Text
    txtSigla.Text = grdMain.cell(grdMain.SelectedRow, 3).Text
    txtDotacao.Text = grdMain.cell(grdMain.SelectedRow, 4).Text
End If
End Sub

Private Sub grdMain_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
Le
End Sub

Private Sub txtDotacao_KeyPress(KeyAscii As Integer)
Tweak txtDotacao, KeyAscii, IntegerPositive
End Sub

Private Sub txtNome_GotFocus()
txtNome.SelStart = 0
txtNome.SelLength = Len(txtNome.Text)
End Sub

Private Sub Grava()
Dim nCodigo As Integer

nCodigo = Val(lblCod.Caption)
Sql = "delete from sc_taxa_correio where codigo=" & nCodigo
cn.Execute Sql, rdExecDirect

Sql = "INSERT sc_taxa_correio (CODIGO,DOTACAO) VALUES("
Sql = Sql & nCodigo & "," & sNullVal(txtDotacao.Text) & ")"
cn.Execute Sql, rdExecDirect
      
grdMain.cell(grdMain.SelectedRow, 4).Text = txtDotacao.Text
      
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
    .AddColumn "kNom", "Nome da secretaria", ecgHdrTextALignLeft, , 320
    .AddColumn "kSig", "Sigla", ecgHdrTextALignLeft, , 100
    .AddColumn "kDot", "Dotação", ecgHdrTextALignLeft, , 80
End With

End Sub

Private Sub txtSigla_GotFocus()
txtSigla.SelStart = 0
txtSigla.SelLength = Len(txtSigla.Text)

End Sub

