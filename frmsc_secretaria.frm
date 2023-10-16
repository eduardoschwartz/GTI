VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form frmsc_secretaria 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro das Secretarias"
   ClientHeight    =   5775
   ClientLeft      =   16470
   ClientTop       =   5190
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   7290
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   6105
      TabIndex        =   8
      ToolTipText     =   "Sair da Tela"
      Top             =   5355
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
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
      MICON           =   "frmsc_secretaria.frx":0000
      PICN            =   "frmsc_secretaria.frx":001C
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
      TabIndex        =   9
      Top             =   4215
      Width           =   7230
      Begin VB.TextBox txtSigla 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   765
         MaxLength       =   20
         TabIndex        =   2
         Top             =   585
         Width           =   1755
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   765
         MaxLength       =   200
         TabIndex        =   1
         Top             =   225
         Width           =   6345
      End
      Begin VB.Label lblCod 
         Alignment       =   1  'Right Justify
         Caption         =   "000"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   6705
         TabIndex        =   13
         Top             =   630
         Width           =   330
      End
      Begin VB.Label Label2 
         Caption         =   "Código:"
         Height          =   195
         Left            =   6165
         TabIndex        =   12
         Top             =   630
         Width           =   645
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sigla.....:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   11
         Top             =   585
         Width           =   645
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome...:"
         Height          =   195
         Index           =   11
         Left            =   90
         TabIndex        =   10
         Top             =   225
         Width           =   645
      End
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   6105
      TabIndex        =   4
      ToolTipText     =   "Cancelar Edição"
      Top             =   5355
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
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
      MICON           =   "frmsc_secretaria.frx":008A
      PICN            =   "frmsc_secretaria.frx":00A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdNovo 
      Height          =   315
      Left            =   90
      TabIndex        =   5
      ToolTipText     =   "Novo Registro"
      Top             =   5355
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
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
      MICON           =   "frmsc_secretaria.frx":0200
      PICN            =   "frmsc_secretaria.frx":021C
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
      Height          =   315
      Left            =   1185
      TabIndex        =   6
      ToolTipText     =   "Editar Registro"
      Top             =   5355
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
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
      MICON           =   "frmsc_secretaria.frx":0376
      PICN            =   "frmsc_secretaria.frx":0392
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
      Height          =   315
      Left            =   2280
      TabIndex        =   7
      ToolTipText     =   "Excluir Registro"
      Top             =   5355
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
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
      MICON           =   "frmsc_secretaria.frx":04EC
      PICN            =   "frmsc_secretaria.frx":0508
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
      Height          =   315
      Left            =   5010
      TabIndex        =   3
      ToolTipText     =   "Gravar os Dados"
      Top             =   5355
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
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
      MICON           =   "frmsc_secretaria.frx":05AA
      PICN            =   "frmsc_secretaria.frx":05C6
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
      TabIndex        =   0
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
Attribute VB_Name = "frmsc_secretaria"
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
            sql = "DELETE FROM sc_secretaria WHERE CODIGO=" & nCodigo
            cn.Execute sql, rdExecDirect
            CarregaLista
'       Else
'            MsgBox "Este produto esta sendo utilizado e não pode ser excluido.", vbExclamation, "Atenção"
'       End If
    End If

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
   cmdNovo.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdSair.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   txtNome.Enabled = True
   txtNome.BackColor = Kde
   txtSigla.Enabled = True
   txtSigla.BackColor = Kde
   grdMain.Enabled = True
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   grdMain.Enabled = False
   txtSigla.Enabled = True
   txtSigla.BackColor = Branco
   txtNome.Enabled = True
   txtNome.BackColor = Branco
   txtNome.SetFocus
End If

End Sub

Private Sub Limpa()
txtNome.Text = ""
txtSigla.Text = ""
End Sub

Private Sub CarregaLista()
Dim sql As String, RdoAux As rdoResultset

grdMain.Clear
sql = "Select codigo,nome,sigla from sc_secretaria where codigo>0  order by nome"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        grdMain.AddRow
        grdMain.CellDetails grdMain.Rows, 1, !codigo, DT_RIGHT
        grdMain.CellDetails grdMain.Rows, 2, !Nome
        grdMain.CellDetails grdMain.Rows, 3, !Sigla
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
End If
End Sub

Private Sub grdMain_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
Le
End Sub

Private Sub txtNome_GotFocus()
txtNome.SelStart = 0
txtNome.SelLength = Len(txtNome.Text)
End Sub

Private Sub Grava()
Dim nCodNovo As Integer

If Evento = "Novo" Then
    sql = "SELECT MAX(CODIGO) AS MAXIMO FROM sc_secretaria"
    Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        nCodNovo = !maximo + 1
       .Close
    End With
    sql = "INSERT sc_secretaria (CODIGO,NOME,SIGLA) VALUES("
    sql = sql & nCodNovo & ",'" & UCase$(Mask(txtNome.Text)) & "','" & UCase$(Mask(txtSigla.Text)) & "')"
Else
    nCodNovo = grdMain.cell(grdMain.SelectedRow, 1).Text
    sql = "UPDATE sc_secretaria SET NOME='" & UCase$(Mask(txtNome.Text)) & "',SIGLA='" & UCase$(Mask(txtSigla.Text)) & "'"
    sql = sql & " WHERE CODIGO=" & nCodNovo
End If
cn.Execute sql, rdExecDirect
      
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
End With

End Sub

Private Sub txtSigla_GotFocus()
txtSigla.SelStart = 0
txtSigla.SelLength = Len(txtSigla.Text)

End Sub
