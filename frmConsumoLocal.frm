VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmsc_secretaria 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro das Secretarias"
   ClientHeight    =   5280
   ClientLeft      =   2385
   ClientTop       =   3510
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5280
   ScaleWidth      =   5640
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   4530
      TabIndex        =   9
      ToolTipText     =   "Sair da Tela"
      Top             =   4860
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
      MICON           =   "frmConsumoLocal.frx":0000
      PICN            =   "frmConsumoLocal.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox lstMain 
      Appearance      =   0  'Flat
      Height          =   4125
      Left            =   45
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   90
      Width           =   5535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Height          =   600
      Left            =   0
      TabIndex        =   5
      Top             =   4215
      Width           =   5655
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   765
         MaxLength       =   500
         TabIndex        =   6
         Top             =   225
         Width           =   4770
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
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   4530
      TabIndex        =   0
      ToolTipText     =   "Cancelar Edição"
      Top             =   4860
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
      MICON           =   "frmConsumoLocal.frx":008A
      PICN            =   "frmConsumoLocal.frx":00A6
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
      TabIndex        =   1
      ToolTipText     =   "Novo Registro"
      Top             =   4860
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
      MICON           =   "frmConsumoLocal.frx":0200
      PICN            =   "frmConsumoLocal.frx":021C
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
      TabIndex        =   2
      ToolTipText     =   "Editar Registro"
      Top             =   4860
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
      MICON           =   "frmConsumoLocal.frx":0376
      PICN            =   "frmConsumoLocal.frx":0392
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
      TabIndex        =   3
      ToolTipText     =   "Excluir Registro"
      Top             =   4860
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
      MICON           =   "frmConsumoLocal.frx":04EC
      PICN            =   "frmConsumoLocal.frx":0508
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
      Left            =   3435
      TabIndex        =   4
      ToolTipText     =   "Gravar os Dados"
      Top             =   4860
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
      MICON           =   "frmConsumoLocal.frx":05AA
      PICN            =   "frmConsumoLocal.frx":05C6
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
Attribute VB_Name = "frmsc_secretaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Evento As String

Private Sub cmdAlterar_Click()
    If lstMain.ListIndex = -1 Then
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

    If lstMain.ListIndex = -1 Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
    
    nCodigo = lstMain.ItemData(lstMain.ListIndex)
    
    If MsgBox("Excluir o local: " & vbCrLf & txtNome.Text & "?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
'       sql = "SELECT * FROM CADASTRORURALPRODUTO WHERE CODPRODUTO=" & Val(txtCod.Text)
'       Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
'       If RdoAux.RowCount = 0 Then
            Sql = "DELETE FROM CONSUMO_LOCAL WHERE CODIGO=" & nCodigo
            cn.Execute Sql, rdExecDirect
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
CarregaLista
Le
Eventos "INICIAR"

End Sub

Private Sub Eventos(Tipo As String)

Dim Ct As Control

If Tipo = "INICIAR" Then
   cmdNovo.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdSair.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   txtNome.Enabled = True
   txtNome.BackColor = Kde
   lstMain.Enabled = True
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   lstMain.Enabled = False
   txtNome.Enabled = True
   txtNome.BackColor = Branco
   txtNome.SetFocus
End If

End Sub

Private Sub Limpa()
txtNome.Text = ""
End Sub

Private Sub CarregaLista()
Dim Sql As String, RdoAux As rdoResultset

Sql = "Select codigo,nome from consumo_local order by nome"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
lstMain.Clear
With RdoAux
    Do Until .EOF
       lstMain.AddItem !Nome
       lstMain.ItemData(lstMain.NewIndex) = !codigo
      .MoveNext
    Loop
   .Close
End With

If lstMain.ListCount > 0 Then
    lstMain.ListIndex = 0
End If

End Sub

Private Sub Le()
If lstMain.ListIndex > -1 Then
    txtNome.Text = lstMain.Text
End If
End Sub

Private Sub lstMain_Click()
Le
End Sub

Private Sub txtNome_GotFocus()
txtNome.SelStart = 0
txtNome.SelLength = Len(txtNome.Text)
End Sub

Private Sub Grava()
Dim nCodNovo As Integer

If Evento = "Novo" Then
    Sql = "SELECT MAX(CODIGO) AS MAXIMO FROM CONSUMO_LOCAL"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        nCodNovo = !maximo + 1
       .Close
    End With
    lstMain.AddItem txtNome.Text
    lstMain.ItemData(lstMain.NewIndex) = nCodNovo
    Sql = "INSERT CONSUMO_LOCAL (CODIGO,NOME) VALUES("
    Sql = Sql & nCodNovo & ",'" & UCase$(Mask(txtNome.Text)) & "')"
Else
    nCodNovo = lstMain.ItemData(lstMain.ListIndex)
    Sql = "UPDATE CONSUMO_LOCAL SET NOME='" & UCase$(Mask(txtNome.Text)) & "'"
    Sql = Sql & " WHERE CODIGO=" & nCodNovo
End If
cn.Execute Sql, rdExecDirect
      
End Sub
