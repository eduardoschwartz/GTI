VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmEmpresaAtividade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empresas por atividade"
   ClientHeight    =   8085
   ClientLeft      =   6555
   ClientTop       =   3945
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8085
   ScaleWidth      =   8715
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1350
      TabIndex        =   1
      Top             =   1080
      Width           =   5865
   End
   Begin VB.Frame Frame1 
      Height          =   2445
      Left            =   135
      TabIndex        =   4
      Top             =   5085
      Width           =   8475
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1980
         Width           =   7080
      End
      Begin VB.TextBox txtFone 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1620
         Width           =   7080
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1260
         Width           =   7080
      End
      Begin VB.TextBox txtEndereco 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   900
         Width           =   7080
      End
      Begin VB.TextBox txtAtividade 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   540
         Width           =   7080
      End
      Begin VB.TextBox txtRazao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   180
         Width           =   7080
      End
      Begin VB.Label Label2 
         Caption         =   "Emaill: "
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   15
         Top             =   2025
         Width           =   1005
      End
      Begin VB.Label txtFone55 
         Caption         =   "Telefone:"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   13
         Top             =   1665
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "Contato:"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   11
         Top             =   1305
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "Endereço: "
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   9
         Top             =   945
         Width           =   1005
      End
      Begin VB.Label Label3 
         Caption         =   "Atividade:"
         Height          =   195
         Left            =   135
         TabIndex        =   7
         Top             =   585
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "Razão Social: "
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   5
         Top             =   225
         Width           =   1005
      End
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   3480
      Left            =   135
      TabIndex        =   3
      Top             =   1575
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   6138
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Razão Social"
         Object.Width           =   14111
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Atividade"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Nome"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Fone"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Email"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Endereco"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.ComboBox cmbAtividade 
      Height          =   315
      Left            =   1350
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   630
      Width           =   7260
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   7515
      TabIndex        =   17
      ToolTipText     =   "Sair da Tela"
      Top             =   7650
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
      MICON           =   "frmEmpresaContato.frx":0000
      PICN            =   "frmEmpresaContato.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdConsultar 
      Height          =   315
      Left            =   7335
      TabIndex        =   19
      ToolTipText     =   "Consulta Cidadãos Cadastrados"
      Top             =   1080
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "C&onsultar"
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
      MICON           =   "frmEmpresaContato.frx":008A
      PICN            =   "frmEmpresaContato.frx":00A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Lista das empesas por atividade OU busque uma empresa pela razão social."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   225
      TabIndex        =   20
      Top             =   180
      Width           =   8250
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Razão Social:"
      Height          =   240
      Left            =   180
      TabIndex        =   18
      Top             =   1125
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Atividade..:"
      Height          =   195
      Left            =   225
      TabIndex        =   2
      Top             =   675
      Width           =   1005
   End
End
Attribute VB_Name = "frmEmpresaAtividade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdConsultar_Click()
Dim Sql As String, RdoAux As rdoResultset, x As Integer

If Trim(txtSearch.Text) = "" Then
    MsgBox "Digite um nome", vbCritical, "erro"
    Exit Sub
End If

Sql = "select * from mobiliario where razaosocial like '%" & Trim(txtSearch.Text) & "%' and dataencerramento is null"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount = 0 Then
    MsgBox "Nenhuma empresa encontrada com este nome", vbCritical, "Erro"
    Exit Sub
End If

lvMain.ListItems.Clear
cmbAtividade.ListIndex = -1
txtRazao.Text = RdoAux!RazaoSocial
txtAtividade.Text = SubNull(RdoAux!ativextenso)
txtNome.Text = SubNull(RdoAux!nomecontato)
txtFone.Text = SubNull(RdoAux!fonecontato)
txtEmail.Text = SubNull(RdoAux!emailcontato)

'MsgBox RdoAux!RazaoSocial & " - " & RdoAux!codatividade
'For x = 0 To cmbAtividade.ListCount - 1
'    If cmbAtividade.ItemData(x) = RdoAux!codatividade Then
'       cmbAtividade.ListIndex = x
'       Exit For
'    End If
'Next

'For x = 1 To lvMain.ListItems.Count
'    If Val(lvMain.SelectedItem.Text) = RdoAux!codigomob Then
'       lvMain.ListItems(x).Selected = True
'       Exit For
'    End If
'Next


End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Sql As String, RdoAux As rdoResultset

Centraliza Me

Sql = "select codatividade, descatividade from atividade order by descatividade"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbAtividade.AddItem (!descatividade)
        cmbAtividade.ItemData(cmbAtividade.NewIndex) = !codatividade
       .MoveNext
    Loop
   .Close
End With

cmbAtividade.ListIndex = 0

End Sub

Private Sub cmbAtividade_Click()
CarregaLista
End Sub

Private Sub CarregaLista()
Dim Sql As String, RdoAux As rdoResultset
Dim z As Long, itmX As ListItem, sEndereco As String
z = SendMessage(lvMain.HWND, LVM_DELETEALLITEMS, 0, 0)
If cmbAtividade.ListIndex = -1 Then Exit Sub
Sql = "select codigomob,razaosocial,ativextenso,nomecontato,fonecontato,emailcontato,LOGRADOURO,numero,descbairro,complemento from vwFULLEMPRESA where dataencerramento is null and codatividade=" & cmbAtividade.ItemData(cmbAtividade.ListIndex) & " order by razaosocial"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       sEndereco = !Logradouro & ", " & Val(SubNull(!Numero)) & " " & SubNull(!Complemento) & " - " & SubNull(!DescBairro)
    
       Set itmX = lvMain.ListItems.Add(, , !codigomob)
       itmX.SubItems(1) = !RazaoSocial
       itmX.SubItems(2) = SubNull(!ativextenso)
       itmX.SubItems(3) = SubNull(!nomecontato)
       itmX.SubItems(4) = SubNull(!fonecontato)
       itmX.SubItems(5) = SubNull(!emailcontato)
       itmX.SubItems(6) = sEndereco
       .MoveNext
    Loop
   .Close
End With
If lvMain.ListItems.Count > 0 Then
    lvMain.ListItems(1).Selected = True
    Le
End If

End Sub

Private Sub lvMain_Click()
Le
End Sub

Private Sub Le()
If lvMain.ListItems.Count = 0 Then Exit Sub

txtRazao.Text = lvMain.SelectedItem.SubItems(1)
txtAtividade.Text = lvMain.SelectedItem.SubItems(2)
txtEndereco.Text = lvMain.SelectedItem.SubItems(6)
txtNome.Text = lvMain.SelectedItem.SubItems(3)
txtFone.Text = lvMain.SelectedItem.SubItems(4)
txtEmail.Text = lvMain.SelectedItem.SubItems(5)


End Sub


Private Sub lvMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
Le
End Sub
