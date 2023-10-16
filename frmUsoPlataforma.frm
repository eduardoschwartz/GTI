VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmUsoPlataforma 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autorização Uso de Plataforma"
   ClientHeight    =   4230
   ClientLeft      =   6540
   ClientTop       =   4215
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4230
   ScaleWidth      =   6360
   Begin prjChameleon.chameleonButton cmdGravar 
      Height          =   315
      Left            =   5175
      TabIndex        =   4
      ToolTipText     =   "Gravar os Dados"
      Top             =   3780
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
      MICON           =   "frmUsoPlataforma.frx":0000
      PICN            =   "frmUsoPlataforma.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame3 
      Caption         =   "Selecione a(s) empresa(s) de ônibus"
      ForeColor       =   &H00000080&
      Height          =   2085
      Left            =   90
      TabIndex        =   7
      Top             =   1575
      Width           =   6180
      Begin VB.ListBox lstEmpresa 
         Appearance      =   0  'Flat
         Height          =   1605
         Left            =   90
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   315
         Width           =   5955
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selecione o Usuário"
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   90
      TabIndex        =   6
      Top             =   810
      Width           =   6180
      Begin VB.ComboBox cmbUser 
         Height          =   315
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   315
         Width           =   5955
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Usuário"
      ForeColor       =   &H00000080&
      Height          =   645
      Left            =   90
      TabIndex        =   5
      Top             =   135
      Width           =   6180
      Begin VB.OptionButton optTipo 
         Caption         =   "Usuário Externo"
         Height          =   195
         Index           =   1
         Left            =   3060
         TabIndex        =   2
         Top             =   315
         Width           =   1590
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Funcionário da Prefeitura"
         Height          =   195
         Index           =   0
         Left            =   585
         TabIndex        =   1
         Top             =   315
         Value           =   -1  'True
         Width           =   2310
      End
   End
End
Attribute VB_Name = "frmUsoPlataforma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bExec


Private Sub cmbUser_Click()
Dim Sql As String, RdoAux As rdoResultset, bFunc As Boolean, nCodigo As Long, x As Integer, lista() As Long, y As Integer

If cmbUser.ListIndex = -1 Then Exit Sub
ReDim lista(0)

bFunc = optTipo(0).value
nCodigo = cmbUser.ItemData(cmbUser.ListIndex)

For x = 0 To lstEmpresa.ListCount - 1
    lstEmpresa.Selected(x) = False
Next

Sql = "select empresa from rodo_uso_plataforma_user where user_id=" & nCodigo & " and funcionario=" & IIf(bFunc, 1, 0)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve lista(UBound(lista) + 1)
        lista(UBound(lista)) = !Empresa
       .MoveNext
    Loop
   .Close
End With

For x = 1 To UBound(lista)
    nCodigo = lista(x)
    For y = 0 To lstEmpresa.ListCount - 1
        If lstEmpresa.ItemData(y) = nCodigo Then
            lstEmpresa.Selected(y) = True
            lstEmpresa.ListIndex = y
            Exit For
        End If
    Next
Next

End Sub

Private Sub cmdGravar_Click()
Dim Sql As String, bFunc As Boolean, nCodigo As Integer, x As Integer

bFunc = optTipo(0).value
nCodigo = cmbUser.ItemData(cmbUser.ListIndex)

Sql = "delete from rodo_uso_plataforma_user where user_id=" & nCodigo & " and funcionario=" & IIf(bFunc, 1, 0)
cn.Execute Sql, rdExecDirect

For x = 0 To lstEmpresa.ListCount - 1
    If lstEmpresa.Selected(x) Then
        Sql = "insert rodo_uso_plataforma_user(user_id,funcionario,empresa) values(" & nCodigo & "," & IIf(bFunc, 1, 0) & "," & lstEmpresa.ItemData(x) & ")"
        cn.Execute Sql, rdExecDirect
    End If
Next

MsgBox "Dados gravados!", vbInformation, "Informação"

End Sub

Private Sub Form_Load()
Dim Sql As String, RdoAux As rdoResultset
Centraliza Me

Sql = "select codcidadao,nomecidadao from cidadao where codcidadao in (SELECT codigo FROM rodo_empresa) order by nomecidadao"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        lstEmpresa.AddItem !nomecidadao
        lstEmpresa.ItemData(lstEmpresa.NewIndex) = !CodCidadao
       .MoveNext
    Loop
   .Close
End With

Carrega_Usuario
'cmbUser.SetFocus
End Sub

Private Sub Carrega_Usuario()
Dim Sql As String, RdoAux As rdoResultset, bFunc As Boolean
bFunc = optTipo(0).value

bExec = False
cmbUser.Clear
If bFunc Then
    Sql = "select id as codigo,nomecompleto as nome from usuario where id>1 and ativo=1 order by nomecompleto"
Else
    Sql = "select id as codigo,nome from usuario_web where ativo=1 order by nome"
End If
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbUser.AddItem !Nome
        cmbUser.ItemData(cmbUser.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With
bExec = True
cmbUser.ListIndex = 0

End Sub


Private Sub optTipo_Click(Index As Integer)
Carrega_Usuario
End Sub
