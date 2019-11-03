VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmHorarioFuncionamento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Horário de funcionanamento por atividade"
   ClientHeight    =   4125
   ClientLeft      =   11910
   ClientTop       =   7230
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4125
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstDestino 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   5940
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   1230
      Width           =   5175
   End
   Begin VB.TextBox txtHorario 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   540
      Width           =   10845
   End
   Begin VB.ListBox lstOrigem 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   210
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1230
      Width           =   5175
   End
   Begin VB.ComboBox cmbHorario 
      Height          =   315
      Left            =   930
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   7305
   End
   Begin prjChameleon.chameleonButton cmdDel 
      Height          =   285
      Left            =   5460
      TabIndex        =   3
      ToolTipText     =   "Remove centro de custos"
      Top             =   2190
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmHorarioFuncionamento.frx":0000
      PICN            =   "frmHorarioFuncionamento.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdAdd 
      Height          =   285
      Left            =   5460
      TabIndex        =   4
      ToolTipText     =   "Adiciona centro de custos"
      Top             =   1860
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmHorarioFuncionamento.frx":0176
      PICN            =   "frmHorarioFuncionamento.frx":0192
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Atividades selecionadas"
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   1
      Left            =   6090
      TabIndex        =   7
      Top             =   990
      Width           =   1845
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Atividades disponíveis"
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1845
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Horário..:"
      Height          =   225
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   825
   End
End
Attribute VB_Name = "frmHorarioFuncionamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbHorario_Click()
Dim Sql As String, RdoAux As rdoResultset, nCodigo As Double, nHorario As Integer
txtHorario.Text = cmbHorario.Text
nHorario = cmbHorario.ItemData(cmbHorario.ListIndex)

lstOrigem.Clear: lstDestino.Clear

Sql = "select * from atividade where horario is null"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        lstOrigem.AddItem !codatividade & " - " & !descatividade
        lstOrigem.ItemData(lstOrigem.NewIndex) = !codatividade
       .MoveNext
    Loop
   .Close
End With
On Error Resume Next
lstOrigem.ListIndex = 0

Sql = "select * from atividade where horario=" & nHorario
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        lstDestino.AddItem !codatividade & " - " & !descatividade
        lstDestino.ItemData(lstDestino.NewIndex) = !codatividade
       .MoveNext
    Loop
   .Close
End With
If lstDestino.ListCount > 0 Then lstDestino.ListIndex = 0


End Sub

Private Sub cmdAdd_Click()
Dim Sql As String, RdoAux As rdoResultset, nCodigo As Double, nHorario As Integer, nPos As Integer

If lstOrigem.ListIndex = -1 Then
    MsgBox "Selecione uma atividade.", vbCritical, "Atenção"
    Exit Sub
End If
nPos = lstOrigem.ListIndex
nHorario = cmbHorario.ItemData(cmbHorario.ListIndex)
nCodigo = lstOrigem.ItemData(lstOrigem.ListIndex)
lstDestino.AddItem lstOrigem.Text
lstDestino.ItemData(lstDestino.NewIndex) = nCodigo
lstOrigem.RemoveItem (lstOrigem.ListIndex)

Sql = "update atividade set horario=" & nHorario & " where codatividade=" & nCodigo
cn.Execute Sql, rdExecDirect
On Error Resume Next
lstOrigem.ListIndex = nPos
lstOrigem.SetFocus
End Sub

Private Sub cmdDel_Click()
Dim Sql As String, RdoAux As rdoResultset, nCodigo As Double, nHorario As Integer

If lstDestino.ListIndex = -1 Then
    MsgBox "Selecione uma atividade.", vbCritical, "Atenção"
    Exit Sub
End If

nHorario = cmbHorario.ItemData(cmbHorario.ListIndex)
nCodigo = lstDestino.ItemData(lstDestino.ListIndex)
lstOrigem.AddItem lstDestino.Text
lstOrigem.ItemData(lstOrigem.NewIndex) = nCodigo
lstDestino.RemoveItem (lstDestino.ListIndex)

Sql = "update atividade set horario=null where codatividade=" & nCodigo
cn.Execute Sql, rdExecDirect

End Sub

Private Sub Form_Load()
Dim Sql As String, RdoAux As rdoResultset
Centraliza Me

Sql = "select id,descricao from horario_funcionamento order by descricao"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbHorario.AddItem !descricao
        cmbHorario.ItemData(cmbHorario.NewIndex) = !id
       .MoveNext
    Loop
   .Close
End With
cmbHorario.ListIndex = 0

End Sub
