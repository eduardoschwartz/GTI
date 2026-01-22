VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmIsencaoVSCodigo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Isenção de VS por inscrição"
   ClientHeight    =   4515
   ClientLeft      =   12120
   ClientTop       =   5475
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4515
   ScaleWidth      =   6900
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   135
      MaxLength       =   6
      TabIndex        =   0
      Top             =   675
      Width           =   1500
   End
   Begin VB.ListBox lstMain 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   135
      TabIndex        =   1
      Top             =   1170
      Width           =   6630
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   360
      Left            =   5625
      TabIndex        =   3
      ToolTipText     =   "Sair da Tela"
      Top             =   4050
      Width           =   1035
      _ExtentX        =   1826
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
      MICON           =   "frmIsencaoVSCodigo.frx":0000
      PICN            =   "frmIsencaoVSCodigo.frx":001C
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
      Left            =   1755
      TabIndex        =   4
      ToolTipText     =   "Gravar os Dados"
      Top             =   675
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "&Adicionar"
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
      MICON           =   "frmIsencaoVSCodigo.frx":008A
      PICN            =   "frmIsencaoVSCodigo.frx":00A6
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
      Left            =   4410
      TabIndex        =   5
      ToolTipText     =   "Excluir Registro"
      Top             =   4050
      Width           =   1035
      _ExtentX        =   1826
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
      MICON           =   "frmIsencaoVSCodigo.frx":044B
      PICN            =   "frmIsencaoVSCodigo.frx":0467
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
      Caption         =   "Informe as inscrições isentas da Vigilância Sanitária"
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   135
      TabIndex        =   2
      Top             =   225
      Width           =   6315
   End
End
Attribute VB_Name = "frmIsencaoVSCodigo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
CarregaLista
End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)
Tweak txtCod, KeyAscii, IntegerPositive
End Sub

Private Sub cmdGravar_Click()
Dim sql As String, RdoAux As rdoResultset, nCodigo As Long, bFind As Boolean

If Not IsNumeric(txtCod.Text) Then
    MsgBox "Informe a inscrição cadastral", vbCritical, "Erro"
    Exit Sub
End If

nCodigo = Val(txtCod.Text)
If nCodigo < 100000 Or nCodigo > 200000 Then
    MsgBox "Inscrição cadastral inválida", vbCritical, "Erro"
    Exit Sub
End If

bFind = False
sql = "select codigo from isencaovs_codigo where codigo=" & nCodigo
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    bFind = True
End If
RdoAux.Close

If bFind Then
    MsgBox "Esta Inscrição já esta na lista de isenção", vbCritical, "Erro"
    Exit Sub
End If

sql = "select razaosocial from mobiliario where codigomob=" & nCodigo
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If RdoAux.RowCount = 0 Then
        MsgBox "Inscrição não cadastrada", vbCritical, "Erro"
        Exit Sub
    Else
        If MsgBox("Incluir a inscrição " & txtCod.Text & " - " & !RazaoSocial & "?", vbQuestion + vbYesNo) = vbYes Then
            lstMain.AddItem txtCod.Text & " - " & !RazaoSocial
            lstMain.ItemData(lstMain.NewIndex) = nCodigo
            
            sql = "insert isencaovs_codigo (codigo) values(" & nCodigo & ")"
            cn.Execute sql, rdExecDirect
        End If
    End If
   .Close
End With
txtCod.Text = ""
End Sub

Private Sub CarregaLista()
Dim sql As String, RdoAux As rdoResultset

sql = "select codigo,razaosocial from isencaovs_codigo i inner join mobiliario m  on i.codigo=m.codigomob order by codigo"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        lstMain.AddItem !Codigo & " - " & !RazaoSocial
        lstMain.ItemData(lstMain.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub cmdExcluir_Click()
Dim sql As String, RdoAux As rdoResultset, nCodigo As Long

If lstMain.ListIndex = -1 Then
    MsgBox "Selecione a inscrição à excluir", vbCritical, "Erro"
    Exit Sub
End If

If MsgBox("Excluir a inscrição " & lstMain.List(lstMain.ListIndex) & "?", vbQuestion + vbYesNo) = vbYes Then
    nCodigo = lstMain.ItemData(lstMain.ListIndex)
    sql = "delete from isencaovs_codigo where codigo=" & nCodigo
    cn.Execute sql, rdExecDirect
    lstMain.RemoveItem (lstMain.ListIndex)
End If

End Sub

