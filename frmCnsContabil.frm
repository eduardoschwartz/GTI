VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmCnsContabil 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Escritórios Contábeis"
   ClientHeight    =   4530
   ClientLeft      =   16755
   ClientTop       =   8415
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   Begin prjChameleon.chameleonButton cmdConsultar 
      Height          =   360
      Left            =   5985
      TabIndex        =   2
      ToolTipText     =   "Consulta Cidadãos Cadastrados"
      Top             =   45
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "frmCnsContabil.frx":0000
      PICN            =   "frmCnsContabil.frx":001C
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
      Height          =   3540
      Left            =   90
      TabIndex        =   1
      Top             =   450
      Width           =   6360
   End
   Begin VB.TextBox txtBusca 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   5775
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   5400
      TabIndex        =   3
      ToolTipText     =   "Sair da Tela"
      Top             =   4095
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
      MICON           =   "frmCnsContabil.frx":0176
      PICN            =   "frmCnsContabil.frx":0192
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
Attribute VB_Name = "frmCnsContabil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Codigo As Integer

Private Sub cmdConsultar_Click()
Dim Sql As String, RdoAux As rdoResultset

lstMain.Clear
Sql = "SELECT CODIGOESC,NOMEESC FROM ESCRITORIOCONTABIL WHERE NOMEESC LIKE '%" & txtBusca.Text & "%' ORDER BY NOMEESC"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        lstMain.AddItem !NOMEESC
        lstMain.ItemData(lstMain.NewIndex) = !CODIGOESC
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub cmdSair_Click()
Codigo = 0

If lstMain.ListCount > 0 Then
    If lstMain.ListIndex > -1 Then
        Codigo = lstMain.ItemData(lstMain.ListIndex)
    Else
        Codigo = lstMain.ItemData(0)
    End If
End If
frmEscContab.CodigoEscritorio = Codigo
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
End Sub

Private Sub txtBusca_KeyPress(KeyAscii As Integer)
Tweak txtBusca, KeyAscii, AllLettersAllCaps
End Sub
