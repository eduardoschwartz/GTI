VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmIsencaoVSAtividade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Isenção de VS por Atividade"
   ClientHeight    =   4800
   ClientLeft      =   3855
   ClientTop       =   3345
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4800
   ScaleWidth      =   6720
   Begin VB.ListBox lstMain 
      Appearance      =   0  'Flat
      Height          =   3630
      Left            =   135
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   540
      Width           =   6465
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   5445
      TabIndex        =   1
      ToolTipText     =   "Sair da Tela"
      Top             =   4350
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmIsencaoVSAtividade.frx":0000
      PICN            =   "frmIsencaoVSAtividade.frx":001C
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
      Left            =   4365
      TabIndex        =   2
      ToolTipText     =   "Gravar os Dados"
      Top             =   4350
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
      MICON           =   "frmIsencaoVSAtividade.frx":008A
      PICN            =   "frmIsencaoVSAtividade.frx":00A6
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
      Caption         =   "Marque as atividades que serão isentas da Vigilância Sanitária"
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   180
      TabIndex        =   3
      Top             =   135
      Width           =   6315
   End
End
Attribute VB_Name = "frmIsencaoVSAtividade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGravar_Click()
Dim x As Integer

Sql = "delete from isencaovs_atividade"
cn.Execute Sql, rdExecDirect

For x = 0 To lstMain.ListCount - 1
    If lstMain.Selected(x) = True Then
        lstMain.ListIndex = x
        Sql = "insert isencaovs_atividade (codigo) values(" & lstMain.ItemData(x) & ")"
        cn.Execute Sql, rdExecDirect
    End If
Next

MsgBox "Isenção gravada com sucesso.", vbInformation, "Atenção"

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
Carrega_Lista
End Sub

Private Sub Carrega_Lista()
Dim Sql As String, RdoAux As rdoResultset, aCodigo() As Integer, x As Integer

ReDim aCodigo(0)

Sql = "select codigo from isencaovs_atividade"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aCodigo(UBound(aCodigo) + 1)
        aCodigo(UBound(aCodigo)) = !Codigo
       .MoveNext
    Loop
   .Close
End With


lstMain.Clear

Sql = "SELECT codatividade,descatividade FROM atividade ORDER BY descatividade"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        lstMain.AddItem !descatividade
        lstMain.ItemData(lstMain.NewIndex) = !codatividade
        For x = 1 To UBound(aCodigo)
            If aCodigo(x) = !codatividade Then
                lstMain.Selected(lstMain.ListCount - 1) = True
                Exit For
            End If
        Next
       .MoveNext
    Loop
   .Close
End With



End Sub

