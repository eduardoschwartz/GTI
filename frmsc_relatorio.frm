VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmsc_relatorio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatórios diversos"
   ClientHeight    =   1260
   ClientLeft      =   14805
   ClientTop       =   6330
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1260
   ScaleWidth      =   7110
   Begin VB.ComboBox cmbAno 
      Height          =   315
      Left            =   3375
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   630
      Width           =   1095
   End
   Begin VB.ComboBox cmbMes 
      Height          =   315
      ItemData        =   "frmsc_relatorio.frx":0000
      Left            =   1530
      List            =   "frmsc_relatorio.frx":002B
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   630
      Width           =   1725
   End
   Begin VB.ComboBox cmbTipo 
      Height          =   315
      Left            =   1530
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   135
      Width           =   5370
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   360
      Left            =   5805
      TabIndex        =   5
      ToolTipText     =   "Imprimir o relatório"
      Top             =   585
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "Imprimir"
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
      MICON           =   "frmsc_relatorio.frx":0094
      PICN            =   "frmsc_relatorio.frx":00B0
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
      Caption         =   "Período Consumo:"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   2
      Top             =   675
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo de Relatório.:"
      Height          =   240
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Width           =   1410
   End
End
Attribute VB_Name = "frmsc_relatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrint_Click()
Dim ntipo As Integer

ntipo = cmbTipo.ItemData(cmbTipo.ListIndex)
Select Case ntipo
    Case 1
        frmReport.ShowReport3 "SC_AGUAMENSAL", frmMdi.HWND, Me.HWND
    Case 2
        frmReport.ShowReport3 "SC_ENERGIAMENSAL", frmMdi.HWND, Me.HWND
    Case 3
        frmReport.ShowReport3 "SC_TELEFONEFIXOMENSAL", frmMdi.HWND, Me.HWND
    Case 4
    Case 5
    Case 6
End Select

End Sub

Private Sub Form_Load()
Dim x As Integer
Centraliza Me

For x = 0 To 11
    cmbMes.ItemData(x) = x + 1
Next
cmbMes.ListIndex = Month(Now) - 1

For x = 2021 To Year(Now)
    cmbAno.AddItem x
Next
cmbAno.ListIndex = cmbAno.ListCount - 1

cmbTipo.AddItem "01-Consumo mensal de água"
cmbTipo.ItemData(cmbTipo.NewIndex) = 1
cmbTipo.AddItem "02-Consumo mensal de energia"
cmbTipo.ItemData(cmbTipo.NewIndex) = 2
cmbTipo.AddItem "03-Consumo mensal de telefonia fixa"
cmbTipo.ItemData(cmbTipo.NewIndex) = 3
cmbTipo.AddItem "04-Consumo mensal de telefonia celular"
cmbTipo.ItemData(cmbTipo.NewIndex) = 4
cmbTipo.AddItem "05-Consumo mensal de internet"
cmbTipo.ItemData(cmbTipo.NewIndex) = 5
cmbTipo.AddItem "06-Consumo mensal de taxas de correio"
cmbTipo.ItemData(cmbTipo.NewIndex) = 6

cmbTipo.ListIndex = 0

End Sub
