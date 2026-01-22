VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmSenhaPre 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pré-Atendimento"
   ClientHeight    =   6270
   ClientLeft      =   5790
   ClientTop       =   2250
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6270
   ScaleWidth      =   3255
   Begin VB.TextBox txtSenha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   6
      Left            =   1845
      TabIndex        =   32
      Text            =   "0"
      Top             =   8775
      Width           =   690
   End
   Begin VB.TextBox txtSenha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   7
      Left            =   1845
      TabIndex        =   31
      Text            =   "0"
      Top             =   9090
      Width           =   690
   End
   Begin VB.TextBox txtSenha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   8
      Left            =   1845
      TabIndex        =   30
      Text            =   "0"
      Top             =   9405
      Width           =   690
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   315
      Left            =   135
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   10575
      Width           =   2985
   End
   Begin VB.CheckBox chkParalela 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00404040&
      Caption         =   "Impressora Paralela"
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   405
      TabIndex        =   24
      Top             =   9945
      Width           =   1725
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   540
      Picture         =   "frmSenhaPre.frx":0000
      ScaleHeight     =   450
      ScaleWidth      =   1530
      TabIndex        =   23
      Top             =   11070
      Width           =   1590
   End
   Begin VB.TextBox txtSenha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   1845
      TabIndex        =   21
      Text            =   "0"
      Top             =   8460
      Width           =   690
   End
   Begin VB.TextBox txtSenha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   1845
      TabIndex        =   18
      Text            =   "0"
      Top             =   8145
      Width           =   690
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2655
      Top             =   5715
   End
   Begin VB.TextBox txtSenha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   1845
      TabIndex        =   13
      Text            =   "0"
      Top             =   7830
      Width           =   690
   End
   Begin VB.TextBox txtSenha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   1845
      TabIndex        =   11
      Text            =   "0"
      Top             =   7515
      Width           =   690
   End
   Begin VB.TextBox txtSenha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   1845
      TabIndex        =   9
      Text            =   "0"
      Top             =   7200
      Width           =   690
   End
   Begin VB.TextBox txtSenha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   1845
      TabIndex        =   6
      Text            =   "0"
      Top             =   6885
      Width           =   690
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4365
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1260
      Width           =   645
   End
   Begin prjChameleon.chameleonButton cmdSenha 
      Height          =   510
      Index           =   0
      Left            =   90
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   180
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "PREFEITURA"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   8454143
      FCOLO           =   8454143
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmSenhaPre.frx":3E02
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdSenha 
      Height          =   510
      Index           =   1
      Left            =   90
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   765
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "PREFERÊNCIAL"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   8454143
      FCOLO           =   8454143
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmSenhaPre.frx":411C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdSenha 
      Height          =   510
      Index           =   2
      Left            =   90
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1350
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "PAT"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   8454143
      FCOLO           =   8454143
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmSenhaPre.frx":4436
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdSenha 
      Height          =   510
      Index           =   3
      Left            =   90
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1935
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "PAT - PREF"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   8454143
      FCOLO           =   8454143
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmSenhaPre.frx":4750
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdConfig 
      Height          =   330
      Left            =   855
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5715
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   582
      BTYPE           =   14
      TX              =   "Configuração"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   12640511
      FCOLO           =   12640511
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmSenhaPre.frx":4A6A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   -1  'True
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdSenha 
      Height          =   510
      Index           =   4
      Left            =   90
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2520
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "PAV-RECEITA F"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   8454143
      FCOLO           =   8454143
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmSenhaPre.frx":4D84
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdSenha 
      Height          =   510
      Index           =   5
      Left            =   90
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3105
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "PAV - PREF"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   8454143
      FCOLO           =   8454143
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmSenhaPre.frx":509E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdSenha 
      Height          =   510
      Index           =   6
      Left            =   90
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   3690
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "REFIS"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   8454143
      FCOLO           =   8454143
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmSenhaPre.frx":53B8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdSenha 
      Height          =   510
      Index           =   7
      Left            =   90
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4275
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "REFIS - PREF"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   8454143
      FCOLO           =   8454143
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmSenhaPre.frx":56D2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdSenha 
      Height          =   510
      Index           =   8
      Left            =   90
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   4860
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "BOLETOS - DAM"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   8454143
      FCOLO           =   8454143
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmSenhaPre.frx":59EC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "REFIS.........:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   9
      Left            =   495
      TabIndex        =   35
      Top             =   8775
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "REFIS-PREF:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   8
      Left            =   495
      TabIndex        =   34
      Top             =   9090
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "BOLETO-DAM..:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   7
      Left            =   495
      TabIndex        =   33
      Top             =   9405
      Width           =   1230
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Impressoras:"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Left            =   135
      TabIndex        =   25
      Top             =   10305
      Width           =   1770
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PAV-PREF...:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   6
      Left            =   495
      TabIndex        =   22
      Top             =   8460
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PAV-RECEI..:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   5
      Left            =   495
      TabIndex        =   19
      Top             =   8145
      Width           =   1230
   End
   Begin VB.Label lblBanda 
      Height          =   240
      Left            =   4230
      TabIndex        =   16
      Top             =   765
      Width           =   1410
   End
   Begin VB.Label lblSenha 
      Height          =   240
      Left            =   4230
      TabIndex        =   15
      Top             =   405
      Width           =   1410
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PAT - PREF.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   4
      Left            =   495
      TabIndex        =   14
      Top             =   7830
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PAT............:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   3
      Left            =   495
      TabIndex        =   12
      Top             =   7515
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Preferencial.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   2
      Left            =   495
      TabIndex        =   10
      Top             =   7200
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Prefeitura....:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   1
      Left            =   495
      TabIndex        =   8
      Top             =   6885
      Width           =   1230
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Próxima Senha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Index           =   0
      Left            =   810
      TabIndex        =   7
      Top             =   6525
      Width           =   1590
   End
End
Attribute VB_Name = "frmSenhaPre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub chkParalela_Click()
If chkParalela.value = vbChecked Then
    cmbPrinter.Enabled = False
Else
    cmbPrinter.Enabled = True
End If
End Sub

Private Sub cmdConfig_Click()
If cmdConfig.value = False Then
    Me.Height = 6735
Else
    Me.Height = 11505
End If
SaveAll
End Sub

Private Sub cmdSenha_Click(Index As Integer)
Dim RdoAux As rdoResultset, sql As String, nSenha As Integer, nBanda As Integer, sParalela As String
Me.Enabled = False

Ocupado
nBanda = Index + 1

If nBanda = 1 Then
    nSenha = Val(txtSenha(0).Text)
ElseIf nBanda = 2 Then
    nSenha = Val(txtSenha(1).Text)
ElseIf nBanda = 3 Then
    nSenha = Val(txtSenha(2).Text)
ElseIf nBanda = 4 Then
    nSenha = Val(txtSenha(3).Text)
ElseIf nBanda = 5 Then
    nSenha = Val(txtSenha(4).Text)
ElseIf nBanda = 6 Then
    nSenha = Val(txtSenha(5).Text)
ElseIf nBanda = 7 Then
    nSenha = Val(txtSenha(6).Text)
ElseIf nBanda = 8 Then
    nSenha = Val(txtSenha(7).Text)
ElseIf nBanda = 9 Then
    nSenha = Val(txtSenha(8).Text)
End If

sql = "SELECT * FROM SSPAC WHERE DATAENTRADA='" & Format(Now, "mm/dd/yyyy") & "' AND SENHA=" & nSenha
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount = 0 Then
    sql = "INSERT SSPAC(DATAENTRADA,HORAENTRADA,SENHA,BANDA,MONITOR) VALUES('" & Format(Now, "mm/dd/yyyy") & "','"
    sql = sql & Format(Now, "hh:mm:ss") & "'," & nSenha & "," & nBanda & ",0)"
    cn.Execute sql, rdExecDirect
    RdoAux.Close
End If

lblSenha.Caption = Format(nSenha, "000")
lblBanda.Caption = Left(cmdSenha(Index).Caption, InStr(1, cmdSenha(Index).Caption, "(") - 2)

txtSenha(Index).Text = Val(txtSenha(Index).Text) + 1
On Error Resume Next
Text1.SetFocus

sParalela = GetSetting("GTI", "PRINTER", "PARALELA")
If sParalela <> "S" Then
    printTermica nSenha, lblBanda.Caption
    Me.Enabled = True
    Liberado
    Exit Sub
End If

On Error GoTo Erro
Open "Lpt1" For Output As #1
Print #1, Spc(2); "========================================="
Print #1, Spc(6); "PREFEITURA MUNICIPAL DE JABOTICABAL"
Print #1, Spc(2); "Sistema Pratico de Atendimento ao Cidadao"
Print #1, Spc(2); "========================================="
Print #1, Spc(8); "Data:" & Format(Now, "dd/mm/yyyy"); Spc(2); "Hora:" & Format(Now, "hh:mm:ss")
Print #1, Spc(12); Chr(27) & Chr(69) & Chr(27) & Chr(14) + "Senha:" & Format(nSenha, "000")
Print #1, Spc(12); Chr(27) & Chr(70) & Chr(27) & Chr(14) + lblBanda.Caption
Print #1, Chr(20)
Print #1, Spc(2); "POR FAVOR AGUARDE."
Print #1, Chr(10) & Chr(13)
Print #1, Chr(10)
Print #1, Chr(10)
Print #1, Chr(10)
Close #1
Liberado

Me.Enabled = True
Exit Sub
Erro:
Liberado
Me.Enabled = True
MsgBox "Impressora não conectada.", vbCritical, "Atenção"

End Sub

Private Sub cmdSenha_MouseOut(Index As Integer)
On Error Resume Next
Text1.SetFocus
End Sub

Private Sub Form_Load()
Dim sql As String, RdoAux As rdoResultset, nSenha As Integer, sParalela As String
Dim p As Printer, x As Integer

For Each p In Printers
  cmbPrinter.AddItem p.DeviceName
Next

sPrinter = GetSetting("GTI", "PRINTER", "SELECTED")

If sPrinter = "" Then
    cmbPrinter.ListIndex = 0
Else
    For x = 0 To cmbPrinter.ListCount - 1
        If cmbPrinter.List(x) = sPrinter Then
            cmbPrinter.ListIndex = x
            Exit For
        End If
    Next
End If
SaveSetting "GTI", "PRINTER", "SELECTED", sPrinter
Ocupado

sql = "SELECT MAX(SENHA) AS MAXIMO FROM SSPAC WHERE YEAR(DATAENTRADA)=" & Year(Now) & " AND MONTH(DATAENTRADA)=" & Month(Now) & " AND "
sql = sql & "DAY(DATAENTRADA)=" & Day(Now) & " AND BANDA=" & 1
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    If IsNull(!maximo) Then
        nSenha = 1
    Else
        nSenha = !maximo + 1
    End If
   .Close
End With
txtSenha(0).Text = nSenha

sql = "SELECT MAX(SENHA) AS MAXIMO FROM SSPAC WHERE YEAR(DATAENTRADA)=" & Year(Now) & " AND MONTH(DATAENTRADA)=" & Month(Now) & " AND "
sql = sql & "DAY(DATAENTRADA)=" & Day(Now) & " AND BANDA=" & 2
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    If IsNull(!maximo) Then
        nSenha = 200
    Else
        nSenha = !maximo + 1
    End If
   .Close
End With
txtSenha(1).Text = nSenha

sql = "SELECT MAX(SENHA) AS MAXIMO FROM SSPAC WHERE YEAR(DATAENTRADA)=" & Year(Now) & " AND MONTH(DATAENTRADA)=" & Month(Now) & " AND "
sql = sql & "DAY(DATAENTRADA)=" & Day(Now) & " AND BANDA=" & 3
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    If IsNull(!maximo) Then
        nSenha = 300
    Else
        nSenha = !maximo + 1
    End If
   .Close
End With
txtSenha(2).Text = nSenha

sql = "SELECT MAX(SENHA) AS MAXIMO FROM SSPAC WHERE YEAR(DATAENTRADA)=" & Year(Now) & " AND MONTH(DATAENTRADA)=" & Month(Now) & " AND "
sql = sql & "DAY(DATAENTRADA)=" & Day(Now) & " AND BANDA=" & 4
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    If IsNull(!maximo) Then
        nSenha = 400
    Else
        nSenha = !maximo + 1
    End If
   .Close
End With
txtSenha(3).Text = nSenha

sql = "SELECT MAX(SENHA) AS MAXIMO FROM SSPAC WHERE YEAR(DATAENTRADA)=" & Year(Now) & " AND MONTH(DATAENTRADA)=" & Month(Now) & " AND "
sql = sql & "DAY(DATAENTRADA)=" & Day(Now) & " AND BANDA=" & 5
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    If IsNull(!maximo) Then
        nSenha = 500
    Else
        nSenha = !maximo + 1
    End If
   .Close
End With
txtSenha(4).Text = nSenha

sql = "SELECT MAX(SENHA) AS MAXIMO FROM SSPAC WHERE YEAR(DATAENTRADA)=" & Year(Now) & " AND MONTH(DATAENTRADA)=" & Month(Now) & " AND "
sql = sql & "DAY(DATAENTRADA)=" & Day(Now) & " AND BANDA=" & 6
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    If IsNull(!maximo) Then
        nSenha = 600
    Else
        nSenha = !maximo + 1
    End If
   .Close
End With
txtSenha(5).Text = nSenha

sql = "SELECT MAX(SENHA) AS MAXIMO FROM SSPAC WHERE YEAR(DATAENTRADA)=" & Year(Now) & " AND MONTH(DATAENTRADA)=" & Month(Now) & " AND "
sql = sql & "DAY(DATAENTRADA)=" & Day(Now) & " AND BANDA=" & 7
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    If IsNull(!maximo) Then
        nSenha = 700
    Else
        nSenha = !maximo + 1
    End If
   .Close
End With
'If nSenha < 1000 Then nSenha = 1104
txtSenha(6).Text = nSenha

sql = "SELECT MAX(SENHA) AS MAXIMO FROM SSPAC WHERE YEAR(DATAENTRADA)=" & Year(Now) & " AND MONTH(DATAENTRADA)=" & Month(Now) & " AND "
sql = sql & "DAY(DATAENTRADA)=" & Day(Now) & " AND BANDA=" & 8
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    If IsNull(!maximo) Then
        nSenha = 800
    Else
        nSenha = !maximo + 1
    End If
   .Close
End With
'If nSenha < 1000 Then nSenha = 2201
txtSenha(7).Text = nSenha

sql = "SELECT MAX(SENHA) AS MAXIMO FROM SSPAC WHERE YEAR(DATAENTRADA)=" & Year(Now) & " AND MONTH(DATAENTRADA)=" & Month(Now) & " AND "
sql = sql & "DAY(DATAENTRADA)=" & Day(Now) & " AND BANDA=" & 9
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    If IsNull(!maximo) Then
        nSenha = 900
    Else
        nSenha = !maximo + 1
    End If
   .Close
End With
'If nSenha < 1000 Then nSenha = 3102
txtSenha(8).Text = nSenha


sParalela = GetSetting("GTI", "PRINTER", "PARALELA")
If sParalela = "" Then
    chkParalela.value = vbUnchecked
    SaveSetting "GTI", "PRINTER", "PARALELA", "N"
Else
    If sParalela = "S" Then
        chkParalela.value = vbChecked
    Else
        chkParalela.value = vbUnchecked
    End If
End If

Le
Liberado
End Sub

Private Sub Le()
On Error GoTo Erro
Dim sql As String, RdoAux As rdoResultset, dData As Date, x As Integer
Dim aCount(8) As Integer
dData = Now
cmdSenha(0).Caption = "PREFEITURA (0)"
cmdSenha(1).Caption = "PREFERÊNCIAL (0)"
cmdSenha(2).Caption = "PAT (0)"
cmdSenha(3).Caption = "PAT - PREF F(0)"
cmdSenha(4).Caption = "PAV-RECEITA F(0)"
cmdSenha(5).Caption = "PAV - PREF (0)"
cmdSenha(6).Caption = "REFIS(0)"
cmdSenha(7).Caption = "REFIS - PREF(0)"
cmdSenha(8).Caption = "BOLETOS - DAM(0)"

sql = "SELECT * FROM SSPAC WHERE YEAR(DATAENTRADA)=" & Year(dData) & " AND MONTH(DATAENTRADA)=" & Month(dData) & " AND "
sql = sql & "DAY(DATAENTRADA)=" & Day(dData) & " AND DATACHAMADA IS NULL"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    
    Do Until .EOF
        aCount(!BANDA - 1) = (aCount(!BANDA - 1)) + 1
       Select Case !BANDA - 1
            Case 0
                cmdSenha(!BANDA - 1).Caption = "PREFEITURA (" & aCount(!BANDA - 1) & ")"
            Case 1
                cmdSenha(!BANDA - 1).Caption = "PREFERENCIAL (" & aCount(!BANDA - 1) & ")"
            Case 2
                cmdSenha(!BANDA - 1).Caption = "PAT (" & aCount(!BANDA - 1) & ")"
            Case 3
                cmdSenha(!BANDA - 1).Caption = "PAT - PREF (" & aCount(!BANDA - 1) & ")"
            Case 4
                cmdSenha(!BANDA - 1).Caption = "PAV-RECEITA F(" & aCount(!BANDA - 1) & ")"
            Case 5
                cmdSenha(!BANDA - 1).Caption = "PAV - PREF (" & aCount(!BANDA - 1) & ")"
            Case 6
                cmdSenha(!BANDA - 1).Caption = "REFIS (" & aCount(!BANDA - 1) & ")"
            Case 7
                cmdSenha(!BANDA - 1).Caption = "REFIS - PREF(" & aCount(!BANDA - 1) & ")"
            Case 8
                cmdSenha(!BANDA - 1).Caption = "BOLETOS - DAM (" & aCount(!BANDA - 1) & ")"
        End Select
        .MoveNext
    Loop
   .Close
End With

Exit Sub
Erro:
MsgBox Err.Description
'Liberado
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
SaveAll
End Sub

Private Sub SaveAll()
If chkParalela.value = vbChecked Then
    SaveSetting "GTI", "PRINTER", "PARALELA", "S"
Else
    SaveSetting "GTI", "PRINTER", "PARALELA", "N"
End If
SaveSetting "GTI", "PRINTER", "SELECTED", cmbPrinter.List(cmbPrinter.ListIndex)

End Sub

Private Sub Timer1_Timer()
Le
End Sub

Private Sub printTermica(nSenha As Integer, sBanda As String)
Dim sPrinter As String

sPrinter = GetSetting("GTI", "PRINTER", "SELECTED")
If sPrinter = "" Then
    sPrinter = cmbPrinter.List(0)
End If
On Error GoTo Erro

Dim p As Printer
defprinter = Printer.DeviceName
For Each p In Printers
  If p.DeviceName = sPrinter Then
    Set Printer = p
    Exit For
  End If
Next

With Printer
    .FontBold = True
    .FontSize = 7
    .CurrentX = 0
    .CurrentY = 10
     Printer.Print "PREFEITURA MUNCIPAL"
    .CurrentX = 4
    .CurrentY = 18
     Printer.Print "DE JABOTICABAL"
    .PaintPicture Picture1.Picture, 90, 0, 90, 50
    .ScaleMode = 2 'Point
    .FontName = "Arial"
    .FontBold = False
    .FontSize = 10
     Printer.Print " "
    .CurrentX = 18
    .CurrentY = 60
     Printer.Print "Data:" & Format(Now, "dd/mm/yyyy"); Spc(2); "Hora:" & Format(Now, "hh:mm:ss")
    
    .FontBold = True
    .FontSize = 20
    .CurrentX = 35
    .CurrentY = 80
     Printer.Print "Senha: " & Format(nSenha, "000")
    .CurrentX = 35
    .CurrentY = 100
    .FontSize = 14
     Printer.Print lblBanda.Caption
    .FontBold = False
    .FontSize = 8
    .CurrentX = 40
    .CurrentY = 130
     Printer.Print "POR FAVOR AGUARDE."
    .EndDoc
End With



Exit Sub
Erro:
MsgBox Err.Description

End Sub

