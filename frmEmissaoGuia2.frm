VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmEmissaoGuia2 
   Appearance      =   0  'Flat
   BackColor       =   &H00FBFBE3&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Composição da guia"
   ClientHeight    =   5460
   ClientLeft      =   10725
   ClientTop       =   3690
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   330
      Left            =   7860
      TabIndex        =   38
      ToolTipText     =   "Imprimir as parcelas"
      Top             =   4530
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
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
      MICON           =   "frmEmissaoGuia2.frx":0000
      PICN            =   "frmEmissaoGuia2.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Tributacao.jcFrames PainelBottom2 
      Height          =   1875
      Left            =   4290
      Top             =   2250
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   3307
      FrameColor      =   12829635
      Style           =   0
      Caption         =   "Observação"
      TextColor       =   128
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.TextBox txtObs 
         Appearance      =   0  'Flat
         Height          =   1515
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   270
         Width           =   4605
      End
   End
   Begin Tributacao.jcFrames PainelBottom1 
      Height          =   2715
      Left            =   30
      Top             =   2250
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   4789
      FrameColor      =   12829635
      Style           =   0
      Caption         =   "Composição do lançamento"
      TextColor       =   128
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin prjChameleon.chameleonButton cmdQtde 
         Height          =   240
         Left            =   3210
         TabIndex        =   12
         ToolTipText     =   "Altera a Qtde do Tributo"
         Top             =   0
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   423
         BTYPE           =   3
         TX              =   "Qtde"
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
         MICON           =   "frmEmissaoGuia2.frx":0176
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ListView lvTrib 
         Height          =   2340
         Left            =   60
         TabIndex        =   11
         Top             =   300
         Width           =   4050
         _ExtentX        =   7144
         _ExtentY        =   4128
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Desc.Tributo"
            Object.Width           =   4047
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Qtde"
            Object.Width           =   1059
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Valor"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descrição Completa"
            Object.Width           =   6068
         EndProperty
      End
   End
   Begin Tributacao.jcFrames PainelMiddle 
      Height          =   1245
      Left            =   30
      Top             =   990
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   2196
      FrameColor      =   12829635
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   ""
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.TextBox txtNumProc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5955
         MaxLength       =   15
         TabIndex        =   40
         Top             =   870
         Width           =   1005
      End
      Begin VB.CheckBox chkRural 
         Appearance      =   0  'Flat
         BackColor       =   &H00FBFBE3&
         Caption         =   "Rural"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4320
         TabIndex        =   39
         Top             =   900
         Width           =   705
      End
      Begin VB.TextBox txtPercUnica 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7935
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "0"
         Top             =   120
         Width           =   585
      End
      Begin VB.ComboBox cmbExercicio 
         Height          =   315
         ItemData        =   "frmEmissaoGuia2.frx":0192
         Left            =   7950
         List            =   "frmEmissaoGuia2.frx":0194
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   840
         Width           =   975
      End
      Begin VB.CheckBox chkUnica 
         Appearance      =   0  'Flat
         BackColor       =   &H00FBFBE3&
         Caption         =   "Parcela única"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   5295
         TabIndex        =   5
         Top             =   165
         Width           =   1350
      End
      Begin VB.TextBox txtQtdeParc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1215
         MaxLength       =   6
         TabIndex        =   2
         Top             =   120
         Width           =   585
      End
      Begin VB.ComboBox cmbAnoTabela 
         BackColor       =   &H00FBFBE3&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmEmissaoGuia2.frx":0196
         Left            =   7950
         List            =   "frmEmissaoGuia2.frx":0198
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtAbate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         TabIndex        =   8
         Text            =   "0,00"
         Top             =   480
         Width           =   885
      End
      Begin prjChameleon.chameleonButton cmdAddData 
         Height          =   270
         Left            =   4365
         TabIndex        =   4
         ToolTipText     =   "Editar Datas de Vencimento"
         Top             =   135
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   476
         BTYPE           =   3
         TX              =   "+"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmEmissaoGuia2.frx":019A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin esMaskEdit.esMaskedEdit mskDataVencimento 
         Height          =   285
         Left            =   3225
         TabIndex        =   3
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BackColor       =   16777215
         MouseIcon       =   "frmEmissaoGuia2.frx":01B6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "99/99/9999"
         SelText         =   ""
         Text            =   "__/__/____"
         HideSelection   =   -1  'True
      End
      Begin esMaskEdit.esMaskedEdit mskDataInicio 
         Height          =   285
         Left            =   3225
         TabIndex        =   7
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BackColor       =   16777215
         MouseIcon       =   "frmEmissaoGuia2.frx":01D2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "99/99/9999"
         SelText         =   ""
         Text            =   "__/__/____"
         HideSelection   =   -1  'True
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BackStyle       =   0  'Transparent
         Caption         =   "Processo.:"
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   9
         Left            =   5130
         TabIndex        =   41
         Top             =   900
         Width           =   795
      End
      Begin VB.Label lblUfir22 
         BackStyle       =   0  'Transparent
         Caption         =   "Índice referência..:"
         Height          =   225
         Left            =   2160
         TabIndex        =   32
         Top             =   900
         Width           =   1335
      End
      Begin VB.Label lblUfir 
         BackStyle       =   0  'Transparent
         Caption         =   "0,0000"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   3600
         TabIndex        =   31
         Top             =   900
         Width           =   645
      End
      Begin VB.Label lblMesProp 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   1770
         TabIndex        =   30
         Top             =   870
         Width           =   315
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Meses proporcionais..:"
         Height          =   225
         Index           =   4
         Left            =   90
         TabIndex        =   29
         Top             =   900
         Width           =   1635
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "% ùnica...:"
         Height          =   225
         Index           =   3
         Left            =   7080
         TabIndex        =   28
         Top             =   165
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Data 1º Vencto.:"
         Height          =   225
         Index           =   15
         Left            =   1965
         TabIndex        =   26
         Top             =   165
         Width           =   1260
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde Parc.....:"
         Height          =   225
         Index           =   14
         Left            =   90
         TabIndex        =   25
         Top             =   165
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Exercício.:"
         Height          =   225
         Index           =   17
         Left            =   7080
         TabIndex        =   24
         Top             =   900
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cálculo proporcional a partir da data de.....:"
         Height          =   225
         Index           =   16
         Left            =   90
         TabIndex        =   23
         Top             =   525
         Width           =   3105
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Tabela.....:"
         Height          =   240
         Index           =   0
         Left            =   7080
         TabIndex        =   22
         Top             =   525
         Width           =   825
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Abatimento em NF:"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   4500
         TabIndex        =   21
         Top             =   525
         Width           =   1425
      End
   End
   Begin Tributacao.jcFrames PainelTop 
      Height          =   915
      Left            =   30
      Top             =   45
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   1614
      FrameColor      =   12829635
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   ""
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   16514019
      ColorTo         =   0
      Begin VB.ListBox lstAtividade 
         Appearance      =   0  'Flat
         Height          =   705
         Left            =   5040
         Style           =   1  'Checkbox
         TabIndex        =   27
         Top             =   90
         Width           =   3795
      End
      Begin VB.ComboBox cmbTipoGuia 
         Height          =   315
         ItemData        =   "frmEmissaoGuia2.frx":01EE
         Left            =   1230
         List            =   "frmEmissaoGuia2.frx":01F0
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   90
         Width           =   2745
      End
      Begin VB.ComboBox cmbLanc 
         Height          =   315
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   3690
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de guia..:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   20
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lançamento..:"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   19
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Atividade..:"
         Height          =   195
         Index           =   2
         Left            =   4140
         TabIndex        =   18
         Top             =   180
         Width           =   855
      End
   End
   Begin Tributacao.jcFrames pnlData 
      Height          =   4785
      Left            =   3180
      Top             =   90
      Visible         =   0   'False
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   8440
      FillColor       =   14745599
      Style           =   4
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Vencimentos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin MSFlexGridLib.MSFlexGrid grdData 
         Height          =   3795
         Left            =   90
         TabIndex        =   35
         Top             =   480
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   6694
         _Version        =   393216
         RowHeightMin    =   280
         BackColor       =   16777215
         BackColorBkg    =   12648447
         Appearance      =   0
         FormatString    =   "Parcela        |^Data               "
      End
      Begin prjChameleon.chameleonButton cmdRetornar 
         Height          =   345
         Left            =   720
         TabIndex        =   36
         ToolTipText     =   "Sair da Tela"
         Top             =   4350
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "&Retornar"
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
         MICON           =   "frmEmissaoGuia2.frx":01F2
         PICN            =   "frmEmissaoGuia2.frx":020E
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
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      Caption         =   $"frmEmissaoGuia2.frx":027C
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Index           =   5
      Left            =   0
      TabIndex        =   37
      Top             =   4980
      Width           =   9075
   End
   Begin VB.Label lblValorParcela 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   5685
      TabIndex        =   34
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Parcela.:"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   4500
      TabIndex        =   33
      Top             =   4680
      Width           =   1035
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Única...:"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   4500
      TabIndex        =   17
      Top             =   4200
      Width           =   1035
   End
   Begin VB.Label lblTotalUnica 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   5685
      TabIndex        =   16
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total.....:"
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   0
      Left            =   4500
      TabIndex        =   15
      Top             =   4440
      Width           =   1035
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   5685
      TabIndex        =   14
      Top             =   4440
      Width           =   975
   End
End
Attribute VB_Name = "frmEmissaoGuia2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sObsIss As String, nAreaIss As Double, nAreaTerreno As Double

Private Sub chkUnica_Click()
If chkUnica.value = vbUnchecked Then
    lblTotalUnica.Caption = "0,0"
    txtPercUnica.Text = "0"
    FillTotal
End If
End Sub

Private Sub cmbLanc_Click()
Dim nCodLanc As Integer, nCodReduz As Long, nTipo As Integer

nCodReduz = Val(frmEmissaoGuia.txtCodigo.Text)
If cmbLanc.ItemData(cmbLanc.ListIndex) = 50 Or cmbLanc.ItemData(cmbLanc.ListIndex) = 65 Then
    txtAbate.Locked = False
    txtAbate.BackColor = Branco
Else
    txtAbate.Locked = True
    txtAbate.BackColor = &HFBFBE3
End If
txtAbate.Text = "0,00"
nCodLanc = cmbLanc.ItemData(cmbLanc.ListIndex)

If nCodReduz < 100000 Then
    nTipo = 1
ElseIf nCodReduz >= 100000 And nCodReduz < 300000 Then
    nTipo = 2
ElseIf nCodReduz >= 500000 And nCodReduz < 700000 Then
    nTipo = 3
End If

lvTrib.ListItems.Clear
If nTipo = 1 And (cmbTipoGuia.ListIndex = 2 Or cmbTipoGuia.ListIndex = 3 Or cmbTipoGuia.ListIndex = 4) Then
    MsgBox "Um imóvel não pode ter lançamentos de ISS, Taxa de Licença e Vigilância Sanitária.", vbExclamation, "Atenção"
    cmbTipoGuia.ListIndex = 0
    Exit Sub
End If

If nTipo > 1 And cmbTipoGuia.ListIndex = 5 Then
    MsgBox "Apenas imóveis podem ter lançamentos de Roçada.", vbExclamation, "Atenção"
    cmbTipoGuia.ListIndex = 0
    Exit Sub
End If

If nTipo > 1 And cmbTipoGuia.ListIndex = 1 Then
    MsgBox "Apenas imóveis podem ter lançamentos de IPTU.", vbExclamation, "Atenção"
    cmbTipoGuia.ListIndex = 0
    Exit Sub
End If

'Select Case nCodLanc
'    Case 1, 2, 3, 5, 6, 11, 13, 14, 33, 48
'        mskDataInicio.Locked = False
'        mskDataInicio.BackColor = Branco
'        chkUnica.Enabled = True
'        txtQtdeParc.Locked = False
'        txtQtdeParc.BackColor = Branco
'        txtPercUnica.Locked = False
'        txtPercUnica.BackColor = Branco
'    Case Else
'        mskDataInicio.Locked = True
'        mskDataInicio.BackColor = Me.BackColor
'        chkUnica.value = vbUnchecked
'        chkUnica.Enabled = False
'        txtQtdeParc.Text = "1"
'        txtQtdeParc.BackColor = Me.BackColor
'        txtQtdeParc.Locked = True
'        txtPercUnica.Text = "0"
'        txtPercUnica.BackColor = Me.BackColor
'        txtPercUnica.Locked = False
'End Select
If mskDataVencimento.Visible Then
    mskDataVencimento.SetFocus
End If


CarregaTributo nCodLanc

End Sub

Private Sub cmbTipoGuia_Click()

lstAtividade.Clear
If cmbTipoGuia.ListIndex = -1 Then Exit Sub
lvTrib.ListItems.Clear
If cmbTipoGuia.ListIndex = 1 Then
    CarregaLancamento 1
ElseIf cmbTipoGuia.ListIndex = 2 Then
    If Val(frmEmissaoGuia.txtCodigo.Text) >= 500000 Then
        CarregaLancamento 6
    Else
        CarregaTL
        If lstAtividade.ListCount > 0 Then
            CarregaLancamento 6
        Else
            MsgBox "Empresa não possui atividade de taxa de licença cadastrada.", vbCritical, "Erro"
            cmbTipoGuia.ListIndex = 0
        End If
    End If
ElseIf cmbTipoGuia.ListIndex = 3 Then
    CarregaISS
    If lstAtividade.ListCount > 0 Then
        CarregaLancamento 14
    Else
        MsgBox "Empresa não possui atividade de ISS Fixo cadastrada.", vbCritical, "Erro"
        cmbTipoGuia.ListIndex = 0
    End If
ElseIf cmbTipoGuia.ListIndex = 4 Then
    CarregaVS
    CarregaLancamento 13
ElseIf cmbTipoGuia.ListIndex = 5 Then
    CarregaLancamento 38
Else
    CarregaLancamento 0
End If

End Sub

Private Sub cmdAddData_Click()
If Val(txtQtdeParc.Text) = 0 Then
    txtQtdeParc.Text = 1
End If

If Val(txtQtdeParc.Text) > 48 Then
    MsgBox "Máximo 48 parcelas.", vbExclamation, "Atenção"
    Exit Sub
End If

If Not IsDate(mskDataVencimento.Text) Then
    MsgBox "Digite o 1º vencimento.", vbExclamation, "Atenção"
    mskDataVencimento.SetFocus
    Exit Sub
End If

If grdData.Rows = 1 Then
    mskDataVencimento_LostFocus
End If

pnlData.Visible = True
pnlData.ZOrder 0
cmdPrint.Enabled = False
cmbTipoGuia.Enabled = False
cmbTipoGuia.BackColor = Me.BackColor
cmbLanc.Enabled = False
cmbLanc.BackColor = Me.BackColor
lstAtividade.BackColor = Me.BackColor
lstAtividade.Enabled = False
lvTrib.Enabled = False
lvTrib.BackColor = Me.BackColor
cmbExercicio.Enabled = False
cmbExercicio.BackColor = Me.BackColor
cmbAnoTabela.Enabled = False
cmbAnoTabela.BackColor = Me.BackColor
txtObs.Enabled = False
txtQtdeParc.Enabled = False
txtAbate.Enabled = False
txtPercUnica.Enabled = False

End Sub

Private Sub cmdPrint_Click()
Dim x As Integer, bFind As Boolean, z As Variant, Sql As String, RdoAux As rdoResultset, NumProc As Long, AnoProc As Integer, nDV As Integer
Dim sValidaProc As String, nCodReduz As Long

nCodReduz = Val(frmEmissaoGuia.txtCodigo.Text)
bFind = False
For x = 0 To Forms.Count - 1
    If Forms(x).Name = "frmEmissaoGuia" Then
        bFind = True
    End If
Next

If Not bFind Then
    MsgBox "Você fechou a tela principal da emissão de guias, você deverá reiniciar novamente a operação.", vbExclamation, "Atenção"
    Unload Me
    Exit Sub
End If

If cmbLanc.ListIndex = -1 Then
    MsgBox "Selecione o Lançamento.", vbExclamation, "Atenção"
    Exit Sub
End If

If Val(txtQtdeParc.Text) = 0 Then
    MsgBox "Digite a qtde de parcelas.", vbExclamation, "Atenção"
    Exit Sub
End If

If chkUnica.value = vbChecked And Val(txtPercUnica.Text) = 0 Then
    MsgBox "Digite o % da parcela única.", vbExclamation, "Atenção"
    Exit Sub
End If

If lblTotal.Caption = "0,00" Then
    MsgBox "Selecione ao menos um tributo.", vbExclamation, "Atenção"
    Exit Sub
End If

If Not IsDate(mskDataVencimento.Text) Then
    MsgBox "Data de 1º vencimento inválida.", vbExclamation, "Atenção"
    Exit Sub
End If

For x = 1 To grdData.Rows - 1
    If Not IsDate(grdData.TextMatrix(x, 1)) Then
        MsgBox "Data de vencimento da parcela nº " & x & " inválida.", vbExclamation, "Atenção"
        Exit Sub
    End If
Next

For x = 1 To grdData.Rows - 2
    If CDate(grdData.TextMatrix(x, 1)) >= CDate(grdData.TextMatrix(x + 1, 1)) Then
        MsgBox "Data de vencimento da parcela nº " & x & " não pode ser maior ou igual ao vencimento da parcela nº " & x + 1 & ".", vbExclamation, "Atenção"
        Exit Sub
    End If
Next

If cmbTipoGuia.ListIndex = 5 Then
    
        
'    z = InputBox("Digite o número do processo.", "Entre com os dados")

    If Trim(txtNumProc.Text) = "" Then
        MsgBox "Digite o número do processo.", vbCritical, "Atenção"
        Exit Sub
    End If
    sValidaProc = ValidaProcesso(txtNumProc.Text)
    NumeroProcesso = txtNumProc.Text
    NumProc = Left$(NumeroProcesso, InStr(1, NumeroProcesso, "/", vbBinaryCompare) - 2)
    AnoProc = Right$(NumeroProcesso, 4)

    If InStr(1, NumeroProcesso, "/", vbBinaryCompare) = 0 Or Len(NumeroProcesso) < 2 Then
        GoTo Erro
    Else
        If Right(NumeroProcesso, 1) = "/" Then
            GoTo Erro
        End If
    End If
    nDV = Val(Right(Left$(NumeroProcesso, InStr(1, NumeroProcesso, "/", vbBinaryCompare) - 1), 1))
    If InStr(1, NumeroProcesso, "/", vbBinaryCompare) < 3 Then
        GoTo Erro
    End If
    NumProc = Left$(NumeroProcesso, InStr(1, NumeroProcesso, "/", vbBinaryCompare) - 2)
    AnoProc = Right$(NumeroProcesso, 4)
    If nDV <> RetornaDVProcesso(NumProc) Then
        GoTo Erro
    End If
    Sql = "SELECT * FROM PROCESSOGTI WHERE ANO=" & AnoProc & " AND NUMERO=" & NumProc
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount = 0 Then
        GoTo Erro
    End If
    txtObs.Text = "Número do Processo: " & NumeroProcesso
    txtObs.SelStart = Len(txtObs.Text)
    
    Sql = "SELECT * FROM ETIQUETAROCADA WHERE CODREDUZIDO=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
            Sql = "INSERT ETIQUETAROCADA(CODREDUZIDO,DATA) VALUES(" & nCodReduz & ",'" & Format(Now, "mm/dd/yyyy") & "')"
        Else
            Sql = "UPDATE ETIQUETAROCADA SET DATA='" & Format(Now, "mm/dd/yyyy") & "' WHERE CODREDUZIDO=" & nCodReduz
        End If
        cn.Execute Sql, rdExecDirect
       .Close
    End With
    EmiteBoletoRocada
    frmReport.ShowReport3 "MULTAINF", frmMdi.HWND, Me.HWND, nCodReduz
    Unload Me
Else
    EmiteBoleto
End If

Exit Sub
Erro:
MsgBox "Número de Processo inválido ou não cadastrado.", vbCritical, "Atenção"
End Sub

Private Sub cmdQtde_Click()
Dim z As Variant, nCodTributo As Integer

If lvTrib.ListItems.Count = 0 Then
    MsgBox "Não existem tributos.", vbExclamation, "Atenção"
    Exit Sub
End If

If lvTrib.SelectedItem.Checked = False Then
    MsgBox "Selecione o item antes de alterar sua qtde.", vbExclamation, "Atenção"
    Exit Sub
End If

nCodTributo = Val(Right$(lvTrib.ListItems(lvTrib.SelectedItem.Index).Key, 3))
Select Case nCodTributo
    Case 11, 12, 13, 14, 25, 170
        MsgBox "A Quantidade deste tributo não pode ser alterada.", vbExclamation, "Atenção"
    Case Else
        z = InputBox("Digite a Qtde.", "Quantidade", lvTrib.ListItems(lvTrib.SelectedItem.Index).SubItems(1))
        If z = "" Then z = lvTrib.ListItems(lvTrib.SelectedItem.Index).SubItems(1)
        z = Ponto2Virg(CStr(z))
        If Not IsNumeric(z) Then
            MsgBox "Qtde inválida.", vbExclamation, "Atenção"
        Else
            If z = 0 Then z = 1
            lvTrib.ListItems(lvTrib.SelectedItem.Index).SubItems(1) = z
            FillTotal
        End If
End Select

End Sub

Private Sub cmdRetornar_Click()
pnlData.Visible = False
cmdPrint.Enabled = True
cmbTipoGuia.Enabled = True
cmbTipoGuia.BackColor = Branco
cmbLanc.Enabled = True
cmbLanc.BackColor = Branco
lstAtividade.BackColor = Branco
lstAtividade.Enabled = True
lvTrib.Enabled = True
lvTrib.BackColor = Branco
cmbExercicio.Enabled = True
cmbExercicio.BackColor = Branco
cmbAnoTabela.Enabled = True
cmbAnoTabela.BackColor = Branco
txtObs.Enabled = True
txtQtdeParc.Enabled = True
txtAbate.Enabled = True
txtPercUnica.Enabled = True
If grdData.Rows > 1 Then
    mskDataVencimento.Text = grdData.TextMatrix(1, 1)
End If

End Sub

Private Sub Form_Activate()
txtQtdeParc.SetFocus
End Sub

Private Sub Form_Load()
Dim x As Integer

Me.Top = frmEmissaoGuia.Top + 2500
Me.Left = frmEmissaoGuia.Left + 1000

cmbTipoGuia.AddItem "(Lançamentos diversos)"
cmbTipoGuia.AddItem "IPTU/ITU"
cmbTipoGuia.AddItem "Taxa de Licença"
cmbTipoGuia.AddItem "ISS Fixo"
cmbTipoGuia.AddItem "Vigilância Sanitária"
cmbTipoGuia.AddItem "Roçada"
cmbTipoGuia.ListIndex = 0

For x = 1994 To Year(Now)
    cmbExercicio.AddItem x
Next
cmbExercicio.Text = Year(Now)

For x = 2011 To Year(Now)
    cmbAnoTabela.AddItem x
Next
cmbAnoTabela.Text = Year(Now)

txtAbate.Locked = True
txtAbate.BackColor = &HFBFBE3

cmbAnoTabela.Text = Year(Now)
If NomeDeLogin = "SCHWARTZ" Or NomeDeLogin = "RENATA" Or NomeDeLogin = "GLEISE" Or NomeDeLogin = "RITA" Or _
    NomeDeLogin = "LEANDRO" Or NomeDeLogin = "LUIZH" Or NomeDeLogin = "SOLANGE" Or NomeDeLogin = "RODRIGOG" Or NomeDeLogin = "AAFMARTINS" Or IsAtendente Then
    cmbAnoTabela.Enabled = True
    cmbAnoTabela.BackColor = Branco
End If

End Sub

Private Sub CarregaLancamento(nCodigo As Integer)
Dim Sql As String, RdoAux As rdoResultset

cmbLanc.Clear
Sql = "select codlancamento, descreduz from lancamento "
If nCodigo = 1 Or nCodigo = 6 Or nCodigo = 13 Or nCodigo = 14 Or nCodigo = 38 Then
    Sql = Sql & "where codlancamento=" & nCodigo
Else
    Sql = Sql & "where codlancamento not in (1,2,3,6,8,12,13,14,20,21,30) "
End If
Sql = Sql & "order by descreduz"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If cmbTipoGuia.ListIndex <> 5 And (!CodLancamento = 38 Or !CodLancamento = 10) Then
        Else
            cmbLanc.AddItem !descreduz
            cmbLanc.ItemData(cmbLanc.NewIndex) = !CodLancamento
        End If
       .MoveNext
    Loop
   .Close
End With

If cmbLanc.ListCount > 0 Then cmbLanc.ListIndex = 0

End Sub

Private Sub grdData_DblClick()
Dim z As Variant, nRow As Integer
nRow = grdData.Row

Inicio:
z = InputBox("Digite o novo vencimento para a parcela " & nRow, "Alteração de Vencimento", grdData.TextMatrix(grdData.Row, 1))
If IsDate(z) Then
    If Year(CDate(z)) < 1996 Or Year(CDate(z)) > Year(Now) + 1 Then
        MsgBox "Data de vencimento inválida.", vbCritical, "Erro"
    Else
        If nRow > 1 Then
            If CDate(z) < CDate(grdData.TextMatrix(nRow - 1, 1)) Then
                MsgBox "Data da parcela " & nRow & " não pode ser inferior ou igual a parcela anterior.", vbExclamation, "Atenção"
                GoTo Inicio
            End If
        End If
        grdData.TextMatrix(nRow, 1) = z
    End If
End If

End Sub

Private Sub lstAtividade_Click()
Dim x As Integer
For x = 1 To lvTrib.ListItems.Count
    If Val(Right$(lvTrib.ListItems(x).Key, 3)) = 11 Or Val(Right$(lvTrib.ListItems(x).Key, 3)) = 25 Or Val(Right$(lvTrib.ListItems(x).Key, 3)) = 14 Then
        lvTrib.ListItems(x).Selected = True
        If lstAtividade.Selected(lstAtividade.ListIndex) = True Then
            lvTrib.SelectedItem.Checked = True
            lvTrib_ItemCheck lvTrib.SelectedItem
        Else
            lvTrib.SelectedItem.Checked = False
            lvTrib.ListItems(x).SubItems(1) = "0"
            lvTrib.ListItems(x).SubItems(2) = "0,0000"
            lvTrib.ListItems(x).ForeColor = vbBlack
            FillTotal
            Exit Sub
        End If
        Exit For
    End If
Next


End Sub

Private Sub lvTrib_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim nAno As Integer, nCodTributo As Integer, nValorTributo As Double, Sql As String, RdoAux As rdoResultset, nCodigo As Long, nValorAliquota As Double, nCodAtividade As Integer
Dim nQtdeProf As Integer, nArea As Double, x As Integer, sCnae As String, nCodCriterio As Integer, y As Integer

nCodigo = Val(frmEmissaoGuia.txtCodigo.Text)
nAno = Val(cmbAnoTabela.Text)
nCodTributo = Val(Right$(lvTrib.ListItems(Item.Index).Key, 3))
Item.Selected = True

If nCodTributo = 154 Or nCodTributo = 155 Or nCodTributo = 156 Then
    MsgBox "Uso de plataforma deve ser emitido através do site da prefeitura.", vbCritical, "Atenção"
    lvTrib.SelectedItem.Checked = False
    Exit Sub
End If

With lvTrib
    If .ListItems(Item.Key).Checked = True Then
        If nCodTributo = 14 Then
            Sql = "SELECT MOBILIARIO.CODATIVIDADE,QTDEPROF ,DESCATIVIDADE,VALORALIQ1,VALORALIQ2,VALORALIQ3,AREATL,CODIGOALIQ FROM MOBILIARIO INNER JOIN "
            Sql = Sql & "ATIVIDADE ON MOBILIARIO.CODATIVIDADE = ATIVIDADE.CODATIVIDADE Where CODIGOMOB =" & nCodigo
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount > 0 Then
                    Select Case !CODIGOALIQ
                        Case 1
                            nValorAliquota = !VALORALIQ1
                        Case 2
                            nValorAliquota = !VALORALIQ2
                        Case 3
                            nValorAliquota = !VALORALIQ3
                    End Select
                    If Not IsNull(!areatl) Then
                       nArea = IIf(!areatl = 0, 1, !areatl)
                    Else
                       nArea = 1
                    End If
                    nQtdeProf = Val(SubNull(!QTDEPROF))
                    If nQtdeProf = 0 Then nQtdeProf = 1
                End If
            End With
            If nValorAliquota < 14 Then
                nValorTributo = (nValorAliquota * RetornaUFIR(nAno) * nArea) * MesesProporcional / 12
            Else
                nValorTributo = (nValorAliquota * RetornaUFIR(nAno)) * MesesProporcional / 12
            End If
            lvTrib.ListItems(Item.Index).SubItems(1) = nQtdeProf
            lvTrib.ListItems(Item.Index).SubItems(2) = FormatNumber(nValorTributo, 4)
            lvTrib.ListItems(Item.Index).ForeColor = vbRed
        ElseIf nCodTributo = 11 Then
            Sql = "SELECT DISTINCT MOBILIARIOATIVIDADEISS.CODMOBILIARIO,MOBILIARIOATIVIDADEISS.CODTRIBUTO,MOBILIARIOATIVIDADEISS.CODATIVIDADE,"
            Sql = Sql & "MOBILIARIOATIVIDADEISS.QTDEISS,MOBILIARIOATIVIDADEISS.VALORISS,"
            Sql = Sql & "TABELAISS.ALIQUOTA FROM MOBILIARIOATIVIDADEISS INNER JOIN TABELAISS ON MOBILIARIOATIVIDADEISS.CODTRIBUTO = TABELAISS.TIPOISS AND "
            Sql = Sql & "MOBILIARIOATIVIDADEISS.CODATIVIDADE = TABELAISS.CODIGOATIV Where MOBILIARIOATIVIDADEISS.CODMOBILIARIO = " & nCodigo
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount > 0 Then
                    nValorTributo = 0
                    Do Until .EOF
                        For x = 0 To lstAtividade.ListCount - 1
                            If lstAtividade.Selected(x) = True Then
                                nCodAtividade = Val(Left(lstAtividade.List(x), 3))
                                If nCodAtividade = !codatividade Then
                                    nValorAliquota = RetornaAliquotaISS(!codatividade, Format(Now, "dd/mm/yyyy"))
                                    nQtdeProf = Val(SubNull(!QTDEISS))
                                    nValorTributo = nValorTributo + (nValorAliquota)
                                End If
                            End If
                        Next
                       .MoveNext
                    Loop
                End If
            End With
            nValorTributo = nValorTributo * RetornaUFIR(nAno) * (MesesProporcional / 12)
            lvTrib.ListItems(Item.Index).SubItems(1) = nQtdeProf
            lvTrib.ListItems(Item.Index).SubItems(2) = FormatNumber(nValorTributo, 4)
            lvTrib.ListItems(Item.Index).ForeColor = vbRed
        ElseIf nCodTributo = 25 Then
            nValorTributo = 0
            For x = 0 To lstAtividade.ListCount - 1
                If lstAtividade.Selected(x) = True Then
                    nCodCriterio = lstAtividade.ItemData(x)
                    Sql = "SELECT distinct mobiliariovs.codigo, mobiliariovs.cnae, mobiliariovs.criterio, mobiliariovs.qtde,  cnaecriteriodesc.valor "
                    Sql = Sql & "FROM mobiliariovs INNER JOIN cnae_criterio ON mobiliariovs.cnae = cnae_criterio.cnae INNER JOIN cnae ON mobiliariovs.cnae = cnae.cnae "
                    Sql = Sql & "INNER JOIN cnaecriteriodesc ON mobiliariovs.criterio = cnaecriteriodesc.criterio WHERE mobiliariovs.codigo = " & nCodigo & " and mobiliariovs.criterio=" & nCodCriterio
                    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux
                        If .RowCount > 0 Then
                            Do Until .EOF
                                sCnae = RetornaNumero(Left(lstAtividade.List(x), 9))
                                If sCnae = !Cnae And nCodCriterio = !criterio Then
                                    nQtdeProf = Val(SubNull(!QTDE))
                                    nValorTributo = nValorTributo + (!Valor)
                                End If
                               .MoveNext
                            Loop
                        End If
                    End With
                End If
            Next
            nValorTributo = nValorTributo * (MesesProporcional / 12)
            lvTrib.ListItems(Item.Index).SubItems(1) = nQtdeProf
            lvTrib.ListItems(Item.Index).SubItems(2) = FormatNumber(nValorTributo, 4)
            lvTrib.ListItems(Item.Index).ForeColor = vbRed
        Else
            Sql = "SELECT VALORALIQ FROM TRIBUTOALIQUOTA WHERE ANO=" & nAno & " AND CODTRIBUTO = " & nCodTributo
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount > 0 Then
                    
                    If Val(lvTrib.ListItems(Item.Index).SubItems(1)) = 0 Then
                        If nCodTributo = 170 Then
                            lvTrib.ListItems(Item.Index).SubItems(1) = Format(nAreaTerreno, "#0.00")
                        Else
                            lvTrib.ListItems(Item.Index).SubItems(1) = "1"
                        End If
                    End If
                    
                    lvTrib.ListItems(Item.Index).SubItems(2) = FormatNumber(!valoraliq, 4)
                    lvTrib.ListItems(Item.Index).ForeColor = vbRed
                Else
                    MsgBox "Não existe tarifa para este tributo." & vbCrLf & "Consulte a Tabela de Preços Públicos.", vbExclamation, "Atenção"
                    lvTrib.ListItems(Item.Index).Checked = False
                    lvTrib.ListItems(Item.Index).ForeColor = vbBlack
                End If
            End With
        End If
    Else
        If nCodTributo = 14 Then
            .ListItems(Item.Key).Checked = True
        Else
            .ListItems(Item.Index).SubItems(1) = "0"
            .ListItems(Item.Index).SubItems(2) = "0,0000"
            lvTrib.ListItems(Item.Index).ForeColor = vbBlack
        End If
    End If
End With
FillTotal


End Sub

Private Sub mskDataInicio_GotFocus()
mskDataInicio.SelStart = 0
mskDataInicio.SelLength = Len(mskDataInicio.Text)

End Sub

Private Sub mskDataVencimento_GotFocus()
mskDataVencimento.SelStart = 0
mskDataVencimento.SelLength = Len(mskDataVencimento.Text)
End Sub

Private Sub mskDataVencimento_LostFocus()
grdData.Rows = 1


If Not IsDate(mskDataVencimento.Text) Then Exit Sub

df = ValidaFeriado(CDate(mskDataVencimento.Text))
If df = 1 Then
    If MsgBox("Data do 1º Vencimento cai no Domingo." & vbCrLf & "Próximo Dia Util é " & RetornaDiaUtil(CDate(mskDataVencimento.Text)) & vbCrLf & vbCrLf & "Aceitar esta Data?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
        mskDataVencimento.Text = Format(RetornaDiaUtil(CDate(mskDataVencimento.Text)), "dd/mm/yyyy")
    End If
ElseIf df = 2 Then
    If MsgBox("Data do 1º Vencimento cai no sábado." & vbCrLf & "Próximo Dia Util é " & RetornaDiaUtil(CDate(mskDataVencimento.Text)) & vbCrLf & vbCrLf & "Aceitar esta Data?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
        mskDataVencimento.Text = Format(RetornaDiaUtil(CDate(mskDataVencimento.Text)), "dd/mm/yyyy")
    Else
        If grdData.Rows = 1 Then
            grdData.Rows = Val(txtNumParc.Text) + 1
            AutoFillDate2
            For x = 1 To Val(txtNumParc.Text)
                grdData.TextMatrix(x, 0) = "Parcela " & Format(x, "00")
            Next
        End If
        Exit Sub
    End If
ElseIf df = 3 Then
    Sql = "SELECT NOMEFERIADO FROM FERIADODEF INNER JOIN "
    Sql = Sql & "FERIADO ON FERIADODEF.CODFERIADO = FERIADO.CODFERIADO "
    Sql = Sql & " Where DIA = " & Day(CDate(mskDataVencimento.Text))
    Sql = Sql & " AND MES=" & Month(CDate(mskDataVencimento.Text)) & " AND ANO=" & Year(CDate(mskDataVencimento.Text))
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            If MsgBox("Data do 1º Vencimento cai no Feriado (" & !NOMEFERIADO & ")" & vbCrLf & "Próximo Dia Util é " & RetornaDiaUtil(CDate(mskDataVencimento.Text)) & vbCrLf & vbCrLf & "Aceitar esta Data?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
                mskDataVencimento.Text = RetornaDiaUtil(CDate(mskDataVencimento.Text))
            Else
                'Exit Sub
            End If
          .Close
        End If
    End With
End If

If grdData.Rows = 1 Then
    grdData.Rows = Val(txtQtdeParc.Text) + 1
    AutoFillDate2
    For x = 1 To Val(txtQtdeParc.Text)
        grdData.TextMatrix(x, 0) = "Parcela " & Format(x, "00")
    Next
End If

End Sub

Private Sub AutoFillDate2()
Dim x As Integer, sData As String, sDiaIni As String
sData = mskDataVencimento.Text

If Not IsDate(sData) Then
    MsgBox "Data inválida", vbExclamation, "Atenção"
    Exit Sub
End If


sDiaIni = Left(sData, 2)
For x = 1 To Val(txtQtdeParc.Text)
    grdData.TextMatrix(x, 1) = sData
    sData = Format(DateAdd("m", 1, sData), "dd/mm/yyyy")

    sData = sDiaIni & "/" & Mid(sData, 4, 2) & "/" & Right(sData, 4)
Inicio:
    If Not IsDate(sData) Then
        sData = Format(Val(Left(sData, 2)) - 1, "00") & "/" & Mid(sData, 4, 2) & "/" & Right(sData, 4)
        GoTo Inicio
    End If
Next

End Sub

Private Sub txtPercUnica_Change()
FillTotal
End Sub

Private Sub txtPercUnica_KeyPress(KeyAscii As Integer)
Tweak txtPercUnica, KeyAscii, IntegerPositive
End Sub

Private Sub txtQtdeParc_Change()
grdData.Rows = 1
If Val(txtQtdeParc.Text) <= 1 Then
   chkUnica.value = vbUnchecked
   txtPercUnica.Text = "0"
End If

FillTotal
End Sub

Private Sub txtQtdeParc_GotFocus()
txtQtdeParc.SelStart = 0
txtQtdeParc.SelLength = Len(txtQtdeParc.Text)

End Sub

Private Sub txtQtdeParc_KeyPress(KeyAscii As Integer)
Tweak txtQtdeParc, KeyAscii, IntegerPositive
End Sub

Private Sub CarregaTributo(CodLancamento As Integer)
Dim Sql As String, RdoAux As rdoResultset, itmX As ListItem, nCodigo As Long

nCodigo = Val(frmEmissaoGuia.txtCodigo.Text)
lvTrib.ListItems.Clear
lblTotal.Caption = "0,00"
lblTotalUnica.Caption = "0,00"

If CodLancamento = 38 Then
    Sql = "select codreduzido,Dt_AreaTerreno from cadimob where codreduzido=" & nCodigo
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    nAreaTerreno = RdoAux!Dt_AreaTerreno
    RdoAux.Close
End If

'Sql = "SELECT CODTRIBUTO,ABREVTRIBUTO,DESCTRIBUTO FROM vwTRIBUTOLANCAMENTO WHERE CODLANCAMENTO=" & CodLancamento & " AND CODTRIBUTO<>3 and codtributo<>13 ORDER BY ABREVTRIBUTO"
Sql = "SELECT CODTRIBUTO,ABREVTRIBUTO,DESCTRIBUTO FROM vwTRIBUTOLANCAMENTO WHERE CODLANCAMENTO=" & CodLancamento & "  ORDER BY ABREVTRIBUTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If CodLancamento = 5 Then
            If !CodTributo = 13 Then
                GoTo Proximo
            End If
        End If
        If !CodTributo <> 124 Then
            Set itmX = lvTrib.ListItems.Add(, "C" & Format(!CodTributo, "000"), !ABREVTRIBUTO)
            itmX.SubItems(1) = "0,00"
            itmX.SubItems(2) = "0,0000"
            itmX.SubItems(3) = !desctributo
        End If
Proximo:
       .MoveNext
    Loop
   .Close
End With


End Sub

Private Sub CarregaISS()
Dim Sql As String, RdoAux As rdoResultset, nCodReduz As Long

nCodReduz = Val(frmEmissaoGuia.txtCodigo.Text)
lstAtividade.Clear
Sql = "SELECT MOBILIARIOATIVIDADEISS.CODATIVIDADE,ATIVIDADEISS.DESCATIVIDADE FROM MOBILIARIOATIVIDADEISS INNER JOIN "
Sql = Sql & "ATIVIDADEISS ON MOBILIARIOATIVIDADEISS.CODATIVIDADE = ATIVIDADEISS.CODATIVIDADE "
Sql = Sql & "Where MOBILIARIOATIVIDADEISS.CODMOBILIARIO = " & nCodReduz & " AND CODTRIBUTO=11"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    Do Until .EOF
       lstAtividade.AddItem Format(!codatividade, "000") & "-" & !descatividade
      .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub CarregaTL()
Dim Sql As String, RdoAux As rdoResultset, nCodReduz As Long

nCodReduz = Val(frmEmissaoGuia.txtCodigo.Text)
lstAtividade.Clear
Sql = "SELECT MOBILIARIO.CODATIVIDADE,ATIVIDADE.DESCATIVIDADE FROM MOBILIARIO INNER JOIN "
Sql = Sql & "ATIVIDADE ON MOBILIARIO.CODATIVIDADE = ATIVIDADE.CODATIVIDADE "
Sql = Sql & "Where MOBILIARIO.CODIGOMOB = " & nCodReduz
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    Do Until .EOF
       lstAtividade.AddItem Format(!codatividade, "000") & "-" & !descatividade
      .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub CarregaVS()
Dim Sql As String, RdoAux As rdoResultset, nCodReduz As Long

nCodReduz = Val(frmEmissaoGuia.txtCodigo.Text)
lstAtividade.Clear
Sql = "SELECT distinct mobiliariovs.codigo, mobiliariovs.cnae, mobiliariovs.criterio, mobiliariovs.qtde, cnae.descricao, cnaecriteriodesc.valor "
Sql = Sql & "FROM mobiliariovs INNER JOIN cnae_criterio ON mobiliariovs.cnae = cnae_criterio.cnae INNER JOIN cnae ON mobiliariovs.cnae = cnae.cnae "
Sql = Sql & "INNER JOIN cnaecriteriodesc ON mobiliariovs.criterio = cnaecriteriodesc.criterio WHERE mobiliariovs.codigo = " & nCodReduz
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    Do Until .EOF
       lstAtividade.AddItem Left(!Cnae, 4) & "-" & Mid(!Cnae, 5, 1) & "/" & Right(!Cnae, 2) & "-" & !Descricao
       lstAtividade.ItemData(lstAtividade.NewIndex) = !criterio
      .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub FillTotal()
Dim nTotal As Double, nTotalUnica As Double, Sql As String, RdoAux As rdoResultset, nValorAbate As Double
Dim nPerc As Double, nValorTotal As Double, nQtde As Integer
nTotal = 0: nTotalUnica = 0

If Val(txtQtdeParc.Text) = 0 Then
    nQtde = 1
Else
    nQtde = Val(txtQtdeParc.Text)
End If

lblMesProp.Caption = MesesProporcional
lblUfir.Caption = Format(RetornaUFIR(Year(Now)), "#0.0000")

For x = 1 To lvTrib.ListItems.Count
    If lvTrib.ListItems(x).Checked = True Then
        If lvTrib.ListItems(x).SubItems(1) = "" Then lvTrib.ListItems(x).SubItems(1) = "0"
        nTotal = nTotal + (CDbl(lvTrib.ListItems(x).SubItems(1)) * CDbl(lvTrib.ListItems(x).SubItems(2)))
    End If
Next

If txtAbate.Text = "" Or txtAbate.Text = "," Then
   nValorAbate = 0
Else
    nValorAbate = CDbl(txtAbate.Text)
End If
nValorTotal = nTotal - nValorAbate
lblTotal.Caption = FormatNumber(nValorTotal, 2)
lblValorParcela.Caption = FormatNumber(nValorTotal / nQtde, 2)
If Val(txtPercUnica.Text) <> 0 Then
    nPerc = CDbl(txtPercUnica.Text / 100)
    lblTotalUnica.Caption = FormatNumber(nValorTotal - (nValorTotal * nPerc), 2)
Else
    nPerc = 0
    lblTotalUnica.Caption = "0,00"
End If



End Sub

Private Function MesesProporcional() As Integer
Dim sDataAtual As String, sDataFim As String
Dim dData As Date

If Not IsDate(mskDataInicio.Text) Then
   dData = CDate(Format(Now, "dd/mm/yyyy"))
Else
   dData = CDate(mskDataInicio.Text)
End If

sDataAtual = Format(dData, "dd/mm/yyyy")
sDataFim = Format(dData, "31/12/yyyy")

MesesProporcional = DateDiff("m", CDate(sDataAtual), CDate(sDataFim)) + 1

End Function

Private Sub EmiteBoleto()
Dim Sql As String, RdoAux As rdoResultset, nNumDoc As Long, nAno As Integer, nLanc As Integer, nSeqLanc As Integer, nQtdeTributo As Double, nSid As Long
Dim nNumParc As Integer, nCompl As Integer, sDataVencto As String, nValor As Double, p As Integer, nCodigo As Long, nPerc As Double, sUF As String
Dim nUserID As Integer, nQtdeParc As Integer, nCodTributo As Integer, nValorTributo As Double, t As Integer, bUnica As Boolean, sCep As String, sLote As String
Dim nValorParcela As Double, nSeqObs As Integer, sDataBase As String, sNome As String, sCPF As String, sEndereco As String, sCidade As String, sQuadra As String
Dim sBairro As String, sNossoNumero As String, dDataBase As String, nFatorVencto As Integer, sQuintoGrupo As String, sBarra As String, sCampo1 As String
Dim sCampo2 As String, sCampo3 As String, sCampo4 As String, sCampo5 As String, sDigitavel As String, sObs As String, sNumDoc As String, sInscricao As String
Dim v1 As String, v2 As String, v3 As String, v4 As String, v5 As String, v6 As String, v7 As String, v8 As String, v9 As String, V10 As String, aNumDoc() As Long
Dim sTributo As String, x As Integer, nCodigoImovel As Variant, nAreaImovel As Double

ReDim aNumDoc(1)
Ocupado

mskDataVencimento_LostFocus

sObsIss = ""
nSid = Int(Rnd(100) * 1000000)
Sql = "delete from ficha_compensacao where sid=" & nSid
cn.Execute Sql, rdExecDirect

bUnica = IIf(chkUnica.value = vbChecked, True, False)
nCodigo = Val(frmEmissaoGuia.txtCodigo.Text)
nAno = Val(cmbExercicio.Text)
nLanc = cmbLanc.ItemData(cmbLanc.ListIndex)
nCompl = 0
nUserID = RetornaUsuarioID(NomeDeLogin)
nQtdeParc = Val(txtQtdeParc.Text)
nPerc = CDbl(txtPercUnica.Text / 100)
sDataBase = Right$(frmMdi.Sbar.Panels(6).Text, 10)
sNome = frmEmissaoGuia.txtNome.Text
sCPF = RetornaNumero(frmEmissaoGuia.txtDoc.Text)
sEndereco = frmEmissaoGuia.txtEndereco.Text
sBairro = frmEmissaoGuia.txtBairro.Text
sCep = frmEmissaoGuia.txtCep.Text
sCidade = frmEmissaoGuia.txtCidade.Text
sUF = frmEmissaoGuia.txtUF.Text
sInscricao = frmEmissaoGuia.txtInscricao.Text
sQuadra = frmEmissaoGuia.txtQuadra.Text
sLote = frmEmissaoGuia.txtLote.Text

sTributo = ""
For x = 1 To lvTrib.ListItems.Count
    If lvTrib.ListItems(x).Checked Then
        If nLanc = 50 Or nLanc = 65 Then
            sTributo = lvTrib.ListItems(x).Text
            nAreaIss = lvTrib.ListItems(x).SubItems(1)
            Exit For
        End If
    End If
Next
nAreaImovel = nAreaIss

If sTributo <> "" And chkRural.value = vbUnchecked And Not bMulta Then
Inicio:
    nCodigoImovel = InputBox("Digite o código do imóvel a ser lançado o ISS Constução Civil", "Campo obrigatório")
    If Trim(nCodigoImovel) = "" Then GoTo Inicio
 '   Sql = "select codreduzido,dt_areaterreno from cadimob where codreduzido=" & Val(nCodigoImovel)
  '  Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   ' If RdoAux.RowCount > 0 Then
   '    nAreaImovel = RdoAux!Dt_AreaTerreno
'        bFind = True
'    End If
'    RdoAux.Close
'   If Not bFind Then GoTo Inicio
    sObsIss = "Emitido guia de ISS construção civil do tipo -> " & sTributo & " com área de: " & Format(nAreaImovel, "#0.00") & " m² para o imóvel:" & nCodigoImovel & " Lançado por " & NomeDeLogin & " no código cidadão:" & nCodigo
End If

Sql = "SELECT MAX(SEQLANCAMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodigo & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!maximo) Then
    nSeqLanc = 0
Else
    nSeqLanc = RdoAux!maximo + 1
End If
RdoAux.Close


If Trim(sObsIss) <> "" Then
    Sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA WHERE CODREDUZIDO=" & nCodigo & " AND ANOEXERCICIO=" & nAno
    Sql = Sql & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND NUMPARCELA=" & nNumParc & " AND CODCOMPLEMENTO=" & nCompl
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If IsNull(!maximo) Then
            nSeqObs = 1
        Else
            nSeqObs = !maximo + 1
        End If
       .Close
    End With

    Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & nCodigo & ","
    Sql = Sql & nAno & "," & nLanc & "," & nSeqLanc & "," & nNumParc & "," & nCompl & "," & nSeqObs & ",'" & Mask(Trim(sObsIss)) & "',"
    Sql = Sql & RetornaUsuarioID(NomeDeLogin) & ",'" & Format(sDataBase, "mm/dd/yyyy") & "')"
    cn.Execute Sql, rdExecDirect

    Sql = "SELECT MAX(SEQ) AS MAXIMO FROM HISTORICO WHERE CODREDUZIDO=" & nCodigoImovel
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If IsNull(!maximo) Then
            nSeqObs = 1
        Else
            nSeqObs = !maximo + 1
        End If
       .Close
    End With
    
    Sql = "INSERT HISTORICO(CODREDUZIDO,SEQ,DATAHIST,DESCHIST,USERID,DATAHIST2) VALUES("
    Sql = Sql & nCodigoImovel & "," & nSeqObs & ",'" & Format(Now, "dd/mm/yyyy") & "','" & sObsIss & "',236,'" & Format(Now, "mm/dd/yyyy") & "')"
    cn.Execute Sql, rdExecDirect

End If


p = IIf(bUnica, 0, 1)
For nNumParc = p To nQtdeParc
    If nNumParc = 0 Or nNumParc = 1 Then
        sDataVencto = mskDataVencimento.Text
    Else
        sDataVencto = grdData.TextMatrix(nNumParc, 1)
    End If
    
    'GRAVA PARCELA
    Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,USERID) VALUES("
    Sql = Sql & nCodigo & "," & nAno & "," & nLanc & "," & nSeqLanc & "," & nNumParc & "," & nCompl & "," & 3 & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','"
    Sql = Sql & Format(sDataBase, "mm/dd/yyyy") & "'," & nUserID & ")"
    cn.Execute Sql, rdExecDirect
    
    nValorParcela = 0
    For t = 1 To lvTrib.ListItems.Count
        If lvTrib.ListItems(t).Checked = True Then
            nCodTributo = Val(Right$(lvTrib.ListItems(t).Key, 3))
            nQtdeTributo = CDbl(lvTrib.ListItems(t).SubItems(1))
            If nNumParc = 0 Then
                nValorTributo = CDbl(lvTrib.ListItems(t).SubItems(2)) * nQtdeTributo
                nValorTributo = FormatNumber(nValorTributo - (nValorTributo * nPerc), 2)
            Else
                nValorTributo = CDbl(lvTrib.ListItems(t).SubItems(2)) * nQtdeTributo / nQtdeParc
            End If
            nValorParcela = nValorParcela + nValorTributo
            
            'GRAVA TRIBUTOS
            Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
            Sql = Sql & nCodigo & "," & nAno & "," & nLanc & "," & nSeqLanc & "," & nNumParc & "," & nCompl & "," & nCodTributo & "," & Virg2Ponto(CStr(nValorTributo)) & ")"
            cn.Execute Sql, rdExecDirect
        End If
    Next

    Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    nNumDoc = RdoAux!maximo + 1
    RdoAux.Close
    aNumDoc(nNumParc) = nNumDoc
    ReDim Preserve aNumDoc(UBound(aNumDoc) + 1)
    
    'GRAVA Nº DOCUMENTO
    Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,VALORGUIA,TIPODOC,USERID,REGISTRADO) VALUES("
    Sql = Sql & nNumDoc & ",'" & Format(Now, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(nValorParcela)) & "," & 3 & "," & nUserID & "," & 1 & ")"
    cn.Execute Sql, rdExecDirect

    'GRAVA PARCELA DOCUMENTO
    Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & nCodigo & ","
    Sql = Sql & nAno & "," & nLanc & "," & nSeqLanc & "," & nNumParc & "," & nCompl & "," & nNumDoc & ")"
    cn.Execute Sql, rdExecDirect
    
    'GRAVA OBS PARCELA
    If Trim(txtObs.Text) <> "" Then
        Sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA WHERE CODREDUZIDO=" & nCodigo & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND NUMPARCELA=" & nNumParc & " AND CODCOMPLEMENTO=" & nCompl
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If IsNull(!maximo) Then
                nSeqObs = 1
            Else
                nSeqObs = !maximo + 1
            End If
           .Close
        End With
        Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & nCodigo & ","
        Sql = Sql & nAno & "," & nLanc & "," & nSeqLanc & "," & nNumParc & "," & nCompl & "," & nSeqObs & ",'" & Mask(Trim(txtObs.Text)) & "',"
        Sql = Sql & nUserID & ",'" & Format(sDataBase, "mm/dd/yyyy") & "')"
        cn.Execute Sql, rdExecDirect
    End If
    
    'GRAVA DOCUMENTO PARA REGISTRO
    Sql = "INSERT ficha_compensacao_documento(numero_documento,data_vencimento,valor_documento,nome,cpf,endereco,bairro,cep,cidade,uf) values(" & nNumDoc & ",'"
    Sql = Sql & Format(sDataVencto, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(nValorParcela)) & ",'" & Mask(Left(sNome, 40)) & "','" & sCPF & "','"
    Sql = Sql & Mask(Left(sEndereco, 40)) & "','" & Mask(Left(sBairro, 15)) & "','" & RetornaNumero(sCep) & "','" & Mask(Left(sCidade, 30)) & "','" & sUF & "')"
    cn.Execute Sql, rdExecDirect
    
Next


'IMPRIMIR PARCELA ÚNICA E PARCELA 1 PELO SITE DO BB

v1 = sNome
v2 = Left(sEndereco & " - " & sBairro, 60)
v3 = mskDataVencimento.Text
v4 = RetornaNumero(sCPF)
v5 = "287353200" & Format(aNumDoc(0), "00000000")
v6 = RetornaNumero(lblTotalUnica.Caption)
v7 = Left(sCidade, 18)
v8 = sUF
v9 = sCep
V10 = NomeDeLogin
If Trim(sCep) = "" Or Trim(sCep) = "-" Then
    v9 = "14870-000"
End If
If bUnica Then
    'ShellExecute HWND, "open", "http://sistemas.jaboticabal.sp.gov.br/gti/Pages/boletoBB.aspx?f1=" & v1 & "&f2=" & v2 & "&f3=" & v3 & "&f4=" & v4 & "&f5=" & v5 & "&f6=" & v6 & "&f7=" & v7 & "&f8=" & v8 & "&f9=" & v9 & "&f10=" & V10, vbNullString, vbNullString, conSwNormal
    ShellExecute HWND, "open", "http://sistemas.jaboticabal.sp.gov.br/gti/Tributario/GateBank?p1=" & v1 & "&p2=" & v2 & "&p3=" & v3 & "&p4=" & v4 & "&p5=" & v5 & "&p6=" & v6 & "&p7=" & v7 & "&p8=" & v8 & "&p9=" & v9, vbNullString, vbNullString, conSwNormal
Else
    v5 = "287353200" & Format(aNumDoc(1), "00000000")
    v6 = RetornaNumero(lblValorParcela.Caption)
    'ShellExecute HWND, "open", "http://sistemas.jaboticabal.sp.gov.br/gti/Pages/boletoBB.aspx?f1=" & v1 & "&f2=" & v2 & "&f3=" & v3 & "&f4=" & v4 & "&f5=" & v5 & "&f6=" & v6 & "&f7=" & v7 & "&f8=" & v8 & "&f9=" & v9 & "&f10=" & V10, vbNullString, vbNullString, conSwNormal
    ShellExecute HWND, "open", "http://sistemas.jaboticabal.sp.gov.br/gti/Tributario/GateBank?p1=" & v1 & "&p2=" & v2 & "&p3=" & v3 & "&p4=" & v4 & "&p5=" & v5 & "&p6=" & v6 & "&p7=" & v7 & "&p8=" & v8 & "&p9=" & v9, vbNullString, vbNullString, conSwNormal
End If

'IMPRIMIR O RESTANTE DOS BOLETOS
If nQtdeParc > 1 Then
    'mskDataVencimento_LostFocus
    For nNumParc = 2 To nQtdeParc
        sNossoNumero = "2873532"
        dDataBase = "07/10/1997"
        nFatorVencto = CDate(sDataVencto) - CDate(dDataBase)
        sQuintoGrupo = Format(nFatorVencto, "0000")
        sQuintoGrupo = sQuintoGrupo & Format(RetornaNumero(FormatNumber(nValorParcela, 2)), "0000000000")
        sBarra = "0019" & Format(nFatorVencto, "0000") & Format(RetornaNumero(FormatNumber(nValorParcela, 2)), "0000000000") & "000000287353200"
        sBarra = sBarra & CStr(aNumDoc(nNumParc)) & "17"
        
        sCampo1 = "0019" & Mid(sBarra, 20, 5)
        sDigitavel = sCampo1 & Val(Calculo_DV10(sCampo1))
        sCampo2 = Mid(sBarra, 24, 10)
        sDigitavel = sDigitavel & sCampo2 & Val(Calculo_DV10(sCampo2))
        sCampo3 = Mid(sBarra, 34, 10)
        sDigitavel = sDigitavel & sCampo3 & Val(Calculo_DV10(sCampo3))
        sCampo5 = Format(nFatorVencto, "0000") & Format(RetornaNumero(FormatNumber(nValorParcela, 2)), "0000000000")
        sCampo4 = Val(Calculo_DV11(sBarra))
        sDigitavel = sDigitavel & sCampo4 & sCampo5
        sBarra = Left(sBarra, 4) & sCampo4 & Mid(sBarra, 5, Len(sBarra) - 4)
        sDigitavel2 = Left(sDigitavel, 5) & "." & Mid(sDigitavel, 6, 5) & " " & Mid(sDigitavel, 11, 5) & "." & Mid(sDigitavel, 16, 6) & " "
        sDigitavel2 = sDigitavel2 & Mid(sDigitavel, 22, 5) & "." & Mid(sDigitavel, 27, 6) & " " & Mid(sDigitavel, 33, 1) & " " & Right(sDigitavel, 14)
        sBarra = Gera2of5Str(sBarra)
        
        sObs = "Guia(s) de pagamento referente à " & cmbLanc.Text
        sNumDoc = "287353200" & Format(nNumDoc, "00000000")
        Sql = "Insert FICHA_COMPENSACAO(SID,SEQ,CODIGO,NOME,CPF,ENDERECO,BAIRRO,CIDADE,CEP,DOCUMENTO,VALOR,VENCIMENTO,PARCELA,DIGITAVEL,CODBARRA,OBS,INSCRICAO,QUADRA,LOTE,UF) VALUES("
        Sql = Sql & nSid & "," & nNumParc & "," & nCodigo & ",'" & Left(Mask(sNome), 80) & "','" & sCPF & "','" & Mask(sEndereco) & "','" & Left(Mask(sBairro), 25) & "','"
        Sql = Sql & Mask(sCidade) & "','" & sCep & "'," & aNumDoc(nNumParc) & "," & Virg2Ponto(Format(nValorParcela, "#0.00")) & ",'" & Format(grdData.TextMatrix(nNumParc, 1), "mm/dd/yyyy") & "','"
        Sql = Sql & Format(nNumParc, "00") & "/" & Format(nQtdeParc, "00") & "','" & sDigitavel2 & "','" & Mask(sBarra) & "','" & sObs & "','" & sInscricao & "','"
        Sql = Sql & Mask(sQuadra) & "','" & Mask(sLote) & "','" & sUF & "')"
        cn.Execute Sql, rdExecDirect
    Next
    
    frmReport.ShowReport3 "FICHACOMPENSACAO", frmMdi.HWND, Me.HWND, nSid
    
    Sql = "delete from ficha_compensacao where sid=" & nSid
    cn.Execute Sql, rdExecDirect
    
End If

Liberado

Unload Me

End Sub

Private Sub EmiteBoletoRocada()
Dim Sql As String, RdoAux As rdoResultset, nNumDoc As Long, nAno As Integer, nLanc As Integer, nSeqLancRocada As Integer, nSeqLancMulta As Integer, nQtdeTributo As Double, nSid As Long
Dim nNumParc As Integer, nCompl As Integer, sDataVencto As String, nValor As Double, p As Integer, nCodigo As Long, nPerc As Double, sUF As String
Dim nUserID As Integer, nQtdeParc As Integer, nCodTributo As Integer, nValorTributo As Double, t As Integer, bUnica As Boolean, sCep As String, sLote As String
Dim nValorParcela As Double, nSeqObsRocada As Integer, nSeqObsMulta As Integer, sDataBase As String, sNome As String, sCPF As String, sEndereco As String, sCidade As String, sQuadra As String
Dim sBairro As String, sNossoNumero As String, dDataBase As String, nFatorVencto As Integer, sQuintoGrupo As String, sBarra As String, sCampo1 As String
Dim sCampo2 As String, sCampo3 As String, sCampo4 As String, sCampo5 As String, sDigitavel As String, sObs As String, sNumDoc As String, sInscricao As String
Dim v1 As String, v2 As String, v3 As String, v4 As String, v5 As String, v6 As String, v7 As String, v8 As String, v9 As String, V10 As String, aNumDoc() As Long
Dim sTributo As String, x As Integer, nCodigoImovel As Variant

ReDim aNumDoc(1)
Ocupado
'Sql = "delete from ficha_compensacao where sid=" & nSid
'cn.Execute Sql, rdExecDirect

mskDataVencimento_LostFocus

sObsIss = ""
'nSid = Int(Rnd(100) * 1000000)
bUnica = IIf(chkUnica.value = vbChecked, True, False)
nCodigo = Val(frmEmissaoGuia.txtCodigo.Text)
nAno = Val(cmbExercicio.Text)
nLanc = cmbLanc.ItemData(cmbLanc.ListIndex)
nCompl = 0
nUserID = RetornaUsuarioID(NomeDeLogin)
nQtdeParc = Val(txtQtdeParc.Text)
nPerc = CDbl(txtPercUnica.Text / 100)
sDataBase = Right$(frmMdi.Sbar.Panels(6).Text, 10)
sNome = frmEmissaoGuia.txtNome.Text
sCPF = RetornaNumero(frmEmissaoGuia.txtDoc.Text)
sEndereco = frmEmissaoGuia.txtEndereco.Text
sBairro = frmEmissaoGuia.txtBairro.Text
sCep = frmEmissaoGuia.txtCep.Text
sCidade = frmEmissaoGuia.txtCidade.Text
sUF = frmEmissaoGuia.txtUF.Text
sInscricao = frmEmissaoGuia.txtInscricao.Text
sQuadra = frmEmissaoGuia.txtQuadra.Text
sLote = frmEmissaoGuia.txtLote.Text

Sql = "SELECT MAX(SEQLANCAMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodigo & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=16"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!maximo) Then
    nSeqLancRocada = 0
Else
    nSeqLancRocada = RdoAux!maximo + 1
End If
RdoAux.Close

Sql = "SELECT MAX(SEQLANCAMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodigo & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=16"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!maximo) Then
    nSeqLancMulta = 0
Else
    nSeqLancMulta = RdoAux!maximo + 1
End If
RdoAux.Close

p = IIf(bUnica, 0, 1)
nNumParc = 1
sDataVencto = grdData.TextMatrix(nNumParc, 1)

'GRAVA PARCELA
Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,USERID) VALUES("
Sql = Sql & nCodigo & "," & nAno & "," & 16 & "," & nSeqLancRocada & "," & nNumParc & "," & nCompl & "," & 3 & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','"
Sql = Sql & Format(sDataBase, "mm/dd/yyyy") & "'," & nUserID & ")"
cn.Execute Sql, rdExecDirect

nValorTributo = 636.53
nValorParcela = nValorTributo
Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
Sql = Sql & nCodigo & "," & nAno & "," & 16 & "," & nSeqLancRocada & "," & nNumParc & "," & nCompl & "," & 20 & "," & Virg2Ponto(CStr(nValorTributo)) & ")"
cn.Execute Sql, rdExecDirect

Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nNumDoc = RdoAux!maximo + 1
RdoAux.Close

'##################################
'########## GERA MULTA ###########
'##################################

'GRAVA Nº DOCUMENTO
Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,VALORGUIA,TIPODOC,USERID,REGISTRADO) VALUES("
Sql = Sql & nNumDoc & ",'" & Format(Now, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(nValorParcela)) & "," & 3 & "," & nUserID & "," & 1 & ")"
cn.Execute Sql, rdExecDirect

'GRAVA PARCELA DOCUMENTO
Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & nCodigo & ","
Sql = Sql & nAno & "," & 16 & "," & nSeqLancRocada & "," & nNumParc & "," & nCompl & "," & nNumDoc & ")"
cn.Execute Sql, rdExecDirect

'GRAVA OBS PARCELA
If Trim(txtObs.Text) <> "" Then
    Sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA WHERE CODREDUZIDO=" & nCodigo & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=16 AND SEQLANCAMENTO=" & nSeqLancRocada & " AND NUMPARCELA=" & nNumParc & " AND CODCOMPLEMENTO=" & nCompl
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If IsNull(!maximo) Then
            nSeqObs = 1
        Else
            nSeqObs = !maximo + 1
        End If
       .Close
    End With
    Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & nCodigo & ","
    Sql = Sql & nAno & "," & 16 & "," & nSeqLancRocada & "," & nNumParc & "," & nCompl & "," & nSeqObs & ",'" & Mask(Trim(txtObs.Text)) & "',"
    Sql = Sql & nUserID & ",'" & Format(sDataBase, "mm/dd/yyyy") & "')"
    cn.Execute Sql, rdExecDirect
End If

Sql = "SELECT MAX(SEQ) AS MAXIMO FROM HISTORICO WHERE CODREDUZIDO=" & nCodigo
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If IsNull(!maximo) Then
        nSeqObs = 1
    Else
        nSeqObs = !maximo + 1
    End If
   .Close
End With

Sql = "INSERT HISTORICO(CODREDUZIDO,SEQ,DATAHIST,DESCHIST,USERID,DATAHIST2) VALUES("
Sql = Sql & nCodigo & "," & nSeqObs & ",'" & Format(Now, "dd/mm/yyyy") & "','MULTA DE INFRAÇÃO(ROÇADA) LANÇADA',236,'" & Format(Now, "mm/dd/yyyy") & "')"
cn.Execute Sql, rdExecDirect

'GRAVA DOCUMENTO PARA REGISTRO
Sql = "INSERT ficha_compensacao_documento(numero_documento,data_vencimento,valor_documento,nome,cpf,endereco,bairro,cep,cidade,uf) values(" & nNumDoc & ",'"
Sql = Sql & Format(sDataVencto, "mm/dd/yyyy") & "'," & Virg2Ponto(Format(CStr(nValorParcela / 2), "#0.00")) & ",'" & Mask(Left(sNome, 40)) & "','" & sCPF & "','"
Sql = Sql & Mask(Left(sEndereco, 40)) & "','" & Mask(Left(sBairro, 15)) & "','" & RetornaNumero(sCep) & "','" & Mask(Left(sCidade, 30)) & "','" & sUF & "')"
cn.Execute Sql, rdExecDirect

nValorParcela = 318.27 '50% de desconto na multa
sNossoNumero = "2873532"
dDataBase = "07/10/1997"
nFatorVencto = CDate(sDataVencto) - CDate(dDataBase)
sQuintoGrupo = Format(nFatorVencto, "0000")
sQuintoGrupo = sQuintoGrupo & Format(RetornaNumero(FormatNumber(nValorParcela, 2)), "0000000000")
sBarra = "0019" & Format(nFatorVencto, "0000") & Format(RetornaNumero(FormatNumber(nValorParcela, 2)), "0000000000") & "000000287353200"
sBarra = sBarra & CStr(nNumDoc) & "17"

sCampo1 = "0019" & Mid(sBarra, 20, 5)
sDigitavel = sCampo1 & Val(Calculo_DV10(sCampo1))
sCampo2 = Mid(sBarra, 24, 10)
sDigitavel = sDigitavel & sCampo2 & Val(Calculo_DV10(sCampo2))
sCampo3 = Mid(sBarra, 34, 10)
sDigitavel = sDigitavel & sCampo3 & Val(Calculo_DV10(sCampo3))
sCampo5 = Format(nFatorVencto, "0000") & Format(RetornaNumero(FormatNumber(nValorParcela, 2)), "0000000000")
sCampo4 = Val(Calculo_DV11(sBarra))
sDigitavel = sDigitavel & sCampo4 & sCampo5
sBarra = Left(sBarra, 4) & sCampo4 & Mid(sBarra, 5, Len(sBarra) - 4)
sDigitavel2 = Left(sDigitavel, 5) & "." & Mid(sDigitavel, 6, 5) & " " & Mid(sDigitavel, 11, 5) & "." & Mid(sDigitavel, 16, 6) & " "
sDigitavel2 = sDigitavel2 & Mid(sDigitavel, 22, 5) & "." & Mid(sDigitavel, 27, 6) & " " & Mid(sDigitavel, 33, 1) & " " & Right(sDigitavel, 14)
sBarra = Gera2of5Str(sBarra)

nSid = Int(Rnd(100) * 1000000)
Sql = "delete from ficha_compensacao where sid=" & nSid
cn.Execute Sql, rdExecDirect

sObs = "Guia(s) de pagamento referente à multa de infração."
sNumDoc = "287353200" & Format(nNumDoc, "00000000")
Sql = "Insert FICHA_COMPENSACAO(SID,SEQ,CODIGO,NOME,CPF,ENDERECO,BAIRRO,CIDADE,CEP,DOCUMENTO,VALOR,VENCIMENTO,PARCELA,DIGITAVEL,CODBARRA,OBS,INSCRICAO,QUADRA,LOTE,UF) VALUES("
Sql = Sql & nSid & "," & nNumParc & "," & nCodigo & ",'" & Left(Mask(sNome), 80) & "','" & sCPF & "','" & sEndereco & "','" & Left(Mask(sBairro), 25) & "','"
Sql = Sql & Mask(sCidade) & "','" & sCep & "'," & sNumDoc & "," & Virg2Ponto(Format(nValorParcela, "#0.00")) & ",'" & Format(mskDataVencimento.Text, "mm/dd/yyyy") & "','"
Sql = Sql & "01/01','" & sDigitavel2 & "','" & Mask(sBarra) & "','" & sObs & "','" & sInscricao & "','"
Sql = Sql & Mask(sQuadra) & "','" & Mask(sLote) & "','" & sUF & "')"
cn.Execute Sql, rdExecDirect

'##################################
'########## GERA ROÇADA ###########
'##################################

Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nNumDoc = RdoAux!maximo + 1
RdoAux.Close

Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,USERID) VALUES("
Sql = Sql & nCodigo & "," & nAno & "," & 38 & "," & nSeqLancMulta & "," & nNumParc & "," & nCompl & "," & 3 & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','"
Sql = Sql & Format(sDataBase, "mm/dd/yyyy") & "'," & nUserID & ")"
cn.Execute Sql, rdExecDirect

nValorTributo = nAreaTerreno * 1.0329
nValorParcela = nValorTributo
Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
Sql = Sql & nCodigo & "," & nAno & "," & 38 & "," & nSeqLancMulta & "," & nNumParc & "," & nCompl & "," & 170 & "," & Virg2Ponto(CStr(nValorTributo)) & ")"
cn.Execute Sql, rdExecDirect

'GRAVA Nº DOCUMENTO
Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,VALORGUIA,TIPODOC,USERID,REGISTRADO) VALUES("
Sql = Sql & nNumDoc & ",'" & Format(Now, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(nValorParcela)) & "," & 3 & "," & nUserID & "," & 1 & ")"
cn.Execute Sql, rdExecDirect

'GRAVA PARCELA DOCUMENTO
Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & nCodigo & ","
Sql = Sql & nAno & "," & 38 & "," & nSeqLancMulta & "," & nNumParc & "," & nCompl & "," & nNumDoc & ")"
cn.Execute Sql, rdExecDirect

'GRAVA OBS PARCELA
If Trim(txtObs.Text) <> "" Then
    Sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA WHERE CODREDUZIDO=" & nCodigo & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=38 AND SEQLANCAMENTO=" & nSeqLancMulta & " AND NUMPARCELA=" & nNumParc & " AND CODCOMPLEMENTO=" & nCompl
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If IsNull(!maximo) Then
            nSeqObs = 1
        Else
            nSeqObs = !maximo + 1
        End If
       .Close
    End With
    Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & nCodigo & ","
    Sql = Sql & nAno & "," & 38 & "," & nSeqLancMulta & "," & nNumParc & "," & nCompl & "," & nSeqObs & ",'" & Mask(Trim(txtObs.Text)) & "',"
    Sql = Sql & nUserID & ",'" & Format(sDataBase, "mm/dd/yyyy") & "')"
    cn.Execute Sql, rdExecDirect
End If

'GRAVA DOCUMENTO PARA REGISTRO
Sql = "INSERT ficha_compensacao_documento(numero_documento,data_vencimento,valor_documento,nome,cpf,endereco,bairro,cep,cidade,uf) values(" & nNumDoc & ",'"
Sql = Sql & Format(sDataVencto, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(nValorParcela)) & ",'" & Mask(Left(sNome, 40)) & "','" & sCPF & "','"
Sql = Sql & Mask(Left(sEndereco, 40)) & "','" & Mask(Left(sBairro, 15)) & "','" & RetornaNumero(sCep) & "','" & Mask(Left(sCidade, 30)) & "','" & sUF & "')"
cn.Execute Sql, rdExecDirect

sNossoNumero = "2873532"
dDataBase = "07/10/1997"
nFatorVencto = CDate(sDataVencto) - CDate(dDataBase)
sQuintoGrupo = Format(nFatorVencto, "0000")
sQuintoGrupo = sQuintoGrupo & Format(RetornaNumero(FormatNumber(nValorParcela, 2)), "0000000000")
sBarra = "0019" & Format(nFatorVencto, "0000") & Format(RetornaNumero(FormatNumber(nValorParcela, 2)), "0000000000") & "000000287353200"
sBarra = sBarra & CStr(nNumDoc) & "17"

sCampo1 = "0019" & Mid(sBarra, 20, 5)
sDigitavel = sCampo1 & Val(Calculo_DV10(sCampo1))
sCampo2 = Mid(sBarra, 24, 10)
sDigitavel = sDigitavel & sCampo2 & Val(Calculo_DV10(sCampo2))
sCampo3 = Mid(sBarra, 34, 10)
sDigitavel = sDigitavel & sCampo3 & Val(Calculo_DV10(sCampo3))
sCampo5 = Format(nFatorVencto, "0000") & Format(RetornaNumero(FormatNumber(nValorParcela, 2)), "0000000000")
sCampo4 = Val(Calculo_DV11(sBarra))
sDigitavel = sDigitavel & sCampo4 & sCampo5
sBarra = Left(sBarra, 4) & sCampo4 & Mid(sBarra, 5, Len(sBarra) - 4)
sDigitavel2 = Left(sDigitavel, 5) & "." & Mid(sDigitavel, 6, 5) & " " & Mid(sDigitavel, 11, 5) & "." & Mid(sDigitavel, 16, 6) & " "
sDigitavel2 = sDigitavel2 & Mid(sDigitavel, 22, 5) & "." & Mid(sDigitavel, 27, 6) & " " & Mid(sDigitavel, 33, 1) & " " & Right(sDigitavel, 14)
sBarra = Gera2of5Str(sBarra)


sObs = "Guia(s) de pagamento referente à Taxa de limpeza de terreno."
sNumDoc = "287353200" & Format(nNumDoc, "00000000")
Sql = "Insert FICHA_COMPENSACAO(SID,SEQ,CODIGO,NOME,CPF,ENDERECO,BAIRRO,CIDADE,CEP,DOCUMENTO,VALOR,VENCIMENTO,PARCELA,DIGITAVEL,CODBARRA,OBS,INSCRICAO,QUADRA,LOTE,UF) VALUES("
Sql = Sql & nSid & "," & 2 & "," & nCodigo & ",'" & Left(Mask(sNome), 80) & "','" & sCPF & "','" & sEndereco & "','" & Left(Mask(sBairro), 25) & "','"
Sql = Sql & Mask(sCidade) & "','" & sCep & "'," & sNumDoc & "," & Virg2Ponto(Format(nValorParcela, "#0.00")) & ",'" & Format(mskDataVencimento.Text, "mm/dd/yyyy") & "','"
Sql = Sql & "01/01','" & sDigitavel2 & "','" & Mask(sBarra) & "','" & sObs & "','" & sInscricao & "','"
Sql = Sql & Mask(sQuadra) & "','" & Mask(sLote) & "','" & sUF & "')"
cn.Execute Sql, rdExecDirect

frmReport.ShowReport3 "FICHACOMPENSACAO_Rocada", frmMdi.HWND, Me.HWND, nSid

Sql = "delete from ficha_compensacao where sid=" & nSid
cn.Execute Sql, rdExecDirect
  

Liberado



End Sub

