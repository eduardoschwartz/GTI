VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmCidadao 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Cidad�o"
   ClientHeight    =   7545
   ClientLeft      =   9810
   ClientTop       =   4170
   ClientWidth     =   7590
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   Begin Tributacao.jcFrames jcFrames3 
      Height          =   1890
      Left            =   60
      Top             =   60
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   3334
      FrameColor      =   12829635
      Style           =   0
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
      Begin VB.TextBox txtProfissao 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   3210
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   79
         Top             =   1450
         Width           =   3690
      End
      Begin VB.TextBox txtCod 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   105
         Width           =   975
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   4
         Top             =   420
         Width           =   6300
      End
      Begin VB.TextBox txtOrgao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4605
         MaxLength       =   25
         TabIndex        =   6
         Top             =   765
         Width           =   2780
      End
      Begin VB.OptionButton optP 
         Caption         =   "Pessoa F�sica"
         Height          =   285
         Index           =   0
         Left            =   4590
         TabIndex        =   2
         Top             =   60
         Value           =   -1  'True
         Width           =   1365
      End
      Begin VB.OptionButton optP 
         Caption         =   "Pessoa Jur�dica"
         Height          =   285
         Index           =   1
         Left            =   5970
         TabIndex        =   3
         Top             =   60
         Width           =   1455
      End
      Begin esMaskEdit.esMaskedEdit mskCPF 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   1110
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   503
         MouseIcon       =   "frmCidadao.frx":0000
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
         MaxLength       =   14
         Mask            =   "999.999.999-99"
         SelText         =   ""
         Text            =   "___.___.___-__"
         HideSelection   =   -1  'True
      End
      Begin esMaskEdit.esMaskedEdit mskCNPJ 
         Height          =   285
         Left            =   4605
         TabIndex        =   8
         Top             =   1110
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   503
         MouseIcon       =   "frmCidadao.frx":001C
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
         MaxLength       =   18
         Mask            =   "99.999.999/9999-99"
         SelText         =   ""
         Text            =   "__.___.___/____-__"
         HideSelection   =   -1  'True
      End
      Begin esMaskEdit.esMaskedEdit mskRG 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   765
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   503
         MouseIcon       =   "frmCidadao.frx":0038
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
         MaxLength       =   25
         SelText         =   ""
         HideSelection   =   -1  'True
      End
      Begin prjChameleon.chameleonButton cmbFindCod 
         Height          =   270
         Left            =   2280
         TabIndex        =   0
         ToolTipText     =   "Consulta por c�digo"
         Top             =   60
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   476
         BTYPE           =   9
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCidadao.frx":0054
         PICN            =   "frmCidadao.frx":0070
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton btCPF 
         Height          =   270
         Left            =   3465
         TabIndex        =   71
         ToolTipText     =   "Situa��o Cadastral"
         Top             =   1125
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   476
         BTYPE           =   9
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCidadao.frx":01CA
         PICN            =   "frmCidadao.frx":01E6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton btCNPJ 
         Height          =   270
         Left            =   6930
         TabIndex        =   72
         ToolTipText     =   "Situa��o Cadastral"
         Top             =   1125
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   476
         BTYPE           =   9
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCidadao.frx":0340
         PICN            =   "frmCidadao.frx":035C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin esMaskEdit.esMaskedEdit mskDtNascto 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   1470
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   503
         MouseIcon       =   "frmCidadao.frx":04B6
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
      Begin prjChameleon.chameleonButton btProfissao 
         Height          =   270
         Left            =   6930
         TabIndex        =   78
         ToolTipText     =   "Selecionar profiss�o"
         Top             =   1470
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   476
         BTYPE           =   9
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCidadao.frx":04D2
         PICN            =   "frmCidadao.frx":04EE
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
         BackStyle       =   0  'Transparent
         Caption         =   "Profiss�o..:"
         Height          =   225
         Index           =   27
         Left            =   2340
         TabIndex        =   75
         Top             =   1485
         Width           =   840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Dta Nascto..:"
         Height          =   225
         Index           =   26
         Left            =   90
         TabIndex        =   74
         Top             =   1485
         Width           =   1020
      End
      Begin VB.Label lblEspolio 
         Alignment       =   2  'Center
         Caption         =   "Esp�lio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   2700
         TabIndex        =   73
         Top             =   90
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "C�digo.........:"
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   57
         Top             =   90
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome..........:"
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   56
         Top             =   435
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CNPJ..:"
         Height          =   225
         Index           =   11
         Left            =   4020
         TabIndex        =   55
         Top             =   1125
         Width           =   660
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CPF.............:"
         Height          =   225
         Index           =   12
         Left            =   90
         TabIndex        =   54
         Top             =   1125
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "RG..............:"
         Height          =   225
         Index           =   13
         Left            =   90
         TabIndex        =   53
         Top             =   780
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Org�o..:"
         Height          =   225
         Index           =   14
         Left            =   4020
         TabIndex        =   52
         Top             =   795
         Width           =   630
      End
   End
   Begin Tributacao.jcFrames jcFrames2 
      Height          =   2370
      Index           =   0
      Left            =   60
      Top             =   1965
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   4180
      FrameColor      =   12829635
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "  Endere�o Residencial"
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
      Begin VB.CheckBox chkWhatsApp 
         Caption         =   "WhatsApp"
         Height          =   225
         Left            =   2070
         TabIndex        =   84
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CheckBox chkFone1 
         Caption         =   "N�o possui telefone"
         Height          =   255
         Left            =   90
         TabIndex        =   82
         Top             =   2010
         Width           =   1725
      End
      Begin prjChameleon.chameleonButton btPais 
         Height          =   270
         Left            =   3330
         TabIndex        =   80
         ToolTipText     =   "Selecionar pa�s"
         Top             =   1710
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   476
         BTYPE           =   9
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCidadao.frx":0648
         PICN            =   "frmCidadao.frx":0664
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtPais 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1050
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   61
         Top             =   1680
         Width           =   2235
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4425
         MaxLength       =   50
         TabIndex        =   60
         Top             =   1665
         Width           =   3000
      End
      Begin VB.TextBox txtFone 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4425
         MaxLength       =   30
         TabIndex        =   59
         Top             =   2025
         Width           =   2985
      End
      Begin VB.CheckBox chkEtiq 
         Alignment       =   1  'Right Justify
         Height          =   225
         Left            =   135
         TabIndex        =   26
         Top             =   0
         Width           =   210
      End
      Begin VB.TextBox txtNomeLog 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2205
         MaxLength       =   50
         TabIndex        =   18
         Top             =   240
         Width           =   5220
      End
      Begin VB.TextBox txtNumLog 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   17
         Top             =   240
         Width           =   930
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   19
         Top             =   585
         Width           =   930
      End
      Begin VB.TextBox txtCompl 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2970
         MaxLength       =   50
         TabIndex        =   20
         Top             =   585
         Width           =   4455
      End
      Begin VB.ComboBox cmbBairro 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1305
         Width           =   2250
      End
      Begin VB.ComboBox cmbCidade 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4425
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   930
         Width           =   3000
      End
      Begin VB.ComboBox cmbUF 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmCidadao.frx":07BE
         Left            =   1080
         List            =   "frmCidadao.frx":07C0
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   915
         Width           =   2250
      End
      Begin esMaskEdit.esMaskedEdit mskCEP 
         Height          =   285
         Left            =   4425
         TabIndex        =   25
         Top             =   1320
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         MouseIcon       =   "frmCidadao.frx":07C2
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
         MaxLength       =   9
         Mask            =   "99999-999"
         SelText         =   ""
         Text            =   "_____-___"
         HideSelection   =   -1  'True
      End
      Begin prjChameleon.chameleonButton cmdAddBairro 
         Height          =   270
         Left            =   3345
         TabIndex        =   24
         ToolTipText     =   "Cadastrar um novo bairro"
         Top             =   1320
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   476
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
         MPTR            =   99
         MICON           =   "frmCidadao.frx":07DE
         PICN            =   "frmCidadao.frx":0AF8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Image imgWhatsApp 
         Height          =   255
         Left            =   3180
         Picture         =   "frmCidadao.frx":0C52
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pa�s.............:"
         Height          =   180
         Index           =   15
         Left            =   90
         TabIndex        =   64
         Top             =   1710
         Width           =   990
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Email...:"
         Height          =   225
         Index           =   10
         Left            =   3825
         TabIndex        =   63
         Top             =   1725
         Width           =   630
      End
      Begin VB.Label lblFone 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone.:"
         Height          =   225
         Left            =   3660
         TabIndex        =   62
         Top             =   2055
         Width           =   795
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "N�mero........:"
         Height          =   225
         Index           =   2
         Left            =   90
         TabIndex        =   43
         Top             =   615
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Logradouro..:"
         Height          =   225
         Index           =   3
         Left            =   90
         TabIndex        =   42
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Complem.:"
         Height          =   225
         Index           =   4
         Left            =   2175
         TabIndex        =   41
         Top             =   645
         Width           =   765
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro...........:"
         Height          =   225
         Index           =   5
         Left            =   90
         TabIndex        =   40
         Top             =   1365
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         Height          =   225
         Index           =   6
         Left            =   3810
         TabIndex        =   39
         Top             =   990
         Width           =   585
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "UF...............:"
         Height          =   225
         Index           =   7
         Left            =   90
         TabIndex        =   38
         Top             =   990
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CEP....:"
         Height          =   225
         Index           =   8
         Left            =   3810
         TabIndex        =   37
         Top             =   1365
         Width           =   585
      End
   End
   Begin prjChameleon.chameleonButton cmdCopy 
      Height          =   315
      Left            =   4515
      TabIndex        =   16
      ToolTipText     =   "Copiar dados do cidad�o selecioando"
      Top             =   7170
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Co&piar"
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
      MICON           =   "frmCidadao.frx":1032
      PICN            =   "frmCidadao.frx":104E
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
      Left            =   5340
      TabIndex        =   15
      ToolTipText     =   "Gravar os Dados"
      Top             =   7170
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
      MICON           =   "frmCidadao.frx":1139
      PICN            =   "frmCidadao.frx":1155
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
      Left            =   3420
      TabIndex        =   1
      ToolTipText     =   "Consulta Cidad�os Cadastrados"
      Top             =   7170
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
      MICON           =   "frmCidadao.frx":14FA
      PICN            =   "frmCidadao.frx":1516
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   6420
      TabIndex        =   10
      ToolTipText     =   "Sair da Tela"
      Top             =   7170
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
      MICON           =   "frmCidadao.frx":1670
      PICN            =   "frmCidadao.frx":168C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   6420
      TabIndex        =   11
      ToolTipText     =   "Cancelar Edi��o"
      Top             =   7170
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
      MICON           =   "frmCidadao.frx":16FA
      PICN            =   "frmCidadao.frx":1716
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
      Left            =   2325
      TabIndex        =   12
      ToolTipText     =   "Excluir Registro"
      Top             =   7170
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
      MICON           =   "frmCidadao.frx":1870
      PICN            =   "frmCidadao.frx":188C
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
      Left            =   1230
      TabIndex        =   13
      ToolTipText     =   "Editar Registro"
      Top             =   7170
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
      MICON           =   "frmCidadao.frx":192E
      PICN            =   "frmCidadao.frx":194A
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
      Left            =   135
      TabIndex        =   14
      ToolTipText     =   "Novo Registro"
      Top             =   7170
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
      MICON           =   "frmCidadao.frx":1AA4
      PICN            =   "frmCidadao.frx":1AC0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Tributacao.jcFrames jcFrames2 
      Height          =   2370
      Index           =   1
      Left            =   60
      Top             =   4350
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   4180
      FrameColor      =   12829635
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "  Endere�o Comercial"
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
      Begin VB.CheckBox chkWhatsApp2 
         Caption         =   "WhatsApp"
         Height          =   225
         Left            =   2010
         TabIndex        =   85
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CheckBox chkFone2 
         Caption         =   "N�o possui telefone"
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   2040
         Width           =   1725
      End
      Begin prjChameleon.chameleonButton btPais2 
         Height          =   270
         Left            =   3300
         TabIndex        =   81
         ToolTipText     =   "Selecionar pais"
         Top             =   1680
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   476
         BTYPE           =   9
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCidadao.frx":1C1A
         PICN            =   "frmCidadao.frx":1C36
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtPais2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1065
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   67
         Top             =   1665
         Width           =   2175
      End
      Begin VB.TextBox txtEmail2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4425
         MaxLength       =   50
         TabIndex        =   66
         Top             =   1650
         Width           =   3000
      End
      Begin VB.TextBox txtFone2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4425
         MaxLength       =   30
         TabIndex        =   65
         Top             =   1995
         Width           =   2985
      End
      Begin VB.CheckBox chkEtiq2 
         Alignment       =   1  'Right Justify
         Height          =   225
         Left            =   135
         TabIndex        =   36
         Top             =   0
         Width           =   210
      End
      Begin VB.ComboBox cmbUF2 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmCidadao.frx":1D90
         Left            =   1080
         List            =   "frmCidadao.frx":1D92
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   915
         Width           =   2250
      End
      Begin VB.ComboBox cmbCidade2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4425
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   930
         Width           =   3000
      End
      Begin VB.ComboBox cmbBairro2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1305
         Width           =   2250
      End
      Begin VB.TextBox txtCompl2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2970
         MaxLength       =   50
         TabIndex        =   30
         Top             =   585
         Width           =   4455
      End
      Begin VB.TextBox txtNum2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   29
         Top             =   585
         Width           =   930
      End
      Begin VB.TextBox txtNumLog2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   27
         Top             =   240
         Width           =   930
      End
      Begin VB.TextBox txtNomeLog2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2205
         MaxLength       =   50
         TabIndex        =   28
         Top             =   240
         Width           =   5220
      End
      Begin esMaskEdit.esMaskedEdit mskCEP2 
         Height          =   285
         Left            =   4425
         TabIndex        =   35
         Top             =   1320
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         MouseIcon       =   "frmCidadao.frx":1D94
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
         MaxLength       =   9
         Mask            =   "99999-999"
         SelText         =   ""
         Text            =   "_____-___"
         HideSelection   =   -1  'True
      End
      Begin prjChameleon.chameleonButton cmdAddBairro2 
         Height          =   270
         Left            =   3345
         TabIndex        =   34
         ToolTipText     =   "Cadastrar um novo bairro"
         Top             =   1320
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   476
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
         MPTR            =   99
         MICON           =   "frmCidadao.frx":1DB0
         PICN            =   "frmCidadao.frx":20CA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Image imgWhatsApp2 
         Height          =   255
         Left            =   3120
         Picture         =   "frmCidadao.frx":2224
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pa�s.............:"
         Height          =   180
         Index           =   25
         Left            =   90
         TabIndex        =   70
         Top             =   1710
         Width           =   990
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Email...:"
         Height          =   225
         Index           =   24
         Left            =   3810
         TabIndex        =   69
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label lblFone2 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone.:"
         Height          =   225
         Left            =   3660
         TabIndex        =   68
         Top             =   2025
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CEP....:"
         Height          =   225
         Index           =   22
         Left            =   3810
         TabIndex        =   50
         Top             =   1365
         Width           =   585
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "UF...............:"
         Height          =   225
         Index           =   21
         Left            =   90
         TabIndex        =   49
         Top             =   990
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         Height          =   225
         Index           =   20
         Left            =   3810
         TabIndex        =   48
         Top             =   990
         Width           =   585
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro...........:"
         Height          =   225
         Index           =   19
         Left            =   90
         TabIndex        =   47
         Top             =   1365
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Complem.:"
         Height          =   225
         Index           =   18
         Left            =   2175
         TabIndex        =   46
         Top             =   645
         Width           =   765
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Logradouro..:"
         Height          =   225
         Index           =   17
         Left            =   90
         TabIndex        =   45
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "N�mero........:"
         Height          =   225
         Index           =   16
         Left            =   90
         TabIndex        =   44
         Top             =   615
         Width           =   975
      End
   End
   Begin VB.ListBox lstNomeLog 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   2175
      ItemData        =   "frmCidadao.frx":2604
      Left            =   2280
      List            =   "frmCidadao.frx":2606
      TabIndex        =   58
      Top             =   2205
      Visible         =   0   'False
      Width           =   5205
   End
   Begin prjChameleon.chameleonButton btObservacao 
      Height          =   315
      Left            =   2273
      TabIndex        =   76
      ToolTipText     =   "Observa��o do Cidad�o"
      Top             =   6780
      Width           =   1400
      _ExtentX        =   2461
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Observa��o"
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
      MICON           =   "frmCidadao.frx":2608
      PICN            =   "frmCidadao.frx":2624
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton btHistorico 
      Height          =   315
      Left            =   3923
      TabIndex        =   77
      ToolTipText     =   "Hist�rico do Cidad�o"
      Top             =   6780
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Hist�rico"
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
      MCOL            =   16711935
      MPTR            =   1
      MICON           =   "frmCidadao.frx":269E
      PICN            =   "frmCidadao.frx":26BA
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
Attribute VB_Name = "frmCidadao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type tReg
    sNome As String
    sRG As String
    sCPF As String
    sCNPJ As String
    sDataNascto As String
    sProfissao As String
    sNomeLogr As String
    sNumeroR As String
    sComplR As String
    sUFR As String
    sCidadeR As String
    sBairroR As String
    sCepR As String
    sPaisR As String
    sFoneR As String
    sEmailR As String
    bTemFoneR As Boolean
    sNomeLogC As String
    sNumeroC As String
    sComplC As String
    sUFC As String
    sCidadeC As String
    sBairroC As String
    sCepC As String
    sPaisC As String
    sFoneC As String
    sEmailC As String
    bTemFoneC As Boolean
End Type

Public bZOrder As Boolean
Dim RdoAux As rdoResultset
Dim Sql As String, bExec As Boolean
Dim Evento As String, sEnd As String
Dim frm As frmCnsCidadao
Dim NomeForm As String, sTipoCid As String

Dim sRet As String
Dim evEdit As Integer, evNew As Integer, evDel As Integer
Dim bEdit As Boolean, bNew As Boolean, bDel As Boolean, aReg(0) As tReg

Public Property Let sForm(sNomeForm As String)
    NomeForm = sNomeForm
End Property

Public Property Let sTipoCidadao(sValue As String)
    sTipoCid = sValue
End Property

Private Sub btCNPJ_Click()
Set frm2 = frmSituacaoTributaria
frm2.sDoc = mskCNPJ.Text
frm2.show

End Sub

Private Sub btCPF_Click()
Set frm2 = frmSituacaoTributaria
frm2.sDoc = mskCPF.Text
frm2.show

End Sub

Private Sub btHistorico_Click()

Dim frm2 As frmHistoricoCidadao
If Val(txtCod.Text) = 0 Then
   MsgBox "N�o existem Registros.", vbCritical, "Aten��o"
   Exit Sub
End If

Set frm2 = frmHistoricoCidadao
frm2.nContribuinte = Val(txtCod.Text)
frm2.show
frm2.ZOrder 0

End Sub

Private Sub btObservacao_Click()
Dim frm2 As frmObsCidadao
If Val(txtCod.Text) = 0 Then
   MsgBox "N�o existem Registros.", vbCritical, "Aten��o"
   Exit Sub
End If

Set frm2 = frmObsCidadao
frm2.nContribuinte = Val(txtCod.Text)
frm2.show vbModal

End Sub

Private Sub btPais_Click()
frmPaises.residencial = True
frmPaises.CodPais = Val(txtPais.Tag)
frmPaises.show vbModal

End Sub

Private Sub btPais2_Click()
frmPaises.residencial = False
frmPaises.CodPais = Val(txtPais2.Tag)
frmPaises.show vbModal

End Sub

Private Sub btProfissao_Click()
frmProfissao.nCodProfissao = Val(txtProfissao.Tag)
frmProfissao.show vbModal
End Sub

Private Sub chkEtiq_Click()
HabilitaEndR
End Sub

Private Sub chkEtiq2_Click()
HabilitaEndC
End Sub

Private Sub chkFone1_Click()
If chkFone1.value = vbChecked Then
    txtFone.Text = ""
    txtFone.Enabled = False
    txtFone.BackColor = Me.BackColor
    lblFone.Enabled = False
Else
    txtFone.Enabled = True
    txtFone.BackColor = Branco
    lblFone.Enabled = True
End If
End Sub

Private Sub chkFone2_Click()
If chkFone2.value = vbChecked Then
    txtFone2.Text = ""
    txtFone2.Enabled = False
    txtFone2.BackColor = Me.BackColor
    lblFone2.Enabled = False
Else
    txtFone2.Enabled = True
    txtFone2.BackColor = Branco
    lblFone2.Enabled = True
End If

End Sub

Public Sub cmbCidade_Click()
If Not bExec Then Exit Sub
If cmbCidade.ListIndex = -1 Then Exit Sub
cmbBairro.Clear
Sql = "SELECT CODBAIRRO,DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & Left$(cmbUF.Text, 2) & "' AND CODCIDADE=" & cmbCidade.ItemData(cmbCidade.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    cmbBairro.AddItem " "
    Do While Not .EOF
        If !DescBairro <> "" Then
            cmbBairro.AddItem !DescBairro
            cmbBairro.ItemData(cmbBairro.NewIndex) = !CodBairro
        End If
       .MoveNext
    Loop
   .Close
End With

End Sub

Public Sub cmbCidade2_Click()
If Not bExec Then Exit Sub
If cmbCidade2.ListIndex = -1 Then Exit Sub
cmbBairro2.Clear
Sql = "SELECT CODBAIRRO,DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & Left$(cmbUF2.Text, 2) & "' AND CODCIDADE=" & cmbCidade2.ItemData(cmbCidade2.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    cmbBairro2.AddItem " "
    Do While Not .EOF
       If !DescBairro <> "" Then
            cmbBairro2.AddItem !DescBairro
            cmbBairro2.ItemData(cmbBairro2.NewIndex) = !CodBairro
        End If
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub cmbFindCod_Click()
Dim z As Variant

z = InputBox("Digite o c�digo do cidad�o.", "Entre com a informa��o")
If Val(z) > 0 Then
    If Val(z) > 700000 Then
        MsgBox "C�digo inv�lido.", vbCritical, "Aten��o"
    Else
        txtCod.Text = Val(z)
        Limpa
        Le
        If NomeForm = "frmProcesso" Then
            Evento = "Alterar"
            Eventos "INCLUIR"
        End If
        
    End If
End If

End Sub

Private Sub cmbUF_Click()

If Not bExec Then Exit Sub
cmbCidade.Clear
cmbBairro.Clear
Sql = "SELECT CODCIDADE,DESCCIDADE FROM CIDADE WHERE SIGLAUF='" & Left$(cmbUF.Text, 2) & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do While Not .EOF
       cmbCidade.AddItem !descCidade
       cmbCidade.ItemData(cmbCidade.NewIndex) = !CodCidade
      .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub cmbUF2_Click()

If Not bExec Then Exit Sub
cmbCidade2.Clear
cmbBairro2.Clear
Sql = "SELECT CODCIDADE,DESCCIDADE FROM CIDADE WHERE SIGLAUF='" & Left$(cmbUF2.Text, 2) & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do While Not .EOF
       cmbCidade2.AddItem !descCidade
       cmbCidade2.ItemData(cmbCidade2.NewIndex) = !CodCidade
      .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub cmdAddBairro_Click()
If cmbUF.ListIndex = -1 Or cmbCidade.ListIndex = -1 Then
    MsgBox "Selecione UF e Cidade do bairro a ser cadastardo.", vbExclamation, "Aten��o"
    Exit Sub
End If

If Left(cmbUF.Text, 2) = "SP" And cmbCidade.ItemData(cmbCidade.ListIndex) = 413 Then
    MsgBox "Apenas bairros de fora podem ser cadastrados.", vbExclamation, "Aten��o"
    Exit Sub
End If

Set frm2 = frmBairro
frm2.FormCall = Me.Name & "R"
frm2.SiglaUF = cmbUF.Text
frm2.CodCidade = cmbCidade.ItemData(cmbCidade.ListIndex)
frmBairro.show

End Sub

Private Sub cmdAddBairro2_Click()
If cmbUF2.ListIndex = -1 Or cmbCidade2.ListIndex = -1 Then
    MsgBox "Selecione UF e Cidade do bairro a ser cadastardo.", vbExclamation, "Aten��o"
    Exit Sub
End If

If Left(cmbUF2.Text, 2) = "SP" And cmbCidade2.ItemData(cmbCidade2.ListIndex) = 413 Then
    MsgBox "Apenas bairros de fora podem ser cadastrados.", vbExclamation, "Aten��o"
    Exit Sub
End If

Set frm2 = frmBairro
frm2.FormCall = Me.Name & "C"
frm2.SiglaUF = cmbUF2.Text
frm2.CodCidade = cmbCidade2.ItemData(cmbCidade2.ListIndex)
frmBairro.show

End Sub

Private Sub cmdAlterar_Click()
    If Val(txtCod.Text) = 0 Then
       MsgBox "N�o existem Registros.", vbCritical, "Aten��o"
       Exit Sub
    End If
    If NomeDeLogin <> "ROSE" And NomeDeLogin <> "FERNANDA.SIMOLIN" And NomeDeLogin <> "JOSEANE" And NomeDeLogin <> "ANGELICA" And NomeDeLogin <> "MARIELA" And NomeDeLogin <> "TICYANNE.OKIMASU" And NomeDeLogin <> "DANIELE.SILVA" And NomeDeLogin <> "MARIELA.CUSTODIO" Then
 '       If Val(txtCod.Text) < 500000 Then
 '          MsgBox "N�o � possivel alterar cidad�os antigos.", vbCritical, "Aten��o"
 '          Exit Sub
 '       End If
    End If
    Eventos "INCLUIR"
    Evento = "Alterar"
    LeCampos
    txtNome.SetFocus
End Sub

Private Sub cmdCancel_Click()
Dim x As Long
    If Evento = "Alterar" Then
       x = Val(txtCod.Text)
       Limpa
       txtCod.Text = x
       Le
    Else
       Limpa
    End If
    lstNomeLog.Visible = False
    Eventos "INICIAR"
    Evento = ""
End Sub

Private Sub cmdConsultar_Click()
bZOrder = False
Set frm = frmCnsCidadao
frm.sForm = Me.Name
frmCnsCidadao.show
frmCnsCidadao.ZOrder 0
End Sub

Private Sub cmdCopy_Click()
Dim z As Variant, Sql As String, RdoAux As rdoResultset
If Val(txtCod.Text) = 0 Then
   MsgBox "N�o existem Registros.", vbCritical, "Aten��o"
   Exit Sub
End If

z = InputBox("Digite o c�digo do cidad�o original.", "Entre com a informa��o")
If Val(z) < 500000 Or Val(z) > 700000 Then
    MsgBox "C�digo inv�lido.", vbCritical, "Erro"
    Exit Sub
End If

Sql = "SELECT * FROM CIDADAO WHERE CODCIDADAO=" & Val(z)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        MsgBox "C�digo n�o cadastrado.", vbCritical, "Erro"
    Else
        If MsgBox("Deseja copiar os dados do cidad�o " & !CodCidadao & "-" & !nomecidadao & " ?", vbYesNo + vbQuestion, "Confirma��o") = vbYes Then
            Log Form, Me.Caption, Altera��o, "Copiado cidad�o " & !CodCidadao & " - " & nomecidadao & " para o c�digo " & txtCod.Text & " - " & txtNome.Text
            Sql = "update cidadao set cpf='" & SubNull(!CPF) & "',cnpj='" & SubNull(!Cnpj) & "',codlogradouro=" & Val(SubNull(!CodLogradouro)) & ","
            Sql = Sql & "numimovel=" & Val(SubNull(!NUMIMOVEL)) & ",complemento='" & SubNull(!Complemento) & "',codbairro=" & Val(SubNull(!CodBairro)) & ","
            Sql = Sql & "codcidade=" & Val(!CodCidade) & ",siglauf='" & SubNull(!SiglaUF) & "' where codcidadao=" & Val(txtCod.Text)
            cn.Execute Sql, rdExecDirect
            z = Val(txtCod.Text)
            Limpa
            txtCod.Text = z
            Le
        End If

    End If
   .Close
End With

End Sub

Private Sub cmdExcluir_Click()
Dim x As Integer
On Error GoTo Erro
If NomeDeLogin <> "SCHWARTZ" And NomeDeLogin <> "RENATA" And NomeDeLogin <> "SOLANGE" And NomeDeLogin <> "LUIZH" Then
   MsgBox "N�o � poss�vel excluir.", vbCritical, "Aten��o"
   Exit Sub
End If

If Val(txtCod.Text) = 0 Then
   MsgBox "N�o existem Registros.", vbCritical, "Aten��o"
   Exit Sub
ElseIf Val(txtCod.Text) < 500000 Then
   MsgBox "Este cidad�o n�o pode ser excluido.", vbCritical, "Aten��o"
   Exit Sub
End If

Sql = "SELECT * FROM PROPRIETARIO WHERE CODCIDADAO=" & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
If RdoAux.RowCount > 0 Then
   MsgBox "N�o � poss�vel excluir este Cidad�o pois ele � Propriet�rio/Propriet�rio Solid�rio de um Im�vel.", vbExclamation, "Aten��o"
   Exit Sub
End If

If MsgBox("Excluir este Cidad�o ?", vbQuestion + vbYesNoCancel, "Aten��o") = vbYes Then
   Sql = "DELETE FROM CIDADAO WHERE CODCIDADAO=" & txtCod.Text
   cn.Execute Sql, rdExecDirect
   Limpa
End If
    
Exit Sub

Erro:
For x = 0 To rdoErrors.Count - 1
   MsgBox rdoErrors(x).Description
Next

End Sub

Private Sub cmdGravar_Click()
If bLocal Then
    Exit Sub
End If

Dim Y As Integer
    If txtNome.Text = "" Then
       MsgBox "Digite o Nome do Cidad�o.", vbExclamation, "Aten��o"
       txtNome.SetFocus
       Exit Sub
    End If
    
    If mskDtNascto.ClipText <> "" Then
        If Not IsDate(mskDtNascto.Text) Then
            MsgBox "Data de nascimento inv�lida.", vbCritical, "Erro"
            Exit Sub
        End If
    End If
    
'    If NomeDeLogin <> "LUIZH" Then
  '      If mskCPF.ClipText = "" And mskCNPJ.ClipText = "" And mskRG.ClipText = "" Then
        If mskCPF.ClipText = "" And mskCNPJ.ClipText = "" Then
           MsgBox "Digite CPF ou CNPJ.", vbExclamation, "Aten��o"
           Exit Sub
        End If
           
        If txtOrgao.Text <> "" And mskRG.ClipText = "" Then
           MsgBox "Digite o RG.", vbExclamation, "Aten��o"
           Exit Sub
        End If
           
        If mskCPF.ClipText <> "" And mskCNPJ.ClipText <> "" Then
           MsgBox "Digite CPF ou CNPJ.", vbExclamation, "Aten��o"
           Exit Sub
        End If
        
        If chkEtiq.value = vbUnchecked And chkEtiq2.value = vbUnchecked And NomeDeLogin <> "SCHWARTZ" Then
           MsgBox "Preencha ao menos um endere�o.", vbExclamation, "Aten��o"
           Exit Sub
        End If
 '   End If
    
    If chkEtiq.value = vbChecked Then
        If Val(txtNum.Text) > 10000 Then
           MsgBox "N� de Im�vel inv�lido.", vbExclamation, "Aten��o"
           Exit Sub
        End If
            
        If Val(txtNumLog.Text) = 0 And cmbCidade.Text = "JABOTICABAL" Then
           MsgBox "Selecione o Logradouro de Jaboticabal.", vbCritical, "Erro de Valida��o."
           txtNumLog.SetFocus
           Exit Sub
        End If
    
        If Val(txtNumLog.Text) = 0 And txtNomeLog.Text = "" Then
           MsgBox "Digite o Logradouro.", vbCritical, "Erro de Valida��o."
           txtNomeLog.SetFocus
           Exit Sub
        End If
    
'        If Trim(txtPais.Text) <> "" And cmbUF.ListIndex > 0 Then
'           MsgBox "Pa�s n�o pode ter cidade/UF.", vbExclamation, "Aten��o"
'           Exit Sub
'        End If
    
        If txtPais.Text = "" And cmbCidade.ListIndex = -1 Then
           MsgBox "Favor selecionar a cidade.", vbExclamation, "Aten��o"
           Exit Sub
        End If
    
        If txtPais = "" And txtNomeLog.Text <> "" And cmbBairro.ListIndex = -1 Then
           MsgBox "Favor selecionar o bairro residencial.", vbExclamation, "Aten��o"
           Exit Sub
        End If
    
        If mskCEP.ClipText = "" And txtPais.Text = "BRASIL" Then
           MsgBox "Favor cadastrar o CEP.", vbExclamation, "Aten��o"
           Exit Sub
        End If
        
        If chkFone1.value = vbUnchecked And Len(txtFone.Text) < 6 Then
           MsgBox "Informe o n� de telefone ou marque a op��o que n�o possui.", vbExclamation, "Aten��o"
           Exit Sub
        End If

        If chkWhatsApp.value = vbChecked And Len(txtFone.Text) < 6 Then
            chkWhatsApp.value = vbUnchecked
        End If

    End If
    
    If chkEtiq2.value = vbChecked Then
        If Val(txtNum2.Text) > 10000 Then
           MsgBox "N� de Im�vel comercial inv�lido.", vbExclamation, "Aten��o"
           Exit Sub
        End If
            
        If Val(txtNumLog2.Text) = 0 And cmbCidade2.Text = "JABOTICABAL" Then
           MsgBox "Selecione o Logradouro de Jaboticabal.", vbCritical, "Erro de Valida��o."
           txtNumLog2.SetFocus
           Exit Sub
        End If
    
        If Val(txtNumLog2.Text) = 0 And txtNomeLog2.Text = "" Then
           MsgBox "Digite o Logradouro comercial.", vbCritical, "Erro de Valida��o."
           txtNomeLog2.SetFocus
           Exit Sub
        End If
    
 '       If Trim(txtPais2.Text) <> "" And cmbUF2.ListIndex > 0 Then
 '          MsgBox "Pa�s n�o pode ter cidade/UF.", vbExclamation, "Aten��o"
 '          Exit Sub
 '       End If
    
        If cmbCidade2.ListIndex = -1 Then
           MsgBox "Favor selecionar a cidade end.comercial.", vbExclamation, "Aten��o"
           Exit Sub
        End If
    
        If txtNomeLog2.Text <> "" And cmbBairro2.ListIndex < 1 Then
           MsgBox "Favor selecionar o bairro end.comercial.", vbExclamation, "Aten��o"
           Exit Sub
        End If
    
        If mskCEP2.ClipText = "" Then
           MsgBox "Favor cadastrar o CEP end.comercial.", vbExclamation, "Aten��o"
           Exit Sub
        End If

        If chkFone2.value = vbUnchecked And Len(txtFone2.Text) < 6 Then
           MsgBox "Informe o n� de telefone ou marque a op��o que n�o possui.", vbExclamation, "Aten��o"
           Exit Sub
        End If


        If chkWhatsApp2.value = vbChecked And Len(txtFone2.Text) < 6 Then
            chkWhatsApp2.value = vbUnchecked
        End If

    End If
       
    If mskCPF.ClipText <> "" Then
        If Not ValidaCPF(mskCPF.ClipText) Or mskCPF.ClipText = "00000000000" Then
           MsgBox "CPF inv�lido.", vbExclamation, "Aten��o"
           Exit Sub
        End If
        
        If Evento = "Novo" Then
             Sql = "SELECT CODCIDADAO,NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO>500000 AND CPF='" & mskCPF.ClipText & "'"
        Else
             Sql = "SELECT CODCIDADAO,NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO>500000 AND CODCIDADAO<>" & Val(txtCod.Text) & " AND CPF='" & mskCPF.ClipText & "'"
        End If
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
               If MsgBox("CPF j� cadastrado ---> " & Format(!CodCidadao, "00000") & " - " & !nomecidadao & vbCrLf & "Deseja gravar assim mesmo?", vbCritical + vbYesNo, "CPF Duplicado !!!") = vbNo Then
                   Exit Sub
               End If
            End If
        End With
    End If
    
    If mskCNPJ.ClipText <> "" Then
       If Not ValidaCGC(mskCNPJ.ClipText) Then
          MsgBox "CNPJ inv�lido.", vbExclamation, "Aten��o"
          Exit Sub
       End If
       If Evento = "Novo" Then
            Sql = "SELECT CODCIDADAO,NOMECIDADAO FROM CIDADAO WHERE CNPJ='" & mskCNPJ.ClipText & "'"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount > 0 Then
                   MsgBox "CNPJ j� cadastrado ---> " & Format(!CodCidadao, "00000") & " - " & !nomecidadao, vbCritical, "CNPJ Duplicado !!!"
                   Exit Sub
                End If
            End With
       End If
    End If
    
    If mskRG.ClipText <> "" Then
       If Evento = "Novo" Then
            Sql = "SELECT CODCIDADAO,NOMECIDADAO FROM CIDADAO WHERE RG='" & mskRG.Text & "'"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount > 0 Then
                   MsgBox "RG j� cadastrado ---> " & Format(!CodCidadao, "00000") & " - " & !nomecidadao, vbCritical, "RG Duplicado !!!"
                   Exit Sub
                End If
            End With
       End If
    End If
    
    Grava
    Eventos "INICIAR"
    If Evento = "Novo" Then
        For x = 0 To Forms.Count - 1
            If Forms(x).Name = "frmCadImob" Then
                For Y = 1 To frmCadImob.tvProp.Nodes.Count
                    On Error GoTo fim
                    If Right$(frmCadImob.tvProp.Nodes(x).Key, 5) = Format(txtCod.Text, "00000") Then
                       MsgBox "Ja existe um Propriet�rio ou Propriet�rio Solid�rio com este nome.", vbCritical, "Aten��o"
                       GoTo fim
                    End If
                Next
                If sTipoCid = "P" Then
                   Set NodX = frmCadImob.tvProp.Nodes.Add("PROP", tvwChild, "PROP" & Format(txtCod.Text, "00000"), txtNome.Text, 1)
                   frmCadImob.tvProp.Nodes("PROP" & Format(txtCod.Text, "00000")).ForeColor = vbBlue
                Else
                   Set NodX = frmCadImob.tvProp.Nodes.Add("COMP", tvwChild, "COMP" & Format(txtCod.Text, "00000"), txtNome.Text, 2)
                   frmCadImob.tvProp.Nodes("COMP" & Format(txtCod.Text, "00000")).ForeColor = vbBlue
                End If
                For Y = 1 To frmCadImob.tvProp.Nodes.Count
                    frmCadImob.tvProp.Nodes(x).EnsureVisible
                Next
fim:
                sTipoCid = ""
                Unload frmCidadao
                frmCadImob.SetFocus
                Exit Sub
            ElseIf Forms(x).Name = "frmCadMob" Then
                For Y = 1 To frmCadMob.grdProp.Rows - 1
                    If Val(frmCadMob.grdProp.TextMatrix(Y, 0)) = Val(txtCod.Text) Then
                       MsgBox "Ja existe um Propriet�rio ou Propriet�rio Solid�rio com este nome.", vbCritical, "Aten��o"
                       GoTo Fim2
                    End If
                Next
                frmCadMob.grdProp.AddItem Format(txtCod.Text, "00000") & Chr(9) & txtNome.Text & Chr(9) & IIf(mskCPF.ClipText <> "", mskCPF.Text, mskCNPJ.Text)
Fim2:
                sTipoCid = ""
                Unload frmCidadao
                frmCadMob.SetFocus
                Exit Sub
            End If
        Next
    End If
End Sub

Public Sub cmdNovo_Click()
    Limpa
    txtCod.Text = ""
    cmbUF.Text = "SP-S�O PAULO"
    cmbUF_Click
    cmbCidade.Text = "JABOTICABAL"
    cmbCidade_Click
    Eventos "INCLUIR"
    Evento = "Novo"
    txtNome.SetFocus
End Sub

Private Sub cmdSair_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Form_Activate()
On Error Resume Next
If CodCidadao > 0 Then
    Limpa
    txtCod.Text = Format(CodCidadao, "00000")
    Le
    
    If NomeForm = "frmProcesso" Then
        Evento = "Alterar"
        Eventos "INCLUIR"
    End If

End If
CodCidadao = 0

If NomeForm = "frmProcesso" Then
    'bZOrder = True
    If frmProcesso.cmdGravar.Visible = False Then
        cmdNovo.Enabled = False
        cmdAlterar.Enabled = False
        cmdExcluir.Enabled = False
        cmdConsultar.Enabled = False
    End If
End If

Liberado

End Sub

Private Sub Form_Load()
bZOrder = True
Centraliza Me
sRet = RetEventUserForm(Me.Name)
Limpa
sEnd = "R"
bExec = False
cmbUF.AddItem " "
Sql = "SELECT SIGLAUF,DESCUF FROM UF"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset)
Do While Not RdoAux.EOF
   cmbUF.AddItem RdoAux!SiglaUF & "-" & RdoAux!DESCUF
   RdoAux.MoveNext
Loop
cmbUF2.AddItem " "
Sql = "SELECT SIGLAUF,DESCUF FROM UF"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset)
Do While Not RdoAux.EOF
   cmbUF2.AddItem RdoAux!SiglaUF & "-" & RdoAux!DESCUF
   RdoAux.MoveNext
Loop


bExec = True
RdoAux.Close

'If NomeForm = "frmProcesso" Then
'    Eventos "INCLUIR"
'Else
Eventos "INICIAR"
' End If
End Sub

Private Sub Le()
Dim RdoS As rdoResultset

Sql = "SELECT cidadao.*, profissao.nome AS profissao_nome,pais_1.nome_pais, pais.nome_pais AS nome_pais2 "
Sql = Sql & "FROM profissao RIGHT OUTER JOIN cidadao LEFT OUTER JOIN pais AS pais_1 ON cidadao.codpais = pais_1.id_pais LEFT OUTER JOIN "
Sql = Sql & "pais ON cidadao.codpais2 = pais.id_pais ON profissao.codigo = cidadao.codprofissao Where CodCidadao = " & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

With RdoAux
    If .RowCount = 0 Then Exit Sub
    txtCod.Text = !CodCidadao
    If Not IsNull(!CPF) Then mskCPF.Text = Format(Trim(!CPF), "000\.000\.000-00")
    If Not IsNull(!Cnpj) Then mskCNPJ.Text = Format(Trim(!Cnpj), "00\.000\.000/0000-00")
    If Not IsNull(!rg) Then mskRG.Text = !rg
    If Not IsNull(!data_nascimento) Then
        mskDtNascto.Text = Format(!data_nascimento, "dd/mm/yyyy")
    End If
    
    txtNome.Text = !nomecidadao
    If !CodLogradouro > 0 Then
       txtNumLog.Text = !CodLogradouro
       txtNumLog_LostFocus
    Else
       txtNumLog.Text = 0
       txtNomeLog.Text = SubNull(!NomeLogradouro)
    End If
    
    If !CodLogradouro2 > 0 Then
       txtNumLog2.Text = !CodLogradouro2
       txtNumLog2_LostFocus
    Else
       txtNumLog2.Text = 0
       txtNomeLog2.Text = SubNull(!NOMELOGRADOURO2)
    End If
    
    txtNum.Text = SubNull(!NUMIMOVEL)
    txtCompl.Text = SubNull(!Complemento)
    
    txtNum2.Text = SubNull(!NUMIMOVEL2)
    txtCompl2.Text = SubNull(!Complemento2)
    
    If IsNull(!codprofissao) Then
        txtProfissao.Text = "(N�o especificado)"
        txtProfissao.Tag = "1"
    Else
        txtProfissao.Text = SubNull(!profissao_nome)
        txtProfissao.Tag = Val(SubNull(!codprofissao))
    End If
    
    If IsNull(!CodPais) Then
        txtPais.Text = "BRASIL"
        txtPais.Tag = "1"
    Else
        txtPais.Text = SubNull(!nome_pais)
        txtPais.Tag = Val(SubNull(!CodPais))
    End If
    
    If IsNull(!CodPais2) Then
        txtPais2.Text = "BRASIL"
        txtPais2.Tag = "1"
    Else
        txtPais2.Text = SubNull(!nome_pais2)
        txtPais2.Tag = Val(SubNull(!CodPais2))
    End If
    
    
    If Trim$(SubNull(!SiglaUF)) <> "" Then
       bExec = False
       Sql = "SELECT DESCUF FROM UF WHERE SIGLAUF='" & !SiglaUF & "'"
       Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
       cmbUF.Text = !SiglaUF & "-" & RdoS!DESCUF
       bExec = True
       cmbUF_Click
       If !CodCidade > 0 Then
          bExec = False
            Sql = "SELECT DESCCIDADE FROM CIDADE WHERE SIGLAUF='" & !SiglaUF & "' AND CODCIDADE=" & !CodCidade
            Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
            cmbCidade.Text = RdoS!descCidade
            bExec = True
           cmbCidade_Click
           If !CodBairro > 0 Then
                Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & !SiglaUF & "' AND CODCIDADE=" & !CodCidade & " AND CODBAIRRO=" & !CodBairro
                Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
                If RdoS.RowCount > 0 Then
                If RdoS!DescBairro <> "" Then
                    cmbBairro.Text = RdoS!DescBairro
                End If
                End If
          End If
       End If
    End If
    
    
    If Trim$(SubNull(!SiglaUF2)) <> "" Then
       bExec = False
       Sql = "SELECT DESCUF FROM UF WHERE SIGLAUF='" & !SiglaUF2 & "'"
       Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
       cmbUF2.Text = !SiglaUF2 & "-" & RdoS!DESCUF
       bExec = True
       cmbUF2_Click
       If !CodCidade2 > 0 Then
          bExec = False
            Sql = "SELECT DESCCIDADE FROM CIDADE WHERE SIGLAUF='" & !SiglaUF2 & "' AND CODCIDADE=" & !CodCidade2
            Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
            cmbCidade2.Text = RdoS!descCidade
            bExec = True
           cmbCidade2_Click
           If !CodBairro2 > 0 Then
                Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & !SiglaUF2 & "' AND CODCIDADE=" & !CodCidade2 & " AND CODBAIRRO=" & !CodBairro2
                Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
                If RdoS!DescBairro <> "" Then
                    cmbBairro2.Text = RdoS!DescBairro
                End If
          End If
       End If
    End If
    
    If Not IsNull(!Cep) Then
       mskCEP.Text = Format(!Cep, "00000-000")
    End If
    If Not IsNull(!Cep2) Then
       mskCEP2.Text = Format(!Cep2, "00000-000")
    End If
    txtFone.Text = SubNull(!telefone)
    txtEmail.Text = SubNull(!Email)
    txtOrgao.Text = SubNull(!ORGAO)
'    txtPais.Text = SubNull(!pais)
'    txtPais2.Text = SubNull(!PAIS2)
    txtFone2.Text = SubNull(!Telefone2)
    txtEmail2.Text = SubNull(!EMAIL2)
    OptP(0).value = True
    If Not IsNull(!juridica) Then
        If !juridica Then
            OptP(1).value = True
        End If
    End If
    If IsNull(!etiqueta) Then
        If cmbBairro.ListIndex > -1 Then
            chkEtiq.value = vbChecked
        Else
            chkEtiq.value = vbUnchecked
        End If
    Else
        If !etiqueta = "S" Then
            chkEtiq.value = vbChecked
        Else
            chkEtiq.value = vbUnchecked
        End If
    End If
   
    If IsNull(!etiqueta2) Then
        If cmbBairro2.ListIndex > -1 Then
            chkEtiq2.value = vbChecked
        Else
            chkEtiq2.value = vbUnchecked
        End If
    Else
        If !etiqueta2 = "S" Then
            chkEtiq2.value = vbChecked
        Else
            chkEtiq2.value = vbUnchecked
        End If
    End If
    If IsNull(!temfone) Then
        chkFone1.value = vbUnchecked
    Else
        If !temfone = True Then
            chkFone1.value = vbChecked
        Else
            chkFone1.value = vbUnchecked
        End If
    End If
    If IsNull(!temfone2) Then
        chkFone2.value = vbUnchecked
    Else
        If !temfone2 = True Then
            chkFone2.value = vbChecked
        Else
            chkFone2.value = vbUnchecked
        End If
    End If
    
    If IsNull(!whatsapp) Then
        chkWhatsApp.value = vbUnchecked
    Else
        If !whatsapp = True Then
            chkWhatsApp.value = vbChecked
        Else
            chkWhatsApp.value = vbUnchecked
        End If
    End If
    
    If IsNull(!whatsapp2) Then
        chkWhatsApp2.value = vbUnchecked
    Else
        If !whatsapp2 = True Then
            chkWhatsApp2.value = vbChecked
        Else
            chkWhatsApp2.value = vbUnchecked
        End If
    End If
    
   .Close
End With

Sql = "SELECT  espolio.codigo, tipousuario.nome FROM  espolio INNER JOIN  tipousuario ON espolio.tipo = tipousuario.codigo WHERE espolio.codigo=" & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    lblEspolio.Caption = RdoAux!Nome
    lblEspolio.Visible = True
Else
    lblEspolio.Visible = False
End If
RdoAux.Close

LeCampos

End Sub

Private Sub LeCampos()
With aReg(0)
    .sBairroC = cmbBairro2.Text
    .sBairroR = cmbBairro.Text
    .sCepC = mskCEP2.Text
    .sCepR = mskCEP.Text
    .sCidadeC = cmbCidade2.Text
    .sCidadeR = cmbCidade.Text
    .sCNPJ = mskCNPJ.Text
    .sComplC = txtCompl2.Text
    .sComplR = txtCompl.Text
    .sCPF = mskCPF.Text
    .sDataNascto = mskDtNascto.Text
    .sEmailC = txtEmail2.Text
    .sEmailR = txtEmail.Text
    .sFoneC = txtFone2.Text
    .sFoneR = txtFone.Text
    .sNome = txtNome.Text
    .sNomeLogC = txtNomeLog2.Text
    .sNomeLogr = txtNomeLog.Text
    .sNumeroC = txtNum2.Text
    .sNumeroR = txtNum.Text
    .sPaisC = txtPais2.Text
    .sPaisR = txtPais.Text
    .sProfissao = txtProfissao.Text
    .sRG = mskRG.Text
    .sUFC = cmbUF2.Text
    .sUFR = cmbUF.Text
     .bTemFoneR = IIf(chkFone1.value = 1, True, False)
     .bTemFoneC = IIf(chkFone2.value = 1, True, False)
End With

End Sub

Private Sub Eventos(Tipo As String)

Dim Ct As Control

If Tipo = "INICIAR" Then
   cmdNovo.Visible = True
   cmdCopy.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdSair.Visible = True
   cmdConsultar.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   For Each Ct In frmCidadao
       If TypeOf Ct Is TextBox Or TypeOf Ct Is ComboBox Then
          Ct.BackColor = Kde
           Ct.Enabled = False
       End If
   Next
   mskDtNascto.Enabled = False
   mskDtNascto.BackColor = Kde
   
   mskCPF.Enabled = False
   mskCPF.BackColor = Kde
   mskRG.Enabled = False
   mskRG.BackColor = Kde
   mskCEP.Enabled = False
   mskCEP.BackColor = Kde
   mskCEP2.Enabled = False
   mskCEP2.BackColor = Kde
   mskCNPJ.Enabled = False
   mskCNPJ.BackColor = Kde
   chkEtiq.Enabled = False
   chkEtiq2.Enabled = False
   OptP(0).Enabled = False
   OptP(1).Enabled = False
   cmdAddBairro.Enabled = False
   cmdAddBairro2.Enabled = False
   btProfissao.Enabled = False
   btPais.Enabled = False
   btPais2.Enabled = False
   chkFone1.Enabled = False
   chkFone2.Enabled = False
   chkWhatsApp.Enabled = False
   chkWhatsApp2.Enabled = False
   imgWhatsApp.Enabled = False
   imgWhatsApp2.Enabled = False
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdCopy.Visible = False
   btProfissao.Enabled = True
   btPais.Enabled = True
   btPais2.Enabled = True
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdConsultar.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   For Each Ct In frmCidadao
       If TypeOf Ct Is TextBox Or TypeOf Ct Is ComboBox Then
          Ct.BackColor = vbWhite
          Ct.Enabled = True
       End If
   Next
   txtProfissao.BackColor = Kde
   mskDtNascto.Enabled = True
   mskDtNascto.BackColor = vbWhite
   mskCPF.Enabled = True
   mskCPF.BackColor = vbWhite
   mskRG.Enabled = True
   mskRG.BackColor = vbWhite
   mskCNPJ.Enabled = True
   mskCNPJ.BackColor = vbWhite
   mskCEP.Enabled = True
   mskCEP.BackColor = vbWhite
   mskCEP2.Enabled = True
   mskCEP2.BackColor = vbWhite
   txtCod.Locked = True
   txtCod.BackColor = Kde
   chkEtiq.Enabled = True
   chkEtiq2.Enabled = True
   OptP(0).Enabled = True
   OptP(1).Enabled = True

    HabilitaEndR
    HabilitaEndC
End If

FormHagana

End Sub

Private Sub HabilitaEndR()

Dim Cor As Long
Dim bValue As Boolean

If chkEtiq.value = vbChecked Then
    bValue = True
Else
    bValue = False
End If

If bValue And cmdNovo.Visible Then Exit Sub


If bValue Then
    Cor = vbWhite
Else
    Cor = Kde
End If

txtNumLog.Enabled = bValue
txtNomeLog.Enabled = bValue
txtNum.Enabled = bValue
txtCompl.Enabled = bValue
cmbUF.Enabled = bValue
cmbCidade.Enabled = bValue
cmbBairro.Enabled = bValue
mskCEP.Enabled = bValue
txtPais.Enabled = bValue
txtFone.Enabled = bValue
txtEmail.Enabled = bValue
cmdAddBairro.Enabled = bValue
chkFone1.Enabled = bValue
chkWhatsApp.Enabled = bValue
imgWhatsApp.Enabled = bValue

txtNumLog.BackColor = Cor
txtNomeLog.BackColor = Cor
txtNum.BackColor = Cor
txtCompl.BackColor = Cor
cmbUF.BackColor = Cor
cmbCidade.BackColor = Cor
cmbBairro.BackColor = Cor
mskCEP.BackColor = Cor
txtPais.BackColor = Cor
txtFone.BackColor = Cor
txtEmail.BackColor = Cor

End Sub

Private Sub HabilitaEndC()

Dim Cor As Long
Dim bValue As Boolean

If chkEtiq2.value = vbChecked Then
    bValue = True
Else
    bValue = False
End If

If bValue And cmdNovo.Visible Then Exit Sub

If bValue Then
    Cor = vbWhite
Else
    Cor = Kde
End If

txtNumLog2.Enabled = bValue
txtNomeLog2.Enabled = bValue
txtNum2.Enabled = bValue
txtCompl2.Enabled = bValue
cmbUF2.Enabled = bValue
cmbCidade2.Enabled = bValue
cmbBairro2.Enabled = bValue
mskCEP2.Enabled = bValue
txtPais2.Enabled = bValue
txtFone2.Enabled = bValue
chkWhatsApp2.Enabled = bValue
imgWhatsApp2.Enabled = bValue

txtEmail2.Enabled = bValue
cmdAddBairro2.Enabled = bValue
chkFone2.Enabled = bValue
txtNumLog2.BackColor = Cor
txtNomeLog2.BackColor = Cor
txtNum2.BackColor = Cor
txtCompl2.BackColor = Cor
cmbUF2.BackColor = Cor
cmbCidade2.BackColor = Cor
cmbBairro2.BackColor = Cor
mskCEP2.BackColor = Cor
txtPais2.BackColor = Cor
txtFone2.BackColor = Cor
txtEmail2.BackColor = Cor

End Sub
Private Sub Limpa()

LimpaMascara mskCPF
LimpaMascara mskCNPJ
LimpaMascara mskRG
txtNome.Text = ""
txtNumLog.Text = 0
txtNomeLog.Text = ""
txtNum.Text = 0
txtCompl.Text = ""
cmbBairro.ListIndex = -1
cmbCidade.ListIndex = -1
cmbUF.ListIndex = -1
LimpaMascara mskCEP
chkEtiq.value = vbUnchecked
chkEtiq2.value = vbUnchecked
lblEspolio.Visible = False
txtNumLog2.Text = 0
txtNomeLog2.Text = ""
txtNum2.Text = 0
txtCompl2.Text = ""
cmbBairro2.ListIndex = -1
cmbCidade2.ListIndex = -1
cmbUF2.ListIndex = -1
LimpaMascara mskCEP2
LimpaMascara mskDtNascto
txtProfissao.Text = ""
txtProfissao.Tag = ""
chkEtiq2.value = vbUnchecked
txtFone.Text = ""
txtEmail.Text = ""
txtPais.Text = ""
txtFone2.Text = ""
txtEmail2.Text = ""
txtPais2.Text = ""
chkFone1.value = vbUnchecked
chkFone2.value = vbUnchecked
chkWhatsApp.value = vbUnchecked
chkWhatsApp2.value = vbUnchecked
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

On Error Resume Next
If NomeForm = "frmProcesso" And Val(txtCod.Text) < 500000 Then
    MsgBox "Selecione apenas cidad�o acima de 500000.", vbExclamation, "Aten��o"
    frmProcesso.HabilitaPainelPrincipal
    frmProcesso.cmdEditCid.Enabled = True
    frmProcesso.cmdEditDoc.Enabled = True
    frmProcesso.cmdEditEnd.Enabled = True
    Exit Sub
End If
If NomeForm = "frmProcesso" Then
    frmProcesso.lblCodCid.Caption = txtCod.Text
    frmProcesso.lblNomeCid.Caption = txtNome.Text
    frmProcesso.HabilitaPainelPrincipal
    If sTipoCid = "A" Then
        frmProcesso.cmbOrigem.BackColor = Kde
        frmProcesso.cmbAssunto.BackColor = Kde
        frmProcesso.chkFisico.BackColor = Kde
        frmProcesso.chkInterno.BackColor = Kde
        frmProcesso.txtCompl.BackColor = Kde
        frmProcesso.cmbOrigem.Enabled = False
        frmProcesso.cmbAssunto.Enabled = False
        frmProcesso.chkFisico.Enabled = False
        frmProcesso.chkInterno.Enabled = False
        frmProcesso.txtCompl.Locked = True
    End If
    frmProcesso.cmdEditCid.Enabled = True
    frmProcesso.cmdEditDoc.Enabled = True
    frmProcesso.cmdEditEnd.Enabled = True
    frmProcesso.SetFocus
    NomeForm = ""
End If
If NomeForm = "frmCPProcesso" And Val(txtCod.Text) < 500000 Then
    MsgBox "Selecione apenas cidad�o acima de 500000.", vbExclamation, "Aten��o"
    frmCPProcesso.HabilitaPainelPrincipal
    frmCPProcesso.cmdEditCid.Enabled = True
    frmCPProcesso.cmdEditDoc.Enabled = True
    frmCPProcesso.cmdEditEnd.Enabled = True
    Exit Sub
End If
If NomeForm = "frmCPProcesso" Then
    frmCPProcesso.lblCodCid.Caption = txtCod.Text
    frmCPProcesso.lblNomeCid.Caption = txtNome.Text
    frmCPProcesso.HabilitaPainelPrincipal
    If sTipoCid = "A" Then
        frmCPProcesso.cmbOrigem.BackColor = Kde
        frmCPProcesso.cmbAssunto.BackColor = Kde
        frmCPProcesso.chkFisico.BackColor = Kde
        frmCPProcesso.chkInterno.BackColor = Kde
        frmCPProcesso.txtCompl.BackColor = Kde
        frmCPProcesso.cmbOrigem.Enabled = False
        frmCPProcesso.cmbAssunto.Enabled = False
        frmCPProcesso.chkFisico.Enabled = False
        frmCPProcesso.chkInterno.Enabled = False
        frmCPProcesso.txtCompl.Locked = True
    End If
    frmCPProcesso.cmdEditCid.Enabled = True
    frmCPProcesso.cmdEditDoc.Enabled = True
    frmCPProcesso.cmdEditEnd.Enabled = True
    frmCPProcesso.SetFocus
    NomeForm = ""
End If
If NomeForm = "frmRequerimento" Then
    frmRequerimento.lblCodRequerente.Caption = txtCod.Text
    frmRequerimento.lblRequerente.Caption = txtNome.Text
ElseIf NomeForm = "frmRequerIPTU" Then
    frmRequerIPTU.lblRequerente.Tag = txtCod.Text
    frmRequerIPTU.lblRequerente.Caption = txtNome.Text
ElseIf NomeForm = "frmRequerIPTU2" Then
    frmRequerIPTU.lblRazao.Tag = txtCod.Text
    frmRequerIPTU.lblRazao.Caption = txtNome.Text
    frmRequerIPTU.lblCNPJ.Caption = mskCNPJ.Text
    frmRequerIPTU.txtEnd.Text = txtNomeLog.Text & ", " & txtNum.Text
ElseIf NomeForm = "frmDeclaraIsento" Then
    frmDeclaraIsento.lblRequerente.Tag = txtCod.Text
    frmDeclaraIsento.lblRequerente.Caption = txtNome.Text
    frmDeclaraIsento.lblCPF.Caption = mskCPF.Text
    frmDeclaraIsento.lblRG.Caption = mskRG.Text
ElseIf NomeForm = "frmRegistroAtendimento" Then
    frmRegistroAtendimento.txtCidadao.Text = Format(txtCod.Text, "000000") & " - " & txtNome.Text
    frmRegistroAtendimento.txtEnd.Text = txtNomeLog.Text & ", " & txtNum.Text
    frmRegistroAtendimento.txtBairro = cmbBairro.Text
    frmRegistroAtendimento.txtCompl.Text = txtCompl.Text
    frmRegistroAtendimento.txtFone.Text = txtFone.Text
End If

'CodCidadao = Val(txtCod.Text)
Unload frmCnsCidadao
Unload frmCidadao
End Sub

Private Sub lstNomeLog_LostFocus()
lstNomeLog.Visible = False
End Sub

Private Sub mskCEP_GotFocus()
mskCEP.SelStart = 0
mskCEP.SelLength = Len(mskCEP.Text)
End Sub

Private Sub mskCEP2_GotFocus()
mskCEP2.SelStart = 0
mskCEP2.SelLength = Len(mskCEP2.Text)
End Sub

Private Sub mskCNPJ_GotFocus()
mskCNPJ.SelStart = 0
mskCNPJ.SelLength = Len(mskCNPJ.Text)
'mskCNPJ.SetFocus
End Sub

Private Sub mskCPF_GotFocus()
mskCPF.SelStart = 0
mskCPF.SelLength = Len(mskCPF.Text)
'mskCPF.SetFocus
End Sub

Private Sub mskDtNascto_GotFocus()
mskDtNascto.SelStart = 0
mskDtNascto.SelLength = Len(mskDtNascto.Text)
mskDtNascto.SetFocus
End Sub

Private Sub mskRG_GotFocus()
mskRG.SelStart = 0
mskRG.SelLength = Len(mskRG.Text)
'mskRG.SetFocus
End Sub

Private Sub txtCompl_GotFocus()
txtCompl.SelStart = 0
txtCompl.SelLength = Len(txtCompl)
End Sub

Private Sub txtCompl2_GotFocus()
txtCompl2.SelStart = 0
txtCompl2.SelLength = Len(txtCompl2)
End Sub

Private Sub txtEmail_GotFocus()
txtEmail.SelStart = 0
txtEmail.SelLength = Len(txtEmail)
End Sub

Private Sub txtFone_GotFocus()
txtFone.SelStart = 0
txtFone.SelLength = Len(txtFone)
End Sub

Private Sub txtNome_GotFocus()
txtNome.SelStart = 0
txtNome.SelLength = Len(txtNome)
End Sub

Private Sub txtNum_GotFocus()
txtNum.SelStart = 0
txtNum.SelLength = Len(txtNum)
End Sub

Private Sub txtNum2_GotFocus()
txtNum2.SelStart = 0
txtNum2.SelLength = Len(txtNum2)
End Sub

Private Sub txtNum_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 44 Or KeyAscii = 45) Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtNum2_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 44 Or KeyAscii = 45) Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtNum_LostFocus()
Dim s As String

LimpaMascara mskCEP
If cmbCidade.Text <> "JABOTICABAL" Then
    Exit Sub
End If


If txtNomeLog.Text <> "" Then
    s = RetornaCEP(Val(txtNumLog.Text), Val(txtNum.Text))
    mskCEP.Text = s
End If

End Sub

Private Sub txtNum2_LostFocus()
Dim s As String

LimpaMascara mskCEP2
If cmbCidade2.Text <> "JABOTICABAL" Then
    Exit Sub
End If

If txtNomeLog2.Text <> "" Then
    s = RetornaCEP(Val(txtNumLog2.Text), Val(txtNum2.Text))
    mskCEP2.Text = s
End If

End Sub

Private Sub Grava()

Dim MaxCod As Long, nBairro As Integer, nCidade As Integer, nPessoa As Integer, nBairro2 As Integer, nCidade2 As Integer, sEtiq As String
Dim RdoAux2 As rdoResultset, nProf As Integer, aObs() As String, x As Integer
ReDim aObs(0)
Sql = "SELECT MAX(CODCIDADAO) AS MAXIMO FROM CIDADAO WHERE CODCIDADAO<700000"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!maximo) Then
    MaxCod = 1
Else
    MaxCod = RdoAux!maximo + 1
End If
RdoAux.Close

If OptP(0).value = True Then
    nPessoa = 0
Else
    nPessoa = 1
End If

If cmbBairro.ListIndex > -1 Then
    nBairro = cmbBairro.ItemData(cmbBairro.ListIndex)
Else
    nBairro = 0
End If
If cmbCidade.ListIndex > -1 Then
    nCidade = cmbCidade.ItemData(cmbCidade.ListIndex)
Else
    nCidade = 0
End If

If cmbBairro2.ListIndex > -1 Then
    nBairro2 = cmbBairro2.ItemData(cmbBairro2.ListIndex)
Else
    nBairro2 = 0
End If
If cmbCidade2.ListIndex > -1 Then
    nCidade2 = cmbCidade2.ItemData(cmbCidade2.ListIndex)
Else
    nCidade2 = 0
End If


If Evento = "Novo" Then
    Sql = "INSERT CIDADAO(CODCIDADAO,NOMECIDADAO,CPF,CNPJ,CODLOGRADOURO,NUMIMOVEL,COMPLEMENTO,CODBAIRRO,CODCIDADE,"
    Sql = Sql & "SIGLAUF,CEP,TELEFONE,EMAIL,RG,NOMELOGRADOURO,ORGAO,JURIDICA,CODPAIS,CODLOGRADOURO2,NUMIMOVEL2,COMPLEMENTO2,CODBAIRRO2,CODCIDADE2,"
    Sql = Sql & "SIGLAUF2,CEP2,NOMELOGRADOURO2,ETIQUETA,CODPAIS2,TELEFONE2,EMAIL2,ETIQUETA2,data_nascimento,codprofissao,temfone,temfone2,whatsapp,whatsapp2) VALUES(" & MaxCod & ",'" & Mask(txtNome.Text) & "','"
    Sql = Sql & IIf(mskCPF.ClipText = "", "", mskCPF.ClipText) & "','" & IIf(mskCNPJ.ClipText = "", "", mskCNPJ.ClipText) & "',"
    Sql = Sql & IIf(Val(txtNumLog.Text) > 0, Val(txtNumLog.Text), "Null") & "," & Val(txtNum.Text) & ",'" & Mask(txtCompl.Text) & "'," & IIf(nBairro > 0, nBairro, "Null") & ","
    Sql = Sql & IIf(nCidade > 0, nCidade, "Null") & ",'" & Left$(cmbUF.Text, 2) & "'," & IIf(Trim(mskCEP.ClipText) = "", "Null", "'" & mskCEP.ClipText & "'") & ",'" & Mask(txtFone.Text) & "','"
    Sql = Sql & Mask(txtEmail.Text) & "','" & IIf(mskRG.ClipText = "", "", Mask(mskRG.ClipText)) & "'," & IIf(Val(txtNumLog.Text) > 0, "Null", "'" & txtNomeLog.Text & "'") & ",'"
    Sql = Sql & Mask(txtOrgao.Text) & "'," & nPessoa & "," & IIf(txtPais.Tag = "", 1, Val(txtPais.Tag)) & "," & IIf(Val(txtNumLog2.Text) > 0, Val(txtNumLog2.Text), "Null") & "," & Val(txtNum2.Text) & ",'"
    Sql = Sql & Mask(txtCompl2.Text) & "'," & IIf(nBairro2 > 0, nBairro2, "Null") & "," & IIf(nCidade2 > 0, nCidade2, "Null") & ",'"
    Sql = Sql & Left$(cmbUF2.Text, 2) & "'," & IIf(Trim(mskCEP2.ClipText) = "", "Null", "'" & mskCEP2.ClipText & "'") & "," & IIf(Val(txtNumLog2.Text) > 0, "Null", "'" & txtNomeLog2.Text & "'") & ",'"
    Sql = Sql & IIf(chkEtiq.value = vbChecked, "S", "N") & "'," & IIf(txtPais2.Tag = "", 1, Val(txtPais2.Tag)) & ",'" & Mask(txtFone2.Text) & "','" & Mask(txtEmail2.Text) & "','" & IIf(chkEtiq2.value = vbChecked, "S", "N") & "',"
    Sql = Sql & IIf(IsDate(mskDtNascto.Text), "'" & Format(mskDtNascto.Text, "mm/dd/yyyy") & "'", "Null") & "," & Val(txtProfissao.Tag) & "," & IIf(chkFone1.value = vbChecked, 1, 0) & ","
    Sql = Sql & IIf(chkFone2.value = vbChecked, 1, 0) & "," & IIf(chkWhatsApp.value = vbChecked, 1, 0) & "," & IIf(chkWhatsApp2.value = vbChecked, 1, 0) & ")"
Else
    Sql = "UPDATE CIDADAO SET NOMECIDADAO='" & Mask(txtNome.Text) & "',CPF='" & IIf(mskCPF.ClipText = "", "", mskCPF.ClipText) & "',"
    Sql = Sql & "CNPJ='" & IIf(mskCNPJ.ClipText = "", "", mskCNPJ.ClipText) & "',CODLOGRADOURO=" & IIf(Val(txtNumLog.Text) > 0, Val(txtNumLog.Text), "Null") & ","
    Sql = Sql & "NUMIMOVEL=" & Val(txtNum.Text) & ",COMPLEMENTO='" & Mask(txtCompl.Text) & "',CODBAIRRO=" & IIf(nBairro > 0, nBairro, "Null") & ","
    Sql = Sql & "CODCIDADE=" & IIf(nCidade > 0, nCidade, "Null") & ",SIGLAUF='" & Left$(cmbUF.Text, 2) & "',"
    Sql = Sql & "CEP=" & IIf(Trim(mskCEP.ClipText) = "", "Null", "'" & mskCEP.ClipText & "'") & ",TELEFONE='" & Mask(txtFone.Text) & "',"
    Sql = Sql & "EMAIL='" & Mask(txtEmail.Text) & "',RG='" & IIf(mskRG.ClipText = "", "", Mask(mskRG.ClipText)) & "',NOMELOGRADOURO=" & IIf(Val(txtNumLog.Text) > 0, "Null", "'" & txtNomeLog.Text & "'") & ","
    Sql = Sql & "ORGAO='" & Mask(txtOrgao.Text) & "',JURIDICA=" & nPessoa & ",CODPAIS=" & IIf(txtPais.Tag = "", 1, Val(txtPais.Tag)) & ",CODLOGRADOURO2=" & IIf(Val(txtNumLog2.Text) > 0, Val(txtNumLog2.Text), "Null") & ","
    Sql = Sql & "NUMIMOVEL2=" & Val(txtNum2.Text) & ",COMPLEMENTO2='" & Mask(txtCompl2.Text) & "',CODBAIRRO2=" & IIf(nBairro2 > 0, nBairro2, "Null") & ",CODCIDADE2=" & IIf(nCidade2 > 0, nCidade2, "Null") & ","
    Sql = Sql & "SIGLAUF2='" & Left$(cmbUF2.Text, 2) & "',CEP2=" & IIf(Trim(mskCEP2.ClipText) = "", "Null", "'" & mskCEP2.ClipText & "'") & ",NOMELOGRADOURO2=" & IIf(Val(txtNumLog2.Text) > 0, "Null", "'" & txtNomeLog2.Text & "'") & ","
    Sql = Sql & "ETIQUETA='" & IIf(chkEtiq.value = vbChecked, "S", "N") & "',CODPAIS2=" & IIf(txtPais2.Tag = "", 1, Val(txtPais2.Tag)) & ",TELEFONE2='" & Mask(txtFone2.Text) & "',EMAIL2='" & Mask(txtEmail2.Text) & "',ETIQUETA2='" & IIf(chkEtiq2.value = vbChecked, "S", "N") & "',"
    Sql = Sql & "data_nascimento=" & IIf(IsDate(mskDtNascto.Text), "'" & Format(mskDtNascto.Text, "mm/dd/yyyy") & "'", "Null") & ",codPROFISSAO=" & IIf(Val(SubNull(txtProfissao.Tag)) = 0, 1, Val(txtProfissao.Tag)) & ",TEMFONE=" & IIf(chkFone1.value = vbChecked, 1, 0) & ","
    Sql = Sql & "TEMFONE2=" & IIf(chkFone2.value = vbChecked, 1, 0) & ",WHATSAPP=" & IIf(chkWhatsApp.value = vbChecked, 1, 0) & ",WHATSAPP2=" & IIf(chkWhatsApp2.value = vbChecked, 1, 0) & " Where CodCidadao = " & Val(txtCod.Text)
End If
 cn.Execute Sql, rdExecDirect

If Evento = "Novo" Then
   txtCod.Text = MaxCod
   
'   Sql = "insert historicocidadao(codigo,data,usuario,obs) values(" & MaxCod & ",'" & Format(Now, sDataFormat & " hh:mm:ss") & "','" & NomeDeLogin & "','"
'   Sql = Sql & "Cidad�o criado atrav�s da tela de cadastro de cidad�o')"
   Sql = "insert historicocidadao(codigo,data,userid,obs) values(" & MaxCod & ",'" & Format(Now, sDataFormat & " hh:mm:ss") & "'," & RetornaUsuarioID(NomeDeLogin) & ",'"
   Sql = Sql & "Cidad�o criado atrav�s da tela de cadastro de cidad�o')"
   cn.Execute Sql, rdExecDirect
   
Else
    MaxCod = Val(txtCod.Text)
    With aReg(0)
        If .sBairroC <> cmbBairro2.Text Then
            ReDim Preserve aObs(UBound(aObs) + 1)
            aObs(UBound(aObs)) = "Alterado bairro end. comercial de " & .sBairroC & " para " & cmbBairro2.Text
        End If
        If .sBairroR <> cmbBairro.Text Then
            ReDim Preserve aObs(UBound(aObs) + 1)
            aObs(UBound(aObs)) = "Alterado bairro end. residencial de " & .sBairroR & " para " & cmbBairro.Text
        End If
        If .sCepC <> mskCEP2.Text Then
            ReDim Preserve aObs(UBound(aObs) + 1)
            aObs(UBound(aObs)) = "Alterado cep end. comercial de " & .sCepC & " para " & mskCEP2.Text
        End If
        If .sCepR <> mskCEP.Text Then
            ReDim Preserve aObs(UBound(aObs) + 1)
            aObs(UBound(aObs)) = "Alterado cep end. residencial de " & .sCepR & " para " & mskCEP.Text
        End If
        If .sCidadeC <> cmbCidade2.Text Then
            ReDim Preserve aObs(UBound(aObs) + 1)
            aObs(UBound(aObs)) = "Alterado cidade end. comercial de " & .sCidadeC & " para " & cmbCidade2.Text
        End If
        If .sCidadeR <> cmbCidade.Text Then
            ReDim Preserve aObs(UBound(aObs) + 1)
            aObs(UBound(aObs)) = "Alterado cidade end. residencial de " & .sCidadeR & " para " & cmbCidade.Text
        End If
        If .sCNPJ <> mskCNPJ.Text Then
            ReDim Preserve aObs(UBound(aObs) + 1)
            aObs(UBound(aObs)) = "Alterado CNPJ de " & .sCNPJ & " para " & mskCNPJ.Text
        End If
        If .sComplC <> txtCompl2.Text Then
            ReDim Preserve aObs(UBound(aObs) + 1)
            aObs(UBound(aObs)) = "Alterado compl. end. comercial de " & .sComplC & " para " & txtCompl2.Text
        End If
        If .sComplR <> txtCompl.Text Then
            ReDim Preserve aObs(UBound(aObs) + 1)
            aObs(UBound(aObs)) = "Alterado compl. end. residencial de " & .sComplR & " para " & txtCompl.Text
        End If
        If .sCPF <> mskCPF.Text Then
            ReDim Preserve aObs(UBound(aObs) + 1)
            aObs(UBound(aObs)) = "Alterado CPF de " & .sCPF & " para " & mskCPF.Text
        End If
        If .sDataNascto <> mskDtNascto.Text Then
            ReDim Preserve aObs(UBound(aObs) + 1)
            aObs(UBound(aObs)) = "Alterado Dt.Nascto de " & .sDataNascto & " para " & mskDtNascto.Text
        End If
        If .sEmailC <> txtEmail2.Text Then
            ReDim Preserve aObs(UBound(aObs) + 1)
            aObs(UBound(aObs)) = "Alterado Email end. comercial de " & .sEmailC & " para " & txtEmail2.Text
        End If
        If .sEmailR <> txtEmail.Text Then
            ReDim Preserve aObs(UBound(aObs) + 1)
            aObs(UBound(aObs)) = "Alterado Email end. residencial de " & .sEmailR & " para " & txtEmail.Text
        End If
        If .sFoneC <> txtFone2.Text Then
            ReDim Preserve aObs(UBound(aObs) + 1)
            aObs(UBound(aObs)) = "Alterado Fone end. comercial de " & .sFoneC & " para " & txtFone2.Text
        End If
        If .sFoneR <> txtFone.Text Then
            ReDim Preserve aObs(UBound(aObs) + 1)
            aObs(UBound(aObs)) = "Alterado Fone end. residencial de " & .sFoneR & " para " & txtFone.Text
        End If
        If .sNome <> txtNome.Text Then
            ReDim Preserve aObs(UBound(aObs) + 1)
            aObs(UBound(aObs)) = "Alterado nome de " & .sNome & " para " & txtNome.Text
        End If
        If .sNomeLogC <> txtNomeLog2.Text Then
            ReDim Preserve aObs(UBound(aObs) + 1)
            aObs(UBound(aObs)) = "Alterado logradouro end. comercial de " & .sNomeLogC & " para " & txtNomeLog2.Text
        End If
        If .sNomeLogr <> txtNomeLog.Text Then
            ReDim Preserve aObs(UBound(aObs) + 1)
            aObs(UBound(aObs)) = "Alterado logradouro end. residencial de " & .sNomeLogr & " para " & txtNomeLog.Text
        End If
        If .sNumeroC <> txtNum2.Text Then
            ReDim Preserve aObs(UBound(aObs) + 1)
            aObs(UBound(aObs)) = "Alterado numero end. comercial de " & .sNumeroC & " para " & txtNum2.Text
        End If
        If .sNumeroR <> txtNum.Text Then
            ReDim Preserve aObs(UBound(aObs) + 1)
            aObs(UBound(aObs)) = "Alterado numero end. residencial de " & .sNumeroR & " para " & txtNum.Text
        End If
        If .sPaisC <> txtPais2.Text Then
            ReDim Preserve aObs(UBound(aObs) + 1)
            aObs(UBound(aObs)) = "Alterado pa�s end. comercial de " & .sPaisC & " para " & txtPais2.Text
        End If
        If .sPaisR <> txtPais.Text Then
            ReDim Preserve aObs(UBound(aObs) + 1)
            aObs(UBound(aObs)) = "Alterado pa�s end. residencial de " & .sPaisR & " para " & txtPais.Text
        End If
        If .sProfissao <> txtProfissao.Text Then
            ReDim Preserve aObs(UBound(aObs) + 1)
            aObs(UBound(aObs)) = "Alterado profiss�o de " & .sProfissao & " para " & txtProfissao.Text
        End If
        If .sRG <> mskRG.Text Then
            ReDim Preserve aObs(UBound(aObs) + 1)
            aObs(UBound(aObs)) = "Alterado RG de " & .sRG & " para " & mskRG.Text
        End If
        If .sUFC <> cmbUF2.Text Then
            ReDim Preserve aObs(UBound(aObs) + 1)
            aObs(UBound(aObs)) = "Alterado UF end. comercial de " & .sUFC & " para " & cmbUF2.Text
        End If
        If .sUFR <> cmbUF.Text Then
            ReDim Preserve aObs(UBound(aObs) + 1)
            aObs(UBound(aObs)) = "Alterado UF end. residencial de " & .sUFR & " para " & cmbUF.Text
        End If
        If .bTemFoneR <> IIf(chkFone1.value = 1, 1, 0) Then
            ReDim Preserve aObs(UBound(aObs) + 1)
            aObs(UBound(aObs)) = "Alterado op��o de tem telefone residencial de " & IIf(.bTemFoneR, "Marcado", "Desmarcado") & " para " & IIf(chkFone1.value = vbChecked, "Marcado", "Desmarcado")
        End If
        If .bTemFoneC <> IIf(chkFone2.value = 1, 1, 0) Then
            ReDim Preserve aObs(UBound(aObs) + 1)
            aObs(UBound(aObs)) = "Alterado op��o de tem telefone comercial de " & IIf(.bTemFoneC, "Marcado", "Desmarcado") & " para " & IIf(chkFone2.value = vbChecked, "Marcado", "Desmarcado")
        End If
        For x = 1 To UBound(aObs)
            'Sql = "insert historicocidadao(codigo,data,usuario,obs) values(" & MaxCod & ",'" & Format(Now, sDataFormat & " hh:mm:ss") & "','" & NomeDeLogin & "','" & aObs(x) & "')"
            Sql = "insert historicocidadao(codigo,data,userid,obs) values(" & MaxCod & ",'" & Format(Now, sDataFormat & " hh:mm:ss") & "'," & RetornaUsuarioID(NomeDeLogin) & ",'" & Mask(aObs(x)) & "')"
            cn.Execute Sql, rdExecDirect
        Next
    End With
End If



'*** INTEGRA��O EICON ****
Sql = "SELECT * FROM mobiliarioproprietario Where mobiliarioproprietario.codcidadao = " & MaxCod
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    Sql = "select codigo from eicon_socio where codigo=" & MaxCod
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux2.RowCount = 0 Then
        Sql = "insert eicon_socio(codigo) values(" & MaxCod & ")"
        cn.Execute Sql, rdExecDirect
        AtualizaSocio
    End If
    RdoAux2.Close
End If
RdoAux.Close

End Sub

Private Sub FormHagana()

evNew = 2
evEdit = 3
evDel = 4

If InStr(1, sRet, Format(evNew, "000"), vbBinaryCompare) > 0 Then bNew = True
If InStr(1, sRet, Format(evEdit, "000"), vbBinaryCompare) > 0 Then bEdit = True
If InStr(1, sRet, Format(evDel, "000"), vbBinaryCompare) > 0 Then bDel = True

If Not bNew Then cmdNovo.Enabled = False
If Not bEdit Then cmdAlterar.Enabled = False
If Not bDel Then cmdExcluir.Enabled = False

End Sub

Private Sub txtNomeLog_Change()
If Trim$(txtNomeLog) = "" Then
   txtNumLog.Text = 0
End If
End Sub

Private Sub txtNomeLog2_Change()
If Trim$(txtNomeLog2) = "" Then
   txtNumLog2.Text = 0
End If
End Sub

Private Sub txtNomeLog_GotFocus()
txtNomeLog.SelStart = 0
txtNomeLog.SelLength = Len(txtNomeLog.Text)
End Sub

Private Sub txtNomeLog2_GotFocus()
txtNomeLog2.SelStart = 0
txtNomeLog2.SelLength = Len(txtNomeLog2)
End Sub

Private Sub txtNomeLog_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   lstNomeLog.Clear
   If txtNomeLog.Text <> "" Then
      Sql = "select codlogradouro,endereco from logradouro where endereco like '%" & Trim$(txtNomeLog.Text) & "%'  order by endereco"
      Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
      With RdoAux
          If .RowCount > 0 Then
             Do Until .EOF
                lstNomeLog.AddItem !Endereco
                lstNomeLog.ItemData(lstNomeLog.NewIndex) = !CodLogradouro
               .MoveNext
             Loop
             lstNomeLog.Left = 2250
             lstNomeLog.Top = 1800
             sEnd = "R"
             lstNomeLog.Visible = True
             lstNomeLog.ZOrder 0
             lstNomeLog.ListIndex = 0
             lstNomeLog.SetFocus
          Else
             MsgBox "Logradouro n�o encontrado.", vbInformation, "Aten��o"
             lstNomeLog.Visible = False
             txtNomeLog.SetFocus
          End If
      End With
   End If
Else
   txtNumLog.Text = 0
End If

End Sub

Private Sub txtNomeLog2_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   lstNomeLog.Clear
   If txtNomeLog2.Text <> "" Then
      Sql = "select codlogradouro,endereco from logradouro where endereco like '%" & Trim$(txtNomeLog2.Text) & "%' order by endereco"
      Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
      With RdoAux
          If .RowCount > 0 Then
             Do Until .EOF
                lstNomeLog.AddItem !Endereco
                lstNomeLog.ItemData(lstNomeLog.NewIndex) = !CodLogradouro
               .MoveNext
             Loop
             lstNomeLog.Left = 2250
             lstNomeLog.Top = 3510
             sEnd = "C"
             lstNomeLog.Visible = True
             lstNomeLog.ZOrder 0
             lstNomeLog.ListIndex = 0
             lstNomeLog.SetFocus
          Else
             MsgBox "Logradouro n�o encontrado.", vbInformation, "Aten��o"
             lstNomeLog.Visible = False
             txtNomeLog2.SetFocus
          End If
      End With
   End If
Else
   txtNumLog2.Text = 0
End If

End Sub

Private Sub lstNomeLog_DblClick()
If lstNomeLog.ListIndex > -1 Then
    If sEnd = "R" Then
        txtNumLog.Text = lstNomeLog.ItemData(lstNomeLog.ListIndex)
        txtNumLog_LostFocus
        txtNum.SetFocus
    Else
        txtNumLog2.Text = lstNomeLog.ItemData(lstNomeLog.ListIndex)
        txtNumLog2_LostFocus
        txtNum2.SetFocus
    End If
End If

lstNomeLog.Visible = False
End Sub

Private Sub lstNomeLog_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If lstNomeLog.ListIndex > -1 Then
        If sEnd = "R" Then
            txtNumLog.Text = lstNomeLog.ItemData(lstNomeLog.ListIndex)
            txtNumLog_LostFocus
            txtNum.SetFocus
        Else
            txtNumLog2.Text = lstNomeLog.ItemData(lstNomeLog.ListIndex)
            txtNumLog2_LostFocus
            txtNum2.SetFocus
        End If
    End If
    lstNomeLog.Visible = False
ElseIf KeyAscii = vbKeyEscape Then
   lstNomeLog.Visible = False
End If

End Sub

Private Sub txtNumLog_LostFocus()
If Val(txtNumLog.Text) > 0 Then
   Sql = "select codlogradouro,endereco from logradouro where codlogradouro=" & Val(txtNumLog.Text)
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   With RdoAux
       If .RowCount > 0 Then
          txtNomeLog.Text = !Endereco
       Else
          txtNomeLog.Text = ""
          MsgBox "Logradouro n�o cadastrado.", vbExclamation, "Aten��o"
          txtNumLog.SetFocus
       End If
      .Close
   End With
End If

End Sub

Private Sub txtNumLog2_LostFocus()
If Val(txtNumLog2.Text) > 0 Then
   Sql = "select codlogradouro,endereco from logradouro where codlogradouro=" & Val(txtNumLog2.Text)
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   With RdoAux
       If .RowCount > 0 Then
          txtNomeLog2.Text = !Endereco
       Else
          txtNomeLog2.Text = ""
          MsgBox "Logradouro n�o cadastrado.", vbExclamation, "Aten��o"
          txtNumLog2.SetFocus
       End If
      .Close
   End With
End If

End Sub


