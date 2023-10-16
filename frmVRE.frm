VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form frmDadosVRE 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Integração Via Rápida Empresa - VRE - Redesim"
   ClientHeight    =   6660
   ClientLeft      =   4890
   ClientTop       =   3420
   ClientWidth     =   12930
   Icon            =   "frmVRE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   12930
   Begin Tributacao.jcFrames frTela 
      Height          =   6645
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   12915
      _ExtentX        =   22781
      _ExtentY        =   11721
      FillColor       =   14745599
      Style           =   4
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Lista de Empresas Importadas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   2
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.TextBox txtFilter 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   585
         Width           =   6135
      End
      Begin VB.ComboBox cmbFilter 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmVRE.frx":014A
         Left            =   1590
         List            =   "frmVRE.frx":015A
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   6165
         Width           =   2505
      End
      Begin vbAcceleratorSGrid6.vbalGrid grdMain 
         Height          =   4965
         Left            =   90
         TabIndex        =   0
         Top             =   1095
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   8758
         GridLines       =   -1  'True
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeaderDragReorderColumns=   0   'False
         HeaderHotTrack  =   0   'False
         HeaderFlat      =   -1  'True
         BorderStyle     =   2
         ScrollBarStyle  =   2
         Editable        =   -1  'True
         DisableIcons    =   -1  'True
         DrawFocusRectangle=   0   'False
         SelectionAlphaBlend=   -1  'True
      End
      Begin prjChameleon.chameleonButton cmdSair 
         Height          =   375
         Left            =   11340
         TabIndex        =   4
         ToolTipText     =   "Sair da Tela"
         Top             =   6165
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   661
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
         MICON           =   "frmVRE.frx":01AF
         PICN            =   "frmVRE.frx":01CB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdOpen 
         Height          =   375
         Left            =   9840
         TabIndex        =   5
         ToolTipText     =   "Detalhes da empresa"
         Top             =   6165
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Detalhes"
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
         MICON           =   "frmVRE.frx":0239
         PICN            =   "frmVRE.frx":0255
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdFilter 
         Height          =   315
         Left            =   7740
         TabIndex        =   2
         ToolTipText     =   "Filtrar Dados"
         Top             =   585
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         BTYPE           =   5
         TX              =   "Filtrar"
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmVRE.frx":0340
         PICN            =   "frmVRE.frx":035C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdCancelFilter 
         Height          =   315
         Left            =   8820
         TabIndex        =   3
         ToolTipText     =   "Limpar Filtro"
         Top             =   585
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         BTYPE           =   5
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmVRE.frx":056F
         PICN            =   "frmVRE.frx":058B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar Dados...:"
         Height          =   240
         Left            =   270
         TabIndex        =   52
         Top             =   630
         Width           =   1230
      End
      Begin VB.Label lblCount 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4710
         TabIndex        =   15
         Top             =   5190
         Width           =   3825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar por:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   390
         TabIndex        =   6
         Top             =   6225
         Width           =   1185
      End
   End
   Begin Tributacao.jcFrames frTela 
      Height          =   6645
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   12915
      _ExtentX        =   22781
      _ExtentY        =   11721
      FillColor       =   14745599
      Style           =   4
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Detalhes da Empresa"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   2
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   45
         Left            =   165
         TabIndex        =   48
         Top             =   2400
         Width           =   12615
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   45
         Left            =   180
         TabIndex        =   30
         Top             =   1410
         Width           =   12615
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   45
         Left            =   180
         TabIndex        =   29
         Top             =   3000
         Width           =   12615
      End
      Begin Tributacao.jcFrames jcFrames1 
         Height          =   2205
         Left            =   240
         Top             =   4155
         Width           =   6825
         _ExtentX        =   12039
         _ExtentY        =   3889
         FillColor       =   14745599
         Style           =   4
         RoundedCornerTxtBox=   -1  'True
         Caption         =   "Lista de Eventos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ThemeColor      =   3
         ColorFrom       =   0
         ColorTo         =   0
         Begin VB.ListBox lstEvento 
            Appearance      =   0  'Flat
            Height          =   1590
            Left            =   90
            TabIndex        =   51
            Top             =   540
            Width           =   6630
         End
      End
      Begin prjChameleon.chameleonButton cmdVoltar1 
         Height          =   375
         Left            =   11160
         TabIndex        =   8
         ToolTipText     =   "Sair da Tela"
         Top             =   5970
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Voltar"
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
         MICON           =   "frmVRE.frx":07A8
         PICN            =   "frmVRE.frx":07C4
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
         Height          =   2205
         Left            =   7380
         Top             =   4155
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   3889
         FillColor       =   14745599
         Style           =   4
         RoundedCornerTxtBox=   -1  'True
         Caption         =   "Lista de Atividades"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ThemeColor      =   3
         ColorFrom       =   0
         ColorTo         =   0
         Begin MSComctlLib.ListView lvCnae 
            Height          =   1545
            Left            =   120
            TabIndex        =   28
            Top             =   540
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   2725
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "CNAE"
               Object.Width           =   1766
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Princ"
               Object.Width           =   1235
            EndProperty
         End
      End
      Begin prjChameleon.chameleonButton cmdGravar 
         Height          =   375
         Left            =   11160
         TabIndex        =   31
         ToolTipText     =   "Gravar os Dados"
         Top             =   5520
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   661
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
         MICON           =   "frmVRE.frx":0832
         PICN            =   "frmVRE.frx":084E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblCodImovel 
         BackStyle       =   0  'Transparent
         Caption         =   "CodImovel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   7110
         TabIndex        =   50
         Top             =   960
         Width           =   1605
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Código Imóvel......:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   28
         Left            =   5490
         TabIndex        =   49
         Top             =   960
         Width           =   1605
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Data deValidade.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   27
         Left            =   7260
         TabIndex        =   47
         Top             =   2610
         Width           =   1575
      End
      Begin VB.Label lblValidade 
         BackStyle       =   0  'Transparent
         Caption         =   "Validade"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   8820
         TabIndex        =   46
         Top             =   2610
         Width           =   1545
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Data de Emissão.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   26
         Left            =   4380
         TabIndex        =   45
         Top             =   2610
         Width           =   1575
      End
      Begin VB.Label lblDataEmissao 
         BackStyle       =   0  'Transparent
         Caption         =   "EndVRE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   6000
         TabIndex        =   44
         Top             =   2610
         Width           =   1545
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº do protocolo.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   25
         Left            =   240
         TabIndex        =   43
         Top             =   2610
         Width           =   1365
      End
      Begin VB.Label lblProtocolo 
         BackStyle       =   0  'Transparent
         Caption         =   "Protocolo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1680
         TabIndex        =   42
         Top             =   2610
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Área Imóvel.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   24
         Left            =   9780
         TabIndex        =   41
         Top             =   3195
         Width           =   1125
      End
      Begin VB.Label lblAreaImovel 
         BackStyle       =   0  'Transparent
         Caption         =   "AreaImovel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   10980
         TabIndex        =   40
         Top             =   3195
         Width           =   1125
      End
      Begin VB.Label lblAreaEstab 
         BackStyle       =   0  'Transparent
         Caption         =   "AreaEstab"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   8010
         TabIndex        =   39
         Top             =   3195
         Width           =   1125
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Área Estab:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   6990
         TabIndex        =   38
         Top             =   3195
         Width           =   1035
      End
      Begin VB.Label lblCodReduzido 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   10710
         TabIndex        =   37
         Top             =   660
         Width           =   1605
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Inscr.Municipal.....:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   9090
         TabIndex        =   36
         Top             =   660
         Width           =   1575
      End
      Begin VB.Label lblSituacao 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Aguardando"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   11010
         TabIndex        =   35
         Top             =   4650
         Width           =   1665
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Situação:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11340
         TabIndex        =   34
         Top             =   4260
         Width           =   1095
      End
      Begin VB.Label lblCidadeGTI 
         BackStyle       =   0  'Transparent
         Caption         =   "CidadeGTI"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   10350
         TabIndex        =   33
         Top             =   1560
         Width           =   2355
      End
      Begin VB.Label lblBairroGti 
         BackStyle       =   0  'Transparent
         Caption         =   "BairroGTI"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   6570
         TabIndex        =   32
         Top             =   1560
         Width           =   2445
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "CEP................:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   8400
         TabIndex        =   27
         Top             =   1995
         Width           =   1245
      End
      Begin VB.Label lblCep 
         BackStyle       =   0  'Transparent
         Caption         =   "Cep"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   9630
         TabIndex        =   26
         Top             =   1995
         Width           =   1065
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Número........:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   5340
         TabIndex        =   25
         Top             =   1995
         Width           =   1215
      End
      Begin VB.Label lblNumero 
         BackStyle       =   0  'Transparent
         Caption         =   "num"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   6570
         TabIndex        =   24
         Top             =   1995
         Width           =   1125
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Complemento.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   1320
         TabIndex        =   23
         Top             =   1995
         Width           =   1365
      End
      Begin VB.Label lblComplemento 
         BackStyle       =   0  'Transparent
         Caption         =   "Complemento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2820
         TabIndex        =   22
         Top             =   1995
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço GTI.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   21
         Top             =   1560
         Width           =   1365
      End
      Begin VB.Label lblEndGTI 
         BackStyle       =   0  'Transparent
         Caption         =   "EndGTI"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1650
         TabIndex        =   20
         Top             =   1560
         Width           =   3495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro GT.I...:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   5340
         TabIndex        =   19
         Top             =   1560
         Width           =   1155
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade GTI....:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   9120
         TabIndex        =   18
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "UF..:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   17
         Top             =   1995
         Width           =   435
      End
      Begin VB.Label lblUF 
         BackStyle       =   0  'Transparent
         Caption         =   "UF"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   720
         TabIndex        =   16
         Top             =   1995
         Width           =   345
      End
      Begin VB.Label lblDataAbertura 
         BackStyle       =   0  'Transparent
         Caption         =   "DataAbertura"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   10710
         TabIndex        =   14
         Top             =   970
         Width           =   1605
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Data de Abertura..:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   9090
         TabIndex        =   13
         Top             =   970
         Width           =   1575
      End
      Begin VB.Label lblCNPJ 
         BackStyle       =   0  'Transparent
         Caption         =   "CNPJ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1470
         TabIndex        =   12
         Top             =   970
         Width           =   2235
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de CNPJ..:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   970
         Width           =   1125
      End
      Begin VB.Label lblRazao 
         BackStyle       =   0  'Transparent
         Caption         =   "Razão Social:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1470
         TabIndex        =   10
         Top             =   660
         Width           =   7485
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Razão Social:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   660
         Width           =   1125
      End
   End
   Begin VB.Menu mnuOpcoes 
      Caption         =   "Opções"
      Visible         =   0   'False
      Begin VB.Menu mnuCriar 
         Caption         =   "Criar Empresa"
      End
      Begin VB.Menu mnuAlterar 
         Caption         =   "Alterar Empresa"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMarcar 
         Caption         =   "Marcar como cadastrada"
      End
      Begin VB.Menu mnuDesmarcar 
         Caption         =   "Marcar como aguardando"
      End
      Begin VB.Menu mnuCadVerificar 
         Caption         =   "Marcar como Cad\Verificar"
      End
   End
End
Attribute VB_Name = "frmDadosVRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type tLogradouro
    Codigo As Integer
    Nome As String
    NomeGTI As String
End Type
Dim aLogradouro() As tLogradouro, bOk As Boolean

Private Sub cmbFilter_Click()
CarregaLista
End Sub

Private Sub cmdCancelFilter_Click()
txtFilter.Text = ""
CarregaLista
End Sub

Private Sub cmdFilter_Click()
CarregaLista
End Sub

Private Sub cmdGravar_Click()
PopupMenu mnuOpcoes
End Sub

Private Sub cmdOpen_Click()

If grdMain.SelectedRow = 0 Then
    MsgBox "Selecione uma empresa!", vbExclamation, "Atenção"
    Exit Sub
End If

CarregaEmpresa grdMain.CellText(grdMain.SelectedRow, 1)
If bOk Then
    MudaTela 1
End If
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdVoltar1_Click()
MudaTela 0
End Sub

Private Sub Form_Load()

Centraliza Me
GridHeader
MudaTela 0
cmbFilter.ListIndex = 0
ReDim aLogradouro(0)
CarregaLogradouro

End Sub

Private Sub MudaTela(nTela As Integer)
Dim x As Integer

For x = 0 To 1
    If nTela <> x Then
        frTela(x).Visible = False
    Else
        frTela(x).Visible = True
    End If
Next

End Sub

Private Sub GridHeader()
With grdMain
    .GridFillLineColor = vbWhite
    .Editable = False
    .GridLines = True
    .HighlightBackColor = Marrom
    .HighlightForeColor = vbWhite
    .RowMode = True
    .DefaultRowHeight = 20
    .AddColumn "kId", "Protocolo", ecgHdrTextALignCentre, , 130
    .AddColumn "kRazao", "Razão Social", ecgHdrTextALignLeft, , 350
    .AddColumn "kCnpj", "CNPJ", ecgHdrTextALignCentre, , 110
    .AddColumn "kImp", "Dt.Licença", ecgHdrTextALignCentre, , 130
    .AddColumn "kArq", "Arquivo", ecgHdrTextALignCentre, , 0
    .AddColumn "kSit", "Situação", ecgHdrTextALignLeft, , 80
End With

End Sub

Private Sub CarregaListaOld()
Dim RdoAux As rdoResultset, Sql As String

grdMain.Clear
grdMain.Redraw = False
Sql = "select * from vre_empresa "
If cmbFilter.ListIndex = 1 Then
    Sql = Sql & " where situacao='Cadastrada'"
ElseIf cmbFilter.ListIndex = 2 Then
    Sql = Sql & " where situacao='Aguardando' or situacao is null "
ElseIf cmbFilter.ListIndex = 3 Then
    Sql = Sql & " where situacao='Cad\Verificar'"
End If
Sql = Sql & " order by razao_social"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    lblCount.Caption = .RowCount & " Empresas listadas"
    Do Until .EOF
        grdMain.AddRow
        grdMain.CellDetails grdMain.Rows, 1, !id, DT_CENTER
        grdMain.CellDetails grdMain.Rows, 2, !razao_social, DT_LEFT
        grdMain.CellDetails grdMain.Rows, 3, !Cnpj, DT_CENTER
        grdMain.CellDetails grdMain.Rows, 4, Format(!data_importacao, "dd/mm/yyyy hh:mm"), DT_CENTER
        grdMain.CellDetails grdMain.Rows, 5, !nome_arquivo, DT_LEFT
        If IsNull(!Situacao) Then
            grdMain.CellDetails grdMain.Rows, 6, "Aguardando", DT_LEFT, , , vbRed
        Else
            If !Situacao = "Aguardando" Then
                grdMain.CellDetails grdMain.Rows, 6, !Situacao, DT_LEFT, , , vbRed
            ElseIf !Situacao = "Cad\Verificar" Then
                grdMain.CellDetails grdMain.Rows, 6, !Situacao, DT_LEFT, , , vbBlue
            Else
                grdMain.CellDetails grdMain.Rows, 6, !Situacao, DT_LEFT, , , &H4000&
            End If
        End If
       .MoveNext
    Loop
   .Close
End With
grdMain.Redraw = True
If grdMain.Rows > 0 Then
    grdMain.SelectedRow = 1
End If
End Sub


Private Sub grdMain_ColumnClick(ByVal lCol As Long)
Dim sTag As String
Dim iSortIndex As Long
      On Error Resume Next
   With grdMain.SortObject
      
      ' This demo allows grouping.  When a column is clicked
      ' for sorting, we only want to remove any grouped rows:
      .ClearNongrouped
      
      ' See if this column is already in the sort object:
      iSortIndex = .IndexOf(lCol)
      If (iSortIndex = 0) Then
         ' If not, we add it:
         iSortIndex = .Count + 1
         .SortColumn(iSortIndex) = lCol
      End If
   
      ' Determine which sort order to apply:
      sTag = grdMain.ColumnTag(lCol)
      If (sTag = "") Then
         sTag = "DESC"
         .SortOrder(iSortIndex) = CCLOrderAscending
      Else
         sTag = ""
         .SortOrder(iSortIndex) = CCLOrderDescending
      End If
      grdMain.ColumnTag(lCol) = sTag
      
      ' Set the type of sorting:
      .SortType(iSortIndex) = grdMain.ColumnSortType(lCol)
   End With
   
   ' Do the sort:
   Screen.MousePointer = vbHourglass
   grdMain.Sort
   Screen.MousePointer = vbDefault
End Sub

Private Sub grdMain_DblClick(ByVal lRow As Long, ByVal lCol As Long)
If lRow > 0 Then
    cmdOpen_Click
End If
End Sub

Private Sub grdMain_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

If grdMain.SelectedRow = 0 Then
    MsgBox "Selecione uma empresa!", vbExclamation, "Atenção"
    Exit Sub
End If

If KeyCode = vbKeyReturn Then
    cmdOpen_Click
End If

End Sub

Private Sub CarregaEmpresaOld(nId As Long)
Dim Sql As String, RdoAux As rdoResultset, sPrefixo As String, sCRCF As String, sCRCJ As String, x As Integer
Dim sTipoLog As String, sTitLog As String, sNomeLog As String, RdoAux2 As rdoResultset, z As Long, sCodigo As String
Dim zNumero As Variant, zDataEmissao As Variant, zDataVencimento As Variant
bOk = True
If grdMain.CellText(grdMain.SelectedRow, 6) = "Aguardando" Then
    lblSituacao.Caption = "Aguardando"
    lblSituacao.ForeColor = vbRed
ElseIf grdMain.CellText(grdMain.SelectedRow, 6) = "Cad\Verificar" Then
    lblSituacao.Caption = "Cad\Verificar"
    lblSituacao.ForeColor = vbBlue
Else
    lblSituacao.Caption = "Cadastrada"
    lblSituacao.ForeColor = &H4000&

End If

lblCodReduzido.Caption = "000000"
lblCodResp.Caption = "000000"
lblEndGTI.Caption = ""
lblNomeContador.Tag = ""
lblNomeContador.Caption = ""
lblEndGTI.Tag = ""
z = SendMessage(lvSocio.HWND, LVM_DELETEALLITEMS, 0, 0)
z = SendMessage(lvCnae.HWND, LVM_DELETEALLITEMS, 0, 0)

Sql = "select * from vre_empresa where id=" & nId
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    lblCodReduzido.Caption = Format(!CODREDUZIDO, "000000")
    lblRazao.Caption = !razao_social
    lblCNPJ.Caption = Format(!Cnpj, "0#\.###\.###/####-##")
    If !tipo_registro = 1 Then
        lblIE.Caption = !numero_registro
    Else
        lblIE.Caption = "Isento"
    End If
    If Val(!setor_quadra_lote) < 40000 Then
        lblCodImovel.Caption = Format(!setor_quadra_lote, "000000")
    Else
        lblCodImovel.Caption = Format(0, "000000")
    End If
    lblDataAbertura.Caption = Format(!data_abertura, "dd/mm/yyyy")
    lblNomeResp.Caption = !nome_responsavel
    lblCPFResp.Caption = Format(!cpf_responsavel, "00#\.###\.###-##")
    lblFone.Caption = SubNull(!fone_contato1) & " " & SubNull(!fone_contato2)
    lblEmail.Caption = SubNull(!email_contato)
    lblEndVre.Caption = !tipo_logradouro & " " & !nome_logradouro
    lblBairroVre.Caption = !Bairro
    lblCidadeVRE.Caption = !Cidade
    lblUF.Caption = !UF
    lblAreaEstab.Caption = Format(!area_estabelecimento, "#0.00")
    lblAreaImovel.Caption = Format(!area_total, "#0.00")
    If Not IsNull(!cpf_responsavel) Then
        Sql = "select * from cidadao where cpf='" & !cpf_responsavel & "'"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            lblCodResp.Caption = Format(RdoAux2!CodCidadao, "000000")
        End If
        RdoAux2.Close
    End If
    
    For x = 1 To UBound(aLogradouro)
        If removeAcentos(UCase(aLogradouro(x).Nome)) = UCase(lblEndVre.Caption) Then
            lblEndGTI.Caption = aLogradouro(x).NomeGTI
            lblEndGTI.Tag = aLogradouro(x).Codigo
            Exit For
        End If
    Next
    
    lblCidadeGTI.Caption = ""
    Sql = "select * from cidade where siglauf='" & !UF & "'"
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
            Do Until .EOF
            If UCase(removeAcentos(!descCidade)) = UCase(RdoAux!Cidade) Then
                lblCidadeGTI.Caption = !descCidade
                lblCidadeGTI.Tag = !CodCidade
                Exit Do
            End If
           .MoveNext
        Loop
       .Close
    End With
    
    lblBairroGti.Caption = ""
    If lblCidadeGTI.Tag <> "" Then
        Sql = "select * from bairro where siglauf='" & !UF & "' and codcidade=" & Val(lblCidadeGTI.Tag)
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
                Do Until .EOF
                If UCase(removeAcentos(!DescBairro)) = UCase(RdoAux!Bairro) Then
                    lblBairroGti.Caption = !DescBairro
                    lblBairroGti.Tag = !CodBairro
                    Exit Do
                End If
               .MoveNext
            Loop
           .Close
        End With
    End If
    
    lblComplemento.Caption = SubNull(!Complemento)
    lblNumero.Caption = IIf(!numero_imovel = 0, "S/N", !numero_imovel)
    lblCep.Caption = Format(!Cep, "00000-000")
    
    If Val(!numero_crc_pf) = 0 Then
        sCRCF = ""
    Else
        sPrefixo = "1"
        sCRCF = sPrefixo & !uf_crc_pf & !numero_crc_pf & "/" & !tipo_crc_pf & "-" & !classif_crc_pf
    End If
    
    If Val(!numero_crc_pj) = 0 Then
        sCRCJ = ""
    Else
        sPrefixo = "1"
        sCRCJ = sPrefixo & !uf_crc_pj & !numero_crc_pj & "/" & !tipo_crc_pj & "-" & !classif_crc_pj
    End If
    
    lblCRCContadorPF.Caption = sCRCF
    lblCPFContador.Caption = Format(!cpf_contador, "00#\.###\.###-##")
    lblCRCContadorPJ.Caption = sCRCJ
    lblCnpjContador.Caption = Format(!cnpj_contador, "0#\.###\.###/####-##")
   
   
   If Val(!cnpj_contador) = 0 And Val(!cpf_contador) = 0 Then
        lblNomeContador.Caption = ""
        lblNomeContador.Tag = ""
   Else
        If Val(!cnpj_contador) > 0 Then
            Sql = "select * from escritoriocontabil where cnpj='" & !cnpj_contador & "' or crc='" & sCRCJ & "'"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount > 0 Then
                lblNomeContador.Caption = RdoAux2!NOMEESC
                lblNomeContador.Tag = CStr(RdoAux2!CODIGOESC)
            End If
            RdoAux2.Close
        End If
        If lblNomeContador.Caption = "" And (Val(!cpf_contador) > 0 Or sCrc <> "") Then
            Sql = "select * from escritoriocontabil where cpf='" & !cpf_contador & "' or crc='" & sCRCF & "'"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount > 0 Then
                lblNomeContador.Caption = RdoAux2!NOMEESC
                lblNomeContador.Tag = CStr(RdoAux2!CODIGOESC)
            End If
            RdoAux2.Close
        End If
    End If
   .Close
End With

Sql = "select * from vre_socio where id=" & nId
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Sql = "select * from cidadao where cpf='" & Format(Val(!Numero), "00000000000") & "'"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            sCodigo = Format(RdoAux2!CodCidadao, "000000")
        Else
            sCodigo = "000000"
        End If
        RdoAux2.Close
        Set itmX = lvSocio.ListItems.Add(, , sCodigo)
        itmX.SubItems(1) = !Nome
        itmX.SubItems(2) = !Numero
       .MoveNext
    Loop
   .Close
End With

Sql = "select * from vre_atividade where id=" & nId
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Set itmX = lvCnae.ListItems.Add(, , !Cnae)
        itmX.SubItems(1) = IIf(!principal, "Sim", "Não")
        itmX.SubItems(2) = IIf(!exercida, "Sim", "Não")
       .MoveNext
    Loop
   .Close
End With

'Sql = "select * from vre_licenciamento where empresa_id=" & nId & " and orgao=1"
Sql = "select * from vre_licenciamento where empresa_id=" & nId & " and orgao_id=287"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        If Val(!Numero) > 0 Then
            lblProtocolo.Caption = Format(RetornaNumero(!Numero), "0000000000000")
            lblProtocolo.Caption = Mid(lblProtocolo.Caption, 1, Len(lblProtocolo.Caption) - 6) & "." & Mid(lblProtocolo.Caption, 8, 4) & "-" & Right(lblProtocolo.Caption, 2)
            lblDataEmissao.Caption = Format(!Data_Emissao, "dd/mm/yyyy")
            lblValidade.Caption = Format(!Data_Vencimento, "dd/mm/yyyy")
        Else
            If MsgBox("Número de protocolo não encontrado, deseja informar agora?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
                zNumero = InputBox("Digite o número do protocolo", "Nº do protocolo")
                If Val(zNumero) = 0 Then
                    MsgBox "Nº de protocolo inválido", vbCritical, "Erro"
                    bOk = False
                Else
                    zDataEmissao = InputBox("Digite a data de emissão.", "Data de emissão")
                    If Not IsDate(zDataEmissao) Then
                        MsgBox "Data de emissão inválida", vbCritical, "Erro"
                        bOk = False
                    Else
                        zDataVencimento = InputBox("Digite a data de vencimento.", "Data de vencimento")
                        If Not IsDate(zDataVencimento) Then
                            MsgBox "Data de vencimento inválida", vbCritical, "Erro"
                            bOk = False
                        Else
                            Sql = "update vre_licenciamento set numero='" & zNumero & "',data_emissao='" & Format(zDataEmissao, "mm/dd/yyyy") & "',data_vencimento='" & Format(zDataVencimento, "mm/dd/yyyy") & "' "
                            Sql = Sql & "Where empresa_id=" & nId & " and orgao_id=1"
                            cn.Execute Sql, rdExecDirect
                        
                            lblProtocolo.Caption = zNumero
                            If InStr(1, lblProtocolo.Caption, ".") > 0 Then
                                lblProtocolo.Caption = Mid(lblProtocolo.Caption, 1, Len(lblProtocolo.Caption) - 6) & "." & Mid(lblProtocolo.Caption, 7, 4) & "-" & Right(lblProtocolo.Caption, 2)
                            End If
                            lblDataEmissao.Caption = zDataEmissao
                            lblValidade.Caption = zDataVencimento
                        End If
                    End If
                End If
            Else
                bOk = False
            End If
        End If
    End If
   .Close
End With

End Sub


Private Sub CarregaLogradouro()
Dim Sql As String, RdoAux As rdoResultset, sTipo As String, sTit As String, sLog As String, sNome As String

Sql = "select * from vwLOGRADOURO order by nomelogradouro"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        sTipo = Trim(SubNull(!NomeTipoLog))
        sTit = Trim(SubNull(!NomeTitLog))
        sLog = Trim(!NomeLogradouro)
        
        sNome = sTipo & " "
        sNome = sNome & IIf(sTit = "", "", sTit & " ")
        sNome = sNome & sLog
        
        ReDim Preserve aLogradouro(UBound(aLogradouro) + 1)
        aLogradouro(UBound(aLogradouro)).Codigo = !CodLogradouro
        aLogradouro(UBound(aLogradouro)).Nome = sNome
        aLogradouro(UBound(aLogradouro)).NomeGTI = !Logradouro
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub mnuAlterar_Click()
Dim RdoAux As rdoResultset, Sql As String, nId As Long

nId = grdMain.CellText(grdMain.SelectedRow, 1)
Sql = "select * from mobiliario where cnpj='" & RetornaNumero(lblCNPJ.Caption) & "' and dataencerramento is not null"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount = 0 Then
    MsgBox "Não localizado empresa cadastrada/ativa com este CNPJ.", vbCritical, "ERRO"
Else
    If MsgBox("Deseja alterar a empresa: " & RdoAux!codigomob & " - " & lblRazao.Caption & " ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
        AlterarEmpresa RdoAux!codigomob
    
        Sql = "update vre_empresa set situacao='Cadastrada',codreduzido=" & RdoAux!codigomob & " where id=" & nId
        cn.Execute Sql, rdExecDirect
        CarregaLista
        cmdVoltar1_Click
    
    End If
End If
RdoAux.Close

End Sub

Private Sub AlterarEmpresa(nCodigo As Long)
Dim RdoAux As rdoResultset, Sql As String, nAreaTLOld As Double, nAreaTLNew As Double, sFoneOld As String, sFoneNew As String, sEmailOld As String, sEmailNew As String
Dim nCodImovelOld As Long, nCodImovelNew As Long, sSilOld As String, sSilNew As String, bFind As Boolean, x As Integer, sDataEmissao As String, sDataValidade
Dim nCodLogradouroOld As Long, nCodLogradouroNew As Long, nNumImovelOld As Integer, nNumImovelNew As Integer, sNomeLogradouroOld As String, sNomeLogradouroNew As String

nAreaTLNew = CDbl(lblAreaEstab.Caption)
sFoneNew = lblFone.Caption
sEmailNew = lblEmail.Caption
nCodImovelNew = Val(lblCodImovel.Caption)
nCodLogradouroNew = Val(lblEndGTI.Tag)
nNumImovelNew = Val(lblNumero.Caption)
sNomeLogradouroNew = lblEndGTI.Caption

sSilNew = lblProtocolo.Caption & " Dt.Emissão: " & lblDataEmissao.Caption & " Área imóvel: " & lblAreaImovel.Caption & "m² Validade: " & lblValidade.Caption
bFind = False
Sql = "select * from vre_licenciamento where empresa_id=" & Val(grdMain.CellText(grdMain.SelectedRow, 1)) & " and orgao=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        sSilNew = !Numero
        sSilNew = Mid(sSilNew, 1, Len(sSilNew) - 6) & "." & Mid(sSilNew, 7, 4) & "-" & Right(sSilNew, 2)
        sDataEmissao = Format(!Data_Emissao, "dd/mm/yyyy")
        sDataValidade = Format(!Data_Vencimento, "dd/mm/yyyy")
       .MoveNext
    Loop
   .Close
End With

Sql = "select * from vwfullempresa where codigomob=" & nCodigo
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nAreaTLOld = !areatl
    sFoneOld = SubNull(!fonecontato)
    sEmailOld = SubNull(!emailcontato)
    nCodImovel = Val(SubNull(!Imovel))
    nCodLogradouroOld = Val(SubNull(!CodLogradouro))
    sNomeLogradouroOld = SubNull(!Logradouro2)
    nNumImovelOld = Val(SubNull(!Numero))
   .Close
End With

If nAreaTLNew <> nAreaTLOld Then
    Sql = "update mobiliario set areatl=" & Virg2Ponto(CStr(nAreaTLNew)) & " where codigomob=" & nCodigo
    cn.Execute Sql, rdExecDirect
    GravaHistorico nCodigo, "Alterado Área estabelecimento de " & nAreaTLOld & "m² para " & nAreaTLNew & "m²"
End If

If sFoneNew <> sFoneOld Then
    Sql = "update mobiliario set areatl=" & Virg2Ponto(CStr(nAreaTLNew)) & " where codigomob=" & nCodigo
    cn.Execute Sql, rdExecDirect
    GravaHistorico nCodigo, "Alterado Área estabelecimento de " & nAreaTLOld & "m² para " & nAreaTLNew & "m²"
End If

If nCodImovelNew <> nCodImovelOld Then
    Sql = "update mobiliario set imovel=" & Val(nCodImovelNew) & " where codigomob=" & nCodigo
    cn.Execute Sql, rdExecDirect
    GravaHistorico nCodigo, "Alterado Código do imóvel de " & nCodImovelOld & " para " & nCodImovelNew
End If

If nCodLogradouroNew > 0 Then
    If nCodLogradouroNew <> nCodLogradouroOld Then
        Sql = "update mobiliario set codlogradouro=" & nCodLogradouroNew & " where codigomob=" & nCodigo
        cn.Execute Sql, rdExecDirect
        GravaHistorico nCodigo, "Alterado logradouro do imóvel de " & sNomeLogradouroOld & " para " & sNomeLogradouroNew
    End If
End If

If nNumImovelNew > 0 Then
    If nNumImovelNew <> nNumImovelOld Then
        Sql = "update mobiliario set numero=" & nNumImovelNew & " where codigomob=" & nCodigo
        cn.Execute Sql, rdExecDirect
        GravaHistorico nCodigo, "Alterado nº do estabelecimento de " & nNumImovelOld & " para " & nNumImovelNew
    End If
End If


End Sub

Private Sub GravaHistorico(nCodigo As Long, sHistorico As String)
Dim RdoAux As rdoResultset, Sql As String, nSeq As Integer

Sql = "SELECT MAX(SEQ) AS MAXIMO FROM MOBILIARIOHIST WHERE CODMOBILIARIO=" & nCodigo
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!maximo) Then
    nSeq = 0
Else
    nSeq = RdoAux!maximo + 1
End If
            
'Sql = "INSERT MOBILIARIOHIST(CODMOBILIARIO,SEQ,DATAHIST,OBS,USUARIO) VALUES("
'Sql = Sql & nCodigo & "," & nSeq & ",'" & Format(Now, "mm/dd/yyyy") & "','" & Mask(sHistorico) & "','" & NomeDeLogin & "\VRE')"
Sql = "INSERT MOBILIARIOHIST(CODMOBILIARIO,SEQ,DATAHIST,OBS,USERID) VALUES("
Sql = Sql & nCodigo & "," & nSeq & ",'" & Format(Now, "mm/dd/yyyy") & "','" & Mask(sHistorico) & "'," & RetornaUsuarioID(NomeDeLogin) & ")"
cn.Execute Sql, rdExecDirect

End Sub


Private Sub mnuCadVerificar_Click()
Dim Sql As String, RdoAux As rdoResultset, nId As Long, sProc As String

If lblSituacao.Caption = "Cad\Verificar" Then
    MsgBox "A empresa já esta definida com a situação 'Cad\Verificar'", vbExclamation, "Atenção"
Else
    sProc = grdMain.CellText(grdMain.SelectedRow, 1)
    Sql = "update redesim_master set situacao='Cad\Verificar' where protocolo='" & sProc & "'"
'    nId = grdMain.CellText(grdMain.SelectedRow, 1)
'    Sql = "update vre_empresa set situacao='Cad\Verificar' where id=" & nId
    cn.Execute Sql, rdExecDirect
    CarregaLista
    cmdVoltar1_Click
End If

End Sub

Private Sub mnuDesmarcar_Click()
Dim Sql As String, RdoAux As rdoResultset, nId As Long, sProc As String

If lblSituacao.Caption = "Aguardando" Then
    MsgBox "A empresa já esta definida com a situação 'Aguardando'", vbExclamation, "Atenção"
Else
    sProc = grdMain.CellText(grdMain.SelectedRow, 1)
    Sql = "update redesim_master set situacao='Aguardando' where protocolo='" & sProc & "'"
'    nId = grdMain.CellText(grdMain.SelectedRow, 1)
'    Sql = "update vre_empresa set situacao='Aguardando' where id=" & nId
    cn.Execute Sql, rdExecDirect
    CarregaLista
    cmdVoltar1_Click
End If

End Sub

Private Sub mnuMarcar_Click()
Dim Sql As String, RdoAux As rdoResultset, nId As Long, sProc As String

If lblSituacao.Caption = "Cadastrada" Then
    MsgBox "A empresa já esta definida com a situação 'Cadastrada'", vbExclamation, "Atenção"
Else
    sProc = grdMain.CellText(grdMain.SelectedRow, 1)
    Sql = "update redesim_master set situacao='Cadastrada' where protocolo='" & sProc & "'"
'    nId = grdMain.CellText(grdMain.SelectedRow, 1)
'    Sql = "update vre_empresa set situacao='Cadastrada' where id=" & nId
    cn.Execute Sql, rdExecDirect
    CarregaLista
    cmdVoltar1_Click
End If

End Sub

Private Sub mnuCriar_Click()
Dim Sql As String, RdoAux As rdoResultset, nId As Long, MinCod As Long, MaxCod As Long, nCodReduz As Long, nCodCidadao As Long
Dim nCodLogradouro As Integer, nCodBairro As Integer, nCodCidade As Integer, x As Integer, sCnae As String, bPrincipal As Boolean, bExercida As Boolean, sSil As String, sDataTmp As String
Dim sDivisao As String, sGrupo As String, sClasse As String, sSubClasse As String, sAtividade As String, sDataEmissao As String, sDataValidade As String, sProtocolo As String
Dim sDDD_NF As String, sTelefone_NF As String, sFone As String, sProc As String

'sFone = RetornaNumero(lblFone.Caption)
'If IsNumeric(sFone) Then
'    sDDD_NF = Left(sFone, 2)
'    sTelefone_NF = Mid(sFone, 3, Len(sFone) - 2)
'End If

sProc = grdMain.CellText(grdMain.SelectedRow, 1)
sAtividade = ""
For x = 1 To lvCnae.ListItems.Count
    sCnae = Format(lvCnae.ListItems(x).Text, "0000000")
    Sql = "select descricao from cnae where cnae='" & sCnae & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount > 0 Then
        sAtividade = sAtividade & RdoAux!Descricao & ";"
    End If
Next
If sAtividade <> "" Then
    sAtividade = Chomp(sAtividade, chomp_righT, 1)
End If

sProtocolo = ""
'Sql = "select * from vre_licenciamento where empresa_id=" & nId & " and orgao=287"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    If .RowCount > 0 Then
'        sprotocolo = !Numero
'        sprotocolo = Mid(sprotocolo, 1, Len(sprotocolo) - 6) & "." & Mid(sprotocolo, 7, 4) & "-" & Right(sprotocolo, 2)
'        sDataEmissao = Format(!data_emissao, "dd/mm/yyyy")
'        sDataValidade = Format(!Data_Vencimento, "dd/mm/yyyy")
'    End If
 '  .Close
'End With
sProtocolo = lblProtocolo.Caption
sDataEmissao = lblDataEmissao.Caption
sDataValidade = lblValidade.Caption
If lblSituacao.Caption = "Cadastrada" Then
    MsgBox "A empresa já esta definida com a situação 'Cadastrada'" & vbCrLf & "Caso queira gravar no GTI uma nova emprea com estes dados, remova primeiro esta marcação.", vbExclamation, "Atenção"
    Exit Sub
End If
'sql = "select * from mobiliario where cnpj='" & RetornaNumero(lblCNPJ.Caption) & "' and dataencerramento is null"
Sql = "select * from mobiliario where cnpj='" & RetornaNumero(lblCNPJ.Caption) & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    MsgBox "Já existe uma inscrição com este Cnpj (" & RdoAux!codigomob & ")", vbCritical, "Erro"
    Exit Sub
End If

If MsgBox("Deseja criar um novo cadastro no GTI?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub

Sql = "select MAX(codigomob) as maximo FROM mobiliario WHERE codigomob<180000"
'Sql = "SELECT CODIGOMOB FROM MOBILIARIO WHERE CODIGOMOB>100000 and CODIGOMOB<200000 ORDER BY CODIGOMOB"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
MaxCod = RdoAux!maximo + 1
nCodReduz = MaxCod

If Val(lblEndGTI.Tag) > 0 Then
    nCodLogradouro = Val(lblEndGTI.Tag)
Else
    nCodLogradouro = 999
End If
If Val(lblBairroGti.Tag) > 0 Then
    nCodBairro = Val(lblBairroGti.Tag)
Else
    nCodBairro = 999
End If
If Val(lblCidadeGTI.Tag) > 0 Then
    nCodCidade = Val(lblCidadeGTI.Tag)
Else
    nCodCidade = 999
End If

If nCodCidade <> 999 And nCodCidade <> 413 Then
    MsgBox "Apenas empresas estabelecidas no município podem ser cadastradas.", vbCritical, "Erro"
    Exit Sub
End If

Sql = "insert mobiliario(codigomob,dvmob,razaosocial,cnpj,areatl,ativextenso,codlogradouro,numero,complemento,codbairro,codcidade,siglauf,cep,dataabertura,imovel) values("
Sql = Sql & nCodReduz & "," & RetornaDVCodReduzido(nCodReduz) & ",'" & Mask(lblRazao.Caption) & "','" & RetornaNumero(lblCNPJ.Caption) & "'," & Virg2Ponto(lblAreaEstab.Caption) & ",'" & UCase(Mask(sAtividade)) & "'," & nCodLogradouro & "," & Val(lblNumero.Caption) & ",'" & Mask(lblComplemento.Caption) & "',"
Sql = Sql & nCodBairro & "," & nCodCidade & ",'" & lblUF.Caption & "','" & RetornaNumero(lblCep.Caption) & "','" & Format(lblDataAbertura.Caption, "mm/dd/yyyy") & "'," & Val(lblCodImovel.Caption) & ")"
cn.Execute Sql, rdExecDirect

'For x = 1 To lvSocio.ListItems.Count
'    nCodCidadao = Val(lvSocio.ListItems(x).Text)
'    If nCodCidadao > 0 Then
'        Sql = "insert mobiliarioproprietario(codmobiliario,codcidadao) values(" & nCodReduz & "," & nCodCidadao & ")"
'        cn.Execute Sql, rdExecDirect
'    Else
'        Sql = "select max(codcidadao) as maximo from cidadao"
'        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'        MaxCod = RdoAux!maximo + 1
'        RdoAux.Close
        
'        Sql = "insert cidadao(codcidadao,nomecidadao,cpf) values(" & MaxCod & ",'" & lvSocio.ListItems(x).SubItems(1) & "','" & Format(lvSocio.ListItems(x).SubItems(2), "00000000000") & "')"
'        cn.Execute Sql, rdExecDirect
        
'        Sql = "insert mobiliarioproprietario(codmobiliario,codcidadao) values(" & nCodReduz & "," & MaxCod & ")"
'        cn.Execute Sql, rdExecDirect
        
'    End If
'Next

For x = 1 To lvCnae.ListItems.Count
    sCnae = Format(lvCnae.ListItems(x).Text, "0000000")
    bPrincipal = IIf(lvCnae.ListItems(x).SubItems(1) = "Sim", True, False)
'    bExercida = IIf(lvCnae.ListItems(x).SubItems(2) = "Sim", True, False)
    sDivisao = Left(sCnae, 2)
    sGrupo = Mid(sCnae, 3, 1)
    sClasse = Mid(sCnae, 4, 2)
    sSubClasse = Right(sCnae, 2)
    sCnae = sDivisao & sGrupo & Left$(sClasse, 1) & "-" & Right$(sClasse, 1) & "/" & sSubClasse
    On Error Resume Next
    Sql = "INSERT MOBILIARIOCNAE(CODMOBILIARIO,SECAO,DIVISAO,GRUPO,CLASSE,SUBCLASSE,PRINCIPAL,CNAE) VALUES("
    Sql = Sql & nCodReduz & ",''," & Val(sDivisao) & "," & Val(sGrupo) & "," & Val(sClasse) & "," & Val(sSubClasse) & "," & IIf(bPrincipal, 1, 0) & ",'" & sCnae & "')"
    cn.Execute Sql, rdExecDirect
    On Error GoTo 0
Next


'If sProtocolo <> "" Then
'    If Not IsDate(sDataEmissao) Then
ini1:
'        sSil = InputBox("Digite manualmente o nº de protocolo ", "Não foi possível carregar o protocolo")
'        If sSil = "" Then
'            GoTo ini1
'        End If
ini2:
'        sDataTmp = InputBox("Digite a data da emissão", "Não foi possível carregar a data da emissão")
'        If Not IsDate(sDataTmp) Then
'            MsgBox "Data inválida!"
'            GoTo ini2
'        End If
'        sDataEmissao = sDataTmp
ini3:
'        sDataTmp = InputBox("Digite a data da validade", "Não foi possível carregar a data da validade")
'        If Not IsDate(sDataTmp) Then
'            MsgBox "Data inválida!"
'            GoTo ini3
'        End If
'        sDataValidade = sDataTmp
        
'    Else
'        sSil = sSil & " Dt.Emissão: " & sDataEmissao & " Área imóvel: " & lblAreaImovel.Caption & "m² Validade: " & sDataValidade
'    End If
'    If sProtocolo <> "" Then
'        Sql = "insert sil(codigo,protocolo,data_emissao,data_validade,area_imovel) values(" & nCodReduz & ",'" & sProtocolo & "','" & Format(sDataEmissao, "mm/dd/yyyy") & "','" & Format(sDataValidade, "mm/dd/yyyy") & "'," & Virg2Ponto(lblAreaImovel.Caption) & ")"
'        cn.Execute Sql, rdExecDirect
'    End If
'End If

Sql = "update redesim_master set situacao='Cadastrada',inscricao=" & nCodReduz & " where protocolo='" & sProc & "'"
'Sql = "update vre_empresa set situacao='Cadastrada',codreduzido=" & nCodReduz & " where id=" & nId
cn.Execute Sql, rdExecDirect

MsgBox "A empresa foi criada com a inscrição municipal nº " & nCodReduz, vbInformation, "Informação"
CarregaLista
cmdVoltar1_Click

End Sub

Private Sub CarregaLista()
Dim RdoAux As rdoResultset, Sql As String, sFiltro As String

sFiltro = Trim(txtFilter.Text)
grdMain.Clear
grdMain.Redraw = False
lstEvento.Clear
Sql = "select * from redesim_master where 1=1 "
If cmbFilter.ListIndex = 1 Then
    Sql = Sql & " and situacao='Cadastrada'"
ElseIf cmbFilter.ListIndex = 2 Then
    Sql = Sql & " and situacao='Aguardando' or situacao is null "
ElseIf cmbFilter.ListIndex = 3 Then
    Sql = Sql & " and situacao='Cad\Verificar'"
End If
If sFiltro <> "" Then
    Sql = Sql & " and (protocolo like '%" & sFiltro & "%' or razao_social like '%" & sFiltro & "%' or cnpj like '%" & sFiltro & "%' )"
End If
Sql = Sql & " order by razao_social"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    lblCount.Caption = .RowCount & " Empresas listadas"
    Do Until .EOF
        grdMain.AddRow
        grdMain.CellDetails grdMain.Rows, 1, !Protocolo, DT_CENTER
        grdMain.CellDetails grdMain.Rows, 2, !razao_social, DT_LEFT
        grdMain.CellDetails grdMain.Rows, 3, !Cnpj, DT_CENTER
        grdMain.CellDetails grdMain.Rows, 4, Format(!data_licenca, "dd/mm/yyyy hh:mm"), DT_CENTER
        grdMain.CellDetails grdMain.Rows, 5, "", DT_LEFT
        If IsNull(!Situacao) Then
            grdMain.CellDetails grdMain.Rows, 6, "Aguardando", DT_LEFT, , , vbRed
        Else
            If !Situacao = "Aguardando" Then
                grdMain.CellDetails grdMain.Rows, 6, !Situacao, DT_LEFT, , , vbRed
            ElseIf !Situacao = "Cad\Verificar" Then
                grdMain.CellDetails grdMain.Rows, 6, !Situacao, DT_LEFT, , , vbBlue
            Else
                grdMain.CellDetails grdMain.Rows, 6, !Situacao, DT_LEFT, , , &H4000&
            End If
        End If
       .MoveNext
    Loop
   .Close
End With
grdMain.Redraw = True
If grdMain.Rows > 0 Then
    grdMain.SelectedRow = 1
End If
End Sub

Private Sub CarregaEmpresa(sProtocolo As String)
Dim Sql As String, RdoAux As rdoResultset, sPrefixo As String, sCRCF As String, sCRCJ As String, x As Integer
Dim sTipoLog As String, sTitLog As String, sNomeLog As String, RdoAux2 As rdoResultset, z As Long, sCodigo As String
Dim zNumero As Variant, zDataEmissao As Variant, zDataVencimento As Variant
bOk = True
If grdMain.CellText(grdMain.SelectedRow, 6) = "Aguardando" Then
    lblSituacao.Caption = "Aguardando"
    lblSituacao.ForeColor = vbRed
ElseIf grdMain.CellText(grdMain.SelectedRow, 6) = "Cad\Verificar" Then
    lblSituacao.Caption = "Cad\Verificar"
    lblSituacao.ForeColor = vbBlue
Else
    lblSituacao.Caption = "Cadastrada"
    lblSituacao.ForeColor = &H4000&

End If
lstEvento.Clear
lblCodReduzido.Caption = "000000"
'lblCodResp.Caption = "000000"
lblEndGTI.Caption = ""
'lblNomeContador.Tag = ""
'lblNomeContador.Caption = ""
lblEndGTI.Tag = ""
'z = SendMessage(lvSocio.HWND, LVM_DELETEALLITEMS, 0, 0)
z = SendMessage(lvCnae.HWND, LVM_DELETEALLITEMS, 0, 0)

Dim b As Bairro
Sql = "SELECT protocolo,data_licenca,data_validade,razao_social,cnpj,rm.logradouro,numero,rm.complemento,rm.cep,rm.cnae_principal,rm.numero_imovel,rm.area_imovel,rm.area_estabelecimento,rm.inscricao,rm.situacao,l.endereco_resumido FROM redesim_master rm "
Sql = Sql & "LEFT OUTER JOIN vwLOGRADOURO l ON rm.logradouro = codlogradouro where protocolo='" & sProtocolo & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If !Inscricao = 0 Then
        Sql = "select codigomob from mobiliario where cnpj='" & !Cnpj & "'"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            lblCodReduzido.Caption = Format(Val(SubNull(RdoAux2!codigomob)), "000000")
        End If
    Else
        lblCodReduzido.Caption = Format(Val(SubNull(!Inscricao)), "000000")
    End If
'    lblCodReduzido.Caption = Format(Val(SubNull(!Inscricao)), "000000")
    lblRazao.Caption = !razao_social
    lblCNPJ.Caption = Format(!Cnpj, "0#\.###\.###/####-##")
    lblDataAbertura.Caption = Format(!data_licenca, "dd/mm/yyyy")
    lblCodImovel.Caption = Format(Val(SubNull(!numero_imovel)), "000000")
    lblAreaEstab.Caption = Format(!area_estabelecimento, "#0.00")
    lblAreaImovel.Caption = Format(!area_imovel, "#0.00")
    lblUF.Caption = "SP"
    lblEndGTI.Caption = SubNull(!endereco_resumido)
    lblEndGTI.Tag = !Logradouro
    lblCidadeGTI.Caption = "JABOTICABAL"
    lblCidadeGTI.Tag = 413
    b = RetornaBairro(!Cep)
    lblBairroGti.Caption = SubNull(b.Nome)
    lblBairroGti.Tag = SubNull(b.Codigo)
    lblComplemento.Caption = SubNull(!Complemento)
    lblNumero.Caption = IIf(!Numero = 0, "S/N", !Numero)
    lblCep.Caption = Format(!Cep, "00000-000")
    lblProtocolo.Caption = !Protocolo
    lblDataEmissao = Format(!data_licenca, "dd/mm/yyyy")
    If IsNull(!Data_Validade) Then
        lblValidade.Caption = "N/D"
    Else
        lblValidade.Caption = Format(!Data_Validade, "dd/mm/yyyy")
    End If
    
    Set itmX = lvCnae.ListItems.Add(, , !Cnae_principal)
    itmX.SubItems(1) = "Sim"
   .Close
End With

Sql = "select * from redesim_cnae where protocolo='" & sProtocolo & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Set itmX = lvCnae.ListItems.Add(, , !Cnae)
        itmX.SubItems(1) = "Não"
       .MoveNext
    Loop
   .Close
End With

Sql = "SELECT protocolo,re.nome FROM redesim_registro_evento rr INNER JOIN redesim_evento re ON re.codigo=rr.evento WHERE protocolo='" & sProtocolo & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        lstEvento.AddItem !Nome
       .MoveNext
    Loop
   .Close
End With


'Sql = "select * from vre_empresa where id=" & nId
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    lblCodReduzido.Caption = Format(!CODREDUZIDO, "000000")
'    lblRazao.Caption = !razao_social
'    lblCNPJ.Caption = Format(!Cnpj, "0#\.###\.###/####-##")
'    If !tipo_registro = 1 Then
'        lblIE.Caption = !numero_registro
'    Else
'        lblIE.Caption = "Isento"
'    End If
'    If Val(!setor_quadra_lote) < 40000 Then
'        lblCodImovel.Caption = Format(!setor_quadra_lote, "000000")
'    Else
'        lblCodImovel.Caption = Format(0, "000000")
'    End If
'    lblDataAbertura.Caption = Format(!data_abertura, "dd/mm/yyyy")
'    lblNomeResp.Caption = !nome_responsavel
'    lblCPFResp.Caption = Format(!cpf_responsavel, "00#\.###\.###-##")
'    lblFone.Caption = SubNull(!fone_contato1) & " " & SubNull(!fone_contato2)
'    lblEmail.Caption = SubNull(!email_contato)
'    lblEndVre.Caption = !tipo_logradouro & " " & !nome_logradouro
'    lblBairroVre.Caption = !Bairro
'    lblCidadeVRE.Caption = !Cidade
'    lblUF.Caption = !UF
'    lblAreaEstab.Caption = Format(!area_estabelecimento, "#0.00")
'    lblAreaImovel.Caption = Format(!area_total, "#0.00")
'    If Not IsNull(!cpf_responsavel) Then
'        Sql = "select * from cidadao where cpf='" & !cpf_responsavel & "'"
'        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'        If RdoAux2.RowCount > 0 Then
'            lblCodResp.Caption = Format(RdoAux2!CodCidadao, "000000")
'        End If
'        RdoAux2.Close
'    End If
'
'    For x = 1 To UBound(aLogradouro)
'        If removeAcentos(UCase(aLogradouro(x).Nome)) = UCase(lblEndVre.Caption) Then
'            lblEndGTI.Caption = aLogradouro(x).NomeGTI
'            lblEndGTI.Tag = aLogradouro(x).Codigo
'            Exit For
'        End If
'    Next
'
'    lblCidadeGTI.Caption = ""
'    Sql = "select * from cidade where siglauf='" & !UF & "'"
'    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'    With RdoAux2
'            Do Until .EOF
'            If UCase(removeAcentos(!descCidade)) = UCase(RdoAux!Cidade) Then
'                lblCidadeGTI.Caption = !descCidade
'                lblCidadeGTI.Tag = !CodCidade
'                Exit Do
'            End If
'           .MoveNext
'        Loop
'       .Close
'    End With
'
'    lblBairroGti.Caption = ""
'    If lblCidadeGTI.Tag <> "" Then
'        Sql = "select * from bairro where siglauf='" & !UF & "' and codcidade=" & Val(lblCidadeGTI.Tag)
'        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'        With RdoAux2
'                Do Until .EOF
'                If UCase(removeAcentos(!DescBairro)) = UCase(RdoAux!Bairro) Then
'                    lblBairroGti.Caption = !DescBairro
'                    lblBairroGti.Tag = !CodBairro
'                    Exit Do
'                End If
'               .MoveNext
'            Loop
'           .Close
'        End With
'    End If
'
'    lblComplemento.Caption = SubNull(!Complemento)
'    lblNumero.Caption = IIf(!numero_imovel = 0, "S/N", !numero_imovel)
'    lblCep.Caption = Format(!Cep, "00000-000")
'
'    If Val(!numero_crc_pf) = 0 Then
'        sCRCF = ""
'    Else
'        sPrefixo = "1"
'        sCRCF = sPrefixo & !uf_crc_pf & !numero_crc_pf & "/" & !tipo_crc_pf & "-" & !classif_crc_pf
'    End If
'
'    If Val(!numero_crc_pj) = 0 Then
'        sCRCJ = ""
'    Else
'        sPrefixo = "1"
'        sCRCJ = sPrefixo & !uf_crc_pj & !numero_crc_pj & "/" & !tipo_crc_pj & "-" & !classif_crc_pj
'    End If
'
'    lblCRCContadorPF.Caption = sCRCF
'    lblCPFContador.Caption = Format(!cpf_contador, "00#\.###\.###-##")
'    lblCRCContadorPJ.Caption = sCRCJ
'    lblCnpjContador.Caption = Format(!cnpj_contador, "0#\.###\.###/####-##")
'
'
'   If Val(!cnpj_contador) = 0 And Val(!cpf_contador) = 0 Then
'        lblNomeContador.Caption = ""
'        lblNomeContador.Tag = ""
'   Else
'        If Val(!cnpj_contador) > 0 Then
'            Sql = "select * from escritoriocontabil where cnpj='" & !cnpj_contador & "' or crc='" & sCRCJ & "'"
'            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'            If RdoAux2.RowCount > 0 Then
'                lblNomeContador.Caption = RdoAux2!NOMEESC
'                lblNomeContador.Tag = CStr(RdoAux2!codigoesc)
'            End If
'            RdoAux2.Close
'        End If
'        If lblNomeContador.Caption = "" And (Val(!cpf_contador) > 0 Or sCrc <> "") Then
'            Sql = "select * from escritoriocontabil where cpf='" & !cpf_contador & "' or crc='" & sCRCF & "'"
'            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'            If RdoAux2.RowCount > 0 Then
'                lblNomeContador.Caption = RdoAux2!NOMEESC
'                lblNomeContador.Tag = CStr(RdoAux2!codigoesc)
'            End If
'            RdoAux2.Close
'        End If
'    End If
'   .Close
'End With
'
'Sql = "select * from vre_socio where id=" & nId
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    Do Until .EOF
'        Sql = "select * from cidadao where cpf='" & Format(Val(!Numero), "00000000000") & "'"
'        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'        If RdoAux2.RowCount > 0 Then
'            sCodigo = Format(RdoAux2!CodCidadao, "000000")
'        Else
'            sCodigo = "000000"
'        End If
'        RdoAux2.Close
'        Set itmX = lvSocio.ListItems.Add(, , sCodigo)
'        itmX.SubItems(1) = !Nome
'        itmX.SubItems(2) = !Numero
'       .MoveNext
'    Loop
'   .Close
'End With
'
'Sql = "select * from vre_atividade where id=" & nId
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    Do Until .EOF
'        Set itmX = lvCnae.ListItems.Add(, , !Cnae)
'        itmX.SubItems(1) = IIf(!principal, "Sim", "Não")
'        itmX.SubItems(2) = IIf(!exercida, "Sim", "Não")
'       .MoveNext
'    Loop
'   .Close
'End With
'
''Sql = "select * from vre_licenciamento where empresa_id=" & nId & " and orgao=1"
'Sql = "select * from vre_licenciamento where empresa_id=" & nId & " and orgao_id=287"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    If .RowCount > 0 Then
'        If Val(!Numero) > 0 Then
'            lblProtocolo.Caption = Format(RetornaNumero(!Numero), "0000000000000")
'            lblProtocolo.Caption = Mid(lblProtocolo.Caption, 1, Len(lblProtocolo.Caption) - 6) & "." & Mid(lblProtocolo.Caption, 8, 4) & "-" & Right(lblProtocolo.Caption, 2)
'            lblDataEmissao.Caption = Format(!Data_Emissao, "dd/mm/yyyy")
'            lblValidade.Caption = Format(!Data_Vencimento, "dd/mm/yyyy")
'        Else
'            If MsgBox("Número de protocolo não encontrado, deseja informar agora?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
'                zNumero = InputBox("Digite o número do protocolo", "Nº do protocolo")
'                If Val(zNumero) = 0 Then
'                    MsgBox "Nº de protocolo inválido", vbCritical, "Erro"
'                    bOk = False
'                Else
'                    zDataEmissao = InputBox("Digite a data de emissão.", "Data de emissão")
'                    If Not IsDate(zDataEmissao) Then
'                        MsgBox "Data de emissão inválida", vbCritical, "Erro"
'                        bOk = False
'                    Else
'                        zDataVencimento = InputBox("Digite a data de vencimento.", "Data de vencimento")
'                        If Not IsDate(zDataVencimento) Then
'                            MsgBox "Data de vencimento inválida", vbCritical, "Erro"
'                            bOk = False
'                        Else
'                            Sql = "update vre_licenciamento set numero='" & zNumero & "',data_emissao='" & Format(zDataEmissao, "mm/dd/yyyy") & "',data_vencimento='" & Format(zDataVencimento, "mm/dd/yyyy") & "' "
'                            Sql = Sql & "Where empresa_id=" & nId & " and orgao_id=1"
'                            cn.Execute Sql, rdExecDirect
'
'                            lblProtocolo.Caption = zNumero
'                            If InStr(1, lblProtocolo.Caption, ".") > 0 Then
'                                lblProtocolo.Caption = Mid(lblProtocolo.Caption, 1, Len(lblProtocolo.Caption) - 6) & "." & Mid(lblProtocolo.Caption, 7, 4) & "-" & Right(lblProtocolo.Caption, 2)
'                            End If
'                            lblDataEmissao.Caption = zDataEmissao
'                            lblValidade.Caption = zDataVencimento
'                        End If
'                    End If
'                End If
'            Else
'                bOk = False
'            End If
'        End If
'    End If
'   .Close
'End With

End Sub

