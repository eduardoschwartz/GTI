VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.MDIForm frmMdi 
   BackColor       =   &H00808080&
   Caption         =   "Prefeitura Municipal de Jaboticabal - Gest�o de Tributa��o Municipal Integrada (G.T.I.)"
   ClientHeight    =   5835
   ClientLeft      =   5325
   ClientTop       =   2640
   ClientWidth     =   12450
   Icon            =   "frmMdi.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Interval        =   60000
      Left            =   900
      Top             =   3240
   End
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   2070
      Top             =   2115
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   12420
      TabIndex        =   25
      Top             =   0
      Width           =   12450
      Begin prjChameleon.chameleonButton cmdJanela 
         Height          =   285
         Left            =   8055
         TabIndex        =   34
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   503
         BTYPE           =   8
         TX              =   "&Janela"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   4
         FOCUSR          =   0   'False
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMdi.frx":08CA
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdPrincipal 
         Height          =   285
         Left            =   45
         TabIndex        =   26
         Top             =   0
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         BTYPE           =   8
         TX              =   "&Principal"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   4
         FOCUSR          =   0   'False
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMdi.frx":08E6
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdParametros 
         Height          =   285
         Left            =   945
         TabIndex        =   27
         Top             =   0
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   503
         BTYPE           =   8
         TX              =   "Par�&metros"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   4
         FOCUSR          =   0   'False
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMdi.frx":0902
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdImobiliario 
         Height          =   285
         Left            =   2025
         TabIndex        =   28
         Top             =   0
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   503
         BTYPE           =   8
         TX              =   "&Imobili�rio"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   4
         FOCUSR          =   0   'False
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMdi.frx":091E
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdMobiliario 
         Height          =   285
         Left            =   3105
         TabIndex        =   29
         Top             =   0
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   503
         BTYPE           =   8
         TX              =   "&Mobili�rio"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   4
         FOCUSR          =   0   'False
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMdi.frx":093A
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdAtende 
         Height          =   285
         Left            =   4095
         TabIndex        =   30
         Top             =   0
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         BTYPE           =   8
         TX              =   "&Atendimento"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   4
         FOCUSR          =   0   'False
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMdi.frx":0956
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdTributo 
         Height          =   285
         Left            =   5310
         TabIndex        =   31
         Top             =   0
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   503
         BTYPE           =   8
         TX              =   "&Tribut�rio"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   4
         FOCUSR          =   0   'False
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMdi.frx":0972
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdProtocolo 
         Height          =   285
         Left            =   6345
         TabIndex        =   32
         Top             =   0
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   503
         BTYPE           =   8
         TX              =   "&Protocolo"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   4
         FOCUSR          =   0   'False
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMdi.frx":098E
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdOutros 
         Height          =   285
         Left            =   7380
         TabIndex        =   33
         Top             =   0
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   503
         BTYPE           =   8
         TX              =   "&Outros"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   4
         FOCUSR          =   0   'False
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMdi.frx":09AA
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.PictureBox picBar 
      Align           =   1  'Align Top
      BackColor       =   &H00E7E3E7&
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   12450
      TabIndex        =   7
      Top             =   345
      Width           =   12450
      Begin prjChameleon.chameleonButton btBar 
         Height          =   555
         Index           =   12
         Left            =   11835
         TabIndex        =   36
         ToolTipText     =   "Agenda"
         Top             =   0
         Visible         =   0   'False
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   979
         BTYPE           =   8
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
         BCOL            =   15197159
         BCOLO           =   15197159
         FCOL            =   12648447
         FCOLO           =   0
         MCOL            =   0
         MPTR            =   99
         MICON           =   "frmMdi.frx":09C6
         PICN            =   "frmMdi.frx":0CE0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton btBar 
         Height          =   555
         Index           =   11
         Left            =   6675
         TabIndex        =   35
         ToolTipText     =   "Chat do Sistema"
         Top             =   30
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   979
         BTYPE           =   8
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
         BCOL            =   15197159
         BCOLO           =   15197159
         FCOL            =   12648447
         FCOLO           =   0
         MCOL            =   15197159
         MPTR            =   99
         MICON           =   "frmMdi.frx":1354
         PICN            =   "frmMdi.frx":166E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton btBar 
         Height          =   555
         Index           =   10
         Left            =   5490
         TabIndex        =   24
         ToolTipText     =   "Certid�es"
         Top             =   30
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   979
         BTYPE           =   8
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
         BCOL            =   15197159
         BCOLO           =   15197159
         FCOL            =   12648447
         FCOLO           =   0
         MCOL            =   0
         MPTR            =   99
         MICON           =   "frmMdi.frx":1AC9
         PICN            =   "frmMdi.frx":1DE3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton btBar 
         Height          =   555
         Index           =   9
         Left            =   6075
         TabIndex        =   23
         ToolTipText     =   "Efetuar Login/Logoff"
         Top             =   30
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   979
         BTYPE           =   8
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
         BCOL            =   15197159
         BCOLO           =   15197159
         FCOL            =   12648447
         FCOLO           =   0
         MCOL            =   15197159
         MPTR            =   99
         MICON           =   "frmMdi.frx":2346
         PICN            =   "frmMdi.frx":2660
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton btBar 
         Height          =   555
         Index           =   8
         Left            =   4890
         TabIndex        =   22
         ToolTipText     =   "Alvar� de Funcionamento"
         Top             =   30
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   979
         BTYPE           =   8
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
         BCOL            =   15197159
         BCOLO           =   15197159
         FCOL            =   12648447
         FCOLO           =   0
         MCOL            =   15197159
         MPTR            =   99
         MICON           =   "frmMdi.frx":32B2
         PICN            =   "frmMdi.frx":35CC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton btBar 
         Height          =   555
         Index           =   7
         Left            =   4290
         TabIndex        =   21
         ToolTipText     =   "Protocolo"
         Top             =   30
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   979
         BTYPE           =   8
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
         BCOL            =   15197159
         BCOLO           =   15197159
         FCOL            =   12648447
         FCOLO           =   0
         MCOL            =   15197159
         MPTR            =   99
         MICON           =   "frmMdi.frx":421E
         PICN            =   "frmMdi.frx":4538
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton btBar 
         Height          =   555
         Index           =   6
         Left            =   3690
         TabIndex        =   20
         ToolTipText     =   "Extrato do Contribuinte"
         Top             =   30
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   979
         BTYPE           =   8
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
         BCOL            =   15197159
         BCOLO           =   15197159
         FCOL            =   12648447
         FCOLO           =   0
         MCOL            =   15197159
         MPTR            =   99
         MICON           =   "frmMdi.frx":518A
         PICN            =   "frmMdi.frx":54A4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton btBar 
         Height          =   555
         Index           =   5
         Left            =   3090
         TabIndex        =   19
         ToolTipText     =   "Consulta Documentos"
         Top             =   30
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   979
         BTYPE           =   8
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
         BCOL            =   15197159
         BCOLO           =   15197159
         FCOL            =   12648447
         FCOLO           =   0
         MCOL            =   15197159
         MPTR            =   99
         MICON           =   "frmMdi.frx":60F6
         PICN            =   "frmMdi.frx":6410
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton btBar 
         Height          =   555
         Index           =   4
         Left            =   2490
         TabIndex        =   18
         ToolTipText     =   "Parcelamento de Divida"
         Top             =   30
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   979
         BTYPE           =   8
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
         BCOL            =   15197159
         BCOLO           =   15197159
         FCOL            =   12648447
         FCOLO           =   0
         MCOL            =   15197159
         MPTR            =   99
         MICON           =   "frmMdi.frx":6910
         PICN            =   "frmMdi.frx":6C2A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton btBar 
         Height          =   555
         Index           =   3
         Left            =   1890
         TabIndex        =   17
         ToolTipText     =   "Emiss�o de Guias"
         Top             =   30
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   979
         BTYPE           =   8
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
         BCOL            =   15197159
         BCOLO           =   15197159
         FCOL            =   12648447
         FCOLO           =   0
         MCOL            =   15197159
         MPTR            =   99
         MICON           =   "frmMdi.frx":787C
         PICN            =   "frmMdi.frx":7B96
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame frDV 
         BackColor       =   &H00E7E3E7&
         Height          =   510
         Left            =   8415
         TabIndex        =   11
         Top             =   0
         Width           =   3390
         Begin VB.ComboBox cmbDV 
            Height          =   315
            ItemData        =   "frmMdi.frx":87E8
            Left            =   720
            List            =   "frmMdi.frx":87F5
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   135
            Width           =   960
         End
         Begin VB.TextBox txtDV 
            Height          =   315
            Left            =   1710
            TabIndex        =   13
            Top             =   135
            Width           =   1050
         End
         Begin VB.CommandButton cmdDV 
            Appearance      =   0  'Flat
            BackColor       =   &H00E7E3E7&
            Height          =   285
            Left            =   2790
            Picture         =   "frmMdi.frx":880E
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Retorna digito verificador"
            Top             =   135
            Width           =   330
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "D.Verif:"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   135
            TabIndex        =   16
            Top             =   210
            Width           =   600
         End
         Begin VB.Label lblDV 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Left            =   3150
            TabIndex        =   15
            Top             =   180
            Width           =   195
         End
      End
      Begin prjChameleon.chameleonButton btBar 
         Height          =   555
         Index           =   2
         Left            =   1285
         TabIndex        =   8
         ToolTipText     =   "Cadastro de Cidad�o"
         Top             =   30
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   979
         BTYPE           =   8
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
         BCOL            =   15197159
         BCOLO           =   15197159
         FCOL            =   12648447
         FCOLO           =   0
         MCOL            =   15197159
         MPTR            =   99
         MICON           =   "frmMdi.frx":8958
         PICN            =   "frmMdi.frx":8C72
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton btBar 
         Height          =   555
         Index           =   0
         Left            =   90
         TabIndex        =   9
         ToolTipText     =   "Cadastro de Im�veis"
         Top             =   30
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   979
         BTYPE           =   8
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
         BCOL            =   15197159
         BCOLO           =   15197159
         FCOL            =   255
         FCOLO           =   255
         MCOL            =   0
         MPTR            =   99
         MICON           =   "frmMdi.frx":98C4
         PICN            =   "frmMdi.frx":9BDE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton btBar 
         Height          =   555
         Index           =   1
         Left            =   695
         TabIndex        =   10
         ToolTipText     =   "Cadastro de Empresas e Servi�os"
         Top             =   30
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   979
         BTYPE           =   8
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
         BCOL            =   15197159
         BCOLO           =   15197159
         FCOL            =   12648447
         FCOLO           =   0
         MCOL            =   15197159
         MPTR            =   99
         MICON           =   "frmMdi.frx":A263
         PICN            =   "frmMdi.frx":A57D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton btBar 
         Height          =   555
         Index           =   13
         Left            =   7290
         TabIndex        =   37
         ToolTipText     =   "Hist�rico de Atualiza��es"
         Top             =   45
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   979
         BTYPE           =   8
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
         BCOL            =   15197159
         BCOLO           =   15197159
         FCOL            =   12648447
         FCOLO           =   0
         MCOL            =   15197159
         MPTR            =   99
         MICON           =   "frmMdi.frx":B1CF
         PICN            =   "frmMdi.frx":B4E9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1035
      Top             =   1380
   End
   Begin VB.PictureBox picBackdrop 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   12450
      TabIndex        =   4
      Top             =   330
      Visible         =   0   'False
      Width           =   12450
      Begin VB.PictureBox picOriginal 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   11520
         Left            =   240
         Picture         =   "frmMdi.frx":BDC3
         ScaleHeight     =   11520
         ScaleWidth      =   15360
         TabIndex        =   6
         Top             =   120
         Width           =   15360
      End
      Begin VB.PictureBox picStretched 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   7260
         Left            =   2040
         ScaleHeight     =   7260
         ScaleWidth      =   4095
         TabIndex        =   5
         Top             =   600
         Width           =   4095
      End
   End
   Begin VB.PictureBox Picture4 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   12390
      TabIndex        =   2
      Top             =   5220
      Visible         =   0   'False
      Width           =   12450
      Begin VB.Label frTeste 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "BASE DE TESTE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   15060
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   12450
      TabIndex        =   0
      Top             =   5535
      Width           =   12450
      Begin MSComctlLib.StatusBar Sbar 
         Height          =   270
         Left            =   45
         TabIndex        =   1
         Top             =   15
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   476
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   6
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Width           =   8185
               MinWidth        =   8185
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   5292
               MinWidth        =   5292
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               Enabled         =   0   'False
               Object.Width           =   531
               MinWidth        =   531
            EndProperty
            BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   2
               Alignment       =   1
               Object.Width           =   1059
               MinWidth        =   1059
               TextSave        =   "NUM"
            EndProperty
            BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   3
               Alignment       =   1
               Enabled         =   0   'False
               Object.Width           =   1059
               MinWidth        =   1059
               TextSave        =   "INS"
            EndProperty
            BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               Object.Width           =   3528
               MinWidth        =   3528
               Text            =   "Data Base: 00/00/0000"
               TextSave        =   "Data Base: 00/00/0000"
            EndProperty
         EndProperty
      End
      Begin VB.Image imStatus 
         Height          =   240
         Index           =   2
         Left            =   5970
         Picture         =   "frmMdi.frx":36095
         Top             =   390
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imStatus 
         Height          =   240
         Index           =   1
         Left            =   6240
         Picture         =   "frmMdi.frx":3641F
         Top             =   330
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imOK 
         Height          =   240
         Left            =   11370
         Picture         =   "frmMdi.frx":367A9
         Top             =   30
         Width           =   240
      End
      Begin VB.Image imWorking 
         Height          =   240
         Left            =   11640
         Picture         =   "frmMdi.frx":36B33
         Top             =   30
         Width           =   240
      End
      Begin VB.Image imStatus 
         Height          =   240
         Index           =   0
         Left            =   5100
         Picture         =   "frmMdi.frx":36EBD
         Top             =   420
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.Menu mnuJanela 
      Caption         =   "&Janela"
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu mnuCloseAll 
         Caption         =   "Fechar Todas"
      End
   End
   Begin VB.Menu mnuCertidao 
      Caption         =   "&Certid�es"
      Visible         =   0   'False
      Begin VB.Menu mnuCertidaoDebito 
         Caption         =   "Certid�o de D�bito"
      End
      Begin VB.Menu mnuCertidaoValorVenal 
         Caption         =   "Certid�o de Valor Venal"
      End
      Begin VB.Menu mnuCertidaoEndereco 
         Caption         =   "Certid�o de Endere�o"
      End
      Begin VB.Menu mnuCertidaoDemolicao 
         Caption         =   "Certid�o de Demoli��o"
      End
      Begin VB.Menu mnuCertidaoIsencao 
         Caption         =   "Certid�o de Isen��o"
      End
      Begin VB.Menu mnuCertidaoIsencaoITBI 
         Caption         =   "Certid�o de N�o Incid�ncia de ITBI"
      End
      Begin VB.Menu mnuRenovaAlvara 
         Caption         =   "Renova��o de Alvar�"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents m_cMenuPrincipal As cPopupMenu
Attribute m_cMenuPrincipal.VB_VarHelpID = -1
Public WithEvents m_cMenuParam As cPopupMenu
Attribute m_cMenuParam.VB_VarHelpID = -1
Public WithEvents m_cMenuImob As cPopupMenu
Attribute m_cMenuImob.VB_VarHelpID = -1
Public WithEvents m_cMenuMob As cPopupMenu
Attribute m_cMenuMob.VB_VarHelpID = -1
Public WithEvents m_cMenuAtende As cPopupMenu
Attribute m_cMenuAtende.VB_VarHelpID = -1
Public WithEvents m_cMenuProt As cPopupMenu
Attribute m_cMenuProt.VB_VarHelpID = -1
Public WithEvents m_cMenuOutro As cPopupMenu
Attribute m_cMenuOutro.VB_VarHelpID = -1
Public WithEvents m_cMenuTrib As cPopupMenu
Attribute m_cMenuTrib.VB_VarHelpID = -1
Private Type FileData
    sName As String
    dDate As Date
End Type

Private Type tVersion
    Major As Integer
    Minor As Integer
    Revision As Integer
    version As String
End Type

Dim RunOnce As Boolean, lngTimer As Long, lngTimer2 As Long, FlagServico As Long
Dim sRet As String
Dim evOpen As Integer
Dim bOpen As Boolean


Private Sub FormHagana()

evOpen = 1

End Sub


Private Sub btBar_Click(Index As Integer)

Select Case Index
    Case 0
        If m_cMenuImob.Enabled(m_cMenuImob.IndexForKey("mnuCadImob")) Then
            frmCadImob.show
            frmCadImob.ZOrder 0
        Else
            MsgBox "Voc� n�o possui permiss�o para acessar este m�dulo!!!", vbCritical, "Seguran�a do GTI"
        End If
    Case 1
        If m_cMenuMob.Enabled(m_cMenuMob.IndexForKey("mnuCadMobiliario")) Then
            frmCadMob.show
            frmCadMob.ZOrder 0
        Else
            MsgBox "Voc� n�o possui permiss�o para acessar este m�dulo!!!", vbCritical, "Seguran�a do GTI"
        End If
    Case 2
        If m_cMenuAtende.Enabled(m_cMenuAtende.IndexForKey("mnuCidadao")) Then
                frmCidadao.show
                frmCidadao.ZOrder 0
        Else
            MsgBox "Voc� n�o possui permiss�o para acessar este m�dulo!!!", vbCritical, "Seguran�a do GTI"
        End If
    Case 3
        If m_cMenuAtende.Enabled(m_cMenuAtende.IndexForKey("mnu2ViaLaser")) Then
            FlagForm = 1
            frmEmissaoGuia.show
            frmEmissaoGuia.ZOrder 0
        Else
            MsgBox "Voc� n�o possui permiss�o para acessar este m�dulo!!!", vbCritical, "Seguran�a do GTI"
        End If
    Case 4
        If m_cMenuAtende.Enabled(m_cMenuAtende.IndexForKey("mnuParcelamento")) Then
            frmParcelamentoNovo.show
            frmParcelamentoNovo.ZOrder 0
        Else
            MsgBox "Voc� n�o possui permiss�o para acessar este m�dulo!!!", vbCritical, "Seguran�a do GTI"
        End If
    Case 5
        If m_cMenuAtende.Enabled(m_cMenuAtende.IndexForKey("mnuCnsNumDoc")) Then
            frmDoc.show
            frmDoc.ZOrder 0
        Else
            MsgBox "Voc� n�o possui permiss�o para acessar este m�dulo!!!", vbCritical, "Seguran�a do GTI"
        End If
    Case 6
        If m_cMenuAtende.Enabled(m_cMenuAtende.IndexForKey("CnsDebitoImob")) Then
            frmDebitoImob.show
            frmDebitoImob.ZOrder 0
        Else
            MsgBox "Voc� n�o possui permiss�o para acessar este m�dulo!!!", vbCritical, "Seguran�a do GTI"
        End If
    Case 7
        If m_cMenuProt.Enabled(m_cMenuProt.IndexForKey("mnuProcesso")) Then
            frmProcesso.show
            frmProcesso.ZOrder 0
        Else
            MsgBox "Voc� n�o possui permiss�o para acessar este m�dulo!!!", vbCritical, "Seguran�a do GTI"
        End If
    Case 8
        If m_cMenuMob.Enabled(m_cMenuMob.IndexForKey("mnuAlvara")) Then
            frmAlvaraNovo.show
            frmAlvaraNovo.ZOrder (0)
        Else
            MsgBox "Voc� n�o possui permiss�o para acessar este m�dulo!!!", vbCritical, "Seguran�a do GTI"
        End If
    Case 9
        If Forms.Count > 1 Then
           If MsgBox("Deseja fechar todas as telas e bloquear o sistema ?", vbQuestion + vbYesNo, "Confirma��o") = vbYes Then
              mnuCloseAll_Click
              modLg "Desconectado do Sistema"
              Sql = "UPDATE USUARIO SET LOGON=0 wHERE NOMELOGIN='" & NomeDeLogin & "'"
              cn.Execute Sql, rdExecDirect
              
           frmLogin.show vbModal
           End If
        Else
           frmLogin.show vbModal
        End If
    Case 10
        sRet = RetEventUserForm("frmCertidao")
        If NomeDeLogin <> "SCHWARTZ" Then
            If InStr(1, sRet, Format(evOpen, "000"), vbBinaryCompare) > 0 Then
               PopupMenu mnuCertidao, , btBar(8).Left, btBar(10).Top + btBar(10).Height + Picture1.Height
            Else
                MsgBox "Voc� n�o possui permiss�o para acessar este m�dulo!!!", vbCritical, "Seguran�a do GTI"
            End If
        Else
              PopupMenu mnuCertidao, , btBar(8).Left, btBar(10).Top + btBar(10).Height + Picture1.Height
        End If
    Case 11
        btBar(11).BackColor = &HE7E3E7
        frmChat.Timer2.Interval = 0
        frmChat.show
        frmChat.ZOrder (0)
    Case 12
        If NomeDeLogin <> "SCHWARTZ" Then
            MsgBox "Em constru��o"
            Exit Sub
        End If
        'frmAgenda.show
        'frmAgenda.ZOrder (0)
        frmRTF.show
     Case 13
        Gera_WhatsNew
End Select
End Sub

Private Sub cmdAtende_Click()
Dim IpAddrs, sTipo As String, Sql As String, RdoAux As rdoResultset, i As Integer
Dim FS As FileSystemObject

lIndex = m_cMenuAtende.ShowPopupMenu(cmdAtende.Left, cmdAtende.Top, cmdAtende.Left, cmdAtende.Top, Me.ScaleWidth - cmdAtende.Left - cmdAtende.Width, cmdAtende.Top + cmdAtende.Height, False)
If (lIndex > 0) Then
    Ocupado
    Select Case m_cMenuAtende.ItemKey(lIndex)
        Case "mnuSenhaControle"
            Set FS = New FileSystemObject
            If FS.FileExists(App.Path & "\gti.ini") Then
            Open App.Path & "\gti.ini" For Input As #1
            Do While Not EOF(1)
                Line Input #1, strLinha
                If Left(strLinha, 6) = "GUICHE" Then
                    nGuiche = Val(Right(strLinha, 2))
                ElseIf Left(strLinha, 4) = "TIPO" Then
                    sTipo = Right(strLinha, 1)
                End If
            Loop
            Close #1
            End If
            Set FS = Nothing
SENHA:
            If sTipo = "G" Then
                Set frm = frmSenhaGuiche
            ElseIf sTipo = "P" Then
                Set frm = frmSenhaPre
            ElseIf sTipo = "M" Then
                Set frm = frmSenhaMonitor
            Else
                MsgBox "Seu computador n�o pode executar o sistema de senhas.", vbCritical, "Aten��o"
                Liberado
                Exit Sub
            End If
        Case "mnuSenhaResumo"
            Set frm = frmSenhaStatus
        Case "mnuSenhaStatus"
             Set frm = frmSenhaStatus
        Case "mnuCidadao"
             Set frm = frmCidadao
        Case "mnu2ViaLaser"
            Set frm = frmEmissaoGuia
            FlagForm = 1
 '            Set frm = frm2ViaLaser
'        Case "mnuEmissaoGuia"
'             Set frm = frmEmissaoGuia
        Case "mnuMovimento"
            Set frm = frmMovEconomico
        Case "mnuAutorizaNF"
            Set frm = frmAutorizaNota
        Case "mnuITBI"
            Set frm = frmITBI
        Case "mnuDeca"
            Set frm = Nothing
            Deca
        Case "mnuRequerimentoProc"
            Set frm = frmRequerimento
        Case "mnuGare"
            Set frm = frmGare
        Case "mnuFundoDespesa"
            Set frm = frmFundoDespesa
        Case "mnuDepositoCRI"
            Set frm = Nothing
            frmReport.ShowReport "DEPOSITOCRI", frmMdi.HWND, Me.HWND
        Case "mnuDeclaraIsentoIPTU"
            Set frm = frmDeclaraIsento
            
        Case "mnuRequerIsentoIPTU"
            Set frm = frmRequerIPTU
        Case "mnuCancelReparc"
            Set frm = frmCancelReparc
'        Case "mnuParcelaDebito"
'            Set frm = frmParcelamento2
        Case "mnuParcelamento"
            Set frm = frmParcelamentoNovo
        Case "mnuLiberaCarne"
            Set frm = frmLiberaCarne
        Case "mnuCancelParcelamento"
            Set frm = frmCancelParcelamento
        Case "mnuCancelParcelamentoAuto"
            Set frm = frmCancelParcAuto
        Case "mnuMalaDiretaParc"
            Set frm = frmMalaDiretaISS
        Case "mnuDesbloquearParc"
            Set frm = frmParcelamentoBloqueio
        Case "mnuPagamentoMensalParc"
            Set frm = frmRelParcPago
        Case "mnuGuiaPratico1"
            Set frm = frmGuiaPratico1
        Case "mnuGuiaPratico2"
            Set frm = frmGuiaPratico2
        Case "mnuGuiaPratico5"
            Set frm = frmGuiaPratico5
        Case "mnuGuiaPratico3"
            Set frm = frmGuiaPratico3
        Case "mnu2Via"
            Set frm = frmEmissaoGuia
            FlagForm = 2
        Case "mnu2ViaEspecial"
            Set frm = frmEmissao2ViaEspecial
        Case "mnuTermConf"
            Set frm = frmConfissaoDivida
        Case "mnuCobrancaJudicial"
            Set frm = Nothing
            frmReport.ShowReport "COMUNICADOJUDICIAL", frmMdi.HWND, Me.HWND
        Case "mnuEmiteDoc"
            Set frm = Nothing
            frmReport.ShowReport2 "DOCEMITIDO", frmMdi.HWND, Me.HWND
        Case "mnuRelRefis"
            Set frm = Nothing
            frmReport.ShowReport2 "REFIS", frmMdi.HWND, Me.HWND
'            Set frm = Nothing
'            Ocupado
 '           GeraRefisDAM (Year(Now))
'            Liberado
        Case "mnuRelRefisTributo"
            Set frm = frmRefisDetalhe
        Case "mnuRelRefisParc"
            Set frm = Nothing
            frmReport.ShowReport2 "REFISPARC", frmMdi.HWND, Me.HWND
        Case "mnuRenovaAlvara"
            Set frm = Nothing
'            frmReport.ShowReport2 "ALVARARENOVA", frmMdi.HWND, Me.HWND
        Case "mnuSenhaISS"
            Set frm = Nothing
            frmReport.ShowReport2 "SENHAISS3", frmMdi.HWND, Me.HWND
        Case "mnuConsultaAte"
            Set frm = Nothing
        Case "mnuCnsNumDoc"
            Set frm = frmDoc
        Case "CnsDebitoImob"
            Set frm = frmDebitoImob
        Case "mnuMalaDiretaCidadao"
            Set frm = frmMalaDiretaCidadao
    End Select
    If Not frm Is Nothing Then
        frm.show
        frm.ZOrder 0
    End If
    Liberado
Else
   lIndex = 0 ' cancelled the menu.
End If
End Sub

Private Sub cmdDV_Click()
If Val(txtDV.Text) = 0 Then
    lblDV.Caption = "?"
    Exit Sub
End If
If cmbDV.ListIndex = 0 Then
    lblDV = RetornaDVNumDoc(txtDV.Text)
ElseIf cmbDV.ListIndex = 1 Then
    lblDV = RetornaDVProcesso(txtDV.Text)
ElseIf cmbDV.ListIndex = 2 Then
    lblDV = RetornaDVCodReduzido(txtDV.Text)
End If

End Sub

Private Sub cmdImobiliario_Click()
Dim z As Variant, z2 As Variant


lIndex = m_cMenuImob.ShowPopupMenu(cmdImobiliario.Left, cmdImobiliario.Top, cmdImobiliario.Left, cmdImobiliario.Top, Me.ScaleWidth - cmdImobiliario.Left - cmdImobiliario.Width, cmdImobiliario.Top + cmdImobiliario.Height, False)
If (lIndex > 0) Then
    Ocupado
    Select Case m_cMenuImob.ItemKey(lIndex)
    Case "mnuCadImob"
        Set frm = frmCadImob
    Case "mnuCondominio"
        Set frm = frmCadCondominio
    Case "mnuCadastroObra"
        Set frm = frmCadastroObra
    Case "mnuLogr"
        Set frm = frmLogradouro
    Case "mnuSegmentoLogradouro"
        Set frm = frmSegmentoLogradouro
    Case "mnuFaceQuadra"
        Set frm = frmFaceQuadra
    Case "mnuCnsImovel"
        Set frm = frmCnsImovel
     Case "mnuCnsAvancadaImob"
        Set frm = frmCnsAvancadaImob
    Case "mnuDetImovel"
        Set frm = frmDadosImovel
    Case "mnuResumoImovel"
        Set frm = frmResumoImobiliario
    Case "mnuDesmem"
        Set frm = frmDesmembramento
    Case "mnuDesdobroCarne"
        Set frm = frmDesdobroCarne
    Case "mnuVVDeclarado"
        Set frm = frmVVDeclarado
    Case "mnuSimuladoPG"
        Set frm = frmSimuladoPG
    Case "mnuEspolio"
        Set frm = frmEspolio
    Case "mnuCorrigeBairro"
        Set frm = frmCepBairro
    Case "mnuUnificacao"
        Set frm = frmUnificacao
    Case "mnuImun"
        Set frm = frmIsencao
    Case "mnuAverbacao"
        Set frm = frmAverbacao
    Case "mnuRolImovel"
        Set frm = frmRolImovel
    Case "mnuDevedorIPTU"
        Set frm = Nothing
        frmReport.ShowReport "DEVEDORIPTUANUAL", frmMdi.HWND, Me.HWND
    Case "mnuCadRural"
        Set frm = frmCadastroRural
    Case "mnuProdutoRural"
        Set frm = frmProdutoRural
    Case "mnuEstradaRural"
        Set frm = frmEstradaRural
    Case "mnuRelCadRural"
        Set frm = Nothing
        frmReport.ShowReport "LISTARURAL", frmMdi.HWND, Me.HWND
    Case "mnuRelCadRuralFull"
        Set frm = Nothing
        frmReport.ShowReport "LISTARURALFULL", frmMdi.HWND, Me.HWND
    Case "mnuRelCadRuralFaixa"
        Set frm = Nothing
        frmReport.ShowReport "LISTARURAL2", frmMdi.HWND, Me.HWND
    Case "mnuRelProdRural"
        Set frm = Nothing
        frmReport.ShowReport "RURALPRODUTO", frmMdi.HWND, Me.HWND
    Case "mnuSimulaRural"
        Set frm = Nothing
        frmReport.ShowReport "LISTARURAL3", frmMdi.HWND, Me.HWND
    Case "mnuEventoRural"
        Set frm = Nothing
        frmReport.ShowReport "EVENTORURAL", frmMdi.HWND, Me.HWND
    Case "mnuLancRocada"
        Set frm = Nothing
        frmReport.ShowReport2 "PAGAMENTOROCADA", frmMdi.HWND, Me.HWND
    Case "mnuMalaDiretaRocada"
        Set frm = Nothing
        z = InputBox("Digite a Data de emiss�o.", "Entre com os dados", Format(Now, "dd/mm/yyyy"))
        If Not IsDate(z) Then
            MsgBox "Data inv�lida.", vbExclamation, "Aten��o"
        Else
            MalaDiretaRo�ada (CDate(z))
            frmReport.ShowReport "ETIQUETAGTI", frmMdi.HWND, Me.HWND
            cn.Execute "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'", rdExecDirect
        End If
    Case "mnuMalaDiretaISSCCivil"
        Set frm = Nothing
        z = InputBox("Digite a Data inicial de emiss�o.", "Entre com os dados", Format(Now, "dd/mm/yyyy"))
        If Not IsDate(z) Then
            MsgBox "Data inv�lida.", vbExclamation, "Aten��o"
        Else
            z2 = InputBox("Digite a Data final de emiss�o.", "Entre com os dados", Format(Now, "dd/mm/yyyy"))
            If Not IsDate(z) Then
                MsgBox "Data inv�lida.", vbExclamation, "Aten��o"
            Else
                MalaDiretaISSCCivil CDate(z), CDate(z2)
                frmReport.ShowReport "ETIQUETAGTI", frmMdi.HWND, Me.HWND
                cn.Execute "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'", rdExecDirect
            End If
        End If
    Case "mnuMalaDiretaRural"
        Set frm = Nothing
        MalaDiretaRural
        frmReport.ShowReport "ETIQUETACONSIST", frmMdi.HWND, Me.HWND
        cn.Execute "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'", rdExecDirect
    Case "mnuSisObraPref"
        Set frm = frmSisObras
    Case "mnuConversorSisObra"
        Set frm = frmConversorSisObraPref
    End Select
    If Not frm Is Nothing Then
        frm.show
        frm.ZOrder 0
    End If
        Liberado
Else
    lIndex = 0 ' cancelled the menu.
End If
End Sub

Private Sub cmdJanela_Click()
PopupMenu mnuJanela, , cmdJanela.Left, cmdJanela.Top + cmdJanela.Height
End Sub

Private Sub cmdMobiliario_Click()
lIndex = m_cMenuMob.ShowPopupMenu(cmdMobiliario.Left, cmdMobiliario.Top, cmdMobiliario.Left, cmdMobiliario.Top, Me.ScaleWidth - cmdMobiliario.Left - cmdMobiliario.Width, cmdMobiliario.Top + cmdMobiliario.Height, False)
If (lIndex > 0) Then
    Ocupado
    Select Case m_cMenuMob.ItemKey(lIndex)
        Case "mnuCadMobiliario"
             Set frm = frmCadMob
        Case "mnuEscContab"
             Set frm = frmEscContab
        Case "mnuHorarioFuncionamento"
             Set frm = frmHorarioFuncionamento
        Case "mnuTabAtivTL"
            Set frm = frmAtiv
        Case "mnuTabAtivISS"
            Set frm = frmAtivISS
        Case "mnuVigSan"
             Set frm = frmVigSanitaria
        Case "mnuVRE"
             Set frm = frmDadosVRE
        Case "mnuVRERedeSim"
             Set frm = frmImportaRedeSim
        Case "mnuCnsEmpresa"
             Set frm = frmCnsMob
        Case "mnuProdutEvento"
             Set frm = frmProdutividadeEvento
        Case "mnuProdutTarefa"
             Set frm = frmProdutividadeControle
        Case "mnuProdutExtratoMes"
             Set frm = frmProdutividadeMensal
        Case "mnuProdutSaldoMes"
             Set frm = frmProdutividadeSaldo
        Case "mnuCnsEmpresaAvancada"
             Set frm = frmCnsAvancadaMob
        Case "mnuCnsNF"
             Set frm = frmNF
        Case "mnuNovaGIA"
             Set frm = frmNovaGIA
        Case "mnuCnsNFDoc"
             Set frm = frmDocNF
        Case "mnuCnsISSVarPago"
             Set frm = frmIssVariavelPago
        Case "mnuCnae"
             Set frm = frmCnaeNovo
        Case "mnuSuspende"
             Set frm = frmSuspReativ
        Case "mnuEmissaoDoc"
             If NomeDeLogin <> "SCHWARTZ" Then
                Liberado
                MsgBox "Acesso Negado!", vbCritical, "GTI"
                Exit Sub
             End If
             Set frm = frmNumeracaoDoc
        Case "mnuRelatorioMob1"
            Set frm = Nothing
            BuildReportMob1
        Case "mnuNotificaISS"
             Set frm = frmNotificaISS
        Case "mnuSalaEmp"
             Set frm = frmSalaEmpreendedor
        Case "mnuAlvara"
            Set frm = frmAlvaraNovo
        Case "mnuSituacaoAlvara"
            Set frm = frmSituacaoAlvara
        Case "mnuDevIssVar"
            Set frm = Nothing
            frmReport.ShowReport "ISSVARIAVELNAOPAGO", frmMdi.HWND, Me.HWND
        Case "mnuDevIssEst"
            Set frm = Nothing
            frmReport.ShowReport "ISSESTIMADONAOPAGO", frmMdi.HWND, Me.HWND
        Case "mnuRelDevTaxaLic"
            Set frm = Nothing
            frmReport.ShowReport "TAXALICENCANPAGA", frmMdi.HWND, Me.HWND
        Case "mnuRelDevTaxaLicAuto"
            Set frm = Nothing
            frmReport.ShowReport "TAXALICENCANPAGAAUTONOMO", frmMdi.HWND, Me.HWND
        Case "mnuRelDevTaxaLicAlvara"
            Set frm = Nothing
            frmReport.ShowReport "TAXALICENCANPAGA2", frmMdi.HWND, Me.HWND
        Case "mnuRelDevVigSanit"
            Set frm = Nothing
            frmReport.ShowReport "VIGILANCIANPAGA", frmMdi.HWND, Me.HWND
        Case "mnuRelDevedorGeral"
            Set frm = Nothing
            frmReport.ShowReport "DEVEDORISSGERAL", frmMdi.HWND, Me.HWND
'        Case "mnuRelDevISSEletro"
'            Set frm = frmReportMob1
        Case "mnuListaSN"
            Set frm = Nothing
        Case "mnuListaPagSN"
            Set frm = frmListaSN
        Case "mnuListaTL"
            Set frm = Nothing
            If MsgBox("Deseja exibir apenas o resumo?", vbYesNo + vbQuestion, "Tipo de Relat�rio") = vbYes Then
                frmReport.ShowReport2 "ATIVIDADETLA", frmMdi.HWND, Me.HWND
            Else
                frmReport.ShowReport2 "ATIVIDADETL", frmMdi.HWND, Me.HWND
            End If
        Case "mnuListaCnae"
            Set frm = Nothing
            frmReport.ShowReport2 "EmpresaCnae", frmMdi.HWND, Me.HWND
        Case "mnuListaTL3"
            Set frm = Nothing
            frmReport.ShowReport "ATIVIDADETL3", frmMdi.HWND, Me.HWND
        Case "mnuListaIssVE"
            Set frm = Nothing
            frmReport.ShowReport "ATIVIDADEISS", frmMdi.HWND, Me.HWND
        Case "mnuListaIssFixo"
            Set frm = Nothing
            frmReport.ShowReport "ATIVIDADEISSFIXO", frmMdi.HWND, Me.HWND
        Case "mnuListaVS"
            Set frm = Nothing
            frmReport.ShowReport "ATIVIDADEVS", frmMdi.HWND, Me.HWND
        Case "mnuRepParcNPIPTU"
            Set frm = Nothing
            frmReport.ShowReport "ReparcNaoPagoIPTU", frmMdi.HWND, Me.HWND
        Case "mnuRepParcNPISS"
            Set frm = Nothing
            frmReport.ShowReport "ReparcNaoPagoISS", frmMdi.HWND, Me.HWND
        Case "mnuMalaDireta"
            Set frm = Nothing
            frmMalaDiretaISS.show: frmMalaDiretaISS.ZOrder (0)
        Case "mnuEmpresaRua"
            Set frm = Nothing
            frmReport.ShowReport "EMPRESAORDEMRUA", frmMdi.HWND, Me.HWND
        Case "mnuEmpresaCNPJ"
            Set frm = Nothing
            frmReport.ShowReport "EMPRESAPORCNPJ", frmMdi.HWND, Me.HWND
        Case "mnuEmpresaAtividade"
            Set frm = Nothing
            frmReport.ShowReport "EMPRESAATIVIDADE", frmMdi.HWND, Me.HWND
        Case "mnuEmpresaContador"
            Set frm = Nothing
            frmReport.ShowReport "MOBILIARIOESCCONTABIL", frmMdi.HWND, Me.HWND
        Case "mnuEmpresaSocio"
            Set frm = Nothing
            frmReport.ShowReport "EMPRESASOCIO", frmMdi.HWND, Me.HWND
        Case "mnuRelEstimado"
            Set frm = Nothing
            frmReport.ShowReport "EMPRESAESTIMADO", frmMdi.HWND, Me.HWND
        Case "mnuRelMEI"
            Set frm = Nothing
            frmReport.ShowReport "EMPRESAMEI", frmMdi.HWND, Me.HWND
        Case "mnuIssPagoAtividade"
            Set frm = frmIssPagoAtividade
        Case "mnuResumoIssCCivil"
            Set frm = Nothing
            frmReport.ShowReport "ISSCONSTRUCAOCIVIL", frmMdi.HWND, Me.HWND
        Case "mnuNFEmitida"
            Set frm = Nothing
            frmReport.ShowReport3 "NF_EMITIDA", frmMdi.HWND, Me.HWND
        Case "mnuRelGiss"
            Set frm = frmResumoIss
        Case "mnuISSMensal"
            Set frm = frmISSMensal
    End Select
    If Not frm Is Nothing Then
        frm.show
        frm.ZOrder 0
    End If
    Liberado
Else
   lIndex = 0 ' cancelled the menu.
End If
End Sub

Private Sub cmdOutros_Click()
Dim z As Variant
lIndex = m_cMenuOutro.ShowPopupMenu(cmdOutros.Left, cmdOutros.Top + cmdOutros.Height, cmdOutros.Left, cmdOutros.Top, Me.ScaleWidth - cmdOutros.Left - cmdOutros.Width, cmdParametros.Top + cmdOutros.Height, False)
If (lIndex > 0) Then
    Ocupado
    Select Case m_cMenuOutro.ItemKey(lIndex)
    Case "mnuUser"
        Set frm = frmUsuario
    Case "mnuSegEvento"
        Set frm = frmEventSecurity
    Case "mnuAtribSeg"
        Set frm = frmSecurity
    Case "mnuSecurityByUser"
        Set frm = frmSecurityByUser
    Case "mnuBaixaIss"
        Set frm = Nothing
        BaixaEicon
    Case "mnuChangeLogin"
        Set frm = Nothing
        z = InputBox("Novo login?", "Change User", NomeDeLogin)
        If z <> "" Then
            NomeDeLogin = z
            frmMdi.Sbar.Panels(2).Text = "Usuario: " & NomeDeLogin
        End If
    Case "mnuSql"
        Set frm = frmSql
'    Case "mnuGravaFoto"
'        Set frm = frmGravaFoto
    Case "mnuGeraDebito"
        Set frm = frmGeraDebito
'    Case "mnuConversor"
 '       Set frm = frmConversor
   ' Case "mnuExporta"
   '     Set frm = frmExporta
    Case "mnuIntegrativa"
        Set frm = New frmArquivoIntegrativa
    Case "mnuEicon"
        Set frm = Nothing
        AtualizaEmpresa
    Case "mnuRelatorioAtendimento"
        Set frm = frmRelatObra
    Case "mnuConectaDB"
        Set frm = Nothing
        ConectaDBTeste
    Case "mnuAnexos"
'        If NomeDeLogin = "SCHWARTZ" Then
            Set frm = frmAnexos
 '       Else
  '          Exit Sub
   '     End If
    Case "mnuConectaBKP"
        Set frm = Nothing
'        ConectaDBBKP
    Case "mnuRegEndereco"
        Set frm = Nothing
        frmReport.ShowReport3 "REGATENDIMENTO_ENDERECO", frmMdi.HWND, Me.HWND
    Case "mnuTabelaEquipe"
        Set frm = New frmParamObra
        frm.sSigla = "EQ"
    Case "mnuTabelaFuncionario"
        Set frm = New frmParamObra
        frm.sSigla = "FC"
    Case "mnuFuncionarioEquipe"
        Set frm = frmTabelaFuncEquipe
    Case "mnutabelaAssunto"
        Set frm = New frmParamObra
        frm.sSigla = "AS"
    Case "mnutabelaAtendente"
        Set frm = New frmParamObra
        frm.sSigla = "AT"
    Case "mnutabelaTipoAtendimento"
        Set frm = New frmParamObra
        frm.sSigla = "TA"
    Case "mnuRegistroAtendimento"
        Set frm = frmRegistroAtendimento
    Case "mnuTabelaDocCP"
        Set frm = frmCPDoc
    Case "mnuTabelaDespachoCP"
        Set frm = frmCPDespacho
    Case "mnuTabelaCCustoCP"
        Set frm = frmCPCentroCusto
    Case "mnuTabelaAssuntoCP"
        Set frm = frmCPAssunto
    Case "mnuProcessoCP"
        Set frm = frmCPProcesso
    'Case "mnuCadCemiterio"
    '    Set frm = frmSerasa
        'Set frm = frmCadCemiterio
    End Select
    If Not frm Is Nothing Then
        frm.show
        frm.ZOrder 0
    End If
        Liberado
Else
    lIndex = 0 ' cancelled the menu.
End If
End Sub

Private Sub cmdParametros_Click()
lIndex = m_cMenuParam.ShowPopupMenu(cmdParametros.Left, cmdParametros.Top, cmdParametros.Left, cmdParametros.Top, Me.ScaleWidth - cmdParametros.Left - cmdParametros.Width, cmdParametros.Top + cmdParametros.Height, False)
Set frm = Nothing
If (lIndex > 0) Then
    Ocupado
    Select Case m_cMenuParam.ItemKey(lIndex)
        Case "mnuBairro"
            Set frm = frmBairro
        Case "mnuCidade"
            Set frm = frmCidade
        Case "mnuTitLog"
            Set frm = frmTitLogradouro
        Case "mnuTipoLog"
            Set frm = frmTipoLogradouro
        Case "mnuTabSistemaBenf"

        Case "mnuTabSistemaPPARC"
            Set frm = frmParamParcela
        Case "mnuTabSistemaTLAN"
            Set frm = frmLancamento
        Case "mnuTabSistemaTTRI"
            Set frm = frmTributo
        Case "mnuTabSistemaTTLA"
            Set frm = frmTributoLanc
        Case "mnuArtigoTributo"
            Set frm = frmTributoArtigo
        Case "mnuBanco"
            Set frm = frmBanco
        Case "mnuFeriado"
            Set frm = frmFeriados
        Case "mnuTabTributoAliq"
            Set frm = frmTributoAliquota
   End Select
   If Not frm Is Nothing Then
        frm.show
        frm.ZOrder 0
   End If
   Liberado
Else
    lIndex = 0 ' cancelled the menu.
End If
End Sub

Private Sub cmdPrincipal_Click()
lIndex = m_cMenuPrincipal.ShowPopupMenu(cmdPrincipal.Left, cmdPrincipal.Top, cmdPrincipal.Left, cmdPrincipal.Top, Me.ScaleWidth - cmdPrincipal.Left - cmdPrincipal.Width, cmdPrincipal.Top + cmdPrincipal.Height, False)
If (lIndex > 0) Then
    Ocupado
    Select Case m_cMenuPrincipal.ItemKey(lIndex)
    Case "mnuLog"
        Set frm = frmLog
    Case "mnuConfig"
        Set frm = frmConfig
    Case "mnuPrintBottom"
        Set frm = Nothing
        If m_cMenuPrincipal.Checked(m_cMenuPrincipal.IndexForKey("mnuPrintBottom")) = True Then
            m_cMenuPrincipal.Checked(m_cMenuPrincipal.IndexForKey("mnuPrintBottom")) = False
            SaveSetting "GTI", "PRINT", "BOTTOM", "N"
        Else
            m_cMenuPrincipal.Checked(m_cMenuPrincipal.IndexForKey("mnuPrintBottom")) = True
            SaveSetting "GTI", "PRINT", "BOTTOM", "S"
        End If
        sPrintBottom = GetSetting("GTI", "PRINT", "BOTTOM")
    Case "mnuSelectPrinter"
        Set frm = frmPrinterTest
    Case "mnuChangeUser"
        Set frm = Nothing
        frmLogin.show vbModal
    Case "mnuClose"
        Set frm = Nothing
        Unload Me
'    Case "mnuHist"
'        Set frm = Nothing
'        frmHist.show: frmHist.ZOrder 0: frmHist.lstLog.ListIndex = frmHist.lstLog.ListCount - 1
    End Select
    
    If Not frm Is Nothing Then
        frm.show
        frm.ZOrder 0
    End If
        Liberado
Else
    lIndex = 0 ' cancelled the menu.
End If
End Sub

Private Sub cmdProtocolo_Click()
lIndex = m_cMenuProt.ShowPopupMenu(cmdProtocolo.Left, cmdProtocolo.Top, cmdProtocolo.Left, cmdProtocolo.Top, Me.ScaleWidth - cmdProtocolo.Left - cmdProtocolo.Width, cmdProtocolo.Top + cmdProtocolo.Height, False)
If (lIndex > 0) Then
    Ocupado
    Select Case m_cMenuProt.ItemKey(lIndex)
    Case "mnuDespacho"
        Set frm = frmDespacho
    Case "mnuCentroCusto"
        Set frm = frmCentroCusto
    Case "mnuTabDocumento"
        Set frm = frmTabDocumento
    Case "mnuAssunto"
        Set frm = frmAssunto
    Case "mnuUsuarioProt"
        Set frm = frmUsuarioFora
    Case "mnuPermissaoProt"
        Set frm = frmUsuario
    Case "mnuProcesso"
        Set frm = frmProcesso
    Case "mnuProcessoArquivado"
        Set frm = frmProcessoArquivado
    Case "mnuEtiquetaProt"
        Set frm = frmLabelProtocolo
    Case "mnuResumoDiario"
        Set frm = frmResumoProtocolo
    Case "mnuPublicacao"
        Set frm = frmPublicaProcesso
    Case "mnuProcessoCC"
        Set frm = frmProcessoCCusto
    Case "mnuTramiteAberto"
        Set frm = Nothing
        frmReport.ShowReport "TRAMITEABERTOLOCAL2", frmMdi.HWND, Me.HWND
    Case "mnuTramiteAtraso"
        Set frm = frmTramiteAtraso
    Case "mnuTramiteEnviado"
        Set frm = frmProcessosEnviados
    Case "mnuProcessoAssunto"
        Set frm = Nothing
        frmReport.ShowReport "PROCESSOASSUNTO", frmMdi.HWND, Me.HWND
    Case "mnuProcessoAno"
        Set frm = Nothing
        frmReport.ShowReport2 "QTDEPROCESSOSANO", frmMdi.HWND, Me.HWND
    Case "mnuAssuntoDoc"
        Set frm = Nothing
        frmReport.ShowReport2 "ASSUNTO_DOC", frmMdi.HWND, Me.HWND
    End Select
    If Not frm Is Nothing Then
        frm.show
        frm.ZOrder 0
    End If
        Liberado
Else
    lIndex = 0 ' cancelled the menu.
End If
End Sub



Private Sub cmdTributo_Click()
lIndex = m_cMenuTrib.ShowPopupMenu(cmdTributo.Left, cmdTributo.Top, cmdTributo.Left, cmdTributo.Top, Me.ScaleWidth - cmdTributo.Left - cmdTributo.Width, cmdTributo.Top + cmdTributo.Height, False)
If (lIndex > 0) Then
    Ocupado
    Select Case m_cMenuTrib.ItemKey(lIndex)
        Case "mnuCalcGeral"
             Set frm = frmCalcGeral
        Case "mnuCalcIPTU"
             Set frm = frmCalculo
        Case "mnuCalcGeralISS"
             Set frm = frmCalcGeralISS
        Case "mnuCalculoCIP"
            Set frm = frmCalculoCIP
        Case "mnuArqLaser"
            Set frm = frmArquivoLaser
        Case "mnuOptanteDA"
            Set frm = frmOptanteDA
        Case "mnuCobranca"
            Set frm = frmCobranca
        Case "mnuOptanteDARel"
            Set frm = Nothing
            frmReport.ShowReport "optantes", frmMdi.HWND, Me.HWND
        Case "mnuAmostraCalculo"
            Set frm = Nothing
            frmReport.ShowReport2 "calculoiptu", frmMdi.HWND, Me.HWND
        Case "mnuImportaArq"
            'Set frm = frmArqBanco
            Set frm = frmImportaBanco
        Case "mnuBuscaArq"
            Set frm = frmBuscaDoc
        Case "mnuBaixaDebito"
  '          Liberado
'            MsgBox "em manuten��o"
            
 '           Exit Sub
            Set frm = frmPagAutomatico
        Case "mnuDebitoCompl"
            Set frm = frmDebitoCompl
        Case "mnuGuiaAmbulante"
            Set frm = frmGuiaAmbulante
        Case "mnuSituacaoTributo"
            Set frm = frmSituacaoTributo
        Case "mnuSituacaoTributaria"
            Set frm = frmSituacaoTributaria
        Case "mnuAnaliseReceita"
 '           Set frm = frmAnaliseReceita
'        Case "mnuAnaliseReceita2"
            Set frm = frmAnalise
        Case "mnuRelPagamento"
            Set frm = frmRelBanco
        Case "mnuNovaImportacao"
            Set frm = frmImportaBanco
        Case "mnuArrecadaSN"
            Set frm = Nothing
            frmReport.ShowReport "ARRECADACAOSNDAF", frmMdi.HWND, Me.HWND
        Case "mnuArrecadaSNDAF"
            Set frm = frmAnaliseReceita
        Case "mnuBaixaSN"
            Set frm = frmBaixaSN
        Case "mnuTabelaSelic"
            Set frm = frmTaxaSelic
        Case "mnuPagamentoDuplicidade"
            Set frm = frmPagamentoDuplicidade
        Case "mnuResumoISSCivil"
            Set frm = Nothing
            frmReport.ShowReport3 "ISSCCIVIL", frmMdi.HWND, Me.HWND
        Case "mnuSimulaSN"
            Set frm = frmSimulaSimples
        Case "mnuImportarSN"
            Set frm = frmOptanteSimples
        Case "mnuDecodificaMEI"
            Set frm = frmDecodificarMEI
        Case "mnuBuscaSN"
             Set frm = frmBuscaSN
        Case "mnuSNCnpj"
             Set frm = frmSimplesCnpj
        Case "mnuSNCnpjReceita"
             Set frm = frmSimplesCNPJ_Receita
        Case "mnuManAluguel"
            Set frm = frmManAluguel
        Case "mnuImpAluguel"
            Set frm = frmEmissaoAluguel
        Case "mnuDividaAtiva"
            Set frm = frmDividaAtiva
        Case "mnuEmiteLivro"
            Set frm = frmGeraLivro
        Case "mnuAjuizaAuto"
            Set frm = Nothing
            AjuizaAuto
        Case "mnuCartaCobranca"
            Set frm = frmCobrancaAmigavel
        Case "mnuRelAjuizamento"
            Set frm = frmRelAjuiza
        Case "mnuListaDevedor"
            Set frm = frmDevedor
        Case "mnuNotificacao"
            Set frm = frmNotificacao
        Case "mnu2vianotificacao"
            Set frm = frm2vianotificacaoiss
        Case "mnuNotificacao2"
            Set frm = frmNotificacao2
        Case "mnuAvisoDebito"
            Set frm = frmAvisoDebito
        Case "mnuDebitoAjPago"
            Set frm = Nothing
            frmReport.ShowReport "DEBITOAJPAGO", frmMdi.HWND, Me.HWND
        Case "mnuPagamentoCC"
            Set frm = Nothing
            frmReport.ShowReport3 "PAGAMENTOCARTACOBRANCA", frmMdi.HWND, Me.HWND
        
        Case "mnuDocEmitido"
            Set frm = Nothing
            frmReport.ShowReport2 "DOCUMENTOSEMITIDOS", frmMdi.HWND, Me.HWND
        Case "mnuComplementoPagto"
            Set frm = Nothing
            frmReport.ShowReport "COMPLEMENTOPAGTO", frmMdi.HWND, Me.HWND
        Case "mnuITBIObs"
            Set frm = Nothing
            frmReport.ShowReport "ITBIOBS", frmMdi.HWND, Me.HWND
        Case "mnuITBIRel"
            Set frm = Nothing
            BuildRelITBI
        Case "mnuCorrecaoCPF"
            Set frm = frmCorrecaoCPF
        Case "mnuPagoTributo"
            Set frm = frmValorPago
        Case "mnuGeraLote"
            Set frm = frmFichaCompensacaoLote

    End Select
    If Not frm Is Nothing Then
        frm.show
        frm.ZOrder 0
    End If
    Liberado
Else
   lIndex = 0 ' cancelled the menu.
End If
End Sub

Private Sub AjuizaAuto()
Dim nCodReduz1 As Long, nCodReduz2 As Long, sData As String, sTmp As String


z = InputBox("Digite o c�digo inicial.", "Aten��o")
If Val(z) = 0 Then
    MsgBox "Opera��o cancelada!", vbInformation, "Aten��o"
    Exit Sub
Else
    nCodReduz1 = CLng(z)
End If

z = InputBox("Digite o c�digo final.", "Aten��o")
If Val(z) = 0 Then
    MsgBox "Opera��o cancelada!", vbInformation, "Aten��o"
    Exit Sub
Else
    nCodReduz2 = CLng(z)
End If

If nCodReduz1 > nCodReduz2 Then
    MsgBox "C�digo inicial n�o pode ser maior que c�digo final. Opera��o cancelada!", vbInformation, "Aten��o"
    Exit Sub
End If

z = InputBox("Digite a data do ajuizamento.", "Aten��o")
If Not IsDate(z) Then
    MsgBox "Data inv�lida. Opera��o cancelada!", vbInformation, "Aten��o"
    Exit Sub
Else
    sData = z
End If

sTmp = "C�digo inicial: " & nCodReduz1 & vbCrLf & "C�digo final: " & nCodReduz2 & vbCrLf & "Data Ajuizamento: " & sData

If MsgBox("Confirme as informa��es" & vbCrLf & vbCrLf & sTmp, vbQuestion + vbYesNo, "Confirma��o") = vbYes Then
    Sql = "update debitoparcela set dataajuiza='" & Format(sData, "mm/dd/yyyy") & "' where codreduzido between " & nCodReduz1 & " and " & nCodReduz2
    Sql = Sql & " and statuslanc=3 and datainscricao is not null and dataajuiza is null and anoexercicio < " & (Year(Now) - 1) & " and numparcela>0"
    cn.Execute Sql, rdExecDirect

    MsgBox "Ajuizamento conclu�do", vbInformation, "Aten��o"
End If

End Sub


Private Sub MDIForm_Resize()




frDV.Left = Me.Width - 4300
imWorking.Left = Me.Width - 600
imOK.Left = Me.Width - 900
'cmdSerasa.Left = Me.Width - 2000
picStretched.Move 0, 0, ScaleWidth, ScaleHeight

' Copy the original picture into picStretched.
picStretched.PaintPicture picOriginal.Picture, 0, 0, picStretched.ScaleWidth, picStretched.ScaleHeight, 0, 0, picOriginal.ScaleWidth, picOriginal.ScaleHeight
        
' Set the MDI form's picture.
Picture = picStretched.Image

Unload frmCnsParcela
End Sub

Private Sub MDIForm_Activate()
Dim x As Integer, RdoAux As rdoResultset, Sql As String
lngTimer = 0



frTeste.Width = Screen.Width
If Not RunOnce Then
     
     
      frmLogin.show 1
    RunOnce = True
Else
    Unload frmLogin
End If

MDIForm_Resize
End Sub

Private Sub MDIForm_Load()

FlagServico = 0

If InStr(1, UCase(Command$), "-INTERNET", vbBinaryCompare) > 0 Then
    bDBInternet = True
Else
    bDBInternet = False
End If

Ocupado
RunOnce = False

Set gtiObj = New gtiProc.Tmuna
Dim lMajor As Long, lMinor As Long, lBuild As Long
MontaMenu
lMajor = App.Major
lMinor = App.Minor
lBuild = App.Revision
Me.Caption = Me.Caption & " - Vers�o: " & lMajor & "." & lMinor & "." & Format(lBuild, "000")
cmbDV.ListIndex = 0
'frmHist.Hide
FormHagana
Liberado
UpdateModule

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next

If MsgBox("Deseja  Sair do Sistema ?", vbQuestion + vbYesNo, "Confirma��o") = vbNo Then
   Cancel = 1
   Liberado
   Exit Sub
End If

frmChat.Timer1.Interval = 0
frmChat.Timer2.Interval = 0
bCloseChat = True
Unload frmChat

Sql = "UPDATE USUARIO SET LOGON=0 WHERE NOMELOGIN='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect
'Sql = "DELETE FROM MACHINES WHERE USUARIO='" & NomeDeLogin & "'"
'cn.Execute Sql, rdExecDirect
'Sql = "DELETE FROM MACHINES WHERE COMPUTER='" & NomeDoComputador & "'"
'cn.Execute Sql, rdExecDirect

Set m_cMenuPrincipal = Nothing
Set m_cMenuParam = Nothing
Set m_cMenuImob = Nothing
Set m_cMenuMob = Nothing
Set m_cMenuAtende = Nothing
Set m_cMenuProt = Nothing
Set m_cMenuOutro = Nothing

Set gtiObj = Nothing
Set DC = Nothing
Unload frmCnsParcela
modLg "Desconectado do Sistema"
modLg000
'If NomeDoComputador <> "BOJUTSU" Then
'    CloseApplication
'End If
End
End Sub

Private Sub mnuCertidaoIsencaoITBI_Click()
frmGuiaPratico4.show: frmGuiaPratico4.ZOrder 0
End Sub

Private Sub mnuCloseAll_Click()
Dim x As Integer
Inicio:
bCloseChat = True
For x = 0 To Forms.Count - 1
    If Forms(x).Name <> "frmMdi" And Forms(x).Name <> "frmHist" Then
       Unload Forms(x)
       GoTo Inicio:
    End If
Next

End Sub

Private Sub mnuCertidaoDebito_Click()
Dim z As Variant, sNumProc As String, nCodReduz As Long, qd As New rdoQuery, RdoAux As rdoResultset, aCodigo() As String
Dim XPos, YPos, x As Integer, nSuspenso As Long
XPos = Screen.Width / 2 - 3000
YPos = Screen.Height / 2 - 1000
nSuspenso = 1
z = InputBox("Digite o n� do processo com digito.", "Aten��o", "", XPos, YPos)
If Val(z) = 0 Then
    MsgBox "Opera��o cancelada!", vbInformation, "Aten��o"
    Exit Sub
Else
    sNumProc = z
End If
    
If InStr(1, sNumProc, "/", vbBinaryCompare) = 0 Then
    MsgBox "N� do processo inv�lido. Formato deve ser: N� do Processo/Ano.", vbCritical, "Aten��o"
    Exit Sub
End If

If Not IsNumeric(Right$(sNumProc, 4)) Then
    MsgBox "N� do processo inv�lido. O ano deve ter 4 digitos.", vbCritical, "Aten��o"
    Exit Sub
End If

If IsNumeric(Right$(sNumProc, 5)) Then
    MsgBox "N� do processo inv�lido. O ano deve ter 4 digitos.", vbCritical, "Aten��o"
    Exit Sub
End If

If Not IsNumeric(Left$(sNumProc, 1)) Then
    MsgBox "N� do processo inv�lido.", vbCritical, "Aten��o"
    Exit Sub
End If
    
    
z = InputBox("Digite o c�digo do(s) contribuinte(s) separados por v�rgula.", "Aten��o", "", XPos, YPos)
If Val(z) = 0 Then
    MsgBox "Opera��o cancelada!", vbInformation, "Aten��o"
    Exit Sub
Else
    aCodigo = Split(z, ",")
'    nCodReduz = z
End If
    
For x = 0 To UBound(aCodigo)
    'CodCidadao = nCodReduz
    CodCidadao = Val(aCodigo(x))
    If CodCidadao > 500000 Then
        MsgBox "N�o � poss�vel emitir certid�o de d�bito para c�digo cidad�o.", vbCritical, "Aten��o"
        Exit Sub
    End If
    NumeroProcesso = sNumProc
    sNumProc = Replace(sNumProc, "-", "")
    Set qd.ActiveConnection = cn
    On Error Resume Next
    RdoAux.Close
    On Error GoTo 0
    qd.Sql = "{ Call spCDB(?,?,?) }"
    qd(0) = CodCidadao
    qd(1) = sNumProc
    qd(2) = NomeDeLogin
    Set RdoAux = qd.OpenResultset(rdOpenKeyset)
    With RdoAux
        If .rdoColumns(3) = "S" Then nSuspenso = 1
        If .rdoColumns(0).value < 3 Or .rdoColumns(0).value = 7 Then
            MsgBox .rdoColumns(1).value, vbCritical, "Aten��o"
        Else
            If CodCidadao < 100000 Then
                If .rdoColumns(0).value = 3 Then
                    frmReport.ShowReport "CDBNEGIM", frmMdi.HWND, Me.HWND
                ElseIf .rdoColumns(0).value = 4 Then
                    frmReport.ShowReport "CDBPOSIM", frmMdi.HWND, Me.HWND
                ElseIf .rdoColumns(0).value = 5 Then
                    frmReport.ShowReport "CDBPSNIM", frmMdi.HWND, Me.HWND, nSuspenso
                ElseIf .rdoColumns(0).value = 6 Then
                    frmReport.ShowReport "CDBPNSIM", frmMdi.HWND, Me.HWND, nSuspenso
                End If
            ElseIf CodCidadao >= 100000 And CodCidadao < 500000 Then
                If .rdoColumns(0).value = 3 Then
                    frmReport.ShowReport "CDBNEGEM", frmMdi.HWND, Me.HWND
                ElseIf .rdoColumns(0).value = 4 Then
                    frmReport.ShowReport "CDBPOSEM", frmMdi.HWND, Me.HWND
                ElseIf .rdoColumns(0).value = 5 Then
                    frmReport.ShowReport "CDBPSNEM", frmMdi.HWND, Me.HWND, nSuspenso
                ElseIf .rdoColumns(0).value = 6 Then
                    frmReport.ShowReport "CDBPNSEM", frmMdi.HWND, Me.HWND, nSuspenso
                End If
            ElseIf CodCidadao > 500000 Then
                If .rdoColumns(0).value = 3 Then
                    frmReport.ShowReport "CDBNEGCD", frmMdi.HWND, Me.HWND
                ElseIf .rdoColumns(0).value = 4 Then
                    frmReport.ShowReport "CDBPOSCD", frmMdi.HWND, Me.HWND
                ElseIf .rdoColumns(0).value = 5 Then
                    frmReport.ShowReport "CDBPSNCD", frmMdi.HWND, Me.HWND, nSuspenso
                ElseIf .rdoColumns(0).value = 6 Then
                    frmReport.ShowReport "CDBPNSCD", frmMdi.HWND, Me.HWND, nSuspenso
                End If
            End If
        End If
        .Close
    End With
Next
modLg "Emiss�o de Certid�o de D�bito - C�digo: " & CodCidadao & " - Processo n�: " & sNumProc
CodCidadao = 0
End Sub

Private Sub mnuCertidaoDemolicao_Click()
frmCertidao.show
frmCertidao.lblTipo.Caption = "CERTID�O DE DEMOLI��O"
frmCertidao.lblCodCert.Caption = 5
End Sub

Private Sub mnuCertidaoEndereco_Click()
frmCertidao.show
frmCertidao.lblTipo.Caption = "CERTID�O DE ENDERE�O ATUALIZADO"
frmCertidao.lblCodCert.Caption = 2
End Sub

Private Sub mnuCertidaoIsencao_Click()
frmCertidao.show
frmCertidao.lblTipo.Caption = "CERTID�O DE ISEN��O"
frmCertidao.lblCodCert.Caption = 6
End Sub

Private Sub mnuCertidaoValorVenal_Click()
frmCertidao.show
frmCertidao.lblTipo.Caption = "CERTID�O DE VALOR VENAL"
frmCertidao.lblCodCert.Caption = 3
End Sub

Private Sub ConectaDBTeste()
Dim DataSourceName As String
Dim DatabaseName As String
Dim Description As String
Dim DriverPath As String
Dim DriverName As String
Dim LastUser As String
Dim Regional As String
Dim Server As String

Dim lResult As Long
Dim hKeyHandle As Long

DataSourceName = "odbcTribTeste"
DatabaseName = "TributacaoTeste"
Description = "Base de Testes do GTI"
DriverPath = "<path to your SQL Server driver>"
LastUser = ""
Server = IPServer
DriverName = "SQL Server"

'Create the new DSN key.

lResult = RegCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\" & _
     DataSourceName, hKeyHandle)

'Set the values of the new DSN key.

lResult = RegSetValueEx(hKeyHandle, "Database", 0&, REG_SZ, _
   ByVal DatabaseName, Len(DatabaseName))
lResult = RegSetValueEx(hKeyHandle, "Description", 0&, REG_SZ, _
   ByVal Description, Len(Description))
lResult = RegSetValueEx(hKeyHandle, "Driver", 0&, REG_SZ, _
   ByVal DriverPath, Len(DriverPath))
lResult = RegSetValueEx(hKeyHandle, "LastUser", 0&, REG_SZ, _
   ByVal LastUser, Len(LastUser))
lResult = RegSetValueEx(hKeyHandle, "Server", 0&, REG_SZ, _
   ByVal Server, Len(Server))

'Close the new DSN key.

lResult = RegCloseKey(hKeyHandle)

'Open ODBC Data Sources key to list the new DSN in the ODBC Manager.
'Specify the new value.
'Close the key.

lResult = RegCreateKey(HKEY_LOCAL_MACHINE, _
   "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", hKeyHandle)
lResult = RegSetValueEx(hKeyHandle, DataSourceName, 0&, REG_SZ, _
   ByVal DriverName, Len(DriverName))
lResult = RegCloseKey(hKeyHandle)

End Sub




Private Sub Deca()
Dim z As Variant, Sql As String, RdoAux As rdoResultset, frm As Object
z = InputBox("Digite o c�digo da empresa.", "Informa��o requerida")
z = Val(z)
If Val(z) = 0 Then
    Set frm = frmDeca
    frm.nCodigoEmpresa = CLng(z)
    frmDeca.show
    Exit Sub
End If

If Val(z) < 100000 Or Val(z) >= 700000 Then
    MsgBox "C�digo inv�lido", vbCritical, "Erro"
    Exit Sub
End If

If Val(z) >= 100000 And Val(z) <= 300000 Then
    Sql = "SELECT CODIGOMOB FROM MOBILIARIO WHERE CODIGOMOB=" & Val(z)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
            MsgBox "C�digo de empresa inv�lido", vbCritical, "Erro"
           .Close
            Exit Sub
        End If
       .Close
    End With
ElseIf Val(z) >= 500000 And Val(z) <= 600000 Then
    Sql = "SELECT CODCIDADAO FROM CIDADAO WHERE CODCIDADAO=" & Val(z)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
            MsgBox "C�digo cidad�o inv�lido", vbCritical, "Erro"
           .Close
            Exit Sub
        End If
       .Close
    End With
End If

Set frm = frmDeca
frm.nCodigoEmpresa = CLng(z)
frmDeca.show

End Sub

Private Sub ImportarSN()
Dim Sql As String, RdoAux As rdoResultset, nCodReduz As Long, RdoAux2 As rdoResultset
Dim sDataIni As String, sDataFim As String
Exit Sub
Ocupado
Sql = "DELETE FROM PERIODOSN"
cn.Execute Sql, rdExecDirect

Sql = "SELECT * FROM IMPORTSN"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       DoEvents
        Sql = "SELECT CODIGOMOB FROM MOBILIARIO WHERE CNPJ LIKE '" & !Cnpj & "%'"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            nCodReduz = RdoAux2!codigomob
        Else
            GoTo Proximo
        End If
        sDataIni = Right$(!Inicio, 2) & "/" & Mid$(!Inicio, 5, 2) & "/" & Left$(!Inicio, 4)
        sDataFim = ""
        If Not IsNull(!FINAL) Then
            If Trim(!FINAL) <> "" And Trim(!FINAL) <> "X" Then
                sDataFim = Right$(!FINAL, 2) & "/" & Mid$(!FINAL, 5, 2) & "/" & Left$(!FINAL, 4)
            End If
        End If
        
        If sDataFim <> "" Then
            Sql = "INSERT PERIODOSN (CODIGO,DATAINI,DATAFIM) VALUES(" & nCodReduz & ",'" & Format(sDataIni, "mm/dd/yyyy") & "','"
            Sql = Sql & Format(sDataFim, "mm/dd/yyyy") & "')"
        Else
            Sql = "INSERT PERIODOSN (CODIGO,DATAINI) VALUES(" & nCodReduz & ",'" & Format(sDataIni, "mm/dd/yyyy") & "')"
        End If
        cn.Execute Sql, rdExecDirect
Proximo:
       .MoveNext
    Loop
   .Close
End With
Liberado
MsgBox "fim"
End Sub

Private Sub mnuRenovaAlvara_Click()
'frmReport.ShowReport2 "ALVARARENOVAVICE", frmMdi.HWND, Me.HWND
'frmReport.ShowReport2 "ALVARARENOVA", frmMdi.HWND, Me.HWND
End Sub

Private Sub Sbar_PanelClick(ByVal Panel As MSComctlLib.Panel)
Dim sRet As String

If bLocal Then
    Exit Sub
End If
'On Error Resume Next
If Panel.Index = 6 Then
    sRet = RetEventUserForm("frmDataBase")
    If InStr(1, sRet, "001", vbBinaryCompare) > 0 Then
        frmDataBase.show vbModeless
        frmDataBase.Mv.Day = Val(Mid$(frmMdi.Sbar.Panels(6).Text, 12, 2))
        frmDataBase.Mv.Month = Val(Mid$(frmMdi.Sbar.Panels(6).Text, 15, 2))
        frmDataBase.Mv.Year = Val(Right$(frmMdi.Sbar.Panels(6).Text, 4))
        frmDataBase.lblDB.Caption = "Data Base: " & frmDataBase.Mv.Day & "/" & frmDataBase.Mv.Month & "/" & frmDataBase.Mv.Year
    Else
        MsgBox "Voc� n�o tem permiss�o para alterar a Data Base", vbCritical, "Aten��o"
    End If
End If

End Sub


Private Sub Timer1_Timer()
lngTimer = lngTimer + 1
If NomeDeLogin = "SCHWARTZ" Or NomeDeLogin = "RITA" Or NomeDeLogin = "LUIZH" Or NomeDeLogin = "RENATA" Or NomeDeLogin = "ROSE" Then
    If lngTimer > 1200 Then
        
        lngTimer = 0
    End If
End If

End Sub

Private Sub Timer2_Timer()
Dim Sql As String, RdoPrm As rdoResultset, sDataBase As String, sOldData As String
If NomeDeLogin = "SCHWARTZ" Then
    Exit Sub
End If

lngTimer2 = lngTimer2 + 1
If lngTimer <= 120 Then Exit Sub
If cn.Connect = "" Then
   lngTimer2 = 0
   Exit Sub
End If
lngTimer2 = 0

On Error Resume Next
If InStr(1, cn.Connect, "TributacaoTeste", vbBinaryCompare) > 0 Then
    Sql = "USE TributacaoTeste"
ElseIf InStr(1, cn.Connect, "TributacaoBKP", vbBinaryCompare) > 0 Then
    Sql = "USE TributacaoBKP"
Else
    Sql = "USE Tributacao"
End If
cn.Execute Sql, rdExecDirect
If NomeDeLogin = "SCHWARTZ" Then
    Exit Sub
End If
Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='DATABASE'"
Set RdoPrm = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoPrm
    If .RowCount = 0 Then
       Sql = "INSERT PARAMETROS(NOMEPARAM,VALPARAM) VALUES('DATABASE'" & ",'" & CStr(Format(Now, "dd/mm/yyyy")) & "')"
       cn.Execute Sql, rdExecDirect
       sDataBase = CStr(Format(Now, "dd/mm/yyyy"))
    Else
       sDataBase = !valparam
    End If
   .Close
End With
sOldData = Right$(frmMdi.Sbar.Panels(6).Text, 10)
If InStr(1, cn.Connect, "Tributacao_", vbBinaryCompare) = 0 And InStr(1, cn.Connect, "TributacaoBKP", vbBinaryCompare) = 0 Then
    If sDataBase <> sOldData Then
       MsgBox "A Data Base foi atualizada para " & sDataBase, vbInformation, "ATEN��O !!!"
       frmMdi.Sbar.Panels(6).Text = "Data Base: " & sDataBase
    End If
End If

'Sql = " select * from eicon_timer"
'Set RdoPrm = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'If RdoPrm!close_app Then
'    Close_GTI_Server
'End If
'RdoPrm.Close

UpdateModule


End Sub

Private Sub Timer3_Timer()
On Error Resume Next
If FlagServico >= 15 Then
    If NomeDeLogin = "RODRIGOC" Or NomeDeLogin = "LUIZH" Or NomeDeLogin = "NOELI" Or NomeDeLogin = "LEANDRO" Or NomeDeLogin = "RITA" Or NomeDeLogin = "DANIELAR" Or NomeDeLogin = "SCHWARTZ" Then
        frmServico.show 1
        frmServico.Refresh
        frmServico.btVerificar_Click
        Unload frmServico
    End If
    FlagServico = 0
End If
FlagServico = FlagServico + 1

End Sub

Private Sub txtDV_Change()
lblDV.Caption = "?"
End Sub

Private Sub MalaDiretaRural()
Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, Sql As String
Dim xId As Long, nNumRec As Long, nCodLogr As Long, sCodInscricao As String, sContribuinte As String
Dim sEnd As String, nNum As Integer, sCep As String, sCompl As String, sBairro As String
Dim sEndEntrega As String, sBairroEntrega As String, sCidEntrega As String, sCepEntrega As String, sUFEntrega As String, sNumEntrega As String
Dim z As Variant

Sql = "SELECT DISTINCT CODREDUZIDO,PROPRIETARIO From CADASTRORURAL ORDER BY CODREDUZIDO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Sql = "SELECT cidadao.codcidadao, cidadao.numimovel, cidadao.complemento, cidadao.codbairro, cidadao.codcidade, cidadao.siglauf, cidade.desccidade, "
        Sql = Sql & "bairro.descbairro, cidadao.codlogradouro, vwLOGRADOURO.ABREVTIPOLOG, vwLOGRADOURO.ABREVTITLOG,vwLOGRADOURO.NOMELOGRADOURO,cidadao.nomecidadao,"
        Sql = Sql & "cidadao.cep, cidadao.nomelogradouro AS Rua FROM cidadao LEFT OUTER JOIN cidade ON "
        Sql = Sql & "cidadao.siglauf = cidade.siglauf AND cidadao.codcidade = cidade.codcidade LEFT OUTER JOIN bairro ON cidadao.siglauf = bairro.siglauf AND "
        Sql = Sql & "cidadao.codcidade = bairro.codcidade AND cidadao.codbairro = bairro.codbairro LEFT OUTER JOIN vwLOGRADOURO ON cidadao.codlogradouro = vwLOGRADOURO.CODLOGRADOURO "
        Sql = Sql & "Where Cidadao.CodCidadao = " & RdoAux!Proprietario
        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux3
            sCodInscricao = Format(RdoAux!CODREDUZIDO, "000000")
            sContribuinte = SubNull(!nomecidadao)
            If IsNull(!NomeLogradouro) Then
                sEnd = !Rua & CStr(SubNull(!NUMIMOVEL))
            Else
                sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & SubNull(!NomeLogradouro) & " N� " & CStr(SubNull(!NUMIMOVEL))
            End If
            sCompl = SubNull(!Complemento)
'            If IsNull(!DescBairro) Then
'                sBairro = SubNull(!NOMEBairro)
'            Else
                sBairro = SubNull(!DescBairro)
 '           End If
  '          If IsNull(!desccidade) Then
   '             sCidade = SubNull(!NomeCidade)
   '         Else
                sCidade = SubNull(!descCidade)
    '        End If
            sCep = SubNull(!Cep)
            sUF = SubNull(!SiglaUF)
            If sCidade = "JABOTICABAL" And Val(SubNull(!CodLogradouro)) > 0 Then
                sCep = RetornaCEP(!CodLogradouro, !NUMIMOVEL)
            End If
            .Close
        End With
        sCompl = SubNull(Left(sCompl, 20))
        'sBairro = SubNull(!DescBairro)
    
        sEndEntrega = sEnd
        sBairroEntrega = sBairro
        sCidEntrega = sCidade
        sCepEntrega = sCep
        sComplEntrega = sCompl
        sUFEntrega = sUF
        
        Sql = "INSERT ETIQUETAGTI (USUARIO,SEQ,CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5) VALUES('"
        Sql = Sql & NomeDeLogin & "'," & xId & ",'" & sCodInscricao & "','" & Mask(sContribuinte) & "','"
        Sql = Sql & Left(sEndEntrega & " " & sComplEntrega, 60) & "','" & sCepEntrega & "   " & sBairroEntrega & "','" & sCidEntrega & "   " & sUFEntrega & "')"
        cn.Execute Sql, rdExecDirect
        xId = xId + 1
PROXIMO2:
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub UpdateModule()
Dim sPathLocal As String, sPathUpdate As String, aLocalFile() As FileData, aRemoteFile() As FileData
Dim nFilePos As Integer, nSearch As Integer, dFileDate As Date, sFileName As String, bAchou As Boolean
Exit Sub
ReDim aLocalFile(0): ReDim aRemoteFile(0)

If Dir(App.Path & "\CONFIG.INI") = "" Then
    Open App.Path & "\CONFIG.INI" For Output As #1
    Print #1, App.Path
    Print #1, "\\192.168.200.130\ATUALIZAGTI"
    Close #1
End If

Open App.Path & "\CONFIG.INI" For Input As #1
Input #1, sPathLocal
Input #1, sPathUpdate
Close #1

Ocupado
For nFilePos = 0 To File1.ListCount - 1
    ReDim Preserve aLocalFile(UBound(aLocalFile) + 1)
    aLocalFile(UBound(aLocalFile)).sName = File1.List(nFilePos)
    aLocalFile(UBound(aLocalFile)).dDate = FileDateTime(sPathLocal & "\REPORT\" & File1.List(nFilePos))
Next

File1.Path = sPathUpdate
For nFilePos = 0 To File1.ListCount - 1
    ReDim Preserve aRemoteFile(UBound(aRemoteFile) + 1)
    aRemoteFile(UBound(aRemoteFile)).sName = File1.List(nFilePos)
    aRemoteFile(UBound(aRemoteFile)).dDate = FileDateTime(sPathUpdate & "\" & File1.List(nFilePos))
Next
For nFilePos = 1 To UBound(aRemoteFile)
    DoEvents
    sFileName = aRemoteFile(nFilePos).sName
    bAchou = False
    For nSearch = 1 To UBound(aLocalFile)
        If UCase(aLocalFile(nSearch).sName) = UCase(sFileName) Then
            bAchou = True
            Exit For
        End If
    Next
    If Not bAchou Then
        FileCopy sPathUpdate & "\" & sFileName, sPathLocal & "\REPORT\" & sFileName
    Else
        dFileDate = aRemoteFile(nFilePos).dDate
        If aRemoteFile(nFilePos).dDate > aLocalFile(nSearch).dDate Then
            FileCopy sPathUpdate & "\" & sFileName, sPathLocal & "\REPORT\" & sFileName
        End If
    End If
Next
Liberado
If FileDateTime(sPathUpdate & "\GTI.EXE") > FileDateTime(sPathLocal & "\GTI.EXE") Then
    MsgBox "Existe uma nova vers�o do sistema GTI." & vbCrLf & "O sistema ser� atualizado automaticamente na pr�xima inicializa��o", vbInformation, "Informa��o de atualiza��o"
End If

End Sub

Private Sub MontaMenu()

Set m_cMenuPrincipal = New cPopupMenu
With m_cMenuPrincipal
    .hwndOwner = Me.HWND
    .GradientHighlight = True
    .ActiveMenuBackgroundColor = &HFFFF80
    .ButtonHighlight = True
    .InActiveMenuForeColor = &H400000
    .MenuBackgroundColor = &H8000000F
    i = .AddItem("Log do Sistema", "", , , , , , "mnuLog")
    i = .AddItem("Prefer�ncias", "", , , , , , "mnuConfig")
    i = .AddItem("Impressora", "", , , , , , "mnuPrinter")
    h = .AddItem("Imprimir na bandeja inferior", "", , i, , , , "mnuPrintBottom")
    h = .AddItem("Alterar impressora padr�o", "", , i, , , , "mnuSelectPrinter")
  '  i = .AddItem("Hist�rico do Sistema", "", , , , , , "mnuHist")
    i = .AddItem("-", "", , , , , , "mnuSep1")
    i = .AddItem("Trocar de usu�rio", "", , , , , , "mnuChangeUser")
    i = .AddItem("Sair do Sistema", "", , , , , , "mnuClose")
End With
   
Set m_cMenuParam = New cPopupMenu
With m_cMenuParam
    .hwndOwner = Me.HWND
    .GradientHighlight = True
    .ActiveMenuBackgroundColor = &HFFFF80
    .ButtonHighlight = True
    .InActiveMenuForeColor = &H400000
    .MenuBackgroundColor = &H8000000F
    i = .AddItem("Territorial", "", 1, , , , , "mnuTerritorial")
    h = .AddItem("Bairro", "", 1, i, , , , "mnuBairro")
    h = .AddItem("Cidade", "", 1, i, , , , "mnuCidade")
    h = .AddItem("Tipo Logradouro", "", 1, i, , , , "mnuTipoLog")
    h = .AddItem("T�tulo Logradouro", "", 1, i, , , , "mnuTitLog")
    i = .AddItem("Tribut�rio", "", 1, , , , , "mnuTribut�rio")
    h = .AddItem("Lan�amentos", "", 1, i, , , , "mnuTabSistemaTLAN")
    h = .AddItem("Pre�os P�blicos", "", 1, i, , , , "mnuTabTributoAliq")
    h = .AddItem("Tributos", "", 1, i, , , , "mnuTabSistemaTTRI")
    h = .AddItem("Tributos por Lan�amento", "", 1, i, , , , "mnuTabSistemaTTLA")
    h = .AddItem("Artigo por Tributo", "", 1, i, , , , "mnuArtigoTributo")
    h = .AddItem("Vencimento das Parcelas", "", 1, i, , , , "mnuTabSistemaPPARC")
    h = .AddItem("Tabela de UFIR", "", 1, i, , , , "mnuTabSistemaUfir")
    i = .AddItem("Outros Par�metros", "", 1, , , , , "mnuOutrosParam")
    h = .AddItem("Bancos", "", 1, i, , , , "mnuBanco")
    h = .AddItem("Feriados", "", 1, i, , , , "mnuFeriado")
End With
   
Set m_cMenuImob = New cPopupMenu
With m_cMenuImob
    .hwndOwner = Me.HWND
    .GradientHighlight = True
    .ActiveMenuBackgroundColor = &HFFFF80
    .ButtonHighlight = True
    .InActiveMenuForeColor = &H400000
    .MenuBackgroundColor = &H8000000F
    i = .AddItem("Cadastro", "", 1, , , , , "mnuCadastro")
    h = .AddItem("Im�vel", "", 1, i, , , , "mnuCadImob")
    h = .AddItem("Condom�nio", "", 1, i, , , , "mnuCondominio")
    h = .AddItem("Logradouro", "", 1, i, , , , "mnuLogr")
    h = .AddItem("Segmento de Logradouro", "", 1, i, , , , "mnuSegmentoLogradouro")
    h = .AddItem("Face de Quadra", "", 1, i, , , , "mnuFaceQuadra")
    i = .AddItem("Consulta", "", 1, , , , , "mnuConsultaImovel")
    h = .AddItem("Consulta de Im�vel", "", 1, i, , , , "mnuCnsImovel")
    h = .AddItem("Consulta avan�ada de im�veis", "", 1, i, , , , "mnuCnsAvancadaImob")
    h = .AddItem("Detalhes do im�vel", "", 1, i, , , , "mnuDetImovel")
    h = .AddItem("Resumo Imobili�rio", "", 1, i, , , , "mnuResumoImovel")
    i = .AddItem("Atividades", "", 1, , , , , "mnuAtividadeImovel")
    h = .AddItem("Desdobro", "", 1, i, , , , "mnuDesmem")
    h = .AddItem("Unifica��o de im�veis", "", 1, i, , , , "mnuUnificacao")
    h = .AddItem("Imunidade/Isen��o", "", 1, i, , , , "mnuImun")
    h = .AddItem("Desdobro de carn�", "", 1, i, , , , "mnuDesdobroCarne")
    h = .AddItem("Valor venal declarado", "", 1, i, , , , "mnuVVDeclarado")
    h = .AddItem("Simulado PG", "", 1, i, , , , "mnuSimuladoPG")
    h = .AddItem("Tipo de propriet�rio", "", 1, i, , , , "mnuEspolio")
    h = .AddItem("Corre��o de Bairros", "", 1, i, , , , "mnuCorrigeBairro")
    i = .AddItem("Relat�rios", "", 1, , , , , "mnuRelatorioImob")
    h = .AddItem("Certid�o de averba��o", "", 1, i, , , , "mnuAverbacao")
    h = .AddItem("Rol dos imoveis cadastrados", "", 1, i, , , , "mnuRolImovel")
    h = .AddItem("Devedor anual de IPTU", "", 1, i, , , , "mnuDevedorIPTU")
    h = .AddItem("Mala direta p/ro�ada", "", 1, i, , , , "mnuMalaDiretaRocada")
    h = .AddItem("Mala direta p/ISS Constr.Civil", "", 1, i, , , , "mnuMalaDiretaISSCCivil")
    h = .AddItem("Lan�amentos de ro�ada", "", 1, i, , , , "mnuLancRocada")
    i = .AddItem("Cadastro Rural", "", 1, , , , , "mnuGTIRural")
    h = .AddItem("Propriedades rurais", "", 1, i, , , , "mnuCadRural")
    h = .AddItem("Produtos rurais", "", 1, i, , , , "mnuProdutoRural")
    h = .AddItem("Estradas rurais", "", 1, i, , , , "mnuEstradaRural")
    h = .AddItem("Relat�rios", "", 1, i, , , , "mnuRural")
    i = .AddItem("Rela��o das propriedades cadastradas", "", 1, h, , , , "mnuRelCadRural")
    i = .AddItem("Dados cadastrais das propriedades", "", 1, h, , , , "mnuRelCadRuralFull")
    i = .AddItem("Rela��o das propriedades por faixa", "", 1, h, , , , "mnuRelCadRuralFaixa")
    i = .AddItem("Rela��o das culturas por propriedade", "", 1, h, , , , "mnuRelProdRural")
    i = .AddItem("Simula��o de c�lculo", "", 1, h, , , , "mnuSimulaRural")
    i = .AddItem("Altera��o no cadastro", "", 1, h, , , , "mnuEventoRural")
    i = .AddItem("Mala direta", "", 1, h, , , , "mnuMalaDiretaRural")
    i = .AddItem("Sistema de Cadastro de Obras", "", 1, , , , , "mnuOutroImob")
    h = .AddItem("Integra��o com o SisobraPref", "", 1, i, , , , "mnuSisObraPref")
End With

Set m_cMenuMob = New cPopupMenu
With m_cMenuMob
    .hwndOwner = Me.HWND
    .GradientHighlight = True
    .ActiveMenuBackgroundColor = &HFFFF80
    .ButtonHighlight = True
    .InActiveMenuForeColor = &H400000
    .MenuBackgroundColor = &H8000000F
    i = .AddItem("Cadastro", "", 1, , , , , "mnuCadastroMob")
    h = .AddItem("Empresas", "", 1, i, , , , "mnuCadMobiliario")
    h = .AddItem("Escritorio Contabil", "", 1, i, , , , "mnuEscContab")
    h = .AddItem("Hor�rio de funcionamento por atividade", "", 1, i, , , , "mnuHorarioFuncionamento")
    h = .AddItem("Atividades do mobili�rio", "", 1, i, , , , "mnuAtividadeCadMob")
    i = .AddItem("Taxa de Licen�a", "", 1, h, , , , "mnuTabAtivTL")
    i = .AddItem("Cobran�a de ISS", "", 1, h, , , , "mnuTabAtivISS")
    i = .AddItem("Vigil�ncia Sanit�ria", "", 1, h, , , , "mnuVigSan")
    i = .AddItem("Consulta", "", 1, , , , , "mnuConsultaMob")
    h = .AddItem("Consulta de empresas", "", 1, i, , , , "mnuCnsEmpresa")
    h = .AddItem("Consulta avan�ada de empresas", "", 1, i, , , , "mnuCnsEmpresaAvancada")
    h = .AddItem("Notas emitidas", "", 1, i, , , , "mnuCnsNF")
    h = .AddItem("Notas fiscais por documento", "", 1, i, , , , "mnuCnsNFDoc")
    h = .AddItem("ISS vari�vel pago nos �ltimos 5 anos", "", 1, i, , , , "mnuCnsISSVarPago")
    h = .AddItem("Tabela CNAE Fiscal 2.0", "", 1, i, , , , "mnuCnae")
    i = .AddItem("Atividades", "", 1, , , , , "mnuAtividadeMob")
    h = .AddItem("Suspens�o/Reativa��o", "", 1, i, , , , "mnuSuspende")
    h = .AddItem("Notifica��o de ISS eletr�nico", "", 1, i, , , , "mnuNotificaISS")
    h = .AddItem("Sala do empreendedor", "", 1, i, , , , "mnuSalaEmp")
    h = .AddItem("Nova GIA", "", 1, i, , , , "mnuNovaGIA")
    h = .AddItem("Integra��o VRE", "", 1, i, , , , "mnuVRE")
    h = .AddItem("Importar dados da RedeSim", "", 1, i, , , , "mnuVRERedeSim")
    j = .AddItem("Produtividade", "", 1, , , , , "mnuProdutividade")
    h = .AddItem("Cadastro de eventos", "", 1, j, , , , "mnuProdutEvento")
    h = .AddItem("Controle di�rio de tarefas", "", 1, j, , , , "mnuProdutTarefa")
    h = .AddItem("Extrato Mensal", "", 1, j, , , , "mnuProdutExtratoMes")
    h = .AddItem("Saldo Mensal", "", 1, j, , , , "mnuProdutSaldoMes")
    i = .AddItem("Relat�rios", "", 1, , , , , "mnuRelatorioMob")
    h = .AddItem("Alvar�s", "", 1, i, , , , "mnuAlvaras")
    z = .AddItem("Alvar� de funcionamento", "", 1, h, , , , "mnuAlvara")
    z = .AddItem("Situa��o dos alvar�s", "", 1, h, , , , "mnuSituacaoAlvara")
    h = .AddItem("Relat�rios de devedores", "", 1, i, , , , "mnuDevedores")
    z = .AddItem("Devedores ISS Vari�vel", "", 1, h, , , , "mnuDevIssVar")
    z = .AddItem("Devedores ISS Estimado", "", 1, h, , , , "mnuDevIssEst")
    z = .AddItem("Devedores Taxa de Licen�a (industria e com�rcio)", "", 1, h, , , , "mnuRelDevTaxaLic")
    z = .AddItem("Devedores Taxa de Licen�a (outros)", "", 1, h, , , , "mnuRelDevTaxaLicAuto")
    z = .AddItem("Devedores Taxa de Licen�a (Alvar�)", "", 1, h, , , , "mnuRelDevTaxaLicAlvara")
    z = .AddItem("Devedores vigil�ncia sanit�ria", "", 1, h, , , , "mnuRelDevVigSanit")
    z = .AddItem("Rela��o de devedores geral", "", 1, h, , , , "mnuRelDevedorGeral")
    h = .AddItem("Optantes do Simples Nacional", "", 1, i, , , , "mnuSN")
    z = .AddItem("Lista de optantes do simples nacional", "", 1, h, , , , "mnuListaSN")
    z = .AddItem("Lista de pagamento dos opt.do simples nacional", "", 1, h, , , , "mnuListaPagSN")
    z = .AddItem("Lista de devedores do simples nacional (Tx.Lic. e Vig.Sanit.)", "", 1, h, , , , "mnuRelatorioMob1")
    h = .AddItem("Lista de atividades", "", 1, i, , , , "mnuListaAtiv")
    z = .AddItem("Empresas por CNAE", "", 1, h, , , , "mnuListaCnae")
    z = .AddItem("Taxa de Licen�a", "", 1, h, , , , "mnuListaTL")
    z = .AddItem("Taxa de Licen�a por Logradouro", "", 1, h, , , , "mnuListaTL3")
    z = .AddItem("ISS Vari�vel/Estimado", "", 1, h, , , , "mnuListaIssVE")
    z = .AddItem("ISS Fixo", "", 1, h, , , , "mnuListaIssFixo")
    z = .AddItem("Vigil�ncia sanit�ria", "", 1, h, , , , "mnuListaVS")
    h = .AddItem("Parcelamentos n�o pagos", "", 1, i, , , , "mnurelParcNPago")
    z = .AddItem("Parcelamentos de IPTU", "", 1, h, , , , "mnuRepParcNPIPTU")
    z = .AddItem("Parcelamentos de ISS", "", 1, h, , , , "mnuRepParcNPISS")
    h = .AddItem("Mala direta", "", 1, i, , , , "mnuMalaDireta")
    h = .AddItem("Arrecada��o mensal de ISS", "", 1, i, , , , "mnuISSMensal")
    h = .AddItem("Relat�rio de empresas por logradouro", "", 1, i, , , , "mnuEmpresaRua")
    h = .AddItem("Relat�rio de empresas por CNPJ", "", 1, i, , , , "mnuEmpresaCNPJ")
    h = .AddItem("Relat�rio de empresas por atividade", "", 1, i, , , , "mnuEmpresaAtividade")
    h = .AddItem("Relat�rio de empresas por contador", "", 1, i, , , , "mnuEmpresaContador")
    h = .AddItem("Relat�rio de empresas por s�cio", "", 1, i, , , , "mnuEmpresaSocio")
    h = .AddItem("Lista de Empresas com ISS Estimado", "", 1, i, , , , "mnuRelEstimado")
    h = .AddItem("Lista de Empresas incluidas no MEI", "", 1, i, , , , "mnuRelMEI")
    h = .AddItem("Iss pago por atividade", "", 1, i, , , , "mnuIssPagoAtividade")
    h = .AddItem("Resumo iss constru��o civil", "", 1, i, , , , "mnuResumoIssCCivil")
    h = .AddItem("Notas fiscais emitidas na CONSIST por per�odo", "", 1, i, , , False, "mnuNFEmitida")
    h = .AddItem("Guias emitidas pela Giss por per�odo", "", 1, i, , , , "mnuRelGiss")
End With

Set m_cMenuAtende = New cPopupMenu
With m_cMenuAtende
    .hwndOwner = Me.HWND
    .GradientHighlight = True
    .ActiveMenuBackgroundColor = &HFFFF80
    .ButtonHighlight = True
    .InActiveMenuForeColor = &H400000
    .MenuBackgroundColor = &H8000000F
    i = .AddItem("Senhas SPAC", "", 1, , , , , "mnuSenha")
    h = .AddItem("Controle de senhas", "", 1, i, , , , "mnuSenhaControle")
    h = .AddItem("Resumo das senhas emitidas", "", 1, i, , , , "mnuSenhaResumo")
    i = .AddItem("Cadastro de cidad�o", "", 1, , , , , "mnuCidadao")
    i = .AddItem("Emiss�o de guias", "", 1, , , , , "mnu2ViaLaser")
   ' i = .AddItem("Emiss�o de guias (Nova Vers�o)", "", 1, , , , , "mnuEmissaoGuia")
    i = .AddItem("Movimento econ�mico", "", 1, , , , , "mnuMovimento")
    i = .AddItem("Autoriza��o de tal�o de nota fiscal", "", 1, , , , , "mnuAutorizaNF")
    i = .AddItem("Emiss�o de ITBI", "", 1, , , , , "mnuITBI")
    i = .AddItem("Declara��o cadastral (DECA)", "", 1, , , , , "mnuDeca")
    i = .AddItem("Requerimento p/abertura de processo", "", 1, , , , , "mnuRequerimentoProc")
    i = .AddItem("Dep�sito CRI", "", 1, , , , , "mnuDepositoCRI")
    i = .AddItem("Recolhimento aos cofres municipais", "", 1, , , , , "mnuGuiaPratico1")
    i = .AddItem("Lan�amento de IPTU complementar", "", 1, , , , , "mnuGuiaPratico2")
    i = .AddItem("Lan�amento de IPTU proporcional", "", 1, , , , , "mnuGuiaPratico5")
    i = .AddItem("Requerimento p/autoriza��o especial de estacionamento", "", 1, , , , , "mnuGuiaPratico3")
    i = .AddItem("Declara��o de Isen��o de IPTU", "", 1, , , , , "mnuDeclaraIsentoIPTU")
    i = .AddItem("Requerimento p/ Isen��o de IPTU", "", 1, , , , , "mnuRequerIsentoIPTU")
    i = .AddItem("Parcelamento de divida", "", 1, , , , , "mnuParcelamentoDivida")
    h = .AddItem("Cancelamento de parcelamento (antigo)", "", 1, i, , , , "mnuCancelReparc")
    'h = .AddItem("Parcelamento de divida fiscal", "", 1, i, , , , "mnuParcelaDebito")
    h = .AddItem("Parcelamento de divida fiscal", "", 1, i, , , , "mnuParcelamento")
    h = .AddItem("Libera��o de carn� de parcelamento", "", 1, i, , , , "mnuLiberaCarne")
    h = .AddItem("Cancelamento manual de parcelamento", "", 1, i, , , , "mnuCancelParcelamento")
    h = .AddItem("Cancelamento autom�tico de parcelamento", "", 1, i, , , , "mnuCancelParcelamentoAuto")
    h = .AddItem("Mala direta p/parcelamento bloqueado", "", 1, i, , , , "mnuMalaDiretaParc")
    h = .AddItem("Desbloquear parcelamentos", "", 1, i, , , , "mnuDesbloquearParc")
    h = .AddItem("Pagamento mensal de parcelamentos", "", 1, i, , , , "mnuPagamentoMensalParc")
    i = .AddItem("Relat�rios", "", 1, , , , , "mnuRelatorioAte")
    h = .AddItem("Emiss�o de 2� Via", "", 1, i, , , , "mnu2Via")
    'h = .AddItem("Emiss�o de 2� Via Especial", "", 1, i, , , , "mnu2ViaEspecial")
    h = .AddItem("Termo de confiss�o de divida", "", 1, i, , , , "mnuTermConf")
    h = .AddItem("Comunicado de cobran�a judicial", "", 1, i, , , , "mnuCobrancaJudicial")
    'h = .AddItem("Renova��o de Alvar�", "", 1, i, , , , "mnuRenovaAlvara")
    h = .AddItem("Documentos emitidos por usu�rio", "", 1, i, , , , "mnuEmiteDoc")
    h = .AddItem("Relat�rio do REFIS DAM", "", 1, i, , , , "mnuRelRefis")
    'h = .AddItem("Relat�rio do REFIS DAM por tributo", "", 1, i, , , False, "mnuRelRefisTributo")
    h = .AddItem("Relat�rio do REFIS Parcelado", "", 1, i, , , , "mnuRelRefisParc")
  '  h = .AddItem("Requerimento de senha para ISS Eletr�nico", "", 1, i, , , , "mnuSenhaISS")
    h = .AddItem("Mala Direta Cidad�o", "", 1, i, , , , "mnuMalaDiretaCidadao")
    i = .AddItem("Consultas", "", 1, , , , , "mnuConsultaAte")
    h = .AddItem("Consulta/Reativa��o de documentos", "", 1, i, , , , "mnuCnsNumDoc")
    h = .AddItem("Extrato do contribuinte", "", 1, i, , , , "CnsDebitoImob")
End With

Set m_cMenuTrib = New cPopupMenu
With m_cMenuTrib
    .hwndOwner = Me.HWND
    .GradientHighlight = True
    .ActiveMenuBackgroundColor = &HFFFF80
    .ButtonHighlight = True
    .InActiveMenuForeColor = &H400000
    .MenuBackgroundColor = &H8000000F
    i = .AddItem("C�lculo", "", 1, , , , , "mnuCalcTit")
    h = .AddItem("C�lculo de IPTU", "", 1, i, , , , "mnuCalcGeral")
    h = .AddItem("C�lculo de IPTU (Novo)", "", 1, i, , , , "mnuCalcIPTU")
    h = .AddItem("C�lculo de ISS", "", 1, i, , , , "mnuCalcGeralISS")
    h = .AddItem("C�lculo CIP", "", 1, i, , , , "mnuCalculoCIP")
    h = .AddItem("Amostra de C�lculo de IPTU", "", 1, i, , , , "mnuAmostraCalculo")
    h = .AddItem("Gera��o de arquivo laser", "", 1, i, , , , "mnuArqLaser")
    i = .AddItem("Atividades banc�rias", "", 1, , , , , "mnuAtivBanco")
    h = .AddItem("Optantes por d�bito autom�tico", "", 1, i, , , , "mnuOptanteDA")
    h = .AddItem("Gera��o de arquivo de cobran�a BB", "", 1, i, , , , "mnuCobranca")
    h = .AddItem("Gera��o de arquivos para registro banc�rio", "", 1, i, , , , "mnuGeraLote")
    h = .AddItem("Relat�rio de Optantes", "", 1, i, , , , "mnuOptanteDARel")
    h = .AddItem("Importa��o de arquivos banc�rios", "", 1, i, , , , "mnuImportaArq")
    h = .AddItem("Busca em arquivos banc�rios", "", 1, i, , , , "mnuBuscaArq")
    h = .AddItem("Baixa de D�bitos", "", 1, i, , , , "mnuBaixaDebito")
    h = .AddItem("An�lise da Receita", "", 1, i, , , , "mnuAnaliseReceita")
    h = .AddItem("Nova importa��o", "", 1, i, , , , "mnuNovaImportacao")
   ' h = .AddItem("An�lise da Receita2", "", 1, i, , , , "mnuAnaliseReceita2")
    h = .AddItem("Relat�rio de pagamento", "", 1, i, , , , "mnuRelPagamento")
    i = .AddItem("Simples Nacional", "", 1, , , , , "mnuSimples")
    h = .AddItem("Arrecada��o do simples nacional", "", 1, i, , , , "mnuArrecadaSN")
    h = .AddItem("Lan�amentos criados pelo simples nacional", "", 1, i, , , , "mnuArrecadaSNDAF")
    h = .AddItem("Efetuar baixa por CNPJ", "", 1, i, , , , "mnuBaixaSN")
    h = .AddItem("Tabela Selic", "", 1, i, , , , "mnuTabelaSelic")
    h = .AddItem("Simula��o de c�lculo", "", 1, i, , , , "mnuSimulaSN")
    h = .AddItem("Importa��o de Per�odos", "", 1, i, , , , "mnuImportarSN")
    h = .AddItem("Decodificar arquivos do MEI", "", 1, i, , , , "mnuDecodificaMEI")
    h = .AddItem("Resumo dos arquivos banc�rios", "", 1, i, , , , "mnuBuscaSN")
    h = .AddItem("CNPJ n�o cadastrados", "", 1, i, , , , "mnuSNCnpj")
    h = .AddItem("Exportar CNPJ para Rec.Federal", "", 1, i, , , , "mnuSNCnpjReceita")
    i = .AddItem("Cobran�a de aluguel", "", 1, , , , , "mnuAlugueis")
    h = .AddItem("Manuten��o dos alugu�is", "", 1, i, , , , "mnuManAluguel")
    h = .AddItem("Emiss�o dos boletos de cobran�a", "", 1, i, , , , "mnuImpAluguel")
    i = .AddItem("Divida Ativa", "", 1, , , , , "mnuDividaAtivaT")
    h = .AddItem("Encerramento dos livros", "", 1, i, , , , "mnuDividaAtiva")
    h = .AddItem("Emiss�o dos livros DA", "", 1, i, , , , "mnuEmiteLivro")
    h = .AddItem("Ajuizamento autom�tico", "", 1, i, , , , "mnuAjuizaAuto")
    h = .AddItem("Resumo do pagamento das cartas de cobran�a", "", 1, i, , , , "mnuPagamentoCC")
    i = .AddItem("Outros", "", 1, , , , , "mnuOutrosT")
    h = .AddItem("Autoriza complementos", "", 1, i, , , , "mnuDebitoCompl")
    h = .AddItem("Emiss�o de Guias p/Ambulantes", "", 1, i, , , , "mnuGuiaAmbulante")
    h = .AddItem("Situa��o dos tributos lan�ados", "", 1, i, , , , "mnuSituacaoTributo")
    h = .AddItem("Situa��o tribut�ria de Contribuinte", "", 1, i, , , , "mnuSituacaoTributaria")
    h = .AddItem("Corre��o de CPF", "", 1, i, , , , "mnuCorrecaoCPF")
    h = .AddItem("Total pago por tributo", "", 1, i, , , , "mnuPagoTributo")
    i = .AddItem("Relat�rios", "", 1, , , , , "mnuRelatorioTrib")
    h = .AddItem("Carta cobran�a amig�vel", "", 1, i, , , , "mnuCartaCobranca")
    h = .AddItem("Relat�rio de ajuizamentos", "", 1, i, , , , "mnuRelAjuizamento")
    h = .AddItem("Lista de devedores", "", 1, i, , , , "mnuListaDevedor")
    h = .AddItem("Notifica��o de imposto devido", "", 1, i, , , , "mnuNotificacao")
'    h = .AddItem("Notifica��o de iss constru��o civil", "", 1, i, , , , "mnuNotificacao2")
    h = .AddItem("Aviso de d�bito", "", 1, i, , , , "mnuAvisoDebito")
    h = .AddItem("D�bito ajuizados pagos", "", 1, i, , , , "mnuDebitoAjPago")
    h = .AddItem("Complementos gerados", "", 1, i, , , , "mnuComplementoPagto")
    h = .AddItem("ITBI emitidos e n�o pagos", "", 1, i, , , , "mnuITBIObs")
    h = .AddItem("Rela��o de ITBI's emitidos", "", 1, i, , , , "mnuITBIRel")
    h = .AddItem("Documentos emitidos", "", 1, i, , , , "mnuDocEmitido")
   ' h = .AddItem("2� via de Notifica��o de ISS", "", 1, i, , , , "mnu2vianotificacao")
    h = .AddItem("Pagamento em Duplicidade", "", 1, i, , , , "mnuPagamentoDuplicidade")
    h = .AddItem("Resumo ISS Constru��o Civil", "", 1, i, , , , "mnuResumoISSCivil")
    
End With

Set m_cMenuProt = New cPopupMenu
With m_cMenuProt
    .hwndOwner = Me.HWND
    .GradientHighlight = True
    .ActiveMenuBackgroundColor = &HFFFF80
    .ButtonHighlight = True
    .InActiveMenuForeColor = &H400000
    .MenuBackgroundColor = &H8000000F
    i = .AddItem("Par�metros do protocolo", "", 1, , , , , "mnuParametroProt")
    h = .AddItem("Tabela de despachos", "", 1, i, , , , "mnuDespacho")
    h = .AddItem("Centro de custos", "", 1, i, , , , "mnuCentroCusto")
    h = .AddItem("Tabela de documentos", "", 1, i, , , , "mnuTabDocumento")
    h = .AddItem("Tabela de assuntos", "", 1, i, , , , "mnuAssunto")
    h = .AddItem("Tabela de funcion�rios", "", 1, i, , , , "mnuUsuarioProt")
    h = .AddItem("Permiss�es de recebimento", "", 1, i, , , , "mnuPermissaoProt")
    i = .AddItem("Processo", "", 1, , , , , "mnuProcesso")
    i = .AddItem("Processos Arquivados", "", 1, , , , , "mnuProcessoArquivado")
    i = .AddItem("Emiss�o de etiquetas", "", 1, , , , , "mnuEtiquetaProt")
    i = .AddItem("Resumo di�rio  dos processos", "", 1, , , , , "mnuResumoDiario")
    i = .AddItem("Publica��o dos processos", "", 1, , , , , "mnuPublicacao")
    i = .AddItem("Processos com tr�mite em aberto por c.custo", "", 1, , , , , "mnuTramiteAberto")
    i = .AddItem("Processos com tr�mite em atraso", "", 1, , , , , "mnuTramiteAtraso")
    i = .AddItem("Processos enviados por c.custo", "", 1, , , , , "mnuTramiteEnviado")
    i = .AddItem("Processos por Assunto", "", 1, , , , , "mnuProcessoAssunto")
    i = .AddItem("Processos que est�o em um C.Custo", "", 1, , , , , "mnuProcessoCC")
    i = .AddItem("Qtde de Processos por Ano", "", 1, , , , , "mnuProcessoAno")
    i = .AddItem("Lista de assuntos por documento", "", 1, , , , , "mnuAssuntoDoc")
End With

Set m_cMenuOutro = New cPopupMenu
With m_cMenuOutro
    .hwndOwner = Me.HWND
    .GradientHighlight = True
    .ActiveMenuBackgroundColor = &HFFFF80
    .ButtonHighlight = True
    .InActiveMenuForeColor = &H400000
    .MenuBackgroundColor = &H8000000F
    i = .AddItem("Administrador", "", 1, , , , , "mnuAdmin")
    h = .AddItem("Seguran�a", "", 1, i, , , , "mnuSeguranca")
    z = .AddItem("Cadastro de usu�rios", "", 1, h, , , , "mnuUser")
    z = .AddItem("Seguran�a por evento", "", 1, h, , , , "mnuSegEvento")
    z = .AddItem("Seguran�a por usu�rio", "", 1, h, , , , "mnuAtribSeg")
    z = .AddItem("Alterar usu�rio", "", 1, h, , , , "mnuChangeLogin")
    z = .AddItem("Security by User", "", 1, h, , , , "mnuSecurityByUser")
    
    h = .AddItem("Sql Builder", "", 1, i, , , , "mnuSql")
'    h = .AddItem("Grava Foto", "", 1, i, , , , "mnuGravaFoto")
    h = .AddItem("Gera��o manual de d�bitos", "", 1, i, , , , "mnuGeraDebito")
    h = .AddItem("Baixa ISS-DAM", "", 1, i, , , , "mnuBaixaIss")
   ' i = .AddItem("Integra��o com ISS Eletr�nico", "", 1, , , , , "mnuExporta")
    i = .AddItem("Integra��o com Sistema de Cobran�a", "", 1, , , , , "mnuIntegrativa")
    i = .AddItem("Integra��o Eicon", "", 1, , , , , "mnuEicon")
    i = .AddItem("Secretaria de Obras", "", 1, , , , , "mnuSecretariaObra")
    h = .AddItem("Par�metros", "", 1, i, , , , "mnuParamObra")
    p = .AddItem("Relat�rios", "", 1, i, , , , "mnuRelatorioAtendimento")
    
    z = .AddItem("Tabela de Equipes", "", 1, h, , , , "mnuTabelaEquipe")
    z = .AddItem("Tabela de Funcion�rios", "", 1, h, , , , "mnuTabelaFuncionario")
    z = .AddItem("Tabela de Assuntos", "", 1, h, , , , "mnutabelaAssunto")
    z = .AddItem("Tabela de Atendentes", "", 1, h, , , , "mnutabelaAtendente")
    z = .AddItem("Tabela de Tipo de Atendimento", "", 1, h, , , , "mnutabelaTipoAtendimento")
    h = .AddItem("Registro de Atendimento", "", 1, i, , , , "mnuRegistroAtendimento")
    h = .AddItem("Reincid�ncia de OS por endere�o", "", 1, i, , , , "mnuRegEndereco")
    i = .AddItem("Controle de anexos", "", 1, , , , , "mnuAnexos")
End With

End Sub

Private Sub MalaDiretaRo�ada(dData As Date)
Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, Sql As String
Dim xId As Long, nNumRec As Long, nCodLogr As Long, sCodInscricao As String, sContribuinte As String
Dim sEnd As String, nNum As Integer, sCep As String, sCompl As String, sBairro As String
Dim sEndEntrega As String, sBairroEntrega As String, sCidEntrega As String, sCepEntrega As String, sUFEntrega As String, sNumEntrega As String
Dim z As Variant

Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

Sql = "SELECT CODREDUZIDO FROM ETIQUETAROCADA WHERE DATA='" & Format(dData, "mm/dd/yyyy") & "'"
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
Do Until RdoAux2.EOF
    'If !CODREDUZIDO = 21656 Then MsgBox "teste"
    Sql = "SELECT * FROM VWFULLIMOVEL2 WHERE CODREDUZIDO=" & RdoAux2!CODREDUZIDO & " ORDER BY CODREDUZIDO"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
'       nNumRec = .RowCount
'        Do Until .EOF
            sContribuinte = !nomecidadao
            
            If !Ee_TipoEnd = 0 Then 'ENDERECO DO IMOVEL
                sEnd = !Logradouro & " N� " & CStr(!Li_Num)
                sCep = RetornaCEP(!CodLogr, !Li_Num)
                sCompl = SubNull(Left(!Li_Compl, 20))
                sBairro = SubNull(!DescBairro)
                sCidEntrega = "JABOTICABAL"
                sUFEntrega = "SP"
            ElseIf !Ee_TipoEnd = 1 Then
                Sql = "select * from vwfullcidadao where codcidadao=" & !CodCidadao
                Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux3
                    sEndEntrega = SubNull(!Endereco) & ", " & Val(SubNull(!NUMIMOVEL))
                    sBairroEntrega = SubNull(!DescBairro)
                    sCidEntrega = SubNull(!descCidade)
                    sCepEntrega = SubNull(!Cep)
                    sComplEntrega = SubNull(!Complemento)
                    sUFEntrega = SubNull(!SiglaUF)
                   .Close
                   GoTo IMPRIMIR
                End With
            ElseIf !Ee_TipoEnd = 2 Then 'ENDERECO DE ENTREGA
                If IsNull(!AbrevTipoLogEE) Then
                    sEnd = SubNull(!Ee_NomeLog)
                Else
                    sEnd = Trim(!AbrevTipoLogEE) & " " & Trim(SubNull(!AbrevTitLogEE)) & " " & !Ee_NomeLog
                End If
                sEnd = sEnd & " N� " & CStr(!Ee_NumImovel)
                sCep = SubNull(!Ee_Cep)
                sCompl = Left(SubNull(!Ee_Complemento), 20)
                sBairro = SubNull(!BairroEE)
                sCidEntrega = SubNull(!CidadeEE)
                sUFEntrega = SubNull(!Ee_Uf)
            End If
            
            sEndEntrega = sEnd
            sBairroEntrega = sBairro
'                sCidEntrega = "JABOTICABAL"
            sCepEntrega = sCep
            sComplEntrega = sCompl
'               sUFEntrega = "SP"
            
IMPRIMIR:
            Sql = "INSERT ETIQUETAGTI (USUARIO,SEQ,CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5) VALUES('"
            Sql = Sql & NomeDeLogin & "'," & xId & ",'" & Format(!CODREDUZIDO, "000000") & "','" & Mask(sContribuinte) & "','"
            Sql = Sql & Left(sEndEntrega & " " & sComplEntrega, 60) & "','" & sCepEntrega & "   " & sBairroEntrega & "','" & Mask(sCidEntrega) & "   " & sUFEntrega & "')"
            cn.Execute Sql, rdExecDirect
            xId = xId + 1
proximo3:
    End With
    RdoAux2.MoveNext
    DoEvents
Loop
End Sub

Private Sub MalaDiretaISSCCivil(dData1 As Date, dData2 As Date)
Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, Sql As String
Dim xId As Long, nNumRec As Long, nCodLogr As Long, sCodInscricao As String, sContribuinte As String
Dim sEnd As String, nNum As Integer, sCep As String, sCompl As String, sBairro As String
Dim sEndEntrega As String, sBairroEntrega As String, sCidEntrega As String, sCepEntrega As String, sUFEntrega As String, sNumEntrega As String
Dim z As Variant

Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

Sql = "SELECT DISTINCT codigo_imovel,numero_notificacao,ano_notificacao,processo FROM notificacao_iss_ccivil WHERE data_gravacao BETWEEN'" & Format(dData1, "mm/dd/yyyy") & "' AND '" & Format(dData2, "mm/dd/yyyy") & "'"
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
Do Until RdoAux2.EOF
    'If !CODREDUZIDO = 21656 Then MsgBox "teste"
    Sql = "SELECT * FROM VWFULLIMOVEL2 WHERE CODREDUZIDO=" & RdoAux2!codigo_imovel & " ORDER BY CODREDUZIDO"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
'       nNumRec = .RowCount
'        Do Until .EOF
            sContribuinte = !nomecidadao
            
            If !Ee_TipoEnd = 0 Then 'ENDERECO DO IMOVEL
                sEnd = !Logradouro & " N� " & CStr(!Li_Num)
                sCep = RetornaCEP(!CodLogr, !Li_Num)
                sCompl = SubNull(Left(!Li_Compl, 20))
                sBairro = SubNull(!DescBairro)
                sCidEntrega = "JABOTICABAL"
                sUFEntrega = "SP"
            ElseIf !Ee_TipoEnd = 1 Then
                Sql = "select * from vwfullcidadao where codcidadao=" & !CodCidadao
                Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux3
                    sEndEntrega = SubNull(!Endereco) & ", " & Val(SubNull(!NUMIMOVEL))
                    sBairroEntrega = SubNull(!DescBairro)
                    sCidEntrega = SubNull(!descCidade)
                    sCepEntrega = SubNull(!Cep)
                    sComplEntrega = SubNull(!Complemento)
                    sUFEntrega = SubNull(!SiglaUF)
                   .Close
                   GoTo IMPRIMIR
                End With
            ElseIf !Ee_TipoEnd = 2 Then 'ENDERECO DE ENTREGA
                If IsNull(!AbrevTipoLogEE) Then
                    sEnd = SubNull(!Ee_NomeLog)
                Else
                    sEnd = Trim(!AbrevTipoLogEE) & " " & Trim(SubNull(!AbrevTitLogEE)) & " " & !Ee_NomeLog
                End If
                sEnd = sEnd & " N� " & CStr(!Ee_NumImovel)
                sCep = SubNull(!Ee_Cep)
                sCompl = Left(SubNull(!Ee_Complemento), 20)
                sBairro = SubNull(!BairroEE)
                sCidEntrega = SubNull(!CidadeEE)
                sUFEntrega = SubNull(!Ee_Uf)
            End If
            
            sEndEntrega = sEnd
            sBairroEntrega = sBairro
'                sCidEntrega = "JABOTICABAL"
            sCepEntrega = sCep
            sComplEntrega = sCompl
'               sUFEntrega = "SP"
            
IMPRIMIR:
            Sql = "INSERT ETIQUETAGTI (USUARIO,SEQ,CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5) VALUES('"
            Sql = Sql & NomeDeLogin & "'," & xId & ",'" & Format(!CODREDUZIDO, "000000") & " Proc: " & RdoAux2!Processo & " Not: " & RdoAux2!numero_notificacao & "/" & RdoAux2!ano_notificacao & "','" & Mask(sContribuinte) & "','"
            Sql = Sql & Left(sEndEntrega & " " & sComplEntrega, 60) & "','" & sCepEntrega & "   " & sBairroEntrega & "','" & Mask(sCidEntrega) & "   " & sUFEntrega & "')"
            cn.Execute Sql, rdExecDirect
            xId = xId + 1
proximo3:
    End With
    RdoAux2.MoveNext
    DoEvents
Loop
End Sub

Private Sub BuildRelITBI()
Dim Sql As String, RdoAux As rdoResultset, ax As String, sNomeArq As String
Dim z As Variant, z2 As Variant

Data1:
    z = InputBox("Digite a data inicial.", "Entre com a informa��o")
    If z = "" Then GoTo Data1
    If Not IsDate(z) Then GoTo Data1
Data2:
    z2 = InputBox("Digite a data final.", "Entre com a informa��o")
    If z2 = "" Then GoTo Data2
    If Not IsDate(z2) Then GoTo Data2


sNomeArq = sPathBin & "\ITBIREL.TXT"
FF1 = FreeFile()
Open sNomeArq For Output As FF1

Print #FF1, "**************************************************************"
Print #FF1, "ITBI'S EMITIDOS ENTRE " & z & " E " & z2
Print #FF1, "IMPRESSO EM " & Format(Now, "dd/mm/yyyy") & " - Fonte: GTI"
Print #FF1, "**************************************************************"
ax = ""
Print #FF1, ax
ax = FillSpace("C�DIGO", 8) & FillSpace("RAZ�O SOCIAL", 42) & "SIT  " & FillSpace("VENCTO", 11) & FillSpace("PROCESSO", 11) & FillLeft("VALOR", 11) & " OBS"
Print #FF1, ax
Print #FF1, "******************************************************************************************************************************************************************************************************"

Sql = "SELECT debitoparcela.codreduzido, debitoparcela.statuslanc, debitoparcela.datavencimento, debitoparcela.numprocesso, vwFULLCIDADAO.nomecidadao, "
Sql = Sql & "debitotributo.CodTributo,debitotributo.ValorTributo, obsparcela.obs FROM debitoparcela INNER JOIN "
Sql = Sql & "vwFULLCIDADAO ON debitoparcela.codreduzido = vwFULLCIDADAO.codcidadao INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
Sql = Sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND debitoparcela.numparcela = debitotributo.numparcela AND debitoparcela.codcomplemento = debitotributo.codcomplemento INNER JOIN "
Sql = Sql & "obsparcela ON debitoparcela.codreduzido = obsparcela.codreduzido AND debitoparcela.anoexercicio = obsparcela.anoexercicio AND debitoparcela.codlancamento = obsparcela.codlancamento AND debitoparcela.seqlancamento = obsparcela.seqlancamento AND "
Sql = Sql & "debitoparcela.NumParcela = obsparcela.NumParcela And debitoparcela.CODCOMPLEMENTO = obsparcela.CODCOMPLEMENTO "
Sql = Sql & "Where (debitoparcela.CodLancamento = 36) And (debitoparcela.statuslanc <> 5) And (debitotributo.CodTributo = 84) and debitoparcela.datavencimento between '" & Format(z, "mm/dd/yyyy") & "' and '" & Format(z2, "mm/dd/yyyy") & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ax = FillSpace(!CODREDUZIDO, 8) & FillSpace(!nomecidadao, 42) & " " & !statuslanc & " " & Format(!DataVencimento, "dd/mm/yyyy") & " " & !NUMPROCESSO & " " & FillLeft(Format(!ValorTributo, "#0.00"), 11) & " " & SubNull(!obs)
        Print #FF1, ax
       .MoveNext
    Loop
   .Close
End With

Close #FF1
ret = Shell("NOTEPAD" & " " & sNomeArq, vbNormalFocus)

End Sub

Private Function FillLeft(sTexto As String, nTamanho As Integer) As String

FillLeft = Space(nTamanho - Len(sTexto)) & sTexto

End Function

Private Function FillSpace(sPalavra As String, nTamanho As Integer) As String
Dim sTmp As String

If Len(sPalavra) > nTamanho Then sPalavra = Left(sPalavra, nTamanho)
sTmp = sPalavra & Space(nTamanho - Len(sPalavra))
FillSpace = sTmp

End Function

Private Sub txtDV_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    cmdDV_Click
End If
End Sub

Private Sub DevedoresTaxaLicenca()
Dim Sql As String, RdoAux As rdoResultset, nCodReduz As Long, nAno As Integer, sNome As String, bFind As Boolean

Sql = "SELECT debitoparcela.codreduzido, mobiliario.razaosocial, debitoparcela.anoexercicio FROM debitoparcela INNER JOIN mobiliario ON debitoparcela.codreduzido = mobiliario.codigomob "
Sql = Sql & "WHERE (debitoparcela.codreduzido BETWEEN 100000 AND 200000) AND (debitoparcela.codlancamento = 6) AND (debitoparcela.numparcela = 1) AND (debitoparcela.statuslanc = 3) AND (mobiliario.dataencerramento IS NULL) "
Sql = Sql & "ORDER BY debitoparcela.codreduzido, debitoparcela.anoexercicio"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        nCodReduz = !CODREDUZIDO
        sNome = !RazaoSocial
        nAno = !AnoExercicio
        
        
        
       .MoveNext
    Loop
   .Close
End With


End Sub

Private Sub BaixaEicon()
Dim nCodReduz As Long, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, nAno As Integer, nLc As Integer, nSq As Integer, RdoAux3 As rdoResultset
Dim nPc As Integer, nCp As Integer, sCnae As String, nPos As Long, nTot As Long, nIni As Integer, nFim As Integer, sMotivo As String, nNumDoc As Long, RdoAux4 As rdoResultset
If NomeDeLogin <> "SCHWARTZ" Then Exit Sub

nPos = 1
Sql = "SELECT * From debitopago WHERE (codreduzido between 100000 and 200000) and (codlancamento = 5) AND (numdocumento > 4000000) AND (codbanco NOT IN (90, 91, 92, 93, 94, 95, 96, 97, 98, 99)) AND "
Sql = Sql & " (seqpag = 0) AND (codcomplemento = 0) and year(datapagamento)>2017 order by numdocumento"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    Do Until .EOF

        nNumDoc = !NumDocumento
        Sql = "select * from parceladocumento where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and codlancamento=" & !CodLancamento & " and "
        Sql = Sql & "seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO & " and numdocumento between 2000000 and 3000000"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                Sql = "select * from parceladocumento where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and codlancamento=" & !CodLancamento & " and "
                Sql = Sql & "seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO & " and numdocumento <> " & RdoAux!NumDocumento
                Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                Do Until RdoAux3.EOF
                    Sql = "select * from damiss where docdam=" & nNumDoc & " and dociss=" & RdoAux3!NumDocumento
                    Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    If RdoAux4.RowCount = 0 Then
                        Sql = "insert damiss (docdam,dociss,baixado) values(" & nNumDoc & "," & RdoAux3!NumDocumento & ",0)"
                        cn.Execute Sql, rdExecDirect
                    End If
                    RdoAux4.Close
                    RdoAux3.MoveNext
                Loop
                RdoAux3.Close
            End If
           .Close
        End With
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With


ConectaEicon

Sql = "SELECT DISTINCT docdam,dociss From damiss where baixado=0 ORDER BY docdam"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        nNumDoc = !docdam
        nNumDocISS = !dociss
        Sql = "select * from debitopago where numdocumento=" & nNumDoc & " and codlancamento=5"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
    
            Sql = "SELECT parceladocumento.codreduzido, debitotributo.valortributo FROM parceladocumento INNER JOIN debitotributo ON parceladocumento.codreduzido = debitotributo.codreduzido AND parceladocumento.anoexercicio = debitotributo.anoexercicio AND parceladocumento.codlancamento = debitotributo.codlancamento AND "
            Sql = Sql & "parceladocumento.SeqLancamento = debitotributo.SeqLancamento And parceladocumento.NumParcela = debitotributo.NumParcela And parceladocumento.CODCOMPLEMENTO = debitotributo.CODCOMPLEMENTO Where parceladocumento.NumDocumento = " & nNumDoc
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            nValorDoc = RdoAux3!ValorTributo
            RdoAux3.Close

            '***** GRAVA BAIXA NA GISS ***************
            Sql = "insert tb_inter_baixa(cod_cliente,cod_banco,num_sequencia,timestamp,data_geracao,nome_arquivo,data_movimento) values("
            Sql = Sql & 2177 & "," & RdoAux2!CodBanco & "," & 0 & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "','" & Format(Now, "mm/dd/yyyy") & "','"
            Sql = Sql & RdoAux2!arquivobanco & "','" & Format(RdoAux2!datarecebimento, "mm/dd/yyyy") & "')"
            cnEicon.Execute Sql, rdExecDirect
            
            Sql = "insert tb_inter_baixa_detalhe(cod_cliente,cod_banco,num_sequencia,num_documento,linha,timestamp,valor_titulo,valor_pago,data_pagamento,valor_encargos,"
            Sql = Sql & "descricao_linha_t,descricao_linha_u) values(" & 2177 & "," & RdoAux2!CodBanco & "," & 0 & "," & nNumDocISS & "," & RdoAux2!SEQPAG & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "',"
            Sql = Sql & Virg2Ponto(CStr(nValorDoc)) & "," & Virg2Ponto(CStr(nValorDoc)) & ",'" & Format(RdoAux2!DataPagamento, "mm/dd/yyyy") & "'," & 0 & ",'"
            Sql = Sql & "" & "','" & "" & "')"
            cnEicon.Execute Sql, rdExecDirect
            
            Sql = "update damiss set baixado=1 where dociss=" & nNumDocISS
            cn.Execute Sql, rdExecDirect
        End If
       .MoveNext
    Loop
   .Close
End With

Sql = "SELECT codreduzido, anoexercicio, codlancamento, seqlancamento, numparcela, codcomplemento, seqpag, datapagamento, datarecebimento, valorpago,CodBanco,"
Sql = Sql & "CodAgencia, restituido, NumDocumento, valorpagoreal, intacto, ValorTarifa, arquivobanco, valordif, datapagamentocalc, dataintegracao, contacorrente "
Sql = Sql & "From debitopago WHERE (numdocumento BETWEEN 2000000 AND 2200000) AND (numdocumento NOT IN (SELECT num_documento FROM GTI_Eicon.dbo.tb_inter_baixa_detalhe)) "
Sql = Sql & " AND (anoexercicio > 2015) ORDER BY numdocumento"

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        DoEvents
        '***** GRAVA BAIXA NA GISS ***************
        Sql = "insert tb_inter_baixa(cod_cliente,cod_banco,num_sequencia,timestamp,data_geracao,nome_arquivo,data_movimento) values("
        Sql = Sql & 2177 & "," & !CodBanco & "," & 0 & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "','" & Format(Now, "mm/dd/yyyy") & "','"
        Sql = Sql & !arquivobanco & "','" & Format(!datarecebimento, "mm/dd/yyyy") & "')"
        cnEicon.Execute Sql, rdExecDirect
        
        Sql = "insert tb_inter_baixa_detalhe(cod_cliente,cod_banco,num_sequencia,num_documento,linha,timestamp,valor_titulo,valor_pago,data_pagamento,valor_encargos,"
        Sql = Sql & "descricao_linha_t,descricao_linha_u) values(" & 2177 & "," & !CodBanco & "," & 0 & "," & !NumDocumento & "," & !SEQPAG & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "',"
        Sql = Sql & Virg2Ponto(CStr(!valorpagoreal)) & "," & Virg2Ponto(CStr(!valorpagoreal)) & ",'" & Format(!DataPagamento, "mm/dd/yyyy") & "'," & 0 & ",'"
        Sql = Sql & "" & "','" & "" & "')"
        cnEicon.Execute Sql, rdExecDirect
    
       .MoveNext
    Loop
   .Close
End With

cnEicon.Close

MsgBox "Baixa finalizada", vbInformation, "Infoma��o"

End Sub

Private Sub Gera_WhatsNew()

Dim Sql As String, RdoAux As rdoResultset, ax As String, sNomeArq As String, sVersion As String, aVersion() As tVersion, x As Integer

ReDim aVersion(0)
Sql = "select distinct major,minor,revision from whats_new order by major desc,minor desc,revision desc"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        sVersion = !Major & "." & !Minor & "." & !Revision
        ReDim Preserve aVersion(UBound(aVersion) + 1)
        aVersion(UBound(aVersion)).version = sVersion
        aVersion(UBound(aVersion)).Major = !Major
        aVersion(UBound(aVersion)).Minor = !Minor
        aVersion(UBound(aVersion)).Revision = !Revision
       .MoveNext
    Loop
   .Close
End With

sNomeArq = sPathBin & "\WHATSNEW.TXT"
FF1 = FreeFile()
Open sNomeArq For Output As FF1

Print #FF1, "**************************************************************"
Print #FF1, "G.T.I. - Hist�rico de Atualiza��es"
Print #FF1, "Impresso em " & Format(Now, "dd/mm/yyyy") & " �S " & Format(Now, "hh:mm")
Print #FF1, "**************************************************************"

For x = 1 To UBound(aVersion)
    Sql = "SELECT * from whats_new where major=" & aVersion(x).Major & " and minor=" & aVersion(x).Minor & " and revision=" & aVersion(x).Revision & " order by data desc,seq"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            If .AbsolutePosition = 1 Then
                ax = ""
                Print #FF1, ax
                ax = "Vers�o: " & aVersion(x).version
                Print #FF1, ax
            End If
            ax = "* " & Format(!Data, "dd/mm/yyyy") & " - " & !obs
            Print #FF1, ax
           .MoveNext
        Loop
       .Close
    End With
Next


Close #FF1
ret = Shell("NOTEPAD" & " " & sNomeArq, vbNormalFocus)

End Sub
