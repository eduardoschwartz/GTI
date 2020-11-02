VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form frmLogradouro 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Logradouros"
   ClientHeight    =   7125
   ClientLeft      =   4620
   ClientTop       =   2055
   ClientWidth     =   7440
   Icon            =   "frmLogradouro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPesq 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   930
      TabIndex        =   0
      Top             =   60
      Width           =   6465
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   5760
      TabIndex        =   40
      ToolTipText     =   "Sair da Tela"
      Top             =   6720
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
      MICON           =   "frmLogradouro.frx":014A
      PICN            =   "frmLogradouro.frx":0166
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
      Left            =   5760
      TabIndex        =   34
      ToolTipText     =   "Cancelar Edição"
      Top             =   6720
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
      MICON           =   "frmLogradouro.frx":01D4
      PICN            =   "frmLogradouro.frx":01F0
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
      Left            =   90
      TabIndex        =   35
      ToolTipText     =   "Novo Registro"
      Top             =   6720
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
      MICON           =   "frmLogradouro.frx":034A
      PICN            =   "frmLogradouro.frx":0366
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
      Left            =   1140
      TabIndex        =   36
      ToolTipText     =   "Editar Registro"
      Top             =   6720
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
      MICON           =   "frmLogradouro.frx":04C0
      PICN            =   "frmLogradouro.frx":04DC
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
      Left            =   2190
      TabIndex        =   37
      ToolTipText     =   "Excluir Registro"
      Top             =   6720
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
      MICON           =   "frmLogradouro.frx":0636
      PICN            =   "frmLogradouro.frx":0652
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
      Left            =   4680
      TabIndex        =   38
      ToolTipText     =   "Gravar os Dados"
      Top             =   6720
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmLogradouro.frx":06F4
      PICN            =   "frmLogradouro.frx":0710
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
      Left            =   4290
      TabIndex        =   39
      ToolTipText     =   "Consulta Cidadãos Cadastrados"
      Top             =   6720
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
      MICON           =   "frmLogradouro.frx":0AB5
      PICN            =   "frmLogradouro.frx":0AD1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Height          =   4050
      Left            =   45
      TabIndex        =   23
      Top             =   2610
      Width           =   7365
      Begin VB.TextBox txtAbreviado 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1890
         MaxLength       =   50
         TabIndex        =   6
         Top             =   950
         Width           =   5355
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Cadastro de Bairro"
         ForeColor       =   &H00000080&
         Height          =   1155
         Left            =   45
         TabIndex        =   42
         Top             =   2835
         Width           =   7260
         Begin VB.ComboBox cmbBairro 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   630
            Width           =   3165
         End
         Begin VB.TextBox txtBairroDe 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1125
            TabIndex        =   17
            Top             =   285
            Width           =   780
         End
         Begin VB.TextBox txtBairroAte 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3015
            TabIndex        =   18
            Top             =   285
            Width           =   825
         End
         Begin prjChameleon.chameleonButton cmdAddBairro 
            Height          =   285
            Left            =   3960
            TabIndex        =   20
            ToolTipText     =   "Adicionar CEP"
            Top             =   330
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   503
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
            MICON           =   "frmLogradouro.frx":0C2B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdDelBairro 
            Height          =   285
            Left            =   3960
            TabIndex        =   21
            ToolTipText     =   "Remover CEP"
            Top             =   660
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   503
            BTYPE           =   3
            TX              =   "-"
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
            MICON           =   "frmLogradouro.frx":0C47
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSFlexGridLib.MSFlexGrid grdBairro 
            Height          =   945
            Left            =   4410
            TabIndex        =   22
            Top             =   120
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   1667
            _Version        =   393216
            Rows            =   1
            Cols            =   4
            FixedCols       =   0
            BackColorFixed  =   15658734
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
            AllowUserResizing=   1
            BorderStyle     =   0
            Appearance      =   0
            FormatString    =   "Inicial  |Final   |CodBairro|Bairro                                                       "
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Inicial....:"
            Height          =   195
            Index           =   12
            Left            =   150
            TabIndex        =   45
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Final......:"
            Height          =   195
            Index           =   10
            Left            =   2025
            TabIndex        =   44
            Top             =   360
            Width           =   885
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro.:"
            Height          =   195
            Index           =   9
            Left            =   150
            TabIndex        =   43
            Top             =   675
            Width           =   555
         End
      End
      Begin esMaskEdit.esMaskedEdit mskData 
         Height          =   285
         Left            =   1890
         TabIndex        =   7
         Top             =   1290
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         MouseIcon       =   "frmLogradouro.frx":0C63
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Cadastro de CEP"
         ForeColor       =   &H00000080&
         Height          =   1155
         Left            =   45
         TabIndex        =   30
         Top             =   1665
         Width           =   7260
         Begin VB.CheckBox chkP 
            Appearance      =   0  'Flat
            BackColor       =   &H00EEEEEE&
            Caption         =   "&Par"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   3210
            TabIndex        =   11
            Top             =   270
            Value           =   1  'Checked
            Width           =   585
         End
         Begin VB.CheckBox chkI 
            Appearance      =   0  'Flat
            BackColor       =   &H00EEEEEE&
            Caption         =   "&Impar"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   2340
            TabIndex        =   10
            Top             =   270
            Value           =   1  'Checked
            Width           =   705
         End
         Begin VB.TextBox txtCEPAte 
            Appearance      =   0  'Flat
            Height          =   255
            Left            =   1170
            TabIndex        =   12
            Top             =   630
            Width           =   1005
         End
         Begin VB.TextBox txtCEPDe 
            Appearance      =   0  'Flat
            Height          =   255
            Left            =   1170
            TabIndex        =   9
            Top             =   330
            Width           =   1005
         End
         Begin prjChameleon.chameleonButton cmdAddCEP 
            Height          =   285
            Left            =   4140
            TabIndex        =   14
            ToolTipText     =   "Adicionar CEP"
            Top             =   330
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   503
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
            MICON           =   "frmLogradouro.frx":0C7F
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdDelCEP 
            Height          =   285
            Left            =   4140
            TabIndex        =   15
            ToolTipText     =   "Remover CEP"
            Top             =   660
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   503
            BTYPE           =   3
            TX              =   "-"
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
            MICON           =   "frmLogradouro.frx":0C9B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSFlexGridLib.MSFlexGrid grdCEP 
            Height          =   945
            Left            =   4500
            TabIndex        =   16
            Top             =   135
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   1667
            _Version        =   393216
            Rows            =   1
            Cols            =   5
            FixedCols       =   0
            BackColorFixed  =   15658734
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
            AllowUserResizing=   1
            BorderStyle     =   0
            Appearance      =   0
            FormatString    =   "Inicial  |Final   |^Imp|^Par|Cep           "
         End
         Begin esMaskEdit.esMaskedEdit mskCEP 
            Height          =   285
            Left            =   2970
            TabIndex        =   13
            Top             =   585
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   503
            MouseIcon       =   "frmLogradouro.frx":0CB7
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
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "CEP....:"
            Height          =   195
            Index           =   7
            Left            =   2310
            TabIndex        =   33
            Top             =   630
            Width           =   555
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Final......:"
            Height          =   195
            Index           =   6
            Left            =   150
            TabIndex        =   32
            Top             =   630
            Width           =   885
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Inicial....:"
            Height          =   195
            Index           =   5
            Left            =   150
            TabIndex        =   31
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.TextBox txtNomeLog 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1890
         MaxLength       =   50
         TabIndex        =   5
         Top             =   585
         Width           =   5355
      End
      Begin VB.TextBox txtCod 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1890
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   180
         Width           =   1005
      End
      Begin VB.ComboBox cmbTipo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3645
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   180
         Width           =   1050
      End
      Begin VB.ComboBox cmbTit 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5670
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   180
         Width           =   1050
      End
      Begin VB.TextBox txtNumOfic 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4845
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1305
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço Resumido....:"
         Height          =   195
         Index           =   13
         Left            =   90
         TabIndex        =   46
         Top             =   990
         Width           =   1725
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Logradouro...:"
         Height          =   195
         Index           =   11
         Left            =   90
         TabIndex        =   29
         Top             =   630
         Width           =   1725
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código do Logradouro.:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   28
         Top             =   240
         Width           =   1725
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo...:"
         Height          =   195
         Index           =   1
         Left            =   3060
         TabIndex        =   27
         Top             =   270
         Width           =   690
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Título.:"
         Height          =   195
         Index           =   2
         Left            =   4995
         TabIndex        =   26
         Top             =   270
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Oficialização.......:"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   25
         Top             =   1365
         Width           =   1725
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº do Ofício....:"
         Height          =   195
         Index           =   4
         Left            =   3540
         TabIndex        =   24
         Top             =   1350
         Width           =   1185
      End
   End
   Begin vbAcceleratorSGrid6.vbalGrid dgMain 
      Height          =   2040
      Left            =   45
      TabIndex        =   1
      Top             =   495
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   3598
      NoHorizontalGridLines=   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BackColor       =   16777215
      HighlightForeColor=   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   2
      DisableIcons    =   -1  'True
      HighlightSelectedIcons=   0   'False
      GroupBoxHintText=   "Arraste as colunas que deseja agrupar"
      SelectionOutline=   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código...:"
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   41
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmLogradouro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim sOldTipo As String, sOldTit As String, sOldNome As String, sOldData As String
Dim RdoAux As rdoResultset
Dim Sql As String
Dim Evento As String
Dim sRet As String
Dim evEdit As Integer, evNew As Integer, evDel As Integer
Dim bEdit As Boolean, bNew As Boolean, bDel As Boolean

Private Sub CarregaCombo()
cmbTit.AddItem ""
Sql = "SELECT CODTITLOG,NOMETITLOG,ABREVTITLOG FROM TITLOGRADOURO WHERE CODTITLOG<>9999"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
cmbTit.AddItem ""
cmbTit.ItemData(cmbTit.NewIndex) = 0
With RdoAux
    Do Until .EOF
       cmbTit.AddItem !AbrevTitLog
       cmbTit.ItemData(cmbTit.NewIndex) = !CODTITLOG
      .MoveNext
    Loop
End With

cmbTipo.AddItem ""
Sql = "SELECT CODTIPOLOG,NOMETIPOLOG,ABREVTIPOLOG FROM TIPOLOGRADOURO WHERE CODTIPOLOG<>9999"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
       cmbTipo.AddItem !AbrevTipoLog
       cmbTipo.ItemData(cmbTipo.NewIndex) = !CODTIPOLOG
      .MoveNext
    Loop
End With

Sql = "SELECT * FROM BAIRRO WHERE SIGLAUF='SP' AND CODCIDADE=413 AND CODBAIRRO<>999 ORDER BY DESCBAIRRO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
       cmbBairro.AddItem !DescBairro
       cmbBairro.ItemData(cmbBairro.NewIndex) = !CodBairro
      .MoveNext
    Loop
End With


End Sub


Private Sub cmdAddBairro_Click()
If Trim$(txtBairroDe.Text) = "" Then txtBairroDe.Text = 0
If Trim$(txtBairroAte.Text) = "" Then txtBairroAte.Text = 0

If Val(txtBairroAte.Text) > 0 Then
     If Val(txtBairroDe.Text) > Val(txtBairroAte.Text) Then
          MsgBox "O valor final tem que ser maior que o valor inicial.", vbExclamation, "Atenção"
          Exit Sub
     End If
End If

If cmbBairro.ListIndex = -1 Then
     MsgBox "Selecione o bairro.", vbExclamation, "Atenção"
     Exit Sub
End If

grdBairro.AddItem txtBairroDe.Text & Chr(9) & txtBairroAte.Text & Chr(9) & cmbBairro.ItemData(cmbBairro.ListIndex) & Chr(9) & cmbBairro.Text

End Sub

Private Sub cmdAddCEP_Click()

If Trim$(txtCEPDe.Text) = "" Then txtCEPDe.Text = 0
If Trim$(txtCEPAte.Text) = "" Then txtCEPAte.Text = 0

If Val(txtCEPAte.Text) > 0 Then
     If Val(txtCEPDe.Text) > Val(txtCEPAte.Text) Then
          MsgBox "O valor final tem que ser maior que o valor inicial.", vbExclamation, "Atenção"
          Exit Sub
     End If
End If

If Len(mskCEP.ClipText) < 8 Then
     MsgBox "Favor digitar o CEP.", vbExclamation, "Atenção"
     Exit Sub
End If

If chkI.value = 0 And chkP.value = 0 Then
     MsgBox "Selecione lado impar,par ou os dois.", vbExclamation, "Atenção"
     Exit Sub
End If

grdCEP.AddItem txtCEPDe.Text & Chr(9) & txtCEPAte.Text & Chr(9) & chkI.value & Chr(9) & chkP.value & Chr(9) & mskCEP.ClipText

End Sub

Private Sub cmdAlterar_Click()
    If txtCod.Text = "" Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
    sOldTipo = cmbTipo.Text
    sOldTit = cmbTit.Text
    sOldNome = txtNomeLog.Text
    sOldData = mskData.Text
    Eventos "INCLUIR"
    Evento = "Alterar"
End Sub

Private Sub cmdCancel_Click()
    Le
    Eventos "INICIAR"
    Evento = ""
End Sub

Private Sub cmdConsultar_Click()
frmCnsRua.show vbModeless
frmCnsRua.ZOrder 0
End Sub

Private Sub cmdDelBairro_Click()
If grdBairro.Row = 0 Then
     MsgBox "Selecione o Bairro a ser excluido.", vbExclamation, "Atenção"
Else
   If grdBairro.Rows > 2 Then
      grdBairro.RemoveItem (grdBairro.Row)
   Else
      grdBairro.Rows = 1
   End If
End If

End Sub

Private Sub cmdDelCEP_Click()

If grdCEP.Row = 0 Then
     MsgBox "Selecione o CEP a ser excluido.", vbExclamation, "Atenção"
Else
   If grdCEP.Rows > 2 Then
      grdCEP.RemoveItem (grdCEP.Row)
   Else
      grdCEP.Rows = 1
   End If
End If

End Sub

Private Sub cmdExcluir_Click()
Dim x As Integer
On Error GoTo Erro

If txtCod.Text = "" Then
   MsgBox "Não existem Registros.", vbCritical, "Atenção"
   Exit Sub
End If

Sql = "SELECT * From vwFACEQUADRA Where codlogr =" & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        MsgBox "Não é possível excluir, este logradouro esta cadastrado em um ou mais imóveis.", vbExclamation, "Atenção"
        Exit Sub
    End If
End With
If MsgBox("Excluir este Logradouro ?", vbQuestion + vbYesNoCancel, "Atenção") = vbYes Then
   Sql = "DELETE FROM CEP WHERE CODLOGR=" & txtCod.Text
   cn.Execute Sql, rdExecDirect
   Sql = "DELETE FROM LOGRADOURO WHERE CODLOGRADOURO=" & txtCod.Text
   cn.Execute Sql, rdExecDirect
   Log Form, Me.Caption, Exclusão, "Excluído registro " & Format(txtCod.Text, "000") & "-" & txtNomeLog.Text
   Limpa
   CarregaLista2
   Le
End If
Exit Sub

Erro:
For x = 0 To rdoErrors.Count - 1
    MsgBox rdoErrors(x).Description
Next

End Sub

Private Sub cmdGravar_Click()
    
 If NomeDeLogin <> "FACTORE" And NomeDeLogin <> "HELOISA" And NomeDeLogin <> "SCHWARTZ" Then
    MsgBox "Sem permissão de gravação.", vbCritical, "Erro"
    Exit Sub
 End If
    
    If Not IsDate(mskData.Text) And mskData.ClipText <> "" Then
       MsgBox "Data Inválida.", vbExclamation, "Atenção"
       mskData.SetFocus
       Exit Sub
    End If
    If IsDate(mskData.Text) Then
        If Year(CDate(mskData.Text)) < 1900 Or Year(CDate(mskData.Text)) > Year(Now) Then
           MsgBox "Data fora de intervalo válido.", vbExclamation, "Atenção"
           mskData.SetFocus
           Exit Sub
        End If
    End If
    If Trim$(txtNomeLog.Text) = "" Then
       MsgBox "Digite o Nome do Logradouro.", vbExclamation, "Atenção"
       txtNomeLog.SetFocus
       Exit Sub
    End If
    Grava
    Eventos "INICIAR"
End Sub


Private Sub cmdNovo_Click()
    Limpa
    Eventos "INCLUIR"
    Evento = "Novo"
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub dgMain_ColumnClick(ByVal lcol As Long)
Dim sTag As String
Dim iSortIndex As Long
      
   With dgMain.SortObject
      
      ' This demo allows grouping.  When a column is clicked
      ' for sorting, we only want to remove any grouped rows:
      .ClearNongrouped
      
      ' See if this column is already in the sort object:
      iSortIndex = .IndexOf(lcol)
      If (iSortIndex = 0) Then
         ' If not, we add it:
         iSortIndex = .Count + 1
         .SortColumn(iSortIndex) = lcol
      End If
   
      ' Determine which sort order to apply:
      sTag = dgMain.ColumnTag(lcol)
      If (sTag = "") Then
         sTag = "DESC"
         .SortOrder(iSortIndex) = CCLOrderAscending
      Else
         sTag = ""
         .SortOrder(iSortIndex) = CCLOrderDescending
      End If
      dgMain.ColumnTag(lcol) = sTag
      
      ' Set the type of sorting:
      .SortType(iSortIndex) = dgMain.ColumnSortType(lcol)
   End With
   
   ' Do the sort:
   Screen.MousePointer = vbHourglass
   dgMain.Sort
   Screen.MousePointer = vbDefault

End Sub

Private Sub dgMain_SelectionChange(ByVal lrow As Long, ByVal lcol As Long)
Le
End Sub

Private Sub Form_Activate()
Liberado
dgMain.SetFocus
End Sub

Private Sub Form_Load()

Centraliza Me
sRet = RetEventUserForm(Me.Name)
CarregaCombo
GridHeader
grdBairro.COLWIDTH(2) = 0
CarregaLista2
dgMain.SelectedRow = 1
Le

Eventos "INICIAR"

End Sub

Private Sub Eventos(Tipo As String)

Dim Ct As Control

If Tipo = "INICIAR" Then
   cmdNovo.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdConsultar.Visible = True
   cmdSair.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   cmdAddCEP.Enabled = False
   cmdDelCEP.Enabled = False
   cmdAddBairro.Enabled = False
   cmdDelBairro.Enabled = False
   For Each Ct In frmLogradouro
       If TypeOf Ct Is TextBox Or TypeOf Ct Is ComboBox Or TypeOf Ct Is esMaskedEdit Then
          Ct.BackColor = Kde
          Ct.Enabled = False
       End If
   Next
   txtPesq.Enabled = True
   txtPesq.BackColor = Branco
   dgMain.Enabled = True
'   grdLogr.Enabled = True
   chkI.Enabled = False
   chkP.Enabled = False
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdConsultar.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   cmdAddCEP.Enabled = True
   cmdDelCEP.Enabled = True
   cmdAddBairro.Enabled = True
   cmdDelBairro.Enabled = True
   For Each Ct In frmLogradouro
       If TypeOf Ct Is TextBox Or TypeOf Ct Is ComboBox Or TypeOf Ct Is esMaskedEdit Then
          Ct.BackColor = Branco
          Ct.Enabled = True
       End If
   Next
   txtCod.BackColor = Kde
   txtCod.Locked = True
   txtPesq.Enabled = False
   txtPesq.BackColor = Kde
   chkI.Enabled = True
   chkP.Enabled = True
   dgMain.Enabled = False
'   grdLogr.Enabled = False
   cmbTipo.SetFocus
End If

FormHagana

End Sub

Private Sub Le()

If dgMain.Rows = 0 Then Exit Sub
txtCod.Text = dgMain.CellText(dgMain.SelectedRow, 1)
If dgMain.CellText(dgMain.SelectedRow, 2) <> "" Then
    cmbTipo.Text = dgMain.CellText(dgMain.SelectedRow, 2)
Else
    cmbTipo.ListIndex = 0
End If
If Trim$(dgMain.CellText(dgMain.SelectedRow, 3)) <> "" Then
   cmbTit.Text = dgMain.CellText(dgMain.SelectedRow, 3)
Else
   cmbTit.ListIndex = 0
End If
txtNomeLog = dgMain.CellText(dgMain.SelectedRow, 4)

Sql = "SELECT DATAOFIC,NUMOFIC,endereco_resumido FROM LOGRADOURO WHERE CODLOGRADOURO=" & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
If RdoAux.RowCount > 0 Then
   If RdoAux!DATAOFIC > CDate("01/01/1920") Then
      mskData.Text = Format(RdoAux!DATAOFIC, "dd/mm/yyyy")
   Else
      LimpaMascara mskData
   End If
   txtAbreviado.Text = SubNull(RdoAux!endereco_resumido)
   txtNumOfic.Text = SubNull(RdoAux!NUMOFIC)
End If

grdCEP.Rows = 1
Sql = "SELECT CEP,VALOR1,VALOR2,IMPAR,PAR FROM CEP "
Sql = Sql & "WHERE CODLOGR=" & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    Do Until .EOF
         grdCEP.AddItem !VALOR1 & Chr(9) & IIf(IsNull(!VALOR2), 0, !VALOR2) & Chr(9) & IIf(!Impar, 1, 0) & Chr(9) & IIf(!Par, 1, 0) & Chr(9) & SubNull(!Cep)
        .MoveNext
    Loop
End With

If grdCEP.Rows > 1 Then
    grdCEP.col = 0
    grdCEP.ColSel = 4
    grdCEP.Row = 1
    grdCEP_Click
Else
    LimpaMascara mskCEP
    txtCEPAte.Text = ""
    txtCEPDe.Text = ""
    chkI.value = 1
    chkP.value = 1
End If

grdBairro.Rows = 1
Sql = "SELECT * FROM logradouro_bairro INNER JOIN bairro ON logradouro_BAIRRO.bairro= Bairro.codbairro "
Sql = Sql & "WHERE logradouro_bairro.logradouro=" & Val(txtCod.Text) & " AND siglauf='SP' AND codcidade=413"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    Do Until .EOF
         grdBairro.AddItem !INICIAL & Chr(9) & !FINAL & Chr(9) & !Bairro & Chr(9) & !DescBairro
        .MoveNext
    Loop
End With


End Sub

Private Sub Limpa()

txtCod.Text = ""
cmbTipo.ListIndex = -1
cmbTit.ListIndex = -1
LimpaMascara mskData
LimpaMascara mskCEP
txtNomeLog.Text = ""
txtNumOfic.Text = ""
chkI.value = 1
chkP.value = 1
txtCEPDe.Text = ""
txtAbreviado.Text = ""
txtCEPAte.Text = ""
grdCEP.Rows = 1
txtBairroDe.Text = ""
txtBairroAte.Text = ""
cmbBairro.ListIndex = -1

End Sub

Private Sub Grava()

Dim MaxCod As Integer, nTipo As Integer, nTit As Integer, sData As String, sNome As String

If cmbTipo.ListIndex = -1 Then
   nTipo = 9999
Else
   nTipo = cmbTipo.ItemData(cmbTipo.ListIndex)
End If
If cmbTit.ListIndex > -1 And cmbTit.Text <> "" Then
   nTit = cmbTit.ItemData(cmbTit.ListIndex)
Else
   nTit = 0
End If
If IsDate(mskData.Text) Then
   sData = Format(mskData.Text, "mm/dd/yyyy")
Else
   sData = ""
End If

Sql = "SELECT MAX(CODLOGRADOURO) AS MAXIMO FROM LOGRADOURO WHERE CODLOGRADOURO<9999"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!maximo) Then
    MaxCod = 1
Else
    MaxCod = RdoAux!maximo + 1
End If
RdoAux.Close

sNome = Trim$(cmbTipo.Text) & IIf(cmbTit.ListIndex < 1, "", " " & Trim$(SubNull(cmbTit.Text))) & " " & txtNomeLog.Text

If Evento = "Novo" Then
    Sql = "INSERT LOGRADOURO(CODLOGRADOURO,CODTIPOLOG,CODTITLOG,NOMELOGRADOURO,DATAOFIC,NUMOFIC,ENDERECO) VALUES("
    Sql = Sql & MaxCod & "," & IIf(nTipo = 0, "Null", nTipo) & "," & IIf(nTit = 0, "Null", nTit) & ",'" & Mask(txtNomeLog.Text) & "'," & IIf(sData = "", "Null", "'" & sData & "'") & "," & IIf(txtNumOfic.Text = "", "Null", "'" & txtNumOfic.Text & "'") & ",'" & Mask(sNome) & "')"
Else
    MaxCod = Val(txtCod.Text)
    Sql = "UPDATE LOGRADOURO SET CODTIPOLOG=" & IIf(nTipo = 0, "Null", nTipo) & ",CODTITLOG=" & IIf(nTit = 0, "Null", nTit) & ",NOMELOGRADOURO='" & Mask(txtNomeLog.Text) & "',DATAOFIC=" & IIf(sData = "", "Null", "'" & sData & "'") & ","
    Sql = Sql & "NUMOFIC=" & IIf(txtNumOfic.Text = "", "Null", "'" & txtNumOfic.Text & "'") & ",ENDERECO='" & Mask(sNome) & "' WHERE  CODLOGRADOURO = " & Val(txtCod.Text)
End If
cn.Execute Sql, rdExecDirect

'Grava CEP
Sql = "DELETE FROM CEP WHERE CODLOGR=" & MaxCod
cn.Execute Sql, rdExecDirect
      
For x = 1 To grdCEP.Rows - 1
    
      Sql = "INSERT CEP (CODLOGR,CEP,VALOR1,VALOR2,IMPAR,PAR) VALUES("
      Sql = Sql & MaxCod & "," & grdCEP.TextMatrix(x, 4) & "," & grdCEP.TextMatrix(x, 0) & "," & grdCEP.TextMatrix(x, 1) & ","
      Sql = Sql & grdCEP.TextMatrix(x, 2) & "," & grdCEP.TextMatrix(x, 3) & ")"
      cn.Execute Sql, rdExecDirect
Next

'Grava Bairro
Sql = "DELETE FROM LOGRADOURO_BAIRRO WHERE LOGRADOURO=" & MaxCod
cn.Execute Sql, rdExecDirect
      
For x = 1 To grdBairro.Rows - 1
    Sql = "INSERT LOGRADOURO_BAIRRO (LOGRADOURO,INICIAL,FINAL,BAIRRO) VALUES("
    Sql = Sql & MaxCod & "," & grdBairro.TextMatrix(x, 0) & "," & grdBairro.TextMatrix(x, 1) & "," & grdBairro.TextMatrix(x, 2) & ")"
    cn.Execute Sql, rdExecDirect
Next

If Evento = "Novo" Then
    dgMain.AddRow
    dgMain.CellDetails dgMain.Rows, 1, MaxCod
    dgMain.CellDetails dgMain.Rows, 2, cmbTipo.Text
    dgMain.CellDetails dgMain.Rows, 3, cmbTit.Text
    dgMain.CellDetails dgMain.Rows, 4, txtNomeLog.Text
    
'   grdLogr.AddItem MaxCod & Chr(9) & cmbTipo.Text & Chr(9) & cmbTit.Text & Chr(9) & txtNomeLog.Text
   txtCod.Text = MaxCod
'   grdLogr.Row = grdLogr.Rows - 1
'   grdLogr.ColSel = 3
   Log Form, Me.Caption, Inclusão, "Inserido registro " & Format(MaxCod, "000") & "-" & cmbTipo.Text & "-" & cmbTit.Text & "-" & txtNomeLog.Text
ElseIf Evento = "Alterar" Then
   dgMain.CellDetails dgMain.Rows, 2, cmbTipo.Text
   dgMain.CellDetails dgMain.Rows, 3, cmbTit.Text
   dgMain.CellDetails dgMain.Rows, 4, txtNomeLog.Text
   Log Form, Me.Caption, Alteração, "Alterado registro " & Format(txtCod.Text, "000") & " de " & sOldTipo & "-" & sOldTit & "-" & sOldNome & " para " & cmbTipo.Text & "-" & cmbTit.Text & "-" & txtNomeLog.Text
End If
      
            
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   If cmdNovo.Visible = True Then
      cmdNovo_Click
   Else
      cmdGravar_Click
   End If
End If
End Sub

Private Sub grdCEP_Click()

If grdCEP.Row > 0 Then
     txtCEPDe = grdCEP.TextMatrix(grdCEP.Row, 0)
     txtCEPAte = grdCEP.TextMatrix(grdCEP.Row, 1)
     chkI.value = grdCEP.TextMatrix(grdCEP.Row, 2)
     chkP.value = grdCEP.TextMatrix(grdCEP.Row, 3)
     If Len(grdCEP.TextMatrix(grdCEP.Row, 4)) = 8 Then
          mskCEP.Text = Format(grdCEP.TextMatrix(grdCEP.Row, 4), "#####-###")
     End If
End If

End Sub

Private Sub txtCEPAte_GotFocus()
txtCEPAte.SelStart = 0
txtCEPAte.SelLength = Len(txtCEPAte.Text)

End Sub

Private Sub txtCEPAte_KeyPress(KeyAscii As Integer)
Tweak txtCEPAte, KeyAscii, DecimalPositive

End Sub

Private Sub txtCEPDe_GotFocus()
txtCEPDe.SelStart = 0
txtCEPDe.SelLength = Len(txtCEPDe.Text)

End Sub

Private Sub txtCEPDe_KeyPress(KeyAscii As Integer)
Tweak txtCEPDe, KeyAscii, DecimalPositive

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

Private Sub txtPesq_Change()
CarregaLista2
End Sub

Private Sub GridHeader()

With dgMain
    .HeaderFlat = True
    .HeaderHeight = 18
    .DefaultRowHeight = 17
    .GridFillLineColor = vbWhite
    .RowMode = True
    .GridLines = True
    .GridLineMode = ecgGridFillControl
    .HighlightBackColor = Marrom
    .HighlightForeColor = vbWhite
        
    .AddColumn "kCodigo", "Código", ecgHdrTextALignLeft, , 50
    .AddColumn "kTipo", "Tipo", ecgHdrTextALignLeft, , 50
    .AddColumn "kTitulo", "Titulo", ecgHdrTextALignLeft, , 50
    .AddColumn "kNome", "Nome do Logradouro", ecgHdrTextALignLeft, , 280
    
End With

End Sub

Private Sub CarregaLista2()
Dim Sql As String, RdoAux As rdoResultset
Dim sPesq As String

sPesq = Trim(txtPesq.Text)

Ocupado
dgMain.Redraw = False
dgMain.Clear
dgMain.Redraw = True
dgMain.Redraw = False

Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO,DATAOFIC,NUMOFIC FROM vwLOGRADOURO WHERE 1=1 "

If sPesq <> "" Then
    If IsNumeric(Left(sPesq, 1)) Then
        Sql = Sql & " AND  CODLOGRADOURO=" & Val(txtPesq.Text)
    Else
        Sql = Sql & " AND  NOMELOGRADOURO like '%" & txtPesq.Text & "%'"
    End If
End If
Sql = Sql & " ORDER BY NOMELOGRADOURO"

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If cGetInputState() <> 0 Then DoEvents
        dgMain.AddRow
        dgMain.CellDetails dgMain.Rows, 1, !CodLogradouro
        dgMain.CellDetails dgMain.Rows, 2, SubNull(!AbrevTipoLog)
        dgMain.CellDetails dgMain.Rows, 3, SubNull(!AbrevTitLog)
        dgMain.CellDetails dgMain.Rows, 4, !NomeLogradouro
       .MoveNext
    Loop
   .Close
End With
Liberado
dgMain.Redraw = True
If dgMain.Rows > 0 Then
    dgMain.SelectedRow = 1
End If

End Sub

Private Sub txtPesq_Click()
txtPesq.SelStart = 0
txtPesq.SelLength = Len(txtPesq.Text)

End Sub

Private Sub txtPesq_GotFocus()
txtPesq.SelStart = 0
txtPesq.SelLength = Len(txtPesq.Text)
End Sub
