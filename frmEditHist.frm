VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmEditHist 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edi��o de Hist�rico do Im�vel"
   ClientHeight    =   4515
   ClientLeft      =   13290
   ClientTop       =   5775
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   7155
   Begin prjChameleon.chameleonButton cmdGravar 
      Height          =   315
      Left            =   5670
      TabIndex        =   4
      ToolTipText     =   "Gravar os Dados"
      Top             =   4110
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
      MICON           =   "frmEditHist.frx":0000
      PICN            =   "frmEditHist.frx":001C
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
      Left            =   4560
      TabIndex        =   9
      ToolTipText     =   "Cancelar Edi��o"
      Top             =   4110
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
      MICON           =   "frmEditHist.frx":03C1
      PICN            =   "frmEditHist.frx":03DD
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
      Left            =   2310
      TabIndex        =   10
      ToolTipText     =   "Excluir Registro"
      Top             =   4110
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
      MICON           =   "frmEditHist.frx":0537
      PICN            =   "frmEditHist.frx":0553
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
      TabIndex        =   11
      ToolTipText     =   "Editar Registro"
      Top             =   4110
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
      MICON           =   "frmEditHist.frx":05F5
      PICN            =   "frmEditHist.frx":0611
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
      Left            =   150
      TabIndex        =   12
      ToolTipText     =   "Novo Registro"
      Top             =   4110
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
      MICON           =   "frmEditHist.frx":076B
      PICN            =   "frmEditHist.frx":0787
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
      Left            =   5670
      TabIndex        =   13
      ToolTipText     =   "Sair da Tela"
      Top             =   4110
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
      MICON           =   "frmEditHist.frx":08E1
      PICN            =   "frmEditHist.frx":08FD
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
      Height          =   1785
      Left            =   30
      TabIndex        =   5
      Top             =   2220
      Width           =   7095
      Begin esMaskEdit.esMaskedEdit mskData 
         Height          =   300
         Left            =   1410
         TabIndex        =   1
         Top             =   165
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         MouseIcon       =   "frmEditHist.frx":096B
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
      Begin VB.TextBox txtHist 
         Appearance      =   0  'Flat
         Height          =   1125
         Left            =   1410
         MaxLength       =   4999
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   525
         Width           =   5535
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5940
         TabIndex        =   2
         Top             =   180
         Width           =   1005
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Data...................:"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   8
         Top             =   210
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Hist�rico.............:"
         Height          =   195
         Index           =   11
         Left            =   60
         TabIndex        =   7
         Top             =   585
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Seq.:......:"
         Height          =   195
         Index           =   0
         Left            =   5160
         TabIndex        =   6
         Top             =   225
         Width           =   705
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdHist 
      Height          =   2145
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3784
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedCols       =   0
      BackColorSel    =   192
      FocusRect       =   0
      SelectionMode   =   1
      Appearance      =   0
      FormatString    =   "^Data                |^Seq     |<Hist�rico                                                              |<Usu�rio                 "
   End
End
Attribute VB_Name = "frmEditHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nCodReduz As Long
Dim Evento As String
Dim NomeForm As String

Public Property Let sForm(sNomeForm As String)
    NomeForm = sNomeForm
End Property

Private Sub cmdAlterar_Click()
If txtSeq.Text = "" Then
   MsgBox "N�o existem Registros.", vbCritical, "Aten��o"
   Exit Sub
End If

If UCase(NomeDeLogin) = "RODRIGOC" Or NomeDeLogin = "RITA" Or UCase(NomeDeLogin) = "CARLOS.SANTOS" Or UCase(NomeDeLogin) = "GLEISE" Or UCase(NomeDeLogin) = "JOAOF" Or UCase(NomeDeLogin) = "DAYANE.IGLESIAS" Or UCase(NomeDeLogin) = "MARIELA.CUSTODIO" Or UCase(NomeDeLogin) = "HELOISA" Or UCase(NomeDeLogin) = "GEOVANI.FARIA" Or UCase(NomeDeLogin) = "MICHELLE.POLETTI" Then
    If grdHist.TextMatrix(grdHist.row, 3) <> "GTI" Then
       GoTo Fim
    Else
        If UCase(NomeDeLogin) <> "RODRIGOC" And UCase(NomeDeLogin) <> "HELOISA" And UCase(NomeDeLogin) <> "CARLOS.SANTOS" And UCase(NomeDeLogin) <> "GLEISE" And UCase(NomeDeLogin) <> "RODRIGOC" And UCase(NomeDeLogin) <> "DAYANE.IGLESIAS" And UCase(NomeDeLogin) <> "MARIELA.CUSTODIO" And NomeDeLogin <> "GEOVANI.FARIA" And NomeDeLogin <> "MICHELLE.POLETTI" Then
            MsgBox "Hist�rico gerado pelo sistema n�o pode ser alterado.", vbCritical, "Erro de Acesso"
            Exit Sub
        Else
            GoTo Fim
        End If
    End If
End If

If grdHist.TextMatrix(grdHist.row, 3) = "GTI/Iss.C.Civil" And (NomeDeLogin = "RODRIGOC" Or NomeDeLogin = "GLEISE" Or NomeDeLogin = "ANA.REIS" Or NomeDeLogin = "ROSANGELA") Then
    GoTo Fim2
'Else
'    MsgBox "Hist�rico de ISS constru��o civil somente pode ser alterado pelos fiscais.", vbCritical, "Erro de Acesso"
'    Exit Sub
End If
If grdHist.TextMatrix(grdHist.row, 3) <> NomeDeLogin Then
    MsgBox "Voce s� pode alterar os hist�ricos criados por voc�.", vbCritical, "Erro de Acesso"
    Exit Sub
End If

Fim:

If grdHist.TextMatrix(grdHist.row, 3) = "GTI/Iss.C.Civil" And NomeDeLogin <> "RODRIGOC" And NomeDeLogin <> "GLEISE" And NomeDeLogin <> "ANA.REIS" And NomeDeLogin <> "ROSANGELA" Then
    MsgBox "Hist�rico de ISS constru��o civil somente pode ser alterado pelos fiscais.", vbCritical, "Erro de Acesso"
    Exit Sub
End If
Fim2:

Eventos "INCLUIR"
Evento = "Alterar"

End Sub

Private Sub cmdCancel_Click()
    Le
    Eventos "INICIAR"
    Evento = ""

End Sub

Private Sub cmdExcluir_Click()
If txtSeq.Text = "" Then
   MsgBox "N�o existem Registros.", vbCritical, "Aten��o"
   Exit Sub
End If

If grdHist.TextMatrix(grdHist.row, 3) = "GTI" And NomeDeLogin <> "SCHWARTZ" Then
    MsgBox "Voce n�o pode excluir um hist�rico gerado pelo sistema!", vbCritical, "ERRO DE ACESSO"
    Exit Sub
End If

If NomeDeLogin = "RODRIGOC" Or NomeDeLogin = "RITA" Or NomeDeLogin = "SCHWARTZ" Or NomeDeLogin = "RODRIGOC" Or NomeDeLogin = "HELOISA" Or NomeDeLogin = "MARIELA" Or NomeDeLogin = "REGINA" Or NomeDeLogin = "JOAOF" Or NomeDeLogin = "DAYANE.IGLESIAS" Or NomeDeLogin = "MARIELA.CUSTODIO" Or NomeDeLogin = "GLEISE" Or NomeDeLogin = "LEANDRO" Then
   GoTo Fim
End If

If grdHist.TextMatrix(grdHist.row, 3) <> NomeDeLogin Then
    MsgBox "Voce s� pode alterar os hist�ricos criados por voc�.", vbCritical, "Erro de Acesso"
    Exit Sub
End If

Fim:
With grdHist
    If .Rows > 2 Then
       .RemoveItem (.row)
    Else
       .Rows = 1
    End If
End With
grdHist_Click
End Sub

Private Sub cmdGravar_Click()
Dim x As Integer, Achou As Boolean

If Evento = "Novo" Then
    With grdHist
        Achou = False
        For x = 1 To .Rows - 1
            If Val(.TextMatrix(x, 1)) = Val(txtSeq.Text) Then
                Achou = True
                Exit For
            End If
        Next
    End With
    If Achou Then
       MsgBox "Sequencia j� cadastrada..", vbExclamation, "Aten��o"
       txtSeq.SetFocus
       Exit Sub
    End If
End If
If Val(Mid(mskData.Text, 4, 2)) > 12 Then
   MsgBox "Data inv�lida.", vbExclamation, "Aten��o"
   mskData.SetFocus
   Exit Sub
End If
If Not IsDate(mskData.Text) Then
   MsgBox "Data inv�lida.", vbExclamation, "Aten��o"
   mskData.SetFocus
   Exit Sub
End If

'If Val(txtSeq.Text) > 100 Then
'   MsgBox "N� de Sequencia inv�lida.", vbExclamation, "Aten��o"
'   txtSeq.SetFocus
'   Exit Sub
'End If

If Val(txtSeq.Text) = 0 Then
   MsgBox "Favor digitar a sequencia.", vbExclamation, "Aten��o"
   txtSeq.SetFocus
   Exit Sub
End If

If txtHist.Text = "" Then
   MsgBox "Favor digitar o hist�rico.", vbExclamation, "Aten��o"
   txtHist.SetFocus
   Exit Sub
End If
Grava
Eventos "INICIAR"

End Sub

Private Sub Grava()

With grdHist
    If Evento = "Novo" Then
        .AddItem mskData.Text & Chr(9) & Format(txtSeq.Text, "00") & Chr(9) & txtHist.Text & Chr(9) & NomeDeLogin
    ElseIf Evento = "Alterar" Then
        .TextMatrix(.row, 0) = mskData.Text
        .TextMatrix(.row, 1) = Format(txtSeq.Text, "00")
        .TextMatrix(.row, 2) = txtHist.Text
        .TextMatrix(.row, 3) = NomeDeLogin
    End If
End With
End Sub

Private Sub cmdNovo_Click()
    Limpa
    Eventos "INCLUIR"
    Evento = "Novo"

End Sub

Private Sub Limpa()

LimpaMascara mskData
txtSeq.Text = ""
txtHist.Text = ""

End Sub

Private Sub cmdSair_Click()
Dim x As Integer

If MsgBox("Deseja atualizar o hist�rico do im�vel ?", vbQuestion + vbYesNo, "Confirma��o !!!") = vbYes Then
    If NomeForm = "frmCadMob" Then
        frmCadMob.grdHist.Rows = 1
        For x = 1 To grdHist.Rows - 1
            With grdHist
                frmCadMob.grdHist.AddItem .TextMatrix(x, 0) & Chr(9) & .TextMatrix(x, 1) & Chr(9) & .TextMatrix(x, 2) & Chr(9) & .TextMatrix(x, 3)
            End With
        Next
Else
        frmCadImob.grdHist.Rows = 1
        For x = 1 To grdHist.Rows - 1
            With grdHist
                frmCadImob.grdHist.AddItem .TextMatrix(x, 0) & Chr(9) & .TextMatrix(x, 1) & Chr(9) & .TextMatrix(x, 2) & Chr(9) & .TextMatrix(x, 3)
            End With
        Next
    End If
End If
Unload Me

End Sub

Private Sub Form_Load()
Eventos "INICIAR"
If NomeForm = "frmCadImob" Then
    nCodReduz = Val(Left$(frmCadImob.lblCodReduz.Caption, 7))
    CodEmpresa = 0
Else
    nCodReduz = 0
    CodEmpresa = Val(frmCadMob.txtCodEmpresa.Text)
End If
CarregaGrid
If grdHist.Rows = 1 Then Exit Sub
grdHist.row = 1
grdHist.ColSel = 3
Le

End Sub

Private Sub CarregaGrid()
Dim x As Integer

grdHist.Rows = 1

If NomeForm = "frmCadMob" Then
    For x = 1 To frmCadMob.grdHist.Rows - 1
        With frmCadMob.grdHist
            grdHist.AddItem .TextMatrix(x, 0) & Chr(9) & .TextMatrix(x, 1) & Chr(9) & .TextMatrix(x, 2) & Chr(9) & .TextMatrix(x, 3)
        End With
    Next
Else
    For x = 1 To frmCadImob.grdHist.Rows - 1
        With frmCadImob.grdHist
            grdHist.AddItem .TextMatrix(x, 0) & Chr(9) & .TextMatrix(x, 1) & Chr(9) & .TextMatrix(x, 2) & Chr(9) & .TextMatrix(x, 3)
        End With
    Next
End If
End Sub

Private Sub grdHist_Click()
Le
End Sub

Private Sub Le()
If grdHist.Rows = 1 Then
   txtHist.Text = ""
   Exit Sub
End If
If grdHist.row > 0 Then
    With grdHist
        mskData.Text = .TextMatrix(.row, 0)
        txtSeq.Text = .TextMatrix(.row, 1)
        txtHist.Text = .TextMatrix(.row, 2)
    End With
End If
End Sub

Private Sub Eventos(Tipo As String)

Dim Ct As Control

If Tipo = "INICIAR" Then
   cmdNovo.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdSair.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   For Each Ct In frmEditHist
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = Kde
          Ct.Enabled = False
       End If
   Next
   mskData.Enabled = False
   mskData.BackColor = Kde
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   For Each Ct In frmEditHist
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = Branco
          Ct.Enabled = True
       End If
   Next
   mskData.Enabled = True
   mskData.BackColor = Branco
   mskData.SetFocus
End If

End Sub

Private Sub txtSeq_KeyPress(KeyAscii As Integer)

Tweak txtSeq, KeyAscii, IntegerPositive

End Sub
