VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmTransfereDivida 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transferência de Dívida"
   ClientHeight    =   8100
   ClientLeft      =   3975
   ClientTop       =   1965
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8100
   ScaleWidth      =   11175
   Begin VB.Frame Frame3 
      Caption         =   "Divisão das parcelas"
      Height          =   3705
      Left            =   90
      TabIndex        =   19
      Top             =   4410
      Width           =   7665
      Begin MSComctlLib.ListView lvDestino 
         Height          =   3255
         Left            =   135
         TabIndex        =   20
         Top             =   315
         Width           =   7350
         _ExtentX        =   12965
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Código(s) de destino"
      Height          =   2760
      Left            =   7830
      TabIndex        =   7
      Top             =   675
      Width           =   3210
      Begin VB.TextBox txtCodigoDestino 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   855
         MaxLength       =   6
         TabIndex        =   10
         Top             =   1980
         Width           =   795
      End
      Begin VB.TextBox txtPerc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   855
         MaxLength       =   6
         TabIndex        =   9
         Top             =   2295
         Width           =   795
      End
      Begin MSComctlLib.ListView lvCodigo 
         Height          =   1455
         Left            =   180
         TabIndex        =   8
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "% Apl"
            Object.Width           =   1305
         EndProperty
      End
      Begin prjChameleon.chameleonButton cmdCancel 
         Height          =   315
         Left            =   2655
         TabIndex        =   13
         ToolTipText     =   "Cancelar Edição"
         Top             =   2250
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
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
         MICON           =   "frmTransfereDivida.frx":0000
         PICN            =   "frmTransfereDivida.frx":001C
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
         Left            =   1755
         TabIndex        =   14
         ToolTipText     =   "Novo Registro"
         Top             =   2250
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
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
         MICON           =   "frmTransfereDivida.frx":0176
         PICN            =   "frmTransfereDivida.frx":0192
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
         Left            =   2210
         TabIndex        =   15
         ToolTipText     =   "Editar Registro"
         Top             =   2250
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
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
         MICON           =   "frmTransfereDivida.frx":02EC
         PICN            =   "frmTransfereDivida.frx":0308
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
         Left            =   2655
         TabIndex        =   16
         ToolTipText     =   "Excluir Registro"
         Top             =   2250
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
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
         MICON           =   "frmTransfereDivida.frx":0462
         PICN            =   "frmTransfereDivida.frx":047E
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
         Left            =   2210
         TabIndex        =   17
         ToolTipText     =   "Gravar os Dados"
         Top             =   2250
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTransfereDivida.frx":0520
         PICN            =   "frmTransfereDivida.frx":053C
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
         Caption         =   "% Apl.:"
         Height          =   225
         Index           =   1
         Left            =   210
         TabIndex        =   12
         Top             =   2340
         Width           =   495
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   2040
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "        Origem do débito"
      Height          =   3660
      Left            =   90
      TabIndex        =   4
      Top             =   675
      Width           =   7665
      Begin VB.CheckBox chkOrigemAll 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   20
         Width           =   240
      End
      Begin MSComctlLib.ListView lvOrigem 
         Height          =   3120
         Left            =   135
         TabIndex        =   5
         Top             =   315
         Width           =   7350
         _ExtentX        =   12965
         _ExtentY        =   5503
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Ano"
            Object.Width           =   1305
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Lc"
            Object.Width           =   776
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "descrição"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Sq"
            Object.Width           =   776
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Pc"
            Object.Width           =   776
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Cp"
            Object.Width           =   776
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "status"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Vencto"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Valor"
            Object.Width           =   1766
         EndProperty
      End
   End
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   2475
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   180
      Width           =   7395
   End
   Begin VB.TextBox txtCodigoOrigem 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   1
      Top             =   180
      Width           =   1005
   End
   Begin prjChameleon.chameleonButton cmdCarregar 
      Height          =   360
      Left            =   9945
      TabIndex        =   3
      ToolTipText     =   "Carregar débito"
      Top             =   135
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "&Carregar"
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
      MCOL            =   0
      MPTR            =   1
      MICON           =   "frmTransfereDivida.frx":08E1
      PICN            =   "frmTransfereDivida.frx":08FD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdExec 
      Height          =   360
      Left            =   9450
      TabIndex        =   18
      ToolTipText     =   "Pré visualizar valores"
      Top             =   3915
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "Pré Visualizar"
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
      MCOL            =   0
      MPTR            =   1
      MICON           =   "frmTransfereDivida.frx":0A84
      PICN            =   "frmTransfereDivida.frx":0AA0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdExec2 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   9450
      TabIndex        =   21
      ToolTipText     =   "Executar o desdobro de dívida"
      Top             =   7650
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "&Executar"
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
      MICON           =   "frmTransfereDivida.frx":0C27
      PICN            =   "frmTransfereDivida.frx":0C43
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Tributacao.XP_ProgressBar PBar 
      Height          =   240
      Left            =   7920
      TabIndex        =   22
      Top             =   7290
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   16777215
      Scrolling       =   1
      ShowText        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Código origem...:"
      Height          =   240
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Width           =   1230
   End
End
Attribute VB_Name = "frmTransfereDivida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Debito
    nAno As Integer
    nLanc As Integer
    sLanc As String
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
    nSituacao As Integer
    sSituacao As String
    sVencto As String
    sDA As String
    sAj As String
    nCodTributo As Double
    nValorTributo As Double
    nValorJuros As Double
    nValorMulta As Double
    nValorCorrecao As Double
    nValorAtual As Double
    nValorHon As Double
    nValorJurApl As Double
    nSaldo As Double
    nCodBanco As Integer
    dDataPag As Date
    sNotificado As String
    sExFiscal As String
    nProt_certidao As Long
    nProt_dtremessa As Date
End Type

Private Type tCodigos
    nCodigo As Long
    nPerc As Double
End Type

Dim Evento As String, aCodigos() As tCodigos

Private Sub CallPb(nVal As Long, nTot As Long)
If nVal > 0 Then
    PBar.Color = &HC0C000
Else
    PBar.Color = vbWhite
End If
If ((nVal * 100) / nTot) <= 100 Then
   PBar.value = (nVal * 100) / nTot
Else
   PBar.value = 100
End If

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

End Sub


Private Sub chkOrigemAll_Click()
Dim x As Integer

For x = 1 To lvOrigem.ListItems.Count
    lvOrigem.ListItems(x).Checked = IIf(chkOrigemAll.value = vbChecked, True, False)
Next

End Sub

Private Sub cmdAlterar_Click()
    If txtCodigoDestino.Text = "" Then
       MsgBox "Selecione um código.", vbCritical, "Atenção"
       Exit Sub
    End If
    Eventos "INCLUIR"
    Evento = "Alterar"
End Sub

Private Sub cmdCancel_Click()
txtCodigoDestino.Text = ""
txtPerc.Text = ""
Eventos "INICIAR"
End Sub

Private Sub cmdCarregar_Click()
Dim nCodigo As Long, Sql As String, RdoAux As rdoResultset

nCodigo = Val(txtCodigoOrigem.Text)
If nCodigo = 0 Then
    MsgBox "Digite um código", vbCritical, "Erro"
    Exit Sub
End If
txtNome.Text = ""
If nCodigo < 100000 Then
    Sql = "SELECT NOMECIDADAO AS NOME FROM vwfullimovel2 WHERE CODREDUZIDO=" & nCodigo
ElseIf nCodigo >= 100000 And nCodigo < 200000 Then
    Sql = "SELECT RAZAOSOCIAL AS NOME FROM MOBILIARIO WHERE CODIGOMOB=" & nCodigo
Else
    Sql = "SELECT NOMECIDADAO AS NOME FROM CIDADAO WHERE CODCIDADAO=" & nCodigo
End If
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount = 0 Then
    MsgBox "Código não cadastrado", vbCritical, "Erro"
    Exit Sub
Else
    txtNome.Text = RdoAux!Nome
End If
RdoAux.Close

Sql = "select * from debitoparcela where codreduzido=" & nCodigo
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount = 0 Then
    MsgBox "Não existe dívida para este código", vbCritical, "Erro"
    Exit Sub
End If
RdoAux.Close

Ocupado
DoEvents
Carrega_Origem nCodigo
Liberado

End Sub

Private Sub cmdExcluir_Click()
    If txtCodigoDestino.Text = "" Then
       MsgBox "Selecione um código.", vbCritical, "Atenção"
       Exit Sub
    End If
    
    lvCodigo.ListItems.Remove (lvCodigo.SelectedItem.Index)
    txtCodigoDestino.Text = ""
    txtPerc.Text = ""
    
End Sub

Private Sub cmdExec_Click()
Dim x As Integer, nSoma As Double, nFalta As Double

nSoma = 0
For x = 1 To lvCodigo.ListItems.Count
    nSoma = nSoma + CDbl(lvCodigo.ListItems(x).SubItems(1))
Next

If nSoma <> 100 Then
    nFalta = 100 - nSoma
    MsgBox "Soma do percentual tem quer ser igual a 100%. " & vbCrLf & "(Soma = " & nSoma & ", diferença: " & nFalta & ")", vbExclamation, "Atenção"
    Exit Sub
End If

If lvOrigem.ListItems.Count = 0 Then
   MsgBox "Selecione a dívida de origem.", vbExclamation, "Atenção"
   Exit Sub
End If


Carrega_Destino

End Sub

Private Sub cmdExec2_Click()
If lvDestino.ListItems.Count = 0 Then
    MsgBox "Nada a transeferir", vbCritical, "Erro"
    Exit Sub
End If

If MsgBox("Deseja efetuar a transferência desta dívida?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
    Ocupado
    DoEvents
    Efetua_Transferencia
    Liberado
End If
End Sub

Private Sub cmdGravar_Click()
Dim x As Integer, bFind As Boolean

lvDestino.ListItems.Clear

If Val(txtCodigoDestino.Text) = 0 Or Val(txtPerc.Text) = 0 Then
   MsgBox "Favor digitar um código e um percentual.", vbExclamation, "Atenção"
   txtCodigoDestino.SetFocus
   Exit Sub
End If

If Val(txtPerc.Text) > 100 Then
   MsgBox "Favor digitar um percentual válido.", vbExclamation, "Atenção"
   txtPerc.SetFocus
   Exit Sub
End If

If Evento = "Novo" Then
    bFind = False
    For x = 1 To lvCodigo.ListItems.Count
        If lvCodigo.ListItems(x).Text = txtCodigoDestino.Text Then
            bFind = True
            Exit For
        End If
    Next
    If bFind Then
       MsgBox "Código já cadastrado.", vbExclamation, "Atenção"
       txtPerc.SetFocus
       Exit Sub
    End If
End If

nCodigo = Val(txtCodigoDestino.Text)
If nCodigo < 100000 Then
    Sql = "SELECT NOMECIDADAO AS NOME FROM vwfullimovel2 WHERE CODREDUZIDO=" & nCodigo
ElseIf nCodigo >= 100000 And nCodigo < 200000 Then
    Sql = "SELECT RAZAOSOCIAL AS NOME FROM MOBILIARIO WHERE CODIGOMOB=" & nCodigo
Else
    Sql = "SELECT NOMECIDADAO AS NOME FROM CIDADAO WHERE CODCIDADAO=" & nCodigo
End If
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount = 0 Then
    MsgBox "Código não cadastrado", vbCritical, "Erro"
    Exit Sub
End If
RdoAux.Close

If Evento = "Novo" Then
    Set itmX = lvCodigo.ListItems.Add(, , txtCodigoDestino.Text)
    itmX.SubItems(1) = txtPerc.Text
Else
    lvCodigo.ListItems(lvCodigo.SelectedItem.Index).Text = txtCodigoDestino.Text
    lvCodigo.ListItems(lvCodigo.SelectedItem.Index).SubItems(1) = txtPerc.Text
End If
Eventos "INICIAR"
End Sub

Private Sub cmdNovo_Click()
    txtCodigoDestino.Text = ""
    txtPerc.Text = ""
    Eventos "INCLUIR"
    Evento = "Novo"
End Sub

Private Sub Form_Load()
Centraliza Me
Eventos "INICIAR"
End Sub



Private Sub lvCodigo_Click()
If lvCodigo.ListItems.Count > 0 Then
    txtCodigoDestino.Text = lvCodigo.ListItems(lvCodigo.SelectedItem.Index).Text
    txtPerc.Text = lvCodigo.ListItems(lvCodigo.SelectedItem.Index).SubItems(1)
End If
End Sub

Private Sub txtPerc_KeyPress(KeyAscii As Integer)
Tweak txtPerc, KeyAscii, DecimalPositive, 2
End Sub

Private Sub txtCodigoDestino_KeyPress(KeyAscii As Integer)
Tweak txtCodigoDestino, KeyAscii, IntegerPositive
End Sub

Private Sub txtCodigoOrigem_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    cmdCarregar_Click
Else
    Tweak txtCodigoOrigem, KeyAscii, IntegerPositive
End If

End Sub

Private Sub Eventos(Tipo As String)

If Tipo = "INICIAR" Then
   cmdNovo.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   lvCodigo.Enabled = True
   txtCodigoDestino.Locked = True
   txtCodigoDestino.BackColor = Kde
   txtPerc.Locked = True
   txtPerc.BackColor = Kde
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   lvCodigo.Enabled = False
   txtCodigoDestino.Locked = False
   txtCodigoDestino.BackColor = Branco
   txtPerc.Locked = False
   txtPerc.BackColor = Branco
   txtCodigoDestino.SetFocus
End If

End Sub

Private Sub Carrega_Origem(Codigo As Long)
Dim qd As New rdoQuery, Sql As String, aDebito() As Debito, x As Integer
ReDim aDebito(0)
lvOrigem.ListItems.Clear
DoEvents
Set qd.ActiveConnection = cn
qd.QueryTimeout = 0
On Error Resume Next
RdoAux.Close
On Error GoTo 0
qd.Sql = "{ Call spEXTRATONEW(?,?) }"
qd(0) = Codigo
qd(1) = Codigo
Set RdoAux = qd.OpenResultset(rdOpenKeyset)

With RdoAux
    If RdoAux.RowCount > 0 Then
        ReDim Preserve aDebito(UBound(aDebito) + 1)
        nEval = UBound(aDebito)
        Do Until .EOF
            'If !statuslanc <> 3 Then GoTo Proximo
            bJuros = False: bMulta = False: bIsentoMJ = False
            If Not IsNull(!NumDocumento) Then
                Sql = "SELECT NUMDOCUMENTO,ISENTOMJ FROM NUMDOCUMENTO WHERE NUMDOCUMENTO=" & RdoAux!NumDocumento
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
                If RdoAux2.RowCount > 0 Then
                    If Val(SubNull(RdoAux2!isentomj)) > 0 Then
                        bIsentoMJ = True
                    End If
                End If
                RdoAux2.Close
            End If
            nEval = UBound(aDebito)
            Achou = False
            For x = 1 To nEval
                If aDebito(x).nAno = !AnoExercicio And aDebito(x).nLanc = !CodLancamento And _
                   aDebito(x).nSeq = !SeqLancamento And _
                   aDebito(x).nParc = !NumParcela And aDebito(x).nCompl = !CODCOMPLEMENTO Then
                   Achou = True
                   Exit For
                End If
            Next
            
            If Not Achou Then
                ReDim Preserve aDebito(UBound(aDebito) + 1)
                nEval = UBound(aDebito)
                aDebito(nEval).nAno = !AnoExercicio
                aDebito(nEval).nLanc = !CodLancamento
                If !CodLancamento = 20 Or !CodLancamento = 8 Then
                   If Not IsNull(!NumProcesso) Then
                      If Val(Right$(!NumProcesso, 4)) >= 2006 Then
                        aDebito(nEval).sLanc = !DESCLANCAMENTO & " (" & Left$(!NumProcesso, InStr(1, !NumProcesso, "/", vbBinaryCompare) - 1) & "-" & RetornaDVProcesso(Left$(!NumProcesso, InStr(1, !NumProcesso, "/", vbBinaryCompare) - 1)) & "/" & Right$(!NumProcesso, 4) & ")"
                      Else
                        aDebito(nEval).sLanc = !DESCLANCAMENTO & " (" & !NumProcesso & ")"
                      End If
                   Else
                      aDebito(nEval).sLanc = !DESCLANCAMENTO
                   End If
                Else
                   aDebito(nEval).sLanc = !DESCLANCAMENTO
                End If
                aDebito(nEval).nSeq = !SeqLancamento
                aDebito(nEval).nParc = !NumParcela
                aDebito(nEval).nCompl = !CODCOMPLEMENTO
                aDebito(nEval).nSituacao = !statuslanc
                aDebito(nEval).sSituacao = !Situacao
                aDebito(nEval).sVencto = Format(!DataVencimento, "dd/mm/yyyy")
                aDebito(nEval).sDA = IIf(IsNull(!datainscricao), "N", "S")
                aDebito(nEval).sAj = IIf(IsNull(!dataajuiza), "N", "S")
                aDebito(nEval).nCodTributo = !CodTributo
                aDebito(nEval).nValorTributo = FormatNumber(!VALORTRIBUTO, 2)
               
                
                If !statuslanc = 1 Or !statuslanc = 2 Or !statuslanc = 7 Then
                    If Not IsNull(!ValorPagoreal) Then
                        aDebito(nEval).nValorAtual = FormatNumber(!ValorPagoreal, 2)
                    Else
                        aDebito(nEval).nValorAtual = FormatNumber(0, 2)
                    End If
                Else
                    If bIsentoMJ Then
                        aDebito(nEval).nValorAtual = FormatNumber(!VALORTRIBUTO + !valorcorrecao, 2)
                    Else
                        aDebito(nEval).nValorAtual = FormatNumber(!ValorTotal, 2)
                    End If
                End If
                If IsNull(!notificado) Then
                    aDebito(nEval).sNotificado = "N"
                Else
                    aDebito(nEval).sNotificado = IIf(!notificado = True, "S", "N")
                End If
                
                sExecFiscal = ""
                If Not IsNull(!processocnj) Then
                    sExecFiscal = !processocnj
                Else
                    If Not IsNull(!anoexecfiscal) Then
                        sExecFiscal = Format(!numexecfiscal, "00000") & "/" & !anoexecfiscal
                    End If
                End If
                aDebito(nEval).sExFiscal = sExecFiscal
                aDebito(nEval).nProt_certidao = Val(SubNull(!prot_certidao))
                If IsNull(!prot_dtremessa) Then
                    aDebito(nEval).nProt_dtremessa = CDate("01/01/1900")
                Else
                    aDebito(nEval).nProt_dtremessa = Format(!prot_dtremessa, "dd/mm/yyyy")
                End If
            Else
                bFind = False
                For k = 1 To UBound(aDebito)
                    If aDebito(k).nAno = !AnoExercicio And aDebito(k).nLanc = !CodLancamento And _
                       aDebito(k).nSeq = !SeqLancamento And aDebito(k).nParc = !NumParcela And _
                       aDebito(k).nCompl = !CODCOMPLEMENTO And aDebito(k).nCodTributo = !CodTributo Then
                       bFind = True
                       Exit For
                    End If
                Next
                
                If Not bFind Then
                    aDebito(x).nValorTributo = FormatNumber(aDebito(x).nValorTributo + !VALORTRIBUTO, 2)
                    If !statuslanc = 1 Or !statuslanc = 2 Or !statuslanc = 7 Then
'                         aDebito(x).nValorAtual = FormatNumber(aDebito(x).nValorAtual + !ValorTributo, 2)
                    Else
                        If bIsentoMJ Then
                            aDebito(x).nValorAtual = FormatNumber(aDebito(x).nValorAtual + !VALORTRIBUTO + !valorcorrecao, 2)
                        Else
                            aDebito(x).nValorAtual = FormatNumber(aDebito(x).nValorAtual + !ValorTotal, 2)
                        End If
                    End If
                End If
            End If
Proximo:
            .MoveNext
        Loop
      End If
   .Close
End With

For x = 1 To UBound(aDebito)
    With aDebito(x)
        If .nAno = 0 Then GoTo Proximo2
        Set itmX = lvOrigem.ListItems.Add(, , .nAno)
        itmX.SubItems(1) = .nLanc
        itmX.SubItems(2) = .sLanc
        itmX.SubItems(3) = .nSeq
        itmX.SubItems(4) = .nParc
        itmX.SubItems(5) = .nCompl
        itmX.SubItems(6) = .nSituacao
        itmX.SubItems(7) = .sVencto
        itmX.SubItems(8) = Format(.nValorTributo, "#0.00")
        
    End With
Proximo2:
Next

End Sub

Private Sub Carrega_Destino()
Dim x As Integer, Sql As String, RdoAux As rdoResultset, nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer
Dim sDataVencto As String, sValor As String, aPerc() As Double, y As Integer
ReDim aPerc(0)
ReDim aCodigos(0)
lvDestino.ListItems.Clear


Inicio:
For x = 1 To lvDestino.ColumnHeaders.Count
    lvDestino.ColumnHeaders.Remove (x)
    GoTo Inicio
Next

lvDestino.ColumnHeaders.Add , , "Ano", 740, lvwColumnLeft
lvDestino.ColumnHeaders.Add , , "Lc", 440, lvwColumnCenter
lvDestino.ColumnHeaders.Add , , "Sq", 440, lvwColumnCenter
lvDestino.ColumnHeaders.Add , , "Pc", 440, lvwColumnCenter
lvDestino.ColumnHeaders.Add , , "Cp", 440, lvwColumnCenter
lvDestino.ColumnHeaders.Add , , "Valor", 1000, lvwColumnCenter

For x = 1 To lvCodigo.ListItems.Count
    lvDestino.ColumnHeaders.Add , , lvCodigo.ListItems(x).Text, 800, lvwColumnCenter
    ReDim Preserve aPerc(UBound(aPerc) + 1)
    aPerc(UBound(aPerc)) = CDbl(lvCodigo.ListItems(x).SubItems(1))
    ReDim Preserve aCodigos(UBound(aCodigos) + 1)
    aCodigos(UBound(aCodigos)).nCodigo = lvCodigo.ListItems(x).Text
    aCodigos(UBound(aCodigos)).nPerc = lvCodigo.ListItems(x).SubItems(1)
Next

For x = 1 To lvOrigem.ListItems.Count
    With lvOrigem.ListItems(x)
        If lvOrigem.ListItems(x).Checked Then
            nAno = .Text
            nLanc = .SubItems(1)
            nSeq = .SubItems(3)
            nParc = .SubItems(4)
            nCompl = .SubItems(5)
            sValor = .SubItems(8)
            
            Set itmX = lvDestino.ListItems.Add(, , nAno)
            itmX.SubItems(1) = nLanc
            itmX.SubItems(2) = nSeq
            itmX.SubItems(3) = nParc
            itmX.SubItems(4) = nCompl
            itmX.SubItems(5) = sValor
            For y = 1 To UBound(aPerc)
                itmX.SubItems(5 + y) = Format(CDbl(sValor) * aPerc(y) / 100, "#0.00")
            Next
        End If
    End With
    
Next

End Sub

Private Sub Efetua_Transferencia()
Dim Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos As Integer, nTot As Integer, nPos2 As Integer, nValor As Double, nCol As Integer
Dim nCodigo_Origem As Long, nAno As Integer, nLanc As Integer, nParc As Integer, nSeqOrigem As Integer, nSeq As Integer, nCompl As Integer, nCodigo_Destino As Long, nPerc As Double
Dim sObs As String

nCodigo_Origem = Val(txtCodigoOrigem.Text)
nCompl = 0
nTot = lvDestino.ListItems.Count

For nPos = 1 To lvDestino.ListItems.Count
    CallPb CLng(nPos), CLng(nTot)
    nAno = lvDestino.ListItems(nPos).Text
    nLanc = lvDestino.ListItems(nPos).SubItems(1)
    nSeqOrigem = lvDestino.ListItems(nPos).SubItems(2)
    nParc = lvDestino.ListItems(nPos).SubItems(3)
    nCompl = lvDestino.ListItems(nPos).SubItems(4)
    
    For nPos2 = 1 To UBound(aCodigos)
        nCodigo_Destino = aCodigos(nPos2).nCodigo
        nPerc = aCodigos(nPos2).nPerc
        nCol = 5 + nPos2
        nValor = lvDestino.ListItems(nPos).SubItems(nCol)
        
        'busca a seq disponivel no codigo de destino
        Sql = "select max(seqlancamento) as maximo from debitoparcela where codreduzido=" & nCodigo_Destino & " and anoexercicio=" & nAno & " and codlancamento=" & nLanc & " and numparcela=" & nParc
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux!maximo) Then
            nSeq = 0
        Else
            nSeq = RdoAux!maximo + 1
        End If
        RdoAux.Close
        
        'GRAVA PARCELA
        Sql = "INSERT DEBITOPARCELA SELECT " & nCodigo_Destino & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & ",0,3,DATAVENCIMENTO,DATADEBASE,CODMOEDA,"
        Sql = Sql & "NUMEROLIVRO,PAGINALIVRO,NUMCERTIDAO,DATAINSCRICAO,DATAAJUIZA,VALORJUROS,NUMPROCESSO,INTACTO,NOTIFICADO,NUMEXECFISCAL,ANOEXECFISCAL,PROCESSOCNJ,"
        Sql = Sql & "SIMPLESNACIONAL,PROTESTO_NRO_TITULO,PROTESTO_DATA_REMESSA,236 FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodigo_Origem & " AND ANOEXERCICIO=" & nAno & " AND "
        Sql = Sql & "CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeqOrigem & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nCompl
        cn.Execute Sql, rdExecDirect
        
       'GRAVA TRIBUTO
        Sql = "SELECT * FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & nCodigo_Origem & " AND ANOEXERCICIO=" & nAno & " AND "
        Sql = Sql & "CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeqOrigem & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nCompl
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
                Sql = Sql & nCodigo_Destino & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & ",0," & !CodTributo & "," & Virg2Ponto(!VALORTRIBUTO * nPerc / 100) & ")"
                cn.Execute Sql, rdExecDirect
               .MoveNext
            Loop
        End With
        
        'GRAVA OBSERVACAO
        sObs = "Débito proveniente da transferência do código de origem " & nCodigo_Origem & " no percentual de " & nPerc & "%."
        Sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & nCodigo_Destino & "," & nAno & ","
        Sql = Sql & nLanc & "," & nSeq & "," & nParc & "," & 0 & "," & 0 & ",'" & sObs & "'," & 236 & ",'" & Format(Now, "mm/dd/yyyy") & "')"
        cn.Execute Sql, rdExecDirect
        
        'ATUALIZA DEBITO ORIGEM
        Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=13 WHERE CODREDUZIDO=" & nCodigo_Origem & " AND ANOEXERCICIO=" & nAno & " AND "
        Sql = Sql & "CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeqOrigem & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nCompl
        cn.Execute Sql, rdExecDirect
        
    Next
Next

PBar.value = 100
Liberado
MsgBox "fim"

End Sub
