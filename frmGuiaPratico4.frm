VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmGuiaPratico4 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Certid�o de Isen��o de ITBI"
   ClientHeight    =   5460
   ClientLeft      =   5955
   ClientTop       =   3480
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9315
   Begin VB.ComboBox cmbAss 
      Height          =   315
      ItemData        =   "frmGuiaPratico4.frx":0000
      Left            =   4275
      List            =   "frmGuiaPratico4.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   5040
      Width           =   1905
   End
   Begin VB.TextBox txtCPF 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   19
      Top             =   405
      Width           =   2985
   End
   Begin VB.Frame frExp 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Natureza da Transa��o:"
      Height          =   735
      Left            =   135
      TabIndex        =   13
      Top             =   1125
      Width           =   9015
      Begin VB.TextBox txtExp 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000080&
         Height          =   420
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   270
         Width           =   8835
      End
      Begin MSComctlLib.TreeView tvMain 
         Height          =   3975
         Left            =   45
         TabIndex        =   2
         Top             =   270
         Visible         =   0   'False
         Width           =   8925
         _ExtentX        =   15743
         _ExtentY        =   7011
         _Version        =   393217
         Style           =   4
         Appearance      =   1
      End
      Begin prjChameleon.chameleonButton cmdExp 
         Height          =   240
         Left            =   2070
         TabIndex        =   1
         ToolTipText     =   "Exibir Lista"
         Top             =   0
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   423
         BTYPE           =   14
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
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   14869218
         BCOLO           =   14869218
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmGuiaPratico4.frx":0029
         PICN            =   "frmGuiaPratico4.frx":0045
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Observa��o"
      Height          =   960
      Index           =   2
      Left            =   135
      TabIndex        =   18
      Top             =   3960
      Width           =   9015
      Begin VB.TextBox txtObs 
         Appearance      =   0  'Flat
         Height          =   645
         Left            =   45
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   225
         Width           =   8880
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Im�vel"
      Height          =   960
      Index           =   1
      Left            =   135
      TabIndex        =   17
      Top             =   2970
      Width           =   9015
      Begin VB.TextBox txtDados 
         Appearance      =   0  'Flat
         Height          =   645
         Left            =   45
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   225
         Width           =   8880
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Transmitente"
      Height          =   960
      Index           =   0
      Left            =   135
      TabIndex        =   16
      Top             =   1980
      Width           =   9015
      Begin VB.TextBox txtTrans 
         Appearance      =   0  'Flat
         Height          =   645
         Left            =   45
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   225
         Width           =   8880
      End
   End
   Begin VB.TextBox txtValor 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1260
      TabIndex        =   6
      Top             =   5040
      Width           =   1455
   End
   Begin prjChameleon.chameleonButton cmdCnsImovel 
      Height          =   315
      Left            =   8685
      TabIndex        =   0
      ToolTipText     =   "Consulta Cidad�o"
      Top             =   90
      Width           =   465
      _ExtentX        =   820
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
      MICON           =   "frmGuiaPratico4.frx":019F
      PICN            =   "frmGuiaPratico4.frx":01BB
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   315
      Left            =   8055
      TabIndex        =   7
      ToolTipText     =   "Imprimir Requerimento"
      Top             =   5040
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
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
      MICON           =   "frmGuiaPratico4.frx":0315
      PICN            =   "frmGuiaPratico4.frx":0331
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
      Caption         =   "Assinatura.:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   2970
      TabIndex        =   20
      Top             =   5085
      Width           =   1380
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor...:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   225
      TabIndex        =   15
      Top             =   5085
      Width           =   1200
   End
   Begin VB.Label lblRequerente 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2700
      TabIndex        =   12
      Top             =   135
      Width           =   5835
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Requerente (Cidad�o)..:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   180
      TabIndex        =   11
      Top             =   135
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CPF/CNPJ..:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   180
      TabIndex        =   10
      Top             =   405
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Endere�o..:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   180
      TabIndex        =   9
      Top             =   720
      Width           =   1110
   End
   Begin VB.Label lblEndereco 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1395
      TabIndex        =   8
      Top             =   720
      Width           =   7815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   90
      X2              =   9225
      Y1              =   1035
      Y2              =   1035
   End
End
Attribute VB_Name = "frmGuiaPratico4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCnsImovel_Click()
Set frm = frmCnsCidadao
frm.sForm = Me.Name
frm.show
frm.ZOrder 0
End Sub

Private Sub cmdExp_Click()
Dim sTexto As String

If cmdExp.value = True Then
    frExp.Height = 4290
    frExp.ZOrder 0
    tvMain.Visible = True
    tvMain.Nodes(1).Selected = True
    txtExp.Visible = False
Else
    sTexto = tvMain.SelectedItem.Text
    If Left(tvMain.SelectedItem.Key, 3) = "ART" Then
        MsgBox "Selecione um Inciso.", vbExclamation, "Aten��o"
        sTexto = ""
    ElseIf Left(tvMain.SelectedItem.Key, 3) = "INC" And tvMain.SelectedItem.Children > 0 Then
        MsgBox "Selecione um tipo de Inciso.", vbExclamation, "Aten��o"
        sTexto = ""
    ElseIf Left(tvMain.SelectedItem.Key, 3) = "ITE" Then
        sTexto = tvMain.SelectedItem.Parent.Text & tvMain.SelectedItem.Text
    End If
    If sTexto <> "" Then
        If tvMain.SelectedItem.Tag <> "" Then
            sTexto = tvMain.Nodes(tvMain.SelectedItem.Tag).Text & " - " & sTexto
        End If
    End If
    
    frExp.Height = 735
    tvMain.Visible = False
    txtExp.Text = sTexto
    txtExp.Visible = True
End If


End Sub

Private Sub cmdPrint_Click()

Sql = "DELETE FROM REPORTTMP WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

Sql = "INSERT REPORTTMP(USUARIO,MEMO1,MEMO2,MEMO3) VALUES('" & NomeDeLogin & "','" & Mask(txtTrans.Text) & "','"
Sql = Sql & Mask(txtDados.Text) & "','" & Mask(txtObs.Text) & "')"
cn.Execute Sql, rdExecDirect

frmReport.ShowReport2 "GUIAPRATICO4", frmMdi.HWND, Me.HWND

Sql = "DELETE FROM REPORTTMP WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub Form_Activate()
If CodCidadao = 0 Then Exit Sub
If CodCidadao > 500000 Then
    Le
    CodCidadao = 0
Else
    MsgBox "C�digo de cidad�o inv�lido.", vbExclamation, "Aten��o"
    Limpa
End If
End Sub

Private Sub Form_Load()
Centraliza Me
cmbAss.ListIndex = 0
Init
End Sub

Private Sub Le()
Dim Sql As String, RdoAux As rdoResultset, sCidade As String
Sql = "SELECT * FROM vwFULLCIDADAO WHERE CODCIDADAO=" & CodCidadao
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    Do Until .EOF
        txtCPF.Text = SubNull(!CPF)
        If txtCPF.Text <> "" Then
            txtCPF.Text = Format(RdoAux!CPF, "00#\.###\.###-##")
        End If
        If txtCPF.Text = "" Then
            If Not IsNull(!Cnpj) Then
                txtCPF.Text = Format(!Cnpj, "0#\.###\.###/####-##")
            End If
        End If
        sCidade = SubNull(!descCidade)
'        If sCidade = "" Then
'            sCidade = SubNull(!descCidade)
 '       End If
'        If sCidade = "" Then
 '           sCidade = SubNull(!descCidade)
'        End If
        lblEndereco.Caption = SubNull(!Endereco) & ", " & SubNull(!NUMIMOVEL) & " " & SubNull(!Complemento) & " " & sCidade & "\" & SubNull(!SiglaUF)
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub Limpa()
'lblCodRequerente.Caption = ""
lblRequerente.Caption = ""
txtCPF.Text = ""
lblEndereco.Caption = ""
txtExp.Text = ""
End Sub

Sub Init()
Dim NodX As Object, x As Integer

Set NodX = tvMain.Nodes.Add(, , "ART111", "Artigo 111")
Set NodX = tvMain.Nodes.Add("ART111", tvwChild, "INC11101", "Inciso I - Substabelecimento de mandato em causa pr�pria, ou com poderes equivalentes, feito para o mandat�rio receber a escritura definitiva do im�vel.")
Set NodX = tvMain.Nodes.Add("ART111", tvwChild, "INC11102", "Inciso II - Transmiss�o de bem im�vel, por for�a de ")
Set NodX = tvMain.Nodes.Add("INC11102", tvwChild, "ITE1110201", "retrovenda")
Set NodX = tvMain.Nodes.Add("INC11102", tvwChild, "ITE1110202", "retrocess�o")
Set NodX = tvMain.Nodes.Add("INC11102", tvwChild, "ITE1110203", "pacto de melhor comprador")
Set NodX = tvMain.Nodes.Add("ART111", tvwChild, "INC11103", "Inciso III - Aquisi��o de im�vel ")
Set NodX = tvMain.Nodes.Add("INC11103", tvwChild, "ITE1110301", "pela Uni�o")
Set NodX = tvMain.Nodes.Add("INC11103", tvwChild, "ITE1110302", "pelo Estado")
Set NodX = tvMain.Nodes.Add("INC11103", tvwChild, "ITE1110303", "pelo Munic�pio")
Set NodX = tvMain.Nodes.Add("INC11103", tvwChild, "ITE1110304", "pelo Distrito Federal")
Set NodX = tvMain.Nodes.Add("INC11103", tvwChild, "ITE1110305", "por autarquia institu�da e mantida pelo Poder P�blico")
Set NodX = tvMain.Nodes.Add("INC11103", tvwChild, "ITE1110306", "por funda��es institu�da e mantida pelo Poder P�blico")
Set NodX = tvMain.Nodes.Add("ART111", tvwChild, "INC11104", "Inciso IV - Aquisi��o de im�vel por ")
Set NodX = tvMain.Nodes.Add("INC11104", tvwChild, "ITE1110401", "partido pol�tico")
Set NodX = tvMain.Nodes.Add("INC11104", tvwChild, "ITE1110402", "funda��o de partido pol�tico")
Set NodX = tvMain.Nodes.Add("INC11104", tvwChild, "ITE1110403", "entidade sindical")
Set NodX = tvMain.Nodes.Add("INC11104", tvwChild, "ITE1110404", "institui��o de educa��o e assist�ncia social sem fins lucrtaivos")
Set NodX = tvMain.Nodes.Add("ART111", tvwChild, "INC11105", "Inciso V - Incorpora��o ao patrim�nio de pessoa jur�dica em realiza��o de capital.")
Set NodX = tvMain.Nodes.Add("ART111", tvwChild, "INC11106", "Inciso VI - ")
Set NodX = tvMain.Nodes.Add("INC11106", tvwChild, "ITE1110601", "fus�o")
Set NodX = tvMain.Nodes.Add("INC11106", tvwChild, "ITE1110602", "incorpora��o")
Set NodX = tvMain.Nodes.Add("INC11106", tvwChild, "ITE1110603", "cis�o")
Set NodX = tvMain.Nodes.Add("INC11106", tvwChild, "ITE1110604", "extin��o de pessoa jur�dica")
Set NodX = tvMain.Nodes.Add("ART111", tvwChild, "INC11107", "Inciso VII - Transmiss�o de im�vel ")
Set NodX = tvMain.Nodes.Add("INC11107", tvwChild, "ITE1110701", "desapropriado para fins de reforma agr�ria")
Set NodX = tvMain.Nodes.Add("INC11107", tvwChild, "ITE1110702", "cedido para fins de reforma agr�ria")
Set NodX = tvMain.Nodes.Add("INC11107", tvwChild, "ITE1110703", "doado para fins de reforma agr�ria")
Set NodX = tvMain.Nodes.Add("ART111", tvwChild, "INC11108", "Inciso VIII - Aquisi��es imobili�rias para fins residenciais, oriundos de programas e conv�nios com o munic�pio, para constru��o de habita��es populares destinadas � fam�lias de baixa renda, pelo sistema de mutir�o e autoconstru��o.")
Set NodX = tvMain.Nodes.Add(, , "ART110", "Artigo 110")
Set NodX = tvMain.Nodes.Add("ART110", tvwChild, "INC11001", "Inciso V - Divis�o amig�vel por extins�o de condom�nio de cota parte ideal.")
Set NodX = tvMain.Nodes.Add(, , "ART150", "Artigo 150")
Set NodX = tvMain.Nodes.Add("ART150", tvwChild, "INC15001", "Inciso VI, letra ""b"" da Constitui��o Federal da Rep�blica Federativa do Brasil, combinada com o Artigo 111, inciso IV, da Lei Complementar 07/92.")
Set NodX = tvMain.Nodes.Add(, , "ART150b", "Artigo 150b")
Set NodX = tvMain.Nodes.Add("ART150b", tvwChild, "INC15001b", "Inciso VI, letra ""b"" da Constitui��o Federal da Rep�blica Federativa do Brasil.")
Set NodX = tvMain.Nodes.Add(, , "ART003", "Artigo 3�")
Set NodX = tvMain.Nodes.Add("ART003", tvwChild, "INC2ART3", " Artigo 3�,Inciso II, da Lei Compelementar 107/2009, referente a aquisi��o de im�vel pela Caixa Econ�mica Federal quando da contrata��o dos empreencimentos habitacionais.")

With tvMain
    For x = 1 To .Nodes.Count
        If Left(.Nodes(x).Key, 6) = "INC111" Or Left(.Nodes(x).Key, 6) = "ITE111" Then
            .Nodes(x).Tag = "ART111"
        End If
        If Left(.Nodes(x).Key, 6) = "INC110" Then
            .Nodes(x).Tag = "ART110"
        ElseIf Left(.Nodes(x).Key, 7) = "ART150b" Then
            .Nodes(x).Tag = "ART150b"
        ElseIf Left(.Nodes(x).Key, 6) = "INC150" Then
            .Nodes(x).Tag = "ART150"
        ElseIf Left(.Nodes(x).Key, 6) = "ART003" Then
            .Nodes(x).Tag = "ART003"
        End If
        
        If Left(.Nodes(x).Key, 3) = "ART" Then
            .Nodes(x).Bold = True
            .Nodes(x).ForeColor = vbBlue
        ElseIf Left(.Nodes(x).Key, 3) = "INC" Then
            .Nodes(x).ForeColor = VerdeAccess
        ElseIf Left(.Nodes(x).Key, 3) = "ITE" Then
            .Nodes(x).Bold = False
            .Nodes(x).ForeColor = vbBlack
       End If
        .Nodes(x).EnsureVisible
    Next
    .Nodes(1).EnsureVisible
End With


End Sub

Private Sub txtCPF_KeyPress(KeyAscii As Integer)
Tweak txtCPF, KeyAscii, DecimalPositive
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
'Tweak txtValor, KeyAscii, DecimalPositive
End Sub
