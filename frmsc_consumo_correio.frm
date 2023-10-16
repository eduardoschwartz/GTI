VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form frmsc_consumo_correio 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consumo de Taxas do Correio"
   ClientHeight    =   4770
   ClientLeft      =   7290
   ClientTop       =   7290
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   6780
   Begin VB.TextBox txtUnidade 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   90
      Width           =   5595
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   4290
      Left            =   45
      TabIndex        =   0
      Top             =   405
      Width           =   4470
      Begin VB.TextBox txtEmpenho 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1485
         TabIndex        =   5
         Top             =   1035
         Width           =   1365
      End
      Begin VB.ComboBox cmbMes 
         Height          =   315
         Left            =   1485
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   630
         Width           =   915
      End
      Begin VB.TextBox txtValor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1485
         MaxLength       =   15
         TabIndex        =   3
         Text            =   "0,00"
         Top             =   1755
         Width           =   1365
      End
      Begin VB.ComboBox cmbAno 
         Height          =   315
         Left            =   2475
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   630
         Width           =   915
      End
      Begin VB.ComboBox cmbMesAno 
         Height          =   315
         Left            =   1485
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   630
         Width           =   1455
      End
      Begin esMaskEdit.esMaskedEdit mskVencto 
         Height          =   285
         Left            =   1485
         TabIndex        =   6
         Top             =   1395
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         MouseIcon       =   "frmsc_consumo_correio.frx":0000
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
      Begin prjChameleon.chameleonButton cmdNovo 
         Height          =   360
         Left            =   2565
         TabIndex        =   7
         ToolTipText     =   "Novo Registro"
         Top             =   3690
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   635
         BTYPE           =   7
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
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmsc_consumo_correio.frx":001C
         PICN            =   "frmsc_consumo_correio.frx":0038
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
         Height          =   360
         Left            =   3465
         TabIndex        =   8
         ToolTipText     =   "Excluir Registro"
         Top             =   3690
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   635
         BTYPE           =   7
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
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmsc_consumo_correio.frx":0192
         PICN            =   "frmsc_consumo_correio.frx":01AE
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
         Height          =   360
         Left            =   3015
         TabIndex        =   9
         ToolTipText     =   "Editar Registro"
         Top             =   3690
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   635
         BTYPE           =   7
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
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmsc_consumo_correio.frx":0250
         PICN            =   "frmsc_consumo_correio.frx":026C
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
         Height          =   360
         Left            =   3015
         TabIndex        =   10
         ToolTipText     =   "Gravar os Dados"
         Top             =   3690
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   635
         BTYPE           =   7
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
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmsc_consumo_correio.frx":03C6
         PICN            =   "frmsc_consumo_correio.frx":03E2
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
         Height          =   360
         Left            =   3465
         TabIndex        =   11
         ToolTipText     =   "Cancelar Edição"
         Top             =   3690
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   635
         BTYPE           =   7
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
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmsc_consumo_correio.frx":0787
         PICN            =   "frmsc_consumo_correio.frx":07A3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Empenho.....:"
         Height          =   195
         Index           =   17
         Left            =   225
         TabIndex        =   19
         Top             =   1065
         Width           =   1185
      End
      Begin VB.Label lblMediaAno 
         BackStyle       =   0  'Transparent
         Caption         =   "00000"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1485
         TabIndex        =   18
         Top             =   2160
         Width           =   1140
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor médio......:"
         Height          =   195
         Index           =   15
         Left            =   225
         TabIndex        =   17
         Top             =   2160
         Width           =   1185
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor em reais...:"
         Height          =   195
         Index           =   12
         Left            =   225
         TabIndex        =   16
         Top             =   1800
         Width           =   1275
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Vencto.....:"
         Height          =   195
         Index           =   9
         Left            =   225
         TabIndex        =   15
         Top             =   1440
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Mês/Ano cons..:"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   14
         Top             =   675
         Width           =   1185
      End
      Begin VB.Label lblCod 
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1485
         TabIndex        =   13
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Código..............:"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   12
         Top             =   315
         Width           =   1185
      End
   End
   Begin Tributacao.jcFrames jcFrames1 
      Height          =   4290
      Left            =   4545
      Top             =   405
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   7567
      FrameColor      =   12829635
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Consumo Anterior"
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
      Begin vbAcceleratorSGrid6.vbalGrid grdMain 
         Height          =   3930
         Left            =   90
         TabIndex        =   21
         Top             =   270
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   6932
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         BackColor       =   16777215
         NoFocusHighlightForeColor=   16777215
         NoFocusHighlightBackColor=   128
         GroupRowBackColor=   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeaderButtons   =   0   'False
         HeaderDragReorderColumns=   0   'False
         HeaderHotTrack  =   0   'False
         HeaderFlat      =   -1  'True
         BorderStyle     =   0
         ScrollBarStyle  =   1
         Editable        =   -1  'True
         DisableIcons    =   -1  'True
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Secretaria..:"
      Height          =   195
      Index           =   4
      Left            =   90
      TabIndex        =   22
      Top             =   135
      Width           =   915
   End
End
Attribute VB_Name = "frmsc_consumo_correio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Evento As String, MesAno As String

Private Sub cmbMesAno_Click()
If cmbMesAno.ListIndex > -1 Then CarregaDados
End Sub

Private Sub cmdAlterar_Click()
Evento = "ALTERAR"
Eventos "INCLUIR"

End Sub

Private Sub cmdCancel_Click()
Eventos "INICIAR"
End Sub

Private Sub cmdExcluir_Click()
Dim Sql As String, RdoAux As rdoResultset, nCodigo As Integer, nAno As Integer, nMes As Integer

If cmbMesAno.ListIndex = -1 Then
    MsgBox "Selecione um período para excluir!", vbCritical, "Erro"
    Exit Sub
End If

nCodigo = Val(lblCod.Caption)
nAno = Val(Right(cmbMesAno.Text, 4))
nMes = Val(Left(cmbMesAno.Text, 2))

If MsgBox("Deseja excluir este período?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmação") = vbYes Then
    Sql = "delete from sc_taxa_correio_consumo where codigo=" & nCodigo & " and ano=" & nAno & " and mes=" & nMes
    cn.Execute Sql, rdExecDirect
    LoadPeriodo
    CarregaDados
End If

End Sub

Private Sub cmdGravar_Click()
Dim Sql As String, RdoAux As rdoResultset, nCodigo As Integer, nAno As Integer, nMes As Integer

nCodigo = Val(lblCod.Caption)
nAno = Val(cmbAno.Text)
nMes = Val(cmbMes.Text)
If txtValor.Text = "" Then txtValor.Text = "0"

If Evento = "ALTERAR" Then
    Sql = "update sc_taxa_correio_consumo set empenho=" & sNull(txtEmpenho.Text) & ","
    Sql = Sql & "valor=" & Virg2Ponto(txtValor.Text) & ",datavencimento=" & sNullData(mskVencto.Text) & " where codigo=" & nCodigo & " and ano=" & nAno & " and mes=" & nMes
    cn.Execute Sql, rdExecDirect
Else
    Sql = "select * from sc_taxa_correio_consumo where codigo=" & nCodigo & " and ano=" & Val(cmbAno.Text) & " and mes=" & nMes
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount > 0 Then
        MsgBox "Periodo de consumo já cadastrado para esta unidade. Clique em Editar para alterar os dados.", vbCritical, "Erro"
        Exit Sub
    End If
    RdoAux.Close
    
    Sql = "insert sc_taxa_correio_consumo(codigo,ano,mes,valor,datavencimento,empenho) values(" & Val(lblCod.Caption) & ","
    Sql = Sql & nAno & "," & nMes & "," & Virg2Ponto(txtValor.Text) & "," & sNullData(mskVencto.Text) & "," & sNull(txtEmpenho.Text) & ")"
    cn.Execute Sql, rdExecDirect
End If

LoadPeriodo
CarregaDados
Eventos "INICIAR"
End Sub

Private Sub cmdNovo_Click()
Dim x As Integer

Evento = "NOVO"
Eventos "INCLUIR"
Limpa

For x = 0 To cmbMes.ListCount - 1
    If Val(cmbMes.List(x)) = Month(Now) Then
        cmbMes.ListIndex = x
        Exit For
    End If
Next

For x = 0 To cmbAno.ListCount - 1
    If Val(cmbAno.List(x)) = Year(Now) Then
        cmbAno.ListIndex = x
        Exit For
    End If
Next


End Sub

Private Sub Form_Load()

Me.Top = frmsc_taxa_correio.Top + 1000
Me.Left = frmsc_taxa_correio.Left + 2000
Header
Eventos "INICIAR"
LoadPeriodo

End Sub

Private Sub Header()
Dim x As Integer
Limpa
GridHeader
lblCod.Caption = frmsc_taxa_correio.lblCod.Caption
txtUnidade.Text = frmsc_taxa_correio.txtNome.Text

For x = 1 To 12
    cmbMes.AddItem Format(x, "00")
Next

For x = 2021 To Year(Now)
    cmbAno.AddItem x
Next

Eventos "INICIAR"

End Sub

Private Sub GridHeader()
With grdMain
    .Clear
    .GridFillLineColor = vbWhite
    .Editable = False
    .GridLines = True
    .HighlightBackColor = Marrom
    .HighlightForeColor = vbWhite
    .RowMode = True
    .DefaultRowHeight = 17
    .AddColumn "kCod", "Cód", ecgHdrTextALignCentre, , 40, False
    .AddColumn "kPer", "Periodo", ecgHdrTextALignCentre, , 60
    .AddColumn "kCon", "Valor", ecgHdrTextALignRight, , 60
End With

End Sub


Private Sub txtValor_KeyPress(KeyAscii As Integer)
Tweak txtValor, KeyAscii, DecimalPositive
End Sub

Private Sub Eventos(Tipo As String)
Dim cor As Long

cor = Me.BackColor
If Tipo = "INICIAR" Then
    cmdNovo.Visible = True
    cmdAlterar.Visible = True
    cmdExcluir.Visible = True
    cmdGravar.Visible = False
    cmdCancel.Visible = False
    txtEmpenho.Locked = True
    txtEmpenho.BackColor = cor
    txtValor.Locked = True
    txtValor.BackColor = cor
    mskVencto.Locked = True
    mskVencto.BackColor = cor
    cmbMes.Visible = False
    cmbAno.Visible = False
    cmbMesAno.Visible = True
    cmbMesAno.Enabled = True
    grdMain.Enabled = True
ElseIf Tipo = "INCLUIR" Then
    cmdNovo.Visible = False
    cmdAlterar.Visible = False
    cmdExcluir.Visible = False
    txtEmpenho.Locked = False
    txtEmpenho.BackColor = Branco
    txtValor.Locked = False
    txtValor.BackColor = Branco
    mskVencto.Locked = False
    mskVencto.BackColor = Branco
    cmdGravar.Visible = True
    cmdCancel.Visible = True
    grdMain.Enabled = False
    If Evento = "ALTERAR" Then
        cmbMesAno.Visible = True
        cmbMesAno.Enabled = False
        cmbMes.Visible = False
        cmbAno.Visible = False
    Else
        cmbMes.Visible = True
        cmbAno.Visible = True
        cmbMesAno.Visible = False
    End If
End If
   


End Sub

Private Sub Limpa()
txtEmpenho.Text = ""
txtValor.Text = ""
LimpaMascara mskVencto
lblMediaAno.Caption = "00000"
End Sub

Private Sub CarregaDados()

Dim Sql As String, RdoAux As rdoResultset, nAno As Integer, nMes As Integer, nCodigo As Integer, sRede As String, sPeriodo As String
Dim nValorMes  As Double, nCount As Integer, nMedia As Double

nCodigo = Val(lblCod.Caption)
nAno = Val(Right(cmbMesAno.Text, 4))
nMes = Val(Left(cmbMesAno.Text, 2))
MesAno = cmbMes.Text

Sql = "SELECT * from sc_taxa_correio_consumo Where codigo = " & nCodigo & " AND ano = " & nAno & " AND mes = " & nMes
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        If Not IsNull(!DataVencimento) Then
            mskVencto.Text = Format(!DataVencimento, "dd/mm/yyyy")
        End If
        txtValor.Text = Format(!valor, "#0.00")
        txtEmpenho.Text = SubNull(!empenho)
        If txtValor.Text = "" Then txtValor.Text = "0"
        nValorMes = CDbl(txtValor.Text)
    End If
    nCount = 1
   .Close
End With


grdMain.Clear
Sql = "SELECT codigo,ano,mes,valor FROM sc_taxa_correio_consumo WHERE codigo=" & nCodigo & " ORDER BY ano DESC, mes desc"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If !ano > nAno Then
            GoTo Proximo
        ElseIf !ano = nAno Then
            If !mes >= nMes Then
                GoTo Proximo
            End If
        End If
        sPeriodo = Format(!mes, "00") & "/" & CStr(!ano)
        grdMain.AddRow
        grdMain.CellDetails grdMain.Rows, 1, !codigo
        grdMain.CellDetails grdMain.Rows, 2, sPeriodo, DT_CENTER
        grdMain.CellDetails grdMain.Rows, 3, Format(!valor, "#0.00"), DT_RIGHT
        nValorMes = CDbl(!valor)
        nCount = nCount + 1

Proximo:
       .MoveNext
    Loop
   .Close
End With
 
If nCount > 0 Then
    nMedia = nValorMes / nCount
End If
lblMediaAno.Caption = Format(nMedia, "#0.00")
 
End Sub

Private Sub LoadPeriodo()
Dim Sql As String, RdoAux As rdoResultset, x As Integer

cmbMesAno.Clear
Sql = "SELECT DISTINCT mes, ano FROM sc_taxa_correio_consumo WHERE codigo=" & Val(lblCod.Caption) & " ORDER BY ano DESC, mes desc"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbMesAno.AddItem Format(!mes, "00") & "/" & CStr(!ano)
       .MoveNext
    Loop
   .Close
End With

If cmbMesAno.ListCount > 0 Then
    cmbMesAno.ListIndex = 0
End If

End Sub



