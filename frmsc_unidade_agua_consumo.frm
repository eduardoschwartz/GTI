VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form frmsc_unidade_agua_consumo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consumo de água da unidade"
   ClientHeight    =   4755
   ClientLeft      =   16230
   ClientTop       =   5145
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6780
   Begin VB.TextBox txtUnidade 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   900
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   90
      Width           =   5775
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
         TabIndex        =   38
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
         Left            =   945
         TabIndex        =   36
         Text            =   "10293/2021"
         Top             =   2790
         Width           =   1365
      End
      Begin VB.ComboBox cmbMes 
         Height          =   315
         Left            =   945
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   180
         Width           =   1140
      End
      Begin prjChameleon.chameleonButton cmdMsg 
         Height          =   300
         Left            =   90
         TabIndex        =   37
         ToolTipText     =   "Mensagens recebidas"
         Top             =   3240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "Mensagens"
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
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   16711935
         MPTR            =   1
         MICON           =   "frmsc_unidade_agua_consumo.frx":0000
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
         Caption         =   "Empenho:"
         Height          =   195
         Index           =   17
         Left            =   90
         TabIndex        =   35
         Top             =   2835
         Width           =   735
      End
      Begin VB.Label lblRede 
         BackStyle       =   0  'Transparent
         Caption         =   "Água e Esgoto"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2835
         TabIndex        =   34
         Top             =   900
         Width           =   1140
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Rede:"
         Height          =   195
         Index           =   16
         Left            =   2295
         TabIndex        =   33
         Top             =   900
         Width           =   510
      End
      Begin VB.Label lblValor 
         BackStyle       =   0  'Transparent
         Caption         =   "0.000.000,00"
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
         Height          =   195
         Left            =   3150
         TabIndex        =   32
         Top             =   3870
         Width           =   1140
      End
      Begin VB.Label lblDias 
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2790
         TabIndex        =   31
         Top             =   2160
         Width           =   285
      End
      Begin VB.Label lblConsumo 
         BackStyle       =   0  'Transparent
         Caption         =   "00000"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1260
         TabIndex        =   30
         Top             =   2160
         Width           =   510
      End
      Begin VB.Label lblMediaAno 
         BackStyle       =   0  'Transparent
         Caption         =   "00000"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3195
         TabIndex        =   29
         Top             =   1215
         Width           =   510
      End
      Begin VB.Label lblMediaMes 
         BackStyle       =   0  'Transparent
         Caption         =   "00000"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1260
         TabIndex        =   28
         Top             =   1215
         Width           =   510
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Média ano.:"
         Height          =   195
         Index           =   15
         Left            =   2295
         TabIndex        =   27
         Top             =   1215
         Width           =   915
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Média mês......:"
         Height          =   195
         Index           =   14
         Left            =   90
         TabIndex        =   26
         Top             =   1215
         Width           =   1140
      End
      Begin VB.Label lblDataCorte 
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1260
         TabIndex        =   25
         Top             =   2475
         Width           =   960
      End
      Begin VB.Label lblDataVencimento 
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000"
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
         Height          =   195
         Left            =   1440
         TabIndex        =   24
         Top             =   3870
         Width           =   1095
      End
      Begin VB.Label lblDataLeituraAnterior 
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2790
         TabIndex        =   23
         Top             =   1845
         Width           =   915
      End
      Begin VB.Label lblDataLeituraAtual 
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2790
         TabIndex        =   22
         Top             =   1530
         Width           =   960
      End
      Begin VB.Label lblLeituraAnterior 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1260
         TabIndex        =   21
         Top             =   1845
         Width           =   555
      End
      Begin VB.Label lblLeituraAtual 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1260
         TabIndex        =   20
         Top             =   1530
         Width           =   555
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Data de Corte..:"
         Height          =   195
         Index           =   13
         Left            =   90
         TabIndex        =   19
         Top             =   2475
         Width           =   1140
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor.:"
         Height          =   195
         Index           =   12
         Left            =   2655
         TabIndex        =   18
         Top             =   3870
         Width           =   555
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Dias.:"
         Height          =   195
         Index           =   11
         Left            =   2295
         TabIndex        =   17
         Top             =   2160
         Width           =   420
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Consumo m³....:"
         Height          =   195
         Index           =   10
         Left            =   90
         TabIndex        =   16
         Top             =   2160
         Width           =   1140
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Vencimento:"
         Height          =   195
         Index           =   9
         Left            =   90
         TabIndex        =   15
         Top             =   3870
         Width           =   1275
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
         Height          =   195
         Index           =   8
         Left            =   2295
         TabIndex        =   14
         Top             =   1845
         Width           =   420
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
         Height          =   195
         Index           =   6
         Left            =   2295
         TabIndex        =   13
         Top             =   1530
         Width           =   555
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Leitura anterior:"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   12
         Top             =   1845
         Width           =   1140
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Leitura atual....:"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   11
         Top             =   1530
         Width           =   1095
      End
      Begin VB.Label lblSituacao 
         BackStyle       =   0  'Transparent
         Caption         =   "Religada"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1260
         TabIndex        =   10
         Top             =   900
         Width           =   780
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Situação.........:"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   9
         Top             =   900
         Width           =   1140
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Mês/Ano.:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   7
         Top             =   225
         Width           =   825
      End
      Begin VB.Label lblHidrometro 
         BackStyle       =   0  'Transparent
         Caption         =   "0000000"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3150
         TabIndex        =   6
         Top             =   585
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Hidrômetro:"
         Height          =   195
         Index           =   7
         Left            =   2295
         TabIndex        =   5
         Top             =   585
         Width           =   825
      End
      Begin VB.Label lblLigacao 
         BackStyle       =   0  'Transparent
         Caption         =   "0000000"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1260
         TabIndex        =   4
         Top             =   585
         Width           =   780
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Ligação.....:"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   3
         Top             =   585
         Width           =   1095
      End
      Begin VB.Label lblCod 
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2700
         TabIndex        =   2
         Top             =   225
         Width           =   510
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cód:"
         Height          =   195
         Index           =   0
         Left            =   2295
         TabIndex        =   1
         Top             =   225
         Width           =   375
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Unidade..:"
      Height          =   195
      Index           =   4
      Left            =   45
      TabIndex        =   40
      Top             =   135
      Width           =   825
   End
End
Attribute VB_Name = "frmsc_unidade_agua_consumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sMsg As String, MesAno As String, bDirty As Boolean

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmbMes_Click()
If cmbMes.ListIndex > -1 Then CarregaDados
End Sub

Private Sub cmdMsg_Click()
MsgBox sMsg, vbInformation, "Mensagens"
End Sub

Private Sub Form_Load()
Me.Top = frmsc_unidade_agua.Top + 1000
Me.Left = frmsc_unidade_agua.Left + 2000
MesAno = ""
Header
End Sub

Private Sub Header()
Dim Sql As String, RdoAux As rdoResultset

GridHeader
lblCod.Caption = frmsc_unidade_agua.lblCod.Caption
lblLigacao.Caption = frmsc_unidade_agua.lblLigacao.Caption
txtUnidade.Text = frmsc_unidade_agua.lblDescricao.Caption

Sql = "SELECT DISTINCT mes, ano FROM sc_ligacao_agua_resumo WHERE codigo=" & Val(lblCod.Caption) & " ORDER BY ano DESC, mes desc"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbMes.AddItem Format(!mes, "00") & "/" & CStr(!ano)
       .MoveNext
    Loop
   .Close
End With

cmbMes.ListIndex = 0

End Sub

Private Sub CarregaDados()

Dim Sql As String, RdoAux As rdoResultset, nAno As Integer, nMes As Integer, nCodigo As Integer, sRede As String, sPeriodo As String
If bDirty Then
   Save
End If

sMsg = ""
nCodigo = Val(lblCod.Caption)
nAno = Val(Right(cmbMes.Text, 4))
nMes = Val(Left(cmbMes.Text, 2))
MesAno = cmbMes.Text

Sql = "SELECT sc_ligacao_agua_resumo.codigo , sc_ligacao_agua_resumo.hidrometro, sc_ligacao_agua_resumo.Situacao, sc_ligacao_agua_resumo.data_leitura, sc_ligacao_agua_resumo.Data_Vencimento, sc_ligacao_agua_resumo.data_corte, sc_ligacao_agua_resumo.data_leitura_anterior, "
Sql = Sql & "sc_ligacao_agua_resumo.ano,sc_ligacao_agua_resumo.mes,sc_ligacao_agua_resumo.mensagem1,sc_ligacao_agua_resumo.mensagem2,sc_ligacao_agua_resumo.mensagem3,sc_ligacao_agua_resumo.mensagem4,sc_ligacao_agua_resumo.mensagem5,sc_ligacao_agua_resumo.mensagem6,"
Sql = Sql & "sc_ligacao_agua_resumo.agua,sc_ligacao_agua_resumo.esgoto,sc_ligacao_agua_consumo.leitura,sc_ligacao_agua_consumo.leitura_calc,sc_ligacao_agua_consumo.leitura_anterior,sc_ligacao_agua_consumo.media_mes ,sc_ligacao_agua_consumo.media_ano,"
Sql = Sql & "sc_ligacao_agua_consumo.consumo_medido,sc_ligacao_agua_consumo.consumo_calc,sc_ligacao_agua_consumo.valor,sc_ligacao_agua_consumo.dias,sc_liagacao_agua_status.nome AS descsituacao,sc_ligacao_agua_resumo.empenho From sc_ligacao_agua_resumo INNER JOIN sc_ligacao_agua_consumo ON sc_ligacao_agua_resumo.codigo = sc_ligacao_agua_consumo.codigo AND "
Sql = Sql & "sc_ligacao_agua_resumo.ano = sc_ligacao_agua_consumo.ano AND sc_ligacao_agua_resumo.mes = sc_ligacao_agua_consumo.mes INNER JOIN sc_liagacao_agua_status ON sc_ligacao_agua_resumo.situacao = sc_liagacao_agua_status.codigo Where sc_ligacao_agua_resumo.codigo = " & nCodigo & " AND sc_ligacao_agua_resumo.ano = " & nAno & " AND sc_ligacao_agua_resumo.mes = " & nMes
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    lblHidrometro.Caption = SubNull(!hidrometro)
    lblMediaMes.Caption = SubNull(!media_mes)
    lblMediaAno.Caption = SubNull(!media_ano)
    If !agua = True Then
        sRede = "Água"
    End If
    If !esgoto = True Then
        If sRede = "" Then
            sRede = "Esgoto"
        Else
            sRede = sRede & " e Esgoto"
        End If
    End If
    lblRede.Caption = sRede
    lblSituacao.Caption = !DescSituacao
    lblLeituraAtual.Caption = Format(Val(SubNull(!leitura)), "000000")
    lblLeituraAnterior.Caption = Format(Val(SubNull(!leitura_anterior)), "000000")
    lblDataLeituraAtual.Caption = IIf(IsNull(!data_leitura), "N/D", Format(!data_leitura, "dd/mm/yyyy"))
    lblDataLeituraAnterior.Caption = IIf(IsNull(!data_leitura_anterior), "N/D", Format(!data_leitura_anterior, "dd/mm/yyyy"))
    lblDataVencimento.Caption = IIf(IsNull(!Data_Vencimento), "N/D", Format(!Data_Vencimento, "dd/mm/yyyy"))
    lblDataCorte.Caption = IIf(IsNull(!data_corte), "N/D", Format(!data_corte, "dd/mm/yyyy"))
    lblConsumo.Caption = Format(!consumo_calc, "00000")
    lblDias.Caption = Format(!dias, "00")
    lblValor.Caption = Format(!valor, "#0.00")
    txtEmpenho.Text = SubNull(!empenho)
    
    If Trim(SubNull(!mensagem1)) <> "" Then
        sMsg = !mensagem1 & ", "
    End If
    If Trim(SubNull(!mensagem2)) <> "" Then
        sMsg = sMsg & !mensagem2 & ", "
    End If
    If Trim(SubNull(!mensagem3)) <> "" Then
        sMsg = sMsg & !mensagem3 & ", "
    End If
    If Trim(SubNull(!mensagem4)) <> "" Then
        sMsg = sMsg & !mensagem4 & ", "
    End If
    If Trim(SubNull(!mensagem5)) <> "" Then
        sMsg = sMsg & !mensagem5 & ", "
    End If
    If Trim(SubNull(!mensagem6)) <> "" Then
        sMsg = sMsg & !mensagem6 & ", "
    End If
    
    If Right(sMsg, 2) = ", " Then
        sMsg = Left(sMsg, Len(sMsg) - 2)
    End If
    
    If sMsg = "" Then
        cmdMsg.Enabled = False
    Else
        cmdMsg.Enabled = True
    End If
    
   .Close
End With

grdMain.Clear
Sql = "SELECT codigo,ano,mes,consumo_calc FROM sc_ligacao_agua_consumo WHERE codigo=" & nCodigo & " ORDER BY ano DESC, mes desc"
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
        grdMain.CellDetails grdMain.Rows, 3, SubNull(!consumo_calc), DT_RIGHT
        
Proximo:
       .MoveNext
    Loop
   .Close
End With

bDirty = False
 
End Sub

Private Sub Save()
Dim Sql As String, nAno As Integer, nMes As Integer
nAno = Val(Right(MesAno, 4))
nMes = Val(Left(MesAno, 2))

Sql = "update sc_ligacao_agua_resumo set empenho=" & sNull(txtEmpenho.Text) & " where codigo=" & Val(lblCod.Caption) & " and ano=" & nAno & " and mes=" & nMes
cn.Execute Sql, rdExecDirect
bDirty = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If bDirty Then
    Save
End If
End Sub

Private Sub txtEmpenho_Change()
bDirty = True
End Sub

Private Sub GridHeader()
With grdMain
    .GridFillLineColor = vbWhite
    .Editable = False
    .GridLines = True
    .HighlightBackColor = Marrom
    .HighlightForeColor = vbWhite
    .RowMode = True
    .DefaultRowHeight = 17
    .AddColumn "kCod", "Cód", ecgHdrTextALignCentre, , 40, False
    .AddColumn "kPer", "Periodo", ecgHdrTextALignCentre, , 60
    .AddColumn "kCon", "Cons.", ecgHdrTextALignRight, , 40
End With

End Sub

