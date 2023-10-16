VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmRelatorioIsencao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Isenção"
   ClientHeight    =   1170
   ClientLeft      =   8895
   ClientTop       =   4275
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   5595
   Begin VB.ComboBox cmbAno 
      Height          =   315
      Left            =   1665
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   1275
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   345
      Left            =   3825
      TabIndex        =   2
      ToolTipText     =   "Imprimir o relatório"
      Top             =   315
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   609
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
      MICON           =   "frmRelatorioIsencao.frx":0000
      PICN            =   "frmRelatorioIsencao.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ProgressBar Pb 
      Height          =   165
      Left            =   180
      TabIndex        =   3
      Top             =   855
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   291
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label lblPB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100 %"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2340
      TabIndex        =   4
      Top             =   855
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Ano de Isenção:"
      Height          =   240
      Left            =   225
      TabIndex        =   0
      Top             =   405
      Width           =   1320
   End
End
Attribute VB_Name = "frmRelatorioIsencao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim Sql As String, RdoAux As rdoResultset
Centraliza Me

Sql = "select distinct anoisencao from isencao where anoisencao>2000 order by anoisencao"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbAno.AddItem !anoisencao
       .MoveNext
    Loop
   .Close
End With
cmbAno.ListIndex = cmbAno.ListCount - 1

End Sub

Private Sub cmdPrint_Click()
Dim Sql As String, RdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, bDef As Boolean

Sql = "delete from relatorio_isencao"
cn.Execute Sql, rdExecDirect
Ocupado
Pb.value = 0
Sql = "SELECT isencao.CODREDUZIDO,isencao.anoisencao,isencao.codisencao,isencao.numprocesso,isencao.percisencao,isencao.motivo,processogti.CODCIDADAO,cidadao.nomecidadao,anoproc,numproc "
Sql = Sql & "From dbo.isencao INNER JOIN dbo.processogti ON isencao.anoproc = processogti.ANO  AND isencao.numproc = processogti.NUMERO INNER JOIN dbo.cidadao ON processogti.CODCIDADAO = cidadao.codcidadao "
Sql = Sql & "Where isencao.anoisencao = " & Val(cmbAno.Text) & "AND isencao.codisencao = 3"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then CallPb nPos, nTot
        Sql = "SELECT despacho FROM tramitacao WHERE ano=" & !AnoProc & " AND numero=" & !NumProc & " AND despacho<3"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount = 0 Then
            bDef = False
        Else
            If RdoAux2!despacho = 1 Then
                bDef = True
            Else
                bDef = False
            End If
        End If
        RdoAux2.Close
        
        Sql = "insert relatorio_isencao(Ano,Codigo,Processo,RequerenteId,Requerente,motivo,percentual,deferido) values("
        Sql = Sql & !anoisencao & "," & !CODREDUZIDO & ",'" & !numprocesso & "'," & !CodCidadao & ",'" & Mask(SubNull(!nomecidadao)) & "','" & Mask(!motivo) & "'," & Virg2Ponto(CStr(!percisencao)) & "," & IIf(bDef, 1, 0) & ")"
        
        cn.Execute Sql, rdExecDirect
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With
Liberado
Pb.value = 0
lblPB.Caption = "0%"
MsgBox "Fim"


End Sub

Private Sub CallPb(nPosF As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents

If ((nPosF * 100) / nTotal) <= 100 Then
   Pb.value = (nPosF * 100) / nTotal
Else
   Pb.value = 100
End If
lblPB.Caption = Int(Pb.value) & "%"

If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub

