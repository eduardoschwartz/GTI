VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmAssuntoPeriodo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Qtde dos assuntos por ano"
   ClientHeight    =   1350
   ClientLeft      =   9885
   ClientTop       =   7740
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   5205
   Begin VB.TextBox txtAno 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1395
      MaxLength       =   4
      TabIndex        =   3
      Top             =   315
      Width           =   870
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   345
      Left            =   3780
      TabIndex        =   1
      Top             =   900
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Imprimir"
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
      MICON           =   "frmAssuntoPeriodo.frx":0000
      PICN            =   "frmAssuntoPeriodo.frx":001C
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
      Left            =   180
      TabIndex        =   2
      Top             =   945
      Width           =   3375
      _ExtentX        =   5953
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
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano Pesquisa"
      Height          =   195
      Index           =   1
      Left            =   225
      TabIndex        =   0
      Top             =   375
      Width           =   1065
   End
End
Attribute VB_Name = "frmAssuntoPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type tAss
    Codigo As Integer
    Nome As String
End Type

Private Type tMain
    Codigo As Integer
    Nome As String
    Qtde As Integer
    Mes As Integer
    Valido As Boolean
End Type

Private Sub cmdPrint_Click()

If Val(txtAno.Text) < 1990 Or Val(txtAno.Text) > Year(Now) Then
   MsgBox "Digite um ano válido.", vbExclamation, "Atenção"
   Exit Sub
End If

Ocupado
Calcula
Liberado
End Sub

Private Sub Form_Load()
Centraliza Me
End Sub

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

Private Sub Calcula()
Dim Sql As String, RdoAux As rdoResultset, nPos As Long, nTot As Long, nMax As Integer
Dim aAss() As tAss, aMain() As tMain, nAno As Integer, nMes As Integer, nAssunto As Integer
Dim x As Long, y As Long, ax As String, Scr_hdc As Long, z As Long, myExcelFile As New ExcelFile
Dim cnExcel As ADODB.Connection, Rs As ADODB.Recordset, nCont As Integer, sFile As String, nRow As Long

nAno = Val(txtAno.Text)

Sql = "select max(codigo) as maximo from assunto"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nMax = RdoAux!maximo + 1
RdoAux.Close

ReDim aAss(nMax)

For nPos = 1 To nMax
    aAss(nPos).Codigo = nPos
    aAss(nPos).Nome = ""
Next

Sql = "select codigo,nome from assunto order by codigo"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        aAss(!Codigo).Nome = !Nome
       .MoveNext
    Loop
   .Close
End With

ReDim aMain(nMax * 12, 12)

x = 0
For nMes = 1 To 12
    For nPos = 1 To nMax
        aMain(nPos, nMes).Codigo = nPos
        aMain(nPos, nMes).Nome = aAss(nPos).Nome
        aMain(nPos, nMes).Qtde = 0
        aMain(nPos, nMes).Mes = nMes
        aMain(nPos, nMes).Valido = False
    Next
    
Next

nPos = 1
Sql = "SELECT ANO,NUMERO,CODASSUNTO,DATAENTRADA FROM processogti WHERE ano=" & nAno
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    Do Until .EOF
        If nPos Mod 50 = 0 Then CallPb nPos, nTot
        nMes = Month(!DATAENTRADA)
        nAssunto = !codassunto
        aMain(nAssunto, nMes).Qtde = aMain(nAssunto, nMes).Qtde + 1
        aMain(nAssunto, 1).Valido = True
       .MoveNext
    Loop
   .Close
End With
PBar.value = 100
Me.Refresh


Scr_hdc = GetDesktopWindow()
          
With myExcelFile
    FileName$ = sPathBin & "\Relatorio.xls"  'create spreadsheet in the current directory
    .CreateFile FileName$
    .SetColumnWidth 1, 1, 50
    .WriteValue xlsText, xlsFont3, xlsrightAlign, xlsNormal, 1, 1, "----"
    .WriteValue xlsText, xlsFont3, xlsrightAlign, xlsNormal, 1, 2, "Jan"
    .WriteValue xlsText, xlsFont3, xlsrightAlign, xlsNormal, 1, 3, "Fev"
    .WriteValue xlsText, xlsFont3, xlsrightAlign, xlsNormal, 1, 4, "Mar"
    .WriteValue xlsText, xlsFont3, xlsrightAlign, xlsNormal, 1, 5, "Abr"
    .WriteValue xlsText, xlsFont3, xlsrightAlign, xlsNormal, 1, 6, "Mai"
    .WriteValue xlsText, xlsFont3, xlsrightAlign, xlsNormal, 1, 7, "Jun"
    .WriteValue xlsText, xlsFont3, xlsrightAlign, xlsNormal, 1, 8, "Jul"
    .WriteValue xlsText, xlsFont3, xlsrightAlign, xlsNormal, 1, 9, "Ago"
    .WriteValue xlsText, xlsFont3, xlsrightAlign, xlsNormal, 1, 10, "Set"
    .WriteValue xlsText, xlsFont3, xlsrightAlign, xlsNormal, 1, 11, "Out"
    .WriteValue xlsText, xlsFont3, xlsrightAlign, xlsNormal, 1, 12, "Nov"
    .WriteValue xlsText, xlsFont3, xlsrightAlign, xlsNormal, 1, 13, "Dez"
    
    nRow = 2
    For x = 2 To nMax
        For y = 1 To 13
            If aMain(x - 1, 1).Valido = False Then
                GoTo Proximo
            Else
                If y = 1 Then
                    .WriteValue xlsText, xlsFont3, xlsLeftAlign, xlsNormal, nRow, y, aMain(x - 1, 1).Nome
                Else
                    .WriteValue xlsnumber, xlsFont3, xlsrightAlign, xlsNormal, nRow, y, aMain(x - 1, y - 1).Qtde
                End If
                        
            End If
        Next
        nRow = nRow + 1
Proximo:
    Next
   .CloseFile
End With

z = ShellExecute(Scr_hdc, "Open", "Relatorio.xls", "", sPathBin, SW_SHOWNORMAL)


End Sub


Private Sub txtAno_KeyPress(KeyAscii As Integer)
Tweak txtAno, KeyAscii, IntegerPositive
End Sub
