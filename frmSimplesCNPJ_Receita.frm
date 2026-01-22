VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmSimplesCNPJ_Receita 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportação de CNPJ para a Rec.Federal"
   ClientHeight    =   1365
   ClientLeft      =   6405
   ClientTop       =   4620
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   5655
   Begin VB.OptionButton optTipo 
      Caption         =   "Arquivo Exclusão"
      Height          =   240
      Index           =   1
      Left            =   2070
      TabIndex        =   4
      Top             =   810
      Value           =   -1  'True
      Width           =   1725
   End
   Begin VB.OptionButton optTipo 
      Caption         =   "Arquivo Inicial"
      Height          =   240
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   810
      Width           =   1725
   End
   Begin VB.CommandButton cmdImportar 
      Caption         =   "Importar"
      Enabled         =   0   'False
      Height          =   330
      Left            =   3735
      TabIndex        =   2
      Top             =   1530
      Width           =   1410
   End
   Begin prjChameleon.chameleonButton cmdExec 
      Height          =   345
      Left            =   4140
      TabIndex        =   0
      ToolTipText     =   "Executar a operação selecionada"
      Top             =   315
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Executar"
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
      MICON           =   "frmSimplesCNPJ_Receita.frx":0000
      PICN            =   "frmSimplesCNPJ_Receita.frx":001C
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
      Left            =   135
      TabIndex        =   1
      Top             =   405
      Width           =   3885
      _ExtentX        =   6853
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
End
Attribute VB_Name = "frmSimplesCNPJ_Receita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type tLaser
    Codigo As Long
    ano As Integer
    Area_Terreno As Double
    Area_Predial As Double
End Type

Private Sub cmdExec_Click()
If optTipo(0).value = True Then
    If MsgBox("Deseja gerar o arquivo inicial?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
        Arquivo_Inicial
    End If
Else
    If MsgBox("Deseja gerar o arquivo de exclusão?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
        Arquivo_Exclusao
    End If
End If

End Sub

Private Sub cmdImportar_Click()
Simples_Cnpj
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

Private Sub Simples_Cnpj()
Dim nCodReduz As Long, sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos2 As Integer, bFind As Boolean
Dim nPos As Long, nTot As Long, a2019() As tLaser, a2020() As tLaser, x As Integer, y As Integer
On Error GoTo Erro

sql = "SELECT cnpj From simplestmp order by cnpj"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 50 = 0 Then
           CallPb nPos, nTot
        End If
        sql = "insert simples_cnpj_receita(cnpj) values('" & !Cnpj & "')"
        cn.Execute sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With


MsgBox "Fim"


Exit Sub

Erro:
'MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub Arquivo_Inicial()

Dim sql As String, RdoAux As rdoResultset, nPos As Long, nTot As Long, sCnpj As String, nCodReduz As Long
Dim RdoAux2 As rdoResultset, dData As Date, nSit As Integer, nEncerrada As Integer, nReg As Integer, nPosReg As Integer

Ocupado

sql = "update simples_cnpj_receita set situacao=null,data=null,codreduz=null,encerrada=null"
cn.Execute sql, rdExecDirect

nPos = 1
'Sql = "select * from simples_cnpj_receita where cnpj='7786772000140' order by cnpj"
sql = "select * from simples_cnpj_receita order by cnpj"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    Do Until .EOF
       If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
            DoEvents
        End If
        sCnpj = !Cnpj
        nCodReduz = 0
        sql = "select codigomob from mobiliario where convert(bigint,cnpj)=" & Val(sCnpj)
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            If RdoAux2.RowCount > 1 Then
                nCodReduz = RdoAux2!codigomob
                RdoAux2.Close
                sql = "select codigomob from mobiliario where convert(bigint,cnpj)=" & Val(sCnpj) & " and dataencerramento is null"
                Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                If RdoAux2.RowCount > 0 Then
                   nCodReduz = RdoAux2!codigomob
                End If
            Else
                nCodReduz = RdoAux2!codigomob
'                If nCodReduz = 120087 Then MsgBox "teste"
            End If
            RdoAux2.Close
        Else
            RdoAux2.Close
            sql = "select codcidadao from cidadao where cnpj='" & sCnpj & "'"
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount > 0 Then
                nCodReduz = RdoAux2!CodCidadao
            End If
        End If
        If nCodReduz = 0 Then
            sql = "update simples_cnpj_receita set situacao=1, data='" & Format(Now, "mm/dd/yyyy") & "' where cnpj='" & sCnpj & "'"
            cn.Execute sql, rdExecDirect
        Else
            sql = "select * from periodosn where codigo=" & nCodReduz & " order by dataini desc"
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount > 0 Then
                dData = RdoAux2!DATAINI
            Else
                dData = Format(Now, "mm/dd/yyyy")
            End If
            RdoAux2.Close
            
            nEncerrada = 0
            sql = "select codigomob,dataencerramento from mobiliario where codigomob=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If Not IsNull(RdoAux2!dataencerramento) And nCodReduz < 500000 Then
                nEncerrada = 1
                nSit = 1
            Else
                nSit = 0
                RdoAux2.Close
                sql = "select * from debitoparcela where codreduzido=" & nCodReduz & " and datavencimento<'" & Format("31/12/2025", "mm/dd/yyyy") & "' and codlancamento in (3,5,6,13) and (statuslanc=3 or statuslanc=42 or statuslanc=43 or statuslanc=39 or statuslanc=40 or statuslanc=41 or statuslanc=39)"
                Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                If RdoAux2.RowCount > 0 Then
                    nSit = 1
                End If
                RdoAux2.Close
            End If
            
            If nSit = 0 Then
                'SUSPENÇÃO
                 sql = "SELECT CODTIPOEVENTO,DATAPROCEVENTO FROM MOBILIARIOEVENTO WHERE CODMOBILIARIO=" & nCodReduz
                 sql = sql & " ORDER BY DATAEVENTO DESC"
                 Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                 With RdoAux2
                     If .RowCount > 0 Then
                         If !CODTIPOEVENTO = 2 Then
                            nSit = 1
                         End If
                     End If
                    .Close
                 End With
            End If
            
            sql = "update simples_cnpj_receita set situacao=" & nSit & ", data='" & Format(dData, "mm/dd/yyyy") & "', codreduz=" & nCodReduz & ", encerrada=" & nEncerrada & " where cnpj='" & sCnpj & "'"
            cn.Execute sql, rdExecDirect
            
        End If
                
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

nReg = 1
nPos = 1

Inicio:
Open sPathBin & "\cnpjsimples.txt" For Output As #1
sql = "select * from simples_cnpj_receita where situacao=1 order by cnpj"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    .Move nPos
    nPosReg = 1
    'Print #1, "INI00000000000"
    Print #1, "00000000000000"
    Do Until .EOF
'        If nPosReg >= 200 Then
'            GoTo Proximo
'        End If
        Print #1, Format(!Cnpj, "00000000000000")
        
        nPosReg = nPosReg + 1
        nPos = nPos + 1
       .MoveNext
    Loop
    Print #1, "99999999999999"
    GoTo Fim
End With

Proximo:
Close #1
nReg = nReg + 1
GoTo Inicio


Fim:
Close #1
Liberado
MsgBox "Gravação concluída", vbInformation, "Informação"

End Sub

Private Sub Arquivo_Exclusao()

Dim sql As String, RdoAux As rdoResultset, nPos As Long, nTot As Long, sCnpj As String, nCodReduz As Long
Dim RdoAux2 As rdoResultset, dData As Date, nSit As Integer, nEncerrada As Integer, nReg As Integer, nPosReg As Integer

Ocupado

sql = "update simples_cnpj_receita set situacao=null,data=null,codreduz=null,encerrada=null"
cn.Execute sql, rdExecDirect

nPos = 1
'Sql = "select * from simples_cnpj_receita where cnpj='44539610000133' order by cnpj"
sql = "select * from simples_cnpj_receita where regularizada is null order by cnpj"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    Do Until .EOF
       If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
            DoEvents
        End If
        sCnpj = !Cnpj
        nCodReduz = 0
        sql = "select codigomob from mobiliario where convert(bigint,cnpj)=" & Val(sCnpj)
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            If RdoAux2.RowCount > 1 Then
                nCodReduz = RdoAux2!codigomob
                RdoAux2.Close
                sql = "select codigomob from mobiliario where convert(bigint,cnpj)=" & Val(sCnpj) & " and dataencerramento is null"
                Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                If RdoAux2.RowCount > 0 Then
                   nCodReduz = RdoAux2!codigomob
                End If
            Else
                nCodReduz = RdoAux2!codigomob
'                If nCodReduz = 120087 Then MsgBox "teste"
            End If
            RdoAux2.Close
        Else
            RdoAux2.Close
            sql = "select codcidadao from cidadao where cnpj='" & sCnpj & "'"
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount > 0 Then
                nCodReduz = RdoAux2!CodCidadao
            End If
        End If
        If nCodReduz = 0 Then
            sql = "update simples_cnpj_receita set situacao=1, data='" & Format(Now, "mm/dd/yyyy") & "' where cnpj='" & sCnpj & "'"
            cn.Execute sql, rdExecDirect
        Else
            sql = "select * from periodosn where codigo=" & nCodReduz & " order by dataini desc"
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount > 0 Then
                dData = RdoAux2!DATAINI
            Else
                dData = Format(Now, "mm/dd/yyyy")
            End If
            RdoAux2.Close
            
            nEncerrada = 0
            sql = "select codigomob,dataencerramento from mobiliario where codigomob=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If Not IsNull(RdoAux2!dataencerramento) And nCodReduz < 500000 Then
                nEncerrada = 1
                nSit = 1
            Else
                nSit = 0
                RdoAux2.Close
                sql = "select * from debitoparcela where codreduzido=" & nCodReduz & " and datavencimento<'" & Format("31/12/2025", "mm/dd/yyyy") & "' and codlancamento in (3,5,6,13) and (statuslanc=3 or statuslanc=42 or statuslanc=43 or statuslanc=39 or statuslanc=40 or statuslanc=41 or statuslanc=39)"
                Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                If RdoAux2.RowCount > 0 Then
                    nSit = 1
                End If
                RdoAux2.Close
            End If
            
            If nSit = 0 Then
                'SUSPENÇÃO
                 sql = "SELECT CODTIPOEVENTO,DATAPROCEVENTO FROM MOBILIARIOEVENTO WHERE CODMOBILIARIO=" & nCodReduz
                 sql = sql & " ORDER BY DATAEVENTO DESC"
                 Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                 With RdoAux2
                     If .RowCount > 0 Then
                         If !CODTIPOEVENTO = 2 Then
                            nSit = 1
                         End If
                     End If
                    .Close
                 End With
            End If
            
            If nSit = 0 Then
                sql = "update simples_cnpj_receita set regularizada=1 where cnpj='" & sCnpj & "'"
                cn.Execute sql, rdExecDirect
            End If
        End If
                
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

nReg = 1
nPos = 1

Inicio:
Open sPathBin & "\cnpjsimplesexc.txt" For Output As #1
sql = "select * from simples_cnpj_receita where regularizada=1 order by cnpj"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    .Move nPos
    nPosReg = 1
    Print #1, "EXC00000000000"
    'Print #1, "00000000000000"
    Do Until .EOF
'        If nPosReg >= 200 Then
'            GoTo Proximo
'        End If
        Print #1, Format(!Cnpj, "00000000000000")
        
        nPosReg = nPosReg + 1
        nPos = nPos + 1
       .MoveNext
    Loop
    Print #1, "99999999999999"
    GoTo Fim
End With

Proximo:
Close #1
nReg = nReg + 1
GoTo Inicio


Fim:
Close #1
Liberado
MsgBox "Gravação concluída", vbInformation, "Informação"

End Sub

Private Sub Arquivo_Complementar()

Dim sql As String, RdoAux As rdoResultset, nPos As Long, nTot As Long, sCnpj As String, nCodReduz As Long
Dim RdoAux2 As rdoResultset, dData As Date, nSit As Integer, nEncerrada As Integer, nReg As Integer, nPosReg As Integer

Ocupado

sql = "update simples_cnpj_receita set situacao=null,data=null,codreduz=null,encerrada=null"
cn.Execute sql, rdExecDirect

nPos = 1
'Sql = "select * from simples_cnpj_receita where cnpj='7786772000140' order by cnpj"
sql = "select * from simples_cnpj_receita order by cnpj"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    Do Until .EOF
       If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
            DoEvents
        End If
        sCnpj = !Cnpj
        nCodReduz = 0
        sql = "select codigomob from mobiliario where convert(bigint,cnpj)=" & Val(sCnpj)
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            If RdoAux2.RowCount > 1 Then
                nCodReduz = RdoAux2!codigomob
                RdoAux2.Close
                sql = "select codigomob from mobiliario where convert(bigint,cnpj)=" & Val(sCnpj) & " and dataencerramento is null"
                Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                If RdoAux2.RowCount > 0 Then
                   nCodReduz = RdoAux2!codigomob
                End If
            Else
                nCodReduz = RdoAux2!codigomob
'                If nCodReduz = 120087 Then MsgBox "teste"
            End If
            RdoAux2.Close
        Else
            RdoAux2.Close
            sql = "select codcidadao from cidadao where cnpj='" & sCnpj & "'"
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount > 0 Then
                nCodReduz = RdoAux2!CodCidadao
            End If
        End If
        If nCodReduz = 0 Then
            sql = "update simples_cnpj_receita set situacao=1, data='" & Format(Now, "mm/dd/yyyy") & "' where cnpj='" & sCnpj & "'"
            cn.Execute sql, rdExecDirect
        Else
            sql = "select * from periodosn where codigo=" & nCodReduz & " order by dataini desc"
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount > 0 Then
                dData = RdoAux2!DATAINI
            Else
                dData = Format(Now, "mm/dd/yyyy")
            End If
            RdoAux2.Close
            
            nEncerrada = 0
            sql = "select codigomob,dataencerramento from mobiliario where codigomob=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If Not IsNull(RdoAux2!dataencerramento) And nCodReduz < 500000 Then
                nEncerrada = 1
                nSit = 1
            Else
                nSit = 0
                RdoAux2.Close
                sql = "select * from debitoparcela where codreduzido=" & nCodReduz & " and datavencimento<'" & Format("31/12/2025", "mm/dd/yyyy") & "' and codlancamento in (3,5,6,13) and (statuslanc=3 or statuslanc=42 or statuslanc=43 or statuslanc=39 or statuslanc=40 or statuslanc=41 or statuslanc=39)"
                Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                If RdoAux2.RowCount > 0 Then
                    nSit = 1
                End If
                RdoAux2.Close
            End If
            
            If nSit = 0 Then
                'SUSPENÇÃO
                 sql = "SELECT CODTIPOEVENTO,DATAPROCEVENTO FROM MOBILIARIOEVENTO WHERE CODMOBILIARIO=" & nCodReduz
                 sql = sql & " ORDER BY DATAEVENTO DESC"
                 Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                 With RdoAux2
                     If .RowCount > 0 Then
                         If !CODTIPOEVENTO = 2 Then
                            nSit = 1
                         End If
                     End If
                    .Close
                 End With
            End If
            
            sql = "update simples_cnpj_receita set situacao=" & nSit & ", data='" & Format(dData, "mm/dd/yyyy") & "', codreduz=" & nCodReduz & ", encerrada=" & nEncerrada & " where cnpj='" & sCnpj & "'"
            cn.Execute sql, rdExecDirect
            
        End If
                
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

nReg = 1
nPos = 1

Inicio:
Open sPathBin & "\cnpjsimples.txt" For Output As #1
sql = "select * from simples_cnpj_receita where situacao=1 order by cnpj"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    .Move nPos
    nPosReg = 1
    'Print #1, "INI00000000000"
    Print #1, "00000000000000"
    Do Until .EOF
'        If nPosReg >= 200 Then
'            GoTo Proximo
'        End If
        Print #1, Format(!Cnpj, "00000000000000")
        
        nPosReg = nPosReg + 1
        nPos = nPos + 1
       .MoveNext
    Loop
    Print #1, "99999999999999"
    GoTo Fim
End With

Proximo:
Close #1
nReg = nReg + 1
GoTo Inicio


Fim:
Close #1
Liberado
MsgBox "Gravação concluída", vbInformation, "Informação"

End Sub

