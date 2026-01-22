VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportaBanco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importação de Arquivos Bancários"
   ClientHeight    =   5130
   ClientLeft      =   9045
   ClientTop       =   3885
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   8325
   Begin VB.CommandButton Command3 
      Caption         =   "Importar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   6570
      TabIndex        =   12
      Top             =   6300
      Width           =   1125
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Enabled         =   0   'False
      Height          =   345
      Left            =   6600
      TabIndex        =   11
      Top             =   6720
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "frmbaixabancaria"
      Enabled         =   0   'False
      Height          =   345
      Left            =   6390
      TabIndex        =   9
      Top             =   5805
      Width           =   1425
   End
   Begin Tributacao.jcFrames frProgress 
      Height          =   1155
      Left            =   1950
      Top             =   1740
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2037
      FrameColor      =   255
      FillColor       =   4210688
      TextBoxColor    =   8454016
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   ""
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
      ThemeColor      =   3
      ColorFrom       =   192
      ColorTo         =   8438015
      Begin Tributacao.XP_ProgressBar pBar 
         Height          =   165
         Left            =   150
         TabIndex        =   8
         Top             =   780
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   291
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
         Color           =   255
         Scrolling       =   1
      End
      Begin VB.Label lblFileName 
         BackStyle       =   0  'Transparent
         Caption         =   "ARRECADA08.ret"
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   360
         Width           =   4155
      End
      Begin VB.Label lblFileNumber 
         BackStyle       =   0  'Transparent
         Caption         =   "Importando Arquivo 0 de 0"
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   90
         Width           =   4305
      End
   End
   Begin VB.TextBox txtResult 
      Appearance      =   0  'Flat
      Height          =   2025
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2460
      Width           =   8145
   End
   Begin VB.ListBox lstNome 
      Height          =   645
      Left            =   20490
      TabIndex        =   2
      Top             =   12750
      Width           =   6105
   End
   Begin VB.ListBox lstArq 
      Appearance      =   0  'Flat
      Height          =   1785
      Left            =   90
      TabIndex        =   1
      Top             =   330
      Width           =   8145
   End
   Begin prjChameleon.chameleonButton cmdArq 
      Height          =   360
      Left            =   6120
      TabIndex        =   0
      ToolTipText     =   "Selecione o arquivo a importar"
      Top             =   4635
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "Selecionar &Arquivos"
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
      MICON           =   "frmImportaBanco.frx":0000
      PICN            =   "frmImportaBanco.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   17820
      Top             =   12840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Color           =   8388608
      DialogTitle     =   "Selecione o arquivo de GIA"
   End
   Begin VB.ListBox List1 
      Height          =   2010
      ItemData        =   "frmImportaBanco.frx":00BA
      Left            =   570
      List            =   "frmImportaBanco.frx":00BC
      TabIndex        =   10
      Top             =   5910
      Width           =   5385
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Resultado da importação"
      Height          =   225
      Index           =   1
      Left            =   150
      TabIndex        =   5
      Top             =   2250
      Width           =   3345
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Arquivos Selecionados"
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3345
   End
End
Attribute VB_Name = "frmImportaBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Banco
    Codigo As Integer
    Nome As String
End Type

Dim aBanco() As Banco, nQtdeArquivos As Integer, nFilePos As Integer, nFileTot As Integer, nCodigoSimples As Long

Private Sub cmdArq_Click()
On Error GoTo Erro:

Dim vFiles As Variant
Dim lFile As Long

txtResult.Text = ""
lstArq.Clear

With cDialog
    
    .FileName = "" 'Clear the filename
    .CancelError = True
    .MaxFileSize = 30000
    .DialogTitle = "Select File(s)..."
    .flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNHideReadOnly 'Flags, allows Multi select, Explorer style and hide the Read only tag
    .Filter = "All files (*.*)|*.*"
    .ShowOpen
    
    vFiles = Split(.FileName, Chr(0)) 'Splits the filename up in segments
    nQtdeArquivos = UBound(vFiles)
    
    If MsgBox("Deseja importar este(s) arquivo(s)?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
        txtResult.Text = "Operação cancelada."
        Exit Sub
    End If
    
    txtResult.Text = IIf(nQtdeArquivos = 0, 1, nQtdeArquivos) & " arquivo(s) selecionado(s)."
    Ocupado
    frProgress.Visible = True
    cmdArq.Enabled = False
    If UBound(vFiles) = 0 Then ' If there is only 1 file then do this
        lstArq.AddItem .FileName
        lstNome.AddItem .FileTitle
        nFileTot = 1
        nFilePos = 1
        ImportaArquivo .FileName, .FileTitle
    Else
        For lFile = 1 To UBound(vFiles) ' More than 1 file then do this until there are no more files
            lstArq.AddItem vFiles(0) + "\" & vFiles(lFile)
            lstNome.AddItem vFiles(lFile)
            nFileTot = UBound(vFiles)
            nFilePos = lFile
            ImportaArquivo vFiles(0) + "\" & vFiles(lFile), CStr(vFiles(lFile))
        Next
    End If
    Liberado
    frProgress.Visible = False
    cmdArq.Enabled = True

End With

MsgBox "Importação concluída!", vbInformation, "Atenção"

Exit Sub
Erro:
If Err.Number = 32755 Then
Else
    MsgBox Err.Description
End If

End Sub

Private Sub Command1_Click()
frmBaixaBancariaNovo.show
End Sub

Private Sub Command2_Click()
List1.Clear
Dim strStartPath As String
    strStartPath = "D:\Tmp\2019\01"
   ' ListFolder strStartPath
    strStartPath = "D:\Tmp\2019\02"
    'ListFolder strStartPath
    strStartPath = "D:\Tmp\2019\03"
    'ListFolder strStartPath
    strStartPath = "D:\Tmp\2019\04"
    'ListFolder strStartPath
    strStartPath = "D:\Tmp\2019\05"
    ListFolder strStartPath
    strStartPath = "D:\Tmp\2019\06"
    ListFolder strStartPath
    strStartPath = "D:\Tmp\2019\07"
    ListFolder strStartPath
    strStartPath = "D:\Tmp\2019\08"
    ListFolder strStartPath
    strStartPath = "D:\Tmp\2019\09"
    ListFolder strStartPath
    strStartPath = "D:\Tmp\2019\10"
    ListFolder strStartPath
    strStartPath = "D:\Tmp\2019\11"
 '   ListFolder strStartPath
    strStartPath = "D:\Tmp\2019\12"
'    ListFolder strStartPath

End Sub

Private Sub ListFolder(sFolderPath As String)
    Dim FS As New FileSystemObject
    Dim FSfolder As Folder
    Dim subfolder As Folder
    Dim i As Integer
    
    Set FSfolder = FS.GetFolder(sFolderPath)
 
    For Each subfolder In FSfolder.SubFolders
        DoEvents
        i = i + 1
        
        '***
        Dim fso As New FileSystemObject
Dim fld As Folder
Dim fil As File
Set fld = fso.GetFolder(subfolder)
For Each fil In fld.Files
  'Debug.Print fil.Name
  List1.AddItem subfolder & "\" & fil.Name
Next
Set fil = Nothing
Set fld = Nothing
Set fso = Nothing
        
        '*****
        
'        List1.AddItem subfolder
    Next subfolder
    Set FSfolder = Nothing
'    MsgBox "Total sub folders in " & sFolderPath & " : " & i
End Sub


Private Sub Command3_Click()
On Error GoTo Erro:

Dim vFiles As Variant
Dim lFile As Long

txtResult.Text = ""
lstArq.Clear


    
    nQtdeArquivos = List1.ListCount
    txtResult.Text = IIf(nQtdeArquivos = 0, 1, nQtdeArquivos) & " arquivo(s) selecionado(s)."
    frProgress.Visible = True
    
        For lFile = 0 To List1.ListCount - 1 ' More than 1 file then do this until there are no more files
            lstArq.AddItem List1.Text
            lstNome.AddItem List1.Text
            nFileTot = nQtdeArquivos
            nFilePos = lFile
            ImportaArquivo List1.List(lFile), GetFileNameFromPath(List1.List(lFile))
        Next
   
    frProgress.Visible = False
    


MsgBox "Importação concluída!", vbInformation, "Atenção"

Exit Sub
Erro:
If Err.Number = 32755 Then
Else
    MsgBox Err.Description
End If

End Sub

Private Sub Form_Load()
Centraliza Me
CarregaBanco
End Sub

Private Sub CarregaBanco()
Dim sql As String, RdoAux As rdoResultset

ReDim aBanco(0)
sql = "select codbanco,nomebanco from banco order by codbanco"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aBanco(UBound(aBanco) + 1)
        aBanco(UBound(aBanco)).Codigo = !CodBanco
        aBanco(UBound(aBanco)).Nome = !NomeBanco
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Function RetornaNomeBanco(Codigo As Integer) As String
Dim x As Integer, bFind As Boolean

bFind = False
For x = 1 To UBound(aBanco) - 1
    If aBanco(x).Codigo = Codigo Then
        bFind = True
        Exit For
    End If
Next
If bFind Then
    RetornaNomeBanco = aBanco(x).Nome
Else
    RetornaNomeBanco = ""
End If

End Function

Private Sub ImportaArquivo(sFileName As String, sFileTitle As String)
Dim fso As FileSystemObject, TS As TextStream, row As String, sTipoArq As String, x As Integer, nRegImportado As Integer, z As Long, nRetorno As Integer
Dim nCodBanco As Integer, sNomeBanco As String, sTipoReg As String, sDataPag As String, sDataCredito As String, sAgencia As String, sContaDeposito As String
Dim nNumDoc As Long, nValorPago As Double, sSitRetorno As String, sql As String, aDataCredito() As String, RdoAux As rdoResultset, nCodReduz As Long, nCodMov As Integer
Dim sDataVencto As String, sCnpj As String, nAno As Integer, nMes As Integer, nValorPrincipal As Double, nValorMulta As Double, nValorJuros As Double, sDataOpcao As String
Dim nValorGuia As Double, sConvenio As String, sCodMov As String, sConta As String, bFind As Boolean, nCodBancoReceptor As Integer, nPos As Long, nTot As Long
Dim nContN As Integer, nContI As Integer, nContP As Integer, nNumDocSimples As Long

On Error GoTo Erro
nContN = 0: nContI = 0: nContP = 0
lblFileNumber.Caption = "Importando arquivo " & nFilePos & " de " & nFileTot
lblFileName.Caption = sFileTitle
PBar.value = 0
nTot = FileRowCount(sFileName)

    
ReDim aDataCredito(0)
Set fso = New FileSystemObject


'**** GRAVA BACKUP *********
Dim sGuid As String, sNomeArquivo As String
sGuid = GerarIDUnico(8)
sNomeArquivo = sFileTitle
sql = "insert bankfileheader(id,nome_arquivo,data_gravado) values('" & sGuid & "','" & Mask(sNomeArquivo) & "','" & Format(Now, "mm/dd/yyyy") & "')"
cn.Execute sql, rdExecDirect

nPos = 1: nRegImportado = 0
Set TS = fso.OpenTextFile(sFileName, ForReading)
Do Until TS.AtEndOfStream
    If nPos Mod 10 = 0 Then CallPb nPos, nTot
    row = TS.ReadLine
    sql = "insert bankfiledata (id,seq,data) values('" & sGuid & "'," & nPos & ",'" & Mask(row) & "')"
    cn.Execute sql, rdExecDirect
    nPos = nPos + 1
Loop
TS.Close


PBar.value = 0
Me.Refresh

Set TS = fso.OpenTextFile(sFileName, ForReading)
row = TS.ReadLine
If Left(row, 1) = "A" And InStr(1, row, "DEBITO AUT") = 0 Then  'ARQUIVO NORMAL
    sTipoArq = "NORMAL"
ElseIf Left(row, 15) = "100000001DAF607" Then 'ARQUIVO SIMPLES
    sTipoArq = "SIMPLES"
    sql = "SELECT MAX(numero_documento) AS MAXIMO FROM importacao_banco where numero_documento<2000000"
    Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurReadOnly)
    
    nNumDocSimples = Val(SubNull(RdoAux!maximo)) + 1
    RdoAux.Close
ElseIf Left(row, 8) = "00100000" Or Left(row, 8) = "03300000" Then  'ARQUIVO COBRANÇA BANCO DO BRASIL OU SANTANDER
    sTipoArq = "COBRANCA"
ElseIf InStr(1, row, "DEBITO AUT") > 0 Then 'ARQUIVO DEBITO AUTOMATICO
    sTipoArq = "DEBAUT"
End If
TS.Close

If sTipoArq = "NORMAL" Then
    GoTo LEARQNORMAL
ElseIf sTipoArq = "SIMPLES" Then
    GoTo LEARQSIMPLES
ElseIf sTipoArq = "COBRANCA" Then
    nTot = nTot / 2
    GoTo LEARQCOBRANCA
ElseIf sTipoArq = "DEBAUT" Then
    GoTo LEDEBAUT
Else
    txtResult.Text = txtResult.Text & vbCrLf & "O arquivo " & sFileName & " é inválido!"
End If

Exit Sub


'******** ARRECADAÇÃO ********
LEARQNORMAL:
nPos = 1: nRegImportado = 0
Set TS = fso.OpenTextFile(sFileName, ForReading)
Do Until TS.AtEndOfStream
    If nPos Mod 10 = 0 Then CallPb nPos, nTot
    row = TS.ReadLine
    If nPos = 1 Then
        nCodBanco = Mid(row, 43, 3)
        sNomeBanco = Trim(Mid(row, 46, 20))
    Else
        sTipoReg = Left(row, 1)
        If sTipoReg = "G" Then
            sDataCredito = ConvDataSerial(Mid(row, 30, 8))
            z = BinarySearchString(aDataCredito, sDataCredito)
            If z = -1 Then
                ReDim Preserve aDataCredito(UBound(aDataCredito) + 1)
                aDataCredito(UBound(aDataCredito)) = sDataCredito
                sql = "select * from importacao_banco where codigo_banco=" & nCodBanco & " and data_credito='" & Format(sDataCredito, "mm/dd/yyyy") & "' and nome_arquivo='" & sFileTitle & "'"
                Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                If RdoAux.RowCount > 0 Then
                   GoTo ARQUIVO_EXISTENTE
                End If
                RdoAux.Close
            End If
            sDataPag = ConvDataSerial(Mid(row, 22, 8))
            sDataVencto = ConvDataSerial(Mid(row, 57, 8))
            nNumDoc = CLng(Mid(row, 65, 9))
            nValorPago = CDbl(Mid(row, 82, 12)) / 100
            sAgencia = Mid(row, 2, 4)
            sConta = Val(Trim(RetornaNumero(Mid(row, 6, 16))))
            sConvenio = Trim(RetornaNumero(Mid(row, 3, 20)))
            sql = "select numdocumento from numdocumento where numdocumento=" & nNumDoc
            Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If RdoAux.RowCount > 0 Then
                sSitRetorno = "00-BAIXA NORMAL"
            Else
                sSitRetorno = "01-DOCUMENTO NÃO ENCONTRADO"
            End If
            RdoAux.Close
            If Not IsDate(sDataVencto) Then sDataVencto = sDataPag
SqlNormal:
            sql = "insert importacao_banco(codigo_banco,data_credito,nome_arquivo,data_pagamento,valor_pago,numero_documento,agencia,situacao_retorno,data_vencimento,convenio,conta_corrente) values("
            sql = sql & nCodBanco & ",'" & Format(sDataCredito, "mm/dd/yyyy") & "','" & sFileTitle & "','" & Format(sDataPag, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(nValorPago)) & ","
            sql = sql & nNumDoc & ",'" & sAgencia & "','" & sSitRetorno & "','" & Format(sDataVencto, "mm/dd/yyyy") & "','" & sConvenio & "','" & sConta & "')"
            cn.Execute sql, rdExecDirect
            nRegImportado = nRegImportado + 1
        End If
    End If
    nPos = nPos + 1
Loop
TS.Close

Grava_Arquivo_Banco sFileName, sFileTitle, sDataCredito, nCodBanco, Val(sAgencia), False
'sql = "INSERT INTO BankFile (data_gravado, FileName, FileData) SELECT '" & Format(sDataCredito, "mm/dd/yyyy") & "','" & sFileTitle & "',BulkColumn FROM OPENROWSET(BULK N'" & sFileName & "', SINGLE_BLOB) AS FileData;"

sql = "update bankfileheader set data_credito='" & Format(sDataCredito, "mm/dd/yyyy") & "',ready=1 where id='" & sGuid & "'"
cn.Execute sql
txtResult.Text = txtResult.Text & vbCrLf & "Importado " & nRegImportado & " registros do arquivo " & sFileName & " - Arrecadação (" & sNomeBanco & ")"
Exit Sub

LEARQSIMPLES:
nPos = 1
Set TS = fso.OpenTextFile(sFileName, ForReading)
Do Until TS.AtEndOfStream
    If nPos Mod 10 = 0 Then CallPb nPos, nTot
    row = TS.ReadLine
    If nPos = 1 Then
        nCodBanco = Mid(row, 76, 3)
        sNomeBanco = RetornaNomeBanco(nCodBanco)
        If sNomeBanco = "" Then
            sNomeBanco = "Sem convênio"
        End If
    Else
        sTipoReg = Left(row, 1)
        If sTipoReg = "2" Then
            sDataCredito = ConvDataSerial(Mid(row, 10, 8))
            z = BinarySearchString(aDataCredito, sDataCredito)
            If z = -1 Then
                ReDim Preserve aDataCredito(UBound(aDataCredito) + 1)
                aDataCredito(UBound(aDataCredito)) = sDataCredito
                sql = "select * from importacao_banco where codigo_banco=" & nCodBanco & " and data_credito='" & Format(sDataCredito, "mm/dd/yyyy") & "' and nome_arquivo='" & sFileTitle & "'"
                Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                If RdoAux.RowCount > 0 Then
                    GoTo ARQUIVO_EXISTENTE
                End If
                RdoAux.Close
            End If
            sDataPag = ConvDataSerial(Mid(row, 10, 8))
            sDataVencto = ConvDataSerial(Mid(row, 18, 8))
            nValorPrincipal = CDbl(Mid(row, 107, 17)) / 100
            nValorMulta = CDbl(Mid(row, 141, 17)) / 100
            nValorJuros = CDbl(Mid(row, 124, 17)) / 100
            nValorGuia = nValorPrincipal + nValorJuros + nValorMulta
            nAno = Val(Mid(row, 101, 4))
            nMes = Val(Mid(row, 105, 2))
            sCnpj = Mid(row, 75, 14)
            sql = "select codigomob from mobiliario where cnpj='" & sCnpj & "'"
            Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If RdoAux.RowCount > 0 Then
                nCodReduz = RdoAux!codigomob
                RdoAux.Close
            Else
                sql = "select codcidadao from cidadao where cnpj='" & sCnpj & "'"
                Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                If RdoAux.RowCount > 0 Then
                    nCodReduz = RdoAux!CodCidadao
                End If
            End If
            
            sAgencia = Mid(row, 223, 4)
            'sSitRetorno = "CNPJ: " & Format(sCNPJ, "0#\.###\.###/####-##")
            
            nCodigoSimples = nCodReduz
            
            nNumDoc = GravaSimples(sCnpj, sDataVencto, nAno, nMes, nValorGuia, sFileTitle, sDataCredito, nCodBanco, sAgencia)
            If nNumDoc > 0 Then
                sql = "insert importacao_banco(codigo_banco,data_credito,nome_arquivo,data_pagamento,valor_pago,numero_documento,agencia,data_vencimento,simples_nacional,cnpj,ano,mes,codigo_reduzido) values("
                sql = sql & nCodBanco & ",'" & Format(sDataCredito, "mm/dd/yyyy") & "','" & sFileTitle & "','" & Format(sDataPag, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(nValorGuia)) & ","
                sql = sql & nNumDocSimples & ",'" & sAgencia & "','" & Format(sDataVencto, "mm/dd/yyyy") & "'," & 1 & ",'" & sCnpj & "'," & nAno & "," & nMes & "," & nCodigoSimples & ")"
                cn.Execute sql, rdExecDirect
            End If
            nNumDocSimples = nNumDocSimples + 1
            nRegImportado = nRegImportado + 1
        End If
    End If
    nPos = nPos + 1
Loop
TS.Close
Grava_Arquivo_Banco sFileName, sFileTitle, sDataCredito, nCodBanco, Val(sAgencia), False
sql = "update bankfileheader set data_credito='" & Format(sDataCredito, "mm/dd/yyyy") & "',ready=1 where id='" & sGuid & "'"
'sql = "INSERT INTO BankFile (data_gravado, FileName, FileData) SELECT '" & Format(sDataCredito, "mm/dd/yyyy") & "','" & sFileTitle & "',BulkColumn FROM OPENROWSET(BULK N'" & sFileName & "', SINGLE_BLOB) AS FileData;"
cn.Execute sql

txtResult.Text = txtResult.Text & vbCrLf & "Incluído o arquivo " & sFileName & " - S.Nacional (" & sNomeBanco & ")"
Exit Sub


LEARQCOBRANCA:
nPos = 1
Set TS = fso.OpenTextFile(sFileName, ForReading)
Do Until TS.AtEndOfStream
    If nPos Mod 10 = 0 Then CallPb nPos, nTot
    row = TS.ReadLine
    If nPos = 1 Then
       'Header de arquivo
        nCodBanco = Left(row, 3)
        sNomeBanco = Trim(Mid(row, 103, 30))
        sConvenio = Mid(row, 35, 7)
    ElseIf nPos = 2 Then
       'Header de lote, nada a importar
    Else
        If Left(row, 8) = "00199999" Or (Left(row, 3) = "033" And Mid(row, 14, 1) = " ") Then
           'Trailer de Lote
            Exit Do
        Else
            sTipoReg = Mid(row, 14, 1)
            If sTipoReg = "T" Then
                sCodMov = Mid(row, 16, 2)
                If sCodMov = "06" Or sCodMov = "17" Or sCodMov = "61" Then
                    If Mid(row, 45, 2) = "00" Then
                        nNumDoc = Val(Mid(row, 47, 8))
                    ElseIf Mid(row, 45, 1) <> "0" Then
                        nNumDoc = Val(Mid(row, 45, 8))
                    Else
                        nNumDoc = Val(Mid(row, 46, 8))
                    End If
                    If nNumDoc > 200000 And nNumDoc < 300000 Then
                        nNumDoc = Val(Mid(sReg, 45, 10))
                    End If
                    
                    sql = "select numdocumento from numdocumento where numdocumento=" & nNumDoc
                    Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                    If RdoAux.RowCount > 0 Then
                        sSitRetorno = "00-BAIXA NORMAL"
                    Else
                        sSitRetorno = "01-DOCUMENTO NÃO ENCONTRADO"
                    End If
                    RdoAux.Close
                    nRetorno = Val(Mid(row, 214, 2))
                    nCodBancoReceptor = Val(Mid(row, 97, 3))
                    sAgencia = Mid(row, 100, 5)
                    sConta = Val(Mid(row, 24, 13))
                    sDataVencto = ConvDataSerialBB(Mid(row, 74, 8))
                    nValorPago = CDbl(Mid(row, 84, 15) / 100)
                    'next line read U
                    row = TS.ReadLine
                    sDataCredito = ConvDataSerialBB(Mid(row, 146, 8))
                    If Not IsDate(sDataCredito) Then
                        sDataCredito = ConvDataSerialBB(Mid(row, 138, 8))
                    End If
                    z = BinarySearchString(aDataCredito, sDataCredito)
                    If z = -1 Then
                        ReDim Preserve aDataCredito(UBound(aDataCredito) + 1)
                        aDataCredito(UBound(aDataCredito)) = sDataCredito
                        sql = "select * from importacao_banco where codigo_banco=" & nCodBanco & " and data_credito='" & Format(sDataCredito, "mm/dd/yyyy") & "' and nome_arquivo='" & sFileTitle & "'"
                        Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        '                If RdoAux.RowCount > 0 Then GoTo ARQUIVO_EXISTENTE
                        RdoAux.Close
                    End If
                    sDataPag = ConvDataSerialBB(Mid(row, 138, 8))
                    If Val(Mid(row, 78, 15)) > 0 Then
                        nValorPago = CDbl(Mid(row, 78, 15) / 100)
                    End If
                    If sDataVencto = "00/00/0000" Or Not IsDate(sDataVencto) Then
                        sDataVencto = sDataPag
                    End If
                                        
                    bFind = False
                    If sCodMov = "17" Then
                        sql = "select * from importacao_banco where codigo_banco=" & nCodBanco & " and data_credito='" & Format(sDataCredito, "mm/dd/yyyy") & "' and "
                        sql = sql & "nome_arquivo='" & sFileName & "' and numero_documento=" & nNumDoc
                        Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                        If RdoAux.RowCount > 0 Then
                            bFind = True
                        End If
                        RdoAux.Close
                    End If
                    
                    If nNumDoc > 0 Then
                        If sCodMov = "06" Or (sCodMov = "17" And bFind = False) Then
                            sql = "insert importacao_banco(codigo_banco,data_credito,nome_arquivo,data_pagamento,valor_pago,numero_documento,agencia,situacao_retorno,data_vencimento,convenio,conta_corrente,codigo_banco_receptor,retorno) values("
                            sql = sql & nCodBanco & ",'" & Format(sDataCredito, "mm/dd/yyyy") & "','" & sFileTitle & "','" & Format(sDataPag, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(nValorPago)) & ","
                            sql = sql & nNumDoc & ",'" & sAgencia & "','" & sSitRetorno & "','" & Format(sDataVencto, "mm/dd/yyyy") & "','" & sConvenio & "','" & sConta & "'," & nCodBancoReceptor & "," & nRetorno & ")"
                            cn.Execute sql, rdExecDirect
                            nRegImportado = nRegImportado + 1
                        End If
                    Else
                        DoEvents
                    End If
                End If
            End If
        End If
    End If
    nPos = nPos + 1
Loop
TS.Close
If IsDate(sDataCredito) Then
    Grava_Arquivo_Banco sFileName, sFileTitle, sDataCredito, nCodBanco, Val(sAgencia), False
    sql = "update bankfileheader set data_credito='" & Format(sDataCredito, "mm/dd/yyyy") & "',ready=1 where id='" & sGuid & "'"
    'sql = "INSERT INTO BankFile (data_gravado, FileName, FileData) SELECT '" & Format(sDataCredito, "mm/dd/yyyy") & "','" & sFileTitle & "',BulkColumn FROM OPENROWSET(BULK N'" & sFileName & "', SINGLE_BLOB) AS FileData;"
    cn.Execute sql
    txtResult.Text = txtResult.Text & vbCrLf & "Incluído o arquivo " & sFileName & " - Cobrança (" & sNomeBanco & ")"
Else
    txtResult.Text = txtResult.Text & vbCrLf & "Arquivo ignorado " & sFileName & " - Cobrança (" & sNomeBanco & ")"
End If
Exit Sub


LEDEBAUT:
nPos = 1
Set TS = fso.OpenTextFile(sFileName, ForReading)
Do Until TS.AtEndOfStream
    If nPos Mod 10 = 0 Then CallPb nPos, nTot
    row = TS.ReadLine
    If nPos = 1 Then
        nCodBanco = Mid(row, 43, 3)
        sNomeBanco = Trim(Mid(row, 46, 20))
    Else
        sTipoReg = Left(row, 1)
        If sTipoReg = "F" Then
            sDataCredito = ConvDataSerial(Mid(row, 45, 8))
            z = BinarySearchString(aDataCredito, sDataCredito)
            If z = -1 Then
                ReDim Preserve aDataCredito(UBound(aDataCredito) + 1)
                aDataCredito(UBound(aDataCredito)) = sDataCredito
                sql = "select * from importacao_banco where codigo_banco=" & nCodBanco & " and data_credito='" & Format(sDataCredito, "mm/dd/yyyy") & "' and nome_arquivo='" & sFileTitle & "'"
                Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                If RdoAux.RowCount > 0 Then
                    GoTo ARQUIVO_EXISTENTE
                End If
                RdoAux.Close
            End If
            sDataPag = sDataCredito
            sDataVencto = sDataCredito
            nCodReduz = Val(Mid(row, 11, 6))
            nValorPago = CDbl(Mid(row, 53, 15)) / 100
            sAgencia = Mid(row, 27, 4)
            sContaDeposito = Trim(Mid(row, 31, 9))
            nCodBancoReceptor = nCodBanco
            nRetorno = Val(Mid(row, 214, 2))
            
            Select Case nRetorno
                Case "00"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "Débito Efetuado"
                Case "01"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "Insuficiência de Fundos"
                Case "02"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "Conta Corrente não Cadastrada"
                Case "04"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "Outras Restrições"
                Case "10"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "Agência em Regime de Encerramento"
                Case "12"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "Valor Inválido"
                Case "13"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "Data de Lançamento inválida"
                Case "14"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "Agência Inválida"
                Case "15"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "DAC da conta corrente inválido"
                Case "18"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "Data do Débito anterior ao do processamento"
                Case "30"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "Sem contrato de débito automático"
                Case "96"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "Manutenção do Cadastro"
                Case "97"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "Cancelamento - Não Encontrado"
                Case "98"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "Cancelamento - não efetuado, fora de tempo habil"
                Case "99"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "Cancelamento - cancelado conforme solicitado"
                Case Else
                       sSitRetorno = Format(nRetorno, "00") & " - " & "Erro Indefinido"
            End Select
            
            sql = "SELECT lancamento.descreduz, debitoparcela.statuslanc, situacaolancamento.descsituacao, debitoparcela.datavencimento, debitoparcela.datadebase,"
            sql = sql & "debitoparcela.codreduzido, debitoparcela.anoexercicio, debitoparcela.codlancamento, debitoparcela.seqlancamento, debitoparcela.numparcela,"
            sql = sql & "debitoparcela.CODCOMPLEMENTO , parceladocumento.NumDocumento FROM lancamento INNER JOIN debitoparcela ON lancamento.codlancamento = debitoparcela.codlancamento INNER JOIN "
            sql = sql & "situacaolancamento ON debitoparcela.statuslanc = situacaolancamento.codsituacao INNER JOIN parceladocumento ON debitoparcela.codreduzido = parceladocumento.codreduzido AND "
            sql = sql & "debitoparcela.anoexercicio = parceladocumento.anoexercicio AND debitoparcela.codlancamento = parceladocumento.codlancamento AND "
            sql = sql & "debitoparcela.seqlancamento = parceladocumento.seqlancamento AND debitoparcela.numparcela = parceladocumento.numparcela AND debitoparcela.CODCOMPLEMENTO = parceladocumento.CODCOMPLEMENTO "
            sql = sql & "WHERE (DEBITOPARCELA.SEQLANCAMENTO=0) AND (DEBITOPARCELA.CODREDUZIDO = " & nCodReduz & ") AND (DEBITOPARCELA.CODLANCAMENTO = 1) AND "
            sql = sql & "(DEBITOPARCELA.NUMPARCELA > 0) AND (DEBITOPARCELA.DATAVENCIMENTO = '" & Format(sDataVencto, "mm/dd/yyyy") & "')"
            'Sql = Sql & "(DEBITOPARCELA.NUMPARCELA > 0) AND (MONTH(DEBITOPARCELA.DATAVENCIMENTO) = '11') AND (YEAR(DEBITOPARCELA.DATAVENCIMENTO) = '2024')"
            Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount = 0 Then
                   ' MsgBox "Código não encontrado nos optantes de débito automático (" & nCodReduz & ")"
                    sql = "delete from importacao_banco where codigo_banco=" & nCodBanco & " and data_credito='" & Format(sDataCredito, "mm/dd/yyyy") & " ' and nome_arquivo='" & sFileName & "' and data_controle is null"
                   ' cn.Execute Sql, rdExecDirect
                    txtResult.Text = txtResult.Text & vbCrLf & "O arquivo " & sFileName & " - Déb.Automático (" & sNomeBanco & ") não foi incluido pois contêm erros."
                    GoTo Close_DA
                Else
                    nNumDoc = !NumDocumento
                End If
               .Close
            End With
            If nNumDoc > 0 Then
                sql = "insert importacao_banco(codigo_banco,data_credito,nome_arquivo,data_pagamento,valor_pago,numero_documento,debito_automatico,agencia,data_vencimento,conta_deposito,codigo_reduzido,codigo_banco_receptor,retorno) values("
                sql = sql & nCodBanco & ",'" & Format(sDataCredito, "mm/dd/yyyy") & "','" & sFileTitle & "','" & Format(sDataPag, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(nValorPago)) & ","
                sql = sql & nNumDoc & ",1" & ",'" & sAgencia & "','" & Format(sDataVencto, "mm/dd/yyyy") & "','" & sContaDeposito & "'," & nCodReduz & "," & Val(nCodBancoReceptor) & "," & nRetorno & ")"
                cn.Execute sql, rdExecDirect
                nRegImportado = nRegImportado + 1
            End If
        ElseIf sTipoReg = "B" Then
            nCodReduz = Val(Mid(row, 11, 6))
            sAgencia = Mid(row, 27, 4)
            sConta = Mid(row, 31, 14)
            sDataOpcao = ConvDataSerial(Mid(row, 45, 8))
            nCodMov = Mid(row, 150, 1)
           sDataCredito = ConvDataSerial(Mid(row, 45, 8))
'            Sql = "SELECT * FROM DEBITOAUTOMATICO WHERE CODREDUZ=" & nCodReduz & " AND CODBANCO=" & nCodBanco
'            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'            If RdoAux.RowCount > 0 Then
               If nCodMov = 1 Then 'EXCLUSAO
                  If CDate(sDataOpcao) > Format(RdoAux!DataOpcao, "dd/mm/yyyy") Then
                      sql = "DELETE FROM DEBITOAUTOMATICO WHERE CODREDUZ=" & nCodReduz & " AND CODBANCO=" & nCodBanco
                      cn.Execute sql, rdExecDirect
                      nContN = nContN + 1
                  Else
                      nContI = nContI + 1
                  End If
               ElseIf nCodMov = 2 Then 'INCLUSAO
                  sql = "DELETE FROM DEBITOAUTOMATICO WHERE CODREDUZ=" & nCodReduz
                  cn.Execute sql, rdExecDirect
                   
                  sql = "INSERT DEBITOAUTOMATICO(CODREDUZ,CODBANCO,CODAGENCIA,NUMEROCONTA,DATAOPCAO,CODIGOPREF) VALUES("
                  sql = sql & nCodReduz & "," & nCodBanco & "," & sAgencia & "," & sConta & ",'" & Format(sDataOpcao, "mm/dd/yyyy") & "'," & nCodReduz & ")"
                  cn.Execute sql, rdExecDirect
                  nContP = nContP + 1
               End If
 '           Else
 '              nContI = nContI + 1
'            End If
'               dDataOpcao = Format(RdoAux!DataOpcao, "dd/mm/yyyy")
'            Else
 '           End If
  '          RdoAux.Close
        
            
        End If
    
    End If
    nPos = nPos + 1
Loop
txtResult.Text = txtResult.Text & vbCrLf & "Incluído o arquivo " & sFileName & " - Déb.Automático (" & sNomeBanco & ")"
If nContP > 0 Or nContI > 0 Or nContN > 0 Then
    txtResult.Text = txtResult.Text & vbCrLf & nContP & " Optantes incluidos, " & nContN & " Optantes removidos e " & nContI & " Optantes ignorados."
End If
Close_DA:
TS.Close
Grava_Arquivo_Banco sFileName, sFileTitle, sDataCredito, nCodBanco, Val(sAgencia), False
sql = "update bankfileheader set data_credito='" & Format(sDataCredito, "mm/dd/yyyy") & "',ready=1 where id='" & sGuid & "'"
'sql = "INSERT INTO BankFile (data_gravado, FileName, FileData) SELECT '" & Format(sDataCredito, "mm/dd/yyyy") & "','" & sFileTitle & "',BulkColumn FROM OPENROWSET(BULK N'" & sFileName & "', SINGLE_BLOB) AS FileData;"
cn.Execute sql
Exit Sub

ARQUIVO_EXISTENTE:
txtResult.Text = txtResult.Text & vbCrLf & "O arquivo " & sFileName & " " & sNomeBanco & " já está baixado, portanto foi ignorado."
'Resume Next
Exit Sub

Erro:
If Err.Number = 62 Then
    txtResult.Text = txtResult.Text & vbCrLf & "O arquivo " & sFileName & " " & sNomeBanco & " é inválido, portanto foi ignorado."
    Resume Next
Else
    If rdoErrors.Count > 1 Then
        If rdoErrors(1).Number = 2627 Then 'duplicidade
            Resume Next
        End If
    Else
        MsgBox Err.Description
        Resume Next
    End If
End If
End Sub

Private Function GravaSimples(sCnpj As String, sDataVencto As String, nAno As Integer, nMes As Integer, nValor As Double, sArquivo As String, sDataCredito As String, nCodBanco As Integer, sAgencia As String) As Long
Dim nNumDoc As Long, sql As String, RdoAux As rdoResultset, nCodReduz As Long, nNumParc As Integer, bAchou As Boolean, nSeq As Integer, nCompl As Integer, dDataVencto As Date, RdoAux3 As rdoResultset, RdoAux4 As rdoResultset
On Error GoTo Erro
DoEvents

'BUSCA CÓDIGO
sql = "SELECT CODIGOMOB,CNPJ FROM MOBILIARIO WHERE DATAENCERRAMENTO IS NULL and CONVERT(BIGINT, cnpj) = " & Val(sCnpj)
sql = sql & " OR CNPJ='" & Format(sCnpj, "00\.000\.000/0000-00") & "' AND DATAENCERRAMENTO IS NULL ORDER BY CODIGOMOB DESC"
Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
    If .RowCount > 0 Then
        nCodReduz = !codigomob
        .Close
    Else
        .Close
        sql = "SELECT CODCIDADAO,CNPJ FROM CIDADAO WHERE CNPJ = '" & RetornaNumero(sCnpj) & "' OR "
        sql = sql & "CNPJ='" & Format(sCnpj, "00\.000\.000/0000-00") & "' ORDER BY CODCIDADAO DESC"
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                nCodReduz = !CodCidadao
            Else
                'CNPJ NÃO LOCALIZADO
                sql = "SELECT * FROM SIMPLESCNPJ WHERE CNPJ='" & sCnpj & "' AND ANOCOMP=" & nAno & " AND MESCOMP=" & nMes
                Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                If RdoAux3.RowCount = 0 Then
                    sql = "INSERT SIMPLESCNPJ (CNPJ,ARQUIVOSHORT,BANCO,DATAARRECADA,DATAVENCTO,ANOCOMP,MESCOMP,PRINCIPAL,JUROS,"
                    sql = sql & "MULTA,AGENCIA,CODREDUZIDO) VALUES('" & RetornaNumero(sCnpj) & "','" & sArquivo & "'," & nCodBanco & ",'" & Format(sDataCredito, "mm/dd/yyyy") & "','"
                    sql = sql & Format(sDataVencto, "mm/dd/yyyy") & "'," & nAno & "," & nMes & "," & Virg2Ponto(CStr(nValor)) & "," & Virg2Ponto(0) & "," & Virg2Ponto(0) & ",'"
                    sql = sql & sAgencia & "'," & 0 & ")"
   '                 cn.Execute Sql, rdExecDirect
                End If
                RdoAux3.Close
                GoTo Fim
            End If
        End With
    End If
End With

'BUSCA LANCAMENTO
sql = "SELECT debitoparcela.codreduzido, debitoparcela.anoexercicio, debitoparcela.codlancamento, debitoparcela.seqlancamento, debitoparcela.numparcela,debitoparcela.CODCOMPLEMENTO, debitoparcela.DataVencimento, debitoparcela.statuslanc, debitotributo.ValorTributo "
sql = sql & "FROM debitoparcela INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
sql = sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND debitoparcela.numparcela = debitotributo.numparcela AND debitoparcela.codcomplemento = debitotributo.codcomplemento Where (debitoparcela.CODREDUZIDO = " & nCodReduz & ") And (debitoparcela.CodLancamento = 5) And (Month(debitoparcela.DataVencimento) = " & Month(CDate(sDataVencto)) & ") And "
sql = sql & "(YEAR(debitoparcela.datavencimento) = " & Year(CDate(sDataVencto)) & ") AND (debitotributo.codtributo = 13) and debitotributo.valortributo =" & Virg2Ponto(CStr(nValor)) & " AND statuslanc<>6 "
Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux3
   'EXISTE LANCAMENTO NESTE MÊS/ANO?
    If .RowCount > 0 Then 'SIM
'        nNumDoc = !NumDocumento
       
        nNumParc = !NumParcela 'CAPTURA A PARCELA
        bAchou = False
       'TEM ALGUMA QUE NÃO ESTA PAGA?
        Do Until .EOF
            If !statuslanc = 3 Then
                bAchou = True
                Exit Do
            End If
           .MoveNext
        Loop
        
       'SE ACHOU PEGA A PARCELA
        If bAchou Then
            nSeq = !SeqLancamento
            nCompl = !CODCOMPLEMENTO '---------------> PARCELA PRONTA PARA USO
            nNumDoc = 0
            GoTo Fim
        Else
           'SE NÃO ACHAR
           .MoveFirst
            nCompl = 0
           'BUSCAR A ÚLTIMA SEQUENCIA DE LANCAMENTO PARA EVITAR DUPLICIDADE
            sql = "SELECT MAX(SEQLANCAMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE (codreduzido = " & nCodReduz & ") AND anoexercicio=" & nAno & " and (codlancamento = 5)"
            Set RdoAux4 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux4
                If IsNull(!maximo) Then
                    nSeq = 0
                Else
                    nSeq = !maximo + 1
                End If
               .Close
            End With
        End If
    
    Else
       'NÃO ACHOU LANCAMENTOS NESTE MÊS/ANO
       'AUMENTA O LANCAMENTO
        sql = "SELECT MAX(SEQLANCAMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE (codreduzido = " & nCodReduz & ") AND (codlancamento = 5) AND (ANOEXERCICIO = " & nAno & ")"
        Set RdoAux4 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux4
            If IsNull(!maximo) Then
                nSeq = 1
            Else
                nSeq = !maximo + 1
            End If
           .Close
        End With
         
        nCompl = 0
        nNumParc = 1
    End If
 
   .Close
End With


sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & 5 & " AND "
sql = sql & "SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nNumParc & " AND CODCOMPLEMENTO=" & nCompl
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurReadOnly)
If RdoAux.RowCount = 0 Then
    'CRIA A PARCELA
    sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO where numdocumento<2000000"
    Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurReadOnly)
    nNumDoc = RdoAux!maximo + 1
    'RdoAux.Close
    
    sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC,ISENTOMJ,emissor) VALUES("
    sql = sql & nNumDoc & ",'" & Format(Now, "mm/dd/yyyy") & "'," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & ",'" & NomeDeLogin & " (BAIXA BANCÁRIA/SN)" & "')"
 '   cn.Execute Sql, rdExecDirect
    
    sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,CODMOEDA,USERID) "
    sql = sql & "VALUES(" & nCodReduz & "," & nAno & "," & 5 & "," & nSeq & "," & nNumParc & "," & nCompl & ","
    sql = sql & 3 & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "'," & 1 & "," & RetornaUsuarioID(NomeDeLogin) & ")"
  '  cn.Execute Sql, rdExecDirect
    
    sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
    sql = sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
    sql = sql & nCodReduz & "," & nAno & "," & 5 & "," & nSeq & ","
    sql = sql & nNumParc & "," & nCompl & "," & 13 & "," & Virg2Ponto(CStr(nValor)) & ")"
 '   cn.Execute Sql, rdExecDirect
    
'    Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
'    Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & nCodReduz & ","
'    Sql = Sql & nAno & "," & 5 & "," & nSeq & "," & nNumParc & ","
'    Sql = Sql & nCompl & "," & nNumDoc & ")"
'    cn.Execute Sql, rdExecDirect


End If
RdoAux.Close

'PROCURA SE O DEBITO JA FOI BAIXADO
sql = "SELECT * FROM COMPLEMENTOSIMPLES WHERE ARQUIVOBANCO='" & sArquivo & "' AND DATACREDITO='" & Format(sDataCredito, "mm/dd/yyyy") & "' AND "
sql = sql & "CNPJ='" & sCnpj & "' AND ANO=" & nAno & " AND MES=" & nMes
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount = 0 Then
    sql = "INSERT COMPLEMENTOSIMPLES(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,ARQUIVOBANCO,DATACREDITO,VALOR,CNPJ,ANO,MES) VALUES(" & nCodReduz & ","
    sql = sql & nAno & "," & 5 & "," & nSeq & "," & nNumParc & "," & nCompl & ",'" & sArquivo & "','" & Format(sDataCredito, "mm/dd/yyyy") & "',"
    sql = sql & Virg2Ponto(CStr(nValor)) & ",'" & sCnpj & "'," & nAno & "," & nMes & ")"
'    cn.Execute Sql, rdExecDirect
    RdoAux.Close
End If
Fim:

nCodigoSimples = nCodReduz
GravaSimples = nNumDoc

Exit Function
Erro:
MsgBox Err.Description
Resume Next
End Function

Private Function ConvDataSerial(sData As String) As String
If Len(sData) = 8 Then
   ConvDataSerial = Right$(sData, 2) & "/" & Mid$(sData, 5, 2) & "/" & Left$(sData, 4)
Else
   ConvDataSerial = Left$(sData, 2) & "/" & Mid$(sData, 3, 2) & "/20" & Right$(sData, 2)
End If
End Function

Private Function ConvDataSerialBB(sData As String) As String
   ConvDataSerialBB = Left(sData, 2) & "/" & Mid$(sData, 3, 2) & "/" & Right(sData, 4)
End Function

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
DoEvents

End Sub

Private Function FileRowCount(sFile As String)
    Const BufSize As Long = 100000
    Dim T0 As Single
    Dim LfAnsi As String
    Dim f As Integer
    Dim FileBytes As Long
    Dim BytesLeft As Long
    Dim Buffer() As Byte
    Dim strBuffer As String
    Dim BufPos As Long
    Dim LineCount As Long

    T0 = Timer()
    LfAnsi = StrConv(vbLf, vbFromUnicode)
    f = FreeFile(0)
    Open sFile For Binary Access Read As #f
    FileBytes = LOF(f)
    ReDim Buffer(BufSize - 1)
    BytesLeft = FileBytes
    Do Until BytesLeft = 0
        If BufPos = 0 Then
            If BytesLeft < BufSize Then ReDim Buffer(BytesLeft - 1)
            Get #f, , Buffer
            strBuffer = Buffer 'Binary copy of bytes.
            BytesLeft = BytesLeft - LenB(strBuffer)
            BufPos = 1
        End If
        Do Until BufPos = 0
            BufPos = InStrB(BufPos, strBuffer, LfAnsi)
            If BufPos > 0 Then
                LineCount = LineCount + 1
                BufPos = BufPos + 1
            End If
        Loop
    Loop
    Close #f
    FileRowCount = LineCount
    'Add 1 to LineCount if last line of your files do not
    'have a trailing CrLf.
  '  MsgBox "Counted " & Format$(LineCount, "#,##0") & " lines in" & vbNewLine _
         & Format$(FileBytes, "#,##0") & " bytes of text." & vbNewLine _
         & Format$(Timer() - T0, "0.0#") & " seconds."
End Function

Function GetFileNameFromPath(strFullPath As String) As String
    GetFileNameFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))
End Function

Private Sub Grava_Arquivo_Banco(sFullPath As String, sNomeArq As String, sDataCredito As String, nCodBanco As Integer, nCodAgencia As Integer, bDA As Boolean)
Dim nSeq As Integer, sql As String, RdoAux3 As rdoResultset, RdoAux As rdoResultset, bFind As Boolean

'MsgBox "Tentando gravar o arquivo na pasta " & sPathArqBanco & "\" & Right$(sDataCredito, 4)

'cria os diretorios
If Dir$(sPathArqBanco & "\" & Right$(sDataCredito, 4), vbDirectory) = "" Then
   'cria o ano
   MkDir sPathArqBanco & "\" & Right$(sDataCredito, 4)
End If
If Dir$(sPathArqBanco & "\" & Right$(sDataCredito, 4) & "\" & Mid$(sDataCredito, 4, 2), vbDirectory) = "" Then
  'cria o mes
   MkDir sPathArqBanco & "\" & Right$(sDataCredito, 4) & "\" & Mid$(sDataCredito, 4, 2)
End If
If Dir$(sPathArqBanco & "\" & Right$(sDataCredito, 4) & "\" & Mid$(sDataCredito, 4, 2) & "\" & Left$(sDataCredito, 2), vbDirectory) = "" Then
  'cria o dia
   MkDir sPathArqBanco & "\" & Right$(sDataCredito, 4) & "\" & Mid$(sDataCredito, 4, 2) & "\" & Left$(sDataCredito, 2)
End If
If Dir$(sPathArqBanco & "\" & Right$(sDataCredito, 4) & "\" & Mid$(sDataCredito, 4, 2) & "\" & Left$(sDataCredito, 2) & "\" & sNomeArq) = "" Then
   FileCopy sFullPath, sPathArqBanco & "\" & Right$(sDataCredito, 4) & "\" & Mid$(sDataCredito, 4, 2) & "\" & Left$(sDataCredito, 2) & "\" & sNomeArq
End If


'Sql = "SELECT * FROM ARQUIVOBANCO WHERE DATACREDITO='" & Format(CDate(sDataCredito), "mm/dd/yyyy") & "' AND NOMEARQ='" & sNomeArq & "'"
sql = "SELECT * FROM ARQUIVOBANCO WHERE DATACREDITO='" & Format(CDate(sDataCredito), "mm/dd/yyyy") & "'"
Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
If RdoAux3.RowCount = 0 Then
   nSeq = 1
   RdoAux3.Close
Else
   sql = "SELECT MAX(SEQ) AS MAXIMO FROM ARQUIVOBANCO WHERE DATACREDITO='" & Format(CDate(sDataCredito), "mm/dd/yyyy") & "'"
   Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
   nSeq = RdoAux3!maximo + 1
   RdoAux3.Close
End If
sql = "INSERT ARQUIVOBANCO(DATACREDITO,SEQ,CODBANCO,CODAGENCIA,DATAINCLUSAO,NOMEARQ,DA) VALUES('"
sql = sql & Format(CDate(sDataCredito), "mm/dd/yyyy") & "'," & nSeq & "," & nCodBanco & "," & nCodAgencia & ",'"
sql = sql & Format(Now, "mm/dd/yyyy") & "','" & sNomeArq & "'," & IIf(bDA, 1, 0) & ")"
cn.Execute sql, rdExecDirect


sql = "SELECT DISTINCT data_credito FROM importacao_banco WHERE Nome_Arquivo ='" & sNomeArq & "'"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If !data_credito <> sDataCredito Then
        
            sql = "SELECT * FROM ARQUIVOBANCO WHERE DATACREDITO='" & Format(!data_credito, "mm/dd/yyyy") & "' AND NOMEARQ='" & sNomeArq & "'"
            Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If RdoAux3.RowCount = 0 Then
                RdoAux3.Close
                sql = "SELECT MAX(SEQ) AS MAXIMO FROM ARQUIVOBANCO WHERE DATACREDITO='" & Format(CDate(!data_credito), "mm/dd/yyyy") & "' AND NOMEARQ='" & sNomeArq & "'"
                Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                nSeq = RdoAux3!maximo + 1
                RdoAux3.Close
                
                
                sql = "INSERT ARQUIVOBANCO(DATACREDITO,SEQ,CODBANCO,CODAGENCIA,DATAINCLUSAO,NOMEARQ,DA) VALUES('"
                sql = sql & Format(!data_credito, "mm/dd/yyyy") & "'," & nSeq & "," & nCodBanco & "," & nCodAgencia & ",'"
                sql = sql & Format(Now, "mm/dd/yyyy") & "','" & sNomeArq & "'," & IIf(bDA, 1, 0) & ")"
                cn.Execute sql, rdExecDirect
            End If
        End If
       .MoveNext
    Loop
   .Close
End With



End Sub
