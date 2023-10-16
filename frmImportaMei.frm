VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportaMei 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importação e geração de débitos do MEI"
   ClientHeight    =   2655
   ClientLeft      =   10860
   ClientTop       =   8790
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2655
   ScaleWidth      =   8355
   Begin VB.ListBox lstFileName 
      Enabled         =   0   'False
      Height          =   450
      Left            =   1170
      TabIndex        =   6
      Top             =   3555
      Visible         =   0   'False
      Width           =   1275
   End
   Begin Tributacao.jcFrames frProgress 
      Height          =   1155
      Left            =   2025
      Top             =   720
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
         TabIndex        =   0
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
      Begin VB.Label lblFileNumber 
         BackStyle       =   0  'Transparent
         Caption         =   "Importando Arquivo 0 de 0"
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   90
         Width           =   4305
      End
      Begin VB.Label lblFileName 
         BackStyle       =   0  'Transparent
         Caption         =   "ARRECADA08.ret"
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   360
         Width           =   4155
      End
   End
   Begin VB.ListBox lstArq 
      Appearance      =   0  'Flat
      Height          =   1785
      Left            =   90
      TabIndex        =   3
      Top             =   300
      Width           =   8145
   End
   Begin prjChameleon.chameleonButton cmdArq 
      Height          =   360
      Left            =   6165
      TabIndex        =   4
      ToolTipText     =   "Selecione o arquivo a importar"
      Top             =   2205
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
      MICON           =   "frmImportaMei.frx":0000
      PICN            =   "frmImportaMei.frx":001C
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
      Left            =   360
      Top             =   3375
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Color           =   8388608
      DialogTitle     =   "Selecione o arquivo de GIA"
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   390
      Left            =   4140
      TabIndex        =   7
      ToolTipText     =   "Exibe os débitos importados"
      Top             =   2205
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   688
      BTYPE           =   3
      TX              =   "&Imprimir Relatório"
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
      MICON           =   "frmImportaMei.frx":00BA
      PICN            =   "frmImportaMei.frx":00D6
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
      Caption         =   "Arquivos Selecionados"
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   90
      Width           =   3345
   End
End
Attribute VB_Name = "frmImportaMei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type tReg
    sNome As String
    sCnpj As String
    sDataVencto As String
    sValor As String
    sCodigo As String
    sEndereco As String
    sNumero As String
    sCompl As String
    sBairro As String
    sCep As String
    sDuplicado As String
    sArquivo As String
End Type

Private Type tCnpj
    sCnpj As String
    sCodigo As String
End Type

Dim aReg() As tReg

Private Sub cmdArq_Click()

On Error GoTo Erro:

Dim vFiles As Variant
Dim lFile As Long, x As Integer
Dim itmX As ListItem
ReDim aReg(0)
lstArq.Clear
lstFileName.Clear
With cDialog
    
    .FileName = "" 'Clear the filename
    .CancelError = True
    .MaxFileSize = 30000
    .DialogTitle = "Select File(s)..."
    .flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNHideReadOnly 'Flags, allows Multi select, Explorer style and hide the Read only tag
    .Filter = "Text files (*.txt)|*.txt"
    .ShowOpen
    
    vFiles = Split(.FileName, Chr(0)) 'Splits the filename up in segments
    nQtdeArquivos = UBound(vFiles)
    
    If MsgBox("Deseja importar este(s) arquivo(s)?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
        Ocupado
        frProgress.Visible = True
        Me.Refresh
        If UBound(vFiles) = 0 Then ' If there is only 1 file then do this
            CallPb CLng(1), CLng(1)
            lblFileNumber = "Lendo Arquivo 1 de 1"
            Me.Refresh
            If ImportaArquivo(.FileName, .FileTitle) Then
                lstArq.AddItem .FileName
                lstFileName.AddItem .FileTitle
                nFileTot = 1
                nFilePos = 1
            Else
                MsgBox "Arquivo " & .FileName & " não é um arquivo válido!", vbCritical, "Atenção"
            End If
        Else
            For lFile = 1 To UBound(vFiles) ' More than 1 file then do this until there are no more files
                CallPb CLng(lFile), CLng(UBound(vFiles))
                lblFileName.Caption = lstFileName.List(lFile)
                lblFileNumber = "Lendo Arquivo " & lFile & " de " & UBound(vFiles)
                Me.Refresh
                If ImportaArquivo(vFiles(0) + "\" & vFiles(lFile), CStr(vFiles(lFile))) Then
                    lstArq.AddItem vFiles(0) + "\" & vFiles(lFile)
                    lstFileName.AddItem vFiles(lFile)
                    nFileTot = UBound(vFiles)
                    nFilePos = lFile
                Else
                    MsgBox "Arquivo " & (vFiles(0) + "\" & vFiles(lFile)) & " não é um arquivo válido!", vbCritical, "Atenção"
                End If
            Next
        End If
    End If
    Liberado
    frProgress.Visible = False
    cmdArq.Enabled = True

End With

If lstArq.ListCount > 0 Then
    For x = 1 To UBound(aReg)
       Set itmX = frmImportaMei2.lvMain.ListItems.Add(, , aReg(x).sCnpj)
       itmX.SubItems(1) = aReg(x).sNome
       itmX.SubItems(2) = aReg(x).sDataVencto
       itmX.SubItems(3) = aReg(x).sValor
       itmX.SubItems(4) = aReg(x).sCodigo
       itmX.SubItems(5) = aReg(x).sEndereco
       itmX.SubItems(6) = aReg(x).sNumero
       itmX.SubItems(7) = aReg(x).sCompl
       itmX.SubItems(8) = aReg(x).sBairro
       itmX.SubItems(9) = aReg(x).sCep
       itmX.SubItems(10) = aReg(x).sDuplicado
       itmX.SubItems(11) = aReg(x).sArquivo
    Next
    frmImportaMei2.show vbModal
End If
CallPb 100, 100
Exit Sub
Erro:
If Err.Number = 32755 Then
Else
    MsgBox Err.Description
End If

End Sub

Private Function ImportaArquivo(sFileName As String, sFileTitle As String) As Boolean
Dim fso As FileSystemObject, TS As TextStream, row As String, header As String, bValido As Boolean, sData As String
Dim sNome As String, sCnpj As String, sDataVencto As String, sValor As String, Sql As String, RdoAux As rdoResultset
Dim sCodigo As String, aCNPJ() As tCnpj, y As Integer, bFind As Boolean, sCPF As String, sEndereco As String, sNum As String
Dim sCompl As String, sBairro As String, sCep As String, sDuplicado As String

On Error GoTo Erro

ReDim aCNPJ(0)
Set fso = New FileSystemObject
Set TS = fso.OpenTextFile(sFileName, ForReading)
row = TS.ReadLine
header = Trim(Mid(row, 1, 12))
bValido = False
If Len(header) = 12 Then
    If Left(header, 1) = "0" Then
        sData = Mid(header, 11, 2) & "/" & Mid(header, 9, 2) & "/" & Mid(header, 5, 4)
        If IsDate(sData) Then
            bValido = True
        End If
    End If
End If

If bValido Then
    
    Do Until TS.AtEndOfStream
        row = TS.ReadLine
        If Left(row, 1) = "1" Then
            sNome = Trim(Mid(row, 534, 150))
            sCPF = Right(sNome, 11)
            sCnpj = Mid(row, 684, 14)
            sEndereco = Trim(LTrim(Mid(row, 713, 3)) & " " & Mid(row, 716, 60))
            sNum = Val(Mid(row, 776, 6))
            sCompl = Trim(Mid(row, 782, 156))
            sBairro = Trim(Mid(row, 938, 50))
            sCep = Mid(row, 1044, 8)
            DoEvents
            sDataVencto = Mid(row, 1064, 2) & "/" & Mid(row, 1062, 2) & "/" & Mid(row, 1058, 4)
            sCodigo = "000000"
            If Year(Now) - Val(Mid(row, 1058, 4)) <= 5 Then 'apenas ultimos 5 anos
              '  If sCnpj = "13633499000116" Then MsgBox "teste"
                For y = 0 To UBound(aCNPJ)
                    If aCNPJ(y).sCnpj = sCnpj Then
                        sCodigo = aCNPJ(y).sCodigo
                        bFind = True
                        Exit For
                    End If
                Next
                If Not bFind Or Val(sCodigo) = 0 Then
                    ReDim Preserve aCNPJ(UBound(aCNPJ) + 1)
                    aCNPJ(UBound(aCNPJ)).sCnpj = sCnpj
                    Sql = "select codigomob from mobiliario where cnpj='" & sCnpj & "'"
                    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    If RdoAux.RowCount > 0 Then
                        sCodigo = RdoAux!codigomob
                        RdoAux.Close
                    Else
                        If IsNumeric(sCPF) Then
                            Sql = "select codcidadao from cidadao where cpf='" & sCPF & "' or nomecidadao='" & sNome & "' or cnpj='" & sCnpj & "'"
                            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                            If RdoAux.RowCount > 0 Then
                                sCodigo = RdoAux!CodCidadao
                            End If
                            RdoAux.Close
                        End If
                    End If
                    aCNPJ(UBound(aCNPJ)).sCodigo = sCodigo
                End If
                sValor = Format(CDbl(Mid(row, 1071, 20)), "#0.00")
                
                Sql = "select codigo from importacao_mei where codigo=" & Val(sCodigo) & " and data_vencimento='" & Format(sDataVencto, "mm/dd/yyyy") & "'"
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                If RdoAux.RowCount = 0 Then
                    sDuplicado = "Não"
                Else
                    sDuplicado = "Sim"
                End If
                RdoAux.Close
                
                ReDim Preserve aReg(UBound(aReg) + 1)
                aReg(UBound(aReg)).sNome = sNome
                aReg(UBound(aReg)).sCnpj = sCnpj
                aReg(UBound(aReg)).sDataVencto = sDataVencto
                aReg(UBound(aReg)).sValor = sValor
                aReg(UBound(aReg)).sCodigo = sCodigo
                aReg(UBound(aReg)).sEndereco = sEndereco
                aReg(UBound(aReg)).sCompl = sCompl
                aReg(UBound(aReg)).sNumero = sNum
                aReg(UBound(aReg)).sBairro = sBairro
                aReg(UBound(aReg)).sCep = sCep
                aReg(UBound(aReg)).sDuplicado = sDuplicado
                aReg(UBound(aReg)).sArquivo = sFileTitle
            End If
        End If
    Loop
    TS.Close
End If


TS.Close
ImportaArquivo = bValido
Exit Function
Erro:
MsgBox Err.Description
Resume Next
End Function

Private Sub cmdPrint_Click()
frmReport.ShowReport3 "IMPORTAMEI", frmMdi.HWND, Me.HWND
End Sub

Private Sub Form_Load()
Centraliza Me
End Sub

Private Sub CallPb(nVal As Long, nTot As Long)
If nVal > 0 Then
    pBar.Color = &HC0C000
Else
    pBar.Color = vbWhite
End If
If ((nVal * 100) / nTot) <= 100 Then
   pBar.value = (nVal * 100) / nTot
Else
   pBar.value = 100
End If

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

End Sub


