VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmEmissaoLivroDivida 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão de Livro de dívida ativa"
   ClientHeight    =   2370
   ClientLeft      =   7455
   ClientTop       =   5025
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2370
   ScaleWidth      =   5475
   Begin VB.ListBox lista 
      Appearance      =   0  'Flat
      Height          =   1005
      ItemData        =   "frmEmissaoLivroDivida.frx":0000
      Left            =   270
      List            =   "frmEmissaoLivroDivida.frx":0007
      TabIndex        =   1
      Top             =   495
      Width           =   2940
   End
   Begin MSComctlLib.ProgressBar Pb 
      Height          =   225
      Left            =   510
      TabIndex        =   2
      Top             =   1845
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin prjChameleon.chameleonButton btExecutar 
      Height          =   360
      Left            =   3510
      TabIndex        =   5
      ToolTipText     =   "Emitir Livro"
      Top             =   1170
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "&Emitir Livro(s)"
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
      MICON           =   "frmEmissaoLivroDivida.frx":0023
      PICN            =   "frmEmissaoLivroDivida.frx":003F
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
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 %"
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   1
      Left            =   180
      TabIndex        =   4
      Top             =   1845
      Width           =   270
   End
   Begin VB.Label lblPB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100 %"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2970
      TabIndex        =   3
      Top             =   1860
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Selecione o Tipo de Livro"
      Height          =   195
      Index           =   0
      Left            =   270
      TabIndex        =   0
      Top             =   225
      Width           =   2085
   End
End
Attribute VB_Name = "frmEmissaoLivroDivida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Endereco
    nCodReduz As Long
    sNome As String
    sEndereco As String
    sCompl As String
    sCep As String
    sBairro As String
    sCidade As String
    sUF As String
End Type

Private Sub Form_Load()
Centraliza Me
lista.ListIndex = 0
End Sub

Private Sub CallPb(nPosF As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents

If ((nPosF * 100) / nTotal) <= 100 Then
   Pb.value = (nPosF * 100) / nTotal
Else
   Pb.value = 100
End If
lblPB.Caption = FormatNumber(Pb.value, 2) & " %"

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub

Private Sub btExecutar_Click()

If lista.ListIndex = 0 Then
    Livro_Parcelamento
End If

End Sub


Private Sub Livro_Parcelamento()
Dim sql As String, RdoAux As rdoResultset, aCod() As Long, nPos As Long, nTot As Long, nCodReduz As Long, RdoEnd As rdoResultset
Dim qd As New rdoQuery, aEndereco() As Endereco, xImovel As clsImovel

ReDim aCod(0): ReDim aEndereco(0)
Set qd.ActiveConnection = cn
qd.QueryTimeout = 180
Set xImovel = New clsImovel

sql = "delete from DIVIDATIVA where usuario='" & NomeDeLogin & "'"
cn.Execute sql, rdExecDirect


Ocupado
btExecutar.Enabled = False
Me.Refresh

sql = "SELECT DISTINCT CODREDUZIDO FROM debitoparcela WHERE codlancamento=20 AND statuslanc IN (3,18) ORDER BY codreduzido"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aCod(UBound(aCod) + 1)
        aCod(UBound(aCod)) = !CODREDUZIDO
       .MoveNext
    Loop
   .Close
End With

nTot = UBound(aCod)
For nPos = 1 To UBound(aCod)
    If nPos Mod 10 = 0 Then
        CallPb nPos, nTot
    End If
    nCodReduz = aCod(nPos)
    
    ReDim Preserve aEndereco(UBound(aEndereco) + 1)
    With xImovel
        If nCodReduz < 100000 Then
           .RetornaEndereco nCodReduz, Imobiliario, Localizacao
            sql = "SELECT proprietario.codcidadao ,cidadao.nomecidadao FROM dbo.proprietario INNER JOIN dbo.cidadao ON proprietario.codcidadao = cidadao.codcidadao WHERE codreduzido=" & nCodReduz & " AND tipoprop='P' AND principal=1"
            Set RdoEnd = cn.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
            If RdoEnd.RowCount > 0 Then
                sProp = RdoEnd!Nomecidadao
                RdoEnd.Close
            End If
        ElseIf nCodReduz >= 100000 And nCodReduz < 500000 Then
           .RetornaEndereco nCodReduz, Mobiliario, Localizacao
            sql = "SELECT RAZAOSOCIAL FROM MOBILIARIO WHERE CODIGOMOB=" & nCodReduz
            Set RdoEnd = cn.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
            If RdoEnd.RowCount > 0 Then
                sProp = RdoEnd!RazaoSocial
                RdoEnd.Close
            End If
        Else
            sql = "SELECT NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO=" & nCodReduz
            Set RdoEnd = cn.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
            If RdoEnd.RowCount > 0 Then
                sProp = RdoEnd!Nomecidadao
                RdoEnd.Close
            End If
           .RetornaEndereco nCodReduz, cidadao, cadastrocidadao
        End If
    
        ReDim Preserve aEndereco(UBound(aEndereco) + 1)
        
        aEndereco(UBound(aEndereco)).nCodReduz = nCodReduz
        aEndereco(UBound(aEndereco)).sNome = sProp
        aEndereco(UBound(aEndereco)).sEndereco = .Endereco & ", " & .Numero
        aEndereco(UBound(aEndereco)).sCompl = .Complemento
        aEndereco(UBound(aEndereco)).sCep = Format(.Cep, "00000-000")
        aEndereco(UBound(aEndereco)).sBairro = .Bairro
        aEndereco(UBound(aEndereco)).sCidade = .Cidade
        aEndereco(UBound(aEndereco)).sUF = .UF
        endPos = UBound(aEndereco)
        
    End With
    
    
    nCodReduz = aCod(nPos)
    On Error Resume Next
    RdoAux.Close
    On Error GoTo 0
    qd.sql = "{ Call spEXTRATOPARCELAMENTOLIVRO(?,?) }"
    qd(0) = nCodReduz
    qd(1) = nCodReduz
    Set RdoAux = qd.OpenResultset(rdOpenKeyset)
    With RdoAux
        Do Until .EOF

            sql = "INSERT DIVIDATIVA (USUARIO,NUMLIVRO,TIPOLIVRO,ANOLIVRO,CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,DATAVENCIMENTO,"
            sql = sql & "CODTRIBUTO,VALORTRIBUTO,VALORJUROS,VALORMULTA,VALORCORRECAO,PROPRIETARIO,ENDERECO,COMPLEMENTO,CEP,BAIRRO,CIDADE,UF,DESCTRIBUTO,PROCESSO) VALUES('"
            sql = sql & NomeDeLogin & " '," & !livro & "," & 0 & "," & 0 & "," & !CODREDUZIDO & "," & !AnoExercicio & "," & !CodLancamento & "," & !SeqLancamento & "," & !NumParcela & "," & !CODCOMPLEMENTO & ",'"
            sql = sql & Format(!DataVencimento, "mm/dd/yyyy") & "'," & !CodTributo & "," & Virg2Ponto(!ValorTributo) & "," & Virg2Ponto(CStr(!ValorJuros)) & "," & Virg2Ponto(CStr(!ValorMulta)) & "," & Virg2Ponto(CStr(!ValorCorrecao)) & ",'"
            sql = sql & Mask(Left$(aEndereco(endPos).sNome, 40)) & "','" & Left$(Mask(aEndereco(endPos).sEndereco), 50) & "','" & Left$(Mask(aEndereco(endPos).sCompl), 35) & "','"
            sql = sql & aEndereco(endPos).sCep & "','" & Left(aEndereco(endPos).sBairro, 40) & "','" & Mask(aEndereco(endPos).sCidade) & "','" & aEndereco(endPos).sUF & "','" & Mask(!abrevTributo) & "','" & SubNull(!NumProcesso) & "')"
            cn.Execute sql, rdExecDirect



           .MoveNext
        Loop
       .Close
    End With
Next

Liberado
btExecutar.Enabled = True
End Sub

