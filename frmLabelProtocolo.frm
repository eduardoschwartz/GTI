VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmLabelProtocolo 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão de etiquetas do Protocolo"
   ClientHeight    =   6210
   ClientLeft      =   15810
   ClientTop       =   5715
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   6585
   Begin VB.CheckBox chkArquivado 
      Caption         =   "Somente não Arquivados"
      Height          =   195
      Left            =   3555
      TabIndex        =   16
      Top             =   90
      Width           =   2850
   End
   Begin VB.ComboBox cmbAssunto 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   720
      Width           =   3930
   End
   Begin VB.OptionButton optEtiq 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Mala Direta por assunto:"
      Height          =   240
      Index           =   2
      Left            =   90
      TabIndex        =   14
      Top             =   765
      Width           =   2130
   End
   Begin VB.TextBox txtAte 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   3
      Top             =   390
      Width           =   585
   End
   Begin VB.TextBox txtDe 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3540
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   2
      Top             =   390
      Width           =   585
   End
   Begin prjChameleon.chameleonButton cmdMalaDireta 
      Height          =   345
      Left            =   1200
      TabIndex        =   10
      ToolTipText     =   "Etiquetas para Mala Direta"
      Top             =   5340
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Mala Direta"
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
      MICON           =   "frmLabelProtocolo.frx":0000
      PICN            =   "frmLabelProtocolo.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.OptionButton optEtiq 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Exibir processos no período:"
      Height          =   240
      Index           =   1
      Left            =   90
      TabIndex        =   1
      Top             =   420
      Width           =   2895
   End
   Begin VB.OptionButton optEtiq 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Exibir apenas processos pendentes"
      Height          =   240
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   2895
   End
   Begin prjChameleon.chameleonButton cmdAbrir 
      Height          =   345
      Left            =   2610
      TabIndex        =   9
      ToolTipText     =   "Abrir processo selecionado"
      Top             =   5340
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Abrir"
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
      MICON           =   "frmLabelProtocolo.frx":007E
      PICN            =   "frmLabelProtocolo.frx":009A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdRemoveAll 
      Height          =   330
      Left            =   90
      TabIndex        =   7
      ToolTipText     =   "Remover Todos"
      Top             =   5340
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   582
      BTYPE           =   3
      TX              =   "-"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLabelProtocolo.frx":0121
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdAddAll 
      Height          =   330
      Left            =   570
      TabIndex        =   8
      ToolTipText     =   "Selecionar Todos"
      Top             =   5340
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   582
      BTYPE           =   3
      TX              =   "+"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLabelProtocolo.frx":013D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   345
      Left            =   5235
      TabIndex        =   5
      ToolTipText     =   "Sair da Tela"
      Top             =   5340
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Sair"
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
      MICON           =   "frmLabelProtocolo.frx":0159
      PICN            =   "frmLabelProtocolo.frx":0175
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
      Height          =   345
      Left            =   3915
      TabIndex        =   6
      ToolTipText     =   "Emitir as etiquetas selecionadas"
      Top             =   5340
      Width           =   1260
      _ExtentX        =   2223
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
      MICON           =   "frmLabelProtocolo.frx":01E3
      PICN            =   "frmLabelProtocolo.frx":01FF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView lvEtiq 
      Height          =   4110
      Left            =   45
      TabIndex        =   4
      Top             =   1125
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7250
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ANO"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nº"
         Object.Width           =   1589
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "REQUERENTE"
         Object.Width           =   4268
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ASSUNTO"
         Object.Width           =   4268
      EndProperty
   End
   Begin prjChameleon.chameleonButton cmdLoad 
      Height          =   315
      Left            =   5550
      TabIndex        =   13
      ToolTipText     =   "Sair da Tela"
      Top             =   360
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   556
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLabelProtocolo.frx":0359
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdCarta 
      Height          =   345
      Left            =   1200
      TabIndex        =   17
      ToolTipText     =   "Emissão de cartas"
      Top             =   5760
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Gera Cartas"
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
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmLabelProtocolo.frx":0375
      PICN            =   "frmLabelProtocolo.frx":0391
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdTramitacao 
      Height          =   345
      Left            =   2610
      TabIndex        =   18
      ToolTipText     =   "Incluir na tramitação"
      Top             =   5760
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Tramitação"
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
      MICON           =   "frmLabelProtocolo.frx":041E
      PICN            =   "frmLabelProtocolo.frx":043A
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
      Caption         =   "Até..:"
      Height          =   225
      Index           =   0
      Left            =   4350
      TabIndex        =   12
      Top             =   450
      Width           =   465
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "De..:"
      Height          =   225
      Index           =   13
      Left            =   3120
      TabIndex        =   11
      Top             =   420
      Width           =   465
   End
End
Attribute VB_Name = "frmLabelProtocolo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, sql As String, NodX As Object

Private Sub cmdAbrir_Click()
    Dim s As String
    If lvEtiq.ListItems.Count = 0 Then Exit Sub
    AnoProcesso = lvEtiq.SelectedItem.Text
    s = lvEtiq.SelectedItem.ListSubItems(1).Text
    CodProcesso = Val(Left$(s, Len(s) - 1))
    frmProcesso.show
    frmProcesso.ZOrder 0

End Sub

Private Sub cmdAddAll_Click()
For n = 1 To lvEtiq.ListItems.Count
    lvEtiq.ListItems(n).Checked = True
Next
End Sub

Private Sub cmdCarta_Click()
Dim x As Integer, nAno As Integer, nNumero As Integer, sProcesso As String, bFind As Boolean

bFind = False
For x = 1 To lvEtiq.ListItems.Count
    If lvEtiq.ListItems(x).Checked Then
        bFind = True
        Exit For
    End If
Next

If Not bFind Then
    MsgBox "Selecione ao menos um processo", vbCritical, "Erro"
    Exit Sub
End If

sql = "delete from carta_protocolo where nomelogin='" & NomeDeLogin & "'"
cn.Execute sql, rdExecDirect

For x = 1 To lvEtiq.ListItems.Count
    If lvEtiq.ListItems(x).Checked Then
        nAno = Val(lvEtiq.ListItems(x).Text)
        nNumero = CLng(Left$(lvEtiq.ListItems(x).SubItems(1), Len(lvEtiq.ListItems(x).SubItems(1)) - 2))
        sProcesso = Format(nNumero, "00000") & "-" & RetornaDVProcesso(CLng(nNumero)) & "/" & nAno
        sql = "insert carta_protocolo(nomelogin,seq,processo) values('" & NomeDeLogin & "'," & x & ",'" & sProcesso & "')"
        cn.Execute sql, rdExecDirect
    End If
Next

frmReport.ShowReport3 "CARTA_PROTOCOLO", frmMdi.HWND, Me.HWND

sql = "delete from carta_protocolo where nomelogin='" & NomeDeLogin & "'"
cn.Execute sql, rdExecDirect


End Sub

Private Sub cmdLoad_Click()


If optEtiq(0).value = True Then Exit Sub
If Val(txtDe.Text) < 1950 Or Val(txtDe.Text) > Year(Now) Then
    MsgBox "Ano inicial inválido", vbCritical, "Erro"
    Exit Sub
End If

If Val(txtAte.Text) < 1950 Or Val(txtAte.Text) > Year(Now) Then
    MsgBox "Ano final inválido", vbCritical, "Erro"
    Exit Sub
End If

If optEtiq(1).value = True Then
    CarregaLista 1
Else
    cmdPrint_Click
End If

End Sub

Private Sub cmdMalaDireta_Click()
Dim nCodCidadao As Long, nAno As Integer, nNumero As Long, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset
Dim sNome As String, sInscricao As String, sEndereco As String, sComplemento As String, nPos As Long
Dim sBairro As String, sCidade As String, sUF As String, sCep As String
Dim sCampo1 As String, sCampo2 As String, sCampo3 As String, sCampo4 As String, sCampo5 As String
Dim bRes As Boolean

sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute sql, rdExecDirect
nSeqP = 0
With lvEtiq
    For x = 1 To .ListItems.Count
        If .ListItems(x).Checked = True Then
            nAno = Val(.ListItems(x).Text)
            nNumero = CLng(Left$(.ListItems(x).SubItems(1), Len(.ListItems(x).SubItems(1)) - 2))
            sql = "SELECT  PROCESSOGTI.ANO, PROCESSOGTI.NUMERO, PROCESSOGTI.CODCIDADAO, cidadao.nomecidadao, PROCESSOGTI.COMPLEMENTO, "
            sql = sql & "CENTROCUSTO.DESCRICAO,processogti.tipoend FROM  PROCESSOGTI LEFT OUTER JOIN  CENTROCUSTO ON PROCESSOGTI.CENTROCUSTO = CENTROCUSTO.CODIGO LEFT OUTER JOIN "
            sql = sql & "cidadao ON PROCESSOGTI.CODCIDADAO = cidadao.codcidadao where PROCESSOGTI.ANO=" & nAno & " AND PROCESSOGTI.NUMERO=" & nNumero
            Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount > 0 Then
                    If IsNull(RdoAux!tipoend) Then
                        bRes = True
                    Else
                        If RdoAux!tipoend = "R" Then
                            bRes = True
                        Else
                            bRes = False
                        End If
                    End If
                    nCodCidadao = !CodCidadao
                    If nCodCidadao = 0 Then GoTo Proximo
                    
                    
                    If IsNull(!tipoend) Then
                        sTipoEnd = "R"
                    Else
                        sTipoEnd = !tipoend
                    End If
                                        
                    If sTipoEnd = "R" Then
                        sql = "SELECT CODCIDADAO,NOMECIDADAO,CPF,CNPJ,CODLOGRADOURO AS fCODLOGRADOURO,NUMIMOVEL AS fNUMIMOVEL,"
                        sql = sql & "COMPLEMENTO AS fCOMPLEMENTO,CODBAIRRO AS fCODBAIRRO,CODCIDADE AS fCODCIDADE,SIGLAUF AS fSIGLAUF,"
                        sql = sql & "CEP AS fCEP,TELEFONE AS fTELEFONE,EMAIL AS fEMAIL,RG AS fRG,NOMELOGRADOURO AS fNOMELOGRADOURO,ORGAO AS fORGAO"
                        sql = sql & " FROM CIDADAO WHERE CODCIDADAO=" & !CodCidadao
                    Else
                        sql = "SELECT CODCIDADAO,NOMECIDADAO,CPF,CNPJ,CODLOGRADOURO2 AS fCODLOGRADOURO,NUMIMOVEL2 AS fNUMIMOVEL,"
                        sql = sql & "COMPLEMENTO2 AS fCOMPLEMENTO,CODBAIRRO2 AS fCODBAIRRO,CODCIDADE2 AS fCODCIDADE,SIGLAUF2 AS fSIGLAUF,"
                        sql = sql & "CEP2 AS fCEP,TELEFONE2 AS fTELEFONE,EMAIL2 AS fEMAIL,RG AS fRG,NOMELOGRADOURO2 AS fNOMELOGRADOURO,ORGAO AS fORGAO"
                        sql = sql & " FROM CIDADAO WHERE CODCIDADAO=" & !CodCidadao
                    End If
                    Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                    On Error Resume Next
                    With RdoAux2
                        If .RowCount > 0 Then
                             sNome = !nomecidadao
                             If Val(SubNull(!FCodLogradouro)) > 0 Then
                                 sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
                                 sql = sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
                                 sql = sql & "ABREVTITLOG,NOMELOGRADOURO "
                                 sql = sql & "FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & !FCodLogradouro
                                 Set RdoS = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                                 With RdoS
                                     If .RowCount > 0 Then
                                        sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                                     Else
                                        sEnd = ""
                                     End If
                                    .Close
                                 End With
                             Else
                                sEnd = SubNull(!FNomeLogradouro)
                             End If
                             sEnd = sEnd & " " & SubNull(RdoAux2!fNUMIMOVEL)
                              
                             sql = "SELECT DESCCIDADE FROM CIDADE WHERE SIGLAUF='" & !fsiglauf & "' AND CODCIDADE=" & !fCodCidade
                             Set RdoS = cn.OpenResultset(sql, rdOpenKeyset)
                             If RdoS.RowCount > 0 Then
                                 sCidade = RdoS!descCidade
                             Else
                                  sCidade = ""
                             End If
                             If Not IsNull(!CodBairro) Then
                                 sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & !fsiglauf & "' AND CODCIDADE=" & !fCodCidade & " AND CODBAIRRO=" & !fCodBairro
                                 Set RdoS = cn.OpenResultset(sql, rdOpenKeyset)
                                 If .RowCount > 0 Then
                                     sBairro = RdoS!DescBairro
                                 Else
                                     sBairro = ""
                                 End If
                             Else
                                 sBairro = ""
                             End If
                             sUF = SubNull(!fsiglauf)
                             sFone = SubNull(!telefone)
                             sCep = SubNull(!FCEP)
                        Else
                            sEnd = ""
                            sBairro = ""
                            sCidade = ""
                            sFone = ""
                            sUF = ""
                            sCep = ""
                        End If
                       .Close
                    End With

                    
                    
'                    If bRes Then
'                        Sql = "SELECT vwFULLCIDADAO.codcidadao, vwFULLCIDADAO.cep, vwFULLCIDADAO.desccidade, vwFULLCIDADAO.nomecidadao, vwFULLCIDADAO.cpf, "
'                        Sql = Sql & "vwFULLCIDADAO.cnpj, vwFULLCIDADAO.codlogradouro, vwLOGRADOURO.abrevtipolog, vwLOGRADOURO.abrevtitlog,vwLOGRADOURO.nomelogradouro, vwFULLCIDADAO.numimovel, vwFULLCIDADAO.complemento, vwFULLCIDADAO.codbairro,"
'                        Sql = Sql & "bairro.descbairro, vwFULLCIDADAO.codcidade, cidade.desccidade AS Expr1, vwFULLCIDADAO.siglauf, uf.descuf,vwFULLCIDADAO.cep AS Expr2, vwFULLCIDADAO.nomelogradouro AS RUA2, vwFULLCIDADAO.desccidade2, vwFULLCIDADAO.codlogradouro2,"
'                        Sql = Sql & "vwFULLCIDADAO.numimovel2, vwFULLCIDADAO.complemento2, vwFULLCIDADAO.codbairro2, vwFULLCIDADAO.codcidade2,vwFULLCIDADAO.siglauf2 , vwFULLCIDADAO.cep2, vwFULLCIDADAO.nomelogradouroc, vwFULLCIDADAO.etiqueta "
'                        Sql = Sql & "FROM vwLOGRADOURO RIGHT OUTER JOIN vwFULLCIDADAO ON vwLOGRADOURO.codlogradouro = vwFULLCIDADAO.codlogradouro LEFT OUTER JOIN "
'                        Sql = Sql & "cidade INNER JOIN bairro ON cidade.siglauf = bairro.siglauf AND cidade.codcidade = bairro.codcidade INNER JOIN uf ON cidade.siglauf = uf.siglauf ON vwFULLCIDADAO.siglauf = bairro.siglauf AND vwFULLCIDADAO.codcidade = bairro.codcidade AND "
'                        Sql = Sql & "vwFULLCIDADAO.codbairro = bairro.codbairro WHERE CODCIDADAO=" & nCodCidadao
'                        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'                        With RdoAux2
'                            sNome = !nomecidadao
'                                If !CodLogradouro > 0 Then
'                                   sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
'                                Else
'                                   sEnd = SubNull(!RUA2)
'                                End If
'                                sEnd = sEnd & ", " & CStr(SubNull(!NUMIMOVEL))
'                                sCEP = RetornaCEP(Val(SubNull(!CodLogradouro)), !NUMIMOVEL)
'                                If sCEP = "00000-000" Then sCEP = SubNull(!Cep)
'                                sComplemento = SubNull(!Complemento)
'                                sBairro = SubNull(!DescBairro)
'                                sCidade = SubNull(!desccidade)
'                                sUF = SubNull(!SiglaUF)
'
'                           .Close
'                        End With
'                    Else
'                        Sql = "SELECT vwFULLCIDADAO.codcidadao, vwFULLCIDADAO.cep, vwFULLCIDADAO.desccidade, vwFULLCIDADAO.nomecidadao, vwFULLCIDADAO.cpf, "
'                        Sql = Sql & "vwFULLCIDADAO.cnpj, vwFULLCIDADAO.codlogradouro, vwLOGRADOURO.abrevtipolog, vwLOGRADOURO.abrevtitlog,vwLOGRADOURO.nomelogradouro, vwFULLCIDADAO.numimovel, vwFULLCIDADAO.complemento, vwFULLCIDADAO.codbairro,"
'                        Sql = Sql & "bairro.descbairro, vwFULLCIDADAO.codcidade, cidade.desccidade AS Expr1, vwFULLCIDADAO.siglauf, uf.descuf,vwFULLCIDADAO.cep AS Expr2, vwFULLCIDADAO.nomelogradouro AS RUA2, vwFULLCIDADAO.desccidade2, vwFULLCIDADAO.codlogradouro2,"
'                        Sql = Sql & "vwFULLCIDADAO.numimovel2, vwFULLCIDADAO.complemento2, vwFULLCIDADAO.codbairro2, vwFULLCIDADAO.codcidade2,vwFULLCIDADAO.siglauf2 , vwFULLCIDADAO.cep2, vwFULLCIDADAO.nomelogradouroc, vwFULLCIDADAO.etiqueta "
'                        Sql = Sql & "FROM vwLOGRADOURO RIGHT OUTER JOIN vwFULLCIDADAO ON vwLOGRADOURO.codlogradouro = vwFULLCIDADAO.codlogradouro LEFT OUTER JOIN "
'                        Sql = Sql & "cidade INNER JOIN bairro ON cidade.siglauf = bairro.siglauf AND cidade.codcidade = bairro.codcidade INNER JOIN uf ON cidade.siglauf = uf.siglauf ON vwFULLCIDADAO.siglauf = bairro.siglauf AND vwFULLCIDADAO.codcidade = bairro.codcidade AND "
'                        Sql = Sql & "vwFULLCIDADAO.codbairro = bairro.codbairro WHERE CODCIDADAO=" & nCodCidadao
'                        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'                        With RdoAux2
'                            sNome = !nomecidadao
'                                If !CodLogradouro > 0 Then
'                                   sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
'                                Else
'                                   sEnd = SubNull(!RUA2)
'                                End If
'                                sEnd = sEnd & ", " & CStr(SubNull(!NUMIMOVEL))
'                                sCEP = RetornaCEP(Val(SubNull(!CodLogradouro)), !NUMIMOVEL)
'                                If sCEP = "00000-000" Then sCEP = SubNull(!Cep)
'                                sComplemento = SubNull(!Complemento)
'                                sBairro = SubNull(!DescBairro)
'                                sCidade = SubNull(!desccidade)
'                                sUF = SubNull(!SiglaUF)
'
'                           .Close
'                        End With
'
'                    End If
                    sql = "INSERT ETIQUETAGTI (USUARIO,SEQ,CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5) VALUES('"
                    sql = sql & NomeDeLogin & "'," & nPos & ",'" & Format(nCodCidadao, "000000") & "','" & Mask(sNome) & "','" & Left(sEnd, 60) & "','" & Left(sBairro & " - " & sCidade, 60) & "','" & sUF & " - " & sCep & "')"
                    cn.Execute sql, rdExecDirect
                End If
               .Close
            End With
        End If
Proximo:
    Next
End With

Fim:
frmReport.ShowReport "ETIQUETAPROTOCOLO", frmMdi.HWND, Me.HWND

sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute sql, rdExecDirect

End Sub

Private Sub cmdRemoveAll_Click()
For n = 1 To lvEtiq.ListItems.Count
    lvEtiq.ListItems(n).Checked = False
Next
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdTramitacao_Click()
Dim x As Integer, nAno As Integer, nNumero As Integer, sProcesso As String, bFind As Boolean, nSeq As Integer, nUserId As Integer
Dim sql As String, RdoAux As rdoResultset, nCentroCusto As Integer, RdoAux2 As rdoResultset, nCodAssunto As Integer

MsgBox "Rotina desativada", vbCritical, "Erro"
Exit Sub

bFind = False
For x = 1 To lvEtiq.ListItems.Count
    If lvEtiq.ListItems(x).Checked Then
        bFind = True
        Exit For
    End If
Next

If Not bFind Then
    MsgBox "Selecione ao menos um processo", vbCritical, "Erro"
    Exit Sub
End If

sql = "delete from carta_protocolo where nomelogin='" & NomeDeLogin & "'"
cn.Execute sql, rdExecDirect

For x = 1 To lvEtiq.ListItems.Count
    If lvEtiq.ListItems(x).Checked Then
        nAno = Val(lvEtiq.ListItems(x).Text)
        nNumero = CLng(Left$(lvEtiq.ListItems(x).SubItems(1), Len(lvEtiq.ListItems(x).SubItems(1)) - 2))
        sProcesso = Format(nNumero, "00000") & "-" & RetornaDVProcesso(CLng(nNumero)) & "/" & nAno
        
        sql = "select * from tramitacaocc where ano=" & nAno & " and numero=" & nNumero & " and seq=1"
        Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux.RowCount = 0 Then
            sql = "SELECT CODASSUNTO FROM processogti WHERE ano=" & nAno & " and numero=" & nNumero
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            nCodAssunto = RdoAux2!codassunto
                        
        
            sql = "SELECT ASSUNTOCC.SEQ,CENTROCUSTO.CODIGO, CENTROCUSTO.DESCRICAO FROM ASSUNTOCC INNER JOIN "
            sql = sql & "CENTROCUSTO ON ASSUNTOCC.CODCC = CENTROCUSTO.CODIGO "
            sql = sql & "WHERE ASSUNTOCC.CODASSUNTO =" & nCodAssunto
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                nSeq = 1
                Do Until .EOF
                    sql = "INSERT TRAMITACAOCC(ANO,NUMERO,SEQ,CCUSTO) VALUES("
                    sql = sql & nAno & "," & nNumero & "," & nSeq & "," & !Codigo & ")"
                    cn.Execute sql, rdExecDirect
                    nSeq = nSeq + 1
                   .MoveNext
                Loop
               .Close
            End With
        End If
    End If
Next

nUserId = RetornaUsuarioID(NomeDeLogin)

For x = 1 To lvEtiq.ListItems.Count
    If lvEtiq.ListItems(x).Checked Then
    
        nAno = Val(lvEtiq.ListItems(x).Text)
        nNumero = CLng(Left$(lvEtiq.ListItems(x).SubItems(1), Len(lvEtiq.ListItems(x).SubItems(1)) - 2))
    
        sql = "select * from tramitacaocc where ano=" & nAno & " and numero=" & nNumero & " and ccusto=14  AND seq NOT IN (select seq from tramitacao where ano=" & nAno & " and numero=" & nNumero & " and ccusto=14)"
        Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux.RowCount > 0 Then
            nSeq = RdoAux!Seq
            nCentroCusto = RdoAux!ccusto
        Else
            GoTo Proximo
        End If

        sql = "select * from tramitacao where ano=" & nAno & " and numero=" & nNumero & " and seq=" & nSeq
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount = 0 Then
            sql = "insert tramitacao (ano,numero,seq,ccusto,datahora,despacho,userid) values("
            sql = sql & nAno & "," & nNumero & "," & nSeq & ",14,'" & Format(Now, "mm/dd/yyyy hh:mm") & "',19," & nUserId & ")"
            'cn.Execute sql, rdExecDirect
        Else
            If RdoAux2!ccusto = 14 And IsNull(RdoAux2!dataHora) Then
                sql = "update tramitacao set datahora='" & Format(Now, "mm/dd/yyyy hh:mm") & "',despacho=19,userid=" & nUserId
                sql = sql & " where ano=" & nAno & " and numero=" & nNumero
            '    cn.Execute sql, rdExecDirect
            End If
        End If
        

    End If
Proximo:
Next

MsgBox "Tramitação incluida nos processos.", vbInformation, "Atenção"

End Sub

Private Sub Form_Load()
Centraliza Me

sql = "SELECT CODIGO,NOME FROM ASSUNTO WHERE ATIVO=1 ORDER BY NOME"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbAssunto.AddItem (!Nome)
        cmbAssunto.ItemData(cmbAssunto.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With

If NomeDeLogin <> "FABIANO.PIRES" And NomeDeLogin <> "PEDROS" And NomeDeLogin <> "RAPHAEL.CARREGARI" And NomeDeLogin <> "ROSELI.BORGES" And NomeDeLogin <> "SCHWARTZ" And NomeDeLogin <> "OTTO.SEGUNDO" Then
    cmdTramitacao.Enabled = False
    cmdCarta.Enabled = False
End If


cmbAssunto.ListIndex = 0
CarregaLista (0)
End Sub

Private Sub CarregaLista(nQual As Integer)
Dim itmX As ListItem, z As Long
z = SendMessage(lvEtiq.HWND, LVM_DELETEALLITEMS, 0, 0)

Ocupado

sql = "SELECT  PROCESSOGTI.ANO, PROCESSOGTI.NUMERO, PROCESSOGTI.CODCIDADAO, cidadao.nomecidadao, PROCESSOGTI.COMPLEMENTO, "
sql = sql & "CENTROCUSTO.DESCRICAO FROM  PROCESSOGTI LEFT OUTER JOIN  CENTROCUSTO ON PROCESSOGTI.CENTROCUSTO = CENTROCUSTO.CODIGO LEFT OUTER JOIN "
sql = sql & "cidadao ON PROCESSOGTI.CODCIDADAO = cidadao.codcidadao "
If nQual = 0 Then
     sql = sql & " Where (PROCESSOGTI.ETIQUETA = 0) And (PROCESSOGTI.FISICO = 1) "
     sql = sql & " AND (PROCESSOGTI.DATACANCEL IS NULL)  AND (PROCESSOGTI.DATAARQUIVA IS NULL) AND (PROCESSOGTI.DATASUSPENSO IS NULL) AND PROCESSOGTI.ETIQUETA=0"
Else
    'Sql = Sql & " Where (PROCESSOGTI.FISICO = 1)  AND (ANO>" & Year(Now) - 5 & ")"
    sql = sql & " Where (PROCESSOGTI.FISICO = 1)  AND (ANO BETWEEN " & Val(txtDe.Text) & " AND " & Val(txtAte.Text) & ")"
End If
sql = sql & " ORDER BY PROCESSOGTI.ANO, PROCESSOGTI.NUMERO"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Set itmX = lvEtiq.ListItems.Add(, "E" & Format(!ano, "0000") & Format(!Numero, "000000"), !ano)
        itmX.SubItems(1) = !Numero & "-" & RetornaDVProcesso(!Numero)
        If Not IsNull(!nomecidadao) Then
            itmX.SubItems(2) = SubNull(!nomecidadao)
        Else
            itmX.SubItems(2) = SubNull(!Descricao)
        End If
        itmX.SubItems(3) = SubNull(!Complemento)
        .MoveNext
    Loop
   .Close
End With

Liberado

End Sub

Private Sub MalaDireta()
Dim sNome As String, sEnd As String, sBairro As String, sCidade As String, sUF As String, sCompl As String, sCep As String
Dim nSeq As Integer, sValidade As String

If Val(txtDe.Text) < 1950 Or Val(txtDe.Text) > Year(Now) Then
    MsgBox "Ano inicial inválido", vbCritical, "Erro"
    Exit Sub
End If

If Val(txtAte.Text) < 1950 Or Val(txtAte.Text) > Year(Now) Then
    MsgBox "Ano final inválido", vbCritical, "Erro"
    Exit Sub
End If

sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute sql, rdExecDirect

nSeq = 1
sql = "SELECT * FROM VWFULLPROCESSO WHERE CODASSUNTO=" & cmbAssunto.ItemData(cmbAssunto.ListIndex) & " AND INTERNO=0 AND ANO BETWEEN " & Val(txtDe.Text) & " AND " & Val(txtAte.Text)
If chkArquivado.value = vbChecked Then
    sql = sql & " AND DATAARQUIVA IS NULL"
End If
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        sNome = !nomecidadao
        sEnd = !Endereco & ", " & !NUMIMOVEL & " " & sCompl
        sBairro = SubNull(!DescBairro)
        sCidade = SubNull(!descCidade)
        sUF = SubNull(!SiglaUF)
        sCompl = SubNull(!Compl)
        sCep = SubNull(!Cep)
                
        If Not IsNull(RdoAux!validade_tipo) Then
            If RdoAux!validade_tipo = 1 Then
                dias = RdoAux!validade_qtde
            ElseIf RdoAux!validade_tipo = 2 Then
                dias = RdoAux!validade_qtde * 30
            Else
                dias = RdoAux!validade_qtde * 365
            End If
            sValidade = "Validade: " & Format(DateAdd("d", dias, Now), "dd/mm/yyyy")
        Else
            sValidade = ""
        End If
        
        If sCidade = "JABOTICABAL" Then
            sCep = RetornaCEP(Val(SubNull(!CodLogradouro)), !NUMIMOVEL)
        End If
        
        sql = "INSERT ETIQUETAGTI (USUARIO,SEQ,CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5) VALUES('"
        sql = sql & NomeDeLogin & "'," & nSeq & ",'" & sValidade & "','" & Mask(sNome) & "','" & Left(sEnd, 60) & "','" & sBairro & " - " & sCidade & "','" & sUF & " - " & sCep & "')"
        cn.Execute sql, rdExecDirect
        
        
        
        nSeq = nSeq + 1
       .MoveNext
    Loop
   .Close
End With

frmReport.ShowReport "ETIQUETAPROTOCOLO", frmMdi.HWND, Me.HWND

sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute sql, rdExecDirect

End Sub

Private Sub lvEtiq_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
With lvEtiq
    .SortKey = ColumnHeader.Position - 1
    .Sorted = True
    .SortOrder = lvwAscending
End With
End Sub

Private Sub optEtiq_Click(Index As Integer)
If Index = 0 Then
    CarregaLista (Index)
    txtDe.Locked = True: txtAte.Locked = True
Else
    txtDe.Locked = False: txtAte.Locked = False
End If

End Sub

Private Sub cmdPrint_Click()
Dim nProc As Long, nAno As Integer, sProc As String, sData As String, sAssunto As String, nAssunto As Integer
Dim sCompl As String, x As Long, y As Integer, aMatriz() As String, sTmp As String, nReq As Long, sReq As String, RdoAux2 As rdoResultset
Dim sEnd As String, sBairro As String, sCidade As String, sFone As String, RdoS As rdoResultset, nSeqP As Integer, sObs1 As String, sObs2 As String
Dim sObs3 As String, bInterno As Boolean, sTipoEnd As String, sCep As String, sValidade As String

If optEtiq(2).value = True Then
    MalaDireta
    Exit Sub
End If

'GoTo fim
sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute sql, rdExecDirect
nSeqP = 0
With lvEtiq
    For x = 1 To .ListItems.Count
        If .ListItems(x).Checked = True Then
        
            nAno = .ListItems(x).Text
            nProc = Val(Left$(.ListItems(x).SubItems(1), Len(.ListItems(x).SubItems(1)) - 2))
            sProc = CStr(nProc) & "-" & RetornaDVProcesso(CStr(nProc)) & "/" & CStr(nAno)
            sql = "UPDATE  PROCESSOGTI SET ETIQUETA=1 "
            sql = sql & " Where PROCESSOGTI.ANO = " & nAno & " And PROCESSOGTI.Numero = " & nProc
            cn.Execute sql, rdExecDirect
            
            sql = "SELECT  PROCESSOGTI.*, ASSUNTO.NOME AS ASSUNTO "
            sql = sql & "FROM  PROCESSOGTI INNER JOIN  ASSUNTO ON PROCESSOGTI.CODASSUNTO = ASSUNTO.CODIGO "
            sql = sql & " Where PROCESSOGTI.ANO = " & nAno & " And PROCESSOGTI.Numero = " & nProc
            Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                    sObs1 = "": sObs2 = "": sObs3 = ""
                    bInterno = !INTERNO
                    If Not IsNull(!OBSERVACAO) Then
                        If Len(!OBSERVACAO) > 1 Then
                            sObs1 = Left$(!OBSERVACAO, 55)
                        End If
                        If Len(!OBSERVACAO) > 55 Then
                            sObs2 = Mid$(!OBSERVACAO, 56, 55)
                        End If
                        If Len(!OBSERVACAO) > 110 Then
                            sObs3 = Mid$(!OBSERVACAO, 111, 55)
                        End If
                    End If
                                        
                    If IsNull(!tipoend) Then
                        sTipoEnd = "R"
                    Else
                        sTipoEnd = !tipoend
                    End If
                                        
                    If sTipoEnd = "R" Then
                        sql = "SELECT CODCIDADAO,NOMECIDADAO,CPF,CNPJ,CODLOGRADOURO AS fCODLOGRADOURO,NUMIMOVEL AS fNUMIMOVEL,"
                        sql = sql & "COMPLEMENTO AS fCOMPLEMENTO,CODBAIRRO AS fCODBAIRRO,CODCIDADE AS fCODCIDADE,SIGLAUF AS fSIGLAUF,"
                        sql = sql & "CEP AS fCEP,TELEFONE AS fTELEFONE,EMAIL AS fEMAIL,RG AS fRG,NOMELOGRADOURO AS fNOMELOGRADOURO,ORGAO AS fORGAO"
                        sql = sql & " FROM CIDADAO WHERE CODCIDADAO=" & !CodCidadao
                    Else
                        sql = "SELECT CODCIDADAO,NOMECIDADAO,CPF,CNPJ,CODLOGRADOURO2 AS fCODLOGRADOURO,NUMIMOVEL2 AS fNUMIMOVEL,"
                        sql = sql & "COMPLEMENTO2 AS fCOMPLEMENTO,CODBAIRRO2 AS fCODBAIRRO,CODCIDADE2 AS fCODCIDADE,SIGLAUF2 AS fSIGLAUF,"
                        sql = sql & "CEP2 AS fCEP,TELEFONE2 AS fTELEFONE,EMAIL2 AS fEMAIL,RG AS fRG,NOMELOGRADOURO2 AS fNOMELOGRADOURO,ORGAO AS fORGAO"
                        sql = sql & " FROM CIDADAO WHERE CODCIDADAO=" & !CodCidadao
                    End If
                    Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                    On Error Resume Next
                    With RdoAux2
                        If .RowCount > 0 Then
                             If Val(SubNull(!FCodLogradouro)) > 0 Then
                                 sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
                                 sql = sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
                                 sql = sql & "ABREVTITLOG,NOMELOGRADOURO "
                                 sql = sql & "FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & !FCodLogradouro
                                 Set RdoS = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                                 With RdoS
                                     If .RowCount > 0 Then
                                        sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                                     Else
                                        sEnd = ""
                                     End If
                                    .Close
                                 End With
                             Else
                                sEnd = SubNull(!FNomeLogradouro)
                             End If
                             sEnd = sEnd & " " & SubNull(RdoAux2!fNUMIMOVEL)
'                             sCEP = RetornaCEP(!fcodlogradouro, !fnumimovel)
                            
                             sql = "SELECT DESCCIDADE FROM CIDADE WHERE SIGLAUF='" & !fsiglauf & "' AND CODCIDADE=" & !fCodCidade
                             Set RdoS = cn.OpenResultset(sql, rdOpenKeyset)
                             If RdoS.RowCount > 0 Then
                                 sCidade = RdoS!descCidade
                             Else
                                  sCidade = ""
                             End If
                             If Not IsNull(!CodBairro) Then
                                 sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & !fsiglauf & "' AND CODCIDADE=" & !fCodCidade & " AND CODBAIRRO=" & !fCodBairro
                                 Set RdoS = cn.OpenResultset(sql, rdOpenKeyset)
                                 If .RowCount > 0 Then
                                     sBairro = RdoS!DescBairro
                                 Else
                                     sBairro = ""
                                 End If
                             Else
                                 sBairro = ""
                             End If
                             sFone = SubNull(RdoAux2!fTELEFONE)
                        Else
                            sEnd = ""
                            sBairro = ""
                            sCidade = ""
                            sFone = ""
                        End If
                       .Close
                    End With

                    sData = Format(!DATAENTRADA, "dd/mm/yyyy")
                    sAssunto = SubNull(!assunto)
                    nAssunto = !codassunto
                                        
                    sql = "SELECT assunto.CODIGO, assunto.NOME, assunto.VALIDADE_TIPO, assunto.VALIDADE_QTDE, tipovalidade.descricao "
                    sql = sql & "FROM assunto  LEFT OUTER JOIN tipovalidade ON assunto.VALIDADE_TIPO = tipovalidade.codigo Where assunto.Codigo = " & nAssunto
                    Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                    If Val(SubNull(RdoAux!validade_qtde)) > 0 Then
                        If RdoAux!validade_tipo = 1 Then
                            dias = RdoAux!validade_qtde
                        ElseIf RdoAux!validade_tipo = 2 Then
                            dias = RdoAux!validade_qtde * 30
                        Else
                            dias = RdoAux!validade_qtde * 365
                        End If
                        sValidade = Format(DateAdd("d", dias, Now), "dd/mm/yyyy")
                    Else
                        sValidade = ""
                    End If
                    
                    
                    sCompl = Left$(SubNull(!Complemento), 200)
                    If !ORIGEM = 2 Then
                        sql = "SELECT CODCIDADAO,NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO=" & !CodCidadao
                        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux2
                            If .RowCount > 0 Then
                                nReq = !CodCidadao
                                sReq = SubNull(!nomecidadao)
                            End If
                           .Close
                        End With
                    ElseIf !ORIGEM = 1 Then
                        sql = "SELECT DESCRICAO FROM CENTROCUSTO WHERE CODIGO=" & !CENTROCUSTO
                        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux2
                            If .RowCount > 0 Then
                                nReq = 0
                                sReq = SubNull(!Descricao)
                            End If
                           .Close
                        End With
                    End If
                   .Close
            End With
            nSeqP = nSeqP + 1
            sql = "INSERT ETIQUETAGTI (USUARIO,SEQ,CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5,PROCESSO) VALUES('"
            sql = sql & NomeDeLogin & "'," & nSeqP & ",'" & sProc & "','" & "Abertura:" & sData & "','" & "Validade: " & sValidade & "','" & "" & "','" & Mask(Left$(sCompl, 50)) & "','" & sProc & "')"
            cn.Execute sql, rdExecDirect
            nSeqP = nSeqP + 1
            On Error GoTo 0
            sql = "INSERT ETIQUETAGTI (USUARIO,SEQ,CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5,PROCESSO) VALUES('"
            If Not bInterno Then
                sql = sql & NomeDeLogin & "'," & nSeqP & ",'" & nReq & "','" & Mask(sReq) & "','" & sEnd & "','" & Left(sBairro & " - " & Mask(sCidade), 60) & "','" & "Telefone: " & sFone & "','" & sProc & "')"
            Else
                sql = sql & NomeDeLogin & "'," & nSeqP & ",'" & IIf(nReq > 0, nReq, "") & "','" & Mask(sReq) & "','" & Mask(sObs1) & "','" & Mask(sObs2) & "','" & Mask(sObs3) & "','" & sProc & "')"
            End If
            cn.Execute sql, rdExecDirect
        
            'CARREGA TODOS OS TRAMITES
            ReDim aMatriz(0)
            sql = "SELECT * FROM tramitacaocc Where ano = " & nAno & " And Numero = " & nProc
            Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount = 0 Then
                    sql = "SELECT ASSUNTOCC.SEQ,CENTROCUSTO.CODIGO, CENTROCUSTO.DESCRICAO,CENTROCUSTO.SIGLA FROM ASSUNTOCC INNER JOIN "
                    sql = sql & "CENTROCUSTO ON ASSUNTOCC.CODCC = CENTROCUSTO.CODIGO "
                    sql = sql & "WHERE ASSUNTOCC.CODASSUNTO =" & nAssunto
                    Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux
                        Do Until .EOF
                            ReDim Preserve aMatriz(UBound(aMatriz) + 1)
                            aMatriz(UBound(aMatriz)) = CStr(UBound(aMatriz)) & " - " & SubNull(!Sigla) & " " & SubNull(!Descricao)
                           .MoveNext
                        Loop
                       .Close
                    End With
                Else
                    sql = "SELECT tramitacaocc.seq, tramitacaocc.ccusto, CENTROCUSTO.DESCRICAO,CENTROCUSTO.SIGLA "
                    sql = sql & "FROM tramitacaocc INNER JOIN CENTROCUSTO ON tramitacaocc.ccusto = CENTROCUSTO.CODIGO "
                    sql = sql & "Where tramitacaocc.ano = " & nAno & " And tramitacaocc.Numero = " & nProc
                    sql = sql & " order by TRAMITACAOCC.SEQ"
                    Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux
                        Do Until .EOF
                            ReDim Preserve aMatriz(UBound(aMatriz) + 1)
                            aMatriz(UBound(aMatriz)) = CStr(UBound(aMatriz)) & " - " & SubNull(!Sigla) & " " & SubNull(!Descricao)
                           .MoveNext
                        Loop
                       .Close
                    End With
                End If
               .Close
            End With
            'FINALIZANDO
            For y = 1 To UBound(aMatriz)
                Select Case y
                   Case 1, 6, 11, 16, 21
                        sTmp = aMatriz(y)
                        nSeqP = nSeqP + 1
                        sql = "INSERT ETIQUETAGTI (USUARIO,SEQ,CAMPO1,PROCESSO) VALUES('"
                        sql = sql & NomeDeLogin & "'," & nSeqP & ",'" & Left(aMatriz(y), 60) & "','" & sProc & "')"
                    Case 2, 7, 12, 17, 22
'                        Sql = "UPDATE ETIQUETAGTI  SET CAMPO2='" & Left(aMatriz(y), 200) & "' WHERE  USUARIO='" & NomeDoUsuario & "' AND CAMPO1='" & sTmp & "'"
                        sql = "UPDATE ETIQUETAGTI  SET CAMPO2='" & Left(aMatriz(y), 200) & "' WHERE  USUARIO='" & NomeDeLogin & "' AND SEQ=" & nSeqP
                    Case 3, 8, 13, 18, 23
                       sql = "UPDATE ETIQUETAGTI  SET CAMPO3='" & Left(aMatriz(y), 60) & "' WHERE   USUARIO='" & NomeDeLogin & "' AND SEQ=" & nSeqP
                     Case 4, 9, 14, 19, 24
                        sql = "UPDATE ETIQUETAGTI  SET CAMPO4='" & Left(aMatriz(y), 60) & "' WHERE  USUARIO='" & NomeDeLogin & "' AND SEQ=" & nSeqP
                    Case 5, 10, 15, 20, 25
                        sql = "UPDATE ETIQUETAGTI  SET CAMPO5='" & Left(aMatriz(y), 60) & "' WHERE   USUARIO='" & NomeDeLogin & "' AND SEQ=" & nSeqP
                End Select
                 cn.Execute sql, rdExecDirect
            Next
        End If
Proximo:
    Next
End With

Fim:
'rpt.RecordSelectionFormula = "{ETIQUETAGTI.USUARIO}='" & NomeDoUsuario & "'"
frmReport.ShowReport "ETIQUETAPROTOCOLO3", frmMdi.HWND, Me.HWND

sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute sql, rdExecDirect

End Sub

Private Sub txtAte_KeyPress(KeyAscii As Integer)
Tweak txtAte, KeyAscii, IntegerPositive
End Sub

Private Sub txtDe_KeyPress(KeyAscii As Integer)
Tweak txtDe, KeyAscii, IntegerPositive
End Sub

