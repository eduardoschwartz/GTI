VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmProcessosAtivos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Processos Ativos"
   ClientHeight    =   3300
   ClientLeft      =   4710
   ClientTop       =   3450
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3300
   ScaleWidth      =   10965
   Begin MSFlexGridLib.MSFlexGrid grdTramite 
      Height          =   2475
      Left            =   135
      TabIndex        =   0
      Top             =   225
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   4366
      _Version        =   393216
      Rows            =   1
      Cols            =   10
      FixedCols       =   0
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmProcessosAtivos.frx":0000
   End
   Begin prjChameleon.chameleonButton cmdOk 
      Height          =   345
      Left            =   9405
      TabIndex        =   1
      ToolTipText     =   "Inserir local selecionado"
      Top             =   2835
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   609
      BTYPE           =   14
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmProcessosAtivos.frx":00D4
      PICN            =   "frmProcessosAtivos.frx":00F0
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
      TabIndex        =   2
      Top             =   2970
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   291
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label lblAno 
      Caption         =   "0"
      Height          =   240
      Left            =   3375
      TabIndex        =   4
      Top             =   2925
      Width           =   1185
   End
   Begin VB.Label lblPB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100 %"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2340
      TabIndex        =   3
      Top             =   2970
      Width           =   480
   End
End
Attribute VB_Name = "frmProcessosAtivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lista_Arquivado() As Integer

Private Sub Form_Load()
Dim Sql As String, RdoAux As rdoResultset
Centraliza Me

ReDim Lista_Arquivado(0)
Sql = "select codigo from despacho_Arquivado order by codigo"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve Lista_Arquivado(UBound(Lista_Arquivado) + 1)
        Lista_Arquivado(UBound(Lista_Arquivado)) = !codigo
       .MoveNext
    Loop
   .Close
End With

End Sub


Private Sub CallPb(nPosF As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents

If ((nPosF * 100) / nTotal) <= 100 Then
   Pb.value = (nPosF * 100) / nTotal
Else
   Pb.value = 100
End If
lblPB.Caption = FormatNumber(Pb.value, 2)

'Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub

Private Function IsArquivado(codigo As Integer) As Boolean
IsArquivado = isInAnyArray(Lista_Arquivado, codigo)
End Function

Private Sub cmdOk_Click()
Dim x As Integer
For x = 2005 To 2020
    lblAno.Caption = x
    Me.Refresh
    Exec (x)
Next

End Sub


Private Sub Exec(nAno As Integer)
Dim nCodReduz As Long, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos2 As Integer, bFind As Boolean, RdoAux3 As rdoResultset
Dim nPos As Long, nTot As Long, sUF As String, sCidade As String, sBairro As String, sCep As String, nCodBairro As Integer, nCodCidade As Integer
Dim Numero As Long, nCodDespacho As Integer, sDespacho As String, nCodigoCC As Integer, sNomeCC As String, sNome As String, nCodigo As Long, ax As String
Dim nMax As Integer, nLinha As Integer, bArquivado As Boolean
On Error GoTo Erro

sNomeArq = sPathBin & "\PROCESSO" & CStr(nAno) & ".TXT"
FF1 = FreeFile()
Open sNomeArq For Output As FF1

Sql = "SELECT p.ANO,p.NUMERO,p.CODASSUNTO,a.NOME AS DESCASSUNTO,p.COMPLEMENTO,p.OBSERVACAO,p.DATAENTRADA,p.INTERNO,p.CODCIDADAO,c.nomecidadao,p.CENTROCUSTO,u.DESCRICAO AS CENTROCUSTONOME "
Sql = Sql & "FROM processogti p LEFT OUTER JOIN assunto a ON p.CODASSUNTO=a.CODIGO LEFT OUTER JOIN cidadao c ON p.CODCIDADAO=c.codcidadao LEFT OUTER JOIN centrocusto u ON p.CENTROCUSTO=u.CODIGO "
Sql = Sql & "WHERE  p.ano=" & nAno & " and p.fisico=1 and p.DATAARQUIVA IS NULL AND p.DATASUSPENSO IS NULL AND p.DATADESCARTE IS NULL AND p.ano>1977 AND SUBSTRING(p.OBSERVACAO,1,9)<>'ARQUIVADO' ORDER BY p.ano,p.numero"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        grdTramite.Rows = 1
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
            DoEvents
        End If
        nAno = !Ano
        nNumero = !Numero
        If !interno Then
            sNome = !centrocustonome
        Else
            sNome = SubNull(!nomecidadao)
        End If
        
        'CARREGA TODOS OS TRAMITES
        Sql = "SELECT * FROM tramitacaocc Where ano = " & nAno & " And Numero = " & nNumero
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount = 0 Then
                Sql = "SELECT ano, numero, seq, ccusto, DESCRICAO From vwTRAMITACAO2 Where ano =" & nAno & " And Numero = " & nNumero & " order by seq"
                Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux3
                    Do Until .EOF
                        grdTramite.AddItem !Seq & Chr(9) & !ccusto & Chr(9) & SubNull(!Descricao)
                        nMax = !Seq
                       .MoveNext
                    Loop
                   .Close
                End With
                
            
                Sql = "SELECT ASSUNTOCC.SEQ,CENTROCUSTO.CODIGO, CENTROCUSTO.DESCRICAO FROM ASSUNTOCC INNER JOIN "
                Sql = Sql & "CENTROCUSTO ON ASSUNTOCC.CODCC = CENTROCUSTO.CODIGO "
                Sql = Sql & "WHERE ASSUNTOCC.CODASSUNTO =" & RdoAux!codassunto
                Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux3
                    Do Until .EOF
                        bAchou = False
                        For x = 1 To grdTramite.Rows - 1
                            If grdTramite.TextMatrix(x, 1) = !codigo Then
                                bAchou = True
                                Exit For
                            End If
                        Next
                        nMax = nMax + 1
                        grdTramite.AddItem nMax & Chr(9) & !codigo & Chr(9) & SubNull(!Descricao)
                       .MoveNext
                    Loop
                   .Close
                End With
            Else
                Sql = "SELECT tramitacaocc.seq, tramitacaocc.ccusto, CENTROCUSTO.DESCRICAO "
                Sql = Sql & "FROM tramitacaocc INNER JOIN CENTROCUSTO ON tramitacaocc.ccusto = CENTROCUSTO.CODIGO "
                Sql = Sql & "Where tramitacaocc.ano = " & nAno & " And tramitacaocc.Numero = " & nNumero
                Sql = Sql & " order by TRAMITACAOCC.SEQ"
                Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux3
                    Do Until .EOF
                        grdTramite.AddItem !Seq & Chr(9) & !ccusto & Chr(9) & SubNull(!Descricao)
                       .MoveNext
                    Loop
                   .Close
                End With
            End If
           .Close
        End With
        
        'VERIFICA OS TRAMITES CONCLUIDOS
        For x = 1 To grdTramite.Rows - 1
            Sql = "SELECT CCUSTO,DESCRICAO,DATAHORA,NOMECOMPLETO,despacho,DESCDESPACHO,dataenvio,nomelogin2 FROM vwTRAMITACAO2 WHERE ANO=" & nAno
            Sql = Sql & " AND NUMERO=" & nNumero & " AND SEQ=" & grdTramite.TextMatrix(x, 0)
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If .RowCount > 0 Then
                    grdTramite.TextMatrix(x, 3) = Format(!DATAHORA, "dd/mm/yyyy")
                    grdTramite.TextMatrix(x, 4) = Format(!DATAHORA, "hh:mm")
                    grdTramite.TextMatrix(x, 5) = SubNull(!NomeCompleto)
                    grdTramite.TextMatrix(x, 6) = SubNull(!DESCDESPACHO)
                    grdTramite.TextMatrix(x, 7) = Val(SubNull(!DESPACHO))
                    If Not IsNull(!DATAENVIO) Then
                        grdTramite.TextMatrix(x, 8) = Format(!DATAENVIO, "dd/mm/yyyy")
                    End If
                    grdTramite.TextMatrix(x, 9) = SubNull(!nomelogin2)
                End If
               .Close
            End With
            Sql = "SELECT * FROM TRAMITACAO WHERE ANO=" & nAno & " AND NUMERO=" & nNumero
            Sql = Sql & " AND SEQ=" & grdTramite.TextMatrix(x, 0)
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            
            grdTramite.Row = x
            grdTramite.col = 0
            grdTramite.ColAlignment(1) = flexAlignRightCenter
            grdTramite.CellForeColor = vbWhite
        Next
        
        '******** verifica o trâmite ********
        bArquivado = False
        nLinha = grdTramite.Rows - 1
Start1:
            If grdTramite.TextMatrix(nLinha, 4) = "" Then
            'se for a primeira linha sai
            If nLinha = 1 Then
                sNomeCC = grdTramite.TextMatrix(nLinha, 2)
                sDespacho = ""
                GoTo Imprime
            End If
            'se não foi recebido, sobe uma linha
            nLinha = nLinha - 1
            GoTo Start1
        Else
            nCodDespacho = Val(grdTramite.TextMatrix(nLinha, 7))
            sDespacho = grdTramite.TextMatrix(nLinha, 6)
            If nCodDespacho > 0 Then
                bArquivado = IsArquivado(nCodDespacho)
            Else
                nLinha = nLinha - 1
                nCodDespacho = Val(grdTramite.TextMatrix(nLinha, 7))
                sDespacho = grdTramite.TextMatrix(nLinha, 6)
                nLinha = nLinha + 1
            End If
            If bArquivado Then
                GoTo Proximo
            End If
            If nLinha < grdTramite.Rows - 1 Then
                If grdTramite.TextMatrix(nLinha, 8) = "" Then
                    nCodigoCC = Val(grdTramite.TextMatrix(nLinha, 1))
                    sNomeCC = grdTramite.TextMatrix(nLinha, 2)
                    If nCodigoCC = 19 Or nCodigoCC = 14 Or nCodigoCC = 172 Then
                        GoTo Proximo
                    End If
                Else
                    sDespacho = grdTramite.TextMatrix(nLinha, 6)
                    nLinha = nLinha + 1
                    nCodigoCC = Val(grdTramite.TextMatrix(nLinha, 1))
                    sNomeCC = grdTramite.TextMatrix(nLinha, 2)
                    
                    If nCodigoCC = 19 Or nCodigoCC = 14 Or nCodigoCC = 172 Then
                        GoTo Proximo
                    End If
                End If
            End If
        End If
        
        sNome = grdTramite.TextMatrix(nLinha, 5)
        If RdoAux!interno Then
            nCodigo = nCodigoCC
            sNome = !centrocustonome
        Else
            nCodigo = RdoAux!CodCidadao
            sNome = RdoAux!nomecidadao
        End If
Imprime:
        ax = nNumero & "-" & RetornaDVProcesso(CStr(nNumero)) & "/" & nAno & "#" & Format(RdoAux!DATAENTRADA, "dd/mm/yyyy") & "#" & RdoAux!descassunto & "#"
        ax = ax & sDespacho & "#" & nCodigo & "#" & sNome & "#" & sNomeCC
        Print #1, ax
                
Proximo:
        nPos = nPos + 1
        
       .MoveNext
    Loop
   .Close
End With
Close #FF1
'MsgBox "Fim"

Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub
