Attribute VB_Name = "gti_Eicon"

Public Sub AtualizaEmpresa()
Dim sql As String, RdoAux As rdoResultset, RdoEmp As rdoResultset, RdoAux2 As rdoResultset, RdoProp As rdoResultset
Dim nCodigo As Long, sIE As String, sRazao As String, sFantasia As String, sNumProcesso As String, sTipoEmpresa As String, nArea As Double
Dim sDoc As String, sDataAbertura As String, sDataEncerramento As String, sTipoLog As String, sTitLog As String, sNomeLog As String, sRegime As String
Dim nNumImovel As Integer, sCompl As String, sBairro As String, sCep As String, sCidade As String, sUF As String, sFone As String, sFax As String, sEmail As String
Dim nCodCidadao As Long, sNome As String, nCodLogr As Long, nTipoEnd As String, sDDD As String

ConectaEicon

sql = "select codigo from eicon_empresa order by codigo"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
        Do Until .EOF
        If .AbsolutePosition > 150 Then Exit Do
        nCodigo = !Codigo
        
       '******* DADOS DA EMPRESA **************************
        
        sql = "select * from vwfullempresa3 where codigomob=" & nCodigo
        Set RdoEmp = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoEmp
            sIE = RetornaNumero(SubNull(!inscestadual))
            sRazao = !RazaoSocial
            sFantasia = SubNull(!NOMEFANTASIA)
            sNumProcesso = SubNull(!NumProcesso)
            If Len(SubNull(!Cnpj)) = 14 Then
                sTipoEmpresa = "J"
                sDoc = RetornaNumero(!Cnpj)
            Else
                sTipoEmpresa = "F"
                sDoc = RetornaNumero(SubNull(!cpf))
            End If
            If Len(sDoc) < 2 Then sTipoEmpresa = "J"
            sDataAbertura = Format(!DataAbertura, "dd/mm/yyyy")
            If IsNull(!dataencerramento) Then
                sDataEncerramento = ""
            Else
                sDataEncerramento = Format(!dataencerramento, "dd/mm/yyyy")
            End If
            sTipoLog = Trim(SubNull(!AbrevTipoLog))
            sTitLog = Trim(SubNull(!AbrevTitLog))
            sNomeLog = Trim(SubNull(!NomeLogradouro))
            nNumImovel = Val(SubNull(!Numero))
            sCompl = SubNull(!Complemento)
            sBairro = SubNull(!DescBairro)
            sCep = RetornaNumero(RetornaCEP(!CodLogradouro, !Numero))
            sCidade = SubNull(!descCidade)
            sUF = SubNull(!SiglaUF)
            sDDD = SubNull(!ddd_nf)
            sFone = Left(SubNull(!telefone_nf), 15)

'            sFone = Left(RetornaNumero(SubNull(!fonecontato)), 15)
            sFax = Left(RetornaNumero(SubNull(!faxcontato)), 15)
            sEmail = SubNull(!emailcontato)
            
            sql = "select codtributo from mobiliarioatividadeiss where codmobiliario=" & nCodigo
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount > 0 Then
                If RdoAux2!CodTributo = 11 Then
                    sRegime = "F"
                ElseIf RdoAux2!CodTributo = 12 Then
                    sRegime = "E"
                ElseIf RdoAux2!CodTributo = 13 Then
                    sRegime = "V"
                Else
                    sRegime = "N"
                End If
            Else
                sRegime = "N"
            End If
            RdoAux2.Close
            If sRegime = "V" Then sRegime = "A"
            If sRegime = "E" Then segime = "T"
            nArea = IIf(IsNull(!areatl), 0, !areatl)
           .Close
           
        End With
        
        sql = "insert tb_inter_empresas(cod_cliente,num_cadastro,timestamp,inscricao,inscricao_estadual,nome_empresa,nome_fantasia,"
        sql = sql & "num_processo,tipo_empresa,cpf_cnpj,data_abertura,data_encerramento,tipo_logradouro,titulo_logradouro,logradouro,"
        sql = sql & "num_imovel,complemento,bairro,cep,cidade,estado,ddd,telefone,fax,email,regime_empresa,status_empresa,classificacao,area_ocupada) "
        sql = sql & "values(2177," & nCodigo & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "'," & nCodigo & "," & IIf(Val(sIE) > 0, Val(sIE), "Null") & ",'" & Mask(sRazao) & "',"
        sql = sql & IIf(sFantasia <> "", "'" & Mask(sFantasia) & "'", "Null") & "," & IIf(sNumProcesso <> "", "'" & sNumProcesso & "'", "Null") & ",'" & sTipoEmpresa & "'," & IIf(Val(sDoc) > 0, Val(sDoc), "Null") & ",'" & Format(sDataAbertura, "m/dd/yyyy") & "',"
        sql = sql & IIf(IsDate(sDataEncerramento), "'" & Format(sDataEncerramento, "mm/dd/yyyy") & "'", "Null") & "," & IIf(sTipoLog <> "", "'" & sTipoLog & "'", "Null") & ","
        sql = sql & IIf(sTitLog <> "", "'" & sTitLog & "'", "Null") & ",'" & Mask(sNomeLog) & "'," & IIf(nNumImovel > 0, "'" & CStr(nNumImovel) & "'", "Null") & "," & IIf(sCompl <> "", "'" & Left(sCompl, 40) & "'", "Null") & ",'"
        sql = sql & Mask(sBairro) & "'," & IIf(Val(sCep) > 0, Val(sCep), "Null") & ",'" & sCidade & "','" & sUF & "'," & IIf(Val(sDDD) > 0, Val(sDDD), "Null") & "," & IIf(Val(sFone) > 0, Val(sFone), "Null") & "," & IIf(Val(sFax) > 0, Val(sFax), "Null") & "," & IIf(sEmail <> "", "'" & sEmail & "'", "Null") & ","
        sql = sql & IIf(sRegime <> "", "'" & sRegime & "'", "Null") & ",'" & IIf(IsDate(sDataEncerramento), "E", "A") & "','" & "N" & "'," & RetornaNumero(FormatNumber(nArea, 2)) & ")"
        cnEicon.Execute sql, rdExecDirect
        
       '******* DADOS DOS SÓCIOS **************************
        sql = "SELECT * FROM mobiliarioproprietario Where mobiliarioproprietario.codmobiliario = " & nCodigo
        Set RdoProp = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoProp
            Do Until .EOF
                nCodCidadao = !CodCidadao
                sql = "select codigo from eicon_socio where codigo=" & nCodCidadao
                Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                If RdoAux2.RowCount = 0 Then
                    sql = "insert eicon_socio(codigo) values(" & nCodCidadao & ")"
                    cn.Execute sql, rdExecDirect
                    AtualizaSocio
                End If
                RdoAux2.Close
               .MoveNext
            Loop
            RdoProp.Close
        End With
                
        
        
        DoEvents
        RdoAux.MoveNext
    Loop
    RdoAux.Close
End With

sql = "delete from eicon_empresa"
'cn.Execute Sql, rdExecDirect

cnEicon.Close
MsgBox "fim"
End Sub

Public Sub AtualizaSocio()

Dim sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset
Dim nCodigo As Long, sTipoLog As String, sTitLog As String, sNomeLog As String, nCodEmpresa As Long, sNome As String, nCodLogr As Long, nTipoEnd As String
Dim nNumImovel As Integer, sCompl As String, sBairro As String, sCep As String, sCidade As String, sUF As String, sFone As String, sFax As String, sEmail As String

ConectaEicon2

'****REMOVER NO FINAL DOS TESTES *****
sql = "delete from tb_inter_socios"
'cnEicon2.Execute sql, rdExecDirect
'*************************************

sql = "select codigo from eicon_socio order by codigo"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
        Do Until .EOF
        nCodigo = !Codigo
        
       '******* DADOS DOS SÓCIOS **************************
        sql = "SELECT vwfullcidadao.*,codmobiliario FROM vwFULLCIDADAO inner join mobiliarioproprietario on vwfullcidadao.codcidadao=mobiliarioproprietario.codcidadao where mobiliarioproprietario.codcidadao = " & nCodigo
        Set RdoProp = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoProp
            Do Until .EOF
                nCodEmpresa = !codmobiliario
                sNome = !nomecidadao
                sql = "select num_cadastro from tb_inter_socios where num_cadastro=" & nCodEmpresa & " and nome_socio='" & sNome & "' and controle is null"
                Set RdoAux2 = cnEicon2.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                If RdoAux2.RowCount > 0 Then
                    RdoAux2.Close
                    GoTo Proximo
                End If
                RdoAux2.Close
                
                sDoc = RetornaNumero(SubNull(!cpf))
                
                If Not IsNull(etiqueta2) Then
                    sTipoEnd = "C"
                Else
                    sTipoEnd = "R"
                End If
                                
                If sTipoEnd = "R" Then
                    sTipoLog = SubNull(!AbrevTipoLog)
                    sTitLog = SubNull(!AbrevTitLog)
                    sNomeLog = SubNull(!NomeLogradouro)
                    If sNomeLog = "" Then
                        sNomeLog = SubNull(!NOMELOGRADOURO2)
                    End If
                    nNumImovel = Val(SubNull(!NUMIMOVEL))
                    sCompl = SubNull(!Complemento)
                    sBairro = SubNull(!DescBairro)
                    nCodLogr = Val(SubNull(!CodLogradouro))
                    If nCodLogr > 0 Then
                        sCep = RetornaNumero(RetornaCEP(nCodLogr, nNumImovel))
                    Else
                        sCep = RetornaNumero(SubNull(!Cep))
                    End If
                    sCidade = SubNull(!descidade)
                    sUF = SubNull(!SiglaUF)
                    sFone = SubNull(!telefone)
                    sEmail = SubNull(!Email)
                Else
                    sTipoLog = SubNull(!AbrevTipoLogC)
                    sTitLog = SubNull(!AbrevTitLogC)
                    sNomeLog = SubNull(!NomeLogradouroC)
'                    If sNomeLog = "" Then
'                        sNomeLog = SubNull(!NOMELOGRADOURO2)
'                    End If
                    nNumImovel = Val(SubNull(!NUMIMOVEL2))
                    sCompl = SubNull(!Complemento2)
                    sBairro = SubNull(!DescBairroC)
                    nCodLogr = Val(SubNull(!CodLogradouro2))
                    If nCodLogr > 0 Then
                        sCep = RetornaNumero(RetornaCEP(nCodLogr, nNumImovel))
                    Else
                        sCep = RetornaNumero(SubNull(!Cep2))
                    End If
                    sCidade = SubNull(!desccidadeC)
                    sUF = SubNull(!SiglaUF2)
                    sFone = SubNull(!Telefone2)
                    sEmail = SubNull(!EMAIL2)
                End If
                
                sql = "insert tb_inter_socios(cod_cliente,num_cadastro,inscricao,cod_socio,nome_socio,timestamp,cpf,tipo_logradouro,titulo_logradouro,logradouro,num_imovel,complemento,bairro,cep,"
                sql = sql & "cidade,estado,telefone,email) "
                sql = sql & "values(2177," & nCodEmpresa & "," & nCodEmpresa & "," & nCodigo & ",'" & Mask(sNome) & "','" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "'," & IIf(Val(sDoc) > 0, Val(sDoc), "Null") & ","
                sql = sql & IIf(sTipoLog <> "", "'" & sTipoLog & "'", "Null") & "," & IIf(sTitLog <> "", "'" & sTitLog & "'", "Null") & ",'" & Mask(sNomeLog) & "'," & IIf(nNumImovel > 0, "'" & CStr(nNumImovel) & "'", "Null") & ","
                sql = sql & IIf(sCompl <> "", "'" & Left(sCompl, 40) & "'", "Null") & "," & IIf(sBairro <> "", "'" & sBairro & "'", "Null") & "," & IIf(Val(sCep) > 0, Val(sCep), "Null") & ",'" & sCidade & "','" & sUF & "',"
                sql = sql & IIf(Val(sFone) > 0, Val(sFone), "Null") & "," & IIf(sEmail <> "", "'" & sEmail & "'", "Null") & ")"
                cnEicon2.Execute sql, rdExecDirect
                 
                sql = "delete from eicon_socio where codigo=" & nCodigo
                cn.Execute sql, rdExecDirect
                 
               .MoveNext
            Loop
            RdoProp.Close
        End With
        
Proximo:
        DoEvents
        RdoAux.MoveNext
    Loop
    RdoAux.Close
End With


cnEicon2.Close


End Sub
