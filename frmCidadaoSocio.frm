VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmCidadaoSocio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inclusão de sócios (Pessoa Física) em empresas de fora "
   ClientHeight    =   3015
   ClientLeft      =   15195
   ClientTop       =   8790
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstMain 
      Appearance      =   0  'Flat
      Height          =   2370
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   7035
   End
   Begin prjChameleon.chameleonButton btAdd 
      Height          =   315
      Left            =   180
      TabIndex        =   0
      ToolTipText     =   "Adicionar Sócio"
      Top             =   2610
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Adicionar"
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
      MICON           =   "frmCidadaoSocio.frx":0000
      PICN            =   "frmCidadaoSocio.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton btDel 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      ToolTipText     =   "Remover Sócio"
      Top             =   2610
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Remover"
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
      MICON           =   "frmCidadaoSocio.frx":0176
      PICN            =   "frmCidadaoSocio.frx":0192
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmCidadaoSocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btAdd_Click()
Dim Sql As String, RdoAux As rdoResultset, z As Variant, nCodigo_Empresa As Long, nCodigo_Socio As Long, sNome As String

nCodigo_Empresa = Val(frmCidadao.txtCod.Text)
z = InputBox("Digite o código do sócio.", "Entrada de dados")
If Val(z) > 0 Then
    nCodigo_Socio = Val(z)
    If nCodigo_Socio = nCodigo_Empresa Then
        MsgBox "Código da empresa e do sócio não podem ser o mesmo.", vbCritical, "Erro"
        Exit Sub
    End If
        
    If nCodigo_Socio >= 500000 And nCodigo_Socio < 700000 Then
        Sql = "select * from cidadao_socio where codigo_empresa=" & nCodigo_Empresa & " and codigo_socio=" & nCodigo_Socio
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux.RowCount > 0 Then
            MsgBox "Este sócio já esta incluso nesta empresa.", vbCritical, "Erro"
            Exit Sub
        Else
            Sql = "select nomecidadao,cpf from cidadao where codcidadao=" & nCodigo_Socio
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux.RowCount = 0 Then
                MsgBox "Cidadão não cadastrado.", vbCritical, "Erro"
                Exit Sub
            Else
                If SubNull(RdoAux!cpf) = "" Then
                    MsgBox "Cidadão não possui cpf cadastrado.", vbCritical, "Erro"
                    Exit Sub
                End If
            End If
            sNome = RdoAux!Nomecidadao
            Sql = "insert cidadao_socio(codigo_empresa,codigo_socio) values(" & nCodigo_Empresa & "," & nCodigo_Socio & ")"
            cn.Execute Sql, rdExecDirect
            lstMain.AddItem nCodigo_Socio & " - " & sNome
        End If
    End If
  
End If


End Sub

Private Sub btDel_Click()
Dim Sql As String, nCodigo_Empresa As Long, nCodigo_Socio As Long

nCodigo_Empresa = Val(frmCidadao.txtCod.Text)
If lstMain.ListCount = 0 Then
    MsgBox "Selecione o sócio a ser excluído.", vbCritical, "Erro"
    Exit Sub
End If

If lstMain.ListIndex = -1 Then
    MsgBox "Selecione o sócio a ser excluído.", vbCritical, "Erro"
    Exit Sub
End If

If MsgBox("Remover o sócio " + lstMain.Text + " desta empresa?", vbYesNo + vbQuestion, "Confirmação") = vbYes Then
    nCodigo_Socio = Val(Left(lstMain.Text, 6))
    Sql = "delete from cidadao_socio where codigo_empresa=" & nCodigo_Empresa & " and codigo_socio=" & nCodigo_Socio
    cn.Execute Sql, rdExecDirect
    Carrega_Lista
End If

End Sub

Private Sub Form_Load()
Centraliza Me
Carrega_Lista
End Sub

Private Sub Carrega_Lista()
Dim Sql As String, RdoAux As rdoResultset, nCodigo_Empresa As Long, nCodigo_Socio As Long, sNome As String
lstMain.Clear
nCodigo_Empresa = Val(frmCidadao.txtCod.Text)

Sql = "SELECT cidadao_socio.codigo_socio,cidadao.nomecidadao From cidadao_socio INNER JOIN cidadao ON "
Sql = Sql & "cidadao_socio.codigo_socio = cidadao.codcidadao Where cidadao_socio.codigo_empresa = " & nCodigo_Empresa & " order by codigo_socio"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        lstMain.AddItem !Codigo_Socio & " - " & !Nomecidadao
       .MoveNext
    Loop
   .Close
End With

End Sub
