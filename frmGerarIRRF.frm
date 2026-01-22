VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmGerarIRRF 
   Caption         =   "Gerar lançamentos de IRRF"
   ClientHeight    =   7755
   ClientLeft      =   9780
   ClientTop       =   5115
   ClientWidth     =   14430
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7755
   ScaleWidth      =   14430
   Begin VB.Frame frTop 
      Height          =   600
      Left            =   45
      TabIndex        =   0
      Top             =   -45
      Width           =   14370
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   7170
      Left            =   45
      TabIndex        =   1
      Top             =   585
      Width           =   14370
      _ExtentX        =   25347
      _ExtentY        =   12647
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Id"
         Object.Width           =   35
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Nº Empenho"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Data Emp."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Nº Liquid."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Data Liq."
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmGerarIRRF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Me.Height = Val(GetSetting("GTI", "WINDOW", "IRRF_HEIGHT"))
If Me.Height < 600 Then
   SaveSetting "GTI", "WINDOW", "IRRF_HEIGHT", 6585
   Me.Height = GetSetting("GTI", "WINDOW", "IRRF_HEIGHT")
End If
Me.Width = Val(GetSetting("GTI", "WINDOW", "IRRF_WIDTH"))
If Me.Width < 2000 Then
   SaveSetting "GTI", "WINDOW", "IRRF_WIDTH", 11655
   Me.Width = GetSetting("GTI", "WINDOW", "IRRF_WIDTH")
End If
Centraliza Me
Carrega_Dados

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

SaveSetting "GTI", "WINDOW", "IRRF_HEIGHT", Me.Height
SaveSetting "GTI", "WINDOW", "IRRF_WIDTH", Me.Width

End Sub

Private Sub Form_Resize()
If Me.Width < 1500 Or Me.Height < 1700 Then Exit Sub
lvMain.Width = Me.Width - 220
lvMain.Height = Me.Height - 1100
frTop.Width = Me.Width - 220

End Sub

Private Sub Carrega_Dados()
Dim RdoAux As rdoResultset, Sql As String, RdoAux2 As rdoResultset, sNumeroEmp As String, sNumeroLiq As String
Dim x As Integer

Ocupado
ConectaSmar
Dim itmX As ListItem
lvMain.ListItems.Clear
x = 1
Sql = "SELECT TOP(30) * FROM Liquidacoes_IRRF  WHERE YEAR( Data_EMPENHO) =2024  ORDER BY Numero_Empenho"
Set RdoAux = cnSmar.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        sNumeroEmp = Format(!numero_empenho, "00000") & "/" & !ano_empenho
        sNumeroLiq = Format(!numero_liquidacao, "00000") & "/" & !ano_liquidacao
        Set itmX = lvMain.ListItems.Add(, , , Format(x, "0000"))
        itmX.SubItems(1) = sNumeroEmp
        itmX.SubItems(2) = Format(!data_empenho, "dd/mm/yyyy")
        itmX.SubItems(3) = sNumeroLiq
        itmX.SubItems(4) = Format(!data_liquidacao, "dd/mm/yyyy")
        x = x + 1
       .MoveNext
    Loop
   .Close
End With

Liberado
cnSmar.Close

End Sub

