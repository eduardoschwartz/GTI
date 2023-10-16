VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmEncerrarEmpresa 
   BackColor       =   &H008080FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Encerrar empresa"
   ClientHeight    =   1695
   ClientLeft      =   11355
   ClientTop       =   7485
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtProcesso 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2130
      MaxLength       =   15
      TabIndex        =   1
      Top             =   720
      Width           =   1425
   End
   Begin esMaskEdit.esMaskedEdit mskData 
      Height          =   285
      Left            =   2130
      TabIndex        =   0
      Top             =   360
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   503
      MouseIcon       =   "frmEncerrarEmpresa.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "99/99/9999"
      SelText         =   ""
      Text            =   "__/__/____"
      HideSelection   =   -1  'True
   End
   Begin prjChameleon.chameleonButton btGravar 
      Height          =   360
      Left            =   1080
      TabIndex        =   4
      Top             =   1260
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "Gravar"
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
      MCOL            =   0
      MPTR            =   1
      MICON           =   "frmEncerrarEmpresa.frx":001C
      PICN            =   "frmEncerrarEmpresa.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton btCancel 
      Height          =   360
      Left            =   2340
      TabIndex        =   5
      Top             =   1260
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "Cancelar"
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
      MICON           =   "frmEncerrarEmpresa.frx":045C
      PICN            =   "frmEncerrarEmpresa.frx":0478
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Data Encerramento.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Index           =   8
      Left            =   270
      TabIndex        =   3
      Top             =   420
      Width           =   1845
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nº do Processo.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   10
      Left            =   285
      TabIndex        =   2
      Top             =   780
      Width           =   1485
   End
End
Attribute VB_Name = "frmEncerrarEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btCancel_Click()
Unload Me
End Sub

Private Sub btGravar_Click()
Dim sValida As String, Sql As String, sDataProcesso As String, nCodigo As Long, nSeq As Integer, sHist As String, RdoAux As rdoResultset

If Not IsDate(mskData.Text) Then
    MsgBox "Data de encerramento inválida", vbCritical, "Erro"
    Exit Sub
End If


If Trim(txtProcesso.Text) = "" Then
    MsgBox "Nº do processo inválido", vbCritical, "Erro"
    Exit Sub
End If

sValida = ValidaProcesso(txtProcesso.Text)
If sValida <> "OK" Then
    MsgBox sValida, vbCritical, "Erro"
    Exit Sub
End If

nCodigo = Val(frmCadMob.txtCodEmpresa.Text)

If MsgBox("Deseja encerrar esta empresa?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmação") = vbNo Then Exit Sub

sDataProcesso = Format(RetornaDataProcesso(Val(Left$(txtProcesso.Text, Len(txtProcesso.Text) - 5)), Val(Right$(txtProcesso.Text, 4))), "dd/mm/yyyy")
frmCadMob.mskDataEn.Text = mskData.Text
frmCadMob.txtNumProcE.Text = txtProcesso.Text
frmCadMob.mskDataPEn.Text = sDataProcesso

Sql = "update mobiliario set dataencerramento='" & Format(mskData.Text, "mm/dd/yyyy") & "', numprocencerramento='" & txtProcesso.Text & "', "
Sql = Sql & "dataprocencerramento='" & Format(sDataProcesso, "mm/dd/yyyy") & "' where codigomob=" & nCodigo
cn.Execute Sql, rdExecDirect

Sql = "select max(seq) as maximo from mobiliariohist where codmobiliario=" & nCodigo
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!maximo) Then
    nSeq = 0
Else
    nSeq = RdoAux!maximo + 1
End If

sHist = "Empresa encerrada"

Sql = "INSERT MOBILIARIOHIST(CODMOBILIARIO,SEQ,DATAHIST,OBS,USERID) VALUES("
Sql = Sql & nCodigo & "," & nSeq & ",'" & Format(Now, "mm/dd/yyyy") & "','" & Mask(sHist) & "'," & RetornaUsuarioID(Mask(NomeDeLogin)) & ")"
cn.Execute Sql, rdExecDirect

'Integração_Eicon
Sql = "select codigo from eicon_empresa where codigo=" & nCodigo
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount = 0 Then
    Sql = "insert eicon_empresa(codigo) values(" & nCodigo & ")"
    cn.Execute Sql, rdExecDirect
End If
RdoAux.Close


Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
mskData.Text = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub mskData_GotFocus()
mskData.SelStart = 0
mskData.SelLength = Len(mskData.Text)
mskData.SetFocus
End Sub

