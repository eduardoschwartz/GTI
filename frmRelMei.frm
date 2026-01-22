VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmRelMei 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório do MEI"
   ClientHeight    =   1650
   ClientLeft      =   6900
   ClientTop       =   3240
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   5085
   Begin VB.OptionButton opt 
      Caption         =   "Empresas que SAIRAM do MEI"
      Height          =   240
      Index           =   1
      Left            =   270
      TabIndex        =   1
      Top             =   585
      Width           =   3210
   End
   Begin VB.OptionButton opt 
      Caption         =   "Empresas que ENTRARAM no MEI"
      Height          =   240
      Index           =   0
      Left            =   270
      TabIndex        =   0
      Top             =   225
      Value           =   -1  'True
      Width           =   3210
   End
   Begin esMaskEdit.esMaskedEdit mskDataIni 
      Height          =   285
      Left            =   1350
      TabIndex        =   2
      Top             =   1125
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   503
      MouseIcon       =   "frmRelMei.frx":0000
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
   Begin esMaskEdit.esMaskedEdit mskDataFim 
      Height          =   285
      Left            =   3750
      TabIndex        =   3
      Top             =   1140
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   503
      MouseIcon       =   "frmRelMei.frx":001C
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
   Begin prjChameleon.chameleonButton cmdPrint 
      Default         =   -1  'True
      Height          =   360
      Left            =   3690
      TabIndex        =   4
      ToolTipText     =   "Gerar o relatório"
      Top             =   450
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "Gerar"
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
      MICON           =   "frmRelMei.frx":0038
      PICN            =   "frmRelMei.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Fim.....:"
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   2730
      TabIndex        =   6
      Top             =   1185
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Início..:"
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   1
      Left            =   315
      TabIndex        =   5
      Top             =   1170
      Width           =   1035
   End
End
Attribute VB_Name = "frmRelMei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrint_Click()

If Not IsDate(mskDataIni.Text) Then
    MsgBox "Data de Inicio inválido", vbExclamation, "atenção"
    Exit Sub
End If

If Not IsDate(mskDataFim.Text) Then
    MsgBox "Data de Fim inválido", vbExclamation, "atenção"
    Exit Sub
End If

If CDate(mskDataIni.Text) > CDate(mskDataFim.Text) Then
    MsgBox "Data de Inicio tem que ser maior que data de termino", vbExclamation, "atenção"
    Exit Sub
End If

GeraRel

End Sub

Private Sub Form_Load()
Centraliza Me
End Sub

Private Sub mskDataFim_GotFocus()
mskDataFim.SelStart = 0
mskDataFim.SelLength = Len(mskDataFim.Text)

End Sub

Private Sub mskDataIni_GotFocus()
mskDataIni.SelStart = 0
mskDataIni.SelLength = Len(mskDataIni.Text)
End Sub

Private Sub GeraRel()

Dim sql As String, rdoAux As rdoResultset, bSaiu As Boolean, FF1 As Integer, sReg As String, sNomeArq As String
FF1 = FreeFile()
sNomeArq = sPathBin & "\ListaMei.txt"
Open sNomeArq For Output As FF1

bSaiu = opt(1).value

If bSaiu Then
    sql = "select id,codigo,datainicio,datafim,razaosocial from periodomei INNER JOIN mobiliario on periodomei.codigo = mobiliario.codigomob "
    sql = sql & " where datafim between '" & Format(mskDataIni.Text, "mm/dd/yyyy") & "' and '" & Format(mskDataFim.Text, "mm/dd/yyyy") & "'"
Else
    sql = "select id,codigo,datainicio,datafim,razaosocial from periodomei INNER JOIN mobiliario on periodomei.codigo = mobiliario.codigomob "
    sql = sql & " where datainicio between '" & Format(mskDataIni.Text, "mm/dd/yyyy") & "' and '" & Format(mskDataFim.Text, "mm/dd/yyyy") & "'"
End If
Set rdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    If .RowCount = 0 Then
        If bSaiu Then
            MsgBox "Nenhuma empresa saiu do MEI no período informado.", vbInformation, "Atenção"
        Else
            MsgBox "Nenhuma empresa entrou no MEI no período informado.", vbInformation, "Atenção"
        End If
        Exit Sub
    Else
        If bSaiu Then
            sReg = "EMPRESAS QUE SAIRAM DO MEI ENTRE " & mskDataIni.Text & " E " & mskDataFim.Text
        Else
            sReg = "EMPRESAS QUE ENTRARAM DO MEI ENTRE " & mskDataIni.Text & " E " & mskDataFim.Text
        End If
    End If
    Print #FF1, sReg
    If Not bSaiu Then
        sReg = "CODIGO RAZAO SOCIAL                                        DT.INICIO           "
    Else
        sReg = "CODIGO RAZAO SOCIAL                                        DT.INICIO   DATA FIM"
    End If
    Print #FF1, sReg
    sReg = "================================================================================="
    Print #FF1, sReg
    
    Do Until .EOF
        sReg = !Codigo
        If IsNull(!datafim) Or Not bSaiu Then
            Print #FF1, sReg & " " & FillSpace(Left(!RazaoSocial, 50), 50) & "  " & Format(!datainicio, "dd/mm/yyyy")
        Else
            Print #FF1, sReg & " " & FillSpace(Left(!RazaoSocial, 50), 50) & "  " & Format(!datainicio, "dd/mm/yyyy") & "  " & Format(!datafim, "dd/mm/yyyy")
        End If
       .MoveNext
    Loop
   .Close
End With
Close #FF1

ret = Shell("NOTEPAD" & " " & sNomeArq, vbNormalFocus)


End Sub
