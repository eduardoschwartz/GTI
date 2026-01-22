VERSION 5.00
Begin VB.Form frmPrintTermica 
   Caption         =   "Form1"
   ClientHeight    =   1965
   ClientLeft      =   8085
   ClientTop       =   6765
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   ScaleHeight     =   1965
   ScaleWidth      =   5850
End
Attribute VB_Name = "frmPrintTermica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Dim p As Printer
defprinter = Printer.DeviceName  ' save name of default printer in case you need to change back
For Each p In Printers
  If UCase(Left(Printer.DeviceName, 5)) = "TANCA" Then
    Set Printer = p
    Exit For
  End If
Next

With Printer

    .ScaleMode = 2 'Point
    .FontName = "Arial"
    .FontBold = True
    .FontSize = 20
    .CurrentX = 19
    .CurrentY = 20
    Printer.Print "My Title"
    .EndDoc
End With
Unload Me
frmSenhaPre.show
End Sub

