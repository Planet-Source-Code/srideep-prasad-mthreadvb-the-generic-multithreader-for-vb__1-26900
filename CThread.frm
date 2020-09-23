VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TH As Long, TID As Long
Dim TDat As TData
Dim A As VBCONSOLE.CONSOLE

Private Sub Form_Click()
TDat.CLASSID = CLSIDFromProgID("VBConsole.Console")
TDat.Ptr = ObjPtr(A)
TH = CreateThread(0, 0, AddressOf TProc, VarPtr(TDat), 0, TID)

End Sub

Sub Hello()
MsgBox "Hello"
End Sub
