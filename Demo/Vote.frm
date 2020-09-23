VERSION 5.00
Begin VB.Form VDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Thanks for appreciating this code...."
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Ok 
      Caption         =   "&Go to Article Page"
      Default         =   -1  'True
      Height          =   390
      Left            =   2625
      TabIndex        =   0
      Top             =   1245
      Width           =   1950
   End
   Begin VB.Label Label2 
      Caption         =   $"Vote.frx":0000
      Height          =   810
      Left            =   150
      TabIndex        =   2
      Top             =   300
      Width           =   4365
   End
   Begin VB.Label Label1 
      Caption         =   $"Vote.frx":00AA
      Height          =   660
      Left            =   30
      TabIndex        =   1
      Top             =   1680
      Width           =   4635
   End
End
Attribute VB_Name = "VDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Const SW_SHOWNORMAL = 1
Const CodeID = 26900

Private Sub Form_Load()
Form1.ShowModalForm.Enabled = False
Form1.ShowForm.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.ShowModalForm.Enabled = True
Form1.ShowForm.Enabled = True

End Sub

Private Sub Ok_Click()
    GotoURL ("http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=26900&lngWId=1")
End Sub



Sub GotoURL(URL As String)
    Dim Res As Long
    Dim TFile As String, Browser As String, Dum As String
    
    TFile = App.Path + "\test.htm"
    Open TFile For Output As #1
    Close
    Browser = String(255, " ")
    Res = FindExecutable(TFile, Dum, Browser)
    Browser = Trim$(Browser)
    
    If Len(Browser) = 0 Then
        MsgBox "Cannot find browser"
        Exit Sub
    End If
    
    Res = ShellExecute(Me.hwnd, "open", Browser, URL, Dum, SW_SHOWNORMAL)
    If Res <= 32 Then
        MsgBox "Cannot open web page"
        Exit Sub
    End If
End Sub


