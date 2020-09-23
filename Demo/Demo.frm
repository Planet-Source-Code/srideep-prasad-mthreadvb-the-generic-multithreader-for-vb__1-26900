VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MThreadVB Multithreader For VB - Prime Finder Demo Application"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   7650
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton ShowModalForm 
      Caption         =   "Show &Modal Form"
      Height          =   345
      Left            =   3045
      TabIndex        =   19
      Top             =   5370
      Width           =   3345
   End
   Begin VB.CommandButton ShowForm 
      Caption         =   "Show &Non Modal Form from Thread"
      Height          =   345
      Left            =   -15
      TabIndex        =   18
      Top             =   5370
      Width           =   2955
   End
   Begin VB.CommandButton VDlg 
      Caption         =   "&Show voting Dialog..."
      Height          =   390
      Left            =   5730
      TabIndex        =   17
      Top             =   4425
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear Log"
      Height          =   330
      Left            =   45
      TabIndex        =   14
      Top             =   4395
      Width           =   1845
   End
   Begin VB.Frame Frame4 
      Caption         =   "Primes Found"
      Height          =   2040
      Left            =   5745
      TabIndex        =   12
      Top             =   2295
      Width           =   1920
      Begin VB.TextBox PText 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   1740
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   13
         Top             =   225
         Width           =   1740
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Event log"
      Height          =   2040
      Left            =   60
      TabIndex        =   10
      Top             =   2295
      Width           =   5250
      Begin VB.TextBox ELog 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   1695
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   11
         Top             =   240
         Width           =   5085
      End
   End
   Begin VB.CommandButton EndThread 
      Caption         =   "Terminate Prime Finder Thread..."
      Enabled         =   0   'False
      Height          =   345
      Left            =   5100
      TabIndex        =   9
      Top             =   1950
      Width           =   2520
   End
   Begin VB.CommandButton StartThread 
      Caption         =   "Start Prime Finder Thread..."
      Height          =   360
      Left            =   2700
      TabIndex        =   8
      Top             =   1935
      Width           =   2265
   End
   Begin VB.Frame Frame2 
      Caption         =   "Parameters"
      Height          =   690
      Left            =   30
      TabIndex        =   5
      Top             =   1260
      Width           =   7590
      Begin VB.ComboBox Pr 
         Height          =   315
         Left            =   690
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   255
         Width           =   4005
      End
      Begin VB.Label Label3 
         Caption         =   "Priority:"
         Height          =   180
         Left            =   135
         TabIndex        =   7
         Top             =   315
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Find Primes:"
      Height          =   720
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   7575
      Begin VB.TextBox EndBox 
         Height          =   285
         Left            =   3945
         TabIndex        =   4
         Text            =   "5000"
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox Start 
         Height          =   285
         Left            =   600
         TabIndex        =   2
         Text            =   "2"
         Top             =   255
         Width           =   2610
      End
      Begin VB.Label Label2 
         Caption         =   "To:"
         Height          =   165
         Left            =   3570
         TabIndex        =   3
         Top             =   285
         Width           =   315
      End
      Begin VB.Label Label1 
         Caption         =   "From:"
         Height          =   180
         Left            =   150
         TabIndex        =   1
         Top             =   300
         Width           =   465
      End
   End
   Begin VB.Label Label7 
      Caption         =   $"Demo.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   15
      TabIndex        =   21
      Top             =   5760
      Width           =   7635
   End
   Begin VB.Label Label6 
      Caption         =   "It seems many of you are having problems showing forms from Threads.... So I have added a ""Form Show"" Demonstration"
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   4935
      Width           =   7425
   End
   Begin VB.Label Label5 
      Caption         =   $"Demo.frx":0128
      Height          =   450
      Left            =   15
      TabIndex        =   16
      Top             =   840
      Width           =   7590
   End
   Begin VB.Label Label4 
      Caption         =   "If you believe this article deserves it, please click the button alongside to get the voting dialog box "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   2145
      TabIndex        =   15
      Top             =   4365
      Width           =   3240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NOTE:MThreadVB does not use the standard apartment model
'threading... As a result global and module level variables
'are now accessible to the thread and they do not remain
'hidden

'Though you can "officially" pass only a variant a parameter
'this really does not hold true in this case
'In this example for instance, the module level
'variables Min and Max are
'the actual parameters which are accessed by the thread
'procedure FindPrimes(). Since VB requires it, the
'multithreaded Sub must also have a single variant argument
'Here this argument is just a dummy though and the actual
'parameters are the variables Min and Max that are
'accessed by the Thread Procedure
'Also the variable Primes holds the prime numbers that
'are found (The output variable)


Dim WithEvents FormThread As MThreadVB.Thread
Attribute FormThread.VB_VarHelpID = -1
Dim WithEvents pThread As MThreadVB.Thread
Attribute pThread.VB_VarHelpID = -1
Dim Min As Long, Max As Long
Dim Primes As String
Dim Frm As New VDialog
Private Sub Picture1_Click()

End Sub

Private Sub Command1_Click()
ELog.Text = ""
End Sub

Private Sub EndThread_Click()
    pThread.TerminateWin32Thread
End Sub

Private Sub Form_Load()
'Instantiate a reference to the multithreader library
Set pThread = New Thread
Set FormThread = New Thread

Pr.AddItem "Lowest"
Pr.AddItem "Below Normal"
Pr.AddItem "Normal"
Pr.AddItem "Above Normal"
Pr.AddItem "Highest"
Pr.ListIndex = 2
End Sub
Sub FindPrimes(DummyArgument As Variant)
'This is the actual multithreaded sub.
'It is called by MthreadVB and made to run on a different
'thread
'All multithreaded subs must accept a variant as an argument
'In this case the argument is a dummy, and the actual
'arguments are Min and Max which are set when the
'Start Prime Thread is clicked
'The variable Primes contains the prime numbers
'Since it is a module level variable it can be accessed by
'any sub in this module
Dim PCnt As Long, I As Long, J As Long
If Max = 0 Then GoTo 20
For I = Min To Max
    If I = 0 Then GoTo 10
    If I = 1 Then GoTo 10
    For J = 2 To I - 1
        If I Mod J = 0 Then
            GoTo 10
        End If
    Next J
    Primes = Primes + CStr(I) + Chr$(13) + Chr$(10)
10 Next I
20 End Sub

Private Sub Form_Unload(Cancel As Integer)
    If pThread.IsThreadRunning = True Or FormThread.IsThreadRunning = True Then
        MsgBox "A thread is running. Terminate it before quitting"
        Cancel = True
    End If
    
    Set pThread = Nothing
    Set FormThread = Nothing
    'VERY IMPORTANT - MUST EXPLICITLY CALL END TO TERMINATE APP
    End
End Sub



Private Sub Pr_Click()
If pThread.IsThreadRunning = True Then
Select Case Pr.List(Pr.ListIndex)
    Case "Lowest"
        pThread.ThreadPriority = THREAD_PRIORITY_LOWEST
    Case "Below Normal"
        pThread.ThreadPriority = THREAD_PRIORITY_BELOW_NORMAL
    Case "Normal"
        pThread.ThreadPriority = THREAD_PRIORITY_NORMAL
    Case "Above Normal"
        pThread.ThreadPriority = THREAD_PRIORITY_ABOVE_NORMAL
    Case "Highest"
        pThread.ThreadPriority = THREAD_PRIORITY_HIGHEST
End Select
End If
End Sub

Private Sub pThread_OnThreadCreateFailure()

    ELog.Text = ELog.Text & Chr$(13) & Chr$(10) & "Thread could not be Created"
End Sub
     
Private Sub pThread_OnThreadCreateSuccess(ByVal ThreadHandle As Long, ByVal ThreadID As Long)
    ELog.Text = ELog.Text & Chr$(13) & Chr$(10) & "Thread Created (Calculations started)"
    StartThread.Enabled = False
    EndThread.Enabled = True
End Sub

Private Sub pThread_OnThreadFinish(ByVal ThreadHandle As Long, ByVal ThreadID As Long)
    ELog.Text = ELog.Text & Chr$(13) & Chr$(10) & "Thread has finished running(Calculations Ended)"
    PText.Text = ""
    PText.Text = Primes
    Primes = ""
    StartThread.Enabled = True
    EndThread.Enabled = False
End Sub

Private Sub pThread_OnThreadPriorityChange(ByVal ThreadHandle As Long, ThreadID As Long, ByVal OldPriority As MThreadVB.ThreadPriorityConsts, ByVal NewPriority As MThreadVB.ThreadPriorityConsts)
    ELog.Text = ELog.Text & Chr$(13) & Chr$(10) & "Thread priority set or changed"
End Sub

Private Sub pThread_OnThreadTerminate(ByVal ThreadHandle As Long, ByVal ThreadID As Long, ByVal ExitCode As Long)
    ELog.Text = ELog.Text & Chr$(13) & Chr$(10) & "Thread has been forcefully terminated"
    StartThread.Enabled = True
    EndThread.Enabled = False

End Sub

Private Sub ShowForm_Click()
ShowModalForm.Enabled = False
ShowForm.Enabled = False
If FormThread.IsThreadRunning = False Then
'If no thread is launched by the FormThread object,
'we launch a new Thread.... The multithreaded procedure
'is called ShowFormMT, and here we actually pass
'a parameter (0) that determines whether the Form
'is modal or not
    FormThread.CreateWin32Thread Me, "ShowFormMT", 0
End If
End Sub

Sub ShowFormMT(ModalFlag As Variant)
'This is the multithreaded procedure... We cannot
'directly load a form here, since doing so will cause
'VB to crash... We must call the Main Thread and make
'it perform the Form load operation...
'The ObjectInThreadContext property returns a reference
'to the original Form (Form1) running on the original
'Thread... We call the ShowFormNow() Sub to show the form
'But this Sub is called in context to the original Thread
FormThread.ObjectInThreadContext.ShowFormNow (CLng(ModalFlag))

'REMEMBER:Here in this case, ModalFlag is a "real"
'parameter unlike in the FindPrimes() Sub
End Sub

Private Sub ShowModalForm_Click()
If FormThread.IsThreadRunning = False Then
'If no thread is launched by the FormThread object,
'we launch a new Thread.... The multithreaded procedure
'is called ShowFormMT, and here we actually pass
'a parameter (1) that determines whether the Form
'is modal or not (It is modal in this case)
    
    FormThread.CreateWin32Thread Me, "ShowFormMT", 1
End If
End Sub

Sub ShowFormNow(ModalFlag As Long)
'This sub is called ultimately, and it ultimately loads
'the Form (ModalFlag is the parameter determining
'whether the form is modal or not)
VDialog.Show ModalFlag
End Sub

Private Sub StartThread_Click()
Min = CLng(Start.Text)
Max = CLng(EndBox.Text)
'This statement creates a thread. The second parameter is
'the Function that has be multithreaded. The first parameter
'is the Object on which the Function is defined...
pThread.CreateWin32Thread Me, "FindPrimes", 0
PText.Text = ""
Select Case Pr.List(Pr.ListIndex)
    Case "Lowest"
        pThread.ThreadPriority = THREAD_PRIORITY_LOWEST
    Case "Below normal"
        pThread.ThreadPriority = THREAD_PRIORITY_BELOW_NORMAL
    Case "Normal"
        pThread.ThreadPriority = THREAD_PRIORITY_NORMAL
    Case "Above Normal"
        pThread.ThreadPriority = THREAD_PRIORITY_ABOVE_NORMAL
    Case "Highest"
        pThread.ThreadPriority = THREAD_PRIORITY_HIGHEST
End Select
End Sub



Private Sub VDlg_Click()
    VDialog.Show 1
End Sub

