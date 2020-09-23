VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThreadWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Dim ThreadIDString As String
Dim wThreadHandle As Long, wThreadID As Long
Dim wThreadParam As Variant
Dim FName As String, ThreadObject As Object
Dim ThreadInterface As ThreadLaunchEX
Dim WithEvents WinThread As MThreadVB.Thread
Attribute WinThread.VB_VarHelpID = -1

Function InitThread(ThreadLaunch As ThreadLaunchEX, TObject As Object, FuncName As String, TParam As Variant, IDString As String) As Long
Set WinThread = New Thread
Set ThreadObject = TObject
FName = FuncName
ThreadParam = TParam
ThreadIDString = IDString
Set ThreadInterface = ThreadLaunch
InitThread = WinThread.CreateWin32Thread(ThreadObject, FName, ThreadParam)
End Function

Function GetThreadClass() As Thread
    Set GetThreadClass = WinThread
End Function

Function GetThreadID() As Long
GetThreadID = wThreadID
End Function

Function GetThreadHandle() As Long
GetThreadHandle = wThreadHandle
End Function

Function GetIDString() As String
    GetIDString = ThreadIDString
End Function

Private Sub Class_Terminate()
Set WinThread = Nothing
Set CallBackObj = Nothing
End Sub


Private Sub WinThread_OnThreadCreateFailure()
    RaiseEvent ThreadInterface.rOnThreadCreateFailure
End Sub

Private Sub WinThread_OnThreadCreateSuccess(ByVal ThreadHandle As Long, ByVal ThreadID As Long)
    Call ThreadInterface.rOnThreadCreateSuccess(ThreadIDString, ThreadHandle, ThreadID)
End Sub

Private Sub WinThread_OnThreadFinish(ByVal ThreadHandle As Long, ByVal ThreadID As Long)
    
    Call ThreadInterface.rOnThreadFinish(ThreadIDString, ThreadHandle, ThreadID)
    Set WinThread = Nothing
    Set ThreadInterface = Nothing
    ThreadIDString = 0
    wThreadHandle = 0
    wThreadID = 0
    wThreadParam = 0
End Sub


Private Sub WinThread_OnThreadPriorityChange(ByVal ThreadHandle As Long, ThreadID As Long, ByVal OldPriority As ThreadPriorityConsts, ByVal NewPriority As ThreadPriorityConsts)
    Call ThreadInterface.rOnThreadPriorityChange(ThreadIDString, ThreadHandle, ThreadID, OldPriority, NewPriority)
End Sub

Private Sub WinThread_OnThreadResume(ByVal ThreadHandle As Long, ByVal ThreadID As Long)
    Call ThreadInterface.rOnThreadResume(ThreadIDString, ThreadHandle, ThreadID)
End Sub

Private Sub WinThread_OnThreadStart(ByVal ThreadHandle As Long, ByVal ThreadID As Long)
    Call ThreadInterface.rOnThreadStart(ThreadIDString, ThreadHandle, ThreadID)
End Sub

Private Sub WinThread_OnThreadSuspend(ByVal ThreadHandle As Long, ByVal ThreadID As Long)
    Call ThreadInterface.rOnThreadSuspend(ThreadIDString, ThreadHandle, ThreadID)
End Sub

Private Sub WinThread_OnThreadTerminate(ByVal ThreadHandle As Long, ByVal ThreadID As Long, ByVal ExitCode As Long)
    Call ThreadInterface.rOnThreadTerminate(ThreadIDString, ThreadHandle, ThreadID, ExitCode)
    Set WinThread = Nothing
    Set ThreadInterface = Nothing
    ThreadIDString = 0
    wThreadHandle = 0
    wThreadID = 0
    wThreadParam = 0
End Sub


