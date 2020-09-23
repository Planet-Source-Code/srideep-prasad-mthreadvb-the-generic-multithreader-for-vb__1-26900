VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThreadLaunchEX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim Threads() As ThreadWrapper, TCount As Long
Public Event OnThreadCreateSuccess(ByVal ThreadIDString As String, ByVal ThreadHandle As Long, ByVal ThreadID As Long)
Public Event OnThreadCreateFailure()
Public Event OnThreadStart(ByVal ThreadIDString As String, ByVal ThreadHandle As Long, ByVal ThreadID As Long)
Public Event OnThreadFinish(ByVal ThreadIDString As String, ByVal ThreadHandle As Long, ByVal ThreadID As Long)
Public Event OnThreadTerminate(ByVal ThreadIDString As String, ByVal ThreadHandle As Long, ByVal ThreadID As Long, ExitCode As Long)
Public Event OnThreadSuspend(ByVal ThreadIDString As String, ByVal ThreadHandle As Long, ByVal ThreadID As Long)
Public Event OnThreadResume(ByVal ThreadIDString As String, ByVal ThreadHandle As Long, ByVal ThreadID As Long)
Public Event OnThreadPriorityChange(ByVal ThreadIDString As String, ByVal ThreadHandle As Long, ThreadID As Long, ByVal OldPriority As ThreadPriorityConsts, ByVal NewPriority As ThreadPriorityConsts)
Function CreateWin32Thread(IDString As String, ThreadObject As Object, ThreadFunc As String, Optional ThreadParam As Variant = 0) As Long
If DoesThreadExist(IDString) = True Then
    RaiseEvent OnThreadCreateFailure
End If
TCount = TCount + 1
ReDim Preserve Threads(TCount)
Set Threads(TCount) = New ThreadWrapper
CreateWin32Thread = Threads(TCount).InitThread(Me, ThreadObject, ThreadFunc, ThreadParam, IDString)
End Function

Friend Sub rOnThreadCreateSuccess(ByVal ThreadIDString As String, ByVal ThreadHandle As Long, ThreadID As Long)
    RaiseEvent OnThreadCreateSuccess(ThreadIDString, ThreadHandle, ThreadID)
End Sub

Friend Sub rOnThreadCreateFailure()
    RaiseEvent OnThreadCreateFailure
End Sub

Friend Sub rOnThreadStart(ByVal ThreadIDString As String, ByVal ThreadHandle As Long, ByVal ThreadID As Long)
    RaiseEvent OnThreadStart(ThreadIDString, ThreadHandle, ThreadID)
End Sub

Friend Sub rOnThreadFinish(ByVal ThreadIDString As String, ByVal ThreadHandle As Long, ByVal ThreadID As Long)
    RemThreadInfo (ThreadIDString)
    RaiseEvent OnThreadFinish(ThreadIDString, ThreadHandle, ThreadID)
End Sub

Friend Sub rOnThreadTerminate(ByVal ThreadIDString As String, ByVal ThreadHandle As Long, ByVal ThreadID As Long, ExitCode As Long)
    RemThreadInfo (ThreadIDString)
    RaiseEvent OnThreadTerminate(ThreadIDString, ThreadHandle, ThreadID, ExitCode)
End Sub

Friend Sub rOnThreadSuspend(ByVal ThreadIDString As String, ByVal ThreadHandle As Long, ByVal ThreadID As Long)
    RaiseEvent OnThreadSuspend(ThreadIDString, ThreadHandle, ThreadID)
End Sub

Friend Sub rOnThreadResume(ByVal ThreadIDString As String, ByVal ThreadHandle As Long, ByVal ThreadID As Long)
    RaiseEvent OnThreadResume(ThreadIDString, ThreadHandle, ThreadID)
End Sub

Friend Sub rOnThreadPriorityChange(ByVal ThreadIDString As String, ByVal ThreadHandle As Long, ThreadID As Long, ByVal OldPriority As ThreadPriorityConsts, ByVal NewPriority As ThreadPriorityConsts)
    RaiseEvent OnThreadPriorityChange(ThreadIDString, ThreadHandle, ThreadID, OldPriority, NewPriority)
End Sub

Function GetThreadID(ThreadIDString As String) As Long
Dim I As Long
    For I = 1 To TCount
        If Threads(I).GetIDString() = ThreadIDString Then
            GetThreadID = Threads(I).GetThreadClass.ThreadID
            Exit For
        End If
    Next I
End Function

Function GetThreadHandle(ThreadIDString As String) As Long
Dim I As Long
    For I = 1 To TCount
        If Threads(I).GetIDString() = ThreadIDString Then
            GetThreadHandle = Threads(I).GetThreadClass.ThreadHandle
            Exit For
        End If
    Next I
End Function

Property Get ThreadPriority(ThreadIDString As String) As ThreadPriorityConsts
Dim ThreadClass As Thread
Dim I As Long
    For I = 1 To TCount
        If Threads(I).GetIDString() = ThreadIDString Then
            Set ThreadClass = Threads(I).GetThreadClass
            Exit For
        End If
    Next I
ThreadPriority = ThreadClass.ThreadPriority
End Property

Property Get IsThreadRunning(ThreadIDString As String) As Boolean
If GetPos(ThreadIDString) <> 0 Then
    IsThreadRunning = Threads(GetPos(ThreadIDString)).GetThreadClass.IsThreadRunning
End If
End Property

Property Get ThreadPseudoHandle(ThreadIDString As String) As Long
If GetPos(ThreadIDString) <> 0 Then
    ThreadPseudoHandle = Threads(GetPos(ThreadIDString)).GetThreadClass.ThreadPseudoHandle
End If
End Property

Property Let ThreadPrority(ThreadIDString As String, NewThreadPriority As ThreadPriorityConsts)
If GetPos(ThreadIDString) <> 0 Then
    Threads(GetPos(ThreadIDString)).GetThreadClass.ThreadPriority = NewThreadPriority
End If
End Property

Sub SuspendWin32Thread(ThreadIDString As String)
If GetPos(ThreadIDString) <> 0 Then
    Threads(GetPos(ThreadIDString)).GetThreadClass.SuspendWin32Thread
End If
End Sub

Sub ResumeWin32Thread(ThreadIDString As String)
If GetPos(ThreadIDString) <> 0 Then
    Threads(GetPos(ThreadIDString)).GetThreadClass.ResumeWin32Thread
End If
End Sub

Sub TerminateWin32Thread(ThreadIDString As String, Optional ExitCode As Long = 0)
If GetPos(ThreadIDString) <> 0 Then
    Threads(GetPos(ThreadIDString)).GetThreadClass.TerminateWin32Thread (ExitCode)
End If
End Sub

Property Get ObjectInThreadContext(ThreadIDString As String) As Object
If GetPos(ThreadIDString) <> 0 Then
    Set ObjectInThreadContext = Threads(GetPos(ThreadIDString)).GetThreadClass.ObjectInThreadContext
End If
End Property

Function DoesThreadExist(ThreadIDString As String) As Boolean
If GetPos(ThreadIDString) <> 0 Then
    DoesThreadExist = True
End If
End Function

Friend Function GetPos(ThreadIDString As String) As Long
Dim I As Long
    For I = 1 To TCount
        If Threads(I).GetIDString() = ThreadIDString Then
            GetPos = I
            Exit For
        End If
    Next I
End Function

Friend Sub RemThreadInfo(ThreadIDString As String)
Dim Pos As Long
Pos = GetPos(ThreadIDString)
If Pos <> 0 Then
    For I = Pos To TCount - 1
    Set Threads(I) = Threads(I + 1)
    Next I
TCount = TCount - 1
Set Threads(TCount + 1) = Nothing
ReDim Preserve Threads(TCount)
End If
End Sub
