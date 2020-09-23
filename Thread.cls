VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Thread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Event OnThreadCreateSuccess(ByVal ThreadHandle As Long, ByVal ThreadID As Long)
Public Event OnThreadCreateFailure()
Public Event OnThreadStart(ByVal ThreadHandle As Long, ByVal ThreadID As Long)
Public Event OnThreadFinish(ByVal ThreadHandle As Long, ByVal ThreadID As Long)
Public Event OnThreadTerminate(ByVal ThreadHandle As Long, ByVal ThreadID As Long, ByVal ExitCode As Long)
Public Event OnThreadSuspend(ByVal ThreadHandle As Long, ByVal ThreadID As Long)
Public Event OnThreadResume(ByVal ThreadHandle As Long, ByVal ThreadID As Long)
Public Event OnThreadPriorityChange(ByVal ThreadHandle As Long, ThreadID As Long, ByVal OldPriority As ThreadPriorityConsts, ByVal NewPriority As ThreadPriorityConsts)
Enum ThreadPriorityConsts
    THREAD_PRIORITY_IDLE = -15
    THREAD_PRIORITY_LOWEST = -2
    THREAD_PRIORITY_BELOW_NORMAL = -1
    THREAD_PRIORITY_NORMAL = 0
    THREAD_PRIORITY_ABOVE_NORMAL = 1
    THREAD_PRIORITY_HIGHEST = 2
    THREAD_PRIORITY_TIME_CRITICAL = 15
End Enum
Dim TH As Long, TID As Long, CObject As Object
Dim TDat As TData, IID_IUnknown As VBGUID
Dim MTFunc As String
'TH HOLDS THE Thread handle and TID holds the Thread ID
Function CreateWin32Thread(ThreadObject As Object, FunctionName As String, Optional FuncParam As Variant) As Long
'Initialize the IUnknown OLE interface ID structure
    With IID_IUnknown
            .Data4(0) = &HC0
            .Data4(7) = &H46
    End With

'Set the parameters required to be passed to the Multithreaded
'procedure TProc
    TDat.CLASSID = CLSIDFromProgID("MThreadVB.ILaunch")
    'Class ID of the dummy object
    TDat.FuncParam = FuncParam
    'Funcation Parameter
    Set TDat.ClientObject = ThreadObject
    'The client Object that contains the function to be
    'multithreaded
    Set TDat.ThreadClass = Me
    'ThreadClass references this class
    TDat.cStream = CoMarshalInterThreadInterfaceInStream(IID_IUnknown, ThreadObject)
    TH = CreateThread(0, 0, AddressOf TProc, VarPtr(TDat), 0, TID)
    'Call CreateThread to Create a raw Win32 thread
    
    'Raise appropriate events depending on the whether
    'the API was able to create the thread or not
    Set CObject = Nothing
    If TH = 0 Then
        RaiseEvent OnThreadCreateFailure
    Else
        RaiseEvent OnThreadCreateSuccess(TH, TID)
    End If
    
    MTFunc = FunctionName
    CreateWin32Thread = TH
End Function
Function TerminateWin32Thread(Optional ExitCode As Long = 0) As Boolean
    TerminateWin32Thread = CBool(TerminateThread(TH, ExitCode))
    If TerminateWin32Thread = True Then
        Set CObject = Nothing
        Set TDat.ClientObject = Nothing
        RaiseEvent OnThreadTerminate(TH, TID, ExitCode)
        TID = 0
        TH = 0
    End If
End Function
Function SuspendWin32Thread() As Boolean
Dim Res As Long
    Res = SuspendThread(TH)
    If Res <> &HFFFFFFFF Then
        RaiseEvent OnThreadSuspend(TH, TID)
        SuspendWin32Thread = True
    End If
End Function

Property Let ThreadPriority(Priority As ThreadPriorityConsts)
Dim OldP As ThreadPriorityConsts
OldP = GetThreadPriority(TH)
    If CBool(SetThreadPriority(TH, Priority)) = True Then
        RaiseEvent OnThreadPriorityChange(TH, TID, OldP, Priority)
    End If
End Property

Property Get ThreadPriority() As ThreadPriorityConsts
    ThreadPriority = GetThreadPriority(TH)
End Property

Function ResumeWin32Thread() As Boolean
    Dim Res As Long
    Res = ResumeThread(TH)
    If Res <> &HFFFFFFFF Then
        RaiseEvent OnThreadResume(TH, TID)
        ResumeWin32Thread = True
    End If
End Function

Property Get ThreadHandle() As Long
    ThreadHandle = TH
End Property
Property Get ThreadPseudoHandle() As Long
    ThreadPseudoHandle = GetCurrentThread()
End Property


Property Get ThreadID() As Long
    ThreadID = TID
End Property

Property Get IsThreadRunning() As Boolean
    If TID <> 0 Then IsThreadRunning = True
End Property

Friend Sub RaiseStart()
    RaiseEvent OnThreadStart(TH, TID)
End Sub

Friend Sub RaiseEND()
    Set TDat.ClientObject = Nothing
    RaiseEvent OnThreadFinish(TH, TID)
    TH = 0
    TID = 0
   
End Sub

Property Get ObjectInThreadContext() As Object
    Set ObjectInThreadContext = CObject
End Property

Friend Sub SetContextObject(Obj As Object)
    Set CObject = Obj
End Sub

Friend Function GetMTFuncName() As String
    GetMTFuncName = MTFunc
End Function

