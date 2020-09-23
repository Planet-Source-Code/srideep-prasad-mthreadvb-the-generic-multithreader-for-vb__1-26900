Attribute VB_Name = "Module1"
'//////////////////////////////////////////////////////////
'SOME IMPORTANT NOTES !
'//////////////////////////////////////////////////////////
'Many of you may be wondering how the CreateThread API works
'Normally it doesn't as far as VB 6 is concerned..
'In fact, I too, after trial and error had concluded that the
'only function that CreateThread works with is a blank one
'or one that only calls the Beep API !

'But a sometime back, things changed ! I came across a demo
'program created BY AN EXTREMENLY TALENTED AND INNOVATIVE
'PROGRAMMER CALLED MATT CURLAND (Unfortunately he is not
'a member of PSC !)

'And in his demo he showed
'the various OLE APIs that need to be called to make
'CreateThread safe... His example was very difficuilt to
'understand owing to the several class modules he used in
'the Demo and the zigzag nature of execution... After
'LOTS OF TRIALS AND ERRORS AND QUITE A FEW FREEZES AND REBOOTS
'I Managed to make the API somewhat safe and have now have
'Created a reasonably safe generic multithreader for VB.
'What you see now as I said before is the result of heavy
'experimentation and LOTS of reboots !
'//////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////
'TECHNICAL STUFF FOR TECH BUFFS AT PSC !!!
'//////////////////////////////////////////////////////////
'How does the multithreader work ?

'Normally, all VB programs are heavily dependent on the
'runtime DLL for its functioning
'In VB 6, within the multithreaded function, any calls
'to the runtime DLL fails(due to some reason) causing your
'program to crash immediately....
'Even an API call is not really compiled in "real" native
'code in VB and is interpreted by the runtime DLL...
'This ultimately involves calling the runtime DLL,
'that causes VB to crash.Most standard VB statements and
'functions such as Set = , For...Next etc also call the runtime
'DLL and ultimately even these fail... (So much for native
'code compilation !)

'If an object could be created within the multithreaded
'procedure, the VB runtime starts behaving properly...
'The trouble is,the standard VB instantiator functions fail
'within the multithreaded procedures...
'Therefore, I have used the ThreadAPI.Tlb type library to
'bypass the Runtime and directly call the OLE/COM o
'APIs (This can be verified using the Dependency Viewer)
'and create a dummy object inside the multithreaded
'prodecure. After this has been done, the runtime DLL starts
'behaving properly and it is possible to call all VB functions
'safely...
'Once again I THANK MATT CURLAND, A VERY TALENTED PROGRAMMER
'FOR GIVING INFO ON USING THE OLE APIs IN VB.
'//////////////////////////////////////////////////////////

'THE THREADAPI.TLB TYPE LIBRARY USED FOR CALLING THE APIs
'SAFELY WITHIN THE MULTITHREADED PROCEDURE (TProc) Can be
'DOWNLOADED SEPERATELY FROM VBACCELERATOR.COM TOO...
'IT IS ALSO INCLUDED WITH THIS PACKAGE

Public Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Public Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Public Declare Function SuspendThread Lib "kernel32" (ByVal hThread As Long) As Long
Public Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
Public Declare Function GetThreadPriority Lib "kernel32" (ByVal hThread As Long) As Long
Public Declare Function GetCurrentThread Lib "kernel32" () As Long


Public Type TData
    CLASSID As CLSID
    cStream As Long
    ThreadClass As Thread
    FuncParam As Variant
    ClientObject As Object
End Type

Function TProc(NewThreadInfo As TData) As Long
Dim hr As Long
Dim pUnk As IUnknown
Dim IID_IUnknown As VBGUID, Obj As Object
Dim FuncName As String
'Initialize the OLE/COM subsystem
Call CoInitialize(0)

        'Initialize the IUnknown interface ID structure
        With IID_IUnknown
            .Data4(0) = &HC0
            .Data4(7) = &H46
        End With
'Create a dummy object referencing the ILaunch class
Call CoCreateInstance(NewThreadInfo.CLASSID, Nothing, CLSCTX_INPROC_SERVER, IID_IUnknown, pUnk)
'Raise the OnThreadStart event
Set Obj = CoGetInterfaceAndReleaseStream(NewThreadInfo.cStream, IID_IUnknown)
NewThreadInfo.ThreadClass.SetContextObject Obj
FuncName = NewThreadInfo.ThreadClass.GetMTFuncName
'Call the function which is to be multithreaded
NewThreadInfo.ThreadClass.RaiseStart
CallByName NewThreadInfo.ClientObject, FuncName, VbMethod, NewThreadInfo.FuncParam




'Raise the OnThreadFinish event
NewThreadInfo.ThreadClass.RaiseEND
NewThreadInfo.ThreadClass.SetContextObject Nothing
Set pUnk = Nothing
Set Obj = Nothing
'Uninitialize the OLE/COM subsystem
Call CoUninitialize

End Function



