Attribute VB_Name = "Module2"
Public Type TDataEx
    CLASSID As CLSID
    cStream As Long
    ThreadClass As ThreadEX
'   Funcname As String
    FuncParam As Variant
    ClientObject As Object
    ThreadIndex As Long
    EventHandle As Long
End Type

Function TProcEx(NewThreadInfo As TDataEx) As Long
Dim hr As Long
Dim pUnk As IUnknown
Dim IID_IUnknown As VBGUID, Obj As Object
Dim FuncName As String
'Initialize the OLE/COM subsystem
Call CoInitialize(0)
Call WaitForSingleObject(NewThreadInfo.EventHandle, INFINITE)
CloseHandle (NewThreadInfo.EventHandle)
        'Initialize the IUnknown interface ID structure
        With IID_IUnknown
            .Data4(0) = &HC0
            .Data4(7) = &H46
        End With
'Create a dummy object referencing the ILaunch class
Call CoCreateInstance(NewThreadInfo.CLASSID, Nothing, CLSCTX_INPROC_SERVER, IID_IUnknown, pUnk)
'Raise the OnThreadStart event
Call WaitForSingleObject(NewThreadInfo.EventHandle, INFINITE)
CloseHandle (NewThreadInfo.EventHandle)
Set Obj = CoGetInterfaceAndReleaseStream(NewThreadInfo.cStream, IID_IUnknown)
NewThreadInfo.ThreadClass.RaiseStart NewThreadInfo.ThreadIndex
'Call the function which is to be multithreaded

FuncName = NewThreadInfo.ThreadClass.GetFunctionName(NewThreadInfo.ThreadIndex)
NewThreadInfo.ThreadClass.SetContextObject NewThreadInfo.ThreadIndex, Obj

CallByName NewThreadInfo.ClientObject, FuncName, VbMethod, NewThreadInfo.FuncParam

'Raise the OnThreadFinish event
NewThreadInfo.ThreadClass.RaiseEND NewThreadInfo.ThreadIndex
Call NewThreadInfo.ThreadClass.SetContextObject(NewThreadInfo.ThreadIndex, Nothing)
Set pUnk = Nothing
Set Obj = Nothing
'Uninitialize the OLE/COM subsystem
Call CoUninitialize
End Function




