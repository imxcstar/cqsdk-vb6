Attribute VB_Name = "DllEntry"
Option Explicit
Private Const DLL_PROCESS_ATTACH  As Long = 1
Private Const DLL_THREAD_ATTACH   As Long = 2
Private Const DLL_PROCESS_DETACH  As Long = 0
Private Const DLL_THREAD_DETACH   As Long = 3
Private Type UUID
    data1       As Long
    data2       As Integer
    data3       As Integer
    data4(7)    As Byte
End Type
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function VBGetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" (ByVal lpModName As Long) As Long
Private Declare Function CoInitializeEx Lib "ole32.dll" (ByVal pvReserved As Long, ByVal dwCoInit As Long) As Long
Private Declare Function VBDllGetClassObject Lib "msvbvm60.dll" (gloaders As Long, gvb As Long, ByVal gvbtab As Long, rclsid As UUID, riid As UUID, ppv As Any) As Long
Private Declare Sub UserDllMain Lib "msvbvm60.dll" (u1 As Long, u2 As Long, ByVal u3_h As Long, ByVal u4_1 As Long, ByVal u5_0 As Long)
Private Declare Function CreateThread Lib "kernel32" (ByVal lpSecurityAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As Long) As Long
Private Declare Function CoInitialize Lib "ole32.dll" (ByVal pvReserved As Long) As Long
Private Declare Sub CoUninitialize Lib "ole32.dll" ()
Private Declare Function CreateIExprSrvObj Lib "msvbvm60.dll" (ByVal p1_0 As Long, ByVal p2_4 As Long, ByVal p3_0 As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Sub FreeLibraryAndExitThread Lib "kernel32" (ByVal hLibModule As Long, ByVal dwExitCode As Long)
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Private Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Long, lpThreadAttributes As Any, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Private Const CREATE_DEFAULT = &H0&
Private Const CREATE_SUSPENDED = &H4&
Private Const STACK_SIZE_CUSTOM = &H10000
Private Const COINIT_MULTITHREADED = &H0&
Private Const COINIT_SPEED_OVER_MEMORY = &H8&
Private hh As Long, VBHeader As Long

Public Sub Main()
    '-------------
End Sub

Public Function DllMain(ByVal hinstDLL As Long, ByVal fdwReason As Long, ByVal lpvReserved As Long) As Long
    Dim nRet As Long
    nRet = 1
    CreateIExprSrvObj 0, 4, 0
    hh = hinstDLL
    init ByVal hinstDLL
    InitCommonControls
    Select Case fdwReason
    Case DLL_PROCESS_ATTACH
        
    Case DLL_THREAD_ATTACH

    Case DLL_PROCESS_DETACH

    Case DLL_THREAD_DETACH

    End Select
    DllMain = nRet
End Function

Public Function init(ByVal hinstDLL As Long)
    Dim fake        As Long
    Dim lp As Long, lvb As Long
    Dim riid        As UUID
    Dim aiid        As UUID
    Dim ofac        As Object
    Dim f0          As Long
    Dim fakehead    As Long
    Dim ll          As Long
    hh = hinstDLL
    CreateIExprSrvObj 0, 4, 0
    Call CoInitialize(0)
    With riid
        .data1 = 1
        .data4(0) = &HC0
        .data4(7) = &H46
    End With
    If hinstDLL = 0 Then hinstDLL = GetModuleHandle(0)
    f0 = GetFakeH(GetModuleHandle(0))
    fakehead = GetFakeH(hinstDLL)
    If f0 = 0 Then
        Call VBDllGetClassObject(GetModuleHandle(0), lvb, ByVal fakehead, aiid, riid, ofac)
    Else
        Call VBDllGetClassObject(hinstDLL, lvb, ByVal fakehead, aiid, riid, ofac)
    End If
    hh = hinstDLL
End Function

Public Function GetFakeH(ByVal hin As Long) As Long
    Dim lPtr          As Long
    Dim lRet          As Long
    Dim isvb          As String
    Dim ll            As Long
    Dim mdat(1033)    As Byte
    GetFakeH = 0
    lPtr = hin
    isvb = StrConv("VB5!", vbFromUnicode)
    Do
        If ReadProcessMemory(-1, ByVal lPtr, mdat(0), 1034, ll) = 0 Then Exit Function
        lRet = InStrB(mdat, isvb)
        If lRet <> 0 Then Exit Do
        lPtr = lPtr + 1024
    Loop
    GetFakeH = lPtr + lRet - 1
End Function

Public Function InitVBdll() As Boolean
    Dim pIID As UUID, pDummy As UUID
    Dim u1 As Long, u2 As Long, u3 As Long
    If VBHeader > 0 Then
        pIID.data1 = &H1&
        pIID.data4(0) = &HC0
        pIID.data4(7) = &H46
        u3 = VBGetModuleHandle(ByVal 0&)
        UserDllMain u1, u2, u3, 1&, 0&
        VBDllGetClassObject u1, u2, VBHeader, pDummy, pIID, pDummy
        InitVBdll = True
    Else
        InitVBdll = False
    End If
End Function

Public Function RunDllHostCallBack(a As Long, b As Long, c As Long) As Long
    RunDllHostCallBack = 0
End Function
