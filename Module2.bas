Attribute VB_Name = "Module2veRecv"
Option Explicit
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function ReadProcessMem Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function WriteProcessMem Lib "kernel32" Alias "WriteProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Long, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lPaketing As String, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lPaketing As Any, ByVal lpFileName As String) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Public Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
Public Declare Function SuspendThread Lib "kernel32" (ByVal hThread As Long) As Long
Public Declare Function OpenThread Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function Thread32First Lib "kernel32" (ByVal hObject As Long, p As THREADENTRY32) As Boolean
Public Declare Function Thread32Next Lib "kernel32" (ByVal hObject As Long, p As THREADENTRY32) As Boolean
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function SetActiveWindow Lib "User32.dll" (ByVal hwnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "User32.dll" (ByVal hwnd As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SendMessageSTRING Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'Public Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
'Public Declare Function WriteProcessMem Lib "kernel32" Alias "WriteProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Const MAILSLOT_NO_MESSAGE  As Long = (-1)
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Type POINTAPI
        X As Long
        Y As Long
End Type

Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPHEAPLIST = &H1
Private Const TH32CS_SNAPTHREAD = &H4
Private Const TH32CS_SNAPMODULE = &H8
Private Const GW_HWNDNEXT = 2
Private Const TH32CS_SNAPALL = TH32CS_SNAPPROCESS + TH32CS_SNAPHEAPLIST + TH32CS_SNAPTHREAD + TH32CS_SNAPMODULE
Private Const MAX_PATH As Integer = 260

Public Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szExeFile As String * MAX_PATH
End Type

Public Type THREADENTRY32
    dwSize As Long
    cntUsage As Long
    th32ThreadID As Long
    th32OwnerProcessID As Long
    tpBasePri As Long
    tpDeltaPri As Long
    dwFlags As Long
End Type

Public Type MODULEINFO
lpBaseOfDLL As Long
SizeOfImage As Long
EntryPoint As Long
End Type
Public Type PROCESS_INFORMATION
         hProcess As Long
         hThread As Long
         dwProcessId As Long
         dwThreadId As Long
End Type

Public Type STARTUPINFO
         cb As Long
         lpReserved As String
         lpDesktop As String
         lpTitle As String
         dwX As Long
         dwY As Long
         dwXSize As Long
         dwYSize As Long
         dwXCountChars As Long
         dwYCountChars As Long
         dwFillAttribute As Long
         dwFlags As Long
         wShowWindow As Integer
         cbReserved2 As Integer
         lpReserved2 As Long
         hStdInput As Long
         hStdOutput As Long
         hStdError As Long
End Type

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type



Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type ItemStruct
    id As Long
    Name As String
End Type

Public Type SkillStruct
    id As Long
    Name As String
    Class As String
    Cooldown As Long
End Type

Public Type LootBoxStruct
BoxID As Long
BoxOpened As Boolean
OpenTime As Long
End Type

Public Type NOTIFYICONDATA
cbSize As Long
hwnd As Long
uId As Long
uFlags As Long
uCallBackMessage As Long
hIcon As Long
szTip As String * 64
End Type

Public Type PROCESS
   id As Long
   ExeFile As String
End Type


'Public Const PAGE_EXECUTE_READWRITE = &H40
'Public Const PAGE_EXECUTE = &H10
'Public Const MEM_RESERVE = &H2000
'
'Public Const MEM_RELEASE = &H8000&
'Public Const MEM_DECOMMIT = &H4000


Public Const TH32CS_INHERIT = &H80000000
Private Const WAIT_ABANDONED& = &H80&
Private Const WAIT_ABANDONED_0& = &H80&
Private Const WAIT_FAILED& = -1&
Private Const WAIT_IO_COMPLETION& = &HC0&
Private Const WAIT_OBJECT_0& = 0
Private Const WAIT_OBJECT_1& = 1
Private Const WAIT_TIMEOUT& = &H102&


Public Type SYSTEM_HANDLE_TABLE_ENTRY_INFO
    UniqueProcessId As Integer
    CreatorBackTraceIndex As Integer
    ObjectTypeIndex As Byte
    HandleAttributes As Byte
    HandleValue As Integer
    Object As Long
    GrantedAccess As Long
End Type

Public Type SYSTEM_HANDLE_INFORMATION
    NumberOfHandles As Long
    Handles() As SYSTEM_HANDLE_TABLE_ENTRY_INFO
End Type



Global son As Integer
Global toplam As Integer
Global son2 As Integer
Global toplam2 As Integer
Public Const lngNull = 0
Public LootBox(1 To 20) As LootBoxStruct
Public Items() As ItemStruct
Public Skills() As SkillStruct
Public TimedSkills() As SkillStruct
Public OtherItems() As ItemStruct
Public Const ERROR_ALREADY_EXISTS = 183&

Private Declare Function CreateWaitableTimer Lib "kernel32" _
    Alias "CreateWaitableTimerA" ( _
    ByVal lpSemaphoreAttributes As Long, _
    ByVal bManualReset As Long, _
    ByVal lpName As String) As Long
    
Private Declare Function OpenWaitableTimer Lib "kernel32" _
    Alias "OpenWaitableTimerA" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal lpName As String) As Long
    
Private Declare Function SetWaitableTimer Lib "kernel32" ( _
    ByVal hTimer As Long, _
    lpDueTime As FILETIME, _
    ByVal lPeriod As Long, _
    ByVal pfnCompletionRoutine As Long, _
    ByVal lpArgToCompletionRoutine As Long, _
    ByVal fResume As Long) As Long
    
Private Declare Function CancelWaitableTimer Lib "kernel32" ( _
    ByVal hTimer As Long)
    
Private Declare Function MsgWaitForMultipleObjects Lib "user32" ( _
    ByVal nCount As Long, _
    pHandles As Long, _
    ByVal fWaitAll As Long, _
    ByVal dwMilliseconds As Long, _
    ByVal dwWakeMask As Long) As Long
    
Private Const QS_HOTKEY& = &H80
Private Const QS_KEY& = &H1
Private Const QS_MOUSEBUTTON& = &H4
Private Const QS_MOUSEMOVE& = &H2
Private Const QS_PAINT& = &H20
Private Const QS_POSTMESSAGE& = &H8

Private Const QS_SENDMESSAGE& = &H40
Private Const QS_TIMER& = &H10
Private Const QS_MOUSE& = (QS_MOUSEMOVE _
                            Or QS_MOUSEBUTTON)
Private Const QS_INPUT& = (QS_MOUSE _
                            Or QS_KEY)
Private Const QS_ALLEVENTS& = (QS_INPUT _
                            Or QS_POSTMESSAGE _
                            Or QS_TIMER _
                            Or QS_PAINT _
                            Or QS_HOTKEY)
Private Const QS_ALLINPUT& = (QS_SENDMESSAGE _
                            Or QS_PAINT _
                            Or QS_TIMER _
                            Or QS_POSTMESSAGE _
                            Or QS_MOUSEBUTTON _
                            Or QS_MOUSEMOVE _
                            Or QS_HOTKEY _
                            Or QS_KEY)
Public PartyIDx(7) As String
Public PartyHPx(7) As String
Public PartyMaxHPx(7) As String
Public PartyClassx(7) As String
Public PartyAdix(7) As String
Public PartyMob(7) As String
Public OncekiMaxHPx(7) As String
Public ptMaxHPx As Long
Public BuffSyc As Date
Public KO_WindowHandle As Long

Public Sub Bekle(Milisaniye As Long)
    Dim ft As FILETIME
    Dim lBusy As Long
    Dim lRet As Long
    Dim dblDelay As Double
    Dim dblDelayLow As Double
    Dim dblUnits As Double
    Dim hTimer As Long
    
    hTimer = CreateWaitableTimer(0, True, App.EXEName & "Timer")
    
    If Err.LastDllError = ERROR_ALREADY_EXISTS Then
        ' If the timer already exists, it does not hurt to open it
        ' as long as the person who is trying to open it has the
        ' proper access rights.
    Else
        ft.dwLowDateTime = -1
        ft.dwHighDateTime = -1
        lRet = SetWaitableTimer(hTimer, ft, 0, 0, 0, 0)
    End If
    
    ' Convert the Units to nanoseconds.
    dblUnits = CDbl(&H10000) * CDbl(&H10000)
    dblDelay = CDbl(Milisaniye) * 10000
    'dblDelay = CDbl(lNumberOfSeconds) * 1000 * 10000  idi ama milisaniye cinsinden daha iyi
    
    ' By setting the high/low time to a negative number, it tells
    ' the Wait (in SetWaitableTimer) to use an offset time as
    ' opposed to a hardcoded time. If it were positive, it would
    ' try to convert the value to GMT.
    ft.dwHighDateTime = -CLng(dblDelay / dblUnits) - 1
    dblDelayLow = -dblUnits * (dblDelay / dblUnits - _
        Fix(dblDelay / dblUnits))
    
    If dblDelayLow < CDbl(&H80000000) Then
        ' &H80000000 is MAX_LONG, so you are just making sure
        ' that you don't overflow when you try to stick it into
        ' the FILETIME structure.
        dblDelayLow = dblUnits + dblDelayLow
        ft.dwHighDateTime = ft.dwHighDateTime + 1
    End If
    
    ft.dwLowDateTime = CLng(dblDelayLow)
    lRet = SetWaitableTimer(hTimer, ft, 0, 0, 0, False)
    
    Do
        ' QS_ALLINPUT means that MsgWaitForMultipleObjects will
        ' return every time the thread in which it is running gets
        ' a message. If you wanted to handle messages in here you could,
        ' but by calling Doevents you are letting DefWindowProc
        ' do its normal windows message handling---Like DDE, etc.
        lBusy = MsgWaitForMultipleObjects(1, hTimer, False, _
            INFINITE, QS_ALLINPUT&)
        DoEvents
    Loop Until lBusy = WAIT_OBJECT_0
    
    ' Close the handles when you are done with them.
    CloseHandle hTimer

End Sub
Public Function ByteDizisiYaz(pAddy As Long, pmem() As Byte, pSize As Long)
WriteProcessMem KO_HANDLE, pAddy, pmem(LBound(pmem)), pSize, 0&
End Function

Public Function ReadMessage(Handle As Long, MailMessage As String, MessagesLeft As Long)
Dim lBytesRead As Long, lNextMsgSize As Long, lpBuffer As String
ReadMessage = False
Call GetMailslotInfo(Handle, ByVal 0&, lNextMsgSize, MessagesLeft, ByVal 0&)
If MessagesLeft > 0 And lNextMsgSize <> MAILSLOT_NO_MESSAGE Then
    lBytesRead = 0
    lpBuffer = String$(lNextMsgSize, Chr$(0))
    Call ReadFile(Handle, ByVal lpBuffer, Len(lpBuffer), lBytesRead, ByVal 0&)
    If lBytesRead <> 0 Then
        MailMessage = Left(lpBuffer, lBytesRead)
        ReadMessage = True
        Call GetMailslotInfo(Handle, ByVal 0&, lNextMsgSize, MessagesLeft, ByVal 0&)
    End If
End If
End Function
Public Function CheckForMessages(Handle As Long, MessageCount As Long)
Dim lMsgCount As Long, lNextMsgSize As Long
CheckForMessages = False
GetMailslotInfo Handle, ByVal 0&, lNextMsgSize, lMsgCount, ByVal 0&
MessageCount = lMsgCount
CheckForMessages = True
End Function
Public Function EstablishMailSlot(ByVal MailSlotName As String, Optional MaxMessageSize As Long = 0, Optional ReadTimeOut As Long = 50) As Long
EstablishMailSlot = CreateMailslot(MailSlotName, MaxMessageSize, ReadTimeOut, ByVal 0&)
End Function
Function writeMailSlot(MailSlotName As String) As Long
Dim KO_MSLOT As Long, pHook As String, p() As Byte, ph() As Byte, CF As Long, WF As Long, CH As Long
KO_MSLOT = VirtualAllocEx(KO_HANDLE, 0, 1024, MEM_COMMIT, PAGE_READWRITE)
If KO_MSLOT <= 0 Then Exit Function: MsgBox "memory could not be opened!", vbCritical
CF = GetProcAddress(GetModuleHandle("kernel32.dll"), "CreateFileA")
WF = GetProcAddress(GetModuleHandle("kernel32.dll"), "WriteFile")
CH = GetProcAddress(GetModuleHandle("kernel32.dll"), "CloseHandle")
Debug.Print Hex(KO_MSLOT)
Hex2Byte StringToHex(MailSlotName), p
ByteDizisiYaz KO_MSLOT + &H400, p, UBound(p) - LBound(p) + 1
pHook = "558BEC83C4F433C08945FC33D28955F86A0068800000006A036A006A01680000004068" & AlignDWORD(KO_MSLOT + &H400) & "E8" & AlignDWORD(getCallDiff(KO_MSLOT + &H27, CF)) & "8945F86A008D4DFC51FF750CFF7508FF75F8E8" & AlignDWORD(getCallDiff(KO_MSLOT + &H3E, WF)) & "8945F4FF75F8E8" & AlignDWORD(getCallDiff(KO_MSLOT + &H49, CH)) & "8BE55DC3" '&H49
Hex2Byte pHook, ph
ByteDizisiYaz KO_MSLOT, ph, UBound(ph) - LBound(ph) + 1
writeMailSlot = KO_MSLOT
End Function
Sub recvHook(MailSlotName As String, RecvFunction As Long, RecvBase As Long)
Dim KO_MSLOT As Long, KO_RCVHK As Long, pHook As String, ph() As Byte
KO_MSLOT = writeMailSlot(MailSlotName)
If KO_MSLOT <= 0 Then Exit Sub: MsgBox "memory could not be opened!", vbCritical
KO_RCVHK = VirtualAllocEx(KO_HANDLE, 0, 1024, MEM_COMMIT, PAGE_READWRITE)
If KO_RCVHK <= 0 Then Exit Sub: MsgBox "memory could not be opened!", vbCritical

pHook = "558BEC83C4F8538B450883C0048B108955FC8B4D0883C1088B018945F8FF75FCFF75F8E8" & AlignDWORD(getCallDiff(KO_RCVHK + &H23, KO_MSLOT)) & "83C4088B0D" & AlignDWORD(KO_PTR_DLG - &H14) & "FF750CFF7508B8" & AlignDWORD(RecvFunction) & "FFD05B59595DC20800"
Hex2Byte pHook, ph
ByteDizisiYaz KO_RCVHK, ph, UBound(ph) - LBound(ph) + 1

pHook = AlignDWORD(KO_RCVHK)
Hex2Byte pHook, ph
ByteDizisiYaz RecvBase, ph, UBound(ph) - LBound(ph) + 1
End Sub

Public Sub FindHook(MailSlotName As String)
Dim KO_RECVHK As Long, KO_RCVHKB As Long
KO_RECVHK = ReadLong(ReadLong(KO_PTR_DLG - &H14)) + &H8 ' &HB844A8 'ReadLong(ReadLong(KO_PTR_DLG - &H14)) + &H8 '&HB844A8 'ReadLong(ReadLong(KO_PTR_DLG - &H14)) + &H8
KO_RCVHKB = ReadLong(KO_RECVHK) '&H54B120 'ReadLong(KO_RECVHK) '&H54B120 '

' KO_RECV_PTR = &HBC4270
' KO_RECV_FNC = &H6FE950
Debug.Print Hex(KO_RECVHK) & "//" & Hex(KO_RCVHKB)
recvHook MailSlotName, KO_RCVHKB, KO_RECVHK
End Sub


Sub DispatchMailSlot(Handle As Long)
Dim MsgCount As Long, rc As Long, MessageBuffer As String, code, PacketType As String
Dim BoxID2, BoxID, ItemID1, ItemID2, ItemID3, ItemID4, ItemID5, ItemID5C, ItemID5L, RecAl1, RecAl2, RecAl4, RecAl3 As Long
MsgCount = 1
Do While MsgCount <> 0
rc = CheckForMessages(Handle, MsgCount)
If CBool(rc) And MsgCount > 0 Then
    If ReadMessage(Handle, MessageBuffer, MsgCount) Then
    code = MessageBuffer
    On Error Resume Next
    
    Select Case Asc(Left(MessageBuffer, 1))
        Case Else
         If mID(StringToHex(MessageBuffer), 1, 2) = "68" Then
           BoxID2 = mID(StringToHex(MessageBuffer), 7, 8)
           Form1.Text6.Text = BoxID2
        End If
           If mID(StringToHex(MessageBuffer), 1, 2) = "68" Then
                BoxID = mID(StringToHex(MessageBuffer), 3, 4)
                ItemID1 = mID(StringToHex(MessageBuffer), 13, 8)
                ItemID2 = mID(StringToHex(MessageBuffer), 25, 8)
                ItemID3 = mID(StringToHex(MessageBuffer), 37, 8)
                ItemID4 = mID(StringToHex(MessageBuffer), 49, 8)
                ItemID5 = mID(StringToHex(MessageBuffer), 109, 8)
                ItemID5C = AlignDWORD("&H" + ItemID5)
                'ItemID5C = AlignDWORD(ItemID5C)
                
                RecAl1 = mID(StringToHex(MessageBuffer), 7, 4)
                RecAl2 = mID(StringToHex(MessageBuffer), 21, 4)
                RecAl3 = mID(StringToHex(MessageBuffer), 33, 4)
                RecAl4 = mID(StringToHex(MessageBuffer), 45, 4)
                
                Form1.Text7.Text = BoxID
                Form1.Text8.Text = ItemID5C
                Form1.Text9.Text = ItemID2
                Form1.Text10.Text = ItemID3
                Form1.Text11.Text = RecAl1
                'If ItemID2 > 0 Then: Paket "26" & BoxID & RecAl1 & ItemID2 & "01" & "00"
                'Bekle 200
                'If ItemID3 > 0 Then: Paket "26" & BoxID & RecAl1 & ItemID3 & "02" & "00"
                'Bekle 200
                'If ItemID4 > 0 Then: Paket "26" & BoxID & RecAl1 & ItemID4 & "03" & "00"
                'Bekle 200
                'If ItemID1 > 0 Then: Paket "26" & BoxID & RecAl1 & ItemID1 & "00" & "00"
          End If
        
        
        
        
    End Select

    End If
End If
Loop
End Sub
Sub DispatchMailSlotkutu(Handle As Long)
Dim MsgCount As Long, rc As Long, MessageBuffer As String, code, PacketType As String
Dim BoxID2, BoxID, ItemID1, ItemID2, ItemID3, ItemID4, RecAl1, RecAl2, RecAl4, RecAl3 As Long
MsgCount = 1
Do While MsgCount <> 0
rc = CheckForMessages(Handle, MsgCount)
If CBool(rc) And MsgCount > 0 Then
    If ReadMessage(Handle, MessageBuffer, MsgCount) Then
    code = MessageBuffer
    On Error Resume Next
    
    Select Case Asc(Left(MessageBuffer, 1))
        Case Else
         If mID(StringToHex(MessageBuffer), 1, 2) = "23" Then
           BoxID2 = mID(StringToHex(MessageBuffer), 7, 8)
           Paket "24" & BoxID2
        End If
           If mID(StringToHex(MessageBuffer), 1, 2) = "24" Then
                BoxID = mID(StringToHex(MessageBuffer), 3, 4)
                ItemID1 = mID(StringToHex(MessageBuffer), 13, 8)
                ItemID2 = mID(StringToHex(MessageBuffer), 25, 8)
                ItemID3 = mID(StringToHex(MessageBuffer), 37, 8)
                ItemID4 = mID(StringToHex(MessageBuffer), 49, 8)
                RecAl1 = mID(StringToHex(MessageBuffer), 7, 4)
                RecAl2 = mID(StringToHex(MessageBuffer), 21, 4)
                RecAl3 = mID(StringToHex(MessageBuffer), 33, 4)
                RecAl4 = mID(StringToHex(MessageBuffer), 45, 4)
                If ItemID2 > 0 Then: Paket "26" & BoxID & RecAl1 & ItemID2 & "01" & "00"
                Bekle 200
                If ItemID3 > 0 Then: Paket "26" & BoxID & RecAl1 & ItemID3 & "02" & "00"
                Bekle 200
                If ItemID4 > 0 Then: Paket "26" & BoxID & RecAl1 & ItemID4 & "03" & "00"
                Bekle 200
                If ItemID1 > 0 Then: Paket "26" & BoxID & RecAl1 & ItemID1 & "00" & "00"
          End If
        
        
        
        
    End Select

    End If
End If
Loop
End Sub
Sub DispatchMailSlotokuuu(Handle As Long)
On Error Resume Next
Dim MsgCount As Long
Dim rc As Long
Dim MessageBuffer As String
Dim fullcode
Dim code
MsgCount = 1
Do While MsgCount <> 0
rc = CheckForMessages(Handle, MsgCount)

If CBool(rc) And MsgCount > 0 Then
If ReadMessage(Handle, MessageBuffer, MsgCount) Then
code = MessageBuffer
fullcode = Strings.Split(MessageBuffer, "")
On Error Resume Next

If Revcfrm.Check2.value = 1 Then
Revcfrm.Text1.SelStart = Len(Revcfrm.Text1.Text)
Select Case Asc(Left(MessageBuffer, 1))

Case &H68
Revcfrm.Text1.SelText = "RECV-->MERCHANT-->" & StringToHex(MessageBuffer) & vbCrLf 'If Revcfrm.List1.Selected(0) = True Then
Case &H69
Revcfrm.Text1.SelText = "RECV-->MERCHANT_INOUT-->" & StringToHex(MessageBuffer) & vbCrLf 'If Revcfrm.List1.Selected(1) = True Then
'Case Else  'UNKNOW
'Revcfrm.Text1.SelText = "RECV-->UNKNOW (" & Left(StringToHex(MessageBuffer), 2) & ")-->" & StringToHex(MessageBuffer) & vbCrLf
End Select
End If
End If
End If
Loop
End Sub
