Attribute VB_Name = "Module1"
Public Enum enSW
    SW_HIDE = 0
    SW_NORMAL = 1
    SW_MAXIMIZE = 3
    SW_MINIMIZE = 6
End Enum

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
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Public Enum enPriority_Class
    NORMAL_PRIORITY_CLASS = &H20
    IDLE_PRIORITY_CLASS = &H40
    HIGH_PRIORITY_CLASS = &H80
End Enum
Public Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Long, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function ReadProcessMem Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function WriteProcessMem Lib "kernel32" Alias "WriteProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Public Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function SetActiveWindow Lib "User32.dll" (ByVal Hwnd As Long) As Long
Public Declare Function GetMailslotInfo Lib "kernel32" (ByVal hMailSlot As Long, lpMaxMessageSize As Long, lpNextSize As Long, lpMessageCount As Long, lpReadTimeout As Long) As Long
Public Declare Function SetForegroundWindow Lib "User32.dll" (ByVal Hwnd As Long) As Long
Public Declare Function CreateMailslot Lib "kernel32" Alias "CreateMailslotA" (ByVal lpName As String, ByVal nMaxMessageSize As Long, ByVal lReadTimeout As Long, lpSecurityAttributes As Any) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal Hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const MEM_RESERVE As Long = &H2000&
Public Const STARTF_USESHOWWINDOW = &H1
Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Const MEM_COMMIT = &H1000
Public Const MEM_RELEASE = &H8000&
Public Const PAGE_READWRITE = &H4&
Public Const INFINITE = &HFFFF

Dim KutuAç, KutununÝçi, kutuiçi1oku, kutuiçi2oku, kutuiçi3oku, kutuiçi4oku, kutuiçi5oku, kutuiçi6oku, kutuid, BoxID As Long
Dim kutu1id As Long
Dim kutu2id As Long
Dim kutu3id As Long

Public BytesAddr As Long
Public FuncPtr As Long
Public ByteMob_Base As Long
Public RecvHandle As Long
Public KO_HANDLE As Long
Public KO_WindowHandle As Long
Public KO_ADR_CHR As Long
Public KO_ADR_DLG As Long
Public KO_PID As Long
Public packetbytes As Long
Public codebytes As Long
Public zMobName As String, zMobZ As Long, zMobID As String
Public ItemLevel As Long
Public BankadakiItemler(191) As String
Public ItemIntID(41) As String
Public Süre As Long
Public id(14) As String
' Pointerler
Public KO_FNC_ISEN As Long
Public KO_PTR_CHR As Long
Public KO_PTR_PKT As Long
Public KO_PTR_DLG As Long
Public KO_NODC As Long
Public KO_SND_FNC As Long
Public KO_SEND_PTR As Long
Public KO_SND_PACKET As Long
Public KO_SH_VALUE As Long
'Clientten Seçmeler


Public KO_PERI As Long

' Offsetler
Public KO_OFF_SWIFT As Long
Public KO_OFF_CLASS As Long
Public KO_OFF_ID As Long
Public KO_OFF_MOB As Long
Public KO_OFF_HP As Long
Public KO_OFF_MAXHP As Long
Public KO_OFF_MP As Long
Public KO_OFF_MAXMP As Long
'Clientten Seçmeler
Public KO_FLDB As Long
Public KO_FMBS As Long
Public KO_FLPZ As Long
Public KO_FPBS As Long
Public KO_FNCB As Long
Public KO_OFF_NAME As Long
Public KO_OFF_NAMELONG As Long


Public KO_OFF_MOBMAX As Long



Public KO_OFF_Y As Long
Public KO_OFF_X As Long
Public KO_OFF_Z As Long
Public KO_OFF_MX As Long
Public KO_OFF_MY As Long
Public KO_OFF_MZ As Long
Public KO_OFF_Go1 As Long
Public KO_OFF_GoX As Long
Public KO_OFF_GoY As Long
Public KO_OFF_Go2 As Long
Public KO_OFF_ZONE As Long
Public KO_OFF_NATION As Long
Public KO_OFF_WH As Long
Public KO_OFF_CHAT As Long
Public KO_RECV_FNC As Long
Public KO_RECV_PTR As Long
Public KO_RECVHK As Long
Public KO_RCVHKB As Long
Public KO_RCVHKB1 As Long
Public KO_RCVHKB2 As Long
Public KO_RCVHKB3 As Long
Public KO_RECVHK1 As Long
Public KO_RECVHK2 As Long
Public KO_RECVHK3 As Long
Public Const KO_OTO_LOGIN_PTR As Long = &HDD1FF0
Public Const KO_OTO_LOGIN_ADR1 As Long = &H4D7480
Public Const KO_OTO_LOGIN_ADR2 As Long = &H4D0950
Public Const KO_OTO_LOGIN_ADR3 As Long = &H4D0410
Public Const KO_OTO_LOGIN_ADR4 As Long = &H4D3700

Public Const KO_BYPASS_ADR1 As Long = &H9A0F91
Public Const KO_BYPASS_ADR2 As Long = &H4BAF8C
Public Const KO_BYPASS_ADR3 As Long = &H4BAFA2
Public Const KO_BYPASS_ADR4 As Long = &H4BAFB1

Public Function AttachKO() As Boolean
    GetWindowThreadProcessId FindWindow(vbNullString, Form1.Text1.Text), KO_PID
    KO_HANDLE = OpenProcess(PROCESS_ALL_ACCESS, False, KO_PID)
        If KO_HANDLE = 0 Then
            AttachKO = False
            Exit Function
        End If
    If KO_PID = 0 Then End
    AttachKO = True
    Dim RecvMailSlot As String
    RecvMailSlot = "\\.\mailslot\Cro" & Hex(GetTickCount)
    RecvHandle = EstablishMailSlot(RecvMailSlot)
    FindHook RecvMailSlot

End Function
Function OffsetleriYükleTWKO()
    ' Pointerler
 KO_PTR_CHR = &HE83C48
 KO_PTR_DLG = &HE69568
 KO_PTR_PKT = &HE69534
 KO_SND_FNC = &H48C610
 KO_RECV_PTR = &HBC4270
 KO_RECV_FNC = &H6FE950
 KO_SND_PACKET = KO_PTR_PKT + &HC5
    KO_OFF_SWIFT = &H7BC
    KO_PERI = &H570230
    KO_OFF_CHAT = &HE1B490
    KO_OFF_NAME = &H688
    KO_OFF_ID = &H680
    KO_OFF_MOB = &H644
    KO_OFF_WH = &H6C0
    KO_OFF_MAXHP = &H6B8
    KO_OFF_HP = &H6BC
    KO_OFF_MAXMP = &HB5C
    KO_OFF_MP = &HB60
    KO_OFF_EXP = &HB78
    KO_OFF_MAXEXP = &HB70
    KO_OFF_NATION = &H6A8
    KO_OFF_CLASS = &H6B0
    KO_OFF_LVL = &H6B4
    KO_OFF_ZONE = &HC00
    KO_OFF_X = &HD8
    KO_OFF_Y = &HE0
    KO_OFF_Z = &HDC
    KO_OFF_Go1 = &HF90
    KO_OFF_GoX = &HF9C
    KO_OFF_GoY = &HFA4
    KO_OFF_Go2 = &H3F0
    KO_OFF_MX = KO_OFF_GoX
    KO_OFF_MY = KO_OFF_GoY
    KO_OFF_MZ = &HFA0
End Function
Function OffsetleriYükle()
KO_PTR_CHR = &HE3DBC8
KO_PTR_DLG = &HE240D4
KO_PTR_PKT = &HE240A0
KO_KEY_PTR = &HE2409C
KO_SND_FNC = &H497C50
KO_FNC_ISEN = &H54EC50
KO_CHAR_SERV = &HC718B4
KO_NODC = &HB75778
KO_MULTI = &HB93ABC
KO_CHAT = &HE26150

KO_OTO_BTN_PTR = &HE240D0
KO_BTN_LEFT = &H4CB200
KO_BTN_RIGHT = &H4CB4A0
KO_BTN_LOGIN = &H4C7A30

KO_FLDB = &HE3DBC4
KO_FNCZ = &H522450
KO_FNCB = &H5225C0
KO_FMBS = &H4F8840
KO_FPBS = &H4F97B0
KO_FNCX = &H522510
KO_FPOZ = &H562030
KO_FPOB = &H5619F0
KO_STMB = &H51B8C0
KO_ITOB = &HE3D9FC
KO_ITEB = &HE3DA04
KO_FPOX = &H69C3B0
'KO_FLPZ = &HE4D184
KO_FLPZ = ReadLong(ReadLong(KO_FLDB + &H40) + 4) + 4
'KO_FLPZ = ReadLong(ReadLong(ReadLong(ReadLong(KO_FLDB) + 40)))
KO_PM = &HC7F0C4
KO_PERI = &H577A90
KO_PERI_CLOOT_STEAM = &HCCEE60
KO_PERI_MLOOT_STEAM = &HCCEE64
KO_PERI_CLOOT_STEAM = &H0
KO_PERI_MLOOT_STEAM = &H0
KO_RECV_FNC = &H54B190
KO_RECV_PTR = &HB766E8
''//== XignCode Pointer ===
KO_CRE_THREAD = &H9CBE41
KO_LANC_BYPASS = &H9D0995
KO_XIGN_BYPASS = &H9CF75D
KO_XIGN_EXBPS1 = &H4B790C
KO_XIGN_EXBPS2 = &H4B7922
KO_XIGN_EXBPS3 = &H4B7931
KO_XIGN_PTR = &HE337A8
KO_XIGN_FNC = &H45B250
KO_XIGN_PRS = &H4B5200
''//==== Offsetler ====
KO_OFF_CLASS = &H6B0
KO_OFF_NT = &H6A8
KO_OFF_MOVE = &HFB8
KO_OFF_TYPE = &H3F0
KO_OFF_MX = &HFC4
KO_OFF_MZ = &HFC8
KO_OFF_MY = &HFCC
KO_OFF_X = &HD8
KO_OFF_Z = &HDC
KO_OFF_Y = &HE0
KO_OFF_MOBX = &H7C
KO_OFF_MOBY = &H80
KO_OFF_MOBZ = &H84
KO_OFF_ID = &H680
KO_OFF_WH = &H6C0
KO_OFF_MobKord = &H428
KO_OFF_PtBase = &H1E8
KO_OFF_PtCount = &H304
KO_OFF_Pt = &H300
KO_OFF_MAXEXP = &HB98
KO_OFF_EXP = &HBA0
KO_OFF_MOB = &H644
KO_OFF_ZONE = &HC28
KO_OFF_NAMELEN = &H698
KO_OFF_NAME = &H688
KO_OFF_GOLD = &HB94
KO_OFF_MAXMP = &HB84
KO_OFF_MP = &HB88
KO_OFF_MAXHP = &H6B8
KO_OFF_HP = &H6BC
KO_OFF_LEVEL = &H6B4
KO_OFF_MINE = &H9BD
KO_OFF_SWIFT = &H7C4
KO_OFF_STATP = &HB7C
KO_OFF_STRP = &HBBC
KO_OFF_HPP = &HBC4
KO_OFF_DEXP = &HBCC
KO_OFF_INTP = &HBD4
KO_OFF_MPP = &HBDC
KO_OFF_SBARBase = &H1EC
KO_OFF_BSkPoint = &H16C
KO_OFF_SPoint1 = &H180
KO_OFF_SPoint2 = &H184
KO_OFF_SPoint3 = &H188
KO_OFF_SPoint4 = &H18C
''//==== Fonksiyon Offsetleri ====
KO_OFF_INV = &H1B8
KO_OFF_INV2 = &H220
KO_OFF_BANKBASE = &H20C
KO_OFF_ITEMBASE = &H208
KO_OFF_ITEMSLOT = &H128
KO_OFF_BANKCONT = &HFC
KO_OFF_SKILLBASE = &H1D0
KO_OFF_SKILLID = &H12C
''//==== Item Silme ====
KO_ITEMDESCALL = &H5E44A0
KO_ITEMDES = &H0
KO_ITEMDES2 = &HE27520
KO_FAKE_ITEM = &H577A90
KO_SH_HOOK = &H4EC24B
KO_SH_VALUE = &HB758C8
KO_SPD_HOOK = &H4EC24B + &H9D
KO_SUB_ADDR0 = &H8A5CF0
KO_SUB_ADDR1 = &H583310
KO_PTR_NRML = &HE240AC
KO_SMMB = &HE3DAC0
KO_SMMB_FNC = &H4FB9A0
KO_FNSB = &H4FB9A0
End Function
Function OffsetleriYükleUSKO()
KO_PTR_CHR = &HE4D140
KO_PTR_DLG = &HE33878
KO_PTR_PKT = &HE33844
KO_KEY_PTR = &HE33840
KO_SND_FNC = &H4977D0
KO_FNC_ISEN = &H54EC00
KO_CHAR_SERV = &HC81844
KO_NODC = &HB83540
KO_MULTI = &HBA22AC
KO_CHAT = &HE358E8

KO_OTO_BTN_PTR = &HE33874
KO_BTN_LEFT = &H4CA320
KO_BTN_RIGHT = &H4CA5C0
KO_BTN_LOGIN = &H4C6B50
KO_FLDB = &HE4D13C
KO_FNCZ = &H5221B0
KO_FNCB = &H522320
KO_FMBS = &H4F8280
KO_FPBS = &H4F91F0
KO_FNCX = &H522270
KO_FPOZ = &H561FF0
KO_FPOB = &H5619B0
KO_STMB = &H51B390
KO_ITOB = &HE4CF74
KO_ITEB = &HE4CF7C
KO_FPOX = &H69A010
'KO_FLPZ = &HE4D184
KO_FLPZ = ReadLong(ReadLong(KO_FLDB + &H40) + 4) + 4
'KO_FLPZ = ReadLong(ReadLong(ReadLong(ReadLong(KO_FLDB) + 40)))
KO_PM = &HC7F0C4
KO_PERI = &H577B40
KO_PERI_CLOOT_USKO = &HCDEE04
KO_PERI_MLOOT_USKO = &HCDEE08
KO_PERI_CLOOT_STEAM = &H0
KO_PERI_MLOOT_STEAM = &H0
KO_RECV_FNC = &H54B120
KO_RECV_PTR = &HB844A8
''//== XignCode Pointer ===
KO_CRE_THREAD = &H9CBE41
KO_LANC_BYPASS = &H9D0995
KO_XIGN_BYPASS = &H9CF75D
KO_XIGN_EXBPS1 = &H4B790C
KO_XIGN_EXBPS2 = &H4B7922
KO_XIGN_EXBPS3 = &H4B7931
KO_XIGN_PTR = &HE337A8
KO_XIGN_FNC = &H45B250
KO_XIGN_PRS = &H4B5200
''//==== Offsetler ====
KO_OFF_CLASS = &H6B0
KO_OFF_NT = &H6A8
KO_OFF_MOVE = &HFB8
KO_OFF_TYPE = &H3F0
KO_OFF_MX = &HFC4
KO_OFF_MZ = &HFC8
KO_OFF_MY = &HFCC
KO_OFF_X = &HD8
KO_OFF_Z = &HDC
KO_OFF_Y = &HE0
KO_OFF_MOBX = &H7C
KO_OFF_MOBY = &H80
KO_OFF_MOBZ = &H84
KO_OFF_ID = &H680
KO_OFF_WH = &H6C0
KO_OFF_MobKord = &H428
KO_OFF_PtBase = &H1E8
KO_OFF_PtCount = &H304
KO_OFF_Pt = &H300
KO_OFF_MAXEXP = &HB98
KO_OFF_EXP = &HBA0
KO_OFF_MOB = &H644
KO_OFF_ZONE = &HC28
KO_OFF_NAMELEN = &H698
KO_OFF_NAME = &H688
KO_OFF_GOLD = &HB94
KO_OFF_MAXMP = &HB84
KO_OFF_MP = &HB88
KO_OFF_MAXHP = &H6B8
KO_OFF_HP = &H6BC
KO_OFF_LEVEL = &H6B4
KO_OFF_MINE = &H9BD
KO_OFF_SWIFT = &H7C4
KO_OFF_STATP = &HB7C
KO_OFF_STRP = &HBBC
KO_OFF_HPP = &HBC4
KO_OFF_DEXP = &HBCC
KO_OFF_INTP = &HBD4
KO_OFF_MPP = &HBDC
KO_OFF_SBARBase = &H1EC
KO_OFF_BSkPoint = &H16C
KO_OFF_SPoint1 = &H180
KO_OFF_SPoint2 = &H184
KO_OFF_SPoint3 = &H188
KO_OFF_SPoint4 = &H18C
''//==== Fonksiyon Offsetleri ====
KO_OFF_INV = &H1B8
KO_OFF_INV2 = &H220
KO_OFF_BANKBASE = &H20C
KO_OFF_ITEMBASE = &H208
KO_OFF_ITEMSLOT = &H128
KO_OFF_BANKCONT = &HFC
KO_OFF_SKILLBASE = &H1D0
KO_OFF_SKILLID = &H12C
''//==== Item Silme ====
KO_ITEMDESCALL = &H5E3BA0
KO_ITEMDES = &H0
KO_ITEMDES2 = &HE36CB8
KO_FAKE_ITEM = &H577B40
KO_SH_HOOK = &H4EBC7B
KO_SH_VALUE = &HB83690
KO_SPD_HOOK = &H4EBC7B + &H9D
KO_SUB_ADDR0 = &H8B0920
KO_SUB_ADDR1 = &H5833C0
KO_PTR_NRML = &HE33850
KO_SMMB = &HE4D038
KO_SMMB_FNC = &H4FB3E0
KO_FNSB = &H4FB3E0
End Function
Public Function LongOku(addr As Long) As Long
Dim value As Long
ReadProcessMem KO_HANDLE, addr, value, 4, 0&
LongOku = value
End Function
Function ByteOku(pAddy As Long, Optional pHandle As Long) As Byte
Dim value As Byte
If pHandle <> 0 Then ReadProcessMem pHandle, pAddy, value, 1, 0& Else ReadProcessMem KO_HANDLE, pAddy, value, 1, 0&
ByteOku = value
End Function
Public Function ConvHEX2ByteArray(pStr As String, pbyte() As Byte)
On Error Resume Next
Dim i As Long
Dim j As Long
ReDim pbyte(1 To Len(pStr) / 2)
j = LBound(pbyte) - 1
For i = 1 To Len(pStr) Step 2
    j = j + 1
    pbyte(j) = CByte("&H" & mID(pStr, i, 2))
Next
End Function
Public Function InjectPatch(addr As Long, pStr As String)
Dim pbytes() As Byte
ConvHEX2ByteArray pStr, pbytes
WriteProcessMem KO_HANDLE, addr, pbytes(LBound(pbytes)), UBound(pbytes) - LBound(pbytes) + 1, 0&
End Function


Public Function ReadLong(addr As Long) As Long 'read a 4 byte value
    Dim value As Long
    ReadProcessMem KO_HANDLE, addr, value, 4, 0&
    ReadLong = value
End Function
Public Function ReadFloat(addr As Long) As Long 'read a float value
On Error Resume Next
    Dim value As Single
    ReadProcessMem KO_HANDLE, addr, value, 4, 0&
    ReadFloat = value
End Function
Public Function WriteFloat(addr As Long, Val As Single) 'write a float value
    WriteProcessMem KO_HANDLE, addr, Val, 4, 0&
End Function
Public Function WriteLong(addr As Long, Val As Long) ' write a 4 byte value
    WriteProcessMem KO_HANDLE, addr, Val, 4, 0&
End Function
Public Function WriteByte(addr As Long, Val As Byte) ' write a 1 byte value
    WriteProcessMem KO_HANDLE, addr, Val, 1, 0&
End Function
Public Function WriteByteArray(pAddy As Long, pmem() As Byte, pSize As Long)
    WriteProcessMem KO_HANDLE, pAddy, pmem(LBound(pmem)), pSize, 0&
End Function
Function KarakterX()
KarakterX = ReadFloat(KO_ADR_CHR + KO_OFF_X)
End Function
Function KarakterY()
KarakterY = ReadFloat(KO_ADR_CHR + KO_OFF_Y)
End Function
Function KarakterZ()
KarakterZ = ReadFloat(KO_ADR_CHR + KO_OFF_Z)
End Function

Function CharX()
CharX = ReadFloat(KO_ADR_CHR + KO_OFF_X)
End Function
Function CharY()
CharY = ReadFloat(KO_ADR_CHR + KO_OFF_Y)
End Function
Function CharZ()
CharZ = ReadFloat(KO_ADR_CHR + KO_OFF_Z)
End Function
Function MobX() As Long
MobX = ReadFloat(ReadLong(ReadLong(KO_PTR_DLG) + &H404) + &H7C)
End Function

Function MobY() As Long
MobY = ReadFloat(ReadLong(ReadLong(KO_PTR_DLG) + &H404) + &H84)
End Function
Function CharId()
CharId = Strings.mID(AlignDWORD(ReadLong(ReadLong(KO_PTR_CHR) + KO_OFF_ID)), 1, 4)
End Function

Function MobZ() As Long
MobZ = ReadFloat(ReadLong(ReadLong(KO_PTR_DLG) + &H404) + &H80)
End Function
Public Function EstablishMailSlot(ByVal MailSlotName As String, Optional MaxMessageSize As Long = 0, Optional ReadTimeOut As Long = 50) As Long
EstablishMailSlot = CreateMailslot(MailSlotName, MaxMessageSize, ReadTimeOut, ByVal 0&)
End Function

'Public Function getCallDiff2(Source As Long, Destination As Long) As Long
'Dim Diff As Long
'Diff = 0
'If Source > Destination Then
'    Diff = Source - Destination
'    If Diff > 0 Then getCallDiff = &HFFFFFFFB - Diff
'Else
'    getCallDiff = Destination - Source - 5
'End If
'End Function
Public Sub Paket(Paket As String)
Dim PaketByte() As Byte
ConvHEX2ByteArray Paket, PaketByte
SendPacket PaketByte
End Sub
Function SendPacket(pPacket() As Byte)
On Error Resume Next
Dim pSize As Long
Dim pCode() As Byte
pSize = UBound(pPacket) - LBound(pPacket) + 1
If BytesAddr = 0 Then
BytesAddr = VirtualAllocEx(KO_HANDLE, 0, 1024, MEM_COMMIT, PAGE_READWRITE)
End If
If BytesAddr <> 0 Then
    WriteByteArray BytesAddr, pPacket, pSize
    Hex2Byte "608B0D" & AlignDWORD(KO_PTR_PKT) & "68" & AlignDWORD(pSize) & "68" & AlignDWORD(BytesAddr) & "BF" & AlignDWORD(KO_SND_FNC) & "FFD761C3", pCode
    ExecuteRemoteCode pCode, True
End If
VirtualFreeEx KO_HANDLE, BytesAddr, 0, MEM_RELEASE&
End Function
Public Function YürüXY(x As Single, Y As Single) As Boolean
    If CInt(CharX) = CInt(x) And CInt(CharY) = CInt(Y) Then YürüXY = True: Exit Function
    WriteLong KO_ADR_CHR + KO_OFF_Go2, 2
    WriteFloat KO_ADR_CHR + KO_OFF_MX, x
    WriteFloat KO_ADR_CHR + KO_OFF_MY, Y
    WriteLong KO_ADR_CHR + KO_OFF_Go1, 1
    YürüXY = False: Exit Function
End Function
Public Function SpeedHack(XKor As Integer, YKor As Integer) As Boolean
If CInt(CharX) = XKor And CInt(CharY) = YKor Then SpeedHack = True: Exit Function
'SeksClub
Dim FarkX As Long, FarkY As Long
Dim ZýplaX As Integer, ZýplaY As Integer, i As Integer
FarkX = XKor - CharX
FarkY = YKor - CharY
ZýplaX = 2
ZýplaY = 2
If CharX = XKor And CharY = YKor Then
Exit Function
End If
For i = 1 To 5
If FarkX = -1 * i Or FarkX = i Then
ZýplaX = 1
ElseIf FarkY = -1 * i Or FarkY = i Then
ZýplaY = 1
End If
Next i
Dim oAnkiX As Long, oAnkiY As Long
oAnkiX = CharX
oAnkiY = CharY
If FarkX <> 0 Or FarkY <> 0 Then
If FarkX < 0 Then
WriteFloat ReadLong(KO_PTR_CHR) + KO_OFF_X, CharX - ZýplaX
ElseIf FarkX > 0 Then
WriteFloat ReadLong(KO_PTR_CHR) + KO_OFF_X, CharX + ZýplaX
End If
If FarkY < 0 Then
WriteFloat ReadLong(KO_PTR_CHR) + KO_OFF_Y, CharY - ZýplaY
ElseIf FarkY > 0 Then
WriteFloat ReadLong(KO_PTR_CHR) + KO_OFF_Y, CharY + ZýplaY
End If
Dim RetX As Long, RetY As Long
RetX = CharX
RetY = CharY
Paket "06" & AlignDWORD(CInt(oAnkiX) * 10, 4) & AlignDWORD(CInt(oAnkiY) * 10, 4) & AlignDWORD(CInt(CharZ) * 10, 4) & "2D0003" & AlignDWORD(CInt(RetX) * 10, 4) & AlignDWORD(CInt(RetY) * 10, 4) & AlignDWORD(CInt(CharZ) * 10, 4)
End If
SpeedHack = False
End Function
Public Function hex2Val(pStrhex As String) As Long
Dim TmpStr As String
Dim Tmphex As String
Dim i As Long
TmpStr = ""
For i = Len(pStrhex) To 1 Step -1
    Tmphex = Hex(Asc(mID(pStrhex, i, 1)))
    If Len(Tmphex) = 1 Then Tmphex = "0" & Tmphex
    TmpStr = TmpStr & Tmphex
Next
hex2Val = CLng("&H" & TmpStr)
End Function
Function ReadString(ByVal pAddy As Long, ByVal LSize As Long) As String
On Error Resume Next
Dim value As Byte
Dim tex() As Byte
On Error Resume Next
If LSize = 0 Then
Exit Function
Else
ReDim tex(1 To LSize)
ReadProcessMem KO_HANDLE, pAddy, tex(1), LSize, 0&
ReadString = StrConv(tex, vbUnicode)
End If
End Function
Function AlignDWORD(Dec As Long, Optional Length As Long = 8) As String

Dim DTH As String
DTH = Hex(Dec)
Select Case Len(Hex(Dec))
    Case 1
    AlignDWORD = Strings.Left("0" & DTH & "000000", Length)
    Case 2
    AlignDWORD = Strings.Left(DTH & "000000", Length)
    Case 3
    AlignDWORD = Strings.Left(Strings.mID(DTH, 2, 2) & "0" & Strings.Left(DTH, 1) & "0000", Length)
    Case 4
    AlignDWORD = Strings.Left(Strings.mID(DTH, 3, 2) & Strings.Left(DTH, 2) & "0000", Length)
    Case 5
    AlignDWORD = Strings.Left(Strings.mID(DTH, 4, 2) & Strings.mID(DTH, 2, 2) & "0" & Strings.Left(DTH, 1), Length) & "00"
    Case 6
    AlignDWORD = Strings.Left(Strings.mID(DTH, 5, 2) & Strings.mID(DTH, 3, 2) & Strings.Left(DTH, 2) & "00", Length)
    Case 7
    AlignDWORD = Strings.Left(Strings.mID(DTH, 6, 2) & Strings.mID(DTH, 4, 2) & Strings.mID(DTH, 2, 2) & "0" & Strings.Left(DTH, 1), Length)
    Case 8
    AlignDWORD = Strings.Left(Strings.mID(DTH, 7, 2) & Strings.mID(DTH, 5, 2) & Strings.mID(DTH, 3, 2) & Strings.Left(DTH, 2), Length)
End Select
End Function
Function CharName()
If ReadLong(ReadLong(KO_PTR_CHR) + &H698) > 15 Then
CharName = ReadString(ReadLong(ReadLong(KO_PTR_CHR) + &H688), ReadLong(ReadLong(KO_PTR_CHR) + &H698))
Else
CharName = ReadString(ReadLong(KO_PTR_CHR) + &H688, ReadLong(ReadLong(KO_PTR_CHR) + &H698))
End If
End Function
Function CharName2() As String
If ReadLong(ReadLong(KO_PTR_CHR) + &H698) > 15 Then
CharName2 = ReadString(ReadLong(ReadLong(KO_PTR_CHR) + KO_OFF_NAME), ReadLong(ReadLong(KO_PTR_CHR) + KO_OFF_NAME + 10))
Else
CharName2 = ReadString(ReadLong(KO_PTR_CHR) + KO_OFF_NAME, ReadLong(ReadLong(KO_PTR_CHR) + KO_OFF_NAME + 10))
End If
End Function
Function CharDC()
CharDC = ReadLong(ReadLong(KO_PTR_PKT) + &H40064)
End Function
Public Function oyunukapa()
On Error Resume Next
ret& = TerminateProcess(KO_HANDLE, 0&)
End Function
Function CharHP()
CharHP = ReadLong(KO_ADR_CHR + KO_OFF_HP)
End Function
Function CharMaxHP()
CharMaxHP = ReadLong(KO_ADR_CHR + KO_OFF_MAXHP)
End Function
Function CharMP()
CharMP = ReadLong(KO_ADR_CHR + KO_OFF_MP)
End Function
Function CharMaxMP()
CharMaxMP = ReadLong(KO_ADR_CHR + KO_OFF_MAXMP)
End Function
Function MobID()
MobID = Strings.mID(AlignDWORD(ReadLong(ReadLong(KO_PTR_CHR) + KO_OFF_MOB)), 1, 4)
End Function
Function MobLID()
MobLID = ReadLong(KO_ADR_CHR + KO_OFF_MOB)
End Function
Public Function Pause(ByVal delay As Single)
delay = Timer + delay
  Do
  DoEvents
  Sleep 1
  Loop While delay > Timer
End Function
Public Function HexItemID(ByVal Slot As Integer) As String
        Dim offset, x, offset3, offset4 As Long
        Dim Base, Sonuc As Long
        offset = ReadLong(KO_ADR_DLG + &H1B4)
        offset = ReadLong(offset + (&H20C + (4 * Slot)))
        
        Sonuc = ReadLong(ReadLong(offset + &H68)) + ReadLong(ReadLong(offset + &H6C))
        HexItemID = Strings.mID(AlignDWORD(Sonuc), 1, 8)
End Function
Public Function LongItemID(ByVal Slot As Integer) As Long
        Dim offset, x, offset3, offset4 As Long
        Dim Base, Sonuc As Long
        offset = ReadLong(KO_ADR_DLG + &H1B4)
        offset = ReadLong(offset + (&H20C + (4 * Slot)))

        LongItemID = ReadLong(ReadLong(offset + &H68)) + ReadLong(ReadLong(offset + &H6C))
        
End Function
Function GetItemCountInInv(ByVal Slot As Integer) As Long
        Dim offset, Offset2 As Long
        offset = ReadLong(KO_ADR_DLG + &H1B4)
        offset = ReadLong(offset + (&H20C + (4 * Slot)))
        Offset2 = ReadLong(offset + &H70)
        GetItemCountInInv = Offset2
End Function
Function GetItemCount() As Integer
        Dim ItemIDAdr As Long
        Dim ItemCount As Integer
        ItemCount = 0
        Dim n As Integer
        For n = 14 To 41
            ItemIDAdr = ReadLong(KO_ADR_DLG + &H1B4)
            ItemIDAdr = ReadLong(ItemIDAdr + (&H20C + (4 * (n))))
            ItemIDAdr = ReadLong(ItemIDAdr + &H68)
            ItemIDAdr = ReadLong(ItemIDAdr)
            If ItemIDAdr > 0 Then
                ItemCount = ItemCount + 1
            End If
        Next
        GetItemCount = ItemCount
    End Function
Function SCKontrol() As Boolean
If GetItemCountInInv(41) <= 26 Then
SCKontrol = False
Else
SCKontrol = True
End If
End Function
Function KarakterID()
KarakterID = Strings.mID(AlignDWORD(ReadLong(ReadLong(KO_PTR_CHR) + KO_OFF_ID)), 1, 4)
End Function
Function SýnýfBul() As Long
SýnýfBul = ReadLong(ReadLong(KO_PTR_CHR) + KO_OFF_CLASS)
End Function
Function JobBul() As String
If SýnýfBul = 201 Or SýnýfBul = 205 Or SýnýfBul = 206 Or SýnýfBul = 101 Or SýnýfBul = 105 Or SýnýfBul = 106 Then
JobBul = "Warrior"
End If
If SýnýfBul = 202 Or SýnýfBul = 207 Or SýnýfBul = 208 Or SýnýfBul = 102 Or SýnýfBul = 107 Or SýnýfBul = 108 Then
JobBul = "Rogue"
End If
If SýnýfBul = 203 Or SýnýfBul = 209 Or SýnýfBul = 210 Or SýnýfBul = 103 Or SýnýfBul = 109 Or SýnýfBul = 110 Then
JobBul = "Mage"
End If
If SýnýfBul = 204 Or SýnýfBul = 211 Or SýnýfBul = 212 Or SýnýfBul = 104 Or SýnýfBul = 111 Or SýnýfBul = 112 Then
JobBul = "Priest"
End If
End Function
Function DüþmanID()
DüþmanID = Strings.mID(AlignDWORD(ReadLong(ReadLong(KO_PTR_CHR) + KO_OFF_MOB)), 1, 4)
End Function
Function FormatHex(strHex As String, inLength As Integer)
On Error Resume Next
Dim newHex As String
Dim ZeroSpaces As Integer
ZeroSpaces = inLength - Len(strHex) '1
newHex = String(ZeroSpaces, "0") + strHex
Select Case Len(newHex)
Case 2
newHex = Left(newHex, 2)
Case 4
newHex = Right(newHex, 2) & Left(newHex, 2)
Case 6
newHex = Right(newHex, 2) & mID(newHex, 3, 2) & Left(newHex, 2)
Case 8
newHex = Right(newHex, 2) & mID(newHex, 5, 2) & mID(newHex, 3, 2) & Left(newHex, 2)
Case Else
End Select
FormatHex = newHex
End Function
Public Function ByteDizisiYaz(pAddy As Long, pmem() As Byte, pSize As Long)
    WriteProcessMem KO_HANDLE, pAddy, pmem(LBound(pmem)), pSize, 0&
End Function
Sub SýraByteOku(addr As Long, pmem() As Byte, pSize As Long)
Dim value As Byte
On Error Resume Next
ReDim pmem(1 To pSize) As Byte
ReadProcessMem KO_HANDLE, addr, pmem(1), pSize, 0&
End Sub
Function MobUzaklýK() As Long
On Error Resume Next
If MobID = "FFFF" Then MobUzaklýK = 255: Exit Function
MobUzaklýK = Sqr((MobX - KarakterX) ^ 2 + (MobY - KarakterY) ^ 2)
End Function
Function MerkezeUzaklýk() As Long
On Error Resume Next
If MobID = "FFFF" Then MerkezeUzaklýk = 255: Exit Function
MerkezeUzaklýk = Sqr((MobX - Label1.Caption) ^ 2 + (MobY - Label2.Caption) ^ 2)
End Function
Public Function StringToHex(ByVal StrToHex As String) As String
Dim strTemp, strReturn As String, i As Long
    For i = 1 To Len(StrToHex)
        strTemp = Hex$(Asc(mID$(StrToHex, i, 1)))
        If Len(strTemp) = 1 Then strTemp = "0" & strTemp
        strReturn = strReturn & strTemp
    Next i
    StringToHex = strReturn
End Function
Public Sub PartyAt(ad As String)
Dim a As String
a = Strings.mID$(AlignDWORD(Len(ad)), 1, 2)
Paket "2f03" + a + "00" + StringToHex(ad)
Paket "2f01" + a + "00" + StringToHex(ad)
End Sub
Function ReadByte(pAddy As Long) As Byte
Dim value As Byte
ReadProcessMem KO_HANDLE, pAddy, value, 1, 0&
ReadByte = value
End Function

Public Function LongYaz(addr As Long, Val As Long)
    WriteProcessMem KO_HANDLE, addr, Val, 4, 0&
End Function
Public Function ByteYaz(addr As Long, pval As Byte)
Dim pbw As Long
WriteProcessMem KO_HANDLE, addr, pval, 1, pbw
End Function
Public Function FloatOku(addr As Long) As Long
On Error Resume Next
    Dim value As Single
    ReadProcessMem KO_HANDLE, addr, value, 4, 0&
    FloatOku = value
End Function

Public Function FloatYaz(addr As Long, Val As Single)
    WriteProcessMem KO_HANDLE, addr, Val, 4, 0&
End Function
Public Sub HookBul()
Dim HookIndex As Integer, TmpAddr As Long
TmpAddr = ReadLong(ReadLong(KO_PTR_DLG)) + &H8
KO_RECVHK = TmpAddr + (HookIndex * 4)
KO_RCVHKB = ReadLong(KO_RECVHK)
End Sub
Function ExecuteRemoteCode(pCode() As Byte, Optional WaitExecution As Boolean = False) As Long
Dim hThread As Long, ThreadID As Long, ret As Long
Dim SE As SECURITY_ATTRIBUTES
SE.nLength = Len(SE)
SE.bInheritHandle = False
ExecuteRemoteCode = 0
If FuncPtr = 0 Then
FuncPtr = VirtualAllocEx(KO_HANDLE, 0, 1024, MEM_COMMIT, PAGE_READWRITE)
End If
If FuncPtr <> 0 Then
    WriteByteArray FuncPtr, pCode, UBound(pCode) - LBound(pCode) + 1
    hThread = CreateRemoteThread(ByVal KO_HANDLE, SE, 0, ByVal FuncPtr, 0&, 0&, ThreadID)
    If hThread Then
    WaitForSingleObject hThread, INFINITE
    ExecuteRemoteCode = ThreadID
    End If
End If
CloseHandle hThread
End Function
Public Function Hex2Byte(Paket As String, pbyte() As Byte)
On Error Resume Next
Dim i As Long
Dim j As Long
ReDim pbyte(1 To Len(Paket) / 2)
j = LBound(pbyte) - 1
For i = 1 To Len(Paket) Step 2
    j = j + 1
    pbyte(j) = CByte("&H" & mID(Paket, i, 2))
Next
End Function


Function periaç()
Dim pCode() As Byte
ConvHEX2ByteArray ("608B0D" + AlignDWORD(KO_PTR_CHR) + "6A006858BFB929B8" + AlignDWORD(KO_PERI) + "FFD061C3"), pCode
ExecuteRemoteCode pCode, True
End Function
