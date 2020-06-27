Attribute VB_Name = "Functions"
Public KOPTRCHR As Long
Public KO_SPD_HOOK As Long
Public Kosu As Boolean
Public SeriTimer As Single
Public SeriHiz As Long
Public DefaultTimer As Single
Public DefaultHiz As Long
Dim memWalk As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function WriteProcessMem Lib "kernel32" Alias "WriteProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
'''''SPEEDHACK BASLANGIC'''''''
Sub SpeedLe()
If memWalk = 0 Then memWalk = VirtualAllocEx(KO_HANDLE, 0, 1024, MEM_COMMIT, &H40&)
InjectYaz KO_SPD_HOOK, "E9" & AlignDWORD(getCallDiff(KO_SPD_HOOK, memWalk)) & "9090"
SpeedEsLe
End Sub

Public Function GetItemID(i As Integer) 'item id bulan fonksiyon
Dim Base As Long, Lng1 As Long, Lng2 As Long, Lng3 As Long, Lng4 As Long
Dim lngItemID As Long, lngItemID_Ext As Long, lngItemNameLen As Long, AdrItemName As Long
Base = ReadLong(KO_PTR_DLG)
Lng1 = ReadLong(Base + &H1B8)
Lng2 = ReadLong(Lng1 + (&H21C + (4 * i)))
Lng3 = ReadLong(Lng2 + &H68)
Lng4 = ReadLong(Lng2 + &H6C)
lngItemID = ReadLong(Lng3)
lngItemID_Ext = ReadLong(Lng4)
lngItemID = lngItemID + lngItemID_Ext
GetItemID = lngItemID
End Function
Public Sub InventoryOku()
      Dim tmpBase As Long, tmpLng1 As Long, tmpLng2 As Long, tmpLng3 As Long, tmpLng4 As Long
      Dim lngItemID As Long, lngItemID_Ext As Long, lngItemNameLen As Long, AdrItemName As Long
      Dim ItemNameB() As Byte
      Dim ItemName As String
      Dim i As Integer
      
      tmpBase = ReadLong(KO_PTR_DLG)
      tmpLng1 = ReadLong(tmpBase + &H1B8)
      Form1.canta.Clear
        For i = 18 To 45
          tmpLng2 = ReadLong(tmpLng1 + (&H210 + (4 * i)))
          tmpLng3 = ReadLong(tmpLng2 + &H68)
          tmpLng4 = ReadLong(tmpLng2 + &H6C)
         
          lngItemID = ReadLong(tmpLng3)
          lngItemID_Ext = ReadLong(tmpLng4)
          lngItemID = lngItemID + lngItemID_Ext
          lngItemNameLen = ReadLong(tmpLng3 + &H1C)
          If lngItemNameLen > 15 Then
          AdrItemName = ReadLong(tmpLng3 + &HC)
          Else
          AdrItemName = tmpLng3 + &HC
          End If
          
          ItemName = ""
          If lngItemNameLen > 0 Then
               SýraByteOku AdrItemName, ItemNameB, lngItemNameLen
               ItemName = StrConv(ItemNameB, vbUnicode) 'convert it to string
          End If
       
            Form1.canta.AddItem Form1.canta.ListCount + 1 & "-) " & ItemName
        If ItemName <> "" Then
        End If
      Next
End Sub
Public Function HexItemID(ByVal Slot As Integer) As String
        Dim offset, x, offset3, offset4 As Long
        Dim Base, Sonuc As Long
        offset = ReadLong(KO_ADR_DLG + &H1B8)
        offset = ReadLong(offset + (&H210 + (4 * Slot))) 'inventory slot
          'item id adress
        
        Sonuc = ReadLong(ReadLong(offset + &H68)) + ReadLong(ReadLong(offset + &H6C))
        HexItemID = Strings.mID(AlignDWORD(Sonuc), 1, 8)
End Function
Public Function LongItemID(ByVal Slot As Integer) As Long
        Dim offset, x, offset3, offset4 As Long
        Dim Base, Sonuc As Long
        offset = ReadLong(KO_ADR_DLG + &H1B8)
        offset = ReadLong(offset + (&H210 + (4 * Slot))) 'inventory slot
          'item id adress
        
        LongItemID = ReadLong(ReadLong(offset + &H68)) + ReadLong(ReadLong(offset + &H6C))
        
End Function
Function GetItemCountInInv(ByVal Slot As Integer) As Long
        Dim offset, Offset2 As Long
        offset = ReadLong(KO_ADR_DLG + &H1B8)
        offset = ReadLong(offset + (&H210 + (4 * Slot)))
        Offset2 = ReadLong(offset + &H70)
        GetItemCountInInv = Offset2
End Function
Function GetItemCount() As Integer
        Dim ItemIDAdr As Long
        Dim ItemCount As Integer
        ItemCount = 0
        Dim n As Integer
        For n = 14 To 41
            ItemIDAdr = ReadLong(KO_ADR_DLG + &H1B8)
            ItemIDAdr = ReadLong(ItemIDAdr + (&H210 + (4 * (n))))
            ItemIDAdr = ReadLong(ItemIDAdr + &H68)
            ItemIDAdr = ReadLong(ItemIDAdr)
            If ItemIDAdr > 0 Then
                ItemCount = ItemCount + 1
            End If
        Next
        GetItemCount = ItemCount
    End Function
    Public Function ItemIDbul()
Dim ItemNo As String
Dim ItemNo1 As String
Dim ItemNo2 As String
Dim ItemNo3 As String
Dim ItemNo4 As String
Dim ItemNo5 As String
Dim ItemNo6 As String
Dim ItemNo7 As String
Dim ItemNo8 As String
Dim ItemNo9 As String

Dim i1 As Integer
Dim i2 As Integer
Dim i3 As Integer
Dim i4 As Integer
Dim i5 As Integer
Dim i6 As Integer
Dim i7 As Integer
Dim i8 As Integer
Dim i9 As Integer
Dim i10 As Integer

For i1 = 18 To 18
 For i2 = 19 To 19
  For i3 = 20 To 20
 For i4 = 21 To 21
  For i5 = 22 To 22
  For i6 = 23 To 23
  For i7 = 24 To 24
 For i8 = 25 To 25
  For i9 = 26 To 26
 For i10 = 27 To 27
 
 
ItemNo = HexItemID(i1)
Form1.Text5.Text = ItemNo
ItemNo1 = HexItemID(i2)
ItemNo2 = HexItemID(i3)
ItemNo3 = HexItemID(i4)
ItemNo4 = HexItemID(i5)
ItemNo5 = HexItemID(i6)
ItemNo6 = HexItemID(i7)
ItemNo7 = HexItemID(i8)
ItemNo8 = HexItemID(i9)
ItemNo9 = HexItemID(i10)


Next i10
Next i9
Next i8
Next i7
Next i6
Next i5
Next i4
Next i3
Next i2
Next i1
    End Function
Public Function Upgrade2()
Dim ScrollID As String
Dim ItemNo As String
Dim ItemNo1 As String
Dim ItemNo2 As String
Dim ItemNo3 As String
Dim ItemNo4 As String
Dim ItemNo5 As String
Dim ItemNo6 As String
Dim ItemNo7 As String
Dim ItemNo8 As String
Dim ItemNo9 As String
Dim ItemNo10 As String
Dim ItemNo11 As String
Dim ItemNo12 As String
Dim ItemNo13 As String
Dim ItemNo14 As String
Dim ItemNo15 As String
Dim ItemNo16 As String
Dim ItemNo17 As String
Dim ItemNo18 As String
Dim ItemNo19 As String
Dim ItemNo20 As String
Dim ItemNo21 As String
Dim ItemNo22 As String
Dim ItemNo23 As String
Dim ItemNo24 As String
Dim ItemNo25 As String
Dim ItemNo26 As String
Dim ItemNo27 As String

Dim i1 As Integer
Dim i2 As Integer
Dim i3 As Integer
Dim i4 As Integer
Dim i5 As Integer
Dim i6 As Integer
Dim i7 As Integer
Dim i8 As Integer
Dim i9 As Integer
Dim i10 As Integer
Dim i11 As Integer
Dim i12 As Integer
Dim i13 As Integer
Dim i14 As Integer
Dim i15 As Integer
Dim i16 As Integer
Dim i17 As Integer
Dim i18 As Integer
Dim i19 As Integer
Dim i20 As Integer
Dim i21 As Integer
Dim i22 As Integer
Dim i23 As Integer
Dim i24 As Integer
Dim i25 As Integer
Dim i26 As Integer
Dim i27 As Integer
Dim i28 As Integer

 For i1 = 18 To 18
 For i2 = 19 To 19
  For i3 = 20 To 20
 For i4 = 21 To 21
  For i5 = 22 To 22
 For i6 = 23 To 23
  For i7 = 24 To 24
 For i8 = 25 To 25
  For i9 = 26 To 26
 For i10 = 27 To 27
  For i11 = 28 To 28
 For i12 = 29 To 29
  For i13 = 30 To 30
 For i14 = 31 To 31
  For i15 = 32 To 32
 For i16 = 33 To 33
  For i17 = 34 To 34
 For i18 = 35 To 35
  For i19 = 36 To 36
 For i20 = 37 To 37
  For i21 = 38 To 38
 For i22 = 39 To 39
  For i23 = 40 To 40
 For i24 = 41 To 41
 For i25 = 42 To 42
  For i26 = 43 To 43
 For i27 = 44 To 44
 For i28 = 45 To 45
 
ItemNo = HexItemID(i1)
ItemNo1 = HexItemID(i2)
ItemNo2 = HexItemID(i3)
ItemNo3 = HexItemID(i4)
ItemNo4 = HexItemID(i5)
ItemNo5 = HexItemID(i6)
ItemNo6 = HexItemID(i7)
ItemNo7 = HexItemID(i8)
ItemNo8 = HexItemID(i9)
ItemNo9 = HexItemID(i10)
ItemNo10 = HexItemID(i11)
ItemNo11 = HexItemID(i12)
ItemNo12 = HexItemID(i13)
ItemNo13 = HexItemID(i14)
ItemNo14 = HexItemID(i15)
ItemNo15 = HexItemID(i16)
ItemNo16 = HexItemID(i17)
ItemNo17 = HexItemID(i18)
ItemNo18 = HexItemID(i19)
ItemNo19 = HexItemID(i20)
ItemNo20 = HexItemID(i21)
ItemNo21 = HexItemID(i22)
ItemNo22 = HexItemID(i23)
ItemNo23 = HexItemID(i24)
ItemNo24 = HexItemID(i25)
ItemNo25 = HexItemID(i26)
ItemNo26 = HexItemID(i27)
ItemNo27 = HexItemID(i28)



        Select Case Form1.Combo1.ListIndex
            Case 0: ScrollID = AlignDWORD(379221000)
            Case 1: ScrollID = AlignDWORD(379205000)
            Case 2: ScrollID = AlignDWORD(379016000)
            Case 3: ScrollID = AlignDWORD(379021000)
        End Select

If Form1.canta.Selected(0) Then
Paket "5B02" + "01" + "1427" + ItemNo + "00" + ScrollID + Form1.Text21.Text + "00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF"
End If
If Form1.canta.Selected(1) Then
Paket "5B02" + "01" + "1427" + ItemNo1 + "01" + ScrollID + Form1.Text21.Text + "00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF"
End If
If Form1.canta.Selected(2) Then
Paket "5B02" + "01" + "1427" + ItemNo2 + "02" + ScrollID + Form1.Text21.Text + "00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF"
End If
If Form1.canta.Selected(3) Then
Paket "5B02" + "01" + "1427" + ItemNo3 + "03" + ScrollID + Form1.Text21.Text + "00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF"
End If
If Form1.canta.Selected(4) Then
Paket "5B02" + "01" + "1427" + ItemNo4 + "04" + ScrollID + Form1.Text21.Text + "00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF"
End If
If Form1.canta.Selected(5) Then
Paket "5B02" + "01" + "1427" + ItemNo5 + "05" + ScrollID + Form1.Text21.Text + "00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF"
End If
If Form1.canta.Selected(6) Then
 Paket "5B02" + "01" + "1427" + ItemNo6 + "06" + ScrollID + Form1.Text21.Text + "00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF"
End If
If Form1.canta.Selected(7) Then
 Paket "5B02" + "01" + "1427" + ItemNo7 + "07" + ScrollID + Form1.Text21.Text + "00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF"
End If
If Form1.canta.Selected(8) Then
  Paket "5B02" + "01" + "1427" + ItemNo8 + "08" + ScrollID + Form1.Text21.Text + "00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF"
End If
If Form1.canta.Selected(9) Then
  Paket "5B02" + "01" + "1427" + ItemNo9 + "09" + ScrollID + Form1.Text21.Text + "00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF"
End If
If Form1.canta.Selected(10) Then
 Paket "5B02" + "01" + "1427" + ItemNo10 + "0A" + ScrollID + Form1.Text21.Text + "00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF"
End If
If Form1.canta.Selected(11) Then
  Paket "5B02" + "01" + "1427" + ItemNo11 + "0B" + ScrollID + Form1.Text21.Text + "00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF"
End If
If Form1.canta.Selected(12) Then
  Paket "5B02" + "01" + "1427" + ItemNo12 + "0C" + ScrollID + Form1.Text21.Text + "00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF"
End If
If Form1.canta.Selected(13) Then
  Paket "5B02" + "01" + "1427" + ItemNo13 + "0D" + ScrollID + Form1.Text21.Text + "00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF"
End If
If Form1.canta.Selected(14) Then
  Paket "5B02" + "01" + "1427" + ItemNo14 + "0E" + ScrollID + Form1.Text21.Text + "00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF"
End If
If Form1.canta.Selected(15) Then
  Paket "5B02" + "01" + "1427" + ItemNo15 + "0F" + ScrollID + Form1.Text21.Text + "00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF"
End If
If Form1.canta.Selected(16) Then
 Paket "5B02" + "01" + "1427" + ItemNo16 + Hex(CLng("16")) + ScrollID + Form1.Text21.Text + "00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF"
End If
If Form1.canta.Selected(17) Then
 Paket "5B02" + "01" + "1427" + ItemNo17 + Hex(CLng("17")) + ScrollID + Form1.Text21.Text + "00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF"
End If
If Form1.canta.Selected(18) Then
 Paket "5B02" + "01" + "1427" + ItemNo18 + Hex(CLng("18")) + ScrollID + Form1.Text21.Text + "00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF"
End If
If Form1.canta.Selected(19) Then
 Paket "5B02" + "01" + "1427" + ItemNo19 + Hex(CLng("19")) + ScrollID + Form1.Text21.Text + "00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF"
End If
If Form1.canta.Selected(20) Then
Paket "5B02" + "01" + "1427" + ItemNo20 + Hex(CLng("20")) + ScrollID + Form1.Text21.Text + "00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF"
End If
If Form1.canta.Selected(21) Then
Paket "5B02" + "01" + "1427" + ItemNo21 + Hex(CLng("21")) + ScrollID + Form1.Text21.Text + "00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF"
End If
If Form1.canta.Selected(22) Then
Paket "5B02" + "01" + "1427" + ItemNo22 + Hex(CLng("22")) + ScrollID + Form1.Text21.Text + "00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF"
End If
If Form1.canta.Selected(23) Then
 Paket "5B02" + "01" + "1427" + ItemNo23 + Hex(CLng("23")) + ScrollID + Form1.Text21.Text + "00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF"
End If
If Form1.canta.Selected(24) Then
  Paket "5B02" + "01" + "1427" + ItemNo24 + Hex(CLng("24")) + ScrollID + Form1.Text21.Text + "00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF"
End If
If Form1.canta.Selected(25) Then
  Paket "5B02" + "01" + "1427" + ItemNo25 + Hex(CLng("25")) + ScrollID + Form1.Text21.Text + "00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF"
End If
If Form1.canta.Selected(26) Then
  Paket "5B02" + "01" + "1427" + ItemNo26 + Hex(CLng("26")) + ScrollID + Form1.Text21.Text + "00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF"
End If
If Form1.canta.Selected(27) Then
 Paket "5B02" + "01" + "1427" + ItemNo27 + Hex(CLng("27")) + ScrollID + Form1.Text21.Text + "00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF00000000FF"
End If

Next i28
Next i27
Next i26
Next i25
Next i24
Next i23
Next i22
Next i21
Next i20
Next i19
Next i18
Next i17
Next i16
Next i15
Next i14
Next i13
Next i12
Next i11
Next i10
Next i9
Next i8
Next i7
Next i6
Next i5
Next i4
Next i3
Next i2
Next i1

End Function
Sub SpeedEsLe()
Dim sHook As String
sHook = "508BC183C0078038007413807801FF7505C600E2EB0880382D7403C6002D90807801FF7516" & _
"C705" & AlignDWORD(KO_SH_VALUE) & AlignDWORD(SingleToHex(DefaultTimer)) & _
"C705" & AlignDWORD(KOPTRCHR + KO_OFF_SWIFT) & AlignDWORD(DefaultHiz) & _
"EB14" & _
"C705" & AlignDWORD(KO_SH_VALUE) & AlignDWORD(SingleToHex(SeriTimer)) & _
"C705" & AlignDWORD(KOPTRCHR + KO_OFF_SWIFT) & AlignDWORD(SeriHiz) & _
"58518B0D" & AlignDWORD(KO_PTR_PKT) & "E9"
sHook = sHook & AlignDWORD(getCallDiff(memWalk + CLng("&H" & Hex(Len(sHook) / 2)), KO_SPD_HOOK + 7))
InjectYaz memWalk, sHook
End Sub
Function InventoryIDAra(ItemID As String) As Long
InventoryOku
Dim i As Integer, a As Long
For i = 19 To 42
a = InStr(1, Right(ItemIntID(i), 1), ItemID, vbTextCompare)
If a <> 0 Then
InventoryIDAra = i
Exit Function
Else
InventoryIDAra = 0
End If
Next
End Function
Sub SpeedAl()
WriteFloat KO_SH_VALUE, DefaultTimer
InjectYaz KO_SPD_HOOK, "518B0D8004DF00"
WriteLong (KO_PTR_CHR + KO_OFF_SWIFT), DefaultHiz
End Sub

Function SingleToHex(ByVal Tmp As Single) As Long
Dim i As Long
CopyMemory i, Tmp, 4
SingleToHex = i
End Function

Function getCallDiff(ByVal Source As Long, ByVal Destination As Long) As Long
If Source > Destination Then
Dim Diff As Long
Diff = Source - Destination
If Diff > 0 Then
getCallDiff = &HFFFFFFFB - Diff
End If
Else
getCallDiff = (Destination - Source) - 5
End If
End Function

Public Function InjectYaz(addr As Long, pStr As String)
Dim pbytes() As Byte
Hex2Byte pStr, pbytes
WriteProcessMem KO_HANDLE, addr, pbytes(LBound(pbytes)), UBound(pbytes) - LBound(pbytes) + 1, 0&
End Function
'''''''SPEED HACK BITIR''''''''''
Function Otokutuac()
On Error Resume Next
Dim mem As Long
mem = VirtualAllocEx(KO_HANDLE, 0, 1024, MEM_COMMIT, &H40)
InjectPatch mem, "558BEC81C4D8FDFFFF53565760FF750CFF7508B8" + AlignDWORD(KO_RECV_FNC) + "FFD0618B45088B400833D28A1083FA2375348D8DF0FEFFFF8D95F1FEFFFF894DFC8B4003C685F0FEFFFF248902608B0D" + AlignDWORD(KO_PTR_PKT) + "6A05FF75FCB8" + AlignDWORD(KO_SND_FNC) + "FFD061E9AA00000033C98A0883F9240F859D0000008B50018955F8BEC8F040008DBDD8FEFFFFB906000000F3A533DB83C00633C98D95D8FEFFFF03C38B1883C006891A4183C20483F90672F066C745F6000033FF8DB5D8FEFFFF833E00744A8D85D8FDFFFF8D95D9FDFFFF8945F08D85DDFDFFFFC685D8FDFFFF268B4DF8890A8D8DE1FDFFFF8B168910668B5DF666891966FF45F6608B0D" + AlignDWORD(KO_PTR_PKT) + "6A0BFF75F0B8" + AlignDWORD(KO_SND_FNC) + "FFD0614783C60483FF0672A85F5E5B8BE55DC20800"
WriteLong KO_RECV_PTR, mem
End Function

Public Function MobHpOku() As Long
Dim xCode() As Byte, Paket As String
Dim MobID As Long
MobID = ReadLong(ReadLong(KO_PTR_CHR) + KO_OFF_MOB)
If FuncPtr = 0 Then: FuncPtr = VirtualAllocEx(KO_HANDLE, 0, 1024, MEM_COMMIT, PAGE_READWRITE)
If MobID > 9999 Then
Paket = "608B0D" & AlignDWORD(KO_FLDB) & "6A01" & "68" & AlignDWORD(MobID) & "BF" & AlignDWORD(KO_FMBS) & "FFD7" & "A3" & AlignDWORD(FuncPtr) & "61C3"
Else
Paket = "608B0D" & AlignDWORD(KO_FLDB) & "6A01" & "68" & AlignDWORD(MobID) & "BF" & AlignDWORD(KO_FPBS) & "FFD7" & "A3" & AlignDWORD(FuncPtr) & "61C3"

End If
Hex2Byte Paket, xCode: ExecuteRemoteCode xCode, True
MobHpOku = ReadLong(ReadLong(FuncPtr) + KO_OFF_HP)
End Function
Public Function MobHpOkuMax() As Long
Dim xCode() As Byte, Paket As String
Dim MobID As Long
MobID = ReadLong(ReadLong(KO_PTR_CHR) + KO_OFF_MOB)
If FuncPtr = 0 Then: FuncPtr = VirtualAllocEx(KO_HANDLE, 0, 1024, MEM_COMMIT, PAGE_READWRITE)
If MobID > 9999 Then
Paket = "608B0D" & AlignDWORD(KO_FLDB) & "6A01" & "68" & AlignDWORD(MobID) & "BF" & AlignDWORD(KO_FMBS) & "FFD7" & "A3" & AlignDWORD(FuncPtr) & "61C3"
Else
Paket = "608B0D" & AlignDWORD(KO_FLDB) & "6A01" & "68" & AlignDWORD(MobID) & "BF" & AlignDWORD(KO_FPBS) & "FFD7" & "A3" & AlignDWORD(FuncPtr) & "61C3"

End If
Hex2Byte Paket, xCode: ExecuteRemoteCode xCode, True
MobHpOkuMax = ReadLong(ReadLong(FuncPtr) + KO_OFF_MAXHP)
End Function

Public Function targetName() As String
Dim PaketByte() As Byte, Paket As String
Dim MobID As Long
MobID = ReadLong(ReadLong(KO_PTR_CHR) + KO_OFF_MOB)
If FuncPtr = 0 Then: FuncPtr = VirtualAllocEx(KO_HANDLE, 0, 1024, MEM_COMMIT, PAGE_READWRITE)
If MobID > 9999 Then
Paket = "608B0D" & AlignDWORD(KO_FLDB) & "6A01" & "68" & AlignDWORD(MobID) & "BF" & AlignDWORD(KO_FMBS) & "FFD7" & "A3" & AlignDWORD(FuncPtr) & "61C3"
Else
Paket = "608B0D" & AlignDWORD(KO_FLDB) & "6A01" & "68" & AlignDWORD(MobID) & "BF" & AlignDWORD(KO_FPBS) & "FFD7" & "A3" & AlignDWORD(FuncPtr) & "61C3"
End If
ConvHEX2ByteArray Paket, PaketByte: ExecuteRemoteCode PaketByte, True
targetName = ReadLong(FuncPtr)
End Function

Function TargetNameBase(Base As Long) As String
If ReadLong(Base + KO_OFF_NAME + 20) >= 20 Then
TargetNameBase = ReadStringAuto(ReadLong(Base + KO_OFF_NAME))
Else
TargetNameBase = ReadStringAuto(Base + KO_OFF_NAME)
End If
End Function

Public Function ReadStringAuto(addr As Long) As String
    Dim aRr(255) As Byte, bu As String
    If KO_HANDLE > 0 Then
        If ReadByte(addr + &H10) > 15 Then
            addr = ReadLong(addr)
        End If
        ReadProcessMem KO_HANDLE, addr, aRr(0), 255, 0&
        For i = 0 To 255
            bu = Chr$(aRr(i))
            If Asc(bu) = 0 Then
                Exit For
            End If
            ret = ret & bu
        Next
        ReadStringAuto = Trim$(ret)
    End If
End Function
Function GetTargetBase(target As Long)
    Dim pCode() As Byte, pStr As String, KO_FNC As Long
    
    If target > 9999 Then
        KO_FNC = KO_FMBS
    ElseIf target > 0 Then
        KO_FNC = KO_FPBS
    Else
        Exit Function
    End If

    If FuncPtr <> 0 Then
        pStr = "60" & _
              "8B0D" & _
              AlignDWORD(KO_FLDB) & _
              "6A01" & _
              "68" & _
              AlignDWORD(target) & _
              "BF" & _
              AlignDWORD(KO_FNC) & _
              "FFD7" & _
              "A3" & _
              AlignDWORD(FuncPtr) & _
              "61C3"
        Hex2Byte pStr, pCode
        ExecuteRemoteCode pCode, True
        GetTargetBase = ReadLong(FuncPtr)
    End If
End Function
Function GetDistance(ChrkorX, ChrkorY, HedefX, HedefY) As Long
On Error Resume Next
GetDistance = Sqr((HedefX - ChrkorX) ^ 2 + (HedefY - ChrkorY) ^ 2)
End Function
Function GetAllPLayer(lsts As ListBox)
Dim EBP As Long, ESI As Long, EAX As Long, FEnd As Long
Dim LDist As Long, CrrDist As Long, LID As Long, LBase As Long, LMoBID As Long, zaman As Long
Dim base_addr As Long
LDist = 99999
zMobName = ""
zMobX = 0
zMobY = 0
zMobZ = 0
zMobID = 0
zMobDistance = 0
EBP = ReadLong(ReadLong(KO_FLDB) + &H40)
FEnd = ReadLong(ReadLong(ReadLong(ReadLong(KO_FLDB) + &H40) + 4) + 4)
ESI = ReadLong(EBP)
While ESI <> EBP
base_addr = ReadLong(ESI + &H10)
If base_addr = 0 Then Exit Function
If ReadLong(base_addr + &H6A8) = 0 And ByteOku(base_addr + &H2A0) <> 10 Then
    CrrDist = GetDistance(CharX, CharY, ReadFloat(base_addr + KO_OFF_X), ReadFloat(base_addr + KO_OFF_Y))
        If CrrDist < LDist Then
            LID = ReadLong(base_addr + KO_OFF_ID)
            LBase = base_addr
            LDist = CrrDist
        End If
End If
EAX = ReadLong(ESI + 8)
    If ReadLong(ESI + 8) <> FEnd Then
        While ReadLong(EAX) <> FEnd
        EAX = ReadLong(EAX)
        Wend
    ESI = EAX
    Else
    EAX = ReadLong(ESI + 4)
        While ESI = ReadLong(EAX + 8)
        ESI = EAX
        EAX = ReadLong(EAX + 4)
        Wend
            If ReadLong(ESI + 8) <> EAX Then
            ESI = EAX
            End If
    End If
Dim Name As String
If ReadLong(LBase + KO_OFF_NAMELEN) >= 15 Then
Name = ReadString(ReadLong(LBase + KO_OFF_NAME), False, ReadLong(LBase + KO_OFF_NAMELEN)) 'BUDA TAMAM
Else
Name = ReadString(LBase + KO_OFF_NAME, False, ReadLong(LBase + KO_OFF_NAMELEN))
End If
If ReadLong(base_addr + &H6A8) = 0 Then
If ListeAra(Name, lsts) = False Then MobListe.List2.AddItem Name
End If
Wend
End Function
Public Function ListeAra(Aranan As String, Liste) As Boolean
Dim i As Long
For i = 0 To Liste.ListCount
If Liste.List(i) = Aranan Then ListeAra = True: Exit For Else: ListeAra = False
Next: End Function

Public Function AraText(Kelime, Cümle) As Boolean
Dim i As Long, Aranan As String
For i = 1 To Len(Cümle): Aranan = mID(Cümle, i, Len(Kelime))
If Aranan = Kelime Then AraText = True: Exit For Else: AraText = False
Next
End Function
Function GetTargetable(Base As Long) As Boolean
    Dim pCode() As Byte, pStr As String

    If FuncPtr <> 0 Then
        pStr = "608B0D" & AlignDWORD(KO_PTR_CHR) & _
            "68" & AlignDWORD(Base) & _
            "B8" & AlignDWORD(KO_FNC_ISEN) & _
            "FFD0A2" & AlignDWORD(FuncPtr) & _
            "61C3"
        Hex2Byte pStr, pCode
        
       ' ExecuteRemoteCode pCode, True
        GetTargetable = True 'ReadByte(FuncPtr)
    End If
End Function
Public Function Runner(crx As Single, cry As Single)
'Sabitle
On Error Resume Next
Dim zipla, x, Y, uzak, a, b, d, e, i, isrtx, isrty
Dim tx As Single, ty As Single
Dim x1 As Single, y1 As Single
Dim bykx, byky, kckx, kcky
zipla = 3.5
tx = ReadFloat(ReadLong(KO_PTR_CHR) + KO_OFF_X)
ty = ReadFloat(ReadLong(KO_PTR_CHR) + KO_OFF_Y)
x = Abs(crx - tx)
Y = Abs(cry - ty)
If tx > crx Then isrtx = -1: bykx = tx: kckx = crx Else isrtx = 1: bykx = crx: kckx = tx
If ty > cry Then isrty = -1: byky = ty: kcky = cry Else isrty = 1: byky = cry: kcky = ty
uzak = Int(Sqr((x ^ 2 + Y ^ 2)))
If uzak > 9999 Then Exit Function
If crx <= 0 Or cry <= 0 Then Exit Function
For i = zipla To uzak Step zipla
a = i ^ 2 * x ^ 2
b = x ^ 2 + Y ^ 2
d = Sqr(a / b)
e = Sqr(i ^ 2 - d ^ 2)
x1 = Int(tx + isrtx * d)
y1 = Int(ty + isrty * e)
If (kckx <= x1 And x1 <= bykx) And (kcky <= y1 And y1 <= byky) Then
WriteFloat ReadLong(KO_PTR_CHR) + KO_OFF_X, x1
WriteFloat ReadLong(KO_PTR_CHR) + KO_OFF_Y, y1
WriteFloat ReadLong(KO_PTR_CHR) + KO_OFF_Z, ReadFloat(ReadLong(KO_PTR_CHR) + KO_OFF_Z)
Paket "06" _
& Left$(AlignDWORD(ReadFloat(ReadLong(KO_PTR_CHR) + KO_OFF_X) * 10), 4) _
& Left$(AlignDWORD(ReadFloat(ReadLong(KO_PTR_CHR) + KO_OFF_Y) * 10), 4) _
& Left$(AlignDWORD(ReadFloat(ReadLong(KO_PTR_CHR) + KO_OFF_Z) * 10), 4) _
& "2D0000" _
& FormatHex(Hex$(CInt(CharX) * 10), 4) & FormatHex(Hex$(CInt(CharY) * 10), 4) & FormatHex(Hex$(CInt(CharZ) * 10), 4)
End If
Next
WriteFloat ReadLong(KO_PTR_CHR) + KO_OFF_X, crx
WriteFloat ReadLong(KO_PTR_CHR) + KO_OFF_Y, cry
WriteFloat ReadLong(KO_PTR_CHR) + KO_OFF_Z, ReadFloat(ReadLong(KO_PTR_CHR) + KO_OFF_Z)

Paket "06" _
& Left$(AlignDWORD(ReadFloat(ReadLong(KO_PTR_CHR) + KO_OFF_X) * 10), 4) _
& Left$(AlignDWORD(ReadFloat(ReadLong(KO_PTR_CHR) + KO_OFF_Y) * 10), 4) _
& Left$(AlignDWORD(ReadFloat(ReadLong(KO_PTR_CHR) + KO_OFF_Z) * 10), 4) _
& "2D0000" _
& FormatHex(Hex$(CInt(CharX) * 10), 4) & FormatHex(Hex$(CInt(CharY) * 10), 4) & FormatHex(Hex$(CInt(CharZ) * 10), 4)
Pause 0.1
End Function
Public Function hex2Val(pStrhex As String) As Long
Dim TmpStr As String
Dim Tmphex As String
Dim i As Long
TmpStr = ""
For i = Len(pStrhex$) To 1 Step -1
    Tmphex$ = Hex$(Asc(mID$(pStrhex$, i, 1)))
    If Len(Tmphex$) = 1 Then Tmphex$ = "0" & Tmphex$
    TmpStr = TmpStr & Tmphex$
Next
hex2Val = CLng("&H" & TmpStr)
End Function
Function ReadString2(ByVal pAddy As Long, ByVal LSize As Long) As String

    On Error Resume Next
    Dim value As Byte
    Dim tex() As Byte
    If LSize = 0 Then Exit Function

    ReDim tex(1 To LSize)
    ReadProcessMem KO_HANDLE, pAddy, tex(1), LSize, 0&
    ReadString2 = StrConv(tex, vbUnicode)

End Function
Function ReadString(ByVal pAddy As Long, ByVal OtoSize As Boolean, Optional ByVal LSize As Long = 1) As String
Dim value As Byte
Dim tex() As Byte
On Error Resume Next
If OtoSize = True Then
ReadProcessMem KO_HANDLE, pAddy, value, 1, 0&
LSize = value
ReDim tex(1 To LSize)
ReadProcessMem KO_HANDLE, pAddy, tex(1), LSize, 0&
ReadString = StrConv(tex, vbUnicode)
Else
If LSize = 0 Then
Exit Function
Else
ReDim tex(1 To LSize)
ReadProcessMem KO_HANDLE, pAddy, tex(1), LSize, 0&
ReadString = StrConv(tex, vbUnicode)
End If
End If
End Function

Public Function MouseX()
MouseX = ReadFloat(ReadLong(KO_PTR_CHR) + KO_OFF_MX)
End Function

Public Function MouseY()
MouseY = ReadFloat(ReadLong(KO_PTR_CHR) + KO_OFF_MY)
End Function


Function MobBilgi(TargetMob As Long)
Dim Ptr As Long, tmpMobBase As Long, tmpBase As Long, IDArray As Long, BaseAddr As Long, Mob As Long, zaman1 As Long
Mob = TargetMob
Ptr = ReadLong(KO_FLDB)
zaman1 = GetTickCount
tmpMobBase = ReadLong(Ptr + &H2C)
tmpBase = ReadLong(tmpMobBase + &H4)
While tmpBase <> 0
If zaman1 - GetTickCount > 50 Then Exit Function
IDArray = ReadLong(tmpBase + &HC)
If IDArray >= Mob Then
If IDArray = Mob Then
BaseAddr = ReadLong(tmpBase + &H10)
End If
tmpBase = ReadLong(tmpBase + &H0)
Else
tmpBase = ReadLong(tmpBase + &H8)
End If
Wend
MobBilgi = BaseAddr
End Function

Public Function etrafoku2() As String
Dim d1pkt As String
Dim i22 As Long, i11 As Long


For i11 = 10000 To 30000 '10000'den 30000'e kadar tüm moblarin listesini serverdan iste.
d1pkt = d1pkt + AlignDWORD(i11)
    If i22 > 200 Then
        d1pkt = "1D" + AlignDWORD(200) + d1pkt
        'List3.AddItem d1pkt
        Paket d1pkt 'paketat veya Sendpacket olabilir. Düzenleyebilirsiniz.
        Pause 0.1 ' burasi arttirilabilir hizli pakette oyun atmasin diye. Delay yerine Projenizde Pause vb. komutlar varsa kullanabilirsiniz. 1ms bekle yapildi.
        
        d1pkt = ""
        i22 = 0
    End If
i22 = i22 + 1
Next

Pause 5 '5 sn bekliyoruz, serverdan tümbilgiler bize ulassin.

Dim EBP As Long, ESI As Long, EAX As Long, FEnd As Long
Dim base_addr As Long
Dim class1 As String, class2 As String, lvl As String, x As String, Y As String, id As String, Name As String
EBP = ReadLong(ReadLong(KO_FLDB) + &H40)

Form1.List1.Clear 'form1'deki listeyi temizle.
FEnd = ReadLong(ReadLong(EBP + 4) + 4)
ESI = ReadLong(EBP)
While ESI <> EBP
base_addr = ReadLong(ESI + &H10)
If ReadLong(base_addr + &H698) > 15 Then
Name = ReadString(ReadLong(base_addr + KO_OFF_NAME), False, ReadLong(base_addr + &H698))
Else
Name = ReadString(base_addr + KO_OFF_NAME, False, ReadLong(base_addr + &H698))
End If

If Name = "Snake Queen" Or Name = "Talos" Then
MsgBox "Boss bulundu"
If (ReadLong(base_addr + KO_OFF_X) <> 0) Then
MsgBox "kordinatx: " & CStr(ReadLong(base_addr + KO_OFF_X)) & " ---   kordinaty: " & CStr(ReadLong(base_addr + KO_OFF_X))
Else
Paket "22" + AlignDWORD(ReadLong(base_addr + KO_OFF_ID)) + "00" 'burasi 00 veya 01 deneyebilirsiniz.
End If
End If

Form1.List1.AddItem Name
EAX = ReadLong(ESI + 8)
   If ReadLong(ESI + 8) <> ReadLong(KO_FLPZ) Then
       While ReadLong(EAX) <> ReadLong(KO_FLPZ)
       EAX = ReadLong(EAX)
       Wend
   ESI = EAX
   Else
   EAX = ReadLong(ESI + 4)
       While ESI = ReadLong(EAX + 8)
       ESI = EAX
       EAX = ReadLong(EAX + 4)
       Wend
       If ReadLong(ESI + 8) <> EAX Then
       ESI = EAX
       End If
End If
Wend
End Function
Public Function etrafoku3() As String
Dim d1pkt As String
Dim i22 As Long, i11 As Long


For i11 = 10000 To 30000 '10000'den 30000'e kadar tüm moblarin listesini serverdan iste.
d1pkt = d1pkt + AlignDWORD(i11)
    If i22 > 200 Then
        d1pkt = "1D" + AlignDWORD(200) + d1pkt
        'List3.AddItem d1pkt
        Paket d1pkt 'paketat veya Sendpacket olabilir. Düzenleyebilirsiniz.
        Pause 0.1 ' burasi arttirilabilir hizli pakette oyun atmasin diye. Delay yerine Projenizde Pause vb. komutlar varsa kullanabilirsiniz. 1ms bekle yapildi.
        
        d1pkt = ""
        i22 = 0
    End If
i22 = i22 + 1
Next

Pause 5 '5 sn bekliyoruz, serverdan tümbilgiler bize ulassin.

Dim EBP As Long, ESI As Long, EAX As Long, FEnd As Long
Dim base_addr As Long
Dim class1 As String, class2 As String, lvl As String, x As String, Y As String, id As String, Name As String
EBP = ReadLong(ReadLong(KO_FLDB) + &H40)

'Form1.PazarList.Refresh 'new with list
Form1.List1.Clear 'form1'deki listeyi temizle.
FEnd = ReadLong(ReadLong(EBP + 4) + 4)
ESI = ReadLong(EBP)
While ESI <> EBP
base_addr = ReadLong(ESI + &H10)
If ReadLong(base_addr + &H698) > 15 Then
Name = ReadString(ReadLong(base_addr + KO_OFF_NAME), False, ReadLong(base_addr + &H698)) '698=namelen
Else
Name = ReadString(base_addr + KO_OFF_NAME, False, ReadLong(base_addr + &H698))
End If

If Name = "Snake Queen" Or Name = "Talos" Then
MsgBox "Boss bulundu"
If (ReadLong(base_addr + KO_OFF_X) <> 0) Then
MsgBox "kordinatx: " & CStr(ReadLong(base_addr + KO_OFF_X)) & " ---   kordinaty: " & CStr(ReadLong(base_addr + KO_OFF_X))
Else
Paket "22" + AlignDWORD(ReadLong(base_addr + KO_OFF_ID)) + "00" 'burasi 00 veya 01 deneyebilirsiniz.
End If
End If

'Form1.List1.AddItem Name
If base_addr <> 0 And Form1.Text1.Text <> "" Then
If NewKordinatUzaklik(CharX, CharY) < 3 Then
'Form1.PazarList.ListItems.Add , , Name
'Form1.PazarList
Form1.List1.AddItem Name
'Form1.List1.ItemData(Form1.List1.NewIndex) = ReadLong(base_addr + KO_OFF_ID)



'Form1.PazarList.ListItems.Add , , Form1.List1.ItemData(0)
End If
End If
EAX = ReadLong(ESI + 8)
If ReadLong(ESI + 8) <> FEnd Then
While ReadLong(EAX) <> FEnd
EAX = ReadLong(EAX)
Wend
ESI = EAX
Else
EAX = ReadLong(ESI + 4)
While ESI = ReadLong(EAX + 8)
ESI = EAX
EAX = ReadLong(EAX + 4)
Wend
If ReadLong(ESI + 8) <> EAX Then
ESI = EAX
End If
End If
Wend

Dim i As Long
For i = 0 To Form1.List1.ListCount - 1
Form1.List1.Selected(i) = True
Form1.PazarList.ListItems.Add , , Form1.List1.Text
Next
Form1.List1.Selected(0) = True
Form1.tmroku.Enabled = False
End Function

Function OkuIDName(OkuNameID As Long)
If ReadLong(OkuNameID + KO_OFF_NAMELEN) >= 15 Then
OkuIDName = ReadString(ReadLong(OkuNameID + KO_OFF_NAME), False, ReadLong(OkuNameID + KO_OFF_NAMELEN))
Else
OkuIDName = ReadString(OkuNameID + KO_OFF_NAME, False, ReadLong(OkuNameID + KO_OFF_NAMELEN))
End If
End Function


Function NewKordinatUzaklik(Target_X As Long, Target_Y As Long) As Long
NewKordinatUzaklýk = Sqr((Target_X - CharX) ^ 2 + (Target_Y - CharY) ^ 2)
End Function

Function KordinatUzaklik(Target_X As Long, Target_Y As Long)
KordinatUzaklýk = Fix((((Target_X - ReadFloat(ReadLong(KO_PTR_CHR) + KO_OFF_X)) * (Target_X - ReadFloat(ReadLong(KO_PTR_CHR) + KO_OFF_X)) + (Target_Y - ReadFloat(ReadLong(KO_PTR_CHR) + KO_OFF_Y)) * (Target_Y - ReadFloat(ReadLong(KO_PTR_CHR) + KO_OFF_Y))) ^ 0.5) / 4)
End Function

Function KordinatArasiFark(X_Bir As Long, Y_Bir As Long, X_Iki As Long, Y_Iki As Long) As Long
KordinatArasýFark = Sqr((X_Bir - X_Iki) ^ 2 + (Y_Bir - Y_Iki) ^ 2)
End Function
Sub f_Sleep(pMS As Long, Optional pDoevents As Boolean = False)
Dim pTime As Long
pTime = GetTickCount
Do While pMS + pTime > GetTickCount
If pDoevents = True Then DoEvents
Loop
End Sub
Public Function okuamk() As String
Dim d1pkt As String
Dim i22 As Long, i11 As Long
Dim zaman1 As Long
zaman1 = GetTickCount

'For i11 = 10000 To 30000 '10000'den 30000'e kadar tüm moblarin listesini serverdan iste.
'd1pkt = d1pkt + AlignDWORD(i11)
'    If i22 > 200 Then
'        d1pkt = "1D" + AlignDWORD(200) + d1pkt
'        'List3.AddItem d1pkt
'        Paket d1pkt 'paketat veya Sendpacket olabilir. Düzenleyebilirsiniz.
'        Pause 0.1 ' burasi arttirilabilir hizli pakette oyun atmasin diye. Delay yerine Projenizde Pause vb. komutlar varsa kullanabilirsiniz. 1ms bekle yapildi.
'
'        d1pkt = ""
'        i22 = 0
'    End If
'i22 = i22 + 1
'Next

Pause 1 '5 sn bekliyoruz, serverdan tümbilgiler bize ulassin.

Dim i As Long, Base As Long, Name As String, UserX As Long, UserY As Long
Dim EBP As Long, ESI As Long, EAX As Long
Dim base_addr As Long

EBP = ReadLong(ReadLong(KO_FLDB) + &H40)
ESI = ReadLong(EBP)
Form1.List4.Clear
While ESI <> EBP
base_addr = ReadLong(ESI + &H10) ' base adresi burda aliyoruz
If base_addr = 0 Then Exit Function
If zaman1 - GetTickCount > 50 Then Exit Function

EAX = ReadLong(ESI + 8)
   If ReadLong(ESI + 8) <> ReadLong(KO_FLPZ) Then
       While ReadLong(EAX) <> ReadLong(KO_FLPZ)
       If zaman1 - GetTickCount > 50 Then Exit Function
       EAX = ReadLong(EAX)
       Wend
   ESI = EAX
   Else
   EAX = ReadLong(ESI + 4)
       While ESI = ReadLong(EAX + 8)
       If zaman1 - GetTickCount > 50 Then Exit Function
       ESI = EAX
       EAX = ReadLong(EAX + 4)
       Wend
       If ReadLong(ESI + 8) <> EAX Then
       ESI = EAX
       End If
   End If
   
OkuIDName (base_addr)
If zaman1 - GetTickCount > 50 Then Exit Function
If ReadLong(base_addr + KO_OFF_NAMELEN) > 15 Then
Name = ReadString(ReadLong(base_addr + KO_OFF_NAME), False, ReadLong(base_addr + KO_OFF_NAMELEN))
Else
Name = ReadString(base_addr + KO_OFF_NAME, False, ReadLong(base_addr + KO_OFF_NAMELEN))
End If
'Name = ReadString(ReadLong(Base + KO_OFF_NAME), ReadLong(Base + KO_OFF_NAME + 4))
UserX = ReadFloat(base_addr + KO_OFF_X)
UserY = ReadFloat(base_addr + KO_OFF_Y)
Form1.Text4.Text = Name


If base_addr <> 0 And Form1.Text4.Text <> "" Then
If NewKordinatUzaklik(UserX, UserY) < 3 Then
Form1.List4.AddItem Name
Form1.List4.ItemData(Form1.List4.NewIndex) = ReadLong(base_addr + KO_OFF_ID)
End If
End If
Wend
End Function
