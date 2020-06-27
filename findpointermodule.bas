Attribute VB_Name = "findpointermodule"

Public Function ReadLong2(addr As Long, Optional OyunHandle As Long) As Long 'read a 4 byte value
On Error Resume Next
    Dim value As Long
    If KO_HANDLE <> 0 Then
        ReadProcessMem OyunHandle, addr, value, 4, 0&
         Else
    ReadProcessMem KO_HANDLE, addr, value, 4, 0&
    End If
    ReadLong2 = value
End Function

Public Function FindPointer(Opcode As String, Begin As Long, Length As Long, Optional Handle As Long) As Long
FindPointer = ReadLong2(SearchHexArray(Opcode, Begin, Length, KO_HANDLE) + (Len(Opcode) / 2), KO_HANDLE)
DoEvents
End Function


Public Function SearchHexArray(HexArray As String, StartOffset As Long, Length As Long, Optional Handle As Long) As Long

On Error Resume Next
ReDim tmpBase(1 To Len(HexArray) / 2) As Byte
ReDim keyword(1 To Len(HexArray) / 2) As Byte
Dim tmpSearchHeap(&H1000) As Byte
Dim handle2 As Long

If Handle = 0 Then handle2 = KO_HANDLE Else handle2 = Handle

Dim i As Long
Dim j As Long
Dim k As Long
ConvHEX2ByteArray HexArray, keyword


For k = StartOffset To StartOffset + Length Step &H1000
ReadProcessMem handle2, k, tmpSearchHeap(0), &H1000, 0&
For i = 0 To &HFFF
    If tmpSearchHeap(i) = keyword(1) Then
        For j = 1 To UBound(keyword)
        
            If tmpSearchHeap(i + j - 1) <> keyword(j) Then GoTo fail
        Next
        SearchHexArray = k + i
        GoTo Fin
fail:
    End If
Next
Next
SearchHexArray = 0
Fin:
End Function


Public Function CallFinder(Opcode As String, Begin As Long, Length As Long) As Long
CallFinder = SearchHexArray(Opcode, Begin, Length) + (Len(Opcode) / 2)
DoEvents
End Function


Function HexFormatla(strHex As String, inLength As Integer)
On Error Resume Next
Dim newHex As String, byte1 As String, byte2 As String, byte3 As String, byte4 As String
Dim ZeroSpaces As Integer
'ABC,4
ZeroSpaces = inLength - Len(strHex) '1
newHex = String(ZeroSpaces, "0") + strHex '0ABC
byte1 = Left(newHex, 2)
byte2 = mID(newHex, 3, 2)
byte3 = mID(newHex, 5, 2)
byte4 = Right(newHex, 2)
Select Case Len(newHex)
Case 2 '0A
newHex = byte1
Case 4 '0ABC
newHex = byte4 & byte1
Case 6 '000ABC
newHex = byte4 & byte2 & byte1
Case 8 '00000ABC
newHex = byte4 & byte3 & byte2 & byte1
Case Else
End Select
HexFormatla = newHex
'\\
End Function


