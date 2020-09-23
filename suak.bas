Attribute VB_Name = "Module2"
'**************************************
' Name: Conversion between Dec, Bin and
'     Hex
' Description:This module contain functi
'     on that are used to convert between deci
'     mal, binary and hexadecimal.
' By: Pierre-Alain Vigeant
'
' Inputs:Depend on the function
'
' Returns:Depend on the function
'
' Assumes:Each function are 'stand-alone
'     '. This mean that u can copy one of them


'     without needing another one.
   ' The conversion Function are written In this way: <from>2<to>
'    Example: The Function 'Dec2Bin' will convert from Decimal To binary
'
'This code is copyrighted and has' limited warranties.Please see http://w
'     ww.Planet-Source-Code.com/xq/ASP/txtCode
'     Id.3242/lngWId.1/qx/vb/scripts/ShowCode.
'     htm'for details.'**************************************

'***************************************
'     ***********************
'* Best Tools *
'* Conversion *
'* v2.1 (Improved performance)*
'*for VB*
'**
'*This module contain a lot of subs and


'     functions for basic*
    '*conversion between Hexadecimal, Binary
    '     and decimal. *
    '***************************************
    '     ***********************
Global xerses



Public Function Bin2Dec(ByVal sBin As String) As Long
    Dim i As Integer


    For i = 1 To Len(sBin)
        Bin2Dec = Bin2Dec + CLng(CInt(Mid(sBin, Len(sBin) - i + 1, 1)) * 2 ^ (i - 1))
    Next i
End Function


Public Function Bin2Hex(ByVal sBin As String) As String
    Dim i As Integer
    Dim nDec As Long
    sBin = String(4 - Len(sBin) Mod 4, "0") & sBin 'Add zero To complete Byte


    For i = 1 To Len(sBin)
        nDec = nDec + CInt(Mid(sBin, Len(sBin) - i + 1, 1)) * 2 ^ (i - 1)
    Next i
    Bin2Hex = Hex(nDec)
    If Len(Bin2Hex) Mod 2 = 1 Then Bin2Hex = "0" & Bin2Hex
End Function


Public Function Dec2Bin(ByVal nDec As Integer) As String
    'This function is the same then Hex2Bin,
    '     but it has been copied to speed up proce
    '     ss
    Dim i As Integer
    Dim j As Integer
    Dim sHex As String
    Const HexChar As String = "0123456789ABCDEF"
    
    sHex = Hex(nDec) 'That the only part that is different


    For i = 1 To Len(sHex)
        nDec = InStr(1, HexChar, Mid(sHex, i, 1)) - 1


        For j = 3 To 0 Step -1
            Dec2Bin = Dec2Bin & nDec \ 2 ^ j
            nDec = nDec Mod 2 ^ j
        Next j
    Next i
    'Remove the first unused 0
    i = InStr(1, Dec2Bin, "1")
    If i <> 0 Then Dec2Bin = Mid(Dec2Bin, i)
End Function


Public Function Hex2Bin(ByVal sHex As String) As String
    Dim i As Integer
    Dim j As Integer
    Dim nDec As Long
    Const HexChar As String = "0123456789ABCDEF"
    


    For i = 1 To Len(sHex)
        nDec = InStr(1, HexChar, Mid(sHex, i, 1)) - 1


        For j = 3 To 0 Step -1
            Hex2Bin = Hex2Bin & nDec \ 2 ^ j
            nDec = nDec Mod 2 ^ j
        Next j
    Next i
    'Remove the first unused 0
    i = InStr(1, Hex2Bin, "1")
    If i <> 0 Then Hex2Bin = Mid(Hex2Bin, i)
End Function


 Function Hex2Dec(ByVal Hexi As String) As String
Hexi = UCase$(Hexi)

    Dim iX As Long, iMult As Long, iY As Long, iDig As String
    iMult = 1
    Hexi = Trim$(Hexi)
    If Len(Hexi) = 0 Then Hexi = "0"
    For iX = 1 To Len(Hexi)
        iDig = Mid$(Hexi, Len(Hexi) - iX + 1, 1)
        If InStr(1, "0123456789", iDig) Then
            iY = iY + iMult * CLng(iDig)
        Else
            iY = iY + iMult * CLng(Asc(iDig) - Asc("A") + 10)
        End If
        iMult = iMult * 16
    Next iX
    Hex2Dec = CStr(iY)
End Function


Public Function HiWord(ByVal DWord As Long) As Long
    HiWord = (DWord \ 65536) And &HFFFF
End Function


Public Function LoWord(ByVal DWord As Long) As Long
    LoWord = DWord And &HFFFF
End Function


Public Function DWord(ByVal HiWord As Long, ByVal LoWord As Long) As Long
    DWord = ((LoWord And 65536) Or ((HiWord And 65536) * 65536))
End Function
