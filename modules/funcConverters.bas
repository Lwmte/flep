Attribute VB_Name = "funcConverters"
'==========================================================================
'  FLEP 1.0 CONVERSION FUNCTIONS
'
'  Main code by Pyuaumch. Refactoring and rewrites by Lwmte.
'
'  Contains various string-to-hex, string-to-decimal, etc. conversion
'  functions that are mainly used to work with user input in forms,
'  textboxes, etc.
'
'==========================================================================
'

Option Explicit

 
Public Type typFloat         ' FLOAT CONVERTER TYPES/VALUES
    F As Single
End Type

Public Type typStringArray2  ' STRING ARRAY CONVERTER TYPE
    Str(1 To 2) As String
End Type

Public Type typByteArray3    ' BYTE ARRAY CONVERTER TYPES/VALUES
    B(1 To 3) As Byte
End Type

Public Type typByteArray4    ' BYTE ARRAY CONVERTER TYPES/VALUES
    B(1 To 4) As Byte
End Type


Public Type typNumString     ' NUM/STRING PARAMETER TYPE
    Number As Integer
    RawString As String
End Type


Public Type typCfgParam      ' CFG PARAMETER TYPE
    Name As String
    Value As String
    Comment As String
End Type


Public MarkError As Boolean  ' Global error conversion flag.

'==========================================================================
' FUNCTION:
'
'==========================================================================




'==========================================================================
' FUNCTION: WRITE PARAMETER
' Converts patch param type to HEX string, according to given type index,
' then writes it to file.
' P.S.: Given type index is identical to NewMainWindow parameter type
' option control array.
'==========================================================================
Function WriteParam(ByVal RawParam As String, ByVal Offset As String, DataType As Integer) As Boolean

On Error GoTo ErrorHandler
 
 Dim cntParamOffsetRGB As Integer  ' Next three variables needed only for RGB type.
 Dim RGBValues() As String
 Dim RGBOffsets() As String
 
 Dim FinalHexString As String
 
 
  WriteParam = False ' reset just in case...
  
 
   Select Case DataType
    
        Case 0: FinalHexString = InvertHex(ValToHex(DecToIEEE(CDbl(StripIn(RawParam, kMaskFloat))), 8))
        Case 1: FinalHexString = ValToHex(RawParam, 2)                  ' Bits(8)
        Case 2: FinalHexString = InvertHex(ValToHex(RawParam, 4))       ' Bits(16)
        Case 3, 5: FinalHexString = InvertHex(ValToHex(RawParam, 2))    ' Byte (signed / unsigned)
        Case 4, 6: FinalHexString = InvertHex(ValToHex(RawParam, 4))    ' Integer (signed / unsigned)
        Case 7: FinalHexString = InvertHex(ValToHex(RawParam, 8))       ' Long
        Case 9:  FinalHexString = InvertHex(BytesToHex(RawParam, 3))    ' RGB
                 RGBOffsets = Split(Offset, kDivider2, 3)
                 RGBValues = Split(RawParam, kDivider, 3)
                 
        Case 8: FinalHexString = vbNullString                           ' String with zero length, never used.
        Case Else
        
             If DataType < 100 Then
                FinalHexString = vbNullString
            Else
            
                ' For string type (which is always > 100), we calculate length by dividing DataType by 100,
                ' multiplying it by 2 (as hex takes 2) and adding 2 extra zeros to the end.
                
                FinalHexString = CharFillR((StringToHex(RawParam)), "0", ((Fix(DataType / 100) * 2))) & "00"
                
            End If
            
    End Select
    
    
    
    If DataType <> 9 Then  ' For RGB datatype, we specify offset workaround,
                           ' in case user wants to specify offset for each color component seperately.
    
        Call WriteHex(FinalHexString, HxVal(Offset), hFile)    ' Default method, single offset.
        
    Else
            Select Case UBound(RGBOffsets)                     ' Alternate method, single or triple offsets (twin gets ignored).
            
                Case 0: Call WriteHex(FinalHexString, HxVal(Offset), hFile)
                Case 2:
                        For cntParamOffsetRGB = 0 To 2
                            Call WriteHex(ValToHex(RGBValues(cntParamOffsetRGB), 2), HxVal(RGBOffsets(cntParamOffsetRGB)), hFile)
                        Next cntParamOffsetRGB
                        
                Case Else: GoTo ErrorHandler
            
            End Select
            
    End If
    
    
 WriteParam = True
 Exit Function
 

ErrorHandler:
    Exit Function

End Function



'==========================================================================
' FUNCTION: CONVERT CONFIG STRING
' Deciphers config string by mask [ParName] = [ParString] and returns
' result as cfgParam type.
'==========================================================================
'
Public Function ConvCFG(ByVal SourceString As String) As typCfgParam

 Dim cntCharCounter          As Long
 Dim cntSrcStringLength      As Long

 Dim cntMarkCommentBeginning As Long
 Dim cntMarkValueBeginning   As Long
               
    SourceString = Trim$(SourceString)

    If LenB(SourceString) = 0 Then Exit Function
    If Asc(SourceString) = 59 Or Asc(SourceString) = 91 Then Exit Function 'if REMARKED, then END FUNCTION NOW!!!
                            
    ConvCFG.Name = vbNullString
    ConvCFG.Value = vbNullString
    ConvCFG.Comment = vbNullString
                
    cntMarkCommentBeginning = 0
    cntMarkValueBeginning = 0
                
    cntSrcStringLength = Len(SourceString)

    For cntCharCounter = cntSrcStringLength To 1 Step -1
                
        Select Case Mid$(SourceString, cntCharCounter, 1)
                    
            Case kCommentary: cntMarkCommentBeginning = cntCharCounter + 1

            Case kEquals: cntMarkValueBeginning = cntCharCounter + 1
                   
        End Select
                
    Next cntCharCounter

                
    If cntMarkValueBeginning = 0 Then Exit Function
    If cntMarkValueBeginning > cntMarkCommentBeginning And cntMarkCommentBeginning > 0 Then Exit Function
                
    ConvCFG.Name = Trim$(Left$(SourceString, cntMarkValueBeginning - 2))
                
    If cntMarkCommentBeginning = 0 Then
                
        ConvCFG.Value = Trim$(Right$(SourceString, (cntSrcStringLength + 1) - cntMarkValueBeginning))
                    
    Else
                
        ConvCFG.Comment = Trim$(Mid$(SourceString, cntMarkCommentBeginning))
        ConvCFG.Value = Trim$(Mid$(SourceString, cntMarkValueBeginning, cntMarkCommentBeginning - cntMarkValueBeginning - 1))
                    
    End If
                
End Function



'==========================================================================
' FUNCTION: VALUE TO HEX-STRING OF SPECIFIED LENGTH
' Converts decimal value (e.g. "11") into true hex value with given length
' (e.g. "0B" in case nativelength=1 or "000B in case nativelength=2)
'==========================================================================
Public Function ValToHex(ByVal SourceValue As String, ByVal DesiredLength As Byte) As String

On Error GoTo ErrorHandler

 Dim SrcLength As Byte
    
    ValToHex = Hex(Val(SourceValue))
    SrcLength = Len(ValToHex)
    
    If SrcLength < DesiredLength Then ValToHex = CharFillL(ValToHex, "0", DesiredLength)
        
    If SrcLength > DesiredLength Then _
       ValToHex = Mid$(ValToHex, (SrcLength - DesiredLength + 1), DesiredLength) 'cuts off excess
    
    Exit Function

ErrorHandler:
    MsgBox "Warning: possible error during DEC > HEX conversion. You have entered incorrect value (" + SourceValue + ")."
    ValToHex = vbNullString
End Function



'==========================================================================
' FUNCTION: VALUE TO HEX-STRING OF SPECIFIED LENGTH (UNSIGNED)
' This function does the same as ValToHex, but with unsigned hexes
'==========================================================================
Public Function ValToHexUnsigned(ByVal SourceValue As String, ByVal DesiredLength As Byte) As String
On Error GoTo ErrorHandler

 Dim SrcLength As Byte
    
    ValToHexUnsigned = UnsignedHex(Val(SourceValue))
    SrcLength = Len(ValToHexUnsigned)

    If SrcLength < DesiredLength Then ValToHexUnsigned = CharFillL(ValToHexUnsigned, "0", DesiredLength)
    
    If SrcLength > DesiredLength Then _
       ValToHexUnsigned = Mid$(ValToHexUnsigned, (SrcLength - DesiredLength + 1), DesiredLength) 'cuts off excess
    
    Exit Function

ErrorHandler:
    MsgBox "Warning: possible error during DEC>HEX conversion. You have entered incorrect value (" + SourceValue + ")."
    ValToHexUnsigned = vbNullString

End Function



'==========================================================================
' FUNCTION: INVERT HEXADECIMAL STRING (ex-Invrt)
' Inverts hexadecimal string to comply with x86 little-endian standard.
'==========================================================================
Public Function InvertHex(ByVal SourceString As String) As String

 Dim cntCurChar As Integer
 Dim LengthInBytes As Integer

'   Check if string contains odd or even amount of symbols, and if it's even,
'   just cut the last symbol:

    If Len(SourceString) Mod 2 = 0 Then _
       LengthInBytes = Len(SourceString) / 2 Else _
       LengthInBytes = Len(SourceString) / 2 - 1
       
       
'   Inversion cycle itself:

    For cntCurChar = 1 To LengthInBytes * 2 Step 2
    
        If cntCurChar <> LengthInBytes * 2 Then
            InvertHex = InvertHex + (Mid$(SourceString, ((LengthInBytes * 2) - cntCurChar), 2))
        End If
        
    Next

End Function



'==========================================================================
' FUNCTION: DECIMAL TO UNSIGNED HEX CONVERSION
' Converts any type of numbers to unsigned HEX string (prevents overflow)
'==========================================================================
Function UnsignedHex(ByVal Value As Variant) As String

 Dim TwoToThe32 As Variant
    
        TwoToThe32 = CDec("2") ^ 32
        
        If CDec(Value) < 0 Or Abs(CDec(Value)) >= TwoToThe32 Then
            UnsignedHex = -1
        Else
            If CDec(Value) >= TwoToThe32 / 2 Then
                Value = CDec(Value) - TwoToThe32
            End If
            UnsignedHex = Hex$(CDec(Value))
        End If
        
End Function



'==========================================================================
' FUNCTION: A,B,C,D PARAMETERS TO BYTES(4)
' Converts 4 divider-separated byte values string into 4 byte array values
'==========================================================================
Public Function ParamsToBytes4(RawString As String, ByVal Nomer As Byte) As typByteArray4

On Error GoTo ErrorHandler 'if overflow or end string, then stop execution

 Dim tmpStringArray() As String
 Dim tmpCurrentValue As Byte
 Dim cntPointer As Byte

        tmpStringArray = Split(RawString, kDivider, 4)
        If UBound(tmpStringArray) > 3 Then ReDim Preserve tmpStringArray(3)
        
        For cntPointer = 0 To UBound(tmpStringArray)
            ParamsToBytes4.B(cntPointer + 1) = CByteL(tmpStringArray(cntPointer))
        Next cntPointer
                
        Exit Function
        
ErrorHandler:   ParamsToBytes4.B(1) = 0 'fuk em...
                ParamsToBytes4.B(2) = 0
                ParamsToBytes4.B(3) = 0
                ParamsToBytes4.B(4) = 0

End Function



'==========================================================================
' FUNCTION: A,B,C PARAMETERS TO BYTES(3) (ex-RGBAConv)
' Converts 3 divider-separated byte values string into 3 byte array values
'==========================================================================
Public Function BytesToHex(RawString As String, Limit As Integer) As String

 Dim tmpStringArray() As String
 Dim cntPointer As Byte

        tmpStringArray = Split(RawString, kDivider, Limit)
        
        For cntPointer = 0 To UBound(tmpStringArray)
            BytesToHex = BytesToHex & ValToHex(tmpStringArray(cntPointer), 2)
        Next cntPointer

        
End Function



'==========================================================================
' FUNCTION: PARAMETERS TO STRING ARRAY
' Converts 2 divider-separated values into string + string values
'==========================================================================
Public Function ParamsToStringArray(RawString As String, Limit As Integer) As String()

On Error GoTo ErrorHandler 'if overflow or end string, then stop execution

 Dim cntPointer As Integer
 Dim tmpStringArray() As String

    ParamsToStringArray = Split(RawString, kDivider, Limit)
    If UBound(ParamsToStringArray) > Limit Or UBound(ParamsToStringArray) < Limit Then ReDim Preserve ParamsToStringArray(Limit)
    
    Exit Function
        
ErrorHandler:
    Exit Function
        
End Function



'==========================================================================
' FUNCTION: A,B PARAMETERS TO INTEGER + STRING
' Converts 2 divider-separated values into integer + string values
'==========================================================================
Public Function ParamsToNumString(RawString As String) As typNumString

On Error GoTo ErrorHandler 'if overflow or end string, then stop execution

 Dim tmpStringArray() As String

    tmpStringArray = Split(RawString, kDivider, 2)
    
    ParamsToNumString.Number = CInt(tmpStringArray(0))
    ParamsToNumString.RawString = tmpStringArray(1)
                                    
    Exit Function
        
ErrorHandler:   ParamsToNumString.Number = 0 'fuk em...
                ParamsToNumString.RawString = vbNullString

End Function



'==========================================================================
' FUNCTION: STRING TO HEXADECIMAL STRING
' Converts standard string to a string hexcode.
'==========================================================================
Public Function StringToHex(ByVal Stroka As String) As String

 Dim cntCharCounter As Byte
 
        For cntCharCounter = 1 To Len(Stroka)
            StringToHex = StringToHex & Hex(AscB(Mid$(Stroka, cntCharCounter, 1)))
        Next
        
End Function



'==========================================================================
' FUNCTION: DECIMAL TO IEEE-754 FLOAT
' Converts decimal long to IEEE-754 float
'==========================================================================
Public Function DecToIEEE(ByVal DecValue As Double) As Long

On Error GoTo ErrorHandler

 Dim B As typByteArray4
 Dim F As typFloat
 Dim t As Long
 
    F.F = DecValue
    LSet B = F
    DecToIEEE = B.B(4) * (2 ^ 24)
    DecToIEEE = DecToIEEE + B.B(3) * (2 ^ 16)
    DecToIEEE = DecToIEEE + B.B(2) * (2 ^ 8)
    DecToIEEE = DecToIEEE + B.B(1)

Exit Function

ErrorHandler:
        MsgBox "Error during DEC > IEEE-754 float conversion. Check if you have set correct value."

End Function



'==========================================================================
' FUNCTION: HEX TO DECIMAL VALUE
' Converts hexadecimal long to a decimal long.
'==========================================================================
Function HxVal(ByVal s As String) As Long

On Error GoTo ErrorHandler

    If LenB(s) <> 0 Then HxVal = CLng("&H" & s) Else HxVal = CLng("&H" & "00")
    Exit Function

ErrorHandler:
    If MarkError = False Then
        MarkError = True
        HxVal = CLng("&H" & "00")
        MsgBox "There was an error when converting some hexadecimal value to a decimal." & vbCrLf & _
               "Make sure that you haven't entered wrong data." & vbCrLf & "Source string: ''" & s & "''"
    End If

End Function

'==========================================================================
' FUNCTION: SINGLE-LINE TO MULTI-LINE (//-TERMINATED)
' Converts single-line //-terminated string into multiline string
'==========================================================================
Function DecipherText(ByVal Origtext As String) As String

    DecipherText = Replace$(Origtext, kTerminator, vbCrLf)

End Function

'==========================================================================
' FUNCTION: MULTI-LINE TO SINGLE-LINE (//-TERMINATED)
' Converts multi-line //-terminated string into single-line string
'==========================================================================
Function CipherText(ByVal SourceString As String) As String

    CipherText = Replace$(SourceString, vbCrLf, kTerminator)

End Function



'==========================================================================
' FUNCTION: PADDING WITH ZEROS FROM LEFT (ex-ZeroFill)
' Padding (char-fill) to the left side of source string with 0 symbol.
'==========================================================================
Function ZeroFill(ByVal Src As String, ByVal DesiredLength As Long) As String

    If Len(Src) > DesiredLength Then Exit Function
    
    ZeroFill = Src
    
    Do Until Len(ZeroFill) = DesiredLength
        ZeroFill = "0" & ZeroFill
    Loop

End Function


'==========================================================================
' FUNCTION: FILL
'
'==========================================================================
Function Fill(ByVal Src As String, ByVal DesiredLength As Long) As String

Dim cnt As Long

    For cnt = 0 To DesiredLength - 1
        Fill = Fill & Src
    Next cnt

End Function



'==========================================================================
' FUNCTION: PADDING (ADD SYMBOLS TO THE LEFT SIDE)
' Padding (char-fill) to the left side of source string.
'==========================================================================
Function CharFillL(ByVal Src As String, ByVal FillChar As String, ByVal DesiredLength As Long) As String

    If Len(Src) > DesiredLength Then CharFillL = Left$(Src, DesiredLength):  Exit Function
    If Len(FillChar) > 1 Then FillChar = Left$(FillChar, 1)
    
    CharFillL = Src
    
    Do Until Len(CharFillL) = DesiredLength
        CharFillL = FillChar & CharFillL
    Loop

End Function



'==========================================================================
' FUNCTION: PADDING (ADD SYMBOLS TO THE RIGHT SIDE)
' Padding (char-fill) to the right side of source string.
'==========================================================================
Function CharFillR(ByVal Src As String, ByVal FillChar As String, ByVal DesiredLength As Long) As String

    If Len(Src) > DesiredLength Then CharFillR = Left$(Src, DesiredLength): Exit Function
    If Len(FillChar) > 1 Then FillChar = Left$(FillChar, 1)
    
    CharFillR = Src
    
    Do Until Len(CharFillR) = DesiredLength
        CharFillR = CharFillR & FillChar
    Loop

End Function



'==========================================================================
' FUNCTION: CUT OFF
' This function cuts off specific amount of symbols from left
'==========================================================================
Function CutOff(ByVal SourceText As String, Length As Byte)

    If Len(SourceText) > Length Then
        CutOff = Mid$(SourceText, Length + 1)
    Else
        CutOff = SourceText
    End If

End Function



'==========================================================================
' FUNCTION: TRUE LENGTH OF STRING WITHOUT "/" SLASH SYMBOLS
'
'==========================================================================
Public Function TrueLOF(SourceString As String) As Integer 'returns true LOF without slashes

 TrueLOF = Len(Replace$(SourceString, "/", vbNullString))
    
End Function



'==========================================================================
' FUNCTION: MERGE ALL MODDED VALUES OF ALL PARAMETERS OF SELECTED PATCH.
' Used to collect all modified param. values for preset / config writing.
'==========================================================================
Public Function MergeModdedValues(PatchNumber As Integer) As String

On Error GoTo ErrorHandler

 Dim tmpStringArray() As String
 Dim cntUnitCounter As Integer

    ReDim tmpStringArray(UBound(PatchArray(PatchNumber).patchParams))
    
    For cntUnitCounter = LBound(PatchArray(PatchNumber).patchParams) To UBound(PatchArray(PatchNumber).patchParams)
        tmpStringArray(cntUnitCounter) = PatchArray(PatchNumber).patchParams(cntUnitCounter).parModdedValue
    Next cntUnitCounter
    
    MergeModdedValues = Join(tmpStringArray, kDivider2)
    
    Exit Function
    
ErrorHandler:
    MergeModdedValues = vbNullString

End Function



'==========================================================================
' FUNCTION: STRIPOUT
' Deletes specific symbols from string.
'==========================================================================
Public Function StripOut(SourceString As String, SymbolsToKill As String) As String

 Dim i As Integer
 
    StripOut = SourceString
    
    For i = 1 To Len(SymbolsToKill)
        StripOut = Replace(StripOut, Mid$(SymbolsToKill, i, 1), vbNullString)
    Next i
 
End Function



'==========================================================================
' FUNCTION: STRIPOUT
' Leaves only specified symbols in a string.
'==========================================================================
Public Function StripIn(SourceString As String, SymbolsToLeave As String) As String

 Dim i, i2 As Integer
 Dim c, s As String
 Dim t As String
 
    StripIn = vbNullString
    t = vbNullString
    
    
    For i = 1 To Len(SourceString)
        For i2 = 1 To Len(SymbolsToLeave)
            c = Mid$(SymbolsToLeave, i2, 1)
            s = Mid$(SourceString, i, 1)
            If s = c Then t = t & c
        Next i2
    Next i
    
    StripIn = t
 
End Function


'==========================================================================
' FUNCTION: FINALIZE
' Finalizes string with desired character, only if there is no such present
'==========================================================================
Public Function Finalize(SourceString As String, EndChar As String) As String

If UCase$(Right$(SourceString, 1)) <> UCase$(Left$(EndChar, 1)) Then Finalize = Finalize & Left$(EndChar, 1) Else Finalize = SourceString

End Function


'==========================================================================
' FUNCTION: CONVERT TO BYTE WITH OVERFLOW PREVENTION
'==========================================================================
Public Function CByteL(ByVal Value As Long) As Byte
    If Value > 255 Then CByteL = 255: Exit Function
    CByteL = CByte(Value)
End Function


'==========================================================================
' FUNCTION: CONVERT TO INTEGER WITH OVERFLOW PREVENTION
'==========================================================================
Public Function CIntL(ByVal Value As Long) As Integer
    If Value > 32767 Then CIntL = CInt(Value - 65536): Exit Function
    CIntL = CInt(Value)
End Function



'==========================================================================
' FUNCTION: BIN-2-DEC
' Converts binary string (e.g. 01010101) into decimal (e.g. 85)
'==========================================================================
Public Function Bin2Dec(Num As String) As Long
  Dim n As Long
  Dim a As Long
  Dim x As String
     n = Len(Num) - 1
     a = n
     Do While n > -1
        x = Mid(Num, ((a + 1) - n), 1)
        Bin2Dec = IIf((x = "1"), Bin2Dec + (2 ^ (n)), Bin2Dec)
        n = n - 1
     Loop
End Function


'==========================================================================
' FUNCTION: DEC-2-BIN 8
' Converts decimal byte into 8 bits as string.
'==========================================================================
Public Function Dec2Bin8(ByVal DecVal As Byte) As String
    Dim i As Integer
    Dim sResult As String

    sResult = Space(8)
    For i = 0 To 7
        If DecVal And (2 ^ i) Then
            Mid(sResult, 8 - i, 1) = "1"
        Else
            Mid(sResult, 8 - i, 1) = "0"
        End If
    Next
    Dec2Bin8 = sResult
End Function


'==========================================================================
' FUNCTION: DEC-2-BIN 16
' Converts decimal byte into 16 bits as string.
'==========================================================================
Public Function Dec2Bin16(ByVal DecVal As Integer) As String
    Dim i As Integer
    Dim sResult As String

    sResult = Space(16)
    For i = 0 To 15
        If DecVal And (2 ^ i) Then
            Mid(sResult, 16 - i, 1) = "1"
        Else
            Mid(sResult, 16 - i, 1) = "0"
        End If
    Next
    Dec2Bin16 = sResult
End Function
