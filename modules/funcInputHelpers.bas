Attribute VB_Name = "funcInputHelpers"
'==========================================================================
'  FLEP 1.0 INPUT HELPERS MODULE
'
'  Main code by Pyuaumch. Refactoring and rewrites by Lwmte.
'
'  Contains all code related to input helping (textbox input filtering, etc.).
'=========================================================================


Option Explicit


'==========================================================================
' FUNCTION: SELECT ALL
' Adds "select all" function to VB textboxes.
'==========================================================================
Public Sub SelectAll(ByRef TextBox As Object, KeyCode, Shift)

    If KeyCode = 65 And Shift = 2 Then
        TextBox.SelStart = 0
        TextBox.SelLength = Len(TextBox.Text)
       
        TextBox.Locked = True
        DoEvents
        TextBox.Locked = False
    End If

End Sub


'==========================================================================
' FUNCTION: CURRENT LINE
' Function used to identify current line in textbox
'==========================================================================
Public Function CurrentLine(ByRef txtBox As TextBox) As Long

    On Error Resume Next
    CurrentLine = SendMessage(txtBox.hwnd, EM_LINEFROMCHAR, -1&, ByVal 0&) + 1
    On Error GoTo 0
    
End Function


'==========================================================================
' FUNCTION: KeyFilterDIScanCode
' Function used to translate key symbol into DirectInput scancode
'==========================================================================
Function KeyFilterDIScanCode(ByVal Letter As String) As String

    If Len(Letter) > 1 Then Letter = Mid$(Letter, 1, 1) 'cutoff garbage symbols

    Select Case UCase(Letter)
        Case "Q":  KeyFilterDIScanCode = 0
        Case "W":  KeyFilterDIScanCode = 1
        Case "E":  KeyFilterDIScanCode = 2
        Case "R":  KeyFilterDIScanCode = 3
        Case "T":  KeyFilterDIScanCode = 4
        Case "Y":  KeyFilterDIScanCode = 5
        Case "U":  KeyFilterDIScanCode = 6
        Case "I":  KeyFilterDIScanCode = 7
        Case "O":  KeyFilterDIScanCode = 8
        Case "P":  KeyFilterDIScanCode = 9
        Case "[": KeyFilterDIScanCode = 10
        Case "]": KeyFilterDIScanCode = 11
        Case "A": KeyFilterDIScanCode = 14
        Case "S": KeyFilterDIScanCode = 15
        Case "D": KeyFilterDIScanCode = 16
        Case "F": KeyFilterDIScanCode = 17
        Case "G": KeyFilterDIScanCode = 18
        Case "H": KeyFilterDIScanCode = 19
        Case "J": KeyFilterDIScanCode = 20
        Case "K": KeyFilterDIScanCode = 21
        Case "L": KeyFilterDIScanCode = 22
        Case ";": KeyFilterDIScanCode = 23
        Case "Z": KeyFilterDIScanCode = 28
        Case "X": KeyFilterDIScanCode = 29
        Case "C": KeyFilterDIScanCode = 30
        Case "V": KeyFilterDIScanCode = 31
        Case "B": KeyFilterDIScanCode = 32
        Case "N": KeyFilterDIScanCode = 33
        Case "M": KeyFilterDIScanCode = 34
    End Select

End Function



'==========================================================================
' FUNCTION: KeyFilterHex
' Filters out non-hexadecimal ASCII keys.
'==========================================================================
Public Function KeyFilterHex(ByVal KeyASCII As Integer) As Integer

    KeyASCII = Asc(UCase(Chr(KeyASCII))) 'converts ASCII key to uppercase anyway.
    
    If KeyASCII <> 8 And KeyASCII < 48 Or _
       KeyASCII > 57 And KeyASCII < 65 Or _
       KeyASCII > 70 Then KeyFilterHex = 0 _
                     Else KeyFilterHex = KeyASCII

End Function



'==========================================================================
' FUNCTION: KeyFilterNumUnsigned
' Filters out non-digital ASCII keys (also filters out minus key).
'==========================================================================
Public Function KeyFilterNumUnsigned(ByVal KeyASCII As Integer) As Integer

    If KeyASCII <> 8 Then
        If KeyASCII < 48 Or KeyASCII > 57 Then KeyFilterNumUnsigned = 0 Else KeyFilterNumUnsigned = KeyASCII
    Else
        KeyFilterNumUnsigned = KeyASCII
    End If

End Function



'==========================================================================
' FUNCTION: KeyFilterNumSigned
' Filters out non-digital ASCII keys (allows minus key).
'==========================================================================
Public Function KeyFilterNumSigned(ByVal KeyASCII As Integer) As Integer

    If KeyASCII <> 8 And KeyASCII <> 45 Then
        If KeyASCII < 48 Or KeyASCII > 57 Then KeyFilterNumSigned = 0 Else KeyFilterNumSigned = KeyASCII
    Else
        KeyFilterNumSigned = KeyASCII
    End If

End Function



'==========================================================================
' FUNCTION: KeyFilterLatin
' Filters out non-latin letters and symbols ASCII keys.
'==========================================================================
Public Function KeyFilterLatin(ByVal KeyASCII As Integer) As Integer

    KeyASCII = Asc(UCase(ChrW(KeyASCII)))

    If KeyASCII > 64 And KeyASCII < 91 Then
        KeyFilterLatin = KeyASCII
    Else
        If KeyASCII <> 8 Then KeyFilterLatin = 0 Else KeyFilterLatin = KeyASCII
    End If

End Function



'==========================================================================
' FUNCTION: KeyFilterFloat
' Filters out non-digital and non-point ASCII keys (also filters out minus).
'==========================================================================
Public Function KeyFilterFloat(ByVal KeyASCII As Integer) As Integer

    If KeyASCII <> 8 And KeyASCII <> 46 Then
       If KeyASCII < 48 Or KeyASCII > 57 Then KeyFilterFloat = 0 Else KeyFilterFloat = KeyASCII
    Else
       KeyFilterFloat = KeyASCII
    End If

End Function



'==========================================================================
' FUNCTION: KeyFilterHexComma
' Filters out non-hex and non-comma ASCII keys (also filters out minus).
'==========================================================================
Public Function KeyFilterHexComma(ByVal KeyASCII As Integer) As Integer


    KeyASCII = Asc(UCase(Chr(KeyASCII))) 'converts ASCII key to uppercase anyway.
    
    If KeyASCII = 44 Then KeyFilterHexComma = KeyASCII: Exit Function
    
    If KeyASCII <> 8 And KeyASCII < 48 Or _
       KeyASCII > 57 And KeyASCII < 65 Or _
       KeyASCII > 70 Then KeyFilterHexComma = 0 _
                     Else: KeyFilterHexComma = KeyASCII
                     
End Function



'==========================================================================
' FUNCTIONS: FILTER MIN/MAX VALUES
' Filters out minimum and maximum values.
'==========================================================================
Public Function FilterMaxValue(ByVal SrcValue As String, ByVal MaxValue As Integer) As Integer
    
    FilterMaxValue = CInt(Val(SrcValue))
    If FilterMaxValue > MaxValue Then FilterMaxValue = MaxValue
    
End Function

Public Function FilterMinValue(ByVal SrcValue As String, ByVal MinValue As Integer) As Integer
    
    FilterMinValue = CInt(Val(SrcValue))
    If FilterMinValue < MinValue Then FilterMinValue = MinValue
    
End Function
'==========================================================================
