Attribute VB_Name = "apiCRC32"
'==========================================================================
'  CRC32 CALCULATION / CHECKING THIRD-PARTY API MODULE
'=========================================================================


Option Explicit
Option Compare Text

Private crc32Table() As Long

Private Sub Class_initialize()

    ' This is the official polynomial used by CRC32 in PKZip.
    ' Often the polynomial is shown reversed (04C11DB7).
    Dim dwPolynomial As Long
    dwPolynomial = &HEDB88320
    Dim i As Integer, j As Integer

    ReDim crc32Table(256)
    Dim dwCrc As Long

    For i = 0 To 255
        dwCrc = i
        For j = 8 To 1 Step -1
            If (dwCrc And 1) Then
                dwCrc = ((dwCrc And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                dwCrc = dwCrc Xor dwPolynomial
            Else
                dwCrc = ((dwCrc And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
            End If
        Next j
        crc32Table(i) = dwCrc
    Next i

End Sub

   
Public Function GetCRCFromFile(FileName As String) As Long

 Dim cStream As New cBinaryFileStream
 Dim cCRC32 As New cCRC32
 Dim lCRC32 As Long
 
   cStream.File = FileName
   GetCRCFromFile = cCRC32.GetFileCrc32(cStream)
   
End Function


