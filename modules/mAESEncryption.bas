Attribute VB_Name = "mAESEncryption"
Option Explicit

Public Enum EncrytionStrength
    enum128Bit = 128
    enum192Bit = 198
    enum256Bit = 256
End Enum

Private m_Rijndael As New cRijndael

Public Function AESEncrypt(ByVal data As String, ByVal password As String, Optional strength As EncrytionStrength) As String
   If strength = 0 Then strength = enum128Bit
   
   Dim retVal As String

    Dim pass()        As Byte
    Dim plaintext()   As Byte
    Dim ciphertext()  As Byte
    Dim KeyBits       As Long
    Dim BlockBits     As Long

    If Len(data) = 0 Then
        retVal = "No Plaintext"
    Else
        If Len(password) = 0 Then
            retVal = "No Password"
        Else
            KeyBits = strength
            BlockBits = strength
            pass = GetPassword(password)

            plaintext = StrConv(data, vbFromUnicode)

            m_Rijndael.SetCipherKey pass, KeyBits
            m_Rijndael.ArrayEncrypt plaintext, ciphertext, 0


            retVal = HexDisplay(ciphertext, UBound(ciphertext) + 1, BlockBits \ 8)

        End If
    End If


    AESEncrypt = retVal
End Function

Public Function AESDencrypt(ByVal data As String, ByVal password As String, Optional strength As EncrytionStrength) As String
    If strength = 0 Then strength = enum128Bit
   
    Dim retVal As String
   
    Dim pass()        As Byte
    Dim plaintext()   As Byte
    Dim ciphertext()  As Byte
    Dim KeyBits       As Long
    Dim BlockBits     As Long

    If Len(data) = 0 Then
        retVal = "No Ciphertext"
    Else
        If Len(password) = 0 Then
            retVal = "No Password"
        Else
            KeyBits = strength
            BlockBits = strength
            pass = GetPassword(password)

            If HexDisplayRev(data, ciphertext) = 0 Then
                retVal = "Text not Hex data"
                Exit Function
            End If

            m_Rijndael.SetCipherKey pass, KeyBits
            If m_Rijndael.ArrayDecrypt(plaintext, ciphertext, 0) <> 0 Then
                Exit Function
            End If

            retVal = StrConv(plaintext, vbUnicode)
            
        End If
    End If
    
    AESDencrypt = retVal
End Function

Private Function GetPassword(ByVal password As String) As Byte()
    Dim data() As Byte

    data = StrConv(password, vbFromUnicode)
    ReDim Preserve data(31)
        
    GetPassword = data
End Function

'Returns a String containing Hex values of data(0 ... n-1) in groups of k
Private Function HexDisplay(data() As Byte, n As Long, k As Long) As String
    Dim i As Long
    Dim j As Long
    Dim c As Long
    Dim data2() As Byte

    If LBound(data) = 0 Then
        ReDim data2(n * 4 - 1 + ((n - 1) \ k) * 4)
        j = 0
        For i = 0 To n - 1
            If i Mod k = 0 Then
                If i <> 0 Then
                    data2(j) = 32
                    data2(j + 2) = 32
                    j = j + 4
                End If
            End If
            c = data(i) \ 16&
            If c < 10 Then
                data2(j) = c + 48     ' "0"..."9"
            Else
                data2(j) = c + 55     ' "A"..."F"
            End If
            c = data(i) And 15&
            If c < 10 Then
                data2(j + 2) = c + 48 ' "0"..."9"
            Else
                data2(j + 2) = c + 55 ' "A"..."F"
            End If
            j = j + 4
        Next i
Debug.Assert j = UBound(data2) + 1
        HexDisplay = data2
    End If

End Function


'Reverse of HexDisplay.  Given a String containing Hex values, convert to byte array data()
'Returns number of bytes n in data(0 ... n-1)
Private Function HexDisplayRev(TheString As String, data() As Byte) As Long
    Dim i As Long
    Dim j As Long
    Dim c As Long
    Dim d As Long
    Dim n As Long
    Dim data2() As Byte

    n = 2 * Len(TheString)
    data2 = TheString

    ReDim data(n \ 4 - 1)

    d = 0
    i = 0
    j = 0
    Do While j < n
        c = data2(j)
        Select Case c
        Case 48 To 57    '"0" ... "9"
            If d = 0 Then   'high
                d = c
            Else            'low
                data(i) = (c - 48) Or ((d - 48) * 16&)
                i = i + 1
                d = 0
            End If
        Case 65 To 70   '"A" ... "F"
            If d = 0 Then   'high
                d = c - 7
            Else            'low
                data(i) = (c - 55) Or ((d - 48) * 16&)
                i = i + 1
                d = 0
            End If
        Case 97 To 102  '"a" ... "f"
            If d = 0 Then   'high
                d = c - 39
            Else            'low
                data(i) = (c - 87) Or ((d - 48) * 16&)
                i = i + 1
                d = 0
            End If
        End Select
        j = j + 2
    Loop
    n = i
    If n = 0 Then
        Erase data
    Else
        ReDim Preserve data(n - 1)
    End If
    HexDisplayRev = n

End Function

