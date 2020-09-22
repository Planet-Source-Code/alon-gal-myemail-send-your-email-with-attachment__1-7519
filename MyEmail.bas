Attribute VB_Name = "Module1"
Option Explicit

Function bin(ByVal dec_number As Integer) As String
    Dim temp As String
    Dim count As Integer
    Dim x As Integer
    Dim length As Integer
    
    count = 0
    temp = ""
    x = 128
    Do While count < dec_number
        If x > dec_number Then
            temp = temp & "0"
        ElseIf count + x > dec_number Then
            temp = temp & "0"
        Else
            temp = temp & "1"
            count = count + x
        End If
        x = x - (x / 2)
    Loop
    length = Len(temp)
    For x = (length + 1) To 8
        temp = temp & 0
    Next x
    bin = temp
End Function

Function dec(ByVal bin_number As String) As Integer
    Dim temp As Integer
    Dim x, c As Integer
    Dim s As String
    
    temp = 0
    x = 128
    c = 1
    Do While c <= 8
        s = Mid(bin_number, c, 1)
        If s = "1" Then
            temp = temp + x
        End If
        x = x - (x / 2)
        c = c + 1
    Loop
    dec = temp
End Function

Function base64_alphabet(ByVal dec_num As Integer) As String
    Dim temp As String
    
    Select Case dec_num
        Case 0: temp = "A"
        Case 1: temp = "B"
        Case 2: temp = "C"
        Case 3: temp = "D"
        Case 4: temp = "E"
        Case 5: temp = "F"
        Case 6: temp = "G"
        Case 7: temp = "H"
        Case 8: temp = "I"
        Case 9: temp = "J"
        Case 10: temp = "K"
        Case 11: temp = "L"
        Case 12: temp = "M"
        Case 13: temp = "N"
        Case 14: temp = "O"
        Case 15: temp = "P"
        Case 16: temp = "Q"
        Case 17: temp = "R"
        Case 18: temp = "S"
        Case 19: temp = "T"
        Case 20: temp = "U"
        Case 21: temp = "V"
        Case 22: temp = "W"
        Case 23: temp = "X"
        Case 24: temp = "Y"
        Case 25: temp = "Z"
        Case 26: temp = "a"
        Case 27: temp = "b"
        Case 28: temp = "c"
        Case 29: temp = "d"
        Case 30: temp = "e"
        Case 31: temp = "f"
        Case 32: temp = "g"
        Case 33: temp = "h"
        Case 34: temp = "i"
        Case 35: temp = "j"
        Case 36: temp = "k"
        Case 37: temp = "l"
        Case 38: temp = "m"
        Case 39: temp = "n"
        Case 40: temp = "o"
        Case 41: temp = "p"
        Case 42: temp = "q"
        Case 43: temp = "r"
        Case 44: temp = "s"
        Case 45: temp = "t"
        Case 46: temp = "u"
        Case 47: temp = "v"
        Case 48: temp = "w"
        Case 49: temp = "x"
        Case 50: temp = "y"
        Case 51: temp = "z"
        Case 52: temp = "0"
        Case 53: temp = "1"
        Case 54: temp = "2"
        Case 55: temp = "3"
        Case 56: temp = "4"
        Case 57: temp = "5"
        Case 58: temp = "6"
        Case 59: temp = "7"
        Case 60: temp = "8"
        Case 61: temp = "9"
        Case 62: temp = "+"
        Case 63: temp = "/"
    End Select
    base64_alphabet = temp
End Function

Function base64_encode(ByVal str_24bits As String) As String
    Dim temp As String
    Dim x, v, w As String
    Dim i, z, y As Integer
    Dim bits_6(4) As String
    Dim bits_8(4) As String
    Dim dec_num(4) As Integer
    Dim base64_val(4) As String
    
    x = ""
    v = ""
    w = ""
    temp = ""
    z = Len(str_24bits)
    For i = 1 To z
        w = Mid(str_24bits, i, 1)
        y = Asc(w)
        v = v & bin(y)
    Next i
    If z < 3 Then
        For i = (z + 1) To 3
            x = x & "00000000"
        Next i
    End If
    v = v & x
    z = 1
    For i = 1 To 4
        x = Mid(v, z, 6)
        z = z + 6
        bits_6(i) = x
    Next i
    For i = 1 To 4
        bits_8(i) = "00" & bits_6(i)
    Next i
    For i = 1 To 4
        dec_num(i) = dec(bits_8(i))
    Next i
    For i = 1 To 4
        base64_val(i) = base64_alphabet(dec_num(i))
    Next i
    For i = 1 To 4
        temp = temp & base64_val(i)
    Next i
    base64_encode = temp
End Function

Function base64_encode_string(str As String) As String
    Dim temp As String
    Dim i, x, v, y As Long
    Dim s, u As String
    Dim crlf() As String
    
    x = Len(str)
    i = x / 76
    i = i + 2
    ReDim crlf(i)
    For i = 1 To x Step 3
        s = Mid(str, i, 3)
        u = base64_encode(s)
        temp = temp & u
    Next i
    u = ""
    i = 0
    v = 1
    y = 1
    x = Len(temp)
    Do While i <= x
        i = i + 1
        If (i Mod 76) = 0 Then
           s = Mid(temp, v, 76)
            crlf(y) = s
            y = y + 1
            v = v + 76
        End If
        crlf(y) = Mid(temp, v, (x + 1) - v)
    Loop
    For i = 1 To y - 1
        crlf(i) = crlf(i) & vbCrLf
    Next i
    temp = ""
    For i = 1 To y
        temp = temp & crlf(i)
    Next i
    base64_encode_string = temp
End Function
