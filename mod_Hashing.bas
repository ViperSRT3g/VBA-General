Attribute VB_Name = "mod_Hashing"
Option Explicit

'HMAC256 Functions
Public Function HMAC_256(ByRef Key As Variant, ByRef InputString As Variant) As Byte()
    With CreateObject("System.Security.Cryptography.HMACSHA256")
        .Key = S2B(Key)
        HMAC_256 = .ComputeHash_2(S2B(InputString))
    End With
End Function
Public Function HMAC_256_String(ByRef Key As Variant, ByRef InputString As Variant, Optional UpperCase As Boolean) As String
    With CreateObject("System.Security.Cryptography.HMACSHA256")
        .Key = S2B(Key)
        Dim BArr() As Byte: BArr = .ComputeHash_2(S2B(InputString))
        HMAC_256_String = B2H(BArr, UpperCase)
    End With
End Function

'HMAC512 Functions
Public Function HMAC_512(ByRef Key As Variant, ByRef InputString As Variant) As Byte()
    With CreateObject("System.Security.Cryptography.HMACSHA512")
        .Key = S2B(Key)
        HMAC_512 = .ComputeHash_2(S2B(InputString))
    End With
End Function
Public Function HMAC_512_String(ByRef Key As Variant, ByRef InputString As Variant, Optional UpperCase As Boolean) As String
    With CreateObject("System.Security.Cryptography.HMACSHA512")
        .Key = S2B(Key)
        Dim BArr() As Byte: BArr = .ComputeHash_2(S2B(InputString))
        HMAC_512_String = B2H(BArr, UpperCase)
    End With
End Function

'SHA256 Functions
Public Function SHA256(ByRef InputString As Variant) As Byte()
    With CreateObject("System.Security.Cryptography.SHA256Managed"): SHA256 = .ComputeHash_2(S2B(InputString)): End With
End Function
Public Function SHA256_String(ByRef InputString As Variant, Optional UpperCase As Boolean) As String
    With CreateObject("System.Security.Cryptography.SHA256Managed")
        Dim BArr() As Byte: BArr = .ComputeHash_2(S2B(InputString))
        SHA256_String = B2H(BArr, UpperCase)
    End With
End Function

'SHA512 Functions
Public Function SHA512(ByRef InputString As Variant) As Byte()
    With CreateObject("System.Security.Cryptography.SHA512Managed"): SHA512 = .ComputeHash_2(S2B(InputString)): End With
End Function
Public Function SHA512_String(ByRef InputString As Variant, Optional UpperCase As Boolean) As String
    With CreateObject("System.Security.Cryptography.SHA512Managed")
        Dim BArr() As Byte: BArr = .ComputeHash_2(S2B(InputString))
        SHA512_String = B2H(BArr, UpperCase)
    End With
End Function


Private Function S2B(ByRef InputStr As Variant) As Byte()
    If VarType(InputStr) = vbArray + vbByte Then
        S2B = InputStr
    ElseIf VarType(InputStr) = vbString Then
        S2B = StrConv(InputStr, vbFromUnicode)
    End If
End Function
Private Function B2H(ByRef ByteArr() As Byte, Optional UpperCase As Boolean) As String
    Dim i As Long: For i = 0 To UBound(ByteArr): B2H = B2H & Right(Hex(256 Or ByteArr(i)), 2): Next
    B2H = IIf(UpperCase, UCase(B2H), LCase(B2H))
End Function
