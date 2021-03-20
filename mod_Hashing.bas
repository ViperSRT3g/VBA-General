Attribute VB_Name = "mod_Hashing"
Option Explicit

'HMAC256 Functions
Public Function HMAC_256(ByRef Key As Variant, ByRef InputString As Variant) As Byte()
    Dim HMAC As Object: Set HMAC = CreateObject("System.Security.Cryptography.HMACSHA256")
    HMAC.Key = S2B(Key)
    HMAC_256 = HMAC.ComputeHash_2(S2B(InputString))
    Set HMAC = Nothing
End Function
Public Function HMAC_256_String(ByRef Key As Variant, ByRef InputString As Variant, Optional UpperCase As Boolean) As String
    Dim HMAC As Object: Set HMAC = CreateObject("System.Security.Cryptography.HMACSHA256")
    HMAC.Key = S2B(Key)
    Dim BArr() As Byte: BArr = HMAC.ComputeHash_2(S2B(InputString))
    HMAC_256_String = B2H(BArr, UpperCase)
    Set HMAC = Nothing
End Function

'HMAC512 Functions
Public Function HMAC_512(ByRef Key As Variant, ByRef InputString As Variant) As Byte()
    Dim HMAC As Object: Set HMAC = CreateObject("System.Security.Cryptography.HMACSHA512")
    HMAC.Key = S2B(Key)
    HMAC_512 = HMAC.ComputeHash_2(S2B(InputString))
    Set HMAC = Nothing
End Function
Public Function HMAC_512_String(ByRef Key As Variant, ByRef InputString As Variant, Optional UpperCase As Boolean) As String
    Dim HMAC As Object: Set HMAC = CreateObject("System.Security.Cryptography.HMACSHA512")
    HMAC.Key = S2B(Key)
    Dim BArr() As Byte: BArr = HMAC.ComputeHash_2(S2B(InputString))
    HMAC_512_String = B2H(BArr, UpperCase)
    Set HMAC = Nothing
End Function

'SHA256 Functions
Public Function SHA256(ByRef InputString As Variant) As Byte()
    Dim SHA256_ As Object: Set SHA256_ = CreateObject("System.Security.Cryptography.SHA256Managed")
    SHA256 = SHA256_.ComputeHash_2(S2B(InputString))
    Set SHA256_ = Nothing
End Function
Public Function SHA256_String(ByRef InputString As Variant, Optional UpperCase As Boolean) As String
    Dim SHA256_ As Object: Set SHA256_ = CreateObject("System.Security.Cryptography.SHA256Managed")
    Dim BArr() As Byte: BArr = SHA256_.ComputeHash_2(S2B(InputString))
    SHA256_String = B2H(BArr, UpperCase)
    Set SHA256_ = Nothing
End Function

'SHA512 Functions
Public Function SHA512(ByRef InputString As Variant) As Byte()
    Dim SHA512_ As Object: Set SHA512_ = CreateObject("System.Security.Cryptography.SHA512Managed")
    SHA512 = SHA512_.ComputeHash_2(S2B(InputString))
    Set SHA512_ = Nothing
End Function
Public Function SHA512_String(ByRef InputString As Variant, Optional UpperCase As Boolean) As String
    Dim SHA512_ As Object: Set SHA512_ = CreateObject("System.Security.Cryptography.SHA512Managed")
    Dim BArr() As Byte: BArr = SHA512_.ComputeHash_2(S2B(InputString))
    SHA512_String = B2H(BArr, UpperCase)
    Set SHA512_ = Nothing
End Function


Private Function S2B(ByRef InputStr As Variant) As Byte()
    If VarType(InputStr) = vbArray + vbByte Then
        S2B = InputStr
    ElseIf VarType(InputStr) = vbString Then
        S2B = StrConv(InputStr, vbFromUnicode)
    Else
        Exit Function
    End If
End Function
Private Function B2H(ByRef ByteArr() As Byte, Optional UpperCase As Boolean) As String
    Dim i As Long: For i = 0 To UBound(ByteArr): B2H = B2H & Right(Hex(256 Or ByteArr(i)), 2): Next
    B2H = IIf(UpperCase, UCase(B2H), LCase(B2H))
End Function

