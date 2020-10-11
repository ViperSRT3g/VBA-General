Attribute VB_Name = "mod_Compatibility"
Option Explicit

Public Function CSWITCH(ParamArray Options() As Variant) As Variant
    On Error GoTo ErrorHandler
    If IsMissing(Options) Then CSWITCH = CVErr(xlErrNA): Exit Function
    If Not WorksheetFunction.IsEven(UBound(Options)) Then CSWITCH = Options(UBound(Options)): Exit Function
    Dim Index As Long
    For Index = LBound(Options) + 1 To UBound(Options) Step 2
        If Options(LBound(Options)) = Options(Index) Then
            CSWITCH = Options(Index + 1): Exit Function
        End If
    Next Index
ErrorHandler:
    CSWITCH = CVErr(xlErrNA)
End Function
