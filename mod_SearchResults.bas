Attribute VB_Name = "mod_SearchResults"
Option Explicit

'Returns a range containing only cells that match the given value
Public Function RangeFindAll(ByRef SearchRange As Range, ByVal Value As Variant, Optional ByVal LookIn As XlFindLookIn = xlValues, Optional LookAt As XlLookAt = xlPart) As Range
    Dim FoundValues As Range, Found As Range, Prev As Range, Looper As Boolean: Looper = True
    Do While Looper
        'If we've found something before, then search from after that point
        If Not Prev Is Nothing Then Set Found = SearchRange.Find(Value, Prev, LookIn, LookAt)
        'If we haven't searched for anything before, then do an initial search
        If Found Is Nothing Then Set Found = SearchRange.Find(Value, LookIn:=LookIn, LookAt:=LookAt)
        If Not Found Is Nothing Then
            'If our search found something
            If FoundValues Is Nothing Then
                'If our found value repository is empty, then set it to what we just found
                Set FoundValues = Found
            Else
                If Not Intersect(Found, FoundValues) Is Nothing Then Looper = False
                'If the found value intersects with what we've already found, then we've looped through the SearchRange
                'Note: This check is performed BEFORE we insert the newly found data into our repository
                
                Set FoundValues = Union(FoundValues, Found)
                'If our found value repository contains data, then add what we just found to it
            End If
            Set Prev = Found
        End If
        If Found Is Nothing And Prev Is Nothing Then Exit Function
    Loop
    Set RangeFindAll = FoundValues
    Set FoundValues = Nothing
    Set Found = Nothing
    Set Prev = Nothing
End Function
