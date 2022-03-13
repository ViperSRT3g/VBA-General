Attribute VB_Name = "mod_FuzzyComparisons"
Option Explicit

'Returns a percentage of similarity between two strings utilizing the N-Gram comparison (aka K-Mer)
'See https://en.wikipedia.org/wiki/N-gram for more information
Public Function NGramScore(ByVal Str1 As String, Str2 As String) As Single
    Dim Score1 As Object: Set Score1 = GenNGram(Str1): If Score1 Is Nothing Then Exit Function
    Dim Score2 As Object: Set Score2 = GenNGram(Str2): If Score2 Is Nothing Then Set Score1 = Nothing: Exit Function
    Dim Index As Long
    For Index = 0 To Score1.Count - 1
        If InStr(1, Str2, Score1.Keys()(Index)) Then NGramScore = NGramScore + 1
    Next Index
    For Index = 0 To Score2.Count - 1
        If InStr(1, Str1, Score2.Keys()(Index)) Then NGramScore = NGramScore + 1
    Next Index
    NGramScore = NGramScore / (Score1.Count + Score2.Count)
    Set Score1 = Nothing: Set Score2 = Nothing
End Function

Private Function GenNGram(ByVal InputStr As String) As Object
    Dim NGrams As Object: Set NGrams = CreateObject("Scripting.Dictionary")
    Dim Index As Long, NGramDegree As Long
    Do Until NGramDegree = Len(InputStr) - 1
        NGramDegree = NGramDegree + 1
        For Index = 1 To Len(InputStr): NGrams(Mid(InputStr, Index, NGramDegree)) = NGrams(Mid(InputStr, Index, NGramDegree)) + 1: Next Index
    Loop
    Set GenNGram = NGrams: Set NGrams = Nothing
End Function
