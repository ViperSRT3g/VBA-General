VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_Collection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Enhanced VBA.Collections
'An attempt at implementing Scripting.Dictionary-like behavior in a VBA.Collection without using any Scripting.Dictionaries
'It also includes Key Sorting (Include keys with all indices if you want to be able to sort your VBA.Collection properly)
'VBA.Collection Key values can only be strings, sorting assumes ASCII encoding
'If you use the Sort method, please limit the number of indices to below 100,000 as runtime begins to balloon as the count grows

Public Enum CollSortEnum
    Descending
    Ascending
    [_First] = Descending
    [_Last] = Ascending
End Enum

Private CollIndex_      As VBA.Collection   'Stores the Item ordered by Index
Private CollKeys_       As VBA.Collection   'Stores any Keys of Items ordered by Index
Private OverwriteMode_  As Boolean


Private Sub Class_Initialize()
    Set CollIndex_ = New VBA.Collection
    Set CollKeys_ = New VBA.Collection
End Sub
Private Sub Class_Terminate()
    Set CollIndex_ = Nothing
    Set CollKeys_ = Nothing
End Sub


Property Get CompareMode() As Long
    CompareMode = VBA.VbCompareMethod.vbTextCompare
End Property

'Returns TRUE if the OverwriteMode is enabled. This mode allows new items to be added to existing keys
Property Get OverwriteMode() As Boolean
    OverwriteMode = OverwriteMode_
End Property
Property Let OverwriteMode(Value As Boolean)
    OverwriteMode_ = Value
End Property

'Returns a LONG containing the number of elements within the VBA.Collection
Property Get Count() As Long
    Count = CollIndex_.Count
End Property

'Returns an ARRAY of the items of the VBA.Collection
Property Get Items() As Variant
    Items = CArray(CollIndex_)
End Property

'Returns an ARRAY of the keys of the VBA.Collection
Property Get Keys() As Variant
    Keys = CArray(CollKeys_)
End Property

'Returns TRUE if the given Item was successfully added to the VBA.Collection
Public Function Add(Item As Variant, Optional Key As String, Optional ByVal Before As Long, Optional ByVal After As Long) As Boolean
    Dim Init As Long: Init = CollIndex_.Count
    
    If OverwriteMode_ And Len(Key) > 0 Then
        If Exists(Key) Then Call Remove(Key)
    End If
    
    If Init = 0 Then
        CollIndex_.Add Item, Key
    Else
        If Before = 0 And After = 0 Then
            CollIndex_.Add Item, Key
        ElseIf (Not Before = 0) And After = 0 Then
            CollIndex_.Add Item, Key, Before
        ElseIf Before = 0 And (Not After = 0) Then
            CollIndex_.Add Item, Key, After:=After
        ElseIf Not (Before = 0 And After = 0) Then
            CollIndex_.Add Item, Key, Before, After
        End If
    End If
    
    If Len(Key) > 0 Then
        If Init = 0 Then
            CollKeys_.Add Key, Key
        Else
            If Before = 0 And After = 0 Then
                CollKeys_.Add Key, Key
            ElseIf (Not Before = 0) And After = 0 Then
                CollKeys_.Add Key, Key, Before
            ElseIf Before = 0 And (Not After = 0) Then
                CollKeys_.Add Key, Key, After:=After
            ElseIf Not (Before = 0 And After = 0) Then
                CollKeys_.Add Key, Key, Before, After
            End If
        End If
    End If
    
    Add = ((Not OverwriteMode_) And (CollIndex_.Count = Init + 1)) Or (OverwriteMode_ And (CollIndex_.Count = Init))
End Function

'Returns TRUE if the given Key exists within the VBA.Collection
Public Function Exists(Key As String) As Boolean
    On Error Resume Next
    Dim VType As Long: VType = CollIndex_(Key)
    'Exists = Not (VType = 0)
    Exists = (Err.Number = 0)
    On Error GoTo 0
End Function

'Returns a VARIANT of the given Index of the VBA.Collection
Public Function Item(Index As Variant) As Variant
    If IsObject(CollIndex_(Index)) Then
        Set Item = CollIndex_(Index)
    Else
        Item = CollIndex_(Index)
    End If
End Function

'Returns a STRING containing the key of the given Index for the VBA.Collection
Public Function Key(Index As Long) As String
    Key = CollKeys_(Index)
End Function

'Returns TRUE only if the given Index was removed from the VBA.Collection
Public Function Remove(Index As Variant) As Boolean
    Dim Init As Long: Init = CollIndex_.Count
    CollIndex_.Remove Index
    CollKeys_.Remove Index
    Remove = (CollIndex_.Count = Init - 1)
End Function

'Returns TRUE if both VBA.Collections have zero items
Public Function RemoveAll() As Boolean
    Set CollIndex_ = New VBA.Collection
    Set CollKeys_ = New VBA.Collection
    RemoveAll = (CollIndex_.Count + CollKeys_.Count = 0)
End Function

'Returns TRUE if the VBA.Collection is successfully sorted by Keys
'This sorting algorithm is a variant of LSD Radix sort
Public Function Sort(Optional Order As CollSortEnum = CollSortEnum.Descending) As Boolean
    If CollKeys_.Count = 0 Then Exit Function
    
    Dim OutputIndex     As New VBA.Collection
    Dim Bucket          As VBA.Collection       'Contains a list of strings for a given ASCII Value
    Dim Buckets         As New VBA.Collection   'Contains a list of Buckets for all ASCII values
    Dim Item            As Variant
    Dim Index           As Long 'Main looping variable
    Dim SIndex          As Long 'Secondary looping variable
    Dim CIndex          As Long 'Current character being evaluated
    Dim Char            As String
    Dim ASCChar         As String
    Dim Sorted          As Boolean
    Dim Exists          As Boolean
    
    For Index = 1 To CollKeys_.Count
        'Track max key length, used for evaluating whether or not the entire array has been sorted
        CIndex = IIf(Len(CollKeys_(Index)) >= CIndex, Len(CollKeys_(Index)), CIndex)
        DoEvents
    Next Index
    
    Do While Not Sorted
        'Apply one pass of sorting
        For Index = 1 To CollKeys_.Count
            If Len(CollKeys_(Index)) >= CIndex Then
                Char = Mid(CollKeys_(Index), CIndex, 1)
                ASCChar = CStr(Asc(Char))
                
                'Get Bucket from Buckets
                On Error Resume Next
                Set Bucket = Buckets(ASCChar)
                Exists = Not Bucket Is Nothing
                If Bucket Is Nothing Then Set Bucket = New VBA.Collection
                On Error GoTo 0
                
                Bucket.Add CollKeys_(Index)
                If Not Exists Then Buckets.Add Bucket, ASCChar
                Set Bucket = Nothing
            Else
                On Error Resume Next
                Set Bucket = Buckets("0")
                Exists = Not Bucket Is Nothing
                If Bucket Is Nothing Then Set Bucket = New VBA.Collection
                On Error GoTo 0
                
                Bucket.Add CollKeys_(Index)
                If Not Exists Then Buckets.Add Bucket, "0"
                Set Bucket = Nothing
            End If
            DoEvents
        Next Index
        
        'Clear original VBA.Collection to be repopulated with its original keys in sort step order
        Set CollKeys_ = New VBA.Collection
        
        For Index = 0 To 255
            On Error Resume Next
            Set Bucket = Buckets(CStr(Index))
            On Error GoTo 0
            
            If Not Bucket Is Nothing Then
                For Each Item In Bucket
                    CollKeys_.Add CStr(Item), CStr(Item)
                Next Item
                
                Set Bucket = Nothing
                Buckets.Remove CStr(Index)
            End If
            DoEvents
        Next Index
        
        Sorted = (CIndex = 1)
        
        'Decrement Index of the character being evaluated for the next loop
        CIndex = CIndex - 1
        DoEvents
    Loop
    
    If Order = CollSortEnum.Ascending Then
        For Index = CollKeys_.Count To 1 Step -1
            OutputIndex.Add CollIndex_(CollKeys_(Index)), CollKeys_(Index)
        Next Index
        
    ElseIf Order = CollSortEnum.Descending Then
        For Index = 1 To CollKeys_.Count
            OutputIndex.Add CollIndex_(CollKeys_(Index)), CollKeys_(Index)
        Next Index
    End If
    
    Set CollIndex_ = OutputIndex
    Set OutputIndex = Nothing
    Set Buckets = Nothing
    
    Sort = Sorted
End Function


'Returns an ARRAY of the contents of a given VBA.Collection
Private Function CArray(Target As VBA.Collection) As Variant
    If Target Is Nothing Then Exit Function
    
    Dim Output As Variant: ReDim Output(Target.Count - 1) As Variant
    Dim Item As Variant
    Dim Index As Long
    For Each Item In Target
        If IsObject(Item) Then
            Set Output(Index) = Item
        Else
            Output(Index) = Item
        End If
        Index = Index + 1
        DoEvents
    Next Item
    CArray = Output
    Erase Output
End Function
