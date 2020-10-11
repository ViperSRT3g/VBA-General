Attribute VB_Name = "mod_ArrayAPIs"
Option Explicit

Public Function IsListed(ByVal Value As Variant, ByRef arr As Variant) As Boolean
    IsListed = (UBound(Filter(arr, Value)) > -1)
End Function

Public Function IdenticalElements(ByRef arr As Variant) As Boolean
    IdenticalElements = UBound(Filter(arr, arr(LBound(arr)))) = UBound(arr)
End Function

Public Function AddToArray(ByVal Value As Variant, ByRef arr As Variant) As Long
    For AddToArray = 0 To UBound(arr)
        If Len(arr(AddToArray)) = 0 Or arr(AddToArray) = 0 Then
            arr(AddToArray) = Value
            Exit Function
        End If
    Next AddToArray
End Function

Public Function CArray(ByVal TargetRange As Range) As Variant
    'Converts cell ranges into a SINGLE DIMENSION ARRAY of values
    'Excel by default creates multidimensional arrays for Variant stored range arrays
    If TargetRange Is Nothing Then Exit Function
    Dim NonBlankCells As Range: Set NonBlankCells = TargetRange.SpecialCells(xlCellTypeConstants, xlLogical + xlNumbers + xlTextValues)
    Dim NewArray As Variant: ReDim NewArray(NonBlankCells.Cells.Count - 1) As Variant
    Dim Index As Long, TCell As Range
    For Each TCell In NonBlankCells
        NewArray(Index) = TCell.Value
        Index = Index + 1
    Next TCell
    Set TCell = Nothing
    Set NonBlankCells = Nothing
    CArray = NewArray
End Function

Public Function UniqueArray(ByRef FullArray As Variant) As Variant
    If UBound(FullArray) < 0 Then Exit Function
    Dim NewArray As Variant: ReDim NewArray(UBound(FullArray)) As Variant
    Dim Index As Long, IndexUBound As Long
    For Index = LBound(FullArray) To UBound(FullArray)
        If Not IsListed(FullArray(Index), NewArray) Then IndexUBound = AddToArray(FullArray(Index), NewArray)
    Next Index
    ReDim Preserve NewArray(IndexUBound) As Variant
    UniqueArray = NewArray
End Function

Public Function AddRange(ByRef MainRange As Range, AddedRange As Range) As Range
    If MainRange Is Nothing Or AddedRange Is Nothing Then Exit Function
    Dim TCell As Range
    For Each TCell In AddedRange
        If Intersect(TCell, AddedRange) Is Nothing Then 'If TCell doesn't intersects with any cell in the Added Range
            Set AddRange = Union(MainRange, TCell)
        End If
    Next
    Set TCell = Nothing
End Function

Public Function RemoveRange(ByRef MainRange As Range, ExcludeRange As Range) As Range
    If MainRange Is Nothing Or ExcludeRange Is Nothing Then Exit Function
    Dim TCell As Range
    For Each TCell In MainRange
        If Intersect(TCell, ExcludeRange) Is Nothing Then 'If TCell intersects with any cell in the Exclusion Range
            If RemoveRange Is Nothing Then
                Set RemoveRange = TCell
            Else
                Set RemoveRange = Union(RemoveRange, TCell)
            End If
        End If
    Next
    Set TCell = Nothing
End Function

Public Sub Sort2DArray(ByRef TargetArray As Variant, Optional ByVal SortBothColumns As Boolean)
    Dim i As Integer, j As Integer, ci As Integer, c As Integer
    Dim temp As Variant
    
    'Bubble sort 1st column
    ci = LBound(TargetArray, 2) '1st column index
    For i = LBound(TargetArray) To UBound(TargetArray) - 1
        For j = i + 1 To UBound(TargetArray)
            If TargetArray(i, ci) < TargetArray(j, ci) Then
                For c = LBound(TargetArray, 2) To UBound(TargetArray, 2)
                    temp = TargetArray(i, c)
                    TargetArray(i, c) = TargetArray(j, c)
                    TargetArray(j, c) = temp
                Next
            End If
        Next
    Next
    
    If SortBothColumns = False Then Exit Sub
    
    'Bubble sort 2nd column, where adjacent rows in 1st column are equal
    ci = LBound(TargetArray, 2) + 1 '2nd column index
    For i = LBound(TargetArray) To UBound(TargetArray) - 1
        For j = i + 1 To UBound(TargetArray)
            If TargetArray(i, ci - 1) = TargetArray(j, ci - 1) Then 'compare adjacent rows in 1st column
                If TargetArray(i, ci) < TargetArray(j, ci) Then
                    For c = LBound(TargetArray, 2) To UBound(TargetArray, 2)
                        temp = TargetArray(i, c)
                        TargetArray(i, c) = TargetArray(j, c)
                        TargetArray(j, c) = temp
                    Next
                End If
            End If
        Next
    Next
End Sub

Public Function QSortInPlace( _
    ByRef InputArray As Variant, _
    Optional ByVal LB As Long = -1&, _
    Optional ByVal UB As Long = -1&, _
    Optional ByVal Descending As Boolean = False, _
    Optional ByVal CompareMode As VbCompareMethod = vbTextCompare, _
    Optional ByVal NoAlerts As Boolean = False) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' QSortInPlace
    '
    ' This function sorts the array InputArray in place -- this is, the original array in the
    ' calling procedure is sorted. It will work with either string data or numeric data.
    ' It need not sort the entire array. You can sort only part of the array by setting the LB and
    ' UB parameters to the first (LB) and last (UB) element indexes that you want to sort.
    ' LB and UB are optional parameters. If omitted LB is set to the LBound of InputArray, and if
    ' omitted UB is set to the UBound of the InputArray. If you want to sort the entire array,
    ' omit the LB and UB parameters, or set both to -1, or set LB = LBound(InputArray) and set
    ' UB to UBound(InputArray).
    '
    ' By default, the sort method is case INSENSTIVE (case doens't matter: "A", "b", "C", "d").
    ' To make it case SENSITIVE (case matters: "A" "C" "b" "d"), set the CompareMode argument
    ' to vbBinaryCompare (=0). If Compare mode is omitted or is any value other than vbBinaryCompare,
    ' it is assumed to be vbTextCompare and the sorting is done case INSENSITIVE.
    '
    ' The function returns TRUE if the array was successfully sorted or FALSE if an error
    ' occurred. If an error occurs (e.g., LB > UB), a message box indicating the error is
    ' displayed. To suppress message boxes, set the NoAlerts parameter to TRUE.
    '
    ''''''''''''''''''''''''''''''''''''''
    ' MODIFYING THIS CODE:
    ''''''''''''''''''''''''''''''''''''''
    ' If you modify this code and you call "Exit Procedure", you MUST decrment the RecursionLevel
    ' variable. E.g.,
    '       If SomethingThatCausesAnExit Then
    '           RecursionLevel = RecursionLevel - 1
    '           Exit Function
    '       End If
    '''''''''''''''''''''''''''''''''''''''
    '
    ' Note: If you coerce InputArray to a ByVal argument, QSortInPlace will not be
    ' able to reference the InputArray in the calling procedure and the array will
    ' not be sorted.
    '
    ' This function uses the following procedures. These are declared as Private procedures
    ' at the end of this module:
    '       IsArrayAllocated
    '       IsSimpleDataType
    '       IsSimpleNumericType
    '       QSortCompare
    '       NumberOfArrayDimensions
    '       ReverseArrayInPlace
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim temp As Variant
    Dim Buffer As Variant
    Dim CurLow As Long
    Dim CurHigh As Long
    Dim CurMidpoint As Long
    Dim Ndx As Long
    Dim pCompareMode As VbCompareMethod
    '''''''''''''''''''''''''
    ' Set the default result.
    '''''''''''''''''''''''''
    QSortInPlace = False
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' This variable is used to determine the level
    ' of recursion  (the function calling itself).
    ' RecursionLevel is incremented when this procedure
    ' is called, either initially by a calling procedure
    ' or recursively by itself. The variable is decremented
    ' when the procedure exits. We do the input parameter
    ' validation only when RecursionLevel is 1 (when
    ' the function is called by another function, not
    ' when it is called recursively).
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Static RecursionLevel As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Keep track of the recursion level -- that is, how many
    ' times the procedure has called itself.
    ' Carry out the validation routines only when this
    ' procedure is first called. Don't run the
    ' validations on a recursive call to the
    ' procedure.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    RecursionLevel = RecursionLevel + 1
    If RecursionLevel = 1 Then
        ''''''''''''''''''''''''''''''''''
        ' Ensure InputArray is an array.
        ''''''''''''''''''''''''''''''''''
        If IsArray(InputArray) = False Then
            If NoAlerts = False Then
                MsgBox "The InputArray parameter is not an array."
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' InputArray is not an array. Exit with a False result.
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''
            RecursionLevel = RecursionLevel - 1
            Exit Function
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Test LB and UB. If < 0 then set to LBound and UBound
        ' of the InputArray.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If LB < 0 Then LB = LBound(InputArray)
        If UB < 0 Then UB = UBound(InputArray)
        
        Select Case NumberOfArrayDimensions(InputArray)
            Case 0
                ''''''''''''''''''''''''''''''''''''''''''
                ' Zero dimensions indicates an unallocated
                ' dynamic array.
                ''''''''''''''''''''''''''''''''''''''''''
                If NoAlerts = False Then
                    MsgBox "The InputArray is an empty, unallocated array."
                End If
                RecursionLevel = RecursionLevel - 1
                Exit Function
            Case 1
                ''''''''''''''''''''''''''''''''''''''''''
                ' We sort ONLY single dimensional arrays.
                ''''''''''''''''''''''''''''''''''''''''''
            Case Else
                ''''''''''''''''''''''''''''''''''''''''''
                ' We sort ONLY single dimensional arrays.
                ''''''''''''''''''''''''''''''''''''''''''
                If NoAlerts = False Then
                    MsgBox "The InputArray is multi-dimensional." & _
                          "QSortInPlace works only on single-dimensional arrays."
                End If
                RecursionLevel = RecursionLevel - 1
                Exit Function
        End Select
        '''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Ensure that InputArray is an array of simple data
        ' types, not other arrays or objects. This tests
        ' the data type of only the first element of
        ' InputArray. If InputArray is an array of Variants,
        ' subsequent data types may not be simple data types
        ' (e.g., they may be objects or other arrays), and
        ' this may cause QSortInPlace to fail on the StrComp
        ' operation.
        '''''''''''''''''''''''''''''''''''''''''''''''''''
        If IsSimpleDataType(InputArray(LBound(InputArray))) = False Then
            If NoAlerts = False Then
                MsgBox "InputArray is not an array of simple data types."
                RecursionLevel = RecursionLevel - 1
                Exit Function
            End If
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' ensure that the LB parameter is valid.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        Select Case LB
            Case Is < LBound(InputArray)
                If NoAlerts = False Then
                    MsgBox "The LB lower bound parameter is less than the LBound of the InputArray"
                End If
                RecursionLevel = RecursionLevel - 1
                Exit Function
            Case Is > UBound(InputArray)
                If NoAlerts = False Then
                    MsgBox "The LB lower bound parameter is greater than the UBound of the InputArray"
                End If
                RecursionLevel = RecursionLevel - 1
                Exit Function
            Case Is > UB
                If NoAlerts = False Then
                    MsgBox "The LB lower bound parameter is greater than the UB upper bound parameter."
                End If
                RecursionLevel = RecursionLevel - 1
                Exit Function
        End Select
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' ensure the UB parameter is valid.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        Select Case UB
            Case Is > UBound(InputArray)
                If NoAlerts = False Then
                    MsgBox "The UB upper bound parameter is greater than the upper bound of the InputArray."
                End If
                RecursionLevel = RecursionLevel - 1
                Exit Function
            Case Is < LBound(InputArray)
                If NoAlerts = False Then
                    MsgBox "The UB upper bound parameter is less than the lower bound of the InputArray."
                End If
                RecursionLevel = RecursionLevel - 1
                Exit Function
            Case Is < LB
                If NoAlerts = False Then
                    MsgBox "the UB upper bound parameter is less than the LB lower bound parameter."
                End If
                RecursionLevel = RecursionLevel - 1
                Exit Function
        End Select
    
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' if UB = LB, we have nothing to sort, so get out.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If UB = LB Then
            QSortInPlace = True
            RecursionLevel = RecursionLevel - 1
            Exit Function
        End If
    End If ' RecursionLevel = 1
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Ensure that CompareMode is either vbBinaryCompare  or
    ' vbTextCompare. If it is neither, default to vbTextCompare.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (CompareMode = vbBinaryCompare) Or (CompareMode = vbTextCompare) Then
        pCompareMode = CompareMode
    Else
        pCompareMode = vbTextCompare
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Begin the actual sorting process.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    CurLow = LB
    CurHigh = UB
    
    If LB = 0 Then
        CurMidpoint = ((LB + UB) \ 2) + 1
    Else
        CurMidpoint = (LB + UB) \ 2 ' note integer division (\) here
    End If
    temp = InputArray(CurMidpoint)
    
    Do While (CurLow <= CurHigh)
        Do While QSortCompare(V1:=InputArray(CurLow), V2:=temp, CompareMode:=pCompareMode) < 0
            CurLow = CurLow + 1
            If CurLow = UB Then Exit Do
        Loop
        Do While QSortCompare(V1:=temp, V2:=InputArray(CurHigh), CompareMode:=pCompareMode) < 0
            CurHigh = CurHigh - 1
            If CurHigh = LB Then Exit Do
        Loop
        If (CurLow <= CurHigh) Then
            Buffer = InputArray(CurLow)
            InputArray(CurLow) = InputArray(CurHigh)
            InputArray(CurHigh) = Buffer
            CurLow = CurLow + 1
            CurHigh = CurHigh - 1
        End If
    Loop
    
    If LB < CurHigh Then QSortInPlace InputArray:=InputArray, LB:=LB, UB:=CurHigh, _
            Descending:=Descending, CompareMode:=pCompareMode, NoAlerts:=True
    
    If CurLow < UB Then QSortInPlace InputArray:=InputArray, LB:=CurLow, UB:=UB, _
            Descending:=Descending, CompareMode:=pCompareMode, NoAlerts:=True
    '''''''''''''''''''''''''''''''''''''
    ' If Descending is True, reverse the
    ' order of the array, but only if the
    ' recursion level is 1.
    '''''''''''''''''''''''''''''''''''''
    If Descending = True And RecursionLevel = 1 Then ReverseArrayInPlace2 InputArray, LB, UB
    RecursionLevel = RecursionLevel - 1
    QSortInPlace = True
End Function

Public Function QSortCompare(V1 As Variant, V2 As Variant, _
    Optional CompareMode As VbCompareMethod = vbTextCompare) As Long
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' QSortCompare
    ' This function is used in QSortInPlace to compare two elements. If
    ' V1 AND V2 are both numeric data types (integer, long, single, double)
    ' they are converted to Doubles and compared. If V1 and V2 are BOTH strings
    ' that contain numeric data, they are converted to Doubles and compared.
    ' If either V1 or V2 is a string and does NOT contain numeric data, both
    ' V1 and V2 are converted to Strings and compared with StrComp.
    '
    ' The result is -1 if V1 < V2,
    '                0 if V1 = V2
    '                1 if V1 > V2
    ' For text comparisons, case sensitivity is controlled by CompareMode.
    ' If this is vbBinaryCompare, the result is case SENSITIVE. If this
    ' is omitted or any other value, the result is case INSENSITIVE.
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim D1 As Double
    Dim D2 As Double
    Dim S1 As String
    Dim S2 As String
    
    Dim Compare As VbCompareMethod
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ' Test CompareMode. Any value other than
    ' vbBinaryCompare will default to vbTextCompare.
    ''''''''''''''''''''''''''''''''''''''''''''''''
    If CompareMode = vbBinaryCompare Or CompareMode = vbTextCompare Then
        Compare = CompareMode
    Else
        Compare = vbTextCompare
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''
    ' If either V1 or V2 is either an array or
    ' an Object, raise a error 13 - Type Mismatch.
    '''''''''''''''''''''''''''''''''''''''''''''''
    If IsArray(V1) = True Or IsArray(V2) = True Then
        Err.Raise 13
        Exit Function
    End If
    If IsObject(V1) = True Or IsObject(V2) = True Then
        Err.Raise 13
        Exit Function
    End If
    
    If IsSimpleNumericType(V1) = True Then
        If IsSimpleNumericType(V2) = True Then
            '''''''''''''''''''''''''''''''''''''
            ' If BOTH V1 and V2 are numeric data
            ' types, then convert to Doubles and
            ' do an arithmetic compare and
            ' return the result.
            '''''''''''''''''''''''''''''''''''''
            D1 = CDbl(V1)
            D2 = CDbl(V2)
            If D1 = D2 Then
                QSortCompare = 0
                Exit Function
            End If
            If D1 < D2 Then
                QSortCompare = -1
                Exit Function
            End If
            If D1 > D2 Then
                QSortCompare = 1
                Exit Function
            End If
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''
    ' Either V1 or V2 was not numeric data type.
    ' Test whether BOTH V1 AND V2 are numeric
    ' strings. If BOTH are numeric, convert to
    ' Doubles and do a arithmetic comparison.
    ''''''''''''''''''''''''''''''''''''''''''''
    If IsNumeric(V1) = True And IsNumeric(V2) = True Then
        D1 = CDbl(V1)
        D2 = CDbl(V2)
        If D1 = D2 Then
            QSortCompare = 0
            Exit Function
        End If
        If D1 < D2 Then
            QSortCompare = -1
            Exit Function
        End If
        If D1 > D2 Then
            QSortCompare = 1
            Exit Function
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''
    ' Either or both V1 and V2 was not numeric
    ' string. In this case, convert to Strings
    ' and use StrComp to compare.
    ''''''''''''''''''''''''''''''''''''''''''''''
    S1 = CStr(V1)
    S2 = CStr(V2)
    QSortCompare = StrComp(S1, S2, Compare)
End Function

Public Function NumberOfArrayDimensions(arr As Variant) As Integer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' NumberOfArrayDimensions
    ' This function returns the number of dimensions of an array. An unallocated dynamic array
    ' has 0 dimensions. This condition can also be tested with IsArrayEmpty.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Ndx As Integer
    Dim Res As Integer
    On Error Resume Next
    ' Loop, increasing the dimension index Ndx, until an error occurs.
    ' An error will occur when Ndx exceeds the number of dimension
    ' in the array. Return Ndx - 1.
    Do
        Ndx = Ndx + 1
        Res = UBound(arr, Ndx)
    Loop Until Err.Number <> 0
    NumberOfArrayDimensions = Ndx - 1
End Function


Public Function ReverseArrayInPlace(InputArray As Variant, _
    Optional NoAlerts As Boolean = False) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ReverseArrayInPlace
    ' This procedure reverses the order of an array in place -- this is, the array variable
    ' in the calling procedure is sorted. An error will occur if InputArray is not an array,
     'if it is an empty, unallocated array, or if the number of dimensions is not 1.
    '
    ' NOTE: Before calling the ReverseArrayInPlace procedure, consider if your needs can
    ' be met by simply reading the existing array in reverse order (Step -1). If so, you can save
    ' the overhead added to your application by calling this function.
    '
    ' The function returns TRUE if the array was successfully reversed, or FALSE if
    ' an error occurred.
    '
    ' If an error occurred, a message box is displayed indicating the error. To suppress
    ' the message box and simply return FALSE, set the NoAlerts parameter to TRUE.
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim temp As Variant
    Dim Ndx As Long
    Dim Ndx2 As Long
    Dim OrigN As Long
    Dim NewN As Long
    Dim NewArr() As Variant
    ''''''''''''''''''''''''''''''''
    ' Set the default return value.
    ''''''''''''''''''''''''''''''''
    ReverseArrayInPlace = False
    '''''''''''''''''''''''''''''''''
    ' Ensure we have an array
    '''''''''''''''''''''''''''''''''
    If IsArray(InputArray) = False Then
       If NoAlerts = False Then
            MsgBox "The InputArray parameter is not an array."
        End If
        Exit Function
    End If
    ''''''''''''''''''''''''''''''''''''''
    ' Test the number of dimensions of the
    ' InputArray. If 0, we have an empty,
    ' unallocated array. Get out with
    ' an error message. If greater than
    ' one, we have a multi-dimensional
    ' array, which is not allowed. Only
    ' an allocated 1-dimensional array is
    ' allowed.
    ''''''''''''''''''''''''''''''''''''''
    Select Case NumberOfArrayDimensions(InputArray)
        Case 0
            '''''''''''''''''''''''''''''''''''''''''''
            ' Zero dimensions indicates an unallocated
            ' dynamic array.
            '''''''''''''''''''''''''''''''''''''''''''
            If NoAlerts = False Then
                MsgBox "The input array is an empty, unallocated array."
            End If
            Exit Function
        Case 1
            '''''''''''''''''''''''''''''''''''''''''''
            ' We can reverse ONLY a single dimensional
            ' arrray.
            '''''''''''''''''''''''''''''''''''''''''''
        Case Else
            '''''''''''''''''''''''''''''''''''''''''''
            ' We can reverse ONLY a single dimensional
            ' arrray.
            '''''''''''''''''''''''''''''''''''''''''''
            If NoAlerts = False Then
                MsgBox "The input array multi-dimensional. ReverseArrayInPlace works only " & _
                       "on single-dimensional arrays."
            End If
            Exit Function
    End Select
    '''''''''''''''''''''''''''''''''''''''''''''
    ' Ensure that we have only simple data types,
    ' not an array of objects or arrays.
    '''''''''''''''''''''''''''''''''''''''''''''
    If IsSimpleDataType(InputArray(LBound(InputArray))) = False Then
        If NoAlerts = False Then
            MsgBox "The input array contains arrays, objects, or other complex data types." & vbCrLf & _
                "ReverseArrayInPlace can reverse only arrays of simple data types."
            Exit Function
        End If
    End If
    
    ReDim NewArr(LBound(InputArray) To UBound(InputArray))
    NewN = UBound(NewArr)
    For OrigN = LBound(InputArray) To UBound(InputArray)
        NewArr(NewN) = InputArray(OrigN)
        NewN = NewN - 1
    Next OrigN
    
    For NewN = LBound(NewArr) To UBound(NewArr)
        InputArray(NewN) = NewArr(NewN)
    Next NewN
    ReverseArrayInPlace = True
End Function


Public Function ReverseArrayInPlace2(InputArray As Variant, _
    Optional LB As Long = -1, Optional UB As Long = -1, _
    Optional NoAlerts As Boolean = False) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ReverseArrayInPlace2
    ' This reverses the order of elements in InputArray. To reverse the entire array, omit or
    ' set to less than 0 the LB and UB parameters. To reverse only part of tbe array, set LB and/or
    ' UB to the LBound and UBound of the sub array to be reversed.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim n As Long
    Dim temp As Variant
    Dim Ndx As Long
    Dim Ndx2 As Long
    Dim OrigN As Long
    Dim NewN As Long
    Dim NewArr() As Variant
    
    ''''''''''''''''''''''''''''''''
    ' Set the default return value.
    ''''''''''''''''''''''''''''''''
    ReverseArrayInPlace2 = False
    
    '''''''''''''''''''''''''''''''''
    ' Ensure we have an array
    '''''''''''''''''''''''''''''''''
    If IsArray(InputArray) = False Then
        If NoAlerts = False Then
            MsgBox "The InputArray parameter is not an array."
        End If
        Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''''
    ' Test the number of dimensions of the
    ' InputArray. If 0, we have an empty,
    ' unallocated array. Get out with
    ' an error message. If greater than
    ' one, we have a multi-dimensional
    ' array, which is not allowed. Only
    ' an allocated 1-dimensional array is
    ' allowed.
    ''''''''''''''''''''''''''''''''''''''
    Select Case NumberOfArrayDimensions(InputArray)
        Case 0
            '''''''''''''''''''''''''''''''''''''''''''
            ' Zero dimensions indicates an unallocated
            ' dynamic array.
            '''''''''''''''''''''''''''''''''''''''''''
            If NoAlerts = False Then
                MsgBox "The input array is an empty, unallocated array."
            End If
            Exit Function
        Case 1
            '''''''''''''''''''''''''''''''''''''''''''
            ' We can reverse ONLY a single dimensional
            ' arrray.
            '''''''''''''''''''''''''''''''''''''''''''
        Case Else
            '''''''''''''''''''''''''''''''''''''''''''
            ' We can reverse ONLY a single dimensional
            ' arrray.
            '''''''''''''''''''''''''''''''''''''''''''
            If NoAlerts = False Then
                MsgBox "The input array multi-dimensional. ReverseArrayInPlace works only " & _
                       "on single-dimensional arrays."
            End If
            Exit Function
    
    End Select
    
    '''''''''''''''''''''''''''''''''''''''''''''
    ' Ensure that we have only simple data types,
    ' not an array of objects or arrays.
    '''''''''''''''''''''''''''''''''''''''''''''
    If IsSimpleDataType(InputArray(LBound(InputArray))) = False Then
        If NoAlerts = False Then
            MsgBox "The input array contains arrays, objects, or other complex data types." & vbCrLf & _
                "ReverseArrayInPlace can reverse only arrays of simple data types."
            Exit Function
        End If
    End If
    
    If LB < 0 Then LB = LBound(InputArray)
    If UB < 0 Then UB = UBound(InputArray)
    
    For n = LB To (LB + ((UB - LB - 1) \ 2))
        temp = InputArray(n)
        InputArray(n) = InputArray(UB - (n - LB))
        InputArray(UB - (n - LB)) = temp
    Next n
    ReverseArrayInPlace2 = True
End Function


Public Function IsSimpleNumericType(v As Variant) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' IsSimpleNumericType
    ' This returns TRUE if V is one of the following data types:
    '        vbBoolean
    '        vbByte
    '        vbCurrency
    '        vbDate
    '        vbDecimal
    '        vbDouble
    '        vbInteger
    '        vbLong
    '        vbSingle
    '        vbVariant if it contains a numeric value
    ' It returns FALSE for any other data type, including any array
    ' or vbEmpty.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If IsSimpleDataType(v) = True Then
        Select Case VarType(v)
            Case vbBoolean, _
                    vbByte, _
                    vbCurrency, _
                    vbDate, _
                    vbDecimal, _
                    vbDouble, _
                    vbInteger, _
                    vbLong, _
                    vbSingle
                IsSimpleNumericType = True
            Case vbVariant
                If IsNumeric(v) = True Then
                    IsSimpleNumericType = True
                Else
                    IsSimpleNumericType = False
                End If
            Case Else
                IsSimpleNumericType = False
        End Select
    Else
        IsSimpleNumericType = False
    End If
End Function

Public Function IsSimpleDataType(v As Variant) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' IsSimpleDataType
    ' This function returns TRUE if V is one of the following
    ' variable types (as returned by the VarType function:
    '    vbBoolean
    '    vbByte
    '    vbCurrency
    '    vbDate
    '    vbDecimal
    '    vbDouble
    '    vbEmpty
    '    vbError
    '    vbInteger
    '    vbLong
    '    vbNull
    '    vbSingle
    '    vbString
    '    vbVariant
    '
    ' It returns FALSE if V is any one of the following variable
    ' types:
    '    vbArray
    '    vbDataObject
    '    vbObject
    '    vbUserDefinedType
    '    or if it is an array of any type.
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error Resume Next
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Test if V is an array. We can't just use VarType(V) = vbArray
    ' because the VarType of an array is vbArray + VarType(type
    ' of array element). E.g, the VarType of an Array of Longs is
    ' 8195 = vbArray + vbLong.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If IsArray(v) = True Then
        IsSimpleDataType = False
        Exit Function
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' We must also explicitly check whether V is an object, rather
    ' relying on VarType(V) to equal vbObject. The reason is that
    ' if V is an object and that object has a default proprety, VarType
    ' returns the data type of the default property. For example, if
    ' V is an Excel.Range object pointing to cell A1, and A1 contains
    ' 12345, VarType(V) would return vbDouble, the since Value is
    ' the default property of an Excel.Range object and the default
    ' numeric type of Value in Excel is Double. Thus, in order to
    ' prevent this type of behavior with default properties, we test
    ' IsObject(V) to see if V is an object.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If IsObject(v) = True Then
        IsSimpleDataType = False
        Exit Function
    End If
    '''''''''''''''''''''''''''''''''''''
    ' Test the value returned by VarType.
    '''''''''''''''''''''''''''''''''''''
    Select Case VarType(v)
        Case vbArray, vbDataObject, vbObject, vbUserDefinedType
            '''''''''''''''''''''''
            ' not simple data types
            '''''''''''''''''''''''
            IsSimpleDataType = False
        Case Else
            ''''''''''''''''''''''''''''''''''''
            ' otherwise it is a simple data type
            ''''''''''''''''''''''''''''''''''''
            IsSimpleDataType = True
    End Select
End Function

Public Function IsArrayAllocated(ByRef arr As Variant) As Boolean
    On Error Resume Next
    IsArrayAllocated = IsArray(arr) And (Not IsError(LBound(arr, 1))) And LBound(arr, 1) <= UBound(arr, 1)
End Function
