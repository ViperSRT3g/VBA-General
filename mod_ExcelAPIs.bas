Attribute VB_Name = "mod_ExcelAPIs"
Option Explicit

'Updates for this module can be found here:
'https://github.com/ViperSRT3g/General-VBA/blob/main/mod_ExcelAPIs.bas

'Returns the row number of the currently selected cell
Public Function ActiveRow() As Long
    ActiveRow = Application.ActiveCell.Row
End Function

'Returns the column number of the currently selected cell
Public Function ActiveCol() As Long
    ActiveCol = Application.ActiveCell.Column
End Function

'Returns a Range of the current cell executing a UDF
Public Function CurrentCell() As Range
    Set CurrentCell = Application.Caller
End Function

'Returns a boolean if the given cell contains a comment
Public Function HasComment(ByRef TargetCell As Range) As Boolean
    HasComment = Not TargetCell.Comment Is Nothing
End Function

'Returns the last row of the specified worksheet number
Public Function GetLastRow(ByRef TargetWorksheet As Worksheet, ByVal ColumnNo As Variant) As Long
    GetLastRow = TargetWorksheet.Cells(TargetWorksheet.Rows.Count, ColumnNo).End(xlUp).Row
End Function

'Returns the last column of the specified worksheet number
Public Function GetLastCol(ByRef TargetWorksheet As Worksheet, ByVal RowNo As Variant) As Long
    GetLastCol = TargetWorksheet.Cells(RowNo, TargetWorksheet.Columns.Count).End(xlToLeft).Column
End Function

'Returns an expanded range of contiguous cells in the given direction from the target range
Public Function Expand(ByRef Target As Range, ByVal Direction As XlDirection) As Range
    If Not Target Is Nothing Then Set Expand = Target.Parent.Range(Target, Target.End(Direction))
End Function

'Adds a given header to the specified worksheet row, and returns the column number the header occupies
Public Function AddHeader(ByRef TargetWorksheet As Worksheet, ByVal RowNo As Variant, ByVal HeaderName As String) As Long
    TargetWorksheet.Cells(RowNo, TargetWorksheet.Cells(RowNo, TargetWorksheet.Columns.Count).End(xlToLeft).Column + 1).Value = HeaderName
    AddHeader = TargetWorksheet.Cells(RowNo, TargetWorksheet.Columns.Count).End(xlToLeft).Column
End Function

'Returns the Column number of the specified header string, from the specified row of the given worksheet
Public Function GetHeader(ByRef TargetWorksheet As Worksheet, ByVal HeaderRow As Long, ByVal HeaderStr As String) As Long
    Dim Header As Range: Set Header = TargetWorksheet.Rows(HeaderRow).Find(HeaderStr, LookAt:=xlWhole)
    If Not Header Is Nothing Then
        GetHeader = Header.Column
        Set Header = Nothing
    End If
End Function

'Returns a Dictionary of all headers in a given row of a given worksheet with their associated column numbers
'Used in conjunction with the GetHeader function
Public Function GetHeaders(ByRef TargetWorksheet As Worksheet, ByVal HeaderRow As Long, Optional CaseSensitive As Boolean) As Object
    Dim Output As Object: Set Output = CreateObject("Scripting.Dictionary")
    Dim ColCounter As Long
    For ColCounter = 1 To GetLastCol(TargetWorksheet, HeaderRow)
        If CaseSensitive Then 'Headers are untouched
            Output(CStr(TargetWorksheet.Cells(HeaderRow, ColCounter).Value)) = ColCounter
        Else 'Headers are all Uppercase
            Output(UCase(CStr(TargetWorksheet.Cells(HeaderRow, ColCounter).Value))) = ColCounter
        End If
    Next ColCounter
    Set GetHeaders = Output
    Set Output = Nothing
End Function

'Returns a URL within a given cell if it contains one
Public Function GetURL(ByRef Target As Range) As String
    'Grab URL if using the insert link method (Just the first one)
    If Target.Hyperlinks.Count > 0 Then
        GetURL = Target.Hyperlinks.Item(1).Address
        Exit Function
    End If
    
    'Grab URL if using the HYPERLINK formula
    If InStr(1, Target.Formula, "HYPERLINK(""", vbTextCompare) Then
        Dim SLeft As Long: SLeft = InStr(1, Target.Formula, "HYPERLINK(""", vbTextCompare)
        Dim SRight As Long: SRight = InStr(SLeft + 11, Target.Formula, """,""", vbTextCompare)
        GetURL = Mid(Target.Formula, SLeft + 11, SRight - (SLeft + 11))
    End If
End Function

'Returns target cell value of a given workbook as a Variant
Public Function PeekFileCell(ByVal FilePath As String, ByVal FileName As String, ByVal WorksheetName As String, ByVal CellRow As Long, ByVal CellCol As Long) As Variant
    PeekFileCell = ExecuteExcel4Macro("'" & FilePath & "\" & "[" & FileName & "]" & WorksheetName & "'!" & Cells(CellRow, CellCol).Address(1, 1, xlR1C1))
End Function

'Returns a shape object containing the added picture
Public Function AddPicture(ByRef TargetSheet As Worksheet, ByVal Path As String, ByVal Left As Single, ByVal Top As Single, _
                             Width As Single, ByVal Height As Single, Optional ByVal ShapeName As String) As Shape
    Set AddPicture = TargetSheet.Shapes.AddPicture(Path, msoFalse, msoTrue, Left, Top, Width, Height)
    If Len(ShapeName) > 0 Then AddPicture.Name = ShapeName
End Function

'Returns a boolean if a given CheckBox exists with a given name in a given worksheet
Public Function CheckBoxExists(ByVal Name As String, ByRef TargetWorksheet As Worksheet) As Boolean
    If TargetWorksheet Is Nothing Then Set TargetWorksheet = ActiveSheet
    Dim TCB As CheckBox
    For Each TCB In TargetWorksheet.CheckBoxes
        If TCB.Name = Name Then
            CheckBoxExists = True: Set TCB = Nothing: Exit Function
        End If
    Next TCB
    Set TCB = Nothing
End Function

'Returns a boolean if a given shape exists in a given worksheet
Public Function ShapeExists(ByVal Name As String, Optional ByRef TargetWorksheet As Worksheet) As Boolean
    On Error Resume Next
    If TargetWorksheet Is Nothing Then Set TargetWorksheet = ActiveSheet
    ShapeExists = Not TargetWorksheet.Shapes(Name) Is Nothing
End Function


'WORKSHEET FUNCTIONS
'Returns a worksheet with the given name, creates a new one if it doesn't already exist
Public Function GetSheet(ByVal SheetName As String, Optional ByRef WB As Workbook, Optional ForceNew As Boolean) As Worksheet
    On Error Resume Next
    If Len(SheetName) = 0 Then Exit Function
    If WB Is Nothing Then Set WB = ThisWorkbook
    Set GetSheet = WB.Worksheets(Left(SheetName, 31)) 'Test if the given named worksheet exists
    
    If ForceNew Then
        Dim Append As String, MatchCounter As Long
        If Not GetSheet Is Nothing Then 'If the given named worksheet exists, then begin appending the default ' (N)' postfix
            Do Until GetSheet Is Nothing
                Append = " (" & MatchCounter & ")"
                Set GetSheet = Nothing
                Set GetSheet = WB.Worksheets(Left(SheetName, 31 - Len(Append)) & Append)
                MatchCounter = MatchCounter + 1
            Loop
        End If
        Set GetSheet = WB.Worksheets.Add(After:=WB.Worksheets(WB.Worksheets.Count))
        GetSheet.Name = Left(SheetName, 31 - Len(Append)) & Append
    Else
        If GetSheet Is Nothing Then 'If the given name does not exist, create a worksheet with the given name
            Set GetSheet = WB.Worksheets.Add(After:=WB.Worksheets(WB.Worksheets.Count))
            GetSheet.Name = Left(SheetName, 31)
        End If
    End If
End Function

'Returns boolean if a given worksheet exists in a given workbook
Public Function SheetExists(ByVal SheetName As String, Optional ByRef WB As Workbook) As Boolean
    On Error Resume Next
    If WB Is Nothing Then Set WB = ActiveWorkbook
    SheetExists = Not WB.Worksheets(SheetName) Is Nothing
End Function


'WORKBOOK FUNCTIONS
'Returns boolean if a given workbook is password protected
Public Function IsWBProtected(ByRef TWB As Workbook) As Boolean
    IsWBProtected = TWB.ProtectWindows Or TWB.ProtectStructure
End Function

'Returns boolean if a given worksheet is password protected
Public Function IsWSProtected(ByRef TWS As Worksheet) As Boolean
    IsWSProtected = TWS.ProtectContents Or TWS.ProtectDrawingObjects Or TWS.ProtectScenarios
End Function

'Returns boolean if a given workbook is currently open
Public Function IsWorkBookOpen(ByVal WorkbookName As String) As Boolean
    On Error GoTo ErrorHandler
    Dim WBO As Workbook: Set WBO = Workbooks(WorkbookName)
    IsWorkBookOpen = Not WBO Is Nothing
ErrorHandler:
    Set WBO = Nothing
End Function

'Returns a workbook object based on a matching name search
Public Function FindWorkbook(ByVal WorkbookName As String) As Workbook
    Dim Index As Long
    For Index = 1 To Workbooks.Count
        If Workbooks(Index).Name Like "*" & WorkbookName & "*" Then Set FindWorkbook = Workbooks(Index)
    Next Index
End Function

'Returns TRUE if a given workbook reference exists and has not been saved
Public Function WBNotSaved(TargetWB As Workbook) As Boolean
    On Error Resume Next
    If TargetWB Is Nothing Then Exit Function
    If Len(TargetWB.Path) > 0 Then Exit Function
    WBNotSaved = Len(TargetWB.Path) = 0
End Function

'Returns TRUE if a given workbook reference is unused. This indicates that the workbook was unexpectedly closed
Public Function WBNullRef(TargetWB As Workbook) As Boolean
    On Error Resume Next
    If TargetWB Is Nothing Then Exit Function
    If Len(TargetWB.Name) = 0 Then
        WBNullRef = Not (Err.Number = 0)
        Err.Clear
    End If
End Function


'SUBROUTINES
'Unmerges a given range of cells, and fills each cell with the originally merged data
Public Sub UnmergeAndFill(ByRef WorkArea As Range)
    Dim TCell As Range, MAddress As String, MVal As String
    For Each TCell In WorkArea.SpecialCells(xlCellTypeConstants, xlLogical + xlNumbers + xlTextValues).Cells
        If TCell.MergeCells Then
            MAddress = TCell.MergeArea.Address
            TCell.MergeCells = False
            Range(MAddress).Value = TCell.Value
        End If
    Next TCell
    Set TCell = Nothing
End Sub

'Adjusts Excel settings for faster VBA processing
Public Sub LudicrousMode(ByVal Toggle As Boolean)
    Application.ScreenUpdating = Not Toggle
    Application.EnableEvents = Not Toggle
    Application.DisplayAlerts = Not Toggle
    Application.EnableAnimations = Not Toggle
    Application.DisplayStatusBar = Not Toggle
    Application.PrintCommunication = Not Toggle
    Application.Calculation = IIf(Toggle, xlCalculationManual, xlCalculationAutomatic)
End Sub

