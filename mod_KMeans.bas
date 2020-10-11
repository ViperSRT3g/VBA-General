Attribute VB_Name = "mod_KMeans"
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : ClusterAnalysis
' Author    : Sheldon Neilson
' Website   : www.neilson.co.za
' Date      : 2011/09/01
' Purpose   : k-Means Cluster Analysis
'---------------------------------------------------------------------------------------

Private Type Records
    Dimension() As Double
    Distance() As Double
    Cluster As Integer
End Type

Dim Table As Range
Dim Record() As Records
Dim Centroid() As Records


Sub Run()
'Run k-Means
    If Not kMeansSelection Then
        Call MsgBox("Error: " & Err.Description, vbExclamation, "kMeans Error")
    End If
End Sub

Function kMeansSelection() As Boolean
'Get user table selection
    On Error Resume Next
    Set Table = Application.InputBox(Prompt:= _
                                     "Please select the range to analyse.", _
                                     Title:="Specify Range", Type:=8)

    If Table Is Nothing Then Exit Function        'Cancelled

    'Check table dimensions
    If Table.Rows.Count < 4 Or Table.Columns.Count < 2 Then
        Err.Raise Number:=vbObjectError + 1000, Source:="k-Means Cluster Analysis", Description:="Table has insufficent rows or columns."
    End If

    'Get number of clusters
    Dim numClusters As Integer
    numClusters = Application.InputBox("Specify Number of Clusters", "k Means Cluster Analysis", Type:=1)

    If Not numClusters > 0 Or numClusters = False Then
        Exit Function        'Cancelled
    End If
    If Err.Number = 0 Then
        If kMeans(Table, numClusters) Then
            outputClusters
        End If
    End If

kMeansSelection_Error:
    kMeansSelection = (Err.Number = 0)
End Function

Function kMeans(Table As Range, Clusters As Integer) As Boolean
'Table - Range of data to group. Records (Rows) are grouped according to attributes/dimensions(columns)
'Clusters - Number of clusters to reduce records into.

    On Error Resume Next

    'Script Performance Variables
    Dim PassCounter As Integer

    'Initialize Data Arrays
    ReDim Record(2 To Table.Rows.Count)
    Dim r As Integer        'record
    Dim d As Integer        'dimension index
    Dim d2 As Integer        'dimension index
    Dim c As Integer        'centroid index
    Dim c2 As Integer        'centroid index
    Dim di As Integer        'distance

    Dim x As Double        'Variable Distance Placeholder
    Dim y As Double        'Variable Distance Placeholder

    For r = LBound(Record) To UBound(Record)
        'Initialize Dimension Value Arrays
        ReDim Record(r).Dimension(2 To Table.Columns.Count)
        'Initialize Distance Arrays
        ReDim Record(r).Distance(1 To Clusters)
        For d = LBound(Record(r).Dimension) To UBound(Record(r).Dimension)
            Record(r).Dimension(d) = Table.Rows(r).Cells(d).Value
        Next d
    Next r

    'Initialize Initial Centroid Arrays
    ReDim Centroid(1 To Clusters)
    Dim uniqueCentroid As Boolean

    For c = LBound(Centroid) To UBound(Centroid)
        'Initialize Centroid Dimension Depth
        ReDim Centroid(c).Dimension(2 To Table.Columns.Count)

        'Initialize record index to next record
        r = LBound(Record) + c - 2

        Do        ' Loop to ensure new centroid is unique
            r = r + 1        'Increment record index throughout loop to find unique record to use as a centroid

            'Assign record dimensions to centroid
            For d = LBound(Centroid(c).Dimension) To UBound(Centroid(c).Dimension)
                Centroid(c).Dimension(d) = Record(r).Dimension(d)
            Next d

            uniqueCentroid = True

            For c2 = LBound(Centroid) To c - 1

                'Loop Through Record Dimensions and check if all are the same
                x = 0
                y = 0
                For d2 = LBound(Centroid(c).Dimension) To _
                    UBound(Centroid(c).Dimension)
                    x = x + Centroid(c).Dimension(d2) ^ 2
                    y = y + Centroid(c2).Dimension(d2) ^ 2
                Next d2

                uniqueCentroid = Not Sqr(x) = Sqr(y)
                If Not uniqueCentroid Then Exit For
            Next c2

        Loop Until uniqueCentroid

    Next c

    'Calculate Distances from Centroids

    Dim lowestDistance As Double
    Dim lastCluster As Integer
    Dim ClustersStable As Boolean

    Do        'While Clusters are not Stable

        PassCounter = PassCounter + 1
        ClustersStable = True        'Until Proved otherwise

        'Loop Through Records
        For r = LBound(Record) To UBound(Record)

            lastCluster = Record(r).Cluster
            lowestDistance = 0        'Reset lowest distance

            'Loop through record distances to centroids
            For c = LBound(Centroid) To UBound(Centroid)

                '======================================================
                '           Calculate Elucidean Distance
                '======================================================
                ' d(p,q) = Sqr((q1 - p1)^2 + (q2 - p2)^2 + (q3 - p3)^2)
                '------------------------------------------------------
                ' X = (q1 - p1)^2 + (q2 - p2)^2 + (q3 - p3)^2
                ' d(p,q) = X

                x = 0
                y = 0
                'Loop Through Record Dimensions
                For d = LBound(Record(r).Dimension) To _
                    UBound(Record(r).Dimension)
                    y = Record(r).Dimension(d) - Centroid(c).Dimension(d)
                    y = y ^ 2
                    x = x + y
                Next d

                x = Sqr(x)        'Get square root

                'If distance to centroid is lowest (or first pass) assign record to centroid cluster.
                If c = LBound(Centroid) Or x < lowestDistance Then
                    lowestDistance = x
                    'Assign distance to centroid to record
                    Record(r).Distance(c) = lowestDistance
                    'Assign record to centroid
                    Record(r).Cluster = c
                End If
            Next c

            'Only change if true
            If ClustersStable Then ClustersStable = Record(r).Cluster = lastCluster

        Next r

        'Move Centroids to calculated cluster average
        For c = LBound(Centroid) To UBound(Centroid)        'For every cluster

            'Loop through cluster dimensions
            For d = LBound(Centroid(c).Dimension) To _
                UBound(Centroid(c).Dimension)

                Centroid(c).Cluster = 0        'Reset nunber of records in cluster
                Centroid(c).Dimension(d) = 0        'Reset centroid dimensions

                'Loop Through Records
                For r = LBound(Record) To UBound(Record)

                    'If Record is in Cluster then
                    If Record(r).Cluster = c Then
                        'Use to calculate avg dimension for records in cluster

                        'Add to number of records in cluster
                        Centroid(c).Cluster = Centroid(c).Cluster + 1
                        'Add record dimension to cluster dimension for later division
                        Centroid(c).Dimension(d) = Centroid(c).Dimension(d) + _
                                                   Record(r).Dimension(d)

                    End If

                Next r

                'Assign Average Dimension Distance
                Centroid(c).Dimension(d) = Centroid(c).Dimension(d) / _
                                           Centroid(c).Cluster
            Next d
        Next c

    Loop Until ClustersStable

    kMeans = (Err.Number = 0)
End Function

Function outputClusters() As Boolean

    Dim c As Integer        'Centroid Index
    Dim r As Integer        'Row Index
    Dim d As Integer        'Dimension Index

    Dim oSheet As Worksheet
    On Error Resume Next

    Set oSheet = addWorksheet("Cluster Analysis", ActiveWorkbook)

    'Loop Through Records
    Dim rowNumber As Integer
    rowNumber = 1

    'Output Headings
    With oSheet.Rows(rowNumber)
        With .Cells(1)
            .Value = "Row Title"
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        With .Cells(2)
            .Value = "Centroid"
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    End With

    'Print by Row
    rowNumber = rowNumber + 1        'Blank Row
    For r = LBound(Record) To UBound(Record)
        oSheet.Rows(rowNumber).Cells(1).Value = Table.Rows(r).Cells(1).Value
        oSheet.Rows(rowNumber).Cells(2).Value = Record(r).Cluster
        rowNumber = rowNumber + 1
    Next r

    'Print Centroids - Headings
    rowNumber = rowNumber + 1
    For d = LBound(Centroid(LBound(Centroid)).Dimension) To UBound(Centroid(LBound(Centroid)).Dimension)
        With oSheet.Rows(rowNumber).Cells(d)
            .Value = Table.Rows(1).Cells(d).Value
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    Next d

    'Print Centroids
    rowNumber = rowNumber + 1
    For c = LBound(Centroid) To UBound(Centroid)
        With oSheet.Rows(rowNumber).Cells(1)
            .Value = "Centroid " & c
            .Font.Bold = True
        End With
        'Loop through cluster dimensions
        For d = LBound(Centroid(c).Dimension) To UBound(Centroid(c).Dimension)
            oSheet.Rows(rowNumber).Cells(d).Value = Centroid(c).Dimension(d)
        Next d
        rowNumber = rowNumber + 1
    Next c

    oSheet.Columns.AutoFit        '//AutoFit columns to contents

outputClusters_Error:
    outputClusters = (Err.Number = 0)
End Function

Function addWorksheet(Name As String, Optional Workbook As Workbook) As Worksheet
    On Error Resume Next
    '// If a Workbook wasn't specified, use the active workbook
    If Workbook Is Nothing Then Set Workbook = ActiveWorkbook
    
    Dim Num As Integer
    '// If a worksheet(s) exist with the same name, add/increment a number after the name
    While WorksheetExists(Name, Workbook)
        Num = Num + 1
        If InStr(Name, " (") > 0 Then Name = Left(Name, InStr(Name, " ("))
        Name = Name & " (" & Num & ")"
    Wend
    
    '//Add a sheet to the workbook
    Set addWorksheet = Workbook.Worksheets.Add
    
    '//Name the sheet
    addWorksheet.Name = Name
End Function

Public Function WorksheetExists(WorkSheetName As String, Workbook As Workbook) As Boolean
    On Error Resume Next
    WorksheetExists = (Workbook.Sheets(WorkSheetName).Name <> "")
    On Error GoTo 0
End Function

