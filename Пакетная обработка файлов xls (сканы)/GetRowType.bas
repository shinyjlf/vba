Attribute VB_Name = "GetRowType"
Option Explicit

Function isHeaderCell(value As String, cell As String) As Boolean

        Dim arr() As String
        Dim a As Integer
        Dim position As Long
        isHeaderCell = False
        
        arr = Split(value, "mySuperSeparator")
        For a = LBound(arr) To UBound(arr)
            position = InStr(1, cell, arr(a), vbTextCompare)
            If position <> 0 Then
             'Call Log(arr(a), cell)
             isHeaderCell = True
            Else
             isHeaderCell = False
             Exit Function
            End If
        Next
        
End Function


Function isRowEmpty(rowId As Integer) As Boolean
    'Set row = Selection.Rows(rowId)
    'If Application.CountA(row) = 0 Then
     'isRowEmpty = True
     'Exit Function
    'End If
    'isRowEmpty = False
    
    Dim j As Long
    For j = 1 To Selection.Columns.Count
       If Cells(rowId, j) <> "" Then
       isRowEmpty = False
       Exit Function
     End If
    Next
    isRowEmpty = True
    
End Function

Function isRowEnumerated(row As Integer) As Boolean
    isRowEnumerated = False
    Dim counter As Long, j As Long

    For j = 1 To Selection.Columns.Count
       If Cells(row, j) <> "" And IsNumeric(Cells(row, j)) And CLng(Cells(row, j) < 15) Then
           counter = counter + 1
       End If
    Next
    
    If counter > Selection.Columns.Count / 2 Then
     isRowEnumerated = True
    End If
 
End Function

Function isRowMerged(row As Integer) As Boolean
    isRowMerged = False
    Dim counter As Long, j As Long

    For j = 1 To Selection.Columns.Count
       If Cells(row, j).MergeCells And Cells(row, j).Address = Cells(row, j).MergeArea.Cells(1).Address And Cells(row, j).MergeArea.Columns.Count > 1 Then
            counter = counter + Cells(row, j).MergeArea.Columns.Count
       End If
    Next
    
    If counter > Selection.Columns.Count / 2 Then
     isRowMerged = True
    End If
 
End Function

Function isRowHeader(row As Integer, dict As Dictionary) As Dictionary
    Set isRowHeader = Nothing
    'Dim counter As Integer
    Dim range As range
    Dim key As Variant
    Dim headerDict As Dictionary
    Set headerDict = New Dictionary
    Dim j As Integer
    
    For j = 1 To Selection.Columns.Count
    
    For Each key In dict("Опознавание столбцов").keyS
        
       If isHeaderCell(dict("Опознавание столбцов").item(key), Cells(row, j)) Then
           Debug.Print "Опознавание столбцов : ", key, j
           headerDict.Add key, j
           'counter = counter + 1
           Call addCellToRange(range, Cells(row, j))
           
       End If
    Next
    Next
    
    Set isRowHeader = Nothing
    
    On Error GoTo MetkaH
    If Not range Is Nothing And range.Cells.Count > 1 Then
        Set isRowHeader = headerDict
        range.Interior.ColorIndex = 34 'светло голубой
    End If
    On Error GoTo 0
     
    'If range Is Nothing Or range.Cells.Count < 2 Then
     ' Set isRowHeader = Nothing
    'Else
    '  Set isRowHeader = headerDict
     ' range.Interior.ColorIndex = 34 'светло голубой
    'End If
MetkaH:
End Function
