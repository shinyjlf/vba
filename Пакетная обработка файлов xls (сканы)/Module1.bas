Attribute VB_Name = "Module1"
Option Explicit

Sub Log(fileName As String, text As String)

    With ThisWorkbook.Sheets("Лог")
        Dim nextRow As Integer
        If IsEmpty(.Cells(1, 1)) Or IsEmpty(.Cells(2, 1)) Then
            .Cells(1, 1) = "Дата и время"
            .Cells(1, 2) = "Файл"
            .Cells(1, 3) = "Действие / описание ошибки"
            .range("A1:C1").Interior.ColorIndex = 34 'светло голубой
            nextRow = 2
        Else
            nextRow = .Cells(1, 1).End(xlDown).Offset(1, 0).row
        End If
        
        .Cells(nextRow, 1) = Date & " " & Time
        .Cells(nextRow, 2) = fileName
        .Cells(nextRow, 3) = text
          'Application.StatusBar = fileName & " " & text
    End With
End Sub
Sub LogVocabulary(key As Variant, newVocabulary As Object)
    'Debug.Print "my key: ", key
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Найденные новые слова")
    Dim LastRow As Long
    Dim key_cell As range, key_column As Integer
         
     If IsEmpty(ws.Cells(1, 1)) Then
         ws.Cells(1, 1) = "Методика"
         ws.Cells(1, 2) = "ОИ"
         ws.Cells(1, 3) = "Показатель"
         ws.Cells(1, 4) = "Диапазон"
         ws.range("A1:D1").Interior.ColorIndex = 34 'светло голубой
         'LastRow = 1
    End If
    
         Set key_cell = ws.range("A1:E1").Find(key)
         If Not key_cell Is Nothing Then
            key_column = key_cell.column
         Else
            Exit Sub
         End If
    
         'Dim ColumnLetter As String: ColumnLetter = Split(ws.Cells(1, key_column).Address, "$")(1)
         'LastRow = ws.Cells(ws.Rows.Count, ColumnLetter).End(xlUp).row
     LastRow = ws.Cells(1, 1).End(xlDown).row
    
     Dim i As Long
     Dim numberOfItemsToAdd As Long: numberOfItemsToAdd = newVocabulary.Count - (LastRow - 1)
     If numberOfItemsToAdd > 0 Then
         For i = LastRow To LastRow + numberOfItemsToAdd - 1
             ws.Cells(i + 1, clm) = "'" + newVocabulary(i - 1)
         Next
     End If
    
End Sub
        
Function in_array(my_array, my_value)
    Dim i As Long
    in_array = False
    
    For i = LBound(my_array) To UBound(my_array)
        If my_array(i) = my_value Then 'If value found
            in_array = True
            Exit For
        End If
    Next
End Function

Function MultipleSearch(ByVal key As String, rng As range, value As String, Optional isReplace As Boolean) As range
    
    Dim rNew As range
    Dim rCell As range
        
    If (key = "Номер") Then
        Dim counter As Long
        Dim idx As Long
        Dim prevCellValue As Long
        'For j = 1 To rng.Cells.Count
        For Each rCell In rng.Cells
           Debug.Print rCell.Address, rCell
           If rCell <> "" Then
           
           If IsNumeric(rCell) Then
               If (counter = 0) Then
                    counter = rCell
               Else
                    counter = counter + 1
                    If counter <> rCell And counter - 1 <> rCell Then
                        If prevCellValue = rCell - 1 Then ' it was not occasional cell value
                            counter = rCell ' resetting counter
                        Else
                            Call addCellToRange(rNew, rCell)
                        End If
                         
                         prevCellValue = rCell
                    End If
                End If
           Else
                        counter = counter + 1
                         Call addCellToRange(rNew, rCell)
           End If
           End If
        Next
    Else
        Dim arr() As String: arr = Split(value, "mySuperSeparator")
        Dim a As Long
        For a = LBound(arr) To UBound(arr)
            If isReplace Then
                Dim keyS As String: keyS = Split(arr(a), "mySuperKeyValueSeparator")(0)
                Dim valS As String: valS = Split(arr(a), "mySuperKeyValueSeparator")(1)
                ' oRange.Replace What:="sht", Replacement:=ChrW(1097), MatchCase:=True
                rng.Replace What:=keyS, Replacement:=valS  '(keyS, valS)
            Else
                Set rCell = rng.Find(arr(a), MatchCase:=True)
                If Not rCell Is Nothing Then
                    Dim firstCellAddress As String: firstCellAddress = rCell.Address
                    Do
                        Call addCellToRange(rNew, rCell)
                        Set rCell = rng.FindNext(rCell)
                        If rCell Is Nothing Then
                            Exit Do
                        End If
                    Loop While (firstCellAddress <> rCell.Address)
                End If
            End If
            
            
        Next
    End If
    
    Set MultipleSearch = rNew
End Function

Function array_intersection(cellValues() As String, vocabulary() As String, ByRef newVocabulary) As Boolean
    
    Dim intersection As Boolean: intersection = False
    Dim item
    Dim list1: Set list1 = CreateObject("System.Collections.ArrayList")
    Dim list2: Set list2 = CreateObject("System.Collections.ArrayList")
    Dim list3: Set list3 = CreateObject("System.Collections.ArrayList")
     
    For Each item In cellValues
        If Not Trim(item) = "" Then
            list1.Add item
        End If
    Next
 
    For Each item In vocabulary
        list2.Add LCase(item)
    Next
    
    For Each item In newVocabulary
        list3.Add LCase(item)
    Next
 
    For Each item In list1
        If list2.Contains(LCase(item)) Then
            intersection = True
        Else
            If Not list3.Contains(LCase(item)) Then
               list3.Add (LCase(item))
               newVocabulary.Add item '
            End If
        End If
    Next
 
    
     
    array_intersection = intersection
 
End Function
Function addCellToRange(ByRef rng1 As range, rng2 As range)
    If rng1 Is Nothing Then
        Set rng1 = rng2
    Else
        Set rng1 = Union(rng1, rng2)
    End If
End Function
Function MultipleSearchByWords(key As Variant, rng As range, separators_ As String, vocabulary_ As String, ByRef newVocabulary As Object)
    Dim rNew As range
    Dim separators() As String: separators = Split(separators_, "mySuperSeparator")
    Dim vocabulary() As String: vocabulary = Split(vocabulary_, "mySuperSeparator")
    Dim cellValue As String
    Dim cellValues() As String
    Dim cell As range
    Dim i As Long
    
    Select Case key
    Case Is = "Номер"
    
        Dim rng1 As Variant
        If rng.Cells.Count > 1 Then
            'rng1 = Application.WorksheetFunction.Transpose(rng.Value2)
            'For i = LBound(rng1) + 1 To UBound(rng1)
            Dim prevCellValue As String
            For Each cell In rng
                If prevCellValue <> "" And Not IsEmpty(prevCellValue) And prevCellValue = cell.Value2 Then
                    Call addCellToRange(rNew, cell)
                End If
                prevCellValue = cell.Value2
            Next
        End If
        
    Case Else
        Dim rCell As range
        For Each rCell In rng
        If Not IsEmpty(rCell) Then
            cellValue = rCell.Value2
            For i = LBound(separators) To UBound(separators)
                cellValue = Replace(cellValue, separators(i), "mySuperSeparator")
            Next
            cellValues = Split(cellValue, "mySuperSeparator")
            If Not array_intersection(cellValues, vocabulary, newVocabulary) Then
                Call addCellToRange(rNew, rCell)
            End If
        End If
        Next
        
    End Select
    
    
    
    
    Set MultipleSearchByWords = rNew
End Function
           

Function UnSelectSomeCells(arrUnselect() As Integer, rng As range) As range
    Dim rSelect As range
    Dim rUnSelect As range
    Dim rNew As range
    Dim rCell As range
    
    Dim column As Integer

    Set rSelect = rng
    column = rng.Columns(1).column
    
    'Set rUnSelect = Application.InputBox("What cells do you want to exclude?", Type:=8)

    For Each rCell In rSelect
        'If Intersect(rCell, rUnSelect) Is Nothing Then
         If Not in_array(arrUnselect, rCell.row) Then
             Call addCellToRange(rNew, rCell)
         End If
    Next
    Set UnSelectSomeCells = rNew
    rNew.Select

    Set rCell = Nothing
    Set rSelect = Nothing
    Set rUnSelect = Nothing
    'Set rNew = Nothing
End Function

Function UnSelectMergedCells(rng As range) As range
     Dim rNew As range
     Dim rCell As range
    'If rng.MergeCells Then ' true if any cell is merged
        For Each rCell In rng
            If rCell.MergeCells = False Or (rCell.MergeCells And rCell.Address = rCell.MergeArea.Cells(1).Address) Then
                 Call addCellToRange(rNew, rCell)
            End If
        Next
    'Else
       ' Set UnSelectMergedCells = rng
       ' Exit Function
    'End If
    
    
    Set UnSelectMergedCells = rNew
End Function



Sub ProcessWorkbook() 'wb As Workbook, dict As Dictionary, buttonNumber As Integer)
    Dim wb As Workbook: Set wb = Workbooks.Open(fileName:="C:\Users\iri inc\source\repos\vba-afterScan\100 примеров\test\Тест1-1.xlsx")
    Dim buttonNumber As Integer: buttonNumber = 1
    Dim dict As Dictionary: Set dict = Settings(buttonNumber)
    'On Error GoTo Metka
    
    
    
    If buttonNumber = 1 Then
    Cells.Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    End If
        
  Dim uRange As range: Set uRange = wb.Sheets(1).UsedRange
  uRange.Select
    
  Dim i As Integer
  
  
  
  
  Dim emptyR() As Integer: ReDim emptyR(0)
  Dim isRowEmpt As Boolean
  Dim enumeratedR() As Integer: ReDim emptyR(0)
  Dim mergedR() As Integer: ReDim emptyR(0)
  Dim headerR() As Integer: ReDim emptyR(0)
  Dim lastRegionRow As Integer
  Dim columnRange As range
  
  
  
  For i = 1 To uRange.Rows.Count
  
    isRowEmpt = isRowEmpty(i)
    If isRowEmpt Then
          emptyR(UBound(emptyR)) = i
          Debug.Print "row " & emptyR(UBound(emptyR)) & "rowType:empty"
          ReDim Preserve emptyR(UBound(emptyR) + 1)
    End If
    
    If isRowEmpt = False And isRowEnumerated(i) Then
          enumeratedR(UBound(enumeratedR)) = i
          Debug.Print "row " & enumeratedR(UBound(enumeratedR)) & "rowType:enumerated"
          ReDim Preserve enumeratedR(UBound(enumeratedR) + 1)
    End If
    
    If isRowEmpt = False And isRowMerged(i) Then
          mergedR(UBound(mergedR)) = i
          Debug.Print "row " & mergedR(UBound(mergedR)) & "rowType:merged"
          ReDim Preserve mergedR(UBound(mergedR) + 1)
    End If
    
    Dim header As Dictionary: Set header = New Dictionary
    Set header = isRowHeader(i, dict)
    If Not header Is Nothing Then
        headerR(UBound(headerR)) = i
        Debug.Print "row " & headerR(UBound(headerR)) & "rowType:header"
        ReDim Preserve headerR(UBound(headerR) + 1)
    End If
    
  Next
   lastRegionRow = uRange.Rows.Count 'допущение, что выделение начато с первой строки листа
   Dim header1 As Dictionary
    
   For i = LBound(headerR) + 1 To UBound(headerR)
        Dim firstRgRow As Integer
        Dim lastRgRow As Integer
        Debug.Print i & ")  UBound :" & UBound(headerR)
        
        firstRgRow = headerR(i - 1)
        If i < UBound(headerR) Then
             lastRgRow = headerR(i)
        Else
             lastRgRow = lastRegionRow + 1
        End If
        If lastRgRow - firstRgRow > 1 Then
           
                
                
                uRange.Select
                Set header1 = New Dictionary
                Set header1 = isRowHeader(firstRgRow, dict)
                Dim key As Variant
                For Each key In header1.keyS
                    Dim col As Integer: col = header1(key)
                    Dim columnRange_exlMerged As range
                    
                    Set columnRange = range(Cells(firstRgRow + 1, col), Cells(lastRgRow - 1, col))
                    columnRange.Select
                                
                    Set columnRange = UnSelectSomeCells(emptyR, columnRange)
                    Set columnRange = UnSelectSomeCells(enumeratedR, columnRange)
                    Set columnRange = UnSelectSomeCells(mergedR, columnRange)
                    
                    If buttonNumber = 1 Then
                        Set columnRange_exlMerged = UnSelectMergedCells(columnRange)
                        
                        If columnRange_exlMerged.Count = 1 Then
                            If IsEmpty(columnRange_exlMerged.Cells(1)) Then
                                columnRange_exlMerged.Cells(1).Interior.ColorIndex = 6 'желтый
                            End If
                        Else
                                On Error Resume Next
                                 With columnRange_exlMerged.SpecialCells(xlCellTypeBlanks) 'handles 8000 rows maximum
                                     .Interior.ColorIndex = 6 'желтый
                                 End With
                                On Error GoTo Metka
                        End If
                                 
                        Dim newVocabulary As Object: Set newVocabulary = CreateObject("System.Collections.ArrayList")
                        Dim item As Variant
                        For Each item In Split(dict("Найденные новые слова").item(key), "mySuperSeparator")
                            newVocabulary.Add item
                        Next
                        Dim yellowRange As range
                        Set yellowRange = MultipleSearchByWords(key, columnRange_exlMerged, dict("Символы разделители слов").item(key), dict("Словари").item(key), newVocabulary)
                        If Not yellowRange Is Nothing Then
                            yellowRange.Interior.ColorIndex = 44 'желтый
                        End If
                        
                        Debug.Print "**** " & key, dict("Красные символы").item(key)
                        Dim redRange As range
                        Set redRange = MultipleSearch(key, columnRange, dict("Красные символы").item(key))
                        If Not redRange Is Nothing Then
                            redRange.Interior.ColorIndex = 3 'красный
                        End If
                        
                        Debug.Print "**-** " & key
                        'If (Not newVocabulary Is Empty) And (newVocabulary.Count > 0) Then
                            Call LogVocabulary(key, newVocabulary)
                        'End If
                    Else
                        
                        'Debug.Print "**** " & key, dict("Красные символы").item(key)
                        Dim replacedRange As range
                        Set replacedRange = MultipleSearch(key, columnRange, dict("Замены").item(key), True)
                        
                    End If
        
        Next
        
        range("A1").Select
        Set header1 = Nothing
        End If
   Next
             
Exit Sub
Metka:
  Call Log(wb.FullName, "error" & " " & Err.Number & " " & Err.Description)
  Err.Clear
  'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Sub Макрос1(buttonNumber As Integer)
Attribute Макрос1.VB_ProcData.VB_Invoke_Func = "g\n14"

    Dim dict As Dictionary ': Set dict = New Dictionary
    Set dict = Settings(buttonNumber)
   
    Dim wb As Workbook
    Dim myPath As String: myPath = ThisWorkbook.Sheets("Найденные новые слова").range("G2").Value2
    
 If Dir(myPath, vbDirectory) = vbNullString Then
    MsgBox "Некорректный путь к каталогу введен в ячейке G2!"
    Exit Sub
 End If
    
    'Optimize Macro Speed
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Application.Calculation = xlCalculationManual
  Application.DisplayAlerts = False
    
    Dim myExtension As String: myExtension = "*.xls*"
    Dim myFile As String: myFile = Dir(myPath & myExtension) 'Target Path with Ending Extention
    On Error Resume Next
  Do While myFile <> "" 'vbNullString 'Loop through each Excel file in folder
    On Error Resume Next
            'Set wb = Workbooks(myFile)
            'wb.Close SaveChanges:=False
            'DoEvents 'Ensure Workbook has closed before moving on to next line of code
      Set wb = Workbooks.Open(fileName:=myPath & myFile)
      DoEvents 'Ensure Workbook has opened before moving on to next line of code
     If Err.Number = 0 And Not wb Is Nothing Then
                Call Log(myFile, "opened")
             On Error GoTo Metka
                Call ProcessWorkbook(wb, dict, buttonNumber)
             On Error GoTo 0
      Else
            Call Log(myFile, "error" & " " & Err.Number & " " & Err.Description)
            Err.Clear
            On Error GoTo 0
      End If
    
Metka:
        On Error GoTo Metka1
          wb.Close SaveChanges:=True
          DoEvents 'Ensure Workbook has closed before moving on to next line of code
          Call Log(myFile, "closed")
        On Error GoTo 0
Metka1:
      myFile = Dir '(myPath & myExtension) 'Get next file name
  Loop


'ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    On Error GoTo 0
  Exit Sub
  
'Metka2:
   '         Call Log(myFile, "error" & " " & Err.Number & " " & Err.Description)
    '        Err.Clear
     '       On Error GoTo 0
End Sub
Sub Обзор()

    Dim FldrPicker As FileDialog: Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)
    Dim myPath As String
    
    With FldrPicker
          .Title = "Select A Target Folder"
          .AllowMultiSelect = False
            If .Show <> -1 Then Exit Sub
            myPath = .SelectedItems(1) & "\"
    End With
    ThisWorkbook.Sheets("Найденные новые слова").range("G2").Value2 = myPath
 
End Sub

Sub кнопка1()
Attribute кнопка1.VB_ProcData.VB_Invoke_Func = "Q\n14"
        Call Макрос1(1)
End Sub

Sub кнопка2()
Attribute кнопка2.VB_ProcData.VB_Invoke_Func = "W\n14"
        Call Макрос1(2)
End Sub
