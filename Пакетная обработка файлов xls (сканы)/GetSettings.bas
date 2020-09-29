Attribute VB_Name = "GetSettings"
Option Explicit

Function RangeToString(myRange As range) As String
    Dim str As String
    Dim idx As Long
    Dim c As range
    Dim i As Long
    'For Each c In myRange
    For i = 1 To myRange.Rows.Count
        Dim s As String: s = myRange.Cells(i, 1).text
        If (myRange.Columns.Count = 2) Then
            s = s & "mySuperKeyValueSeparator" & myRange.Cells(i, 2).text
        End If
        
        If (myRange.Rows.Count - 1 = idx) Then
            str = str & s
        Else
            str = str & s & "mySuperSeparator"
        End If
        idx = idx + 1
    Next
    
    'Next c

    RangeToString = str
End Function

Function getValueFromRange(ws As Worksheet, columnNumber As Integer, Optional step As Integer) As String
    Dim LastRow As Long
    Dim myRange As range
    Debug.Print "step:", step
    Dim ColumnLetter As String: ColumnLetter = Split(ws.Cells(1, columnNumber).Address, "$")(1)
    Dim ColumnLetter1 As String: ColumnLetter1 = Split(ws.Cells(1, columnNumber + step).Address, "$")(1)
    LastRow = ws.Cells(ws.Rows.Count, ColumnLetter).End(xlUp).row
   ' Debug.Print "step:", step
    If LastRow <> 1 Then
        Set myRange = ws.range(ColumnLetter & (2 + step) & ":" & ColumnLetter1 & LastRow)
        getValueFromRange = RangeToString(myRange)
    Else
        getValueFromRange = ""
    End If
End Function

Function getDictionary(name As String, justOneRow As Boolean) As Dictionary
    Dim dict As Dictionary
    Set dict = New Dictionary
    Dim ws As Worksheet
     Dim i As Integer, key As String, value As String
     
     'Dim wS As Worksheet
    Set ws = ThisWorkbook.Sheets(name)
    Dim myRange As range: Set myRange = ws.UsedRange
    
    If name = "Замены" Then
    
        For i = 1 To myRange.Columns.Count Step 2
            If Not dict.Exists(myRange.Cells(1, i)) Then
             key = myRange.Cells(1, i)
             If Not IsEmpty(key) Then
                    value = getValueFromRange(ws, i, 1)
                 
                    dict.Add key, value
                    'Debug.Print dict(key)
                End If
             End If
        Next
    Else
          For i = 1 To myRange.Columns.Count
             If Not IsEmpty(myRange.Cells(1, i)) And Not dict.Exists(myRange.Cells(1, i)) Then
                key = myRange.Cells(1, i)
             If Not IsEmpty(key) Then
                 If justOneRow = True Then
                    value = Replace(myRange.Cells(2, i), ";", "mySuperSeparator")
                 Else
                    value = getValueFromRange(ws, i)
                 End If
                    dict.Add key, value
                    'Debug.Print dict(key)
             End If
             End If
        Next
    End If
    
    Set getDictionary = dict
    
End Function

Function Settings(buttonNumber As Integer) As Dictionary
    Dim mainDict As Dictionary: Set mainDict = New Dictionary
      
    Dim dict1 As Dictionary: Set dict1 = New Dictionary
    Set dict1 = getDictionary("Опознавание столбцов", True)
    mainDict.Add "Опознавание столбцов", dict1
    'Debug.Print "!!!" & mainDict("Опознавание столбцов").Item("Методика")
    
    If buttonNumber = 1 Then
    
          Dim dict2 As Dictionary: Set dict2 = New Dictionary
          Set dict2 = getDictionary("Словари", False)
          mainDict.Add "Словари", dict2
          'Debug.Print "!!!" & mainDict("Словари").Item("Методика")
        
          Dim dict3 As Dictionary: Set dict3 = New Dictionary
          Set dict3 = getDictionary("Красные символы", False)
          mainDict.Add "Красные символы", dict3
          'Debug.Print "!!!" & mainDict("Красные символы").Item("Методика")
          
          Dim dict4 As Dictionary: Set dict4 = New Dictionary
          Set dict4 = getDictionary("Символы разделители слов", False)
          mainDict.Add "Символы разделители слов", dict4
          'Debug.Print "!!!" & mainDict("Символы разделители слов").Item("Методика")
          
          Dim dict5 As Dictionary: Set dict5 = New Dictionary
          Set dict5 = getDictionary("Найденные новые слова", False)
          mainDict.Add "Найденные новые слова", dict5
          Debug.Print "!!!" & mainDict("Найденные новые слова").item("Методика")
     
     Else
     
        Dim dict6 As Dictionary: Set dict6 = New Dictionary
        Set dict6 = getDictionary("Замены", False)
        mainDict.Add "Замены", dict6
        Debug.Print "!!!" & mainDict("Замены").item("Методика")
        
     End If

    Set Settings = mainDict
  
End Function
