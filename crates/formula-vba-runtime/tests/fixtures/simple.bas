Sub WriteHello()
    Range("A1").Value = "Hello"
    Cells(2, 1).Value = 42
End Sub

Sub FillColumn()
    Dim i
    For i = 1 To 3
        Cells(i, 1).Value = i
    Next i
End Sub

Sub SelectAndWrite()
    Range("B2").Select
    ActiveCell.Value = "X"
End Sub

Sub HandleError()
    On Error GoTo ErrHandler
    Range().Value = 1
    Exit Sub
ErrHandler:
    Range("A1").Value = "handled"
End Sub

Sub Workbook_Open()
    Range("A1").Value = "opened"
End Sub

Sub Worksheet_Change(Target)
    Target.Value = "changed"
End Sub

Sub Worksheet_SelectionChange(Target)
    Target.Value = "selected"
End Sub

Sub ArrayTest()
    Dim arr
    arr = Array(1, 2, 3)
    Cells(1, 1).Value = arr(1)
End Sub

Function AddTwo(x)
    AddTwo = x + 2
End Function

Sub CallFunction()
    Cells(1, 1).Value = AddTwo(40)
End Sub

Sub CollectionTest()
    Dim c
    Set c = New Collection
    c.Add 1
    c.Add 2
    Cells(1, 1).Value = c.Count
    Cells(2, 1).Value = c.Item(2)
End Sub

Sub Infinite()
    Do While True
    Loop
End Sub
