Option Explicit

' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' YouTube video: https://youtu.be/MF0nm5kk1vg
' Description: Sum the Sales and Amount for each fruit and write

Public Sub createReport()
    Dim dict As New Dictionary
    Dim rng As Range
    Dim i As Long
    Dim name As String
    Dim fruit As clsFruit
    
    Set rng = Sheet1.Range("B3").CurrentRegion
    'read the data
    For i = 2 To rng.Rows.Count
        name = rng.Cells(i, 1).Value
        'check if the fruit name already exists in the dictionary
        If dict.Exists(name) = False Then
            Set fruit = New clsFruit
            fruit.name = name
            dict.Add key:=fruit.name, Item:=fruit
        Else
            'if it is already in dictionary set to fruit variable
            Set fruit = dict(name)
        End If
        
        'update the values = add values together
        With fruit
            .sales = .sales + rng.Cells(i, 3).Value
            .amount = .amount + rng.Cells(i, 4).Value
        End With
    Next
    
    'write the data to the worksheet
    Call writeDataToWorkSheet(dict)
    
End Sub

Sub writeDataToWorkSheet(dict As Dictionary)
    Dim rng As Range
    Dim key As Variant
    Dim fruit As clsFruit
    Dim curRow As Long
    
    'get output range
    Set rng = Sheet1.Range("G3")
    'clear the existing contents. use offset to preserve the header information
    rng.CurrentRegion.Offset(1).ClearContents
    
    curRow = 1
    'read through the keys in dictionary
    For Each key In dict.Keys
        Set fruit = dict(key)
        'write values to the worksheet
        With fruit
            rng.Cells(curRow, 1).Value = .name
            rng.Cells(curRow, 2).Value = .sales
            rng.Cells(curRow, 3).Value = .amount
            
            curRow = curRow + 1
            
        End With
    Next key
End Sub