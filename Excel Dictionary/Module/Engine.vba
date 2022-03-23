Option Explicit

Dim excelFilePath As String
Dim datarange1 As Variant
Public dataMappings As Scripting.Dictionary

Public Sub runEngine()
    Dim checkCell As Range
    Dim strFileExists As String
    'turn off excel default behaviour
    Call defaultBehaviourOff
    
    Set dataMappings = New Scripting.Dictionary
    
    'check if mandatory fields have been populated - if not exit the script.
    
    'File Path
    Set checkCell = ThisWorkbook.Worksheets("Engine").Cells(6, 4)
    If IsEmpty(checkCell) Then
        MsgBox "File Path required (" & checkCell.Address & ")."
        Exit Sub
    Else
        strFileExists = Dir(ThisWorkbook.Worksheets("Engine").Cells(6, 4))
        If strFileExists = "" Then
            MsgBox "file Path doesn't exist (" & checkCell.Value & ")."
            Exit Sub
        End If
    End If
    
    'Sheet name
    Set checkCell = ThisWorkbook.Worksheets("Engine").Cells(7, 4)
    If IsEmpty(checkCell) Then
        MsgBox "Sheet name required (" & checkCell.Address & ")."
        Exit Sub
    End If
    
    'Product
    Set checkCell = ThisWorkbook.Worksheets("Engine").Cells(11, 4)
    If IsEmpty(checkCell) Then
        MsgBox "Product Column number required (" & checkCell.Address & ")."
        Exit Sub
    End If
    
    'read source document and populate dictionary
    readExcelFile
    
    If dataMappings Is Nothing Then
        Debug.Print "file not found"
    Else
        'create new sheet
        createDataSheet
        'populate new sheet with hard coded headers
        populateDataHeaders
        'populate new sheet with dictionary data
        populateNewSheet
        'autofit excel columns
        AutoFitColumns
        datarange1 = Empty
    End If
    
    Call defaultBehaviourOn
    
    Reset

End Sub

Private Function readExcelFile()
    Dim oXL As Object
    Dim oWB As Workbook
    Dim oWS As Worksheet
    Dim i As Long
    Dim strFileExists As String
    Dim excelFilePath As String
    Dim wsName As String
    'user inputs
    Dim uProduct As String
    Dim iProduct As Integer
    Dim uYear198990 As String
    Dim iYear198990 As String
    Dim dYear198990 As String
    Dim uYear201819 As String
    Dim iYear201819 As Integer
    Dim dYear201819 As String
    Dim last_row As Integer
    Dim last_column As Integer
    
    'user inputs
    uProduct = Sheets("Engine").Cells(11, 4)
    iProduct = userValueCheck(UCase(uProduct))
    
    uYear198990 = Sheets("Engine").Cells(12, 4)
    iYear198990 = userValueCheck(UCase(uYear198990))
    
    uYear201819 = Sheets("Engine").Cells(13, 4)
    iYear201819 = userValueCheck(UCase(uYear201819))
    
    excelFilePath = Sheets("Engine").Cells(6, 4)
    wsName = Sheets("Engine").Cells(7, 4)
    '----
    
    'check if file exists
    strFileExists = Dir(excelFilePath)
    
    If strFileExists = "" Then
        'exit function if file not found
        Exit Function
    Else
    End If
    
    If IsEmpty(datarange1) Then
        Set oXL = CreateObject("excel.application")
        Set oWB = oXL.Workbooks.Open(excelFilePath, UpdateLinks:=False, ReadOnly:=True)
        Set oWS = oWB.Sheets(wsName)
        oXL.Visible = False
        'read excel sheet
        last_row = oWS.UsedRange.Rows.Count
        last_column = oWS.UsedRange.Columns.Count
        datarange1 = oWS.Range(oWS.Cells(1, 1), oWS.Cells(last_row, last_column))
        'close excel and release the application variables
        oWB.Close (False)
        oXL.Application.Quit
        Set oWB = Nothing
        Set oXL = Nothing
        
        'set default compare mode.
        'binary compare = Upper and lower case are different. this is the default option. Value and value are not the same
        'TextCompare = Upper and lower case are identical. Value and value are the same.
        dataMappings.CompareMode = CompareMethod.BinaryCompare
        
        For i = LBound(datarange1) + 1 To UBound(datarange1)
            If Len(datarange1(i, iProduct)) > 0 Then
                
                'optional - manipulation of the variable data - assign updated data to 'd' custom values
                If iYear198990 = -1 Then
                    dYear198990 = "-"
                Else
                    dYear198990 = datarange1(i, iYear198990)
                End If
                
                If iYear201819 = -1 Then
                    dYear201819 = "-"
                Else
                    dYear201819 = datarange1(i, iYear201819)
                End If
                
                newData datarange1(i, iProduct), dYear198990, dYear201819
                
            End If
            
        Next i
        
    Else
        Exit Function
    End If
    
End Function

Private Sub newData(ByVal pProduct As String, ByVal pYear198990 As String, ByVal pYear201819 As String)
    Dim cDataMap As clsDataMap
    'ensure unique values in dictionary - check if exists - else add to dictionary
    If dataMappings.Exists(pProduct) Then
        'do nothing
    Else
        'assign the class
        Set cDataMap = New clsDataMap
        cDataMap.SetAll pProduct, pYear198990, pYear201819
        'add to dictionary
        dataMappings.Add key:=pProduct, Item:=cDataMap
        'clear variable
        Set cDataMap = Nothing
    End If

End Sub

Private Sub populateNewSheet()
    Dim key As Variant
    Dim i As Integer
    i = 2
    If worksheetExists("dataMap") Then
        For Each key In dataMappings.Keys
            With dataMappings.Item(key)
                Sheets("dataMap").Cells(i, 1) = .product
                Sheets("dataMap").Cells(i, 2) = .year198990
                Sheets("dataMap").Cells(i, 3) = .year201819
            End With
            i = i + 1
        Next key
    End If
End Sub


Function userValueCheck(ByVal userField As String) As Integer
    Dim result As Integer
    
    If userField = "" Then
        result = -1
    Else
        If IsNumeric(userField) = True Then
            result = CInt(userField)
        Else
            result = CInt(ConvertLetterToNumber(userField))
        End If
    End If
    
    userValueCheck = result
    
End Function

Function ConvertLetterToNumber(ByVal strSource As String) As String
    Dim i As Integer
    Dim strResult As String
    
    For i = 1 To Len(strSource)
        Select Case Asc(Mid(strSource, i, 1))
            Case 65 To 90:
                strResult = strResult & Asc(Mid(strSource, i, 1)) - 64
            Case Else
                strResult = strResult & Mid(strSource, i, 1)
        End Select
    Next
    ConvertLetterToNumber = strResult

End Function

Private Sub createDataSheet()
    If worksheetExists("dataMap") Then
        Sheets("dataMap").Delete
        createDataSheet
    Else
        Sheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)).Name = "dataMap"
    End If
End Sub

Function worksheetExists(worksheetName As String) As Boolean
    Dim wb As Workbook
    Set wb = ThisWorkbook
    With wb
        On Error Resume Next
        worksheetExists = (.Sheets(worksheetName).Name = worksheetName)
        On Error GoTo 0
    End With
End Function


Private Sub populateDataHeaders()
    If worksheetExists("dataMap") Then
        Sheets("dataMap").Cells(1, 1) = "Product"
        Sheets("dataMap").Cells(1, 1).Interior.ColorIndex = 6
        
        Sheets("dataMap").Cells(1, 2) = "1989-90"
        Sheets("dataMap").Cells(1, 2).Interior.ColorIndex = 6
        
        Sheets("dataMap").Cells(1, 3) = "2018-19"
        Sheets("dataMap").Cells(1, 3).Interior.ColorIndex = 6
        
    End If
End Sub
Private Sub AutoFitColumns()
    If worksheetExists("dataMap") Then
        Sheets("dataMap").Columns(1).AutoFit
        Sheets("dataMap").Columns(2).AutoFit
        Sheets("dataMap").Columns(3).AutoFit
    End If
End Sub

Sub defaultBehaviourOn()
  ' turn on excel defaults
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Sub defaultBehaviourOff()
    ' disable excel defaults to speed up processing
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True
End Sub
