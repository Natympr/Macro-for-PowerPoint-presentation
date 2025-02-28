Function OpenWorkbookFromCell(iCellAddress As String) As Workbook
    
    Dim extractedFileName As String
    Dim excelFilePath As String
    Dim excelFileName As String
    Dim excelWorkbook As Workbook
    
    ' Get the file name from the specified cell
    extractedFileName = ThisWorkbook.Sheets(1).Range(iCellAddress).Value
    
    If extractedFileName = "" Then
        MsgBox "No file name found in cell " & iCellAddress & ".", vbExclamation
        Exit Function ' Exit the function without returning a Workbook object
    End If
    
    ' Get the Excel file path
    excelFilePath = ThisWorkbook.Path & "\"
    
    ' Find and open the Excel file
    excelFileName = Dir(excelFilePath & extractedFileName)
    
    If excelFileName = "" Then
        MsgBox "Excel file not found (" & extractedFileName & ").", vbExclamation
        Exit Function ' Exit the function
    End If
    
    Set excelWorkbook = Workbooks.Open(excelFilePath & excelFileName)
    Set OpenWorkbookFromCell = excelWorkbook ' Return the opened workbook
    
End Function

'============================================================================================================

Function SetSlideTitle(pptPresentation As Object, iSlideNumber As Integer, iCellAddress As String)
    
    Dim pptSlide As Object

    Dim pslideTitle As String

    Set pptSlide = pptPresentation.Slides(iSlideNumber)
    pslideTitle = ThisWorkbook.Sheets(2).Range(iCellAddress).Value
    
    With pptSlide.Shapes(1).TextFrame.TextRange
        .Text = pslideTitle
        .Font.Size = 20
        .Font.Bold = True
    End With
    
    
End Function
'============================================================================================================

Function CopyCharts(excelSheet As Worksheet) As Boolean
    Dim numCharts As Integer
    Dim i As Integer
    
    ' Count the number of charts on the sheet
    numCharts = excelSheet.ChartObjects.Count
    
    ' Check if there are charts to select
    If numCharts = 0 Then
        MsgBox "There are no charts on the sheet.", vbExclamation, "Notice"
        CopyCharts = False
        Exit Function
    End If
    
    ' Select all charts
    For i = 1 To numCharts
        If i = 1 Then
            excelSheet.ChartObjects(i).chart.Parent.Select
        Else
            excelSheet.ChartObjects(i).chart.Parent.Select False
        End If
    Next i
    
    ' Copy the selected charts
    Selection.Copy
    
    ' Return true if the copy was successful
    CopyCharts = True
End Function

