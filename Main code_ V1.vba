Sub ExcelToPowerPoint()
    
    'Application.EnableEvents = False
    'Application.DisplayAlerts = False
    'Application.ScreenUpdating = False

    Dim pptApp As Object
    Dim pptPresentation As Object
    Dim pExcelTable As Range
    Dim excelWorkbook As Workbook
    Dim pptSlide As Object
    Dim slideTitle As Object
    Dim pExcelSheet As Worksheet
    Dim excelFilePath As String
    Dim excelFileName As String
    Dim extractedFileName As String
    Dim pptShape As Object
    Dim powerPointApp As Object
    Dim pptFileName As String
    Dim pptFilePath As String
    Dim text1 As String, text2 As String, text3 As String, title As String
    Dim pNumSlideIndex As Integer
    Dim excelChart As chartObject
    Dim chart As chart
    Dim pExcelTableu As Workbook
    Dim pExcelFranchise As Workbook
    Dim pExcelHAV As Workbook
    Dim pSheetName As String
    Dim pNewPptFileName As String
    Dim pNewPptFilePath As String


    ' Initialize PowerPoint
    On Error Resume Next
    Set powerPointApp = GetObject(, "PowerPoint.Application")
    If Err.Number <> 0 Then
        Set powerPointApp = CreateObject("PowerPoint.Application")
    End If
    On Error GoTo 0
    
    If powerPointApp Is Nothing Then
        MsgBox "Could not open PowerPoint. Make sure it is installed.", vbCritical
        Exit Sub
    End If
    
    powerPointApp.Visible = True
    
    ' Get presentation name and path from cell A3
    pptFileName = ThisWorkbook.Sheets(1).Range("B5").Value
    pptFilePath = ThisWorkbook.Path & "\" & pptFileName
    
    ' Open the presentation
    Set pptPresentation = powerPointApp.Presentations.Open(pptFilePath)


    ' Open the Excel files at the beginning
    Set pExcelTableu = OpenWorkbookFromCell("B2")
    Set pExcelFranchise = OpenWorkbookFromCell("B3")
    Set pExcelHAV = OpenWorkbookFromCell("B4")
    
    
    '    '===========================================SLIDE 1 ================================================================
    '
    ' Create a new text box on the first slide
    Set pptSlide = pptPresentation.Slides(1)
    Set pptShape = pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 400, 100)
    
    ' Get text for the slide
    title = ThisWorkbook.Sheets(2).Range("B2").Value
    text1 = ThisWorkbook.Sheets(2).Range("C2").Value
    text2 = ThisWorkbook.Sheets(2).Range("D2").Value
    
    ' Set each text line in the text box
    With pptShape.TextFrame.TextRange
        .Text = title ' First line
        .Font.Size = 28
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        
        .InsertAfter vbCrLf & vbCrLf & text1
        .Characters(Len(.Text) - Len(text1) + 1, Len(text1)).Font.Size = 9
        .Characters(Len(.Text) - Len(text1) + 1, Len(text1)).Font.Bold = True
        .Characters(Len(.Text) - Len(text1) + 1, Len(text1)).Font.Color = RGB(255, 255, 255)
        
        .InsertAfter vbCrLf & vbCrLf & vbCrLf & text2
        .Characters(Len(.Text) - Len(text2) + 1, Len(text2)).Font.Size = 9
        .Characters(Len(.Text) - Len(text2) + 1, Len(text2)).Font.Bold = True
        .Characters(Len(.Text) - Len(text2) + 1, Len(text2)).Font.Color = RGB(255, 255, 255)
    End With
    
    '======================================================= SLIDE 2 ===================================================================
    

    Set pExcelSheet = pExcelTableu.Sheets(pExcelTableu.Sheets.Count)
    
    Set pExcelTable = pExcelSheet.Range("A1:R27")
    If pExcelTable Is Nothing Then
        MsgBox "No valid range found between rows 1 and 27."

        Exit Sub
    End If
    
    pExcelTable.Copy
    
    'Call function for title
    
    SetSlideTitle pptPresentation, 2, "B3"
    
    
    ' Paste as Excel object in PowerPoint (keeps format and is editable)
    Set pptSlide = pptPresentation.Slides(2)
    pptSlide.Shapes.PasteSpecial DataType : = 10
    
    
    Application.CutCopyMode = False

    ' =========================================================== SLIDE 4 ===============================================================

    
    pExcelFranchise.Charts("Graphique1").CopyPicture Appearance : = xlScreen, Format : = xlPicture

    SetSlideTitle pptPresentation, 4, "B5"
    
    Set pptSlide = pptPresentation.Slides(4)
    pptSlide.Shapes.PasteSpecial DataType : = ppPasteMetafile

    
    
    With pptSlide.Shapes(pptSlide.Shapes.Count)
        .Left = 200
        .Top = 100
        .Height = 420
    End With

    Application.CutCopyMode = False

    ' =========================================================== SLIDE 5 ===============================================================

    pSheetName = ThisWorkbook.Sheets(2).Range("E6").Value

    Set pExcelSheet = pExcelHAV.Sheets(pSheetName)
    pExcelSheet.Activate

    Call CopyCharts(pExcelSheet)
    
    SetSlideTitle pptPresentation, 5, "B6"

    Set pptSlide = pptPresentation.Slides(5)
    pptSlide.Shapes.Paste
    
    Application.CutCopyMode = False

    ' =========================================================== SLIDE 6 ===============================================================
    pSheetName = ThisWorkbook.Sheets(2).Range("E7").Value
    Set pExcelSheet = pExcelHAV.Sheets(pSheetName)
    pExcelSheet.Activate
    
    Call CopyCharts(pExcelSheet)
    
    SetSlideTitle pptPresentation, 6, "B7"
    
    Set pptSlide = pptPresentation.Slides(6)
    pptSlide.Shapes.Paste
    
    Application.CutCopyMode = False
    

    ' =========================================================== SLIDE 7 ===============================================================
    pSheetName = ThisWorkbook.Sheets(2).Range("E8").Value
    Set pExcelSheet = pExcelHAV.Sheets(pSheetName)
    pExcelSheet.Activate

    Call CopyCharts(pExcelSheet)
    
    SetSlideTitle pptPresentation, 7, "B8"
    
    Set pptSlide = pptPresentation.Slides(7)
    pptSlide.Shapes.Paste
    
    Application.CutCopyMode = False

    ' =========================================================== SLIDE 8 ===============================================================
    pSheetName = ThisWorkbook.Sheets(2).Range("E9").Value
    Set pExcelSheet = pExcelHAV.Sheets(pSheetName)
    pExcelSheet.Activate

    Call CopyCharts(pExcelSheet)
    
    SetSlideTitle pptPresentation, 8, "B9"
    
    Set pptSlide = pptPresentation.Slides(8)
    pptSlide.Shapes.Paste
    
    Application.CutCopyMode = False

    ' =========================================================== SLIDE 9 ===============================================================

    pSheetName = ThisWorkbook.Sheets(2).Range("E10").Value
    Set pExcelSheet = pExcelHAV.Sheets(pSheetName)
    pExcelSheet.Activate

    Call CopyCharts(pExcelSheet)
    
    SetSlideTitle pptPresentation, 9, "B10"
    
    Set pptSlide = pptPresentation.Slides(9)
    pptSlide.Shapes.Paste
    
    Application.CutCopyMode = False


    ' =========================================================== SLIDE 10 ===============================================================

    pSheetName = ThisWorkbook.Sheets(2).Range("E11").Value
    Set pExcelSheet = pExcelHAV.Sheets(pSheetName)
    pExcelSheet.Activate

    Call CopyCharts(pExcelSheet)
    
    SetSlideTitle pptPresentation, 10, "B11"
    
    Set pptSlide = pptPresentation.Slides(10)
    pptSlide.Shapes.Paste
    
    Application.CutCopyMode = False


    ' =========================================================== SLIDE 11 ===============================================================

    pSheetName = ThisWorkbook.Sheets(2).Range("E12").Value
    Set pExcelSheet = pExcelFranchise.Sheets(pSheetName)
    pExcelSheet.Activate

    Call CopyCharts(pExcelSheet)
    
    SetSlideTitle pptPresentation, 11, "B12"
    
    Set pptSlide = pptPresentation.Slides(11)
    pptSlide.Shapes.Paste
    
    Application.CutCopyMode = False

    ' =========================================================== SLIDE 12 ===============================================================

    pSheetName = ThisWorkbook.Sheets(2).Range("E13").Value
    Set pExcelSheet = pExcelFranchise.Sheets(pSheetName)
    pExcelSheet.Activate

    Call CopyCharts(pExcelSheet)
    
    SetSlideTitle pptPresentation, 12, "B13"
    
    Set pptSlide = pptPresentation.Slides(12)
    pptSlide.Shapes.Paste
    
    Application.CutCopyMode = False


    ' =========================================================== SLIDE 13 ===============================================================

    pSheetName = ThisWorkbook.Sheets(2).Range("E14").Value
    Set pExcelSheet = pExcelFranchise.Sheets(pSheetName)
    pExcelSheet.Activate

    Call CopyCharts(pExcelSheet)
    
    SetSlideTitle pptPresentation, 13, "B14"
    
    Set pptSlide = pptPresentation.Slides(13)
    pptSlide.Shapes.Paste
    
    Application.CutCopyMode = False


    ' =========================================================== SLIDE 14 ===============================================================

    pSheetName = ThisWorkbook.Sheets(2).Range("E15").Value
    Set pExcelSheet = pExcelFranchise.Sheets(pSheetName)
    pExcelSheet.Activate

    Call CopyCharts(pExcelSheet)
    
    SetSlideTitle pptPresentation, 14, "B15"
    
    Set pptSlide = pptPresentation.Slides(14)
    pptSlide.Shapes.Paste
    
    Application.CutCopyMode = False

    ' =========================================================== SLIDE 15 ===============================================================

    pSheetName = ThisWorkbook.Sheets(2).Range("E16").Value
    Set pExcelSheet = pExcelFranchise.Sheets(pSheetName)
    pExcelSheet.Activate

    Call CopyCharts(pExcelSheet)
    
    SetSlideTitle pptPresentation, 15, "B16"
    
    Set pptSlide = pptPresentation.Slides(15)
    pptSlide.Shapes.Paste
    
    Application.CutCopyMode = False


    ' =========================================================== SLIDE 16 ===============================================================

    pSheetName = ThisWorkbook.Sheets(2).Range("E17").Value
    Set pExcelSheet = pExcelFranchise.Sheets(pSheetName)
    pExcelSheet.Activate

    Call CopyCharts(pExcelSheet)
    
    SetSlideTitle pptPresentation, 16, "B17"
    
    Set pptSlide = pptPresentation.Slides(16)
    pptSlide.Shapes.Paste
    
    Application.CutCopyMode = False


    ' =========================================================== SLIDE 17 ===============================================================

    pSheetName = ThisWorkbook.Sheets(2).Range("E18").Value
    Set pExcelSheet = pExcelFranchise.Sheets(pSheetName)
    pExcelSheet.Activate

    Call CopyCharts(pExcelSheet)
    
    SetSlideTitle pptPresentation, 17, "B18"
    
    Set pptSlide = pptPresentation.Slides(17)
    pptSlide.Shapes.Paste
    
    Application.CutCopyMode = False

    ' =========================================================== SLIDE 18 ===============================================================

    pSheetName = ThisWorkbook.Sheets(2).Range("E19").Value
    Set pExcelSheet = pExcelFranchise.Sheets(pSheetName)
    pExcelSheet.Activate

    Call CopyCharts(pExcelSheet)
    
    SetSlideTitle pptPresentation, 18, "B19"
    
    Set pptSlide = pptPresentation.Slides(18)
    pptSlide.Shapes.Paste
    
    Application.CutCopyMode = False

    ' =========================================================== SLIDE 19 ===============================================================

    pSheetName = ThisWorkbook.Sheets(2).Range("E20").Value
    Set pExcelSheet = pExcelFranchise.Sheets(pSheetName)
    pExcelSheet.Activate

    Call CopyCharts(pExcelSheet)
    
    SetSlideTitle pptPresentation, 19, "B20"
    
    Set pptSlide = pptPresentation.Slides(19)
    pptSlide.Shapes.Paste
    
    Application.CutCopyMode = False

    ' =========================================================== SLIDE 20 ===============================================================

    pSheetName = ThisWorkbook.Sheets(2).Range("E21").Value
    Set pExcelSheet = pExcelFranchise.Sheets(pSheetName)
    pExcelSheet.Activate

    Call CopyCharts(pExcelSheet)
    
    SetSlideTitle pptPresentation, 20, "B21"
    
    Set pptSlide = pptPresentation.Slides(20)
    pptSlide.Shapes.Paste
    
    Application.CutCopyMode = False
    
    
    ' =========================================================== SLIDE 21 ===============================================================

    pSheetName = ThisWorkbook.Sheets(2).Range("E22").Value
    Set pExcelSheet = pExcelFranchise.Sheets(pSheetName)
    pExcelSheet.Activate

    Call CopyCharts(pExcelSheet)
    
    SetSlideTitle pptPresentation, 21, "B22"
    
    Set pptSlide = pptPresentation.Slides(21)
    pptSlide.Shapes.Paste
    
    Application.CutCopyMode = False
    

    '============================================== Add text at the bottom left corner for all slides ============================================
    
    For pNumSlideIndex = 1 To pptPresentation.Slides.Count
        Set pptSlide = pptPresentation.Slides(pNumSlideIndex)
        Set pptShape = pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 500, 100, 60)
        pptShape.TextFrame.TextRange.Text = "Page " & pNumSlideIndex & vbCrLf & Date & vbCrLf & "Confidential"
        pptShape.TextFrame.TextRange.Font.Size = 6
        pptShape.TextFrame.TextRange.Font.Bold = False
        pptShape.TextFrame.TextRange.ParagraphFormat.Alignment = 1
    Next pNumSlideIndex

    '============================================== Closing and saving  ============================================


    ' Close all Excels file without saving changes
    pExcelTableu.Close False
    pExcelFranchise.Close False
    pExcelHAV.Close False


    'Closing and saving PP presentation

    pNewPptFileName = ThisWorkbook.Sheets(1).Range("B6").Value

    pNewPptFilePath = ThisWorkbook.Path & "\" & pNewPptFileName
    
    pptPresentation.SaveAs pNewPptFilePath
    
    pptPresentation.Close
    powerPointApp.Quit
    
    ' Confirmation message
    MsgBox "Table successfully copied to PowerPoint as an Excel object.", vbInformation
    
    'Application.EnableEvents = True
    ' Application.DisplayAlerts = True
    'Application.ScreenUpdating = True

End Sub

