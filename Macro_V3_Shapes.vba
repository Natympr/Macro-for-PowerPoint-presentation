
' EN ESTA VERSION LAS GRAFICAS, TEXTOS, LINEAS SE COPIAN Y PEGAN COMO SHAPES. NO HACEN FALTA LAS FUNCIONES PARA PEGAR CHARTS,
' YA QUE USANDO COMO OBJETO SHAPE SE LOGRA.

Sub ExcelToPowerPoint()

    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Dim powerPointApp As Object
    Dim pptPresentation As Object
    Dim pExcelTable As Range
    Dim pptSlide As Object
    Dim pExcelSheet As Worksheet
    Dim pptFileName As String
    Dim pptFilePath As String
    Dim text1 As String, text2 As String, title As String
    Dim pNumSlideIndex As Integer
    Dim pExcelTableu As Workbook
    Dim pExcelFranchise As Workbook
    Dim pExcelHAV As Workbook
    Dim pSheetName As String
    Dim pNewPptFileName As String
    Dim pNewPptFilePath As String
    Dim i As Integer
    Dim startRow As Integer
    Dim endRow As Integer
             
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
    pptSlide.Shapes.PasteSpecial DataType:=10


    Application.CutCopyMode = False
    
   ' =========================================================== SLIDE 3 ===============================================================
   

   pExcelFranchise.Charts("Graphique1").CopyPicture Appearance:=xlScreen, Format:=xlPicture
    
   SetSlideTitle pptPresentation, 3, "B5"

   Set pptSlide = pptPresentation.Slides(3)
   pptSlide.Shapes.PasteSpecial DataType:=ppPasteMetafile
  


   With pptSlide.Shapes(pptSlide.Shapes.Count)
    .Left = 200
    .Top = 100
    .Height = 420
   End With
   
   Application.CutCopyMode = False
   
' =========================================================== SLIDEs FROM 4 TO 21 ===============================================================
 
  ' Last row with data
  endRow = ThisWorkbook.Sheets(2).Cells(Rows.Count, "E").End(xlUp).Row

  ' initial row for slide (sheet2)
  startRow = 5

  ' Loop to create slides
  For i = startRow To endRow


    pSheetName = ThisWorkbook.Sheets(2).Cells(i, "E").Value

    ' File to use
    If i <= 11 Then ' pExcelHAV file (Filas 5-11)
      Set pExcelSheet = pExcelHAV.Sheets(pSheetName)
    Else 'pExcelFranchise file (Filas 12-22)
      Set pExcelSheet = pExcelFranchise.Sheets(pSheetName)
    End If

    pExcelSheet.Activate
        
'    For Each shapeObj In pExcelSheet.Shapes
'        shapeObj.Copy
'        Set pptSlide = pptPresentation.Slides(i - 1)
'        pptSlide.Shapes.PasteSpecial DataType:=ppPasteMetafile
'    Next shapeObj
    
        
    For Each shapeObj In pExcelSheet.Shapes
    shapeObj.Copy
    DoEvents ' Permite que el sistema procese la copia antes de continuar
    
    SetSlideTitle pptPresentation, i - 1, "B" & i

    ' Selecciona la diapositiva correcta
    Set pptSlide = pptPresentation.Slides(i - 1)
    
    ' Pega el objeto y maneja el error si no se pega
    On Error Resume Next
    Set pastedShape = pptSlide.Shapes.PasteSpecial(DataType:=ppPasteMetafile)
    On Error GoTo 0

    ' Si el pegado falla, intenta ppPasteEnhancedMetafile
    If pastedShape Is Nothing Then
        Set pastedShape = pptSlide.Shapes.PasteSpecial(DataType:=ppPasteEnhancedMetafile)
    End If
    Next shapeObj
    
    Application.CutCopyMode = False

'    ' function to copy charts
'    Call CopyCharts(pExcelSheet)
'
'
'    SetSlideTitle pptPresentation, i - 1, "B" & i
'
'    ' Paste charts
'    Set pptSlide = pptPresentation.Slides(i - 1)
'    pptSlide.Shapes.PasteSpecial DataType:=ppPasteEnhancedMetafile
'
   ' Application.CutCopyMode = False
    
    ' Buscar la línea roja y el texto en la hoja de Excel
   
  Next i
  
    
  Set pptSlide = Nothing
  Set pExcelSheet = Nothing


      
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

    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub

   
