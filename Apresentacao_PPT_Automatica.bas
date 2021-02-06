Sub pastePPT(control As IRibbonControl)
    
    If Not sentFile Then
    
        Application.ScreenUpdating = False
        
        Dim rng As Range
        Dim newPres As Boolean
        
        Dim ppApp As PowerPoint.Application
        Dim ppPres As PowerPoint.Presentation
        Dim ppSld As PowerPoint.Slide
        Dim ppShp As PowerPoint.Shape
        
        Set ppApp = getPowerPointInstance(newPres)
        
        If newPres Then
            On Error Resume Next
            ppApp.presentations.Open Filename:=Application.ThisWorkbook.Path & "\templates\GMtemplate.pptx"
            
            If Err.Number <> 0 Then
                MsgBox "Template não encontrado!" & vbCrLf & "Operação cancelada!"
                ppApp.Quit
                Exit Sub
            End If
            Err.Clear
            On Error GoTo 0
        End If
        
        Set ppPres = ppApp.ActivePresentation
        
        mes = Month(Application.ThisWorkbook.Sheets("MENU").Range("D3").value)
        adjustTable mes
        Set rng = Application.ThisWorkbook.Sheets("MENU").Range("D14:S81")
        
        Application.Wait (Now + TimeValue("00:00:01"))
        Set ppSld = ppPres.Slides.Add(ppPres.Slides.Count + 1, ppLayoutBlank)
        ppApp.ActiveWindow.View.GotoSlide ppSld.SlideIndex
        
        rng.Copy
        Application.Wait (Now + TimeValue("00:00:01"))
        ppSld.Shapes.PasteSpecial DataType:=2
        
        ppSld.Shapes(ppSld.Shapes.Count).Height = 380
        ppSld.Shapes(ppSld.Shapes.Count).Left = (ppPres.PageSetup.SlideWidth - ppSld.Shapes(ppSld.Shapes.Count).Width) / 2
        ppSld.Shapes(ppSld.Shapes.Count).Top = 70
        
        Application.Wait (Now + TimeValue("00:00:01"))
        Set ppShp = ppSld.Shapes.AddTextbox(1, 0, 0, ppPres.PageSetup.SlideWidth, 50)
        With ppShp.TextFrame.TextRange
            .Text = Application.ThisWorkbook.Sheets("MENU").Range("F84").value
            .Font.Name = "GM Global Sans Bold"
            .Font.Size = 28
            .ParagraphFormat.Alignment = ppAlignCenter
            .Font.Color = RGB(0, 0, 0)
        End With
        
        readjustTable
        Application.CutCopyMode = False
        
        Application.ThisWorkbook.Sheets("GRÁFICO").ChartArea.Copy
        
        Application.Wait (Now + TimeValue("00:00:01"))
        Set ppSld = ppPres.Slides.Add(ppPres.Slides.Count + 1, ppLayoutBlank)
        ppApp.ActiveWindow.View.GotoSlide ppSld.SlideIndex
        
        ppApp.CommandBars.ExecuteMso "PasteExcelChartSourceFormatting"
        
        Application.Wait (Now + TimeValue("00:00:02"))
        On Error Resume Next
        Do
            Err.Clear
            Set ppShp = ppSld.Shapes(ppSld.Shapes.Count)
        Loop While Err.Number <> 0
        On Error GoTo 0
            
        LockAspectRatio = msoFalse
        Application.Wait (Now + TimeValue("00:00:01"))
        ppShp.Height = ppPres.PageSetup.SlideHeight
        ppShp.Width = ppPres.PageSetup.SlideWidth
        ppShp.Left = 0
        ppShp.Top = 0
        ppSld.DisplayMasterShapes = msoFalse
        
        If newPres Then
            With Application.ThisWorkbook.Sheets("MENU")
                ppPres.Slides(1).Delete
                ppPres.SaveAs Application.ThisWorkbook.Path & "\Market Share relativo_" & Format(.Range("D3").value, "dd-mm-yyyy") _
                & " - " & .Range("M5").value & ".pptx"
            End With
        Else
            Application.DisplayAlerts = False
            ppPres.Save
        End If
        
        Application.ThisWorkbook.Sheets("MENU").Range("A1").Select
        Application.ScreenUpdating = True
        
    Else
        MsgBox "Ação não permitida!"
    End If
    
End Sub
Private Function getPowerPointInstance(ByRef newPres As Boolean) As PowerPoint.Application
    On Error Resume Next
    Set getPowerPointInstance = GetObject(Class:="PowerPoint.Application")
    On Error GoTo 0
    newPres = False
    Err.Clear
    If getPowerPointInstance Is Nothing Then
        Application.Wait (Now + TimeValue("00:00:01"))
        Set getPowerPointInstance = CreateObject("PowerPoint.Application")
        'getPowerPointInstance.Visible = msoTrue
        newPres = True
    End If
    If Err.Number = 429 Then
        MsgBox "PowerPoint could not be found, aborting."
        Exit Function
    End If
End Function
Private Function adjustTable(ByVal mes As Integer)
    Application.ThisWorkbook.Sheets("MENU").Rows("37:55").EntireRow.Hidden = True
    Application.ThisWorkbook.Sheets("MENU").Rows("58:76").EntireRow.Hidden = True
    Select Case mes
        Case 1
            Application.ThisWorkbook.Sheets("MENU").Columns("H:R").EntireColumn.Hidden = True
        Case 2
            Application.ThisWorkbook.Sheets("MENU").Columns("I:R").EntireColumn.Hidden = True
        Case 3
            Application.ThisWorkbook.Sheets("MENU").Columns("J:R").EntireColumn.Hidden = True
        Case 4
            Application.ThisWorkbook.Sheets("MENU").Columns("K:R").EntireColumn.Hidden = True
        Case 5
            Application.ThisWorkbook.Sheets("MENU").Columns("L:R").EntireColumn.Hidden = True
        Case 6
            Application.ThisWorkbook.Sheets("MENU").Columns("M:R").EntireColumn.Hidden = True
        Case 7
            Application.ThisWorkbook.Sheets("MENU").Columns("N:R").EntireColumn.Hidden = True
        Case 8
            Application.ThisWorkbook.Sheets("MENU").Columns("O:R").EntireColumn.Hidden = True
        Case 9
            Application.ThisWorkbook.Sheets("MENU").Columns("P:R").EntireColumn.Hidden = True
        Case 10
            Application.ThisWorkbook.Sheets("MENU").Columns("Q:R").EntireColumn.Hidden = True
        Case 11
            Application.ThisWorkbook.Sheets("MENU").Columns("R:R").EntireColumn.Hidden = True
    End Select
End Function
Private Function readjustTable()
    Application.ThisWorkbook.Sheets("MENU").Rows("37:55").EntireRow.Hidden = False
    Application.ThisWorkbook.Sheets("MENU").Rows("58:76").EntireRow.Hidden = False
    Application.ThisWorkbook.Sheets("MENU").Columns("H:R").EntireColumn.Hidden = False
End Function
