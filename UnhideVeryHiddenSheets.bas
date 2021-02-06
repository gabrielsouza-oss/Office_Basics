''Unhide VeryHiddenSheets
Sub UnhideAllSheets()
    Dim wks As Worksheet
 
    For Each wks In ActiveWorkbook.Worksheets
        wks.Visible = xlSheetVisible
    Next wks
End Sub
'https://www.ablebits.com/office-addins-blog/2017/12/20/very-hidden-sheets-excel/