# auto_gen_salary_slip
auto generate salary slip in excel
Sub auto_generate_salary_slip()

    Dim header_range As Range
    Dim name_range As Range
    Dim cell As Range
    Dim wb As Workbook
    Dim salary_slip_folder As Object
    Dim path As String
    Dim filename As String
    Dim cell_row As Long
    Dim code As Long
    
    

On Error GoTo errorhandle

    MkDir "G:\公司薪資\薪資\109年薪資\109年薪資通知單\" & "109 " & [E4] & " 薪資通知單\"
    path = "G:\公司薪資\薪資\109年薪資\109年薪資通知單\" & "109 " & [E4] & " 薪資通知單\"
    

'asking the user for the name they wish to create a salary slip
    Set header_range = Application.InputBox("請選擇標題範圍", "Range Selection", , , , , , 8)
    Set name_range = Application.InputBox("請選擇姓名範圍", "Range Selection", , , , , , 8)

    
    
    

'loop though the range to fix each cell

    For Each cell In name_range
        
        cell_row = cell.Row
    
        
    
        filename = cell.Value & "109 " & [E4] & " 薪資通知單"
        
        code = Range("C" & cell_row).Value
    
        cell.EntireRow.Select
        
        Selection.Copy
        
        Workbooks.Add
        
        Range("A4").PasteSpecial Paste:=xlPasteValues
        
        
        
        Columns("A:C").Delete
        
        header_range.Copy
        
        Range("A1").Select
        
        ActiveSheet.Paste
        
        Range("A4").NumberFormatLocal = "yyyy/m/d"
        
        Cells.EntireColumn.AutoFit
        
        Call add_border
        
        Range("A1").Select
        
        
        ActiveWorkbook.SaveAs filename:=path & filename, Password:=code
        
        
        
        ActiveWorkbook.Close
        
        ThisWorkbook.Activate
        
        
            
     
        
        
    Next cell
    
    Exit Sub
    
errorhandle:

        Select Case Err.Number
        
            Case 424
                Exit Sub
            Case Else
                MsgBox "發生錯誤"
                
        End Select
        
    
End Sub

Sub add_border()
'
'
'

'
    Range("A4:R4").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
End Sub
