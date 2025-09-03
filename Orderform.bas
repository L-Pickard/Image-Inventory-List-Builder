Option Explicit

Function Set_Worksheet(The_Code_Name As String, wbk As Workbook) _
        As Worksheet

    Dim Wks As Worksheet
    
    For Each Wks In wbk.Worksheets
        If Wks.CodeName = The_Code_Name Then
           Set Set_Worksheet = Wks
           Exit For
        End If
    Next Wks

End Function

Function GetUniqueValues(ws As Worksheet, startRow As Long, endRow As Long, columnNum As Long) As Variant
    Dim i As Long
    Dim cellValue As String
    Dim unique_array() As String
    Dim array_size As Long
    Dim exists As Boolean

    array_size = 0
    
    For i = startRow To endRow
        cellValue = ws.Cells(i, columnNum).Value2
        
        If array_size > 0 Then
            exists = False
            Dim j As Long
            For j = LBound(unique_array) To UBound(unique_array)
                If unique_array(j) = cellValue Then
                    exists = True
                    Exit For
                End If
            Next j
        Else
            exists = False
        End If
        
        If Not exists Then
            ReDim Preserve unique_array(array_size)
            unique_array(array_size) = cellValue
            array_size = array_size + 1
        End If
    Next i
    
    GetUniqueValues = unique_array
End Function

Function Insert_Logo(entity As String, _
                     targetSheet As Worksheet, _
                     imgLeft As Double, _
                     imgTop As Double, _
                     scaleH As Double) As Shape
            

    Dim logo As Shape
    Dim filePath As String
    
    If entity = "Shiner Ltd" Then
        
        filePath = "P:\Sales\Karl\Sales Builder Templates\Logos\SHINER_LOGO_BLK_LTD.png"
        
    ElseIf entity = "Shiner B.V" Then
        
        filePath = "P:\Sales\Karl\Sales Builder Templates\Logos\SHINER_LOGO_BLK_B.V.png"
        
    ElseIf entity = "Shiner LLC" Then
    
        filePath = "P:\Sales\Karl\Sales Builder Templates\Logos\SHINER_LOGO_BLK_LLC.png"
        
    Else
    
        filePath = "P:\Sales\Karl\Sales Builder Templates\Logos\SHINER_LOGO_BLK_GEN.png"
        
    End If
    

    Set logo = targetSheet.Shapes.AddPicture( _
        filename:=filePath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=imgLeft, _
        Top:=imgTop, _
        Width:=-1, _
        Height:=-1)

    logo.LockAspectRatio = msoTrue
    logo.ScaleHeight scaleH, msoTrue

    Set Insert_Logo = logo
    
    
End Function


Public Sub PickAFolder()
    Dim objFileDialog As FileDialog
    Dim objSelectedFolder As Variant
    Dim wbk As Workbook
    Dim ws_controls As Worksheet
    
    Set wbk = ThisWorkbook
    
    Set ws_controls = Set_Worksheet("Controls", wbk)
    
    Set objFileDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With objFileDialog
        .ButtonName = "Select"
        .Title = "Select a folder"
        .InitialView = msoFileDialogViewList
        .Show
        For Each objSelectedFolder In .SelectedItems
            ws_controls.Range("Save_Path").Value = objSelectedFolder
        Next
    End With
End Sub

Sub clear_filename()

    Dim wbk As Workbook
    Dim ws_controls As Worksheet
    
    Set wbk = ThisWorkbook
    
    Set ws_controls = Set_Worksheet("Controls", wbk)
    
    ws_controls.Range("File_Name").Value = ""

End Sub

Sub Filter_Data()

    Dim wbk As Workbook
    Dim ws_data As Worksheet
    Dim ws_piv As Worksheet
    Dim ws_controls As Worksheet
    Dim ws_test As Worksheet
    Dim answer As Integer

    Set wbk = ThisWorkbook
    
    Set ws_data = Set_Worksheet("Item_Data", wbk)
    Set ws_piv = Set_Worksheet("Pivot_Category", wbk)
    Set ws_controls = Set_Worksheet("Controls", wbk)
    
    answer = MsgBox("Do You want to create an xlsx sales list?" & vbNewLine & vbNewLine & "Click yes to run the vba code, no to exit." _
             , vbQuestion + vbYesNo + vbDefaultButton2, "Shiner Sales List Builder")

    If answer = vbNo Then

        Exit Sub

    End If
    
    If ws_controls.Range("Save_Path").Value = "" Then
    
        MsgBox "Please select output folder before creating xlsx.", vbExclamation, "Shiner Sales List Builder"
    
        Exit Sub
    
    End If
    
    If ws_controls.Range("Custom_FileName").Value = True And ws_controls.Range("File_Name").Value = "" Then
    
        MsgBox "You have selected to use a custom file name, you must provide one or unselect that option before proceeding.", vbExclamation, "Shiner Sales List Builder"

        Exit Sub

    End If
    
    Application.ScreenUpdating = False

    Dim piv_rows As Long
    Dim season_values As Variant
    Dim brand_values As Variant
    Dim category_values As Variant
    Dim group_values As Variant
    Dim price_values As Variant
    
    piv_rows = ws_piv.Range("A2").CurrentRegion.Rows.Count
    
    season_values = GetUniqueValues(ws_piv, 2, piv_rows, 2)
    brand_values = GetUniqueValues(ws_piv, 2, piv_rows, 3)
    category_values = GetUniqueValues(ws_piv, 2, piv_rows, 4)
    group_values = GetUniqueValues(ws_piv, 2, piv_rows, 5)
    
    ' ü = on sale
    ' û = Full Price
    
    If ws_controls.Range("Include_Sale") = True And ws_controls.Range("Include_Full_Price") = True Then
        price_values = Array("ü", "û")
        
    ElseIf ws_controls.Range("Include_Sale") = True And ws_controls.Range("Include_Full_Price") = False Then
        price_values = "ü"
        
    ElseIf ws_controls.Range("Include_Sale") = False And ws_controls.Range("Include_Full_Price") = True Then
        price_values = "û"
        
     ElseIf ws_controls.Range("Include_Sale") = False And ws_controls.Range("Include_Full_Price") = False Then
        price_values = ""
        
    End If
    
    Dim i As Long
    Dim row As Long

    Dim tbl As ListObject
    Set tbl = ws_data.ListObjects(1)

    With tbl
        If Not .AutoFilter Is Nothing Then .AutoFilter.ShowAllData
    End With

    tbl.Range.AutoFilter Field:=1, Criteria1:=ws_controls.Range("B7").Value, Operator:=xlFilterValues
    tbl.Range.AutoFilter Field:=17, Criteria1:=price_values, Operator:=xlFilterValues
    
    tbl.Range.AutoFilter Field:=15, Criteria1:=season_values, Operator:=xlFilterValues
    tbl.Range.AutoFilter Field:=5, Criteria1:=brand_values, Operator:=xlFilterValues
    tbl.Range.AutoFilter Field:=13, Criteria1:=category_values, Operator:=xlFilterValues
    tbl.Range.AutoFilter Field:=14, Criteria1:=group_values, Operator:=xlFilterValues
    
    Dim new_wb As Workbook
    Dim ws_new As Worksheet
    
    On Error Resume Next
    On Error GoTo error_handle
    
    If tbl.DataBodyRange.SpecialCells(xlCellTypeVisible).CountLarge > 1 Then
        Set new_wb = Workbooks.Add
        Set ws_new = new_wb.Sheets(1)
        
        tbl.Range.SpecialCells(xlCellTypeVisible).Copy
        ws_new.Range("B5").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        
    End If
    
    On Error GoTo 0
    
    If ws_controls.Range("Currency").Value = "GBP" Then
        
        ws_new.Columns("U:X").Delete
    
    ElseIf ws_controls.Range("Currency").Value = "EUR" Then
    
        ws_new.Columns("W:X").Delete
        ws_new.Columns("S:T").Delete
        
    ElseIf ws_controls.Range("Currency").Value = "USD" Then
    
        ws_new.Columns("S:V").Delete
        
    End If
    
    If ws_controls.Range("Entity").Value = "Shiner Ltd" Then
        
        ws_new.Columns("E").NumberFormat = "_-£* #,##0.00_-;-£* #,##0.00_-;_-£* ""-""??_-;_-@_-"
        
    ElseIf ws_controls.Range("Entity").Value = "Shiner B.V" Then
    
        ws_new.Columns("E").NumberFormat = "_-[$€-x-euro2] * #,##0.00_-;_-[$€-x-euro2] * -#,##0.00_-;_-[$€-x-euro2] * ""-""??_-;_-@_-"
        
    ElseIf ws_controls.Range("Entity").Value = "Shiner LLC" Then
    
        ws_new.Columns("E").NumberFormat = "_-[$$-en-US]* #,##0.00_-;_-[$$-en-US]* -#,##0.00_-;_-[$$-en-US]* ""-""??_-;_-@_-"
        
    End If
        
    
    If ws_controls.Range("Use_Size").Value = "Size 1" Then

        ws_new.Columns("J:K").Delete

    ElseIf ws_controls.Range("Use_Size").Value = "EU Size" Then

        ws_new.Columns("I").Delete
        ws_new.Columns("J").Delete

    ElseIf ws_controls.Range("Use_Size").Value = "US Size" Then

        ws_new.Columns("I:J").Delete

    End If
    
    Dim dt_val As Long
    Dim refresh_text As String
    
    dt_val = ws_new.Range("C6").Value
    
    refresh_text = "Last Refresh: " & Format(ws_new.Range("C6").Value, "dd/mm/yyyy hh:mm")
    
    ws_new.Columns("B:C").Delete

    Dim logo As Shape

    Set logo = Insert_Logo(ws_controls.Range("Entity").Value, ws_new, 30, 15, 0.25)


    ws_new.Columns("O").Copy
    
    ws_new.Columns("X").PasteSpecial Paste:=xlPasteFormats
    
    ws_new.Columns("Q").Copy
    
    ws_new.Columns("W").PasteSpecial Paste:=xlPasteFormats
    
    ws_new.Columns("B:U").AutoFit
    
    ws_new.Columns("A").ColumnWidth = 4
    
    ws_new.Range("E2").Value = refresh_text
    
    ws_new.Range("V5").Value = "Image"
    ws_new.Columns("V").ColumnWidth = 30
    ws_new.Range("W5").Value = "Quantity"
    ws_new.Range("X5").Value = "Line Total"
    
    Dim ws_new_rows As Long
    Dim cell As Range
    
    ws_new_rows = ws_new.Cells(ws_new.Rows.Count, 14).End(xlUp).row
    
    ws_new.Rows("6:" & ws_new_rows).RowHeight = 130
    
    With ws_new.Rows("5:" & ws_new_rows)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    For Each cell In ws_new.Range(ws_new.Cells(6, 14), ws_new.Cells(ws_new_rows, 14))

        cell.Font.Name = "Wingdings"
        
        If cell.Value = "ü" Then
        
            cell.Font.Color = vbGreen
            
        Else
            cell.Font.Color = vbRed
        
        End If
    
    Next cell
    
'below is the code to add the images
    
    Dim filename As String
    Dim im As Integer
    Dim insert_cell As Range
    
    On Error GoTo ErrorHandler
    
    Application.DisplayAlerts = False
    
    For im = 6 To ws_new_rows
    
        filename = Cells(im, 21).Value
        Set insert_cell = Cells(im, 22)
        
        insert_cell.Select
    
        If filename = "" Then
            insert_cell.Value = "No Image"
        Else
           
            If Dir(filename) <> "" Then
                insert_cell.InsertPictureInCell (filename)
            Else
                insert_cell.Value = "Invalid Path"
            End If
        End If
    
    Next im
    
    Application.DisplayAlerts = True
    
    On Error GoTo 0

    ' Below is the code after the image has been added
    
    ws_new.Columns("U").Delete
    
    ws_new.Range("V6:V" & ws_new_rows).Value = 0
    
    ws_new.Range("W6:W" & ws_new_rows).Formula = "=$V6 * $O6"
    
    ws_new.Cells(ws_new_rows + 1, 22).Formula = "=SUM($V$6:$V$" & ws_new_rows & ")"
    
    ws_new.Cells(ws_new_rows + 1, 23).Formula = "=SUM($W$6:$W$" & ws_new_rows & ")"
    
    ws_new.Range(Cells(ws_new_rows, 22), Cells(ws_new_rows, 23)).Copy
    ws_new.Range(Cells(ws_new_rows + 1, 22), Cells(ws_new_rows + 1, 23)).PasteSpecial Paste:=xlPasteFormats
    
    ws_new.Columns("V").ColumnWidth = 10
    ws_new.Columns("W").ColumnWidth = 11
    
    
    ' Below is the code for the formatting
    
    Dim main_colour As Long
    Dim alt_colour As Long
    Dim back_colour As Long
    
    If ws_controls.Range("Entity").Value = "Shiner Ltd" Then
    
        main_colour = RGB(35, 113, 182)
        alt_colour = RGB(0, 61, 171)
        back_colour = RGB(209, 239, 250)
        
    ElseIf ws_controls.Range("Entity").Value = "Shiner B.V" Then
        
        main_colour = RGB(252, 168, 0)
        alt_colour = RGB(249, 111, 0)
        back_colour = RGB(252, 237, 214)
        
    ElseIf ws_controls.Range("Entity").Value = "Shiner LLC" Then
        
        main_colour = RGB(45, 171, 102)
        alt_colour = RGB(8, 115, 41)
        back_colour = RGB(226, 239, 218)
        
    End If
  
  
  'border and gridline formats

    With ws_new.Range(Cells(6, 2), Cells(ws_new_rows, 23)).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = alt_colour
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With ws_new.Range(Cells(6, 2), Cells(ws_new_rows, 23)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = alt_colour
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With ws_new.Range(Cells(6, 2), Cells(ws_new_rows, 23)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = alt_colour
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With ws_new.Range(Cells(6, 2), Cells(ws_new_rows, 23)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = alt_colour
        .TintAndShade = 0
        .Weight = xlMedium
    End With

    
    With ws_new.Range(Cells(6, 2), Cells(ws_new_rows, 23)).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = alt_colour
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    ' header formats
    
    With ws_new.Range("B5:W5").Font
        .Name = "Aptos Narrow"
        .FontStyle = "Bold"
        .Size = 11
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    With ws_new.Range("B5:W5").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = alt_colour
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With ws_new.Range("B5:W5").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = alt_colour
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With ws_new.Range("B5:W5").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = alt_colour
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With ws_new.Range("B5:W5").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = alt_colour
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With ws_new.Range("B5:W5").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = main_colour
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    'Move prices to teh end doh!
    
    ws_new.Columns("O:P").Cut
    ws_new.Columns("U:V").insert Shift:=xlToRight
    
    'totals formats
    
    With ws_new.Range(Cells(6, 22), Cells(ws_new_rows + 1, 23)).Font
        .Name = "Aptos Narrow"
        .FontStyle = "Bold"
        .Size = 11
        .Color = alt_colour
    End With
    
    With ws_new.Range(Cells(6, 22), Cells(ws_new_rows, 23)).Borders(xlEdgeLeft)
        .LineStyle = xlDash
        .Color = alt_colour
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With ws_new.Range(Cells(6, 22), Cells(ws_new_rows, 23)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = alt_colour
        .TintAndShade = 0
        .Weight = xlMedium
    End With

    With ws_new.Range(Cells(6, 22), Cells(ws_new_rows, 23)).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = back_colour
    End With
    
' cost column

    With ws_new.Range(Cells(6, 3), Cells(ws_new_rows, 3)).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 65535
    End With
    
    With ws_new.Range(Cells(6, 3), Cells(ws_new_rows, 3)).Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    
    ws_new.Range(Cells(6, 3), Cells(ws_new_rows, 3)).Font.Bold = True
    
    With ws_new.Range("E2").Font
        .Name = "Aptos Narrow"
        .FontStyle = "Bold Italic"
        .Size = 11
        .Color = alt_colour
    End With
    
    ActiveWindow.DisplayGridlines = False
    
    ws_new.Name = Format(dt_val, "dd.mm.yyyy")
    
    ws_new.Tab.Color = main_colour
    
    Dim output_path As String
    
    If ws_controls.Range("Custom_FileName").Value = True Then
    
        output_path = ws_controls.Range("Save_Path").Value & "\" & ws_controls.Range("File_Name").Value & ".xlsx"
        
    Else
        output_path = ws_controls.Range("Save_Path").Value & "\" & ws_controls.Range("Entity").Value & " " & "Sales List " & Format(Now(), "dd.mm.yyyy hh.mm.ss") & ".xlsx"
        
    End If
    
    new_wb.SaveAs filename:=output_path, FileFormat:=xlOpenXMLWorkbook
    
    If ws_controls.Range("Close_On_Finish").Value = True Then
    
        new_wb.Close SaveChanges:=False
        
    End If
    
    Application.ScreenUpdating = True
    
    
    MsgBox "You have successfully created the below xlsx sales list:" & vbNewLine & vbNewLine & _
       output_path & vbNewLine & vbNewLine & "@ " & Format(Now(), "dd/mm/yyyy hh:mm:ss"), _
       vbOKOnly + vbInformation, "Shiner Sales List Builder"
    
    Exit Sub
    
error_handle:

    Application.ScreenUpdating = True

    MsgBox "No data matches the filer criteria, please select again.", vbExclamation, "Shiner Sales List Builder"
    
    Exit Sub
    
ErrorHandler:

    Resume Next
      
End Sub
