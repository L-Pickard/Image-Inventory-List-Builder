Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

    Dim wbk As Workbook
    Dim ws_piv As Worksheet
    Dim ws_controls As Worksheet
    Dim pt As PivotTable
    Dim filter_Value As String
    Dim entity As Range
    
    Set wbk = ThisWorkbook
    Set ws_controls = Set_Worksheet("Controls", wbk)
    
    Set entity = ws_controls.Range("Entity")
    
    If Not Intersect(Target, entity) Is Nothing Then
        
        Set ws_piv = Set_Worksheet("Pivot_Category", wbk)
        
        filter_Value = entity.Value
        
        Set pt = ws_piv.PivotTables("Pivot_Category")
        
    With pt.PivotFields("Entity")
        .ClearAllFilters
        
        Dim pItem As PivotItem
        For Each pItem In .PivotItems
            If pItem.Name = filter_Value Then
                pItem.Visible = True
            Else
                pItem.Visible = False
            End If
        Next pItem
    End With
        
    End If
    
End Sub
