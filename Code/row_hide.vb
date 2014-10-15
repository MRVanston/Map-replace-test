Sub RangeHide()
    Set rColumnD = Range("D2", "D75")
    Set rColumnE = Range("E2", "E75")
    
    For Each Row In Worksheet
        If ((rColumnD.cell.Value = "{}") And (rColumnE.cell.Value = "{}")) Then
            Row.EntireRow.Hidden = True
        End If
    Next
    
    'For Each cell In rColumnD
        'If (cell.Value = "{}") Then
            'Rows(cell.Row).EntireRow.Hidden = True
        
        'Else
        
    'For Each cell In rColumnE
        'If cell.Value = "{}" Then
            'Rows(cell.Row).EntireRow.Hidden = True
        'End If
        
    'Next
 
End Sub