Sub RangeHide()
    Set rColumnD = Range("D2", "D75")
    Set rColumnE = Range("E2", "E75")
    
    For Each cell In rColumnD
        If (cell.Value = "{}") Then
            Rows(cell.Row).EntireRow.Hidden = True
        End If
                
    Next
 
End Sub