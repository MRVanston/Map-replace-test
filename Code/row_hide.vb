Sub RangeHide()

Dim rCell As Range
    
    For Each rCell In Range("D2:D75")
        If rCell = "{}" And rCell.Offset(0, 1) = "{}" Then
            rCell.EntireRow.Hidden = True
        Else
            rCell.EntireRow.Hidden = False
            
        End If
    
    Next
 
End Sub