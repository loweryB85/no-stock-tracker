Attribute VB_Name = "Rollers"

'This function is responsible for tracking the number of rollers on the active worksheet
'NOTE - rollers are of the format #X where # is the number of rolling picks in a particular
'order. Example: a route with 2 rolling picks would show 2X
Function FindRollers(rng As Range)

    Dim total As Integer
    Dim cell As Range
    
    'For each cell in the specified range
    For Each cell In rng
    
        'Check to see if the rightmost character in the cell is an X (lower-case or capital), thus indicating a roller
        If (Right(cell.Value, 1) = "X") Or (Right(cell.Value, 1) = "x") Then
        
            'We have found a roller. Get the number of rolling picks by retrieving the value to the left of the X
            total = total + Val(Left(cell.Value, Len(cell.Value) - 1))
            
        End If
        
    Next cell
    
    'Return the total number of rollers found in "rng"
    FindRollers = total

End Function

