Attribute VB_Name = "ShiftSummary"
'calculateFlags
'This sub tracks the running total of flags for the current shift and
'calculates the number of flags that have been worked on the current shift.
Sub calculateFlags()

    Dim difference As Integer   'This will be used to do some math
    
    'get difference between previous total and current total from worksheet
    difference = Range("A60").Value - Range("W26").Value
    
    'The sign (+ / -) of the difference reveals whether we worked or gained flags
    If difference > 0 Then
        'difference is positive - we worked flags
        Range("G35").Value = (Range("G35").Value + difference)
    Else
        'difference is negative - we gained flags
        Range("G34").Value = (Range("G34").Value - difference)
    End If
    
    'reset previous total = current total
    Range("A60").Value = Range("W26").Value
    
End Sub

'clearShift
'This sub resets the shift summary section of the active worksheet
Sub clearShift()

    Range("A60").Value = 0  'clears the hidden cell containing running total
    Range("G34").Value = 0  'clears "Total Flags This Shift"
    Range("G35").Value = 0  'clears "Flags Worked This Shift"=
    
End Sub

'*************************
'Set Previous Shift to 1st
'*************************
Sub setPrevFirst()

ActiveSheet.Range("G33") = "='1st Shift'!W26"

End Sub

'*************************
'Set Previous Shift to 2nd
'*************************
Sub setPrevSecond()

ActiveSheet.Range("G33") = "='2nd Shift'!W26"

End Sub
'*************************
'Set Previous Shift to 3rd
'*************************
Sub setPrevThird()

ActiveSheet.Range("G33") = "='3rd Shift'!W26"

End Sub
'*************************
'Set Previous Shift to Last Day
'*************************
Sub setPrevLast()

ActiveSheet.Range("G33") = "='LAST DAY'!W26"

End Sub


