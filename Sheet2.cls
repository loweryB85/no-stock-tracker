VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private dk41Range, dk24Range, dk25Range, dk1wRange, dk81Range, dk82Range, dk23Range, dk55Range, dk71Range, dk16Range, flagCells As Range

'initialize the union that will assist in catching worksheet changes
Sub initUnion()

    'set module level ranges
    Set dk41Range = Range("E3:J3,E5:J5,E7:J7,E9:J9,E11:J11,E13:J13,E15:J15,E17:J17,E19:J19,E21:J21,E23:J23,E25:J25")
    Set dk24Range = Range("O3:V3,O5:V5,O7:V7,O9:V9,O11:V11,O13:V13,O15:V15,O17:V17,O19:V19,O21:V21,O23:V23,O25:V25")
    Set dk25Range = Range("AC3:AI3,AC5:AI5,AC7:AI7,AC9:AI9,AC11:AI11,AC13:AI13,AC15:AI15,AC17:AI17,AC19:AI19,AC21:AI21,AC23:AI23,AC25:AI25")
    Set dk1wRange = Range("AN3:AO3,AN3:AO14")
    Set dk81Range = Range("AT3:AY14")
    Set dk82Range = Range("BD3:BE14")
    Set dk23Range = Range("BJ3:BK14")
    Set dk55Range = Range("AN17:AO25")
    Set dk71Range = Range("AT17:AW25")
    Set dk16Range = Range("BD17:BE25")

    'set union
    Set flagCells = Union(dk41Range, dk24Range, dk25Range, dk1wRange, dk81Range, dk82Range, dk23Range, dk55Range, dk71Range, dk16Range)

End Sub


' Worksheet_Change
' Note: Target(Range) - the range to watch for changes in
Private Sub Worksheet_Change(ByVal Target As Range)

    'If the worksheet changes in the ranges that can contain flags
    If Intersect(Target, flagCells) Is Nothing Then
        Exit Sub
        
    Else
        
        'Check to see if the change being made is a deletion. In order to prevent flags from being counted more than once in the
        'event that a cell's value is reduced by clearing the cell prior to changing the value, we must ignore calculations after
        'the change event thrown by the delete key.
         If (Target.Cells(1).Value <> "") Then
        
            Call ShiftSummary.calculateFlags
            
        Else
            
            'value was removed with delete key - to prevent user error from impacting shift summary, we must compensate
            Dim diff As Integer
            
            'determine deleted qty
            diff = Range("A60").Value - Range("W26").Value

            'compensate holding cell
            Range("A60") = Range("A60").Value - diff
            
            'compensate total flags
            Range("G34") = Range("G34").Value - diff
        
        End If
    End If
    

End Sub

'ClearFlags
'This sub performs ClearContents on each cell of the active worksheet that could potentially contain flags (DOES NOT include times or Kanban cycles)
Sub ClearFlags()

    'disable application events to prevent the Worksheet_Change code from executing (would result in a run-time error)
    Application.EnableEvents = False

    'dk41Range.ClearContents
    flagCells.ClearContents

    're-enable application events as the Worksheet_Change code needs to be able to run for shift summary to load properly
    Application.EnableEvents = True

End Sub










