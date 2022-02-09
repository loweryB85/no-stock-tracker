VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} optionsForm 
   Caption         =   "Options"
   ClientHeight    =   3330
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9435.001
   OleObjectBlob   =   "optionsForm.frx":0000
End
Attribute VB_Name = "optionsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private shift As String

Private Sub UserForm_Initialize()
'PURPOSE: Position userform to center of Excel Window (important for dual monitor compatibility)
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

'Start Userform Centered inside Excel Screen (for dual monitors)
  Me.StartUpPosition = 0
  Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
  Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub

'disable x button to close
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    'CloseMode other than 1 indicates that unload was not initiated via code
    If CloseMode <> 1 Then
        
        Cancel = 1
        MsgBox "Please make a selection."
        
    End If
    
    
End Sub

'action for begin shift button
Private Sub beginShift_Click()
    
    'get radio button input
    If Me.option1st Then
        shift = "1st Shift"
    ElseIf Me.option2nd Then
        shift = "2nd Shift"
    ElseIf Me.option3rd Then
        shift = "3rd Shift"
    ElseIf Me.optionLast Then
        shift = "Last Day"
    Else
        MsgBox "No selection made."
        Exit Sub
    End If
    
    'present user with a chance to cancel
    If messageBox Then
        
        Sheets(shift).Activate
        
        Call ShiftSummary.clearShift
        Call Sheets(shift).ClearFlags
        
        Unload Me
        
        'MsgBox "True"
            
    End If
    
End Sub

'present user with option to cancel
Function messageBox()

    'check for cancel
    result = MsgBox("This will clear current flags from " & shift & "." & vbNewLine & vbNewLine & "Press Cancel to make a different selection.", vbOKCancel + vbExclamation, "WARNING")
    
        If result = vbOK Then
        
            result = True
            
        ElseIf result = vbCancel Then
        
            result = False
            
        End If
    
    messageBox = result

End Function

'action for continue shift button
Private Sub continueShift_Click()

    'get radio button input
    If Me.option1st Then
        shift = "1st Shift"
    ElseIf Me.option2nd Then
        shift = "2nd Shift"
    ElseIf Me.option3rd Then
        shift = "3rd Shift"
    ElseIf Me.optionLast Then
        shift = "Last Day"
    Else
        MsgBox "No selection made."
        Exit Sub
    End If
        
    'activate chosen sheet
    Sheets(shift).Activate
    
    Unload Me

End Sub
