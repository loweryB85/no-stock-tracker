VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' This code is executed upon the opening of the file
Private Sub Workbook_Open()

    'protect the sheets in case they were not already protected
    protectSheets
    
    'initialize flag cells to watch for change
    Sheet1.initUnion
    Sheet2.initUnion
    Sheet3.initUnion
    Sheet4.initUnion
    
    'load userform
    Load optionsForm
    optionsForm.Show


End Sub

' This code is executed upon closing the file
Private Sub Workbook_Close()
    
    'protect the sheets upon closing the file as a precaution
    protectSheets

End Sub

'Password protect all sheets and workbook itself
'We will make use of the UserInterfaceOnly flag so as not to interfere with macros
Private Sub protectSheets()
    
    Dim pass As String
    pass = "somePassword"
    
    ' check shared and stop sharing
    With ActiveWorkbook
    If .MultiUserEditing Then
    Application.DisplayAlerts = False
    .ExclusiveAccess
    Application.DisplayAlerts = True
    End If
    End With
    
    'protection
    ThisWorkbook.Protect Password:=pass, Structure:=True
    Sheets("1st Shift").Protect Password:=pass, UserInterfaceOnly:=True
    Sheets("2nd Shift").Protect Password:=pass, UserInterfaceOnly:=True
    Sheets("3rd Shift").Protect Password:=pass, UserInterfaceOnly:=True
    Sheets("LAST DAY").Protect Password:=pass, UserInterfaceOnly:=True
    
    'resave with sharing switched on
    With ActiveWorkbook
    Application.DisplayAlerts = False
    .SaveAs Filename:=.FullName, AccessMode:=xlShared
    Application.DisplayAlerts = True
    End With


End Sub

