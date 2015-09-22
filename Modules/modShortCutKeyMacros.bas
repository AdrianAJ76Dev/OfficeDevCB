Attribute VB_Name = "modShortCutKeyMacros"
Option Explicit

'Common Macros

'**********************************************************************************************
'Shortcut Key: Alt-S,N ---- S for Spell, N for Number
'**********************************************************************************************
Public Sub InterfaceForSpellNumber()
    On Error GoTo ErrorHandler

    Selection.Text = SpellNumber(Selection.Text)

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "InterfaceForSpellNumber"
    End With
    Resume Exit_Here
End Sub

'**********************************************************************************************
'Shortcut Key: Alt-F,P ---- F for Format, P for Phone
'**********************************************************************************************
Public Sub FormatPhoneNumber() 'Shortcut Key: Alt-F, P "P" for Phone
    Dim strPhoneNumberDigits As String
    On Error GoTo ErrorHandler
    
    'Check selection to see if 10 digits are selected
    strPhoneNumberDigits = Trim(Selection.Words(1).Text)
    If IsNumeric(strPhoneNumberDigits) And Len(strPhoneNumberDigits) = 10 Then
        strPhoneNumberDigits = Format(strPhoneNumberDigits, "(###)-###-####")
        Selection.Words(1).Text = strPhoneNumberDigits
    Else
        MsgBox "Your selection does not solely consist of numbers " & vbCr & _
        "or consists of more than or less than 10 digits " & "Number Count: " & Len(strPhoneNumberDigits), _
        vbExclamation, "Convert Number to Phone Format"
    End If

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "FormatPhoneNumber"
    End With
    Resume Exit_Here
End Sub

'**********************************************************************************************
'Shortcut Key: Alt-F, D ---- F for Format, D for Date
'**********************************************************************************************
Public Sub FormatDateSpellOutMonth()
    On Error GoTo ErrorHandler
    
    'Use a regular expression here
    Selection.Text = Format(Selection, "mmmm dd, yyyy")

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "FormatDateSpellOutMonth"
    End With
    Resume Exit_Here
End Sub

'Rider Macros

'**********************************************************************************************
'Shortcut Key: Alt-R, A ---- R for Rider, A for ACTIVATE
'**********************************************************************************************
Public Sub ActivateRider()
    Dim objCBContract As clsCBContract
    On Error GoTo ErrorHandler
    
    Set objCBContract = New clsCBContract
    objCBContract.ActivateRider

Exit_Here:
    Set objCBContract = Nothing
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "ActivateRider"
    End With
    Resume Exit_Here
End Sub


'**********************************************************************************************
'Shortcut Key: Alt-R, D ---- R for Rider, D for DELETE
'**********************************************************************************************
Public Sub DeleteUnncessaryRiders()
    Dim objCBContract As clsCBContract
    On Error GoTo ErrorHandler
    
    Set objCBContract = New clsCBContract
    objCBContract.DeleteUnncessaryRiders

Exit_Here:
    Set objCBContract = Nothing
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "DeleteUnncessaryRiders"
    End With
    Resume Exit_Here
End Sub

