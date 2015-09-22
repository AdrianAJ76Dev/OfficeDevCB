Attribute VB_Name = "modCommonTools"
Option Explicit

Public Sub ReverseName()
    On Error GoTo ErrorHandler
    
    Dim varManagerName As Variant
    varManagerName = Split(Selection, " ")
    Selection = varManagerName(1) + Space(1) + varManagerName(0)
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "ReverseName"
    End With
    Resume Exit_Here
End Sub

Public Sub TestFileSaveAsDialog()
    Dim strPath As String
    On Error GoTo ErrorHandler

    strPath = "Q:\RAS Contracts Management\Pivotal\K12 Contracts - Pivotal\_PSAT Agreements\OH\Cincinnati\Agreements\2013-2014\Draft"
    With Application.FileDialog(msoFileDialogSaveAs)
        .InitialFileName = strPath
        .Show
    End With

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "TestFileSaveAsDialog"
    End With
    Resume Exit_Here
End Sub


Public Function SpellNumber(ByVal MyNumber) As String
    Dim Temp As String
    Dim DecimalDigits As String
    Dim NumberAsString As String
    Dim DecimalPlace, Count
    
    On Error GoTo ErrorHandler
    
    ReDim Place(9) As String
    Place(2) = " Thousand "
    Place(3) = " Million "
    Place(4) = " Billion "
    Place(5) = " Trillion "

    'String representation of amount.
    If IsNumeric(MyNumber) Then
        MyNumber = Trim(Str(MyNumber))

        'Position of decimal place 0 if none.
        DecimalPlace = InStr(MyNumber, ".")
    
        'Convert cents and set MyNumber to dollar amount.
        If DecimalPlace > 0 Then
            DecimalDigits = GetTens(Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2))
            MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
        End If
        
        Count = 1
        Do While MyNumber <> ""
            Temp = GetHundreds(Right(MyNumber, 3))
            If Temp <> "" Then NumberAsString = Temp & Place(Count) & NumberAsString
            If Len(MyNumber) > 3 Then
                MyNumber = Left(MyNumber, Len(MyNumber) - 3)
            Else
                MyNumber = ""
            End If
            Count = Count + 1
        Loop
    Else
        MsgBox "The current selection is not a number or part of it contains letters." & vbCr _
        & "Please select only numbers to be converted.", vbExclamation, "Spell Number"
        NumberAsString = Trim(MyNumber)
    End If
        
    SpellNumber = NumberAsString

Exit_Here:
    Exit Function
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "SpellNumber"
    End With
    Resume Exit_Here
End Function

'*******************************************

' Converts a number from 100-999 into text *

'*******************************************

Function GetHundreds(ByVal MyNumber)
    Dim Result As String

    On Error GoTo ErrorHandler

    If Val(MyNumber) = 0 Then Exit Function

    MyNumber = Right("000" & MyNumber, 3)



    ' Convert the hundreds place.

    If Mid(MyNumber, 1, 1) <> "0" Then

        Result = GetDigit(Mid(MyNumber, 1, 1)) & " Hundred "

    End If



    ' Convert the tens and ones place.

    If Mid(MyNumber, 2, 1) <> "0" Then

        Result = Result & GetTens(Mid(MyNumber, 2))

    Else

        Result = Result & GetDigit(Mid(MyNumber, 3))

    End If



    GetHundreds = Result

Exit_Here:
    Exit Function
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "GetHundreds"
    End With
    Resume Exit_Here
End Function

'*********************************************

' Converts a number from 10 to 99 into text. *

'*********************************************

Function GetTens(TensText)

    Dim Result As String

    On Error GoTo ErrorHandler

    Result = ""           ' Null out the temporary function value.

    If Val(Left(TensText, 1)) = 1 Then   ' If value between 10-19...

        Select Case Val(TensText)

            Case 10: Result = "Ten"

            Case 11: Result = "Eleven"

            Case 12: Result = "Twelve"

            Case 13: Result = "Thirteen"

            Case 14: Result = "Fourteen"

            Case 15: Result = "Fifteen"

            Case 16: Result = "Sixteen"

            Case 17: Result = "Seventeen"

            Case 18: Result = "Eighteen"

            Case 19: Result = "Nineteen"

            Case Else

        End Select

    Else                                 ' If value between 20-99...

        Select Case Val(Left(TensText, 1))

            Case 2: Result = "Twenty "

            Case 3: Result = "Thirty "

            Case 4: Result = "Forty "

            Case 5: Result = "Fifty "

            Case 6: Result = "Sixty "

            Case 7: Result = "Seventy "

            Case 8: Result = "Eighty "

            Case 9: Result = "Ninety "

            Case Else

        End Select

        Result = Result & GetDigit(Right(TensText, 1))  ' Retrieve ones place.

    End If

    GetTens = Result

Exit_Here:
    Exit Function
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "GetTens"
    End With
    Resume Exit_Here
End Function

'*******************************************

' Converts a number from 1 to 9 into text. *

'*******************************************

Function GetDigit(Digit)
    
    On Error GoTo ErrorHandler
    Select Case Val(Digit)

        Case 1: GetDigit = "One"

        Case 2: GetDigit = "Two"

        Case 3: GetDigit = "Three"

        Case 4: GetDigit = "Four"

        Case 5: GetDigit = "Five"

        Case 6: GetDigit = "Six"

        Case 7: GetDigit = "Seven"

        Case 8: GetDigit = "Eight"

        Case 9: GetDigit = "Nine"

        Case Else: GetDigit = ""

    End Select

Exit_Here:
    Exit Function
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "GetDigit"
    End With
    Resume Exit_Here
End Function

Function SpellNumberOutToCurrency(ByVal MyNumber)
    Dim Dollars, Cents, Temp
    Dim DecimalPlace, Count

    On Error GoTo ErrorHandler

    ReDim Place(9) As String

    Place(2) = " Thousand "
    Place(3) = " Million "
    Place(4) = " Billion "
    Place(5) = " Trillion "



    ' String representation of amount.

    MyNumber = Trim(Str(MyNumber))



    ' Position of decimal place 0 if none.

    DecimalPlace = InStr(MyNumber, ".")

    ' Convert cents and set MyNumber to dollar amount.

    If DecimalPlace > 0 Then

        Cents = GetTens(Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2))

        MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))

    End If



    Count = 1

    Do While MyNumber <> ""

        Temp = GetHundreds(Right(MyNumber, 3))

        If Temp <> "" Then Dollars = Temp & Place(Count) & Dollars

        If Len(MyNumber) > 3 Then

            MyNumber = Left(MyNumber, Len(MyNumber) - 3)

        Else

            MyNumber = ""

        End If

        Count = Count + 1

    Loop



    Select Case Dollars

        Case ""

            Dollars = "No Dollars"

        Case "One"

            Dollars = "One Dollar"

        Case Else

            Dollars = Dollars & " Dollars"

    End Select



    Select Case Cents

        Case ""

            Cents = " and No Cents"

        Case "One"

            Cents = " and One Cent"

        Case Else

            Cents = " and " & Cents & " Cents"

    End Select



    SpellNumberOutToCurrency = Dollars & Cents

Exit_Here:
    Exit Function
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "SpellNumberOutToCurrency"
    End With
    Resume Exit_Here
End Function

Public Sub CopyDocumentProperties()
    Dim tmplCopyFrom As Template
    Dim docCopyFrom As Document
    Dim tmplCopyTo As Template
    Dim strTemplatePath As String
    Dim strTemplateCopyFrom As String
    Dim objCustDocProp As DocumentProperty
    
    Const strTEMPLATE_NAME_HED_RP_TERMS As String = "HED - RP Terminations.dotx"
    Const strPATH_TEMPLATE As String = "C:\Documents and Settings\ajones\Application Data\Microsoft\Templates\"
    
    On Error GoTo ErrorHandler
    
    'May need to throw this into the global template
    Set tmplCopyFrom = ThisDocument.AttachedTemplate
    
    If ActiveDocument.Name <> ThisDocument.Name Then
        For Each objCustDocProp In ThisDocument.CustomDocumentProperties
            If objCustDocProp.Name <> "ExecutionDate" Then
                ActiveDocument.CustomDocumentProperties.Add _
                    Name:=objCustDocProp.Name, LinkToContent:=False, Value:=objCustDocProp.Value, _
                    Type:=msoPropertyTypeString
            End If
        Next objCustDocProp
    End If
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "CopyDocumentProperties"
    End With
    Resume Exit_Here
End Sub

Public Function FirstSecondThirdForth(ByRef strDigit As String) As String
    Dim strLastDigit As String
    On Error GoTo ErrorHandler
    
    strLastDigit = Right(Trim(strDigit), 1)
    Select Case strLastDigit
        Case 1
            strDigit = strDigit & "st"
            
        Case 2
            strDigit = strDigit & "nd"
            
        Case 3
            strDigit = strDigit & "rd"
            
        Case Else
            strDigit = strDigit & "th"
        
    End Select
    
    FirstSecondThirdForth = strDigit
    
Exit_Here:
    Exit Function
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "FirstSecondThirdForth"
    End With
    Resume Exit_Here
End Function

Public Function FormatDateForAmendment(ByRef strAmendmentDate As String) As String
    Dim datAmendmentDate As Date
    On Error GoTo ErrorHandler
    
    'Grab the day for 1st, 2nd, 3rd, 4th formatting
    datAmendmentDate = strAmendmentDate
    FormatDateForAmendment = FirstSecondThirdForth(Day(strAmendmentDate)) & " day of " & Format(datAmendmentDate, "mmmm") & ", " & Format(datAmendmentDate, "yyyy")
    
Exit_Here:
    Exit Function
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "FormatDateForAmendment"
    End With
    Resume Exit_Here
End Function

Public Function GetAmendmentDate() As String
    Dim strEnteredDate As String
    On Error GoTo ErrorHandler
    
    strEnteredDate = InputBox("Enter Date of the Main Agreement in m/dd/yy", "Main Agreement Date")
    GetAmendmentDate = FormatDateForAmendment(strEnteredDate)
    
Exit_Here:
    Exit Function
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "GetAmendmentDate"
    End With
    Resume Exit_Here
End Function

Public Sub RunInputAmendmentDate()
    On Error GoTo ErrorHandler
    
    Selection.Text = GetAmendmentDate

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "RunInputAmendmentDate"
    End With
    Resume Exit_Here
End Sub
