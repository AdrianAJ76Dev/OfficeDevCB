Attribute VB_Name = "modCBContractTools_v2"
Option Explicit
Dim mobjContract As clsContract

Public Sub Clean_Up_Contract()
    
    Const strREGEX_DATE_FORMAT_EBS As String = "" 'EBS = Enrollment Budget Schedule
    
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = wdAlertsNone
    Application.ScreenUpdating = False 'I may have to comment this out because I'm using the Selection Object and I think that NEEDS for Screenupdating to be TRUE
    Set mobjContract = New clsContract
    With ActiveDocument
        .TrackRevisions = True
        'Turn off seeing revisions or "Final" so the Find in the Clean_Up code works correctly
        .ShowRevisions = False
    End With
    
    With mobjContract
        .Clean_Up_Riders
        .FormatDates
        .Clean_Up_EnrollmentBudgetTable
        .CleanUp_General
    End With
    
    'Turn on seeing revisions or "Final: Show Markup"
    ActiveDocument.ShowRevisions = True
    
ExitHere:
    Application.DisplayAlerts = wdAlertsAll
    Application.ScreenUpdating = True
    Set mobjContract = Nothing
    Exit Sub
    
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCrLf & .Source & vbCrLf & .Description, vbCritical + vbOKOnly, "Clean_Up_Contract"
    End With
    Resume ExitHere
End Sub

Public Sub Clean_Up_Riders()
    '************************************************************************************************
    'Author: Adrian Jones
    'Date Created: 08/24/2015
    'This routine exists in the class as well as here. The purpose for it HERE is to tie a keystroke
    'to this routine.
    'Alt C, R for C = Clean Up, R = Rider
    '************************************************************************************************
    
    Dim varRiders As Variant
    Dim fldRider As Field
    Dim idx As Integer
    
    Set mobjContract = New clsContract
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    varRiders = mobjContract.FindRidersUsingFieldObject()
    For idx = 0 To UBound(varRiders)
        Set fldRider = varRiders(idx)
        fldRider.Unlink
    Next idx
    
    For Each fldRider In ActiveDocument.Fields
        If fldRider.Type = wdFieldIf Then
            fldRider.Code.Paragraphs(1).Range.Delete
        End If
    Next fldRider

ExitHere:
    Application.ScreenUpdating = True
    Set mobjContract = Nothing
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCrLf & .Description, vbCritical + vbOKOnly, "Clean_Up_Riders"
    End With
    Resume ExitHere
End Sub

Public Sub Clean_Up_Riders_OLD()
'************************************************************************************************
'Author: Adrian Jones
'Date Created: 07/22/2015
'This routine exists in the class as well as here. The purpose for it HERE is to tie a keystroke
'to this routine.
'Alt C, R for C = Clean Up, R = Rider
'************************************************************************************************
    Dim fld As Field
    Dim rngRiderTitleParagraph As Range
    Dim rngRiderStart As Range
    Dim strRiderTitle As String
    Dim intRiderFieldCount As Integer
    Dim para As Paragraph
    Dim paras As Paragraphs
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    For Each fld In ActiveDocument.Fields
        'Only look at "Rider" Fields
        If fld.Type = wdFieldIf Then
            intRiderFieldCount = intRiderFieldCount + 1
            Set rngRiderTitleParagraph = fld.Result.Paragraphs(1).Range
            Set rngRiderStart = rngRiderTitleParagraph.Next(Unit:=wdParagraph, Count:=1)
            strRiderTitle = rngRiderTitleParagraph.Text
            If fld.Result.Paragraphs.Count = 1 Then
                'Debug.Print "Rider Title: " & strRiderTitle & "Paragraphs Count: " & fld.Result.Paragraphs.Count
            Else
                'Full justify all paragraphs
                Set paras = fld.Result.Paragraphs
                fld.Unlink
                For Each para In paras
                    para.Range.Select
                    If para.Alignment = wdAlignParagraphLeft Then
                        para.Alignment = wdAlignParagraphJustify
                    End If
                Next
                'Debug.Print "Active Rider: " & strRiderTitle & "With a Paragraph Count of: " & fld.Result.Paragraphs.Count
                'Debug.Print "UNLINKED FIELD/RIDER " & strRiderTitle
                'Place a page break before the rider
                rngRiderStart.ParagraphFormat.PageBreakBefore = True
            End If
            'Debug.Print
            rngRiderTitleParagraph.Delete
            'Debug.Print "DELETED FIELD/RIDER " & strRiderTitle
        End If
    Next fld
    'Debug.Print "Count of 'Rider Fields' " & intRiderFieldCount
    
ExitHere:
    Application.ScreenUpdating = True
    Set rngRiderTitleParagraph = Nothing
    Set rngRiderStart = Nothing
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCrLf & .Description, vbCritical + vbOKOnly, "Clean_Up_Riders_Old"
    End With
    Resume ExitHere
End Sub

Public Sub GetRevisions()
    Dim objRevision As Revision
    
    On Error GoTo ErrorHandler
    
    'Enumerate Revisions for the current document
    For Each objRevision In ActiveDocument.Revisions
        
    Next objRevision

ExitHere:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCrLf & .Description, vbCritical + vbOKOnly, "Clean_Up_Riders"
    End With
    Resume ExitHere
End Sub
