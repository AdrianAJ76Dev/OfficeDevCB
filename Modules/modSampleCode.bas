Attribute VB_Name = "modSampleCode"
Option Explicit

Public Sub TestClass()
    Dim objClassTested As clsCBContract
    On Error GoTo ErrorHandler
    
    '*********************************************************************************************************
    Set objClassTested = New clsCBContract 'Change here to test any created class
'    objClassTested.DetermineActiveRiders
    objClassTested.ActivateRider
'    If objClassTested.DetermineIfRider(Selection.Range) Then
'        MsgBox "A Rider is selected", vbExclamation, "Testing Class Function"
'    Else
'        MsgBox "A Rider is NOT selected", vbExclamation, "Testing Class Function"
'    End If
'    objClassTested.DeleteUnncessaryRiders
    '*********************************************************************************************************

Exit_Here:
    Set objClassTested = Nothing
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "TestClass"
    End With
    Resume Exit_Here
End Sub

Public Sub TestSFConnect()
    Dim sfSession As SForceOfficeToolkitLib4.SForceSession4
    Dim blnLoginResult As Boolean
    
    Const strURL_CB_SERVER As String = "https://na10.salesforce.com/"
    Const strURL_DEV_SERVER As String = "https://na31.salesforce.com/"
    
    Const strURL_Login As String = "https://login.salesforce.com/"
    Const strURL_SOAP As String = "https://www.salesforce.com/services/Soap/c/23.0"
    Const strURL_TEST As String = "https://test.salesforce.com"
    Const strURL_TEST_1 As String = "https://test.salesforce.com/services/Soap/c/"
    Const strURL_TEST_2 As String = "https://test.salesforce.com/services/Soap/universal/"
    
    Const strUSER_ID As String = "ajones@collegboard.org"
    Const strPASSWORD As String = "WimWenders6"
    Const strUSER_ID_DEV As String = "adrianaj@msn.com"
    Const strPASSWORD_DEV As String = "WimWenders5"
    
    On Error GoTo ErrorHandler

    Set sfSession = New SForceSession4
    Debug.Print "Default URL: ", sfSession.ServerUrl
    sfSession.SetServerUrl strURL_DEV_SERVER & "services/Soap/universal/"
    Debug.Print "New URL: ", sfSession.ServerUrl
    
    blnLoginResult = sfSession.Login(strUSER_ID_DEV, strPASSWORD_DEV)
    
    
    Debug.Print blnLoginResult
    Debug.Print sfSession.ErrorMessage
    Debug.Print sfSession.Error
    
    If blnLoginResult Then
        MsgBox "You logged in!"
    Else
        MsgBox "You didn't Logon"
    End If
    
    
    
'    ServerUrl = sfSession.ServerUrl
'    SessionId = sfSession.SessionId
    
'    MsgBox sfSession.ServerUrl + " " + sfSession.SessionId
'    MsgBox sfSession

Exit_Here:
    Set sfSession = Nothing
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "TestSFConnect"
    End With
    Resume Exit_Here
End Sub

Sub Insert_Updated_Dates()
Attribute Insert_Updated_Dates.VB_Description = "Inserts updated, correct dates for the fiscal year into the riders"
Attribute Insert_Updated_Dates.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Insert_Updated_Dates"
'
' Insert_Updated_Dates Macro
' Inserts updated, correct dates for the fiscal year into the riders
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "September 14, 2013"
        .Replacement.Text = "September 13, 2013"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "September 14, 2013"
        .Replacement.Text = "September 13, 2013"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "September 3, 2013"
        .Replacement.Text = "September 2, 2013"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub Replace2012()
Attribute Replace2012.VB_Description = "Search and Replace 2012-2013 with 2013-2014"
Attribute Replace2012.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Replace2012"
'
' Replace2012 Macro
' Search and Replace 2012-2013 with 2013-2014
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "2012-2013"
        .Replacement.Text = "2013-2014"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub TableFormat()
Attribute TableFormat.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.TableFormat"
'
' TableFormat Macro
'
'
    Selection.Tables(1).Rows.Alignment = wdAlignRowLeft
    Selection.Tables(1).Rows.WrapAroundText = False
End Sub
