Attribute VB_Name = "modMain"
Option Explicit
Dim mclsContent As clsCoverLetterVariables
Private Const mstrDOCVAR_PROGRAM As String = "Program"

'Public Sub AutoNew()
'    On Error GoTo ErrorHandler
'
'    Main
'
'Exit_Here:
'    Exit Sub
'
'ErrorHandler:
'    With Err
'        MsgBox .Number & vbCr & .Description, vbCritical, "AutoNew"
'    End With
'    Resume Exit_Here
'End Sub

Public Sub Main()
    On Error GoTo ErrorHandler
    
    frmCoverLetter.RetrieveDocumentVariables = True
    frmCoverLetter.Show
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "Main"
    End With
    Resume Exit_Here
End Sub

Public Sub RetrieveDocumentVariables()
    On Error GoTo ErrorHandler
    
    frmCoverLetter.Show
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "RetrieveDocumentVariables"
    End With
    Resume Exit_Here
End Sub

Public Sub CreateDocumentVariables()
    On Error GoTo ErrorHandler
    
    Set mclsContent = New clsCoverLetterVariables
    mclsContent.CreateDocumentVariables
    
Exit_Here:
    Set mclsContent = Nothing
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "Create Document Variables"
    End With
    Resume Exit_Here
End Sub

Public Sub DocVarExist()
    Dim objDocVar As Variable
    On Error GoTo ErrorHandler
    
    For Each objDocVar In ActiveDocument.Variables
        If objDocVar.Name = mstrDOCVAR_PROGRAM Then
            Debug.Print "THIS IS THE ONE I WANT ----> " & objDocVar.Name
        Else
            Debug.Print objDocVar.Name
        End If
    Next objDocVar

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub

Public Sub ShowDocumentVariables()
    On Error GoTo ErrorHandler
    
    frmDocumentVariables.Show

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "ShowDocumentVariables"
    End With
    Resume Exit_Here
End Sub
