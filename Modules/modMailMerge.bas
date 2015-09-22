Attribute VB_Name = "modMailMerge"
Option Explicit

Public Sub ConnectToPivotalDataInExcel()
    On Error GoTo ErrorHandler
    
    'SAMPLE CODE from Help
    'With docNew.MailMerge
    '    .MainDocumentType = wdFormLetters
    '    .OpenDataSource _
    '    Name:="C:\Program Files\Microsoft Office\Office" & _
    '    "\Samples\Northwind.mdb", _
    '    LinkToSource:=True, AddToRecentFiles:=False, _
    '    Connection:="TABLE Customers"
    '    MsgBox .DataSource.ConnectString
    'End With

'    ActiveDocument.MailMerge.DataSource.ConnectString
    
    
    
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "ConnectToPivotalDataInExcel"
    End With
    Resume Exit_Here
End Sub
