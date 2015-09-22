VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSFConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Login()
    Dim sfSession As SForceOfficeToolkitLib4.SForceSession4
    
    On Error GoTo ErrorHandler

    Set sfSession = New SForceSession4
    
    sfSession.Login "ajones@collegboard.org", "WimWenders6"
    
'    ServerUrl = sfSession.ServerUrl
'    SessionId = sfSession.SessionId
    
    MsgBox sfSession.ServerUrl + " " + sfSession.SessionId
    

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        .Raise .Number & vbCr & .Description, "Login"
    End With
End Sub
