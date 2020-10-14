VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "class_Email"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Static class that allows sending multiple emails without having to create an Outlook.Application each time."
'@ModuleDescription "Static class that allows sending multiple emails without having to create an Outlook.Application each time."
'@PredeclaredId
Option Explicit

Private mMailApp As Object
Private Const APP_NAME As String = "Outlook.Application"

Public Sub GenerateEmail( _
    ByVal ToRecipient As String, _
    ByVal EmailSubject As String, _
    ByVal EmailBody As String, _
    Optional ByVal AutoSend As Boolean, _
    Optional ByVal CCRecipient As String, _
    Optional ByVal BCCRecipient As String _
)
    If mMailApp Is Nothing Then Set mMailApp = CreateObject(APP_NAME)
    With mMailApp
        Dim OutMail As Object: Set OutMail = .CreateItem(0)
        With OutMail
            .To = ToRecipient
            .CC = CCRecipient
            .BCC = BCCRecipient
            .Subject = EmailSubject
            .Body = EmailBody
            If AutoSend Then
                .Send
            Else
                .Display
            End If
        End With
    End With
End Sub

Public Sub Destroy()
    On Error GoTo ErrorHandler
    If Not mMailApp Is Nothing Then
        mMailApp.Quit
        Set mMailApp = Nothing
    End If

    Exit Sub
ErrorHandler:
    If Err.Number > 0 Then
        Debug.Print _
            "Error Disposing of " & APP_NAME; ":" & vbCrLf _
            & Err.Description
        Err.Clear
        Resume Next
    End If
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub


