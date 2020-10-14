VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mod_Email"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Static class that allows sending multiple emails without having to create an Outlook.Application each time."
Option Explicit

private mMailApp As Object

Public Sub GenerateEmail(ByVal ToRecipient As String, ByVal EmailSubject As String, ByVal EmailBody As String, _
                              Optional AutoSend As Boolean, Optional CCRecipient As String, Optional BCCRecipient As String)
    Attribute GenerateEmail.VB_Description = "Sends an email from outlook."
    If mMailApp is Nothing Then Set mMailApp = CreateObject("Outlook.Application")
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
    Set mMailApp = Nothing
End Sub