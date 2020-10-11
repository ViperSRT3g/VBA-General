Attribute VB_Name = "mod_Email"
Option Explicit

Public Sub GenerateEmail(ByVal ToRecipient As String, ByVal EmailSubject As String, ByVal EmailBody As String, _
                              Optional AutoSend As Boolean, Optional CCRecipient As String, Optional BCCRecipient As String)
    With CreateObject("Outlook.Application")
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
        Set OutMail = Nothing
    End With
End Sub
