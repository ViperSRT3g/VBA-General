Attribute VB_Name = "mod_HTMLXML"
Option Explicit

Public Function FilterHTML(ByVal RawHTML As String) As String
    If Len(RawHTML) = 0 Then Exit Function
    Dim HTMLEntities As Variant, HTMLReplacements As Variant, Index As Long
    Const REG_HTMLTAGS = "(<[\w\s""':.=-]*>|<\/[\w\s""':.=-]*>)"
    HTMLEntities = Array("&nbsp;", "&lt;", "&gt;", "&amp;", "&quot;", "&apos;")
    HTMLReplacements = Array(" ", "<", ">", "&", """", "'")
    'Parse HTML Entities into plaintext
    For Index = 0 To UBound(HTMLEntities)
        RawHTML = Replace(RawHTML, HTMLEntities(Index), HTMLReplacements(Index))
    Next Index
    'Remove any stray HTML tags
    Dim TargetTags() As String: TargetTags = RegExer(RawHTML, REG_HTMLTAGS)
    'Preemptively remove new line characters with actual new lines to separate any conjoined lines.
    RawHTML = Replace(RawHTML, "<b>", " ")
    For Index = 0 To UBound(TargetTags)
        RawHTML = Replace(RawHTML, TargetTags(Index), "")
    Next Index
    FilterHTML = RawHTML
End Function

Public Function ParseJSON(ByVal JSON As String) As Object
    With CreateObject("ScriptControl")
        .Language = "JScript"
        Set ParseJSON = .Eval("(" & JSON & ")")
    End With
End Function

Public Function GetHTTP(ByVal URL As String) As String
    On Error Resume Next
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        .Open "GET", URL, False
        .Send
        GetHTTP = .ResponseText
    End With
End Function

Public Function GetXML(ByVal URL As String) As String
    On Error Resume Next
    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", URL, False
        .Send
        GetXML = .ResponseText
    End With
End Function

Public Function GetHTML(ByVal URL As String) As String
    On Error GoTo ErrorEscape
    Dim IE As Object: Set IE = CreateObject("InternetExplorer.Application")
    Dim HTML As Object
    
    IE.navigate URL
    Do Until IE.ReadyState = 4: DoEvents: Loop
    Do While IE.Busy: DoEvents: Loop
    Set HTML = IE.Document.Body
    GetHTML = HTML.innerHTML
    IE.Quit
    
ErrorEscape:
    Set IE = Nothing
    Set HTML = Nothing
End Function

Public Function CheckHTTP(ByVal URL As String) As Integer
    On Error GoTo ErrorEscape
    Dim URLReq As Object: Set URLReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    If UCase(Left(URL, 4)) <> "HTTP" Then URL = "HTTP://" & URL
    URLReq.Open "GET", URL
    URLReq.Send
    CheckHTTP = CInt(URLReq.Status)
ErrorEscape:
    Set URLReq = Nothing
End Function
