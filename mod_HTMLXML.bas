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
    On Error GoTo ErrorHandler
    Dim IE As Object: Set IE = CreateObject("InternetExplorer.Application")
    Dim HTML As Object
    
    IE.navigate URL
    Do Until IE.ReadyState = 4: DoEvents: Loop
    Do While IE.Busy: DoEvents: Loop
    Set HTML = IE.Document.Body
    GetHTML = HTML.innerHTML
    IE.Quit
    
ErrorHandler:
    Set IE = Nothing
    Set HTML = Nothing
End Function

Public Function CheckHTTP(ByVal URL As String) As Integer
    On Error GoTo ErrorHandler
    Dim URLReq As Object: Set URLReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    If UCase(Left(URL, 4)) <> "HTTP" Then URL = "HTTP://" & URL
    URLReq.Open "GET", URL
    URLReq.Send
    CheckHTTP = CInt(URLReq.Status)
ErrorHandler:
    Set URLReq = Nothing
End Function

Public Function DownloadFileHTTP(ByVal SourceURL As String, ByVal LocalFile As String, Optional ByVal Username As String, Optional ByRef Password As String) As Boolean
    On Error GoTo ErrorHandler
    If Len(SourceURL) = 0 Or Len(LocalFile) = 0 Then Exit Function
    Dim WHTTPReq As Object: Set WHTTPReq = CreateObject("Microsoft.XMLHTTP")
    WHTTPReq.Open "GET", SourceURL, False, Username, Password
    WHTTPReq.Send
    
    If Not WHTTPReq.Status = 200 Then Exit Function
    With CreateObject("ADODB.Stream")
        .Open
        .Type = 1
        .Write WHTTPReq.responseBody
        .SaveToFile LocalFile, 2 ' 1 = no overwrite, 2 = overwrite
        .Close
    End With
    DownloadFileHTTP = CreateObject("Scripting.FileSystemObject").FileExists(LocalFile)
ErrorHandler:
    Set WHTTPReq = Nothing
End Function
