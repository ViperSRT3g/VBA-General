Attribute VB_Name = "mod_DownloadFile"
Option Explicit

Private Declare Function URLDownloadToFile Lib "urlmon" _
    Alias "URLDownloadToFileA" _
    (ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) As Long

Private Declare Function DeleteUrlCacheEntry Lib "Wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long

Private Const ERROR_SUCCESS As Long = 0
Private Const BINDF_GETNEWESTVERSION As Long = &H10
Private Const INTERNET_FLAG_RELOAD As Long = &H80000000

Public Function DownloadFile(ByVal SourceURL As String, ByVal LocalFile As String) As Boolean
  'Download the file. BINDF_GETNEWESTVERSION forces the API to download from the specified source.
  'Passing 0& as dwReserved causes the locally-cached copy to be downloaded, if available. If the API
  'returns ERROR_SUCCESS (0), DownloadFile returns True.
   DownloadFile = URLDownloadToFile(0&, SourceURL, LocalFile, BINDF_GETNEWESTVERSION, 0&) = ERROR_SUCCESS
End Function

Public Function DownloadFileHTTP(ByVal SourceURL As String, ByVal LocalFile As String, Optional ByVal Username As String, Optional ByRef Password As String) As Boolean
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
End Function
