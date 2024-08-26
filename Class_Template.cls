VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_Template"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'MODULE DESCRIPTION
'A class module template that has basic amenities such as internet connectivity checks
'and a self-contained property enumeration that can be directly converted into JSON

'MODULE DEPENDENCIES
'JsonConverter.bas


'PUBLIC ENUMS


'PRIVATE ENUMS
Private Enum Prop
    ID_
    Name_
    [_First] = ID_
    [_Last] = Name_
End Enum


'PUBLIC CONSTS


'PRIVATE CONSTS


'PRIVATE DECLARATIONS
#If Win64 And VBA7 Then
    Private Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
#Else
    Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
#End If


'PUBLIC VARIABLES
Public Properties As Object


'PRIVATE VARIABLES


'CLASS PROCEDURES
Private Sub Class_Initialize()
    Set Properties = CreateObject("Scripting.Dictionary")
End Sub
Private Sub Class_Terminate()
    If Not Properties Is Nothing Then Properties.RemoveAll
    
    Set Properties = Nothing
End Sub


'CLASS PROPERTIES
Property Get ToJSON() As String
    On Error Resume Next
    ToJSON = ConvertToJson(Properties)
    On Error GoTo 0
End Property


'PUBLIC FUNCTIONS



'PRIVATE FUNCTIONS
Private Function GetHTTP(ByVal URL As String) As String
    On Error Resume Next
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        .Open "GET", URL, False
        .Send
        GetHTTP = .ResponseText
    End With
End Function

Private Function IsInternetConnected() As Boolean
    IsInternetConnected = InternetGetConnectedState(0&, 0&)
End Function

Private Function P(Value As Prop) As String
    Select Case Value
        Case Prop.ID_: P = "ID"
        Case Prop.Name_: P = "Name"
    End Select
End Function
