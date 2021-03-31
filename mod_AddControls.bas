Attribute VB_Name = "mod_AddControls"
Option Explicit

Public Enum ControlType
    CheckBoxControl = 1
    ComboBoxControl = 2
    CommandButtonControl = 3
    FrameControl = 4
    ImageControl = 5
    LabelControl = 6
    ListBoxControl = 7
    MultiPageControl = 8
    OptionButtonControl = 9
    ScrollBarControl = 10
    SpinButtonControl = 11
    TabStripControl = 12
    TextBoxControl = 13
    ToggleButtonControl = 14
End Enum

Public Function AddControl(ByRef TargetForm As UserForm, TargetControl As ControlType, Optional ControlName As String, Optional Visibility As Boolean) As Control
    If (TargetForm Is Nothing) Or (TargetControl = 0) Then Exit Function
    If Len(ControlName) = 0 Then
        Set AddControl = TargetForm.Controls.Add(GetProgID(TargetControl), Visible:=Visibility)
    Else
        Set AddControl = TargetForm.Controls.Add(GetProgID(TargetControl), ControlName, Visibility)
    End If
End Function

Private Function GetProgID(ByVal TargetControl As ControlType) As String
    Select Case TargetControl
        Case CheckBoxControl: GetProgID = "Forms.CheckBox.1"
        Case ComboBoxControl: GetProgID = "Forms.ComboBox.1"
        Case CommandButtonControl: GetProgID = "Forms.CommandButton.1"
        Case FrameControl: GetProgID = "Forms.Frame.1"
        Case ImageControl: GetProgID = "Forms.Image.1"
        Case LabelControl: GetProgID = "Forms.Label.1"
        Case ListBoxControl: GetProgID = "Forms.ListBox.1"
        Case MultiPageControl: GetProgID = "Forms.MultiPage.1"
        Case OptionButtonControl: GetProgID = "Forms.OptionButton.1"
        Case ScrollBarControl: GetProgID = "Forms.ScrollBar.1"
        Case SpinButtonControl: GetProgID = "Forms.SpinButton.1"
        Case TabStripControl: GetProgID = "Forms.TabStrip.1"
        Case TextBoxControl: GetProgID = "Forms.TextBox.1"
        Case ToggleButtonControl: GetProgID = "Forms.ToggleButton.1"
    End Select
End Function
