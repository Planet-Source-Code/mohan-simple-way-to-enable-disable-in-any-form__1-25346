Attribute VB_Name = "Module1"
Option Explicit
Public ctl As Control
'Enable controls
'Enable controls
Public Sub EnabDisabCtl(f As Form, Enab As Boolean)
For Each ctl In f.Controls
    If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is ListBox Or TypeOf ctl Is CheckBox Or TypeOf ctl Is OptionButton Then
        ctl.Enabled = Enab
    End If
Next
End Sub

'Controls names
Public Function CtlList(f As Form) As String
For Each ctl In f.Controls
    If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is ListBox Or TypeOf ctl Is CheckBox Or TypeOf ctl Is OptionButton Or TypeOf ctl Is MSFlexGrid Then
        If Len(CtlList) = 0 Then
            CtlList = ctl.Name
        Else
            CtlList = CtlList & "," & ctl.Name
        End If
    End If
Next

End Function

