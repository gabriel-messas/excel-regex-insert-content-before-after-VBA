Attribute VB_Name = "Module1"
Function RegxFunc(strInput As String, regexPattern As String, insertBefore As String, insertAfter As String) As String
    Dim regEx As New RegExp
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = regexPattern
    End With

    If regEx.Test(strInput) Then
        Set matches = regEx.Execute(strInput)
        RegxFunc = insertBefore & matches(0).Value & insertAfter
    Else
        RegxFunc = "not matched"
    End If
End Function
