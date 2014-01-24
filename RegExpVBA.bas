Attribute VB_Name = "Module1"
' search regular expression in Value based on Pattern, return place of first occurence or false
Function SearchRX(Value As String, Pattern As String, Optional IgnoreCase As Boolean = False)
    Dim r As New VBScript_RegExp_55.RegExp
    Dim result As MatchCollection
    r.Pattern = Pattern
    r.IgnoreCase = IgnoreCase
    If r.Test(Value) Then
        Set result = r.Execute(Value)
        SearchRX = result.Item(0).FirstIndex
    Else
        SearchRX = False
    End If
End Function
' search regular expression in Value based on Pattern, return first occurence or false
Function FirstRX(Value As String, Pattern As String, Optional IgnoreCase As Boolean = False)
    Dim r As New VBScript_RegExp_55.RegExp
    Dim result As MatchCollection
    r.Pattern = Pattern
    r.IgnoreCase = IgnoreCase
    If r.Test(Value) Then
        Set result = r.Execute(Value)
        FirstRX = result.Item(0).Value
    Else
        FirstRX = False
    End If
End Function
' search regular expression in Value based on Pattern, return Value up to first occurence or false
Function LeftRX(Value As String, Pattern As String, Optional IgnoreCase As Boolean = False)
    Dim r As New VBScript_RegExp_55.RegExp
    r.Pattern = Pattern
    r.IgnoreCase = IgnoreCase
    If r.Test(Value) Then
        Set result = r.Execute(Value)
        LeftRX = Left(Value, result.Item(0).FirstIndex)
    Else
        LeftRX = False
    End If
End Function
' search regular expression in Value based on Pattern, return number of counts or false
Function MatchRX(Value As String, Pattern As String, Optional IgnoreCase As Boolean = False)
    Dim r As New VBScript_RegExp_55.RegExp
    r.Pattern = Pattern
    r.IgnoreCase = IgnoreCase
    If r.Test(Value) Then
        Set result = r.Execute(Value)
        MatchRX = result.Count
    Else
        MatchRX = False
    End If
End Function
' search regular expression in Value based on Pattern, return string "Matches " and pattern or empty string
Function M(Value As String, Pattern As String, Optional IgnoreCase As Boolean = False)
    Dim r As New VBScript_RegExp_55.RegExp
    r.Pattern = Pattern
    r.IgnoreCase = IgnoreCase
    If r.Test(Value) Then
        M = "Matches '" & Pattern & "'"
    Else
        M = ""
    End If
End Function
' search regular expression in Value based on Pattern in the beginning, return string "Starts with " and pattern or empty string
Function StartsWith(Value As String, Pattern As String, Optional IgnoreCase As Boolean = False)
    Dim r As New VBScript_RegExp_55.RegExp
    r.Pattern = "^" & Pattern
    r.IgnoreCase = IgnoreCase
    If r.Test(Value) Then
        StartsWith = "Starts with '" & Pattern & "'"
    Else
        StartsWith = ""
    End If
End Function
' search regular expression in Value based on Pattern at the end, return string "Ends with " and pattern or empty string
Function EndsWith(Value As String, Pattern As String, Optional IgnoreCase As Boolean = False)
    Dim r As New VBScript_RegExp_55.RegExp
    r.Pattern = Pattern & "$"
    r.IgnoreCase = IgnoreCase
    If r.Test(Value) Then
        EndsWith = "Ends with '" & Pattern & "'"
    Else
        EndsWith = ""
    End If
End Function
' search regular expression in Value based on Pattern, return Value with Pattern replaced by ReplaceWith
Function S(Value As String, Pattern As String, ReplaceWith As String, Optional IgnoreCase As Boolean = False)
    Dim r As New VBScript_RegExp_55.RegExp
    r.Pattern = Pattern
    r.IgnoreCase = IgnoreCase
    r.Global = True
    S = r.Replace(Value, ReplaceWith)
End Function





