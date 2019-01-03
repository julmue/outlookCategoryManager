Option Explicit

Public Function isTruePrefix(s1 As String, s2 As String) As Boolean
    If Len(s1) = Len(s2) Then
        isTruePrefix = False
    Else
        If InStr(1, s2, s1) = 1 Then
            isTruePrefix = True
        Else
            isTruePrefix = False
        End If
    End If
End Function

Public Function replaceSepByWs(s As String) As String
    Dim tmp As String
    tmp = Replace(s, ";", " ")
    tmp = Replace(tmp, ",", " ")
    replaceSepByWs = tmp
End Function

Public Function normalizeWs(s As String) As String
    Dim tmp As String
    tmp = Trim(s)
    tmp = removeDupeWs(tmp)
    normalizeWs = tmp
End Function

Public Function removeDupeWs(s As String) As String
    Dim tmp As String
    removeDupeWs = s
    Do
        tmp = removeDupeWs
        removeDupeWs = Replace(removeDupeWs, Space(2), Space(1))
    Loop Until tmp = removeDupeWs
End Function
