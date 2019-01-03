Option Explicit

Private Sub runner()
    Call removeDupeWsUnit1
    Call removeDupeWsUnit2
    Call removeDupeWsUnit3
    Call removeDupeWsUnit4
    Call removeDupeWsUnit5
    Call removeDupeWsUnit6
    Call removeDupeWsUnit7
    
    Call normalizeWsUnit1
    Call normalizeWsUnit2
    Call normalizeWsUnit3
    Call normalizeWsUnit4
    Call normalizeWsUnit5
    
    Call replaceSepByWsUnit1
    Call replaceSepByWsUnit2
    Call replaceSepByWsUnit3
    Call replaceSepByWsUnit4
    Call replaceSepByWsUnit5
    Call replaceSepByWsUnit6
    Call replaceSepByWsUnit7
    
    Call isTruePrefixUnit1
    Call isTruePrefixUnit2
    Call isTruePrefixUnit3
    Call isTruePrefixUnit4
    Call isTruePrefixUnit5
End Sub

Private Sub removeDupeWsUnit1()
    Dim s As String
    s = Strings.removeDupeWs("")
    
    Debug.Assert s = ""
End Sub

Private Sub removeDupeWsUnit2()
    Dim s As String
    s = Strings.removeDupeWs(" ")
    
    Debug.Assert s = " "
End Sub

Private Sub removeDupeWsUnit3()
    Dim s As String
    s = Strings.removeDupeWs("X")
    
    Debug.Assert s = "X"
End Sub

Private Sub removeDupeWsUnit4()
    Dim s As String
    s = Strings.removeDupeWs(" X ")
    
    Debug.Assert s = " X "
End Sub

Private Sub removeDupeWsUnit5()
    Dim s As String
    s = Strings.removeDupeWs("  X  ")
    
    Debug.Assert s = " X "
End Sub

Private Sub removeDupeWsUnit6()
    Dim s As String
    s = Strings.removeDupeWs("hello world")
    
    Debug.Assert s = "hello world"
End Sub

Private Sub removeDupeWsUnit7()
    Dim s As String
    s = Strings.removeDupeWs("hello  world")
    
    Debug.Assert s = "hello world"
End Sub

Private Sub normalizeWsUnit1()
    Dim s As String
    s = Strings.normalizeWs("")
    
    Debug.Assert s = ""
End Sub

Private Sub normalizeWsUnit2()
    Dim s As String
    s = Strings.normalizeWs(" ")
    
    Debug.Assert s = ""
End Sub

Private Sub normalizeWsUnit3()
    Dim s As String
    s = Strings.normalizeWs(" X ")
    
    Debug.Assert s = "X"
End Sub

Private Sub normalizeWsUnit4()
    Dim s As String
    s = Strings.normalizeWs(" hello world ")
    
    Debug.Assert s = "hello world"
End Sub

Private Sub normalizeWsUnit5()
    Dim s As String
    s = Strings.normalizeWs("  hello  world  ")
    
    Debug.Assert s = "hello world"
End Sub

Private Sub replaceSepByWsUnit1()
    Dim s As String
    s = Strings.replaceSepByWs("")
    
    Debug.Assert s = ""
End Sub

Private Sub replaceSepByWsUnit2()
    Dim s As String
    s = Strings.replaceSepByWs(" ")
    
    Debug.Assert s = " "
End Sub

Private Sub replaceSepByWsUnit3()
    Dim s As String
    s = Strings.replaceSepByWs(";")
    
    Debug.Assert s = " "
End Sub

Private Sub replaceSepByWsUnit4()
    Dim s As String
    s = Strings.replaceSepByWs("hello;")
    
    Debug.Assert s = "hello "
End Sub

Private Sub replaceSepByWsUnit5()
    Dim s As String
    s = Strings.replaceSepByWs("hello;world")
    
    Debug.Assert s = "hello world"
End Sub

Private Sub replaceSepByWsUnit6()
    Dim s As String
    s = Strings.replaceSepByWs("hello,")
    
    Debug.Assert s = "hello "
End Sub

Private Sub replaceSepByWsUnit7()
    Dim s As String
    s = Strings.replaceSepByWs("hello,world")
    
    Debug.Assert s = "hello world"
End Sub

Private Sub isTruePrefixUnit1()
    Dim b As Boolean
    b = Not Strings.isTruePrefix("", "")
    
    Debug.Assert b
End Sub

Private Sub isTruePrefixUnit2()
    Dim b As Boolean
    b = Strings.isTruePrefix("", " ")
    
    Debug.Assert b
End Sub

Private Sub isTruePrefixUnit3()
    Dim b As Boolean
    b = Not Strings.isTruePrefix(" ", " ")
    
    Debug.Assert b
End Sub

Private Sub isTruePrefixUnit4()
    Dim b As Boolean
    b = Not Strings.isTruePrefix("hello", "hello")
    
    Debug.Assert b
End Sub

Private Sub isTruePrefixUnit5()
    Dim b As Boolean
    b = Strings.isTruePrefix("hello", "hello.world")
    
    Debug.Assert b
End Sub
