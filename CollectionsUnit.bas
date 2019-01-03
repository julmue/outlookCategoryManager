Option Explicit

Private Sub runner()
    Call sortUnit1
    Call sortUnit2
    Call sortUnit3
    
    Call toArrayUnit1
    
    Call distinctUnit1
    Call distinctUnit2
    Call distinctUnit3
    
    Call setdiffUnit1
    Call setdiffUnit2
    Call setdiffUnit3
End Sub

Private Sub sortUnit1()
    Dim col As Collection
    Set col = New Collection

    col.Add "1"
    Collections.sort col
    
    Debug.Assert col.Item(1) = "1"
End Sub

Private Sub sortUnit2()
    Dim col As Collection
    Set col = New Collection

    col.Add "2"
    col.Add "1"
    Collections.sort col
    
    Debug.Assert col.Item(1) = "1"
    Debug.Assert col.Item(2) = "2"
End Sub

Private Sub sortUnit3()
    Dim col As Collection
    Set col = New Collection

    col.Add "p.duran.projects"
    col.Add "p.duran"
    col.Add "@x"
    col.Add "p"
    Collections.sort col
    
    Debug.Assert col.Item(1) = "@x"
    Debug.Assert col.Item(2) = "p"
    Debug.Assert col.Item(3) = "p.duran"
    Debug.Assert col.Item(4) = "p.duran.projects"
End Sub

Private Sub toArrayUnit1()
    Dim col As Collection
    Set col = New Collection
    Dim a As Variant

    col.Add "1"
    a = toArray(col)
    
    Debug.Assert a(0) = "1"
End Sub

Private Sub fromArrayUnit1()
    Dim col As Collection
    Dim i(0) As Variant

    i(0) = "1"
    Set col = Collections.fromArray(i)
    
    Debug.Assert col.Count = 1
    Debug.Assert col.Item(1) = "1"
End Sub

Private Sub distinctUnit1()
    Dim col As Collection
    Dim res As Collection
    
    Set col = New Collection

    col.Add "1"
    Set res = Collections.distinct(col)
    
    Debug.Assert res(1) = "1"
End Sub

Private Sub distinctUnit2()
    Dim col As Collection
    Dim res As Collection
    
    Set col = New Collection

    col.Add "1"
    col.Add "2"
    Set res = Collections.distinct(col)
    
    Debug.Assert res.Count = 2
    Debug.Assert Collections.contains("1", res)
    Debug.Assert Collections.contains("2", res)
End Sub

Private Sub distinctUnit3()
    Dim col As Collection
    Dim res As Collection
    
    Set col = New Collection

    col.Add "1"
    col.Add "1"
    Set res = Collections.distinct(col)
    
    Debug.Assert res.Count = 1
    Debug.Assert Collections.contains("1", res)
End Sub

Private Sub setdiffUnit1()
    Dim col1 As Collection
    Dim col2 As Collection
    Dim res As Collection
    
    Set col1 = New Collection
    Set col2 = New Collection

    col1.Add "1"
    
    Set res = Collections.setdiff(col1, col2)
    
    Debug.Assert res.Count = 1
    Debug.Assert Collections.contains("1", res)
End Sub

Private Sub setdiffUnit2()
    Dim col1 As Collection
    Dim col2 As Collection
    Dim res As Collection
    
    Set col1 = New Collection
    Set col2 = New Collection

    col1.Add "1"
    col2.Add "1"
    
    Set res = Collections.setdiff(col1, col2)
    
    Debug.Assert res.Count = 0
End Sub

Private Sub setdiffUnit3()
    Dim col1 As Collection
    Dim col2 As Collection
    Dim res As Collection
    
    Set col1 = New Collection
    Set col2 = New Collection

    col1.Add "1"
    col1.Add "2"
    col2.Add "1"
    
    Set res = Collections.setdiff(col1, col2)
    
    Debug.Assert res.Count = 1
    Debug.Assert Collections.contains("2", res)
End Sub

Private Sub setunionUnit1()
    Dim col1 As Collection
    Dim col2 As Collection
    Dim res As Collection
    
    Set col1 = New Collection
    Set col2 = New Collection
    
    Set res = Collections.setdiff(col1, col2)
    
    Debug.Assert res.Count = 0
End Sub

Private Sub setunionUnit2()
    Dim col1 As Collection
    Dim col2 As Collection
    Dim res As Collection
    
    Set col1 = New Collection
    Set col2 = New Collection
    
    col1.Add "1"
    
    Set res = Collections.setdiff(col1, col2)
    
    Debug.Assert res.Count = 1
End Sub

Private Sub setunionUnit3()
    Dim col1 As Collection
    Dim col2 As Collection
    Dim res As Collection
    
    Set col1 = New Collection
    Set col2 = New Collection
    
    col1.Add "1"
    col2.Add "1"
    
    Set res = Collections.setunion(col1, col2)
    
    Debug.Assert res.Count = 1
    Debug.Assert Collections.contains("1", res)
End Sub

Private Sub setunionUnit4()
    Dim col1 As Collection
    Dim col2 As Collection
    Dim res As Collection
    
    Set col1 = New Collection
    Set col2 = New Collection
    
    col1.Add "1"
    col2.Add "2"
    
    Set res = Collections.setunion(col1, col2)
    
    Debug.Assert res.Count = 2
    Debug.Assert Collections.contains("1", res)
    Debug.Assert Collections.contains("2", res)
End Sub

Private Sub setunionUnit5()
    Dim col1 As Collection
    Dim col2 As Collection
    Dim res As Collection
    
    Set col1 = New Collection
    Set col2 = New Collection
    
    col1.Add "1"
    col1.Add "2"
    col2.Add "2"
    col2.Add "3"
    
    Set res = Collections.setunion(col1, col2)
    
    Debug.Assert res.Count = 3
    Debug.Assert Collections.contains("1", res)
    Debug.Assert Collections.contains("2", res)
    Debug.Assert Collections.contains("3", res)
End Sub
