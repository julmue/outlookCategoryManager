Option Explicit
Option Base 1

' ------------------------------------------------------------------------------
' https://stackoverflow.com/questions/3587662/how-do-i-sort-a-collection
' Add faster sorting Algorithm later - this is bubble sort

Public Sub sort(ByRef col As Collection)
'    Dim col As Collection
    Dim vItm As Variant
    Dim i As Long, j As Long
    Dim vTemp As Variant
    
    Set col = distinct(col)

    For i = 1 To col.Count - 1
        For j = i + 1 To col.Count
            If col(i) > col(j) Then
                'store the lesser item
                vTemp = col(j)
                'remove the lesser item
                col.Remove j
                're-add the lesser item before the
                'greater Item
                col.Add vTemp, vTemp, i
            End If
        Next j
    Next i

End Sub

' https://stackoverflow.com/questions/3587662/how-do-i-sort-a-collection
' Returns an array which exactly matches this collection.
' Note: This function is not safe for concurrent modification.
Public Function toArray(col As Collection) As Variant
    Dim a() As Variant
    ReDim a(0 To col.Count)
    Dim i As Long
    For i = 0 To col.Count - 1
        a(i) = col(i + 1)
    Next i
    toArray = a()
End Function

' https://stackoverflow.com/questions/3587662/how-do-i-sort-a-collection
' Returns a Collection which exactly matches the given Array
' Note: This function is not safe for concurrent modification.
Public Function fromArray(a() As Variant) As Collection
    Dim col As Collection
    Set col = New Collection
    Dim element As Variant
    For Each element In a
        col.Add element
    Next element
    Set fromArray = col
End Function

'Returns True if the Collection contains an element equal to value
Public Function contains(value As Variant, col As Collection) As Boolean
    contains = (indexOf(value, col) >= 0)
End Function


'Returns the first index of an element equal to value. If the Collection
'does not contain such an element, returns -1.
Public Function indexOf(value As Variant, col As Collection) As Long

    Dim index As Long
    
    For index = 1 To col.Count Step 1
        If col(index) = value Then
            indexOf = index
            Exit Function
        End If
    Next index
    indexOf = -1
End Function

' get disctinct elements in collection
Public Function distinct(ByRef col As Collection) As Collection
    Dim acc As Collection
    Dim v As Variant
    
    Set acc = New Collection
    
    For Each v In col
        If Not contains(v, acc) Then
            acc.Add v
        End If
    Next

    Set distinct = acc
End Function

' compute the set difference col1 - col2
Public Function setdiff(ByRef col1 As Collection, ByRef col2 As Collection) As Collection
    Dim acc As Collection
    Dim v As Variant
    
    Set acc = New Collection
    
    For Each v In col1
        If Not contains(v, col2) Then
            acc.Add v
        End If
    Next

    Set setdiff = acc
End Function

Public Function setunion(ByRef col1 As Collection, ByRef col2 As Collection) As Collection
    Dim acc As Collection
    Dim v1 As Variant
    Dim v2 As Variant
    
    Set acc = New Collection
    
    For Each v1 In col1
        acc.Add v1
    Next
    
    For Each v2 In col2
        acc.Add v2
    Next
    
    Set acc = distinct(acc)

    Set setunion = acc
End Function
