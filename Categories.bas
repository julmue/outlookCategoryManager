Option Explicit

Private Sub ListCategoryIDs()
    Dim objNameSpace As NameSpace
    Dim objCategory As category
    Dim strOutput As String
 
    ' Obtain a NameSpace object reference.
    Set objNameSpace = Application.GetNamespace("MAPI")
 
    ' Check if the Categories collection for the Namespace
    ' contains one or more Category objects.
    If objNameSpace.Categories.Count > 0 Then
 
        ' Enumerate the Categories collection.
        For Each objCategory In objNameSpace.Categories
 
        ' Add the name and ID of the Category object to
        ' the output string.
        strOutput = strOutput & objCategory.Name & ": " & objCategory.CategoryID & vbCrLf
        Next
    End If
 
    ' Clean up.
    Set objCategory = Nothing
    Set objNameSpace = Nothing
 
End Sub

Public Function showItemCategories(ByRef obj As Object) As String
    Dim cats As Collection
    Dim cn As Variant
    Dim strOutput As String
    
    Set cats = parseCategoryNames(obj.Categories)
    
    For Each cn In cats
        strOutput = strOutput & cn & vbCrLf
    Next
    
     showItemCategories = strOutput
End Function

' Add category to object with category property
Public Sub addItemCategories(ByRef obj As Object, catsToAdd As String, _
                             Optional expanded As Boolean = False)
    Dim objNameSpace As NameSpace
    Dim objCategories As Collection
    Dim objCategory As category
    Dim v As Variant
    
    Dim acc As Collection
    Dim expansion As Collection
    Dim contraction As Collection
    
    Set objCategories = New Collection
    Set acc = distinct(parseCategoryNames(obj.Categories & "; " & catsToAdd))
    Set contraction = normalizeCategoriesMin(acc)
    Set expansion = normalizeCategoriesMax(contraction)
     
    ' assertion: (expanded) categories have to be in namespace
    Set objNameSpace = Application.GetNamespace("MAPI")
    For Each objCategory In objNameSpace.Categories
        objCategories.Add objCategory.Name
    Next
    
    For Each v In expansion
        If Not Collections.contains(v, objCategories) Then
            MsgBox "Unknown category: " & v
            Exit Sub
        End If
    Next
    ' end assertion

    If expanded Then
        obj.Categories = renderCategoryNames(expansion)
    Else
        obj.Categories = renderCategoryNames(contraction)
    End If
    obj.Save
End Sub

Public Sub deleteAllItemCategories(ByRef obj As Object)
    obj.Categories = ""
    obj.Save
End Sub


Public Sub deleteItemCategories(ByRef obj As Object, catsToDelete As String)
    Dim objCatsOld As Collection
    Dim delCats As Collection
    Dim objCatsNew As Collection
    
    
    Set objCatsOld = parseCategoryNames(obj.Categories)
    Set delCats = parseCategoryNames(catsToDelete)
    
    Set objCatsNew = setdiff(objCatsOld, delCats)
    
    obj.Categories = renderCategoryNames(objCatsNew)
    obj.Save
End Sub


' Parse category names separated by whitespace, ',' or ';'
Public Function parseCategoryNames(s As String) As Collection
    Dim cats() As String
    Dim acc As Collection
    Dim numcats As Integer
    Dim i As Integer
    Set acc = New Collection
    
    s = Strings.replaceSepByWs(s)
    s = Strings.normalizeWs(s)
    
    cats() = Split(s, " ")
    numcats = UBound(cats()) - LBound(cats()) + 1
    
    
    While i < numcats
        acc.Add cats(i)
        i = i + 1
    Wend
    
    Set parseCategoryNames = acc
End Function

' render category names
Public Function renderCategoryNames(ByRef catNames As Collection) As String
    Dim acc As String
    Dim i As Integer
    
    For i = 1 To catNames.Count
        If i = 1 Then
            acc = catNames(i)
        Else
            acc = acc & "; " & catNames(i)
        End If
    Next i
    
    renderCategoryNames = acc
End Function

' Get new category names
Public Function getNewCategories(ByRef catNamesExist As Collection, ByRef catNamesAdd As Collection) As Collection
    Set getNewCategories = Collections.setdiff(catNamesAdd, catNamesExist)
End Function

' Maximal Normal Form of Categories:
' * only leave / deepest categories
Public Function normalizeCategoriesMax(catNames As Collection) As Collection
    Dim col1 As Collection
    Dim col2 As Collection
    Dim res As Collection
    Dim v1 As Variant
    Dim v2 As Variant
    Dim cn1 As String
    Dim cn2 As String
    
    Set col1 = catNames
    Set col2 = New Collection
    Set res = New Collection
    Set col1 = normalizeCategoriesMin(col1)

    For Each v1 In col1
        cn1 = v1
        Set col2 = expandCategory(cn1)
        
        For Each v2 In col2
            cn2 = v2
            res.Add cn2
        Next
    Next
            
    Collections.sort res
    
    Set normalizeCategoriesMax = res

End Function

' Minimal Normal Form of Categories:
' * only leave / deepest categories
Public Function normalizeCategoriesMin(catNames As Collection) As Collection
    Dim col As Collection
    
    Set col = catNames
    
    Set col = contractCategories(col)
    Collections.sort col
    
    Set normalizeCategoriesMin = col
End Function


' Expand category name to all subcategories
' * Premis: category name already normalized (no leading/trailing whitespace ...)
Public Function expandCategory(catName As String) As Collection

    Debug.Assert Len(catName) = Len(Trim(catName))
    
    Dim col As Collection
    Set col = New Collection

    Dim levels() As String
    Dim depth As Integer
    Dim i As Integer

    Dim acc As String

    levels = Split(catName, ".")

    ' number of elements in a vba array is not trivial
    depth = UBound(levels) - LBound(levels) + 1

    If depth < 1 Then
        Exit Function
    End If

    acc = levels(0)
    col.Add acc

    ' -- work with an accumulator and backinsert
    For i = 1 To (depth - 1)
        acc = acc & "." & levels(i)
        If Not Collections.contains(acc, col) Then
            col.Add acc
        End If
    Next i
    
    Set col = Collections.distinct(col)
    Collections.sort col

    Set expandCategory = col

End Function


' Contract categories by category names module inheritance relation
' * keeps only leaves
' * returns categories in order
Public Function contractCategories(catNames As Collection) As Collection
    Dim res As Collection
    Dim v As Variant
    Dim cn As String

    Set res = New VBA.Collection

    For Each v In catNames
       cn = v
       If Not hasChildCategory(cn, catNames) And Not Collections.contains(cn, res) Then
          res.Add (cn)
      End If
    Next
    
    Set res = Collections.distinct(res)
    Collections.sort res
    
    Set contractCategories = res
    
End Function


' Detect child categories by category names and prefix relation
' Doesn't take into consideration '.' as separator
Public Function hasChildCategory(catName As String, ByRef catNames As Collection) As Boolean
    Dim v As Variant
    Dim cn As String
    
    hasChildCategory = False
    
    For Each v In catNames
        cn = v
        If Strings.isTruePrefix(catName, cn) Then
            hasChildCategory = True
            Exit For
        End If
    Next
End Function
