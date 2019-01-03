Option Explicit

Private Sub runner()
    Call hasChildCategoryUnit1
    Call hasChildCategoryUnit2
    
    Call contractCategoriesUnit1
    Call contractCategoriesUnit2
    Call contractCategoriesUnit3
    Call contractCategoriesUnit4
    Call contractCategoriesUnit5
    
    Call ExpandCategoryUnit1
    Call ExpandCategoryUnit2
    
    Call normalizeCategoriesMinUnit1
    
    Call normalizeCategoriesMaxUnit1
    Call normalizeCategoriesMaxUnit2
    Call normalizeCategoriesMaxUnit3
    
    Call getNewCategoriesUnit1
    Call getNewCategoriesUnit2
    Call getNewCategoriesUnit3
    Call getNewCategoriesUnit4
    Call getNewCategoriesUnit5
End Sub

Private Sub hasChildCategoryUnit1()
    Dim catNames As Collection
    Set catNames = New Collection
    
    catNames.Add "p"
    
    Debug.Assert Not Categories.hasChildCategory("p", catNames)
End Sub

Private Sub hasChildCategoryUnit2()
    Dim catNames As Collection
    Set catNames = New Collection
    
    catNames.Add "p"
    catNames.Add "p.duran"
 
    Debug.Assert Categories.hasChildCategory("p", catNames)
End Sub

Private Sub contractCategoriesUnit1()
    Dim catNames As Collection
    Dim contraction As Collection
    Set catNames = New Collection
    Dim t As Boolean
    
    catNames.Add "p"
    
    Set contraction = Categories.contractCategories(catNames)
    
    Debug.Assert Collections.contains("p", contraction)
End Sub

Private Sub contractCategoriesUnit2()
    Dim catNames As Collection
    Dim contraction As Collection
    Set catNames = New Collection
    
    catNames.Add "p"
    catNames.Add "p.duran"
    
    Set contraction = Categories.contractCategories(catNames)
    
    Debug.Assert Not Collections.contains("p", contraction)
    Debug.Assert Collections.contains("p.duran", contraction)
End Sub


Private Sub contractCategoriesUnit3()
    Dim catNames As Collection
    Dim contraction As Collection
    Set catNames = New Collection

    catNames.Add "p"
    catNames.Add "p.duran"
    catNames.Add "@x"
    
    Set contraction = Categories.contractCategories(catNames)
    
    Debug.Assert Not Collections.contains("p", contraction)
    Debug.Assert Collections.contains("p.duran", contraction)
    Debug.Assert Collections.contains("@x", contraction)
End Sub

Private Sub contractCategoriesUnit4()
    Dim catNames As Collection
    Dim contraction As Collection
    Set catNames = New Collection
    
    catNames.Add "p"
    catNames.Add "p.duran"
    catNames.Add "p.duran.kategorien"
    
    Set contraction = Categories.contractCategories(catNames)
    
    Debug.Assert Not Collections.contains("p", contraction)
    Debug.Assert Not Collections.contains("p.duran", contraction)
    Debug.Assert Collections.contains("p.duran.kategorien", contraction)
End Sub

Private Sub contractCategoriesUnit5()
    Dim catNames As Collection
    Dim contraction As Collection
    Set catNames = New Collection
    
    catNames.Add "p"
    catNames.Add "p.duran"
    catNames.Add "p.duran.kategorien"
    catNames.Add "p.duran.formulare"
    
    Set contraction = Categories.contractCategories(catNames)
    
    Debug.Assert Not Collections.contains("p", contraction)
    Debug.Assert Not Collections.contains("p.duran", contraction)
    Debug.Assert Collections.contains("p.duran.kategorien", contraction)
    Debug.Assert Collections.contains("p.duran.formulare", contraction)
End Sub

Private Sub ExpandCategoryUnit1()
    Dim expansion As Collection
    
    Set expansion = expandCategory("p")
    
    Debug.Assert expansion.Count = 1
    Debug.Assert Collections.contains("p", expansion)
End Sub

Private Sub ExpandCategoryUnit2()
    Dim expansion As Collection
    
    Set expansion = expandCategory("p.duran.categories")
    
    Debug.Assert expansion.Count = 3
    Debug.Assert Collections.contains("p", expansion)
    Debug.Assert Collections.contains("p.duran", expansion)
    Debug.Assert Collections.contains("p.duran.categories", expansion)
End Sub

Private Sub normalizeCategoriesMinUnit1()
    Dim catNames As Collection
    Dim minform As Collection
    Set catNames = New Collection
    
    catNames.Add "p"
    catNames.Add "p.duran"
    catNames.Add "p.duran.kategorien"
    catNames.Add "p.duran.formulare"
    catNames.Add "p.siemens"
    catNames.Add "p.siemens.buchen"
    catNames.Add "p.siemens.rechnung"
    
    Set minform = Categories.normalizeCategoriesMin(catNames)
    
    Debug.Assert minform.Count = 4
    Debug.Assert Not Collections.contains("p", minform)
    Debug.Assert Not Collections.contains("p.duran", minform)
    Debug.Assert Not Collections.contains("p.siemens", minform)
    Debug.Assert Collections.contains("p.duran.kategorien", minform)
    Debug.Assert Collections.contains("p.duran.formulare", minform)
    Debug.Assert Collections.contains("p.siemens.buchen", minform)
    Debug.Assert Collections.contains("p.siemens.rechnung", minform)
End Sub

Private Sub normalizeCategoriesMaxUnit1()
    Dim catNames As Collection
    Dim maxform As Collection
    Set catNames = New Collection
    
    catNames.Add "p"
    
    Set maxform = Categories.normalizeCategoriesMax(catNames)
    
    Debug.Assert maxform.Count = 1
    Debug.Assert Collections.contains("p", maxform)
End Sub

Private Sub normalizeCategoriesMaxUnit2()
    Dim catNames As Collection
    Dim maxform As Collection
    Set catNames = New Collection
    
    catNames.Add "p.duran"
    
    Set maxform = Categories.normalizeCategoriesMax(catNames)
    
    Debug.Assert maxform.Count = 2
    Debug.Assert Collections.contains("p", maxform)
    Debug.Assert Collections.contains("p.duran", maxform)
End Sub

Private Sub normalizeCategoriesMaxUnit3()
    Dim catNames As Collection
    Dim maxform As Collection
    Set catNames = New Collection
    
    catNames.Add "p.duran.kategorien"
    catNames.Add "p.duran.formulare"
    catNames.Add "p.siemens.buchen"
    catNames.Add "p.siemens.rechnung"
    
    Set maxform = Categories.normalizeCategoriesMax(catNames)
    
    Debug.Assert maxform.Count = 7
    Debug.Assert Collections.contains("p", maxform)
    Debug.Assert Collections.contains("p.duran", maxform)
End Sub

Private Sub getNewCategoriesUnit1()
    Dim catsExist As Collection
    Dim catsAdd As Collection
    Dim res As Collection
    
    Set catsExist = New Collection
    Set catsAdd = New Collection
    
    Set res = getNewCategories(catsExist, catsAdd)
    
    Debug.Assert res.Count = 0
End Sub

Private Sub getNewCategoriesUnit2()
    Dim catsExist As Collection
    Dim catsAdd As Collection
    Dim res As Collection
    
    Set catsExist = New Collection
    Set catsAdd = New Collection
    
    catsAdd.Add "1"
    
    Set res = getNewCategories(catsExist, catsAdd)
    
    Debug.Assert res.Count = 1
    Debug.Assert Collections.contains("1", res)
End Sub

Private Sub getNewCategoriesUnit3()
    Dim catsExist As Collection
    Dim catsAdd As Collection
    Dim res As Collection
    
    Set catsExist = New Collection
    Set catsAdd = New Collection
    
    catsExist.Add "1"
    
    Set res = getNewCategories(catsExist, catsAdd)
    
    Debug.Assert res.Count = 0
End Sub

Private Sub getNewCategoriesUnit4()
    Dim catsExist As Collection
    Dim catsAdd As Collection
    Dim res As Collection
    
    Set catsExist = New Collection
    Set catsAdd = New Collection
    
    catsExist.Add "1"
    catsExist.Add "2"
    
    catsAdd.Add "2"
    catsAdd.Add "3"
        
    Set res = getNewCategories(catsExist, catsAdd)
    
    Debug.Assert res.Count = 1
    Debug.Assert Collections.contains("3", res)
End Sub

Private Sub getNewCategoriesUnit5()
    Dim catsExist As Collection
    Dim catsAdd As Collection
    Dim res As Collection
    
    Set catsExist = New Collection
    Set catsAdd = New Collection
    
    catsExist.Add "1"
    catsExist.Add "2"
    
    catsAdd.Add "2"
    catsAdd.Add "3"
    catsAdd.Add "4"
        
    Set res = getNewCategories(catsExist, catsAdd)
    
    Debug.Assert res.Count = 2
    Debug.Assert Collections.contains("3", res)
    Debug.Assert Collections.contains("4", res)
End Sub

Private Sub renderCategoryNamesUnit1()
    Dim catNames As Collection
    
    Set catNames = New Collection
    
    Debug.Assert Categories.renderCategoryNames(catNames) = ""
End Sub

Private Sub renderCategoryNamesUnit2()
    Dim catNames As Collection
    Dim s As String
    
    Set catNames = New Collection
    
    catNames.Add "p"
    s = Categories.renderCategoryNames(catNames)
    
    Debug.Assert s = "p"
End Sub

Private Sub renderCategoryNamesUnit3()
    Dim catNames As Collection
    Dim s As String
    
    Set catNames = New Collection
    
    catNames.Add "p"
    catNames.Add "p.duran"
    s = Categories.renderCategoryNames(catNames)
    
    Debug.Assert s = "p; p.duran"
End Sub

Private Sub parseCategoryNamesUnit1()
    Dim catNames As Collection
    
    Set catNames = Categories.parseCategoryNames("")
    
    Debug.Assert catNames.Count = 0
End Sub

Private Sub parseCategoryNamesUnit2()
    Dim catNames As Collection
    
    Set catNames = Categories.parseCategoryNames("p")
    
    Debug.Assert catNames.Count = 1
    Debug.Assert Collections.contains("p", catNames)
End Sub

Private Sub parseCategoryNamesUnit3()
    Dim catNames As Collection
    
    Set catNames = Categories.parseCategoryNames("p q")
    
    Debug.Assert catNames.Count = 2
    Debug.Assert Collections.contains("p", catNames)
    Debug.Assert Collections.contains("q", catNames)
End Sub

Private Sub parseCategoryNamesUnit4()
    Dim catNames As Collection
    
    Set catNames = Categories.parseCategoryNames("p; q")
    
    Debug.Assert catNames.Count = 2
    Debug.Assert Collections.contains("p", catNames)
    Debug.Assert Collections.contains("q", catNames)
End Sub

Private Sub parseCategoryNamesUnit5()
    Dim catNames As Collection
    
    Set catNames = Categories.parseCategoryNames("p, q")
    
    Debug.Assert catNames.Count = 2
    Debug.Assert Collections.contains("p", catNames)
    Debug.Assert Collections.contains("q", catNames)
End Sub

Private Sub parseCategoryNamesUnit6()
    Dim catNames As Collection
    
    Set catNames = Categories.parseCategoryNames(" p,  q;; l")
    
    Debug.Assert catNames.Count = 3
    Debug.Assert Collections.contains("p", catNames)
    Debug.Assert Collections.contains("q", catNames)
    Debug.Assert Collections.contains("l", catNames)
End Sub

Private Sub addItemCategoriesUnit1()
    Dim cats As Collection
    Dim it As MailItem
    Dim s As String
    Set it = Application.CreateItem(olMailItem)
    Set cats = New Collection
    
    it.Categories = ""
    it.Save

    Call Categories.addItemCategories(it, "")

    Debug.Assert it.Categories = ""
    
End Sub

Private Sub addItemCategoriesUnit2()
'    Dim cats As Collection
    Dim it As MailItem
    Dim s As String
    Set it = Application.CreateItem(olMailItem)
'    Set cats = New Collection
    
    it.Categories = "p; c"
    it.Save

    Call Categories.addItemCategories(it, "")

    Debug.Assert it.Categories = "c; p"
End Sub

Private Sub addItemCategoriesUnit3()
    Dim it As MailItem
    Dim s As String
    Set it = Application.CreateItem(olMailItem)
    
    it.Categories = "p; c"
    it.Save

    Call Categories.addItemCategories(it, "@w")

    Debug.Assert it.Categories = "@w; c; p"
End Sub

Private Sub addItemCategoriesUnit4()
    Dim it As MailItem
    Dim s As String
    Set it = Application.CreateItem(olMailItem)

    Call Categories.addItemCategories(it, "@w p p.duran p.duran.formulare")

    Debug.Assert it.Categories = "@w; p.duran.formulare"
End Sub

Private Sub addItemCategoriesUnit5()
    Dim it As MailItem
    Dim s As String
    Set it = Application.CreateItem(olMailItem)

    Call Categories.addItemCategories(it, "@w p.duran p p.duran.formulare")

    Debug.Assert it.Categories = "@w; p.duran.formulare"
End Sub
