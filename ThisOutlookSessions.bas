Option Explicit

Public Sub CategoryManager()
    Dim s As String

    s = InputBox("Action:", "Category Manager", "a")
    s = Trim(s)
    
    Select Case s
    Case "a"
        Call addItemCategories
    Case "o"
        Call overwriteItemCategories
    Case "d"
        Call deleteItemCategories
    Case "D"
        Call deleteAllItemCategories
    Case "s"
        Call showItemCategories
    Case "A"
        Call addCategories
    Case "S"
        Call showCategories
    Case ""
        MsgBox "No Command"
    Case Else
        MsgBox "Unknown Command"
    End Select
End Sub

Public Sub addCategories()
    Dim objNameSpace As NameSpace
    Dim cat1 As category
    Dim cat2 As category
    Dim newCategoryNames As Collection
    Dim cn1 As Variant
    Dim cn2 As Variant
    Dim cnParts() As String
    Dim cnParent As String
    Dim levels As Integer
    Dim i As Integer

    Dim s As String
    s = InputBox("Categories:")

    Set newCategoryNames = parseCategoryNames(s)
    
    If newCategoryNames.Count = 0 Then
        MsgBox "No Category to be added"
    End If
    
    Collections.sort newCategoryNames

    Set objNameSpace = Application.GetNamespace("MAPI")

    For Each cn1 In newCategoryNames
        For Each cat1 In objNameSpace.Categories
            If cn1 = cat1.Name Then
                MsgBox "Category " & cn1 & " already defined ... exiting"
                Exit Sub
            End If
        Next
    Next
        
    For Each cn2 In newCategoryNames
        cnParts = Split(cn2, ".")
        levels = UBound(cnParts()) - LBound(cnParts()) + 1
        
        If levels = 1 Then
            MsgBox ("Root Categories have to be added manually!")
            Exit Sub
        End If
        
        For i = 0 To i = levels - 1
            If i = 0 Then
                cnParent = cnParts(i)
            Else
                cnParent = cnParent & "." & cnParts(i)
            End If
        Next i
        
        For Each cat2 In objNameSpace.Categories
            If cat2.Name = cnParent Then
                objNameSpace.Categories.Add cn2, cat2.Color
                Exit For
            End If
        Next
    Next
End Sub

Public Sub addItemCategories()
  Dim coll As Collection
  Dim obj As Object
  Dim s As String

  Set coll = getCurrentItems
  If coll.Count = 0 Then Exit Sub

  s = InputBox("Categories:")

  For Each obj In coll
    Call Categories.addItemCategories(obj, s)
  Next
End Sub

Public Sub overwriteItemCategories()
  Dim col As Collection
  Dim obj As Object
  Dim s As String

  Set col = getCurrentItems
  If col.Count = 0 Then Exit Sub

  s = InputBox("Categories:")

  For Each obj In col
    Call Categories.deleteAllItemCategories(obj)
    Call Categories.addItemCategories(obj, s)
  Next
End Sub

Public Sub showItemCategories()
  Dim col As Collection

  Set col = getCurrentItems
  If col.Count = 0 Then Exit Sub
  
  ' call for first object only
  MsgBox "Categories:" & vbCrLf & vbCrLf & Categories.showItemCategories(col(1))
End Sub

' shows first 10 categories with given prefix
Public Sub showCategories()
    Dim objNameSpace As NameSpace
    Dim objCategory As category
    Dim cats As String
    Dim prefix As String
    Dim strOutput As String
  
    Dim childCategories As Collection
    
    Set childCategories = New Collection
    
    cats = InputBox("Categories:")
    
    prefix = Categories.parseCategoryNames(cats)(1)
  
    ' Obtain a NameSpace object reference.
    Set objNameSpace = Application.GetNamespace("MAPI")
    
    If objNameSpace.Categories.Count = 0 Then
        MsgBox "No Categories in NameSpace"
        Exit Sub
    End If
  
    For Each objCategory In objNameSpace.Categories
        If Strings.isTruePrefix(prefix, objCategory.Name) Then
            childCategories.Add objCategory.Name
        End If
    Next
    
    If childCategories.Count() = 0 Then
        MsgBox "No Child Categories for Prefix: " & prefix
        Exit Sub
    End If
    
    Collections.sort childCategories
    
    Dim i As Integer
    i = 1
    
    While i <= childCategories.Count() And i <= 10
        strOutput = strOutput & childCategories(i) & vbCrLf
        i = i + 1
    Wend

    MsgBox "Subcategories: " & vbCrLf & vbCrLf & strOutput
End Sub

Public Sub deleteAllItemCategories()
  Dim coll As VBA.Collection
  Dim obj As Object
  Dim s$

  Set coll = getCurrentItems
  If coll.Count = 0 Then Exit Sub

  For Each obj In coll
    Call Categories.deleteAllItemCategories(obj)
  Next
End Sub

Public Sub deleteItemCategories()
  Dim coll As VBA.Collection
  Dim obj As Object
  Dim s As String

  Set coll = getCurrentItems
  If coll.Count = 0 Then Exit Sub
  
  s = InputBox("Categories:")
  
  For Each obj In coll
    Call Categories.deleteItemCategories(obj, s)
  Next
End Sub

' ----------------------------------------------------------
' Util: Next Action
' Needs some refactoring

Public Sub NextActionManager()
    Dim s As String

    s = InputBox("Action:", "Next Action Manager", "a")
    s = Trim(s)
    
    Select Case s
    Case "a"
        Call setNextAction
    Case "d"
        Call deleteNextAction
    Case "s"
        Call getNextAction
    Case ""
        MsgBox "No Command"
    Case Else
        MsgBox "Unknown Command"
    End Select
End Sub

Public Sub setNextAction()
  Dim coll As Collection
  Dim it As MailItem
  Dim s As String

  Set coll = getCurrentItems
  If coll.Count = 0 Then Exit Sub
  
  s = InputBox("Next Action:")
  
  For Each it In coll
    it.Mileage = s
    it.Save
  Next
End Sub

Public Sub deleteNextAction()
  Dim coll As Collection
  Dim it As MailItem

  Set coll = getCurrentItems
  If coll.Count = 0 Then Exit Sub
  
  For Each it In coll
    it.Mileage = ""
    it.Save
  Next
End Sub

Public Sub getNextAction()
  Dim coll As Collection
  Dim it As MailItem
  Dim s As String

  Set coll = getCurrentItems
  If coll.Count = 0 Then Exit Sub
  
  For Each it In coll
    MsgBox it.Mileage
  Next
End Sub

Public Sub moveToProc()
    Dim objNameSpace As Outlook.NameSpace
    Dim inbox As Outlook.Folder
    Dim target As Outlook.Folder
    Dim coll As Collection
    Dim it As MailItem

    Set objNameSpace = Application.GetNamespace("MAPI")
    Set inbox = objNameSpace.GetDefaultFolder(olFolderInbox)
    Set target = inbox.Parent.Folders("proc")
    
    Set coll = getCurrentItems
    If coll.Count = 0 Then
        Exit Sub
    End If
  
    For Each it In coll
        it.Move target
    Next
End Sub
                                                        
                                                        
' ----------------------------------------------------------
' Util: Get Items in Scope

Private Function getCurrentItems() As VBA.Collection
  Dim coll As VBA.Collection
  Dim Win As Object
  Dim Sel As Outlook.Selection
  Dim obj As Object
  Dim i&

  Set coll = New VBA.Collection
  Set Win = Application.ActiveWindow

  If TypeOf Win Is Outlook.Inspector Then
    coll.Add Win.CurrentItem
  Else
    Set Sel = Win.Selection
    If Not Sel Is Nothing Then
      For i = 1 To Sel.Count
        coll.Add Sel(i)
      Next
    End If
  End If
  Set getCurrentItems = coll
End Function


