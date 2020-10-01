Attribute VB_Name = "mTest"
Option Explicit
Option Private Module

Private dctTest As Dictionary

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = ThisWorkbook.Name & " mTest." & sProc
End Function

Public Sub Test_DctAdd_00_Regression()
    
    Const PROC = "Test_DctAdd_Regression"
    BoP ErrSrc(PROC)
    
    Test_DctAdd_01_KeyIsValue
    Test_DctAdd_02_KeyIsObjectWithNameProperty
    Test_DctAdd_03_ItemIsObjectWithNameProperty
    Test_DctAdd_04_InsertKeyBefore
    Test_DctAdd_05_InsertKeyAfter
    Test_DctAdd_06_InsertItemBefore
    Test_DctAdd_07_InsertItemAfter
    Test_DctAdd_08_NumKey
'    Test_DctAdd_09_StayWithFirst_Key
'    Test_DctAdd_10_StayWithFirst_Item
'    Test_DctAdd_11_UpdateFirst_Key
    Test_DctAdd_12_AddDuplicate_Item
    
    EoP ErrSrc(PROC)

End Sub

Private Sub Test_DctAdd_12_AddDuplicate_Item()
' ---------------------------------------------------------------
' When add criteria is by item, the item already exists but with
' a different key and staywithfirst = False (the default) the
' item is added.
' ---------------------------------------------------------------
    Const PROC = "Test_DctAdd_12_AddDuplicate_Item"

    Set dctTest = Nothing
    DctAdd dct:=dctTest, key:="A", item:=60, order:=order_byitem, seq:=seq_ascending
    DctAdd dct:=dctTest, key:="BB", item:=50, order:=order_byitem, seq:=seq_ascending
    DctAdd dct:=dctTest, key:="CCC", item:=30, order:=order_byitem, seq:=seq_ascending
    DctAdd dct:=dctTest, key:="DDDD", item:=30, order:=order_byitem, seq:=seq_ascending
    DctAdd dct:=dctTest, key:="EEEEE", item:=20, order:=order_byitem, seq:=seq_ascending
    DctAdd dct:=dctTest, key:="FFFFFF", item:=10, order:=order_byitem, seq:=seq_ascending
    Test_DctAdd_DisplayResult dctTest, "staywithfirst=False"
    Debug.Assert dctTest.Count = 6
    
    Set dctTest = Nothing
    DctAdd dct:=dctTest, key:="A", item:=60, order:=order_byitem, seq:=seq_ascending, staywithfirst:=True
    DctAdd dct:=dctTest, key:="BB", item:=50, order:=order_byitem, seq:=seq_ascending, staywithfirst:=True
    DctAdd dct:=dctTest, key:="CCC", item:=30, order:=order_byitem, seq:=seq_ascending, staywithfirst:=True
    DctAdd dct:=dctTest, key:="DDDD", item:=30, order:=order_byitem, seq:=seq_ascending, staywithfirst:=True
    DctAdd dct:=dctTest, key:="EEEEE", item:=20, order:=order_byitem, seq:=seq_ascending, staywithfirst:=True
    DctAdd dct:=dctTest, key:="FFFFFF", item:=10, order:=order_byitem, seq:=seq_ascending, staywithfirst:=True
    Test_DctAdd_DisplayResult dctTest, "staywithfirst=True"
    Debug.Assert dctTest.Count = 5
    
End Sub

Private Sub Test_DctAdd_01_KeyIsValue()
' -----------------------------------------------
' Note: Since a 100% reverse key order added in mode ascending
' is the worst case regarding performance this test sorts 100 items
' with 50% already in seq and the other 50% to be inserted.
' -----------------------------------------------
    Const PROC = "Test_DctAdd_01_KeyIsValue"
    Dim i   As Long
    
    BoP ErrSrc(PROC)
    Set dctTest = Nothing
    For i = 1 To 9 Step 2
        DctAdd dct:=dctTest, key:=i, item:=ThisWorkbook, seq:=seq_ascending ' by key case sensitive is the default
    Next i
    For i = 10 To 2 Step -2
        DctAdd dct:=dctTest, key:=i, item:=ThisWorkbook, seq:=seq_ascending ' by key case sensitive is the default
    Next i
    
    '~~ Add an already existing key, ignored when the item is neither numeric nor a string
    DctAdd dct:=dctTest, key:=5, item:=ThisWorkbook, seq:=seq_ascending ' by key case sensitive is the default
    
    EoP ErrSrc(PROC)
    
End Sub

Private Sub Test_DctAdd_02_KeyIsObjectWithNameProperty()
' ----------------------------------------------------
' Added items with a key which is an object.
' The order by key uses the object's name property.
' ----------------------------------------------------
    Const PROC = "Test_DctAdd_02_KeyIsObjectWithNameProperty"
    Dim i   As Long
    Dim vbc As VBComponent
    
    BoP ErrSrc(PROC)
    Set dctTest = Nothing
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        DctAdd dct:=dctTest, key:=vbc, item:=vbc.Name, seq:=seq_ascending ' by key case sensitive is the default
    Next vbc
    Debug.Assert dctTest.Count = 9
    Debug.Assert dctTest.Items()(0) = "clsCallStack"
    Debug.Assert dctTest.Items()(dctTest.Count - 1) = "wsDct"
    
    '~~ Add an already existing key = update the item
    Set vbc = ThisWorkbook.VBProject.VBComponents("mTest")
    DctAdd dct:=dctTest, key:=vbc, item:=vbc.Name, seq:=seq_ascending ' by key case sensitive is the default
    Debug.Assert dctTest.Count = 9
    Debug.Assert dctTest.Items()(0) = "clsCallStack"
    Debug.Assert dctTest.Items()(dctTest.Count - 1) = "wsDct"
    EoP ErrSrc(PROC)
        
End Sub

Private Sub Test_DctAdd_03_ItemIsObjectWithNameProperty()
' ----------------------------------------------------
' Added items with a key which is an object.
' The order by key uses the object's name property.
' ----------------------------------------------------
    Const PROC = "Test_DctAdd_03_ItemIsObjectWithNameProperty"
    Dim i   As Long
    Dim vbc As VBComponent
    
    BoP ErrSrc(PROC)
    Set dctTest = Nothing
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        DctAdd dct:=dctTest, key:=vbc.Name, item:=vbc, order:=order_byitem, seq:=seq_ascending
    Next vbc
    Debug.Assert dctTest.Count = 9
    Test_DctAdd_DisplayResult dctTest
    Debug.Assert dctTest.Items()(0).Name = "clsCallStack"
    Debug.Assert dctTest.Items()(dctTest.Count - 1).Name = "wsDct"
    
    '~~ Add an already existing key = update the item
    Set vbc = ThisWorkbook.VBProject.VBComponents("mTest")
    DctAdd dct:=dctTest, key:=vbc.Name, item:=vbc, order:=order_byitem, seq:=seq_ascending
    Debug.Assert dctTest.Count = 9
    Debug.Assert dctTest.Items()(0).Name = "clsCallStack"
    Debug.Assert dctTest.Items()(dctTest.Count - 1).Name = "wsDct"
    EoP ErrSrc(PROC)
        
End Sub

Private Sub Test_DctAdd_04_InsertKeyBefore()
    
    Const PROC = "Test_DctAdd_04_InsertKeyBefore"
    Dim vbc_second As VBComponent
    Dim vbc_first As VBComponent
    
    BoP ErrSrc(PROC)
    
    '~~ Preparation
    Test_DctAdd_02_KeyIsObjectWithNameProperty
    Debug.Assert dctTest.Keys()(0).Name = "clsCallStack"
    Debug.Assert dctTest.Keys()(1).Name = "clsCallStackItem"
    Set vbc_second = ThisWorkbook.VBProject.VBComponents("clsCallStackItem")
    Set vbc_first = ThisWorkbook.VBProject.VBComponents("clsCallStack")
    dctTest.Remove vbc_second
    Debug.Assert dctTest.Count = 8
    
    '~~ Test
    DctAdd dctTest, vbc_second, vbc_second.Name, order:=order_bykey, seq:=seq_beforetarget, target:=vbc_first
    Debug.Assert dctTest.Count = 9
    Debug.Assert dctTest.Keys()(0).Name = "clsCallStackItem"
    Debug.Assert dctTest.Keys()(1).Name = "clsCallStack"
    EoP ErrSrc(PROC)
    
End Sub

Private Sub Test_DctAdd_05_InsertKeyAfter()
    
    Const PROC = "Test_DctAdd_05_InsertKeyAfter"
    Dim vbc_second As VBComponent
    Dim vbc_first As VBComponent
    
    BoP ErrSrc(PROC)
    
    '~~ Preparation
    Test_DctAdd_02_KeyIsObjectWithNameProperty
    Debug.Assert dctTest.Keys()(0).Name = "clsCallStack"
    Debug.Assert dctTest.Keys()(1).Name = "clsCallStackItem"
    Set vbc_first = ThisWorkbook.VBProject.VBComponents("clsCallStack")
    Set vbc_second = ThisWorkbook.VBProject.VBComponents("clsCallStackItem")
    
    '~~ Test
    dctTest.Remove vbc_first
    Debug.Assert dctTest.Count = 8
    DctAdd dctTest, key:=vbc_first, item:=vbc_first.Name, order:=order_bykey, seq:=seq_aftertarget, target:=vbc_second
    Debug.Assert dctTest.Count = 9
    Debug.Assert dctTest.Keys()(0).Name = "clsCallStackItem"
    Debug.Assert dctTest.Keys()(1).Name = "clsCallStack"
    EoP ErrSrc(PROC)
    
End Sub

Private Sub Test_DctAdd_06_InsertItemBefore()
    
    Const PROC = "Test_DctAdd_06_InsertItemBefore"
    Dim vbc_second As VBComponent
    Dim vbc_first As VBComponent
    
    BoP ErrSrc(PROC)
    
    '~~ Preparation
    Test_DctAdd_03_ItemIsObjectWithNameProperty
    Debug.Assert dctTest.Keys()(0) = "clsCallStack"
    Debug.Assert dctTest.Keys()(1) = "clsCallStackItem"
    Set vbc_second = ThisWorkbook.VBProject.VBComponents("clsCallStackItem")
    Set vbc_first = ThisWorkbook.VBProject.VBComponents("clsCallStack")
    dctTest.Remove vbc_second.Name
    Debug.Assert dctTest.Count = 8
    
    '~~ Test
    DctAdd dctTest, vbc_second.Name, vbc_second, order:=order_byitem, seq:=seq_beforetarget, target:=vbc_first
    Debug.Assert dctTest.Count = 9
    Debug.Assert dctTest.Items()(0).Name = "clsCallStackItem"
    Debug.Assert dctTest.Items()(1).Name = "clsCallStack"
    EoP ErrSrc(PROC)
    
End Sub

Private Sub Test_DctAdd_07_InsertItemAfter()
    
    Const PROC = "Test_DctAdd_07_InsertItemAfter"
    Dim vbc_second As VBComponent
    Dim vbc_first As VBComponent
    
    BoP ErrSrc(PROC)
    
    '~~ Preparation
    Test_DctAdd_03_ItemIsObjectWithNameProperty
    Debug.Assert dctTest.Keys()(0) = "clsCallStack"
    Debug.Assert dctTest.Keys()(1) = "clsCallStackItem"
    Set vbc_second = ThisWorkbook.VBProject.VBComponents("clsCallStackItem")
    Set vbc_first = ThisWorkbook.VBProject.VBComponents("clsCallStack")
    dctTest.Remove vbc_first.Name
    Debug.Assert dctTest.Count = 8
    
    '~~ Test
    DctAdd dctTest, vbc_first.Name, vbc_first, order:=order_byitem, seq:=seq_aftertarget, target:=vbc_second
    Debug.Assert dctTest.Count = 9
    Debug.Assert dctTest.Items()(0).Name = vbc_second.Name
    Debug.Assert dctTest.Items()(1).Name = vbc_first.Name
    EoP ErrSrc(PROC)
    
End Sub

Private Sub Test_DctAdd_08_NumKey()
    Const PROC = "Test_DctAdd_08_NumKey"
    BoP ErrSrc(PROC)
    Set dctTest = Nothing
    
    DctAdd dctTest, 2, 5, seq:=seq_ascending
    DctAdd dctTest, 5, 2, seq:=seq_ascending
    DctAdd dctTest, 3, 4, seq:=seq_ascending
    
    Debug.Assert dctTest.Count = 3
    Debug.Assert dctTest.Keys()(0) = 2
    Debug.Assert dctTest.Keys()(dctTest.Count - 1) = 5
    
    EoP ErrSrc(PROC)
    
End Sub

Private Sub Test_DctAdd_09_Performance_n(ByVal lAdds As Long)
    Const PROC = "Test_DctAdd_09_Performance_n"
    Dim i   As Long
    
    BoP ErrSrc(PROC)
    Set dctTest = Nothing
    For i = 1 To lAdds - 1 Step 2
        DctAdd dct:=dctTest, key:=i, item:=ThisWorkbook, seq:=seq_ascending ' by key case sensitive is the default
    Next i
    For i = lAdds To 2 Step -2
        DctAdd dct:=dctTest, key:=i, item:=ThisWorkbook, seq:=seq_ascending ' by key case sensitive is the default
    Next i
    EoP ErrSrc(PROC)
    
End Sub

Private Sub Test_DctAdd_99_Performance()
    Const PROC = "Test_DctAdd_09_Performance"
    
    BoP ErrSrc(PROC)
    
    Test_DctAdd_09_Performance_n 100
    Test_DctAdd_09_Performance_n 500
    Test_DctAdd_09_Performance_n 1000
    Test_DctAdd_09_Performance_n 1500
    Test_DctAdd_09_Performance_n 2000
    Test_DctAdd_09_Performance_n 2500
    Test_DctAdd_09_Performance_n 3000
    
    EoP ErrSrc(PROC)
    
End Sub

Public Sub Test_DctAdd_DisplayResult( _
           ByVal dct As Dictionary, _
  Optional ByVal s As String)
' -----------------------------------------
    Dim v           As Variant
    Dim sKey        As String
    Dim sItem       As String
    Dim lKeyMax     As Long
    Dim lItemMax    As Long
    
    For Each v In dct
        If VarType(v) = vbObject Then sKey = v.Name Else sKey = v
        lKeyMax = Max(lKeyMax, Len(sKey))
        If VarType(dct.item(v)) = vbObject Then sItem = dct.item(v).Name Else sItem = dct.item(v)
        lItemMax = Max(lItemMax, Len(sItem))
    Next v
    
    Debug.Print ">> ----- " & s & " --------------"
    For Each v In dct
        If VarType(v) = vbObject Then sKey = v.Name Else sKey = v
        If VarType(dct.item(v)) = vbObject Then sItem = dct.item(v).Name Else sItem = dct.item(v)
        Debug.Print "Key: '" & sKey & "'," & Space(lKeyMax - Len(sKey)) & " Item: '" & sItem & "'"
    Next v
    Debug.Print "<< ----- " & s & " --------------"
    
End Sub

Private Sub Test_DctDiffers()
' -------------------------------------------
' Precondition: DctAdd is tested
' -------------------------------------------
    Const PROC = "Test_DctDiffers"
    Dim dct1 As Dictionary
    Dim dct2 As Dictionary
    Dim vbc  As VBComponent
    
    BoP ErrSrc(PROC)
    Set dct1 = Nothing
    Set dct2 = Nothing
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        DctAdd dct:=dct1, key:=vbc, item:=vbc, seq:=seq_ascending ' by key case sensitive is the default
    Next vbc
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        DctAdd dct:=dct2, key:=vbc, item:=vbc, seq:=seq_ascending ' by key case sensitive is the default
    Next vbc
    
    '~~ Test: Differs in keys
    Debug.Assert Not DctDiffers(dct1, dct2)
    dct1.Remove ThisWorkbook.VBProject.VBComponents("mTest")
    dct2.Remove ThisWorkbook.VBProject.VBComponents("mBasic")
    Debug.Assert DctDiffers(dct1, dct2)
    Set dct1 = Nothing
    Set dct2 = Nothing
    EoP ErrSrc(PROC)
    
End Sub

