Attribute VB_Name = "mTest"
Option Explicit
Option Private Module

Public Type tMsgSection                 ' ---------------------
       sLabel As String                 ' Structure of the
       sText As String                  ' UserForm's message
       bMonspaced As Boolean            ' area which consists
End Type                                ' of 4 message sections
Public Type tMsg                        ' Attention: 4 is a
       section(1 To 4) As tMsgSection   ' design constant!
End Type                                ' ---------------------

Private dctTest As Dictionary

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = ThisWorkbook.Name & " mTest." & sProc
End Function

Public Sub Test_DctAdd_00_Regression()
    
    Const PROC = "Test_DctAdd_Regression"
    mTrc.BoP ErrSrc(PROC)
    
    Test_DctAdd_01_Performance_KeyIsValue
    Test_DctAdd_02_KeyIsObjectWithNameProperty
    Test_DctAdd_03_ItemIsObjectWithNameProperty
    Test_DctAdd_04_InsertKeyBefore
    Test_DctAdd_05_InsertKeyAfter
    Test_DctAdd_06_InsertItemBefore
    Test_DctAdd_07_InsertItemAfter
    Test_DctAdd_08_NumKey
    Test_DctAdd_12_AddDuplicate_Item
    Test_DctAdd_99_Performance
    
    mTrc.EoP ErrSrc(PROC)

End Sub

Private Sub Test_DctAdd_12_AddDuplicate_Item()
' ---------------------------------------------------------------
' When add criteria is by item, the item already exists but with
' a different key and staywithfirst = False (the default) the
' item is added.
' ---------------------------------------------------------------
    Const PROC = "Test_DctAdd_12_AddDuplicate_Item"

    Set dctTest = Nothing
    DctAdd add_dct:=dctTest, add_key:="A", add_item:=60, add_order:=order_byitem, add_seq:=seq_ascending
    DctAdd add_dct:=dctTest, add_key:="BB", add_item:=50, add_order:=order_byitem, add_seq:=seq_ascending
    DctAdd add_dct:=dctTest, add_key:="CCC", add_item:=30, add_order:=order_byitem, add_seq:=seq_ascending
    DctAdd add_dct:=dctTest, add_key:="DDDD", add_item:=30, add_order:=order_byitem, add_seq:=seq_ascending
    DctAdd add_dct:=dctTest, add_key:="EEEEE", add_item:=20, add_order:=order_byitem, add_seq:=seq_ascending
    DctAdd add_dct:=dctTest, add_key:="FFFFFF", add_item:=10, add_order:=order_byitem, add_seq:=seq_ascending
'    Test_DctAdd_DisplayResult dctTest, "staywithfirst=False"
    Debug.Assert dctTest.Count = 6
    
    Set dctTest = Nothing
    DctAdd add_dct:=dctTest, add_key:="A", add_item:=60, add_order:=order_byitem, add_seq:=seq_ascending, add_staywithfirst:=True
    DctAdd add_dct:=dctTest, add_key:="BB", add_item:=50, add_order:=order_byitem, add_seq:=seq_ascending, add_staywithfirst:=True
    DctAdd add_dct:=dctTest, add_key:="CCC", add_item:=30, add_order:=order_byitem, add_seq:=seq_ascending, add_staywithfirst:=True
    DctAdd add_dct:=dctTest, add_key:="DDDD", add_item:=30, add_order:=order_byitem, add_seq:=seq_ascending, add_staywithfirst:=True
    DctAdd add_dct:=dctTest, add_key:="EEEEE", add_item:=20, add_order:=order_byitem, add_seq:=seq_ascending, add_staywithfirst:=True
    DctAdd add_dct:=dctTest, add_key:="FFFFFF", add_item:=10, add_order:=order_byitem, add_seq:=seq_ascending, add_staywithfirst:=True
'    Test_DctAdd_DisplayResult dctTest, "staywithfirst=True"
    Debug.Assert dctTest.Count = 5
    
End Sub

Private Sub Test_DctAdd_01_Performance_KeyIsValue()
' -----------------------------------------------
' Note: Since a 100% reverse key order added in mode ascending
' is the worst case regarding performance this test sorts 100 items
' with 50% already in seq and the other 50% to be inserted.
' -----------------------------------------------
    Const PROC = "Test_DctAdd_01_Performance_KeyIsValue"
    Dim i       As Long
    Dim j       As Long: j = 999
    Dim jStep   As Long: jStep = 2
    Dim k       As Long: k = 1000
    Dim kStep   As Long: kStep = -2
    
    mTrc.BoP ErrSrc(PROC), "added items = ", k
    Set dctTest = Nothing
    For i = 1 To j Step jStep
        DctAdd add_dct:=dctTest, add_key:=i, add_item:=ThisWorkbook, add_seq:=seq_ascending ' by key case sensitive is the default
    Next i
    For i = k To jStep Step kStep
        DctAdd add_dct:=dctTest, add_key:=i, add_item:=ThisWorkbook, add_seq:=seq_ascending ' by key case sensitive is the default
    Next i
    
    '~~ Add an already existing key, ignored when the item is neither numeric nor a string
    DctAdd add_dct:=dctTest, add_key:=5, add_item:=ThisWorkbook, add_seq:=seq_ascending ' by key case sensitive is the default
     
    mTrc.EoP ErrSrc(PROC)
    
End Sub

Private Sub Test_DctAdd_02_KeyIsObjectWithNameProperty()
' ----------------------------------------------------
' Added items with a key which is an object.
' The order by key uses the object's name property.
' ----------------------------------------------------
    Const PROC = "Test_DctAdd_02_KeyIsObjectWithNameProperty"
    Dim i   As Long
    Dim vbc As VBComponent
    
    mTrc.BoP ErrSrc(PROC)
    Set dctTest = Nothing
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        DctAdd add_dct:=dctTest, add_key:=vbc, add_item:=vbc.Name, add_seq:=seq_ascending ' by key case sensitive is the default
    Next vbc
    Debug.Assert dctTest.Count = ThisWorkbook.VBProject.VBComponents.Count
    Debug.Assert dctTest.Items()(0) = "fMsg"
    Debug.Assert dctTest.Items()(dctTest.Count - 1) = "wsDct"
    
    '~~ Add an already existing key = update the item
    Set vbc = ThisWorkbook.VBProject.VBComponents("mTest")
    DctAdd add_dct:=dctTest, add_key:=vbc, add_item:=vbc.Name, add_seq:=seq_ascending ' by key case sensitive is the default
    Debug.Assert dctTest.Count = ThisWorkbook.VBProject.VBComponents.Count
    Debug.Assert dctTest.Items()(0) = "fMsg"
    Debug.Assert dctTest.Items()(dctTest.Count - 1) = "wsDct"
    mTrc.EoP ErrSrc(PROC)
        
End Sub

Private Sub Test_DctAdd_03_ItemIsObjectWithNameProperty()
' ----------------------------------------------------
' Added items with a key which is an object.
' The order by key uses the object's name property.
' ----------------------------------------------------
    Const PROC = "Test_DctAdd_03_ItemIsObjectWithNameProperty"
    Dim i   As Long
    Dim vbc As VBComponent
    
    mTrc.BoP ErrSrc(PROC)
    Set dctTest = Nothing
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        DctAdd add_dct:=dctTest, add_key:=vbc.Name, add_item:=vbc, add_order:=order_byitem, add_seq:=seq_ascending
    Next vbc
    Debug.Assert dctTest.Count = ThisWorkbook.VBProject.VBComponents.Count
'    Test_DctAdd_DisplayResult dctTest
    Debug.Assert dctTest.Items()(0).Name = "fMsg"
    Debug.Assert dctTest.Items()(dctTest.Count - 1).Name = "wsDct"
    
    '~~ Add an already existing key = update the item
    Set vbc = ThisWorkbook.VBProject.VBComponents("mTest")
    DctAdd add_dct:=dctTest, add_key:=vbc.Name, add_item:=vbc, add_order:=order_byitem, add_seq:=seq_ascending
    Debug.Assert dctTest.Count = ThisWorkbook.VBProject.VBComponents.Count
    Debug.Assert dctTest.Items()(0).Name = "fMsg"
    Debug.Assert dctTest.Items()(dctTest.Count - 1).Name = "wsDct"
    mTrc.EoP ErrSrc(PROC)
        
End Sub

Private Sub Test_DctAdd_04_InsertKeyBefore()
    
    Const PROC = "Test_DctAdd_04_InsertKeyBefore"
    Dim vbc_second As VBComponent
    Dim vbc_first As VBComponent
    
    mTrc.BoP ErrSrc(PROC)
    
    '~~ Preparation
    Test_DctAdd_02_KeyIsObjectWithNameProperty
    Debug.Assert dctTest.Keys()(0).Name = "fMsg"
    Debug.Assert dctTest.Keys()(1).Name = "mDct"
    Set vbc_second = ThisWorkbook.VBProject.VBComponents("mTrc")
    Set vbc_first = ThisWorkbook.VBProject.VBComponents("mTest")
    dctTest.Remove vbc_second
    Debug.Assert dctTest.Count = ThisWorkbook.VBProject.VBComponents.Count - 1
    
    '~~ Test
    DctAdd dctTest, vbc_second, vbc_second.Name, add_order:=order_bykey, add_seq:=seq_beforetarget, add_target:=vbc_first
    Debug.Assert dctTest.Count = ThisWorkbook.VBProject.VBComponents.Count
    Debug.Assert dctTest.Keys()(0).Name = "fMsg"
    Debug.Assert dctTest.Keys()(1).Name = "mDct"
    mTrc.EoP ErrSrc(PROC)
    
End Sub

Private Sub Test_DctAdd_05_InsertKeyAfter()
    
    Const PROC = "Test_DctAdd_05_InsertKeyAfter"
    Dim vbc_second As VBComponent
    Dim vbc_first As VBComponent
    
    mTrc.BoP ErrSrc(PROC)
    
    '~~ Preparation
    Test_DctAdd_02_KeyIsObjectWithNameProperty
    Debug.Assert dctTest.Keys()(0).Name = "fMsg"
    Debug.Assert dctTest.Keys()(1).Name = "mDct"
    Set vbc_first = ThisWorkbook.VBProject.VBComponents(1)
    Set vbc_second = ThisWorkbook.VBProject.VBComponents(2)
    
    '~~ Test
    dctTest.Remove vbc_first
    Debug.Assert dctTest.Count = ThisWorkbook.VBProject.VBComponents.Count - 1
    DctAdd dctTest, add_key:=vbc_first, add_item:=vbc_first.Name, add_order:=order_bykey, add_seq:=seq_aftertarget, add_target:=vbc_second
    Debug.Assert dctTest.Count = ThisWorkbook.VBProject.VBComponents.Count
    Debug.Assert dctTest.Keys()(0).Name = "fMsg"
    Debug.Assert dctTest.Keys()(1).Name = "mDct"
    mTrc.EoP ErrSrc(PROC)
    
End Sub

Private Sub Test_DctAdd_06_InsertItemBefore()
    
    Const PROC = "Test_DctAdd_06_InsertItemBefore"
    Dim vbc_second As VBComponent
    Dim vbc_first As VBComponent
    
    mTrc.BoP ErrSrc(PROC)
    
    '~~ Preparation
    Test_DctAdd_03_ItemIsObjectWithNameProperty
    Debug.Assert dctTest.Keys()(0) = "fMsg"
    Debug.Assert dctTest.Keys()(1) = "mDct"
    Set vbc_second = ThisWorkbook.VBProject.VBComponents("fMsg")
    Set vbc_first = ThisWorkbook.VBProject.VBComponents("mDct")
    dctTest.Remove vbc_second.Name
    Debug.Assert dctTest.Count = ThisWorkbook.VBProject.VBComponents.Count - 1
    
    '~~ Test
    DctAdd dctTest, vbc_second.Name, vbc_second, add_order:=order_byitem, add_seq:=seq_beforetarget, add_target:=vbc_first
    Debug.Assert dctTest.Count = ThisWorkbook.VBProject.VBComponents.Count
    Debug.Assert dctTest.Items()(0).Name = "fMsg"
    Debug.Assert dctTest.Items()(1).Name = "mDct"
    mTrc.EoP ErrSrc(PROC)
    
End Sub

Private Sub Test_DctAdd_07_InsertItemAfter()
    
    Const PROC = "Test_DctAdd_07_InsertItemAfter"
    Dim vbc_second As VBComponent
    Dim vbc_first As VBComponent
    
    mTrc.BoP ErrSrc(PROC)
    
    '~~ Preparation
    Test_DctAdd_03_ItemIsObjectWithNameProperty
    Debug.Assert dctTest.Keys()(0) = "fMsg"
    Debug.Assert dctTest.Keys()(1) = "mDct"
    Set vbc_second = ThisWorkbook.VBProject.VBComponents("mDct")
    Set vbc_first = ThisWorkbook.VBProject.VBComponents("fMsg")
    dctTest.Remove vbc_first.Name
    Debug.Assert dctTest.Count = ThisWorkbook.VBProject.VBComponents.Count - 1
    
    '~~ Test
    DctAdd dctTest, vbc_first.Name, vbc_first, add_order:=order_byitem, add_seq:=seq_aftertarget, add_target:=vbc_second
    Debug.Assert dctTest.Count = ThisWorkbook.VBProject.VBComponents.Count
    Debug.Assert dctTest.Items()(0).Name = vbc_second.Name
    Debug.Assert dctTest.Items()(1).Name = vbc_first.Name
    mTrc.EoP ErrSrc(PROC)
    
End Sub

Private Sub Test_DctAdd_08_NumKey()
    Const PROC = "Test_DctAdd_08_NumKey"
    mTrc.BoP ErrSrc(PROC)
    Set dctTest = Nothing
    
    DctAdd dctTest, 2, 5, add_seq:=seq_ascending
    DctAdd dctTest, 5, 2, add_seq:=seq_ascending
    DctAdd dctTest, 3, 4, add_seq:=seq_ascending
    
    Debug.Assert dctTest.Count = 3
    Debug.Assert dctTest.Keys()(0) = 2
    Debug.Assert dctTest.Keys()(dctTest.Count - 1) = 5
    
    mTrc.EoP ErrSrc(PROC)
    
End Sub

Private Sub Test_DctAdd_09_Performance_n(ByVal lAdds As Long)
    Const PROC = "Test_DctAdd_09_Performance_n"
    Dim i   As Long
    
    mTrc.BoP ErrSrc(PROC), "items added ordered = ", lAdds
    Set dctTest = Nothing
    For i = 1 To lAdds - 1 Step 2
        DctAdd add_dct:=dctTest, add_key:=i, add_item:=ThisWorkbook, add_seq:=seq_ascending ' by key case sensitive is the default
    Next i
    For i = lAdds To 2 Step -2
        DctAdd add_dct:=dctTest, add_key:=i, add_item:=ThisWorkbook, add_seq:=seq_ascending ' by key case sensitive is the default
    Next i
    mTrc.EoP ErrSrc(PROC)
    
End Sub

Private Sub Test_DctAdd_99_Performance()
    Const PROC = "Test_DctAdd_09_Performance"
    
    mTrc.BoP ErrSrc(PROC)
    
    Test_DctAdd_09_Performance_n 100
    Test_DctAdd_09_Performance_n 500
    Test_DctAdd_09_Performance_n 1000
    Test_DctAdd_09_Performance_n 1500
    Test_DctAdd_09_Performance_n 2000
    
    mTrc.EoP ErrSrc(PROC)
    
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
    
    mTrc.BoP ErrSrc(PROC)
    Set dct1 = Nothing
    Set dct2 = Nothing
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        DctAdd add_dct:=dct1, add_key:=vbc, add_item:=vbc, add_seq:=seq_ascending ' by key case sensitive is the default
    Next vbc
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        DctAdd add_dct:=dct2, add_key:=vbc, add_item:=vbc, add_seq:=seq_ascending ' by key case sensitive is the default
    Next vbc
    
    '~~ Test: Differs in keys
    Debug.Assert Not DctDiffers(dct1, dct2)
    dct1.Remove ThisWorkbook.VBProject.VBComponents("mTest")
    dct2.Remove ThisWorkbook.VBProject.VBComponents("mBasic")
    Debug.Assert DctDiffers(dct1, dct2)
    Set dct1 = Nothing
    Set dct2 = Nothing
    mTrc.EoP ErrSrc(PROC)
    
End Sub

