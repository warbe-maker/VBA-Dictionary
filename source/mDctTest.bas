Attribute VB_Name = "mDctTest"
Option Explicit
Option Private Module

Public Type tMsgSection                 ' ---------------------
       sLabel As String                 ' Structure of the
       sText As String                  ' UserForm's message
       bMonspaced As Boolean            ' area which consists
End Type                                ' of 4 message sections
Public Type tMsg                        ' Attention: 4 is a
       Section(1 To 4) As tMsgSection   ' design constant!
End Type                                ' ---------------------

Private dctTest As Dictionary

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed 'Application' error number not conflicts with the
' number of a 'VB Runtime Error' or any other system error.
' - Returns a given positive 'Application Error' number (app_err_no) into a
'   negative by adding the system constant vbObjectError
' - Returns the original 'Application Error' number when called with a negative
'   error number.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mDctTest." & sProc
End Function

Public Sub Test_00_Regression()
' ------------------------------------------------------------------------------
' Requires the Cond. Comp. Args.:
' Debugging = 1 : ErHComp = 1 : MsgComp = 1 : XcTrc_mTrc = 1
'
' Attention! This Regression test may takes up to 30 seconds due to the included
'            performance test which writes the results to the "Test" sheet.
' ------------------------------------------------------------------------------
    Const PROC = "Test_Regression"
    
    On Error GoTo eh
    
    '~~ Initialization of a new Trace Log File for this Regression test
    '~~ ! must be done prior the first BoP !
    mTrc.FileName = "RegessionTest.ExecTrace.log"
    mTrc.Title = "Regression Test module mDct"
    mTrc.NewFile
    
    mBasic.BoP ErrSrc(PROC)
    
    mErH.Regression = True ' prevent display of asserted errors
    Set dctTest = Nothing
    
    Test_01_DctAdd_Performance_KeyIsValue
    Test_02_DctAdd_KeyIsObjectWithNameProperty
    Test_03_DctAdd_ItemIsObjectWithNameProperty
    Test_04_DctAdd_InsertKeyBefore
    Test_05_DctAdd_InsertKeyAfter
    Test_06_DctAdd_InsertItemBefore
    Test_07_DctAdd_InsertItemAfter
    Test_08_DctAdd_NumKey
    Test_10_DctAdd_AddDuplicate_Item
    Test_99_DctAdd_Performance
    
xt: mErH.Regression = False
    mBasic.EoP ErrSrc(PROC)
    mTrc.Dsply
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_01_DctAdd_Performance_KeyIsValue()
' ----------------------------------------------------------------------------
' Note: Since a 100% reverse key order added in mode ascending is the worst
' case regarding performance this test sorts 100 items with 50% already in seq
' and the other 50% to be inserted.
' ----------------------------------------------------------------------------
    Const PROC = "Test_01_DctAdd_Performance_KeyIsValue"
    
    On Error GoTo eh
    Dim i       As Long
    Dim j       As Long: j = 99
    Dim jStep   As Long: jStep = 2
    Dim k       As Long: k = 100
    Dim kStep   As Long: kStep = -2
    
    mBasic.BoP ErrSrc(PROC) ' , "added items = ", k
    Set dctTest = New Dictionary
    For i = 1 To j Step jStep
        If Not dctTest.Exists(i) _
        Then DctAdd add_dct:=dctTest, add_key:=i, add_item:=ThisWorkbook, add_seq:=seq_ascending ' by key case sensitive is the default
    Next i
    For i = k To jStep Step kStep
        If Not dctTest.Exists(i) _
        Then DctAdd add_dct:=dctTest, add_key:=i, add_item:=ThisWorkbook, add_seq:=seq_ascending ' by key case sensitive is the default
    Next i
    
    '~~ Add an already existing key, ignored when the item is neither numeric nor a string
    DctAdd add_dct:=dctTest, add_key:=5, add_item:=ThisWorkbook, add_seq:=seq_ascending ' by key case sensitive is the default
    
xt: mBasic.EoP ErrSrc(PROC)
    Set dctTest = Nothing
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_02_DctAdd_KeyIsObjectWithNameProperty()
' ----------------------------------------------------------------------------
' Added items with a key which is an object. The order by key uses the
' object's name property.
' This test procedure is also used to provide a test Dictionary with all
' components of this VB-Project.
' ----------------------------------------------------------------------------
    Const PROC = "Test_02_DctAdd_KeyIsObjectWithNameProperty"
    
    On Error GoTo eh
    Dim i   As Long
    Dim vbc As VBComponent
    
    mBasic.BoP ErrSrc(PROC)
    Set dctTest = Nothing
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        DctAdd add_dct:=dctTest, add_key:=vbc, add_item:=vbc.Name, add_seq:=seq_ascending ' by key case sensitive is the default
    Next vbc
    Debug.Assert dctTest.Count = ThisWorkbook.VBProject.VBComponents.Count
    Debug.Assert dctTest.Items()(0) = "fMsg"
    Debug.Assert dctTest.Items()(dctTest.Count - 1) = "wsDct"
    
    '~~ Add an already existing key = update the item
    Set vbc = ThisWorkbook.VBProject.VBComponents("mDctTest")
    DctAdd add_dct:=dctTest, add_key:=vbc, add_item:=vbc.Name, add_seq:=seq_ascending ' by key case sensitive is the default
    Debug.Assert dctTest.Count = ThisWorkbook.VBProject.VBComponents.Count
    Debug.Assert dctTest.Items()(0) = "fMsg"
    Debug.Assert dctTest.Items()(dctTest.Count - 1) = "wsDct"
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_03_DctAdd_ItemIsObjectWithNameProperty()
' ----------------------------------------------------------------------------
' Added items with a key which is an object. The order by key uses the
' object's name property.
' ----------------------------------------------------------------------------
    Const PROC = "Test_03_DctAdd_ItemIsObjectWithNameProperty"
    
    On Error GoTo eh
    Dim i   As Long
    Dim vbc As VBComponent
    
    mBasic.BoP ErrSrc(PROC)
    Set dctTest = Nothing
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        DctAdd add_dct:=dctTest, add_key:=vbc.Name, add_item:=vbc, add_order:=order_byitem, add_seq:=seq_ascending
    Next vbc
    Debug.Assert dctTest.Count = ThisWorkbook.VBProject.VBComponents.Count
'    Test_DisplayResult dctTest
    Debug.Assert dctTest.Items()(0).Name = "fMsg"
    Debug.Assert dctTest.Items()(dctTest.Count - 1).Name = "wsDct"
    
    '~~ Add an already existing key = update the item
    Set vbc = ThisWorkbook.VBProject.VBComponents("mDctTest")
    DctAdd add_dct:=dctTest, add_key:=vbc.Name, add_item:=vbc, add_order:=order_byitem, add_seq:=seq_ascending
    Debug.Assert dctTest.Count = ThisWorkbook.VBProject.VBComponents.Count
    Debug.Assert dctTest.Items()(0).Name = "fMsg"
    Debug.Assert dctTest.Items()(dctTest.Count - 1).Name = "wsDct"
        
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_04_DctAdd_InsertKeyBefore()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_04_DctAdd_InsertKeyBefore"
    
    On Error GoTo eh
    Dim vbc_second As VBComponent
    Dim vbc_first As VBComponent
    
    mBasic.BoP ErrSrc(PROC)
    
    '~~ Preparation
    Test_02_DctAdd_KeyIsObjectWithNameProperty
    Debug.Assert dctTest.Keys()(0).Name = "fMsg"
    Debug.Assert dctTest.Keys()(1).Name = "mBasic"
    Set vbc_second = ThisWorkbook.VBProject.VBComponents("mTrc")
    Set vbc_first = ThisWorkbook.VBProject.VBComponents("mDctTest")
    dctTest.Remove vbc_second
    Debug.Assert dctTest.Count = ThisWorkbook.VBProject.VBComponents.Count - 1
    
    '~~ Test
    DctAdd dctTest, vbc_second, vbc_second.Name, add_order:=order_bykey, add_seq:=seq_beforetarget, add_target:=vbc_first
    Debug.Assert dctTest.Count = ThisWorkbook.VBProject.VBComponents.Count
    Debug.Assert dctTest.Keys()(0).Name = "fMsg"
    Debug.Assert dctTest.Keys()(1).Name = "mBasic"
        
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_05_DctAdd_InsertKeyAfter()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_05_DctAdd_InsertKeyAfter"
    
    On Error GoTo eh
    Dim vbc_second As VBComponent
    Dim vbc_first As VBComponent
    
    mBasic.BoP ErrSrc(PROC)
    
    '~~ Preparation
    Test_02_DctAdd_KeyIsObjectWithNameProperty
    Debug.Assert dctTest.Keys()(0).Name = "fMsg"
    Debug.Assert dctTest.Keys()(1).Name = "mBasic"
    Set vbc_first = ThisWorkbook.VBProject.VBComponents(1)
    Set vbc_second = ThisWorkbook.VBProject.VBComponents(2)
    
    '~~ Test
    dctTest.Remove vbc_first
    Debug.Assert dctTest.Count = ThisWorkbook.VBProject.VBComponents.Count - 1
    DctAdd dctTest, add_key:=vbc_first, add_item:=vbc_first.Name, add_order:=order_bykey, add_seq:=seq_aftertarget, add_target:=vbc_second
    Debug.Assert dctTest.Count = ThisWorkbook.VBProject.VBComponents.Count
    Debug.Assert dctTest.Keys()(0).Name = "fMsg"
    Debug.Assert dctTest.Keys()(1).Name = "mBasic"
            
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_06_DctAdd_InsertItemBefore()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_06_DctAdd_InsertItemBefore"
    
    On Error GoTo eh
    Dim vbc_second As VBComponent
    Dim vbc_first As VBComponent
    
    mBasic.BoP ErrSrc(PROC)
    
    '~~ Preparation
    Test_03_DctAdd_ItemIsObjectWithNameProperty
    Debug.Assert dctTest.Keys()(0) = "fMsg"
    Debug.Assert dctTest.Keys()(1) = "mBasic"
    Set vbc_second = ThisWorkbook.VBProject.VBComponents("fMsg")
    Set vbc_first = ThisWorkbook.VBProject.VBComponents("mBasic")
    dctTest.Remove vbc_second.Name
    Debug.Assert dctTest.Count = ThisWorkbook.VBProject.VBComponents.Count - 1
    
    '~~ Test
    DctAdd dctTest, vbc_second.Name, vbc_second, add_order:=order_byitem, add_seq:=seq_beforetarget, add_target:=vbc_first
    Debug.Assert dctTest.Count = ThisWorkbook.VBProject.VBComponents.Count
    Debug.Assert dctTest.Items()(0).Name = "fMsg"
    Debug.Assert dctTest.Items()(1).Name = "mBasic"
        
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_07_DctAdd_InsertItemAfter()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_07_DctAdd_InsertItemAfter"
    
    On Error GoTo eh
    Dim vbc_second As VBComponent
    Dim vbc_first As VBComponent
    
    mBasic.BoP ErrSrc(PROC)
    
    '~~ Preparation
    Test_03_DctAdd_ItemIsObjectWithNameProperty
    Debug.Assert dctTest.Keys()(0) = "fMsg"
    Debug.Assert dctTest.Keys()(1) = "mBasic"
    Set vbc_second = ThisWorkbook.VBProject.VBComponents("mBasic")
    Set vbc_first = ThisWorkbook.VBProject.VBComponents("fMsg")
    dctTest.Remove vbc_first.Name
    Debug.Assert dctTest.Count = ThisWorkbook.VBProject.VBComponents.Count - 1
    
    '~~ Test
    DctAdd dctTest, vbc_first.Name, vbc_first, add_order:=order_byitem, add_seq:=seq_aftertarget, add_target:=vbc_second
    Debug.Assert dctTest.Count = ThisWorkbook.VBProject.VBComponents.Count
    Debug.Assert dctTest.Items()(0).Name = vbc_second.Name
    Debug.Assert dctTest.Items()(1).Name = vbc_first.Name
        
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_08_DctAdd_NumKey()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_08_DctAdd_NumKey"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    Set dctTest = Nothing
    
    DctAdd dctTest, 2, 5, add_seq:=seq_ascending
    DctAdd dctTest, 5, 2, add_seq:=seq_ascending
    DctAdd dctTest, 3, 4, add_seq:=seq_ascending
    
    Debug.Assert dctTest.Count = 3
    Debug.Assert dctTest.Keys()(0) = 2
    Debug.Assert dctTest.Keys()(dctTest.Count - 1) = 5
        
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_09_DctAdd_Performance_n(ByVal lAdds As Long)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_09_DctAdd_Performance_n"
    
    On Error GoTo eh
    Dim i As Long
    
    mBasic.BoP ErrSrc(PROC), "items added ordered = " & lAdds
    Set dctTest = Nothing
    For i = 1 To lAdds - 1 Step 2
        DctAdd add_dct:=dctTest, add_key:=i, add_item:=ThisWorkbook, add_seq:=seq_ascending ' by key case sensitive is the default
    Next i
    For i = lAdds To 2 Step -2
        DctAdd add_dct:=dctTest, add_key:=i, add_item:=ThisWorkbook, add_seq:=seq_ascending ' by key case sensitive is the default
    Next i
        
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_10_DctAdd_AddDuplicate_Item()
' ----------------------------------------------------------------------------
' When add criteria is by item, the item already exists but with a different
' key and staywithfirst = False (the default) the item is added.
' ----------------------------------------------------------------------------
    Const PROC = "Test_10_DctAdd_AddDuplicate_Item"

    On Error GoTo eh
    mBasic.BoP ErrSrc(PROC)
    
    Set dctTest = Nothing
    DctAdd add_dct:=dctTest, add_key:="A", add_item:=60, add_order:=order_byitem, add_seq:=seq_ascending
    DctAdd add_dct:=dctTest, add_key:="BB", add_item:=50, add_order:=order_byitem, add_seq:=seq_ascending
    DctAdd add_dct:=dctTest, add_key:="CCC", add_item:=30, add_order:=order_byitem, add_seq:=seq_ascending
    DctAdd add_dct:=dctTest, add_key:="DDDD", add_item:=30, add_order:=order_byitem, add_seq:=seq_ascending
    DctAdd add_dct:=dctTest, add_key:="EEEEE", add_item:=20, add_order:=order_byitem, add_seq:=seq_ascending
    DctAdd add_dct:=dctTest, add_key:="FFFFFF", add_item:=10, add_order:=order_byitem, add_seq:=seq_ascending
'    Test_DisplayResult dctTest, "staywithfirst=False"
    Debug.Assert dctTest.Count = 6
    
    Set dctTest = Nothing
    DctAdd add_dct:=dctTest, add_key:="A", add_item:=60, add_order:=order_byitem, add_seq:=seq_ascending, add_staywithfirst:=True
    DctAdd add_dct:=dctTest, add_key:="BB", add_item:=50, add_order:=order_byitem, add_seq:=seq_ascending, add_staywithfirst:=True
    DctAdd add_dct:=dctTest, add_key:="CCC", add_item:=30, add_order:=order_byitem, add_seq:=seq_ascending, add_staywithfirst:=True
    DctAdd add_dct:=dctTest, add_key:="DDDD", add_item:=30, add_order:=order_byitem, add_seq:=seq_ascending, add_staywithfirst:=True
    DctAdd add_dct:=dctTest, add_key:="EEEEE", add_item:=20, add_order:=order_byitem, add_seq:=seq_ascending, add_staywithfirst:=True
    DctAdd add_dct:=dctTest, add_key:="FFFFFF", add_item:=10, add_order:=order_byitem, add_seq:=seq_ascending, add_staywithfirst:=True
'    Test_DisplayResult dctTest, "staywithfirst=True"
    Debug.Assert dctTest.Count = 5
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_20_DctDiffers_InKeysAsObject()
' ----------------------------------------------------------------------------
' Precondition: DctAdd is tested
' ----------------------------------------------------------------------------
    Const PROC = "Test_20_DctDiffers_InKeysAsObject"
    
    On Error GoTo eh
    Dim dct1 As Dictionary
    Dim dct2 As Dictionary
    Dim vbc  As VBComponent
    
    mBasic.BoP ErrSrc(PROC)
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
    dct1.Remove ThisWorkbook.VBProject.VBComponents("mDctTest")
    dct2.Remove ThisWorkbook.VBProject.VBComponents("mBasic")
    Debug.Assert DctDiffers(dd_dct1:=dct1 _
                          , dd_dct2:=dct2 _
                          , dd_diff_items:=False _
                          , dd_diff_keys:=True)
    Set dct1 = Nothing
    Set dct2 = Nothing
        
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_21_DctDiffers_InItemsAsObject()
' ----------------------------------------------------------------------------
' Precondition: DctAdd is tested
' ----------------------------------------------------------------------------
    Const PROC = "Test_21_DctDiffers_InItemsAsObject"
    
    On Error GoTo eh
    Dim dct1 As Dictionary
    Dim dct2 As Dictionary
    Dim vbc  As VBComponent
    
    mBasic.BoP ErrSrc(PROC)
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
    dct1.Remove ThisWorkbook.VBProject.VBComponents("mDctTest")
    dct2.Remove ThisWorkbook.VBProject.VBComponents("mBasic")
    Debug.Assert DctDiffers(dd_dct1:=dct1 _
                          , dd_dct2:=dct2 _
                          , dd_diff_items:=True _
                          , dd_diff_keys:=False)
    Set dct1 = Nothing
    Set dct2 = Nothing
        
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_22_DctDiffers_InItemsAsString()
' ----------------------------------------------------------------------------
' Precondition: DctAdd is tested
' ----------------------------------------------------------------------------
    Const PROC = "Test_21_DctDiffers_InItemsAsObject"
    
    On Error GoTo eh
    Dim dct1    As Dictionary
    Dim dct2    As Dictionary
    Dim vbc     As VBComponent
    Dim v       As Variant
    Dim i       As Long
    
    mBasic.BoP ErrSrc(PROC)
    
    '~~ Prepare
    Set dct1 = Nothing
    Set dct2 = Nothing
    With ThisWorkbook.VBProject.VBComponents("mBasic").CodeModule
        For i = 1 To .CountOfLines
            DctAdd add_dct:=dct1, add_key:=i, add_item:=.Lines(i, 1)
        Next i
    End With
    With ThisWorkbook.VBProject.VBComponents("mBasic").CodeModule
        For i = 1 To .CountOfLines
            DctAdd add_dct:=dct2, add_key:=i, add_item:=.Lines(i, 1)
        Next i
    End With
    
    '~~ Test and assert
    Debug.Assert DctDiffers(dd_dct1:=dct1 _
                          , dd_dct2:=dct2 _
                          , dd_diff_items:=True _
                          , dd_diff_keys:=False) = False
    
    Debug.Assert DctDiffers(dd_dct1:=dct1 _
                          , dd_dct2:=dct2 _
                          , dd_diff_items:=True _
                          , dd_diff_keys:=False _
                          , dd_ignore_items_empty:=True) = False
    
    '~~ Prepare
    Set dct1 = Nothing
    Set dct2 = Nothing
    With ThisWorkbook.VBProject.VBComponents("mBasic").CodeModule
        For i = 1 To .CountOfLines
            DctAdd add_dct:=dct1, add_key:=i, add_item:=.Lines(i, 1)
        Next i
    End With
    With ThisWorkbook.VBProject.VBComponents("mBasic").CodeModule
        For i = 1 To .CountOfLines
            DctAdd add_dct:=dct2, add_key:=i, add_item:=.Lines(i, 1)
        Next i
    End With
    
    Debug.Print dct1.Count
    Debug.Print dct2.Count
    mDct.RemoveEmptyItems dct1
    Debug.Print dct1.Count
    Debug.Print dct2.Count
    
    Debug.Assert DctDiffers(dd_dct1:=dct1 _
                          , dd_dct2:=dct2 _
                          , dd_diff_items:=True _
                          , dd_diff_keys:=False _
                          , dd_ignore_items_empty:=False) = True
    
    Debug.Assert DctDiffers(dd_dct1:=dct1 _
                          , dd_dct2:=dct2 _
                          , dd_diff_items:=True _
                          , dd_diff_keys:=False _
                          , dd_ignore_items_empty:=True) = False
    
    Set dct1 = Nothing
    Set dct2 = Nothing
        
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_99_DctAdd_Performance()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_09_DctAdd_Performance"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    
    Test_09_DctAdd_Performance_n 100
    Test_09_DctAdd_Performance_n 500
    Test_09_DctAdd_Performance_n 1000
    Test_09_DctAdd_Performance_n 1250
    Test_09_DctAdd_Performance_n 1500
        
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_DisplayResult(ByVal dct As Dictionary, _
                     Optional ByVal s As String)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_DisplayResult"
    
    Dim v           As Variant
    Dim sKey        As String
    Dim sItem       As String
    Dim MaxKey      As Long
    Dim MaxlItem    As Long
    
    On Error GoTo eh
    For Each v In dct
        If VarType(v) = vbObject Then sKey = v.Name Else sKey = v
        MaxKey = Max(MaxKey, Len(sKey))
        If VarType(dct.Item(v)) = vbObject Then sItem = dct.Item(v).Name Else sItem = dct.Item(v)
        MaxlItem = Max(MaxlItem, Len(sItem))
    Next v
    
    Debug.Print ">> ----- " & s & " --------------"
    For Each v In dct
        If VarType(v) = vbObject Then sKey = v.Name Else sKey = v
        If VarType(dct.Item(v)) = vbObject Then sItem = dct.Item(v).Name Else sItem = dct.Item(v)
        Debug.Print "Key: '" & sKey & "'," & Space(MaxKey - Len(sKey)) & " Item: '" & sItem & "'"
    Next v
    Debug.Print "<< ----- " & s & " --------------"
          
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

