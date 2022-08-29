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

Private Sub BoC(ByVal boc_id As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' (B)egin-(o)f-(C)ode with id (boc_id) trace. Procedure to be copied as Private
' into any module potentially using the Common VBA Execution Trace Service. Has
' no effect when Conditional Compile Argument is 0 or not set at all.
' Note: The begin id (boc_id) has to be identical with the paired EoC statement.
' ------------------------------------------------------------------------------
    Dim s As String: If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")
#If ExecTrace = 1 Then
    mTrc.BoC boc_id, s
#End If
End Sub

Private Sub BoP(ByVal b_proc As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' (B)egin-(o)f-(P)rocedure named (b_proc). Procedure to be copied as Private
' into any module potentially either using the Common VBA Error Service and/or
' the Common VBA Execution Trace Service. Has no effect when Conditional Compile
' Arguments are 0 or not set at all.
' ------------------------------------------------------------------------------
    Dim s As String: If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")
#If ErHComp = 1 Then
    mErH.BoP b_proc, s
#ElseIf ExecTrace = 1 Then
    mTrc.BoP b_proc, s
#End If
End Sub

Private Sub EoC(ByVal eoc_id As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' (E)nd-(o)f-(C)ode id (eoc_id) trace. Procedure to be copied as Private into
' any module potentially using the Common VBA Execution Trace Service. Has no
' effect when the Conditional Compile Argument is 0 or not set at all.
' Note: The end id (eoc_id) has to be identical with the paired BoC statement.
' ------------------------------------------------------------------------------
    Dim s As String: If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")
#If ExecTrace = 1 Then
    mTrc.BoC eoc_id, s
#End If
End Sub

Private Sub EoP(ByVal e_proc As String, _
       Optional ByVal e_inf As String = vbNullString)
' ------------------------------------------------------------------------------
' (E)nd-(o)f-(P)rocedure named (e_proc). Procedure to be copied as Private Sub
' into any module potentially either using the Common VBA Error Service and/or
' the Common VBA Execution Trace Service. Has no effect when Conditional Compile
' Arguments are 0 or not set at all.
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    mErH.EoP e_proc
#ElseIf ExecTrace = 1 Then
    mTrc.EoP e_proc, e_inf
#End If
End Sub

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service. Displays a debugging option button
' when the Conditional Compile Argument 'Debugging = 1' and an optional
' additional "About the error:" section when information is concatenated with
' the error message by two vertical bars (||).
'
' May be copied as Private Function into any module. Considers the Common VBA
' Message Service and the Common VBA Error Services as optional components.
' When neither is installed the error message is displayed by the VBA.MsgBox.
'
' Usage: Example with the Conditional Compile Argument 'Debugging = 1'
'
'        Private/Public <procedure-name>
'            Const PROC = "<procedure-name>"
'
'            On Error Goto eh
'            ....
'        xt: Exit Sub/Function/Property
'
'        eh: Select Case ErrMsg(ErrSrc(PROC))
'               Case vbResume:  Stop: Resume
'               Case Else:      GoTo xt
'            End Select
'        End Sub/Function/Property
'
' Note:  The above may seem to be a lot of code but will be a godsend in case
'        of an error!
'
' Uses:
' - AppErr For programmed application errors (Err.Raise AppErr(n), ....) to
'          turn tem into negative and in the error mesaage back into a positive
'          number.
' - ErrSrc To provide an unambigous procedure name - prefixed by the module name
'
' W. Rauschenberger Berlin, Nov 2021
'
' See:
' https://warbe-maker.github.io/vba/common/2022/02/15/Personal-and-public-Common-Components.html
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ When Common VBA Error Services (mErH) is availabel in the VB-Project
    '~~ (which includes the mMsg component) the mErh.ErrMsg service is invoked.
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
#ElseIf MsgComp = 1 Then
    '~~ When (only) the Common Message Service (mMsg, fMsg) is available in the
    '~~ VB-Project, mMsg.ErrMsg is invoked for the display of the error message.
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
#End If
    '~~ When neither of the Common Component is available in the VB-Project
    '~~ the error message is displayed by means of the VBA.MsgBox
    Dim ErrBttns    As Variant
    Dim ErrAtLine   As String
    Dim ErrDesc     As String
    Dim ErrLine     As Long
    Dim ErrNo       As Long
    Dim ErrSrc      As String
    Dim ErrText     As String
    Dim ErrTitle    As String
    Dim ErrType     As String
    Dim ErrAbout    As String
        
    '~~ Obtain error information from the Err object for any argument not provided
    If err_no = 0 Then err_no = Err.Number
    If err_line = 0 Then ErrLine = Erl
    If err_source = vbNullString Then err_source = Err.Source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
    
    '~~ Consider extra information is provided with the error description
    If InStr(err_dscrptn, "||") <> 0 Then
        ErrDesc = Split(err_dscrptn, "||")(0)
        ErrAbout = Split(err_dscrptn, "||")(1)
    Else
        ErrDesc = err_dscrptn
    End If
    
    '~~ Determine the type of error
    Select Case err_no
        Case Is < 0
            ErrNo = AppErr(err_no)
            ErrType = "Application Error "
        Case Else
            ErrNo = err_no
            If err_dscrptn Like "*DAO*" _
            Or err_dscrptn Like "*ODBC*" _
            Or err_dscrptn Like "*Oracle*" _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & ErrDesc & vbLf & vbLf & "Source: " & vbLf & err_source & ErrAtLine
    If ErrAbout <> vbNullString Then ErrText = ErrText & vbLf & vbLf & "About: " & vbLf & ErrAbout
    
#If Debugging Then
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & "Debugging:" & vbLf & "Yes    = Resume Error Line" & vbLf & "No     = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    ErrMsg = MsgBox(Title:=ErrTitle, Prompt:=ErrText, Buttons:=ErrBttns)

xt:
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mDctTest." & sProc
End Function

Public Sub Test_00_Regression()
' ------------------------------------------------------------------------------
' Attention! This Regression tes takes about 30 seconds due to the included
'            performance test which writes the results to the "Test" sheet.
' ------------------------------------------------------------------------------
    Const PROC = "Test_Regression"
    
    On Error GoTo eh
    
    '~~ Initialization of a new Trace Log File for this Regression test
    '~~ ! must be done prior the first BoP !
    mTrc.LogFile = Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "Regression Test.log")
    mTrc.LogTitle = "Regression Test module mDct"
    
    BoP ErrSrc(PROC)
    
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
    EoP ErrSrc(PROC)
    mTrc.Dsply
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
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
    Dim j       As Long: j = 999
    Dim jStep   As Long: jStep = 2
    Dim k       As Long: k = 1000
    Dim kStep   As Long: kStep = -2
    
    BoP ErrSrc(PROC) ' , "added items = ", k
    Set dctTest = Nothing
    For i = 1 To j Step jStep
        DctAdd add_dct:=dctTest, add_key:=i, add_item:=ThisWorkbook, add_seq:=seq_ascending ' by key case sensitive is the default
    Next i
    For i = k To jStep Step kStep
        DctAdd add_dct:=dctTest, add_key:=i, add_item:=ThisWorkbook, add_seq:=seq_ascending ' by key case sensitive is the default
    Next i
    
    '~~ Add an already existing key, ignored when the item is neither numeric nor a string
    DctAdd add_dct:=dctTest, add_key:=5, add_item:=ThisWorkbook, add_seq:=seq_ascending ' by key case sensitive is the default
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
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
    
    BoP ErrSrc(PROC)
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
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
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
    
    BoP ErrSrc(PROC)
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
        
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
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
    
    BoP ErrSrc(PROC)
    
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
        
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
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
    
    BoP ErrSrc(PROC)
    
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
            
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
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
    
    BoP ErrSrc(PROC)
    
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
        
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
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
    
    BoP ErrSrc(PROC)
    
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
        
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
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
    
    BoP ErrSrc(PROC)
    Set dctTest = Nothing
    
    DctAdd dctTest, 2, 5, add_seq:=seq_ascending
    DctAdd dctTest, 5, 2, add_seq:=seq_ascending
    DctAdd dctTest, 3, 4, add_seq:=seq_ascending
    
    Debug.Assert dctTest.Count = 3
    Debug.Assert dctTest.Keys()(0) = 2
    Debug.Assert dctTest.Keys()(dctTest.Count - 1) = 5
        
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
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
    
    BoP ErrSrc(PROC), "items added ordered = ", lAdds
    Set dctTest = Nothing
    For i = 1 To lAdds - 1 Step 2
        DctAdd add_dct:=dctTest, add_key:=i, add_item:=ThisWorkbook, add_seq:=seq_ascending ' by key case sensitive is the default
    Next i
    For i = lAdds To 2 Step -2
        DctAdd add_dct:=dctTest, add_key:=i, add_item:=ThisWorkbook, add_seq:=seq_ascending ' by key case sensitive is the default
    Next i
        
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
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
    BoP ErrSrc(PROC)
    
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
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
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
    
    BoP ErrSrc(PROC)
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
        
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
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
    
    BoP ErrSrc(PROC)
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
        
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
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
    
    BoP ErrSrc(PROC)
    
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
        
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
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
    
    BoP ErrSrc(PROC)
    
    Test_09_DctAdd_Performance_n 100
    Test_09_DctAdd_Performance_n 500
    Test_09_DctAdd_Performance_n 1000
    Test_09_DctAdd_Performance_n 1500
    Test_09_DctAdd_Performance_n 2000
        
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
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

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

