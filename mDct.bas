Attribute VB_Name = "mDct"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mDct: Procedures for Dictionaries
'
' Note: 1. Procedures of the mDct module do not use the Common VBA Error Handler.
'          However, test module uses the mErrHndlr module for test purpose.
'
'       2. This module is developed, tested, and maintained in the dedicated
'          Common Component Workbook Dct.xlsm available on Github
'          https://Github.com/warbe-maker/VBA-Basic-Procedures
'
' Methods:
' Requires Reference to:
' - "Microsoft Scripting Runtime"
' - "Microsoft Visual Basic Application Extensibility .." (for test only!)
'
' W. Rauschenberger, Berlin Sept 2020
' ----------------------------------------------------------------------------
Private bSeqAscending       As Boolean
Private bSeqDescending      As Boolean
Private bOrderByItem        As Boolean
Private bOrderByKey         As Boolean
Private bCaseSensitive      As Boolean
Private bCaseIgnored        As Boolean
Private bSeqEntry           As Boolean
Private bSeqAfterTarget     As Boolean
Private bSeqBeforeTarget    As Boolean

Public Enum enDctAddOptions ' Dictionary add/insert modes
    sense_caseignored
    sense_casesensitive
    order_byitem
    order_bykey
    seq_aftertarget
    seq_ascending
    seq_beforetarget
    seq_descending
    seq_entry
End Enum

Private Function AppErr(ByVal lNo As Long) As Long
    AppErr = IIf(lNo < 0, lNo - vbObjectError, vbObjectError + lNo)
End Function

Public Sub DctAdd(ByRef dct As Dictionary, _
                  ByVal addkey As Variant, _
                  ByVal additem As Variant, _
         Optional ByVal order As enDctAddOptions = order_bykey, _
         Optional ByVal seq As enDctAddOptions = seq_entry, _
         Optional ByVal sense As enDctAddOptions = sense_casesensitive, _
         Optional ByVal target As Variant, _
         Optional ByVal keepduplicates As Boolean = True)
' ----------------------------------------------------------------------------
' Adds the item (additem) to the Dictionary (dct) with the key (addkey).
' Supports various key sequences, case and case insensitive key as well
' as adding items before or after an existing item.
' - When the key (addkey) already exists the item is updated when it is
'   numeric or a string, else it is ignored.
' - When the dictionary (dct) is Nothing it is setup on the fly.
' - When dctmode = before or after target is obligatory and an
'   error is raised when not provided.
' - When the item's key is an object any dctmode other then by seq
'   requires an object with a name property. If not the case an error is
'   raised.

' W. Rauschenberger, Berlin Oct 2020
' ----------------------------------------------------------------------------
    Const PROC = "DctAdd"
    Dim dctTemp             As Dictionary
    Dim vEntryKey           As Variant
    Dim vEntryItem          As Variant
    Dim vValueEntryNew      As Variant ' the argument key's/item's value
    Dim vValueEntryExisting As Variant ' the entry's key/item value for the comparison with the vValueEntryNew
    Dim vValueEntryTarget   As Variant ' the add before/after key/item's value
    Dim bDone               As Boolean
    Dim bAddedBefore        As Boolean
    Dim bAddedAfter         As Boolean
    Dim vItem               As Variant
    
    On Error GoTo on_error
    
    If dct Is Nothing Then Set dct = New Dictionary
    
    '~~ Plausibility checks
    Select Case order
        Case order_byitem:  bOrderByItem = True:    bOrderByKey = False
        Case order_bykey:   bOrderByItem = False:   bOrderByKey = True
        Case Else: Err.Raise AppErr(1), ErrSrc(PROC), "Vaue for argument order neither item nor key!"
    End Select
    
    Select Case seq
        Case seq_ascending:     bSeqAscending = True:   bSeqDescending = False: bSeqEntry = False:  bSeqAfterTarget = False:    bSeqBeforeTarget = False
        Case seq_descending:    bSeqAscending = False:  bSeqDescending = True:  bSeqEntry = False:  bSeqAfterTarget = False:    bSeqBeforeTarget = False
        Case seq_entry:         bSeqAscending = False:  bSeqDescending = False: bSeqEntry = True:   bSeqAfterTarget = False:    bSeqBeforeTarget = False
        Case seq_aftertarget:   bSeqAscending = False:  bSeqDescending = False: bSeqEntry = False:  bSeqAfterTarget = True:     bSeqBeforeTarget = False
        Case seq_beforetarget:  bSeqAscending = False:  bSeqDescending = False: bSeqEntry = False:  bSeqAfterTarget = False:    bSeqBeforeTarget = True
        Case Else: Err.Raise AppErr(2), ErrSrc(PROC), "Vaue for argument seq neither ascending, descendingy, after, before!"
    End Select
    
    Select Case sense
        Case sense_caseignored:     bCaseIgnored = True:    bCaseSensitive = False
        Case sense_casesensitive:   bCaseIgnored = False:    bCaseSensitive = True
        Case Else: Err.Raise AppErr(3), ErrSrc(PROC), "Vaue for argument sense neither case sensitive nor case ignored!"
    End Select
    
    If bOrderByKey And (bSeqBeforeTarget Or bSeqAfterTarget) And dct.Exists(addkey) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The to be added key (order value = '" & DctAddOrderValue(addkey, additem) & "') for an add before/after already exists!"

    '~~ When no target is specified for add before/after seq descending/ascending is used instead
    If IsMissing(target) Then
        If bSeqBeforeTarget Then seq = seq_descending
        If bSeqBeforeTarget Then seq = seq_ascending
    Else
        '~~ When a target is specified it must exist
        If (bSeqAfterTarget Or bSeqBeforeTarget) And bOrderByKey Then
            If Not dct.Exists(target) _
            Then Err.Raise mBasic.AppErr(2), ErrSrc(PROC), "The target key for an add before/after key does not exists!"
        ElseIf (bSeqAfterTarget Or bSeqBeforeTarget) And bOrderByItem Then
            If Not DctAddItemExists(dct, target) _
            Then Err.Raise AppErr(2), ErrSrc(PROC), "The target itemfor an add before/after item does not exists!"
        End If
        vValueEntryTarget = DctAddOrderValue(target, target)
    End If
        
    With dct
        '~~ When it is the very first item or the order option is entry seq the item will just be added
        If .Count = 0 Or bSeqEntry Then ' the very first item is just added
            .Add addkey, additem
            GoTo end_proc
        End If
        
        If bOrderByKey And .Exists(addkey) Then
            
            '~~ When the order is by key and the key already exists the item is updated
            If VarType(additem) = vbObject _
            Then Set .Item(addkey) = additem _
            Else .Item(addkey) = additem
            GoTo end_proc
        End If
    End With
        
    '~~ When the order argument is an object but does not have a name property raise an error
    If bOrderByKey Then
        If VarType(addkey) = vbObject Then
            On Error Resume Next
            If Err.Number <> 0 _
            Then Err.Raise AppErr(3), ErrSrc(PROC), "The order option is by key, the key is an object but does not have a name property!"
        End If
    Else ' order mode is by item
        On Error Resume Next
        If Err.Number <> 0 _
        Then Err.Raise AppErr(4), ErrSrc(PROC), "The order option is by item, the item is an object but does not have a name property!"
    End If
    
    vValueEntryNew = DctAddOrderValue(addkey, additem)
    
    With dct
        '~~ Get the last entry's order value
        vValueEntryExisting = DctAddOrderValue(.Keys()(.Count - 1), .Items()(.Count - 1))
        
        '~~ When the order mode is ascending and the last entry's key or item
        '~~ is less than the order argument just add it and exit
        If bSeqAscending And vValueEntryNew > vValueEntryExisting Then
            .Add addkey, additem
            GoTo end_proc
        End If
    End With
        
    '~~ Since the new key/item couldn't simply be added to the Dictionary it will
    '~~ be inserted before or after the key/item as specified.
    Set dctTemp = New Dictionary
    bDone = False
    For Each vEntryKey In dct
        
        If IsObject(dct.Item(vEntryKey)) _
        Then Set vEntryItem = dct.Item(vEntryKey) _
        Else vEntryItem = dct.Item(vEntryKey)
        
        With dctTemp
            If bDone Then
                '~~ All remaining items just transfer
                .Add vEntryKey, vEntryItem
            Else
                vValueEntryExisting = DctAddOrderValue(vEntryKey, vEntryItem)
            
                If vValueEntryExisting = vValueEntryTarget And bSeqBeforeTarget Then
                    '~~ The add before target key/item has been reached
                    .Add addkey, additem
                    .Add vEntryKey, vEntryItem
                    bDone = True ' Add done
                    bAddedBefore = True
                ElseIf vValueEntryExisting = vValueEntryTarget And bSeqAfterTarget Then
                    '~~ The add after target key/item has been reached
                    .Add vEntryKey, vEntryItem
                    .Add addkey, additem
                    bDone = True
                    bAddedAfter = True
                ElseIf bSeqAscending And vValueEntryExisting > vValueEntryNew Then
                    .Add addkey, additem
                    .Add vEntryKey, vEntryItem
                    bDone = True ' Add done
                ElseIf bSeqDescending And vValueEntryNew > vValueEntryExisting Then
                    '~~> Add before target key has been reached
                    .Add addkey, additem
                    .Add vEntryKey, vEntryItem
                    bDone = True
                Else
                    .Add vEntryKey, vEntryItem ' transfer item
                End If
            End If
        End With ' dctTemp
    Next vEntryKey
    
    '~~ Return the temporary dictionary with the new item added and all exiting items in dct transfered to it
    '~~ provided ther was not a add before/after error
    If bSeqBeforeTarget And Not bAddedBefore _
    Then Err.Raise AppErr(5), ErrSrc(PROC), "The key/item couldn't be added before because the target key/item did not exist!"
    If bSeqAfterTarget And Not bAddedAfter _
    Then Err.Raise AppErr(6), ErrSrc(PROC), "The key/item couldn't be added before because the target key/item did not exist!"
    
    Set dct = dctTemp
    Set dctTemp = Nothing

end_proc:
    Exit Sub

on_error:
#If Debugging Then
    Debug.Print Err.Description: Stop: Resume
#End If
    ErrMsg errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Sub

Private Sub ErrMsg(ByVal errnumber As Long, _
                  ByVal errsource As String, _
                  ByVal errdscrptn As String, _
                  ByVal errline As String)
' ----------------------------------------------------
' Display the error message by means of the VBA MsgBox
' ----------------------------------------------------
    
    Dim sErrMsg     As String
    Dim sErrPath    As String
    
    sErrMsg = "Description: " & vbLf & ErrMsgErrDscrptn(errdscrptn) & vbLf & vbLf & _
              "Source:" & vbLf & errsource & ErrMsgErrLine(errline)
    sErrPath = vbNullString ' only available with the mErrHndlr module
    If sErrPath <> vbNullString _
    Then sErrMsg = sErrMsg & vbLf & vbLf & _
                   "Path:" & vbLf & sErrPath
    If ErrMsgInfo(errdscrptn) <> vbNullString _
    Then sErrMsg = sErrMsg & vbLf & vbLf & _
                   "Info:" & vbLf & ErrMsgInfo(errdscrptn)
    MsgBox sErrMsg, vbCritical, ErrMsgErrType(errnumber, errsource) & " in " & errsource & ErrMsgErrLine(errline)

End Sub

Private Function ErrMsgErrType( _
        ByVal errnumber As Long, _
        ByVal errsource As String) As String
' ------------------------------------------
' Return the kind of error considering the
' Err.Source and the error number.
' ------------------------------------------

   If InStr(1, Err.Source, "DAO") <> 0 _
   Or InStr(1, Err.Source, "ODBC Teradata Driver") <> 0 _
   Or InStr(1, Err.Source, "ODBC") <> 0 _
   Or InStr(1, Err.Source, "Oracle") <> 0 Then
      ErrMsgErrType = "Database Error"
   Else
      If errnumber > 0 _
      Then ErrMsgErrType = "VB Runtime Error" _
      Else ErrMsgErrType = "Application Error"
   End If
   
End Function

Private Function ErrMsgErrLine(ByVal errline As Long) As String
    If errline <> 0 _
    Then ErrMsgErrLine = " (at line " & errline & ")" _
    Else ErrMsgErrLine = vbNullString
End Function

Private Function ErrMsgErrDscrptn(ByVal s As String) As String
' -------------------------------------------------------------------
' Return the string before a "||" in the error description. May only
' be the case when the error has been raised by means of err.Raise
' which means when it is an "Application Error".
' -------------------------------------------------------------------
    If InStr(s, DCONCAT) <> 0 _
    Then ErrMsgErrDscrptn = Split(s, DCONCAT)(0) _
    Else ErrMsgErrDscrptn = s
End Function

Private Function ErrMsgInfo(ByVal s As String) As String
' -------------------------------------------------------------------
' Return the string after a "||" in the error description. May only
' be the case when the error has been raised by means of err.Raise
' which means when it is an "Application Error".
' -------------------------------------------------------------------
    If InStr(s, DCONCAT) <> 0 _
    Then ErrMsgInfo = Split(s, DCONCAT)(1) _
    Else ErrMsgInfo = vbNullString
End Function

Private Function DctAddOrderValue(ByVal dctkey As Variant, _
                                  ByVal dctitem As Variant) As Variant
' --------------------------------------------------------------------
' When keyoritem is an object its name becomes the order value else
' the keyoiritem value as is.
' --------------------------------------------------------------------
    If bOrderByKey Then
    
        If VarType(dctkey) = vbObject _
        Then DctAddOrderValue = dctkey.Name _
        Else DctAddOrderValue = dctkey
        
    ElseIf bOrderByItem Then
    
        If VarType(dctitem) = vbObject _
        Then DctAddOrderValue = dctitem.Name _
        Else DctAddOrderValue = dctitem
    
    End If
    
    If TypeName(DctAddOrderValue) = "String" And bCaseIgnored Then DctAddOrderValue = LCase$(DctAddOrderValue)

End Function

Public Function DctDiffers( _
                ByVal dct1 As Dictionary, _
                ByVal dct2 As Dictionary, _
       Optional ByVal diffitems As Boolean = True, _
       Optional ByVal diffkeys As Boolean = True) As Boolean
' ----------------------------------------------------------
' Returns TRUE when Dictionary 1 (dct1) is different from
' Dictionary 2 (dct2). diffitems and diffkeys may be False
' when only either of the two differences matters.
' When both are false only a differns in the number of
' entries constitutes a difference.
' ----------------------------------------------------------
Const PROC  As String = "DictDiffers"
Dim i       As Long
Dim v       As Variant

    On Error GoTo on_error
    
    '~~ Difference in entries
    DctDiffers = dct1.Count <> dct2.Count
    If DctDiffers Then GoTo exit_proc
    
    If diffitems Then
        '~~ Difference in items
        For i = 0 To dct1.Count - 1
            If VarType(dct1.Items()(i)) = vbObject And VarType(dct1.Items()(i)) = vbObject Then
                DctDiffers = Not dct1.Items()(i) Is dct2.Items()(i)
                If DctDiffers Then GoTo exit_proc
            ElseIf (VarType(dct1.Items()(i)) = vbObject And VarType(dct1.Items()(i)) <> vbObject) _
                Or (VarType(dct1.Items()(i)) <> vbObject And VarType(dct1.Items()(i)) = vbObject) Then
                DctDiffers = True
                GoTo exit_proc
            ElseIf dct1.Items()(i) <> dct2.Items()(i) Then
                DctDiffers = True
                GoTo exit_proc
            End If
        Next i
    End If
    
    If diffkeys Then
        '~~ Difference in keys
        For i = 0 To dct1.Count - 1
            If VarType(dct1.Keys()(i)) = vbObject And VarType(dct1.Keys()(i)) = vbObject Then
                DctDiffers = Not dct1.Keys()(i) Is dct2.Keys()(i)
                If DctDiffers Then GoTo exit_proc
            ElseIf (VarType(dct1.Keys()(i)) = vbObject And VarType(dct1.Keys()(i)) <> vbObject) _
                Or (VarType(dct1.Keys()(i)) <> vbObject And VarType(dct1.Keys()(i)) = vbObject) Then
                DctDiffers = True
                GoTo exit_proc
            ElseIf dct1.Keys()(i) <> dct2.Keys()(i) Then
                DctDiffers = True
                GoTo exit_proc
            End If
        Next i
    End If
       
exit_proc:
    Exit Function
    
on_error:
#If Debugging Then
    Debug.Print Err.Description: Stop: Resume
#End If
    ErrMsg errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Function

Private Function DctAddItemExists( _
                 ByVal dct As Dictionary, _
                 ByVal dctitem As Variant) As Boolean
    
    Dim v As Variant
    DctAddItemExists = False
    
    For Each v In dct
        If VarType(dct.Item(v)) = vbObject Then
            If dct.Item(v) Is dctitem Then
                DctAddItemExists = True
                Exit Function
            End If
        Else
            If dct.Item(v) = dctitem Then
                DctAddItemExists = True
                Exit Function
            End If
        End If
    Next v
    
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = ThisWorkbook.Name & " mDct." & sProc
End Function



