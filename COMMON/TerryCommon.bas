Attribute VB_Name = "TerryCommon"
'Common functions written by Terry for use in EsiFina, EsiTime, and POM

Option Explicit

Private bSqlRows As Boolean

'ComboBoxKeyUp parameters
Private keysSoFar As String 'keys typed in combobox
Private keysTyped As Integer

'Public Const UseNewPurchaseQty As Boolean = True 'implemented for AWJ by Terry

' Sheet Inventory parameters
Private UsingSheetInventory_Initialized As Boolean  'if true, used cached value below
Private UsingSheetInventory_Flag As Boolean
Private MostRecentSheetPartNo As String
Private MostRecentSheetConvQty As Integer       ' PartTable.PAPURCONV for part number above
Private MostRecentSheetStatus As Boolean        'default = false

Private Declare Function GetTickCount Lib "kernel32" () As Long

'Private StopwatchStartTime As Date
Private StopWatchStartTickCount As Long

Public Sub StopwatchStart()
'   StopwatchStartTime = time
   StopWatchStartTickCount = GetTickCount()
End Sub

Public Sub StopwatchStop(msg As String)
'   Dim StopTime As Date
   Dim StopwatchStopTickCount As Long
   StopwatchStopTickCount = GetTickCount()
   
'   StopTime = time - StopwatchStartTime
'   StopwatchStopTickCount = GetTickCount()
'   Dim minutes As Integer, seconds As Integer
'   minutes = CInt(Mid(StopTime, 4, 2))
'   seconds = CInt(Mid(StopTime, 7, 2))
   'MsgBox msg & " " & CStr(60 * minutes + seconds) & " seconds"
   
   MsgBox msg & " " & CStr((StopwatchStopTickCount - StopWatchStartTickCount) / 1000) & " seconds"
End Sub



Function CheckCurrency(Amount As String, Optional ZeroOK As Boolean = True) As String
   'Test for a valid currency amount and return currency to 2 decimal places
   'Syntax:  cur = CheckCurrency(txt)
   'return = "*" if invalid currency field
   
   On Error Resume Next
   CheckCurrency = "*"
   Dim cUR As Currency
   cUR = CCur(Amount)
   cUR = Round(cUR, 2)
   If Err Then
      MsgBox "Invalid currency amount: " & Amount
      Exit Function
   Else
      '        If cur = 0 Then
      '            MsgBox "Zero currency amount not allowed"   why not?
      '            Exit Function
      '        End If
   End If
   
   CheckCurrency = Format(cUR, "0.00")
End Function

Public Sub CheckCurrencyTextBox(txt As TextBox, Optional ZeroOK As Boolean = True)
   Dim S As String
   S = CheckCurrency(txt.Text, False)
   If S = "*" Then
      txt.SetFocus
   Else
      txt.Text = S
   End If
End Sub

Function CheckDecimal(Amount As String, Fmt As String, Optional BlankOK As Boolean = False) As String
   'Test for a valid decimal amount
   'Syntax:  cur = CheckDecimal(txt,"##0.00")
   'return = "*" if invalid amount
   If BlankOK And Trim(Amount) = "" Then
      CheckDecimal = ""
      Exit Function
   End If
   
   On Error Resume Next
   Dim C As Currency
   C = Format(Amount, Fmt)
   If Err Then
      MsgBox "Number must be in format: " & Replace(Fmt, "0", "#")
      CheckDecimal = "*"
   Else
      CheckDecimal = C
   End If
   
End Function


Public Function IsValidAccount(sAcct As String) As Boolean
   Dim rdo As ADODB.Recordset
   On Error GoTo whoops
   Dim S As String
   'Dim bSqlRows As Boolean
   IsValidAccount = False
   S = Compress(sAcct)
   sSql = "select * from GlacTable where GLACCTREF = '" & S & "'" & vbCrLf _
          & "and GLACCTREF not in (select GLMASTER from GlacTable)"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo)
   If bSqlRows Then
      IsValidAccount = True
   Else
      Dim sMsg As String
      sMsg = sAcct & " is not a valid account"
      MsgBox sMsg, vbInformation, "Invalid account"
   End If
   rdo.Close
   Set rdo = Nothing
   Exit Function
   
whoops:
   Exit Function
End Function

Public Sub CheckAccountInTextBox(txt As TextBox)
   'validate account in textbox.
   'if invalid, force user to remain in textbox unless striking cancel or control
   Dim S As String
   On Error Resume Next
   S = Trim(txt.Text)
   If Not IsValidAccount(S) Then
      txt.SetFocus
   End If
End Sub

Public Sub CheckAccountInComboBox(cbo As ComboBox)
   'validate account in textbox.
   'if invalid, force user to remain in textbox unless striking cancel or control
   Dim S As String
   On Error Resume Next
   S = Trim(cbo.Text)
   If Not IsValidAccount(S) Then
      cbo.SetFocus
   End If
End Sub

Public Function SARound(ByVal X As Variant, Optional ByVal DecPlaces As Integer = 0) As Variant
   ' Symetric Arithmetic Rounding
   ' VB Round Function uses bankers rounding, ie. round .785 down and .775 up
   'test cases in TestSARound
   
   If IsNull(X) Then
      X = 0
   End If
   SARound = Fix(X * 10 ^ DecPlaces + 0.5 * Sgn(X)) / 10 ^ DecPlaces
   If VarType(X) = vbCurrency Then
      SARound = CCur(SARound)
   End If
End Function

Public Function TestSARound() As Integer

   'returns 0 if all test cases work
   If SARound(34.499999999999, 0) <> 34 Then
      TestSARound = 1
   End If
   If SARound(35.499999999999, 0) <> 35 Then
      TestSARound = 2
   End If
   If SARound(34.5, 0) <> 35 Then
      TestSARound = 3
   End If
   If SARound(35.5, 0) <> 36 Then
      TestSARound = 4
   End If
   
   If SARound(34.449999999999, 1) <> 34.4 Then
      TestSARound = 5
   End If
   If SARound(35.449999999999, 1) <> 35.4 Then
      TestSARound = 6
   End If
   If SARound(34.45, 1) <> 34.5 Then
      TestSARound = 7
   End If
   If SARound(35.45, 1) <> 35.5 Then
      TestSARound = 8
   End If
End Function

Public Function NormalRound(ByVal X As Variant, Optional ByVal DecPlaces As Integer = 0) As Variant
   ' VB Round Function uses bankers rounding, ie. round .785 down and .775 up
   'this rounding function always rounds +...5 up and -...5 down
   
   If IsNull(X) Then
      X = 0
   End If
   NormalRound = CLng(X * 10 ^ DecPlaces + 0.5 * Sgn(X)) / 10 ^ DecPlaces     'clng still does bankers round
   If VarType(X) = vbCurrency Then
      NormalRound = CCur(NormalRound)
   End If
End Function

Public Sub CenterForm(frm As Form)
   frm.Top = (Screen.Height - frm.Height) / 2
   frm.Left = (Screen.Width - frm.Width) / 2
End Sub

Public Sub LoadComboWithPartsForOpenRuns(cbo As ComboBox, Optional bAllowSC As Boolean = False)
   'load the part numbers of all open runs into a combobox
   
   Dim rdo As ADODB.Recordset
   cbo.Clear
   
   If (bAllowSC = True) Then
      sSql = "Select distinct PARTNUM from PartTable " & vbCrLf _
          & "join RunsTable runs on RUNREF = PARTREF " & vbCrLf _
          & "and RUNSTATUS <> 'CL' and RUNSTATUS <> 'CA' and RUNSTATUS <> 'CO'" & vbCrLf _
          & "join RnopTable ops on ops.OPREF = runs.RUNREF" & vbCrLf _
          & "and ops.OPRUN = runs.RUNNO" & vbCrLf _
          & "and ops.OPCOMPDATE is null" & vbCrLf _
          & "order by PARTNUM"
   Else
   
    sSql = "Select distinct PARTNUM from PartTable " & vbCrLf _
           & "join RunsTable runs on RUNREF = PARTREF " & vbCrLf _
           & "and RUNSTATUS <> 'SC' and RUNSTATUS <> 'CL' and RUNSTATUS <> 'CA' and RUNSTATUS <> 'CO'" & vbCrLf _
           & "join RnopTable ops on ops.OPREF = runs.RUNREF" & vbCrLf _
           & "and ops.OPRUN = runs.RUNNO" & vbCrLf _
           & "and ops.OPCOMPDATE is null" & vbCrLf _
           & "order by PARTNUM"
   End If
   
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      With rdo
         While Not .EOF
            cbo.AddItem Trim(!PartNum)
            .MoveNext
         Wend
      End With
      
      Set rdo = Nothing
      
      If cbo.ListCount > 0 Then
         cbo.ListIndex = 0
      End If
   Else
      MsgBox "No open Manufacturing Orders", vbExclamation ', sSysCaption
   End If
   
End Sub

Public Sub LoadComboWithOpenRunsForPart(cbo As ComboBox, sPart As String, Optional bAllowSC As Boolean = False)
   'load the open run numbers for a part into a combobox
   
    Dim rdo As ADODB.Recordset
    cbo.Clear
    ' Adding empty RUn No
    cbo.AddItem ""
    
    If (bAllowSC = True) Then
        sSql = "Select distinct RUNNO from RunsTable runs" & vbCrLf _
             & "join RnopTable ops on ops.OPREF = runs.RUNREF" & vbCrLf _
             & "and ops.OPRUN = runs.RUNNO" & vbCrLf _
             & "where RUNREF = '" & Compress(sPart) & "'" & vbCrLf _
             & "and RUNSTATUS <> 'CL' and RUNSTATUS <> 'CA' and RUNSTATUS <> 'CO'" & vbCrLf _
             & "and ops.OPCOMPDATE is null" & vbCrLf _
             & "order by RUNNO"
    Else
   
        sSql = "Select distinct RUNNO from RunsTable runs" & vbCrLf _
             & "join RnopTable ops on ops.OPREF = runs.RUNREF" & vbCrLf _
             & "and ops.OPRUN = runs.RUNNO" & vbCrLf _
             & "where RUNREF = '" & Compress(sPart) & "'" & vbCrLf _
             & "and RUNSTATUS <> 'SC' and RUNSTATUS <> 'CL' and RUNSTATUS <> 'CA' and RUNSTATUS <> 'CO'" & vbCrLf _
             & "and ops.OPCOMPDATE is null" & vbCrLf _
             & "order by RUNNO"
            'and RUNSTATUS <> 'CO'
    End If
    
    If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
       With rdo
          While Not .EOF
             cbo.AddItem CStr(!Runno)
             .MoveNext
          Wend
       End With
       If cbo.ListCount > 0 Then
          cbo.ListIndex = 0
          
       End If
       Set rdo = Nothing
       '    Else
       '        MsgBox "No open Manufacturing Orders for this part", vbExclamation ', sSysCaption
    End If
   
End Sub

Public Sub LoadComboWithOpenOpsForRun(cbo As ComboBox, sPart As String, nRun As Long)
   'load the open ops for a run into a combobox
   
   Dim rdo As ADODB.Recordset
   Dim bTCSerOp As Boolean
   Dim bOnlyOpenOp As Boolean
   cbo.Clear
   
   bOnlyOpenOp = GetShowOpenOpOnly()
   bTCSerOp = GetTCServiceOp()
    If (bTCSerOp = True) Then
    
        If (bOnlyOpenOp = True) Then
        
         sSql = "Select OPNO from RnopTable " & vbCrLf _
                & " WHERE OPREF = '" & Compress(sPart) & "'" & vbCrLf _
                & " AND OPRUN = " & nRun & vbCrLf _
                & " AND OPCOMPLETE = 0 " & vbCrLf _
                & " ORDER BY OPNO"
        Else
        
         sSql = "Select OPNO from RnopTable " & vbCrLf _
                & " WHERE OPREF = '" & Compress(sPart) & "'" & vbCrLf _
                & " AND OPRUN = " & nRun & vbCrLf _
                & " ORDER BY OPNO"
        End If
    Else
         ' MM 6/1/2010
        If (bOnlyOpenOp = True) Then
            sSql = "SELECT OPNO FROM RnopTable " & vbCrLf _
                     & " WHERE OPREF = '" & Compress(sPart) & "'" & vbCrLf _
                     & " AND OPRUN = " & nRun & vbCrLf _
                     & " AND (LTRIM(RTRIM(OPSERVPART)) = '' OR OPSERVPART IS NULL) " & vbCrLf _
                     & " AND OPCOMPLETE = 0 " & vbCrLf _
                     & " ORDER BY OPNO"
        Else
        
            sSql = "SELECT OPNO FROM RnopTable " & vbCrLf _
                     & " WHERE OPREF = '" & Compress(sPart) & "'" & vbCrLf _
                     & " AND OPRUN = " & nRun & vbCrLf _
                     & " AND (LTRIM(RTRIM(OPSERVPART)) = '' OR OPSERVPART IS NULL) " & vbCrLf _
                     & " ORDER BY OPNO"
        End If
'        sSql = "SELECT RnopTable.OPNO OPNO FROM RnopTable, RtopTable " & vbCrLf _
'                    & " WHERE RnopTable.OPREF = RtopTable.OPREF " & vbCrLf _
'                        & " AND RnopTable.OPNO = RtopTable.OPNO " & vbCrLf _
'                        & " AND RnopTable.OPREF = '" & Compress(sPart) & "'" & vbCrLf _
'                        & " AND OPRUN = " & nRun & vbCrLf _
'                        & " AND OPSERVICE <> 1" & vbCrLf _
'                        & "ORDER BY RnopTable.OPNO"
                        '& " AND OPCOMPDATE IS NULL " & vbCrLf
    End If
    
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      With rdo
         While Not .EOF
            cbo.AddItem CStr(!opNo)
            .MoveNext
         Wend
      End With
      If cbo.ListCount > 0 Then
         cbo.ListIndex = 0
      End If
      Set rdo = Nothing
      
   Else
      MsgBox "No open operations for this MO", vbExclamation ', sSysCaption
   End If
   
End Sub

'Public Function Debugging() As Boolean
'    If InStr(1, Command, "/debug", vbTextCompare) Then
'        Debugging = True
'    Else
'        Debugging = False
'    End If
'End Function
'

Public Sub SortFlex(FlexGrid As MSFlexGrid, TheCol As Integer, SortType As Integer, Optional SortColOverride As Integer = 0)
   
   Dim Headline As String, Ascend As Boolean, Decend As Boolean, I As Integer, j As Integer
   If (SortColOverride > 0) Then FlexGrid.Col = SortColOverride Else FlexGrid.Col = TheCol
   For I = 0 To FlexGrid.Cols - 1
      Headline = FlexGrid.TextMatrix(0, I)
      Ascend = Right$(Headline, 1) = "+"
      Decend = Right$(Headline, 1) = "-"
      If Ascend Or Decend Then Headline = Left$(Headline, Len(Headline) - 1)
      If I = TheCol Then
         If Ascend Then
            FlexGrid.TextMatrix(0, I) = Headline & "-"
            Select Case SortType
               Case 0 'None
                  FlexGrid.Sort = flexSortNone
               Case 1 'Generic
                  FlexGrid.Sort = flexSortGenericDescending
               Case 2 'Numeric
                  FlexGrid.Sort = flexSortNumericDescending
               Case 3 'String
                  FlexGrid.Sort = flexSortStringDescending
            End Select
         Else
            FlexGrid.TextMatrix(0, I) = Headline & "+"
            Select Case SortType
               Case 0 'None
                  FlexGrid.Sort = flexSortNone
               Case 1 'Generic
                  FlexGrid.Sort = flexSortGenericAscending
               Case 2 'Numeric
                  FlexGrid.Sort = flexSortNumericAscending
               Case 3 'String
                  FlexGrid.Sort = flexSortStringAscending
            End Select
         End If
      Else
         FlexGrid.TextMatrix(0, I) = Headline
      End If
   Next I
   DoEvents
End Sub

'Public Sub ComboKeyUp(cbo As ComboBox, KeyCode As Integer, keysTyped As Integer)
'    'select first item with matching leading characters
'    'keysTyped = # of keys typed prior to the current one
'
'    Dim keysSoFar As String
'    keysSoFar = Left(cbo.Text, keysTyped) & Chr(KeyCode)
'    keysTyped = keysTyped + 1
'    Dim i As Integer
'    For i = 0 To cbo.ListCount - 1
'        If StrComp(Left(cbo.List(i), keysTyped), keysSoFar, vbTextCompare) = 0 Then
'            If cbo.ListIndex <> i Then
'                cbo.ListIndex = i
'                cbo.Refresh
'                DoEvents
'            End If
'            Debug.Print cbo.Name & " keysSoFar=" & keysSoFar & " match(" & i & ")=" & cbo.Text
'            Exit Sub
'        End If
'    Next
'End Sub

Public Sub ComboGotFocus(cbo As ComboBox)
   keysTyped = 0
   keysSoFar = ""
   'Debug.Print "ComboGotFocus " & cbo.Name
End Sub

Public Sub ComboKeyUp(cbo As ComboBox, KeyCode As Integer)
   'select first item with matching leading characters
   'keysTyped = # of keys typed prior to the current one
   'YOU MUST CALL ComboGotFocus from the combobox's GotFocus event in order for this to work
   
   If KeyCode < 32 Then
      Exit Sub
   End If
   
   keysSoFar = keysSoFar & Chr(KeyCode)
   keysTyped = keysTyped + 1
   Dim I As Integer
   For I = 0 To cbo.ListCount - 1
      If StrComp(Left(cbo.List(I), keysTyped), keysSoFar, vbTextCompare) >= 0 Then
         If cbo.ListIndex <> I Then
            cbo.ListIndex = I
            cbo.Refresh
            DoEvents
         End If
         Debug.Print cbo.Name & " keysSoFar=" & keysSoFar & " match(" & I & ")=" & cbo.Text & " KeyCode=" & KeyCode & " chr=" & Chr(KeyCode)
         Exit Sub
      End If
   Next
End Sub

Public Sub Calendar_KeyDown(KeyCode As Integer)
   'Debug.Print KeyCode
   On Error GoTo whoops
   Dim Fmt As String
   Debug.Print "Before: " & Screen.ActiveControl.Text
   If Len(Screen.ActiveControl.Text) <= 8 Then
      Fmt = "mm/dd/yy"
   Else
      Fmt = "mm/dd/yyyy"
   End If
   
   If KeyCode = vbKeyUp Then 'up
      Screen.ActiveControl.Text = Format(DateAdd("d", 1, Screen.ActiveControl.Text), Fmt)
   ElseIf KeyCode = vbKeyDown Then 'down
      Screen.ActiveControl.Text = Format(DateAdd("d", -1, Screen.ActiveControl.Text), Fmt)
   End If
   Debug.Print " After: " & Screen.ActiveControl.Text
   Exit Sub
whoops:
End Sub

Public Function Calendar_LostFocus_Showing_Calendar() As Boolean
   'returns True if leaving cbo box to show calendar for date selection
   'in this case, don't do any validation -- just return
   If Screen.ActiveForm.Name = "Calendar" Then
      Calendar_LostFocus_Showing_Calendar = True
   Else
      Calendar_LostFocus_Showing_Calendar = False
   End If
End Function

'Sub UpdateTimeCardTotals(EmployeeNumber As Long, CardDate As Date)
'   'update daily totals for timecard
'   'TchdTable.TMREGHRS,TMOVTHRS,TMDBLHRS,TMSTART,TMSTOP
'
'   sSql = "UpdateTimeCardTotals " & EmployeeNumber & ", '" & Format(CardDate, "mm/dd/yyyy") & "'"
'   RdoCon.Execute sSql, rdExecDirect
'End Sub
'
Public Sub SetComboBox(cbo As ComboBox, sValue As String)
   Dim I As Integer
   For I = 0 To cbo.ListCount - 1
      If cbo.List(I) = sValue Then
         cbo.ListIndex = I
         Exit Sub
      End If
   Next
   
   'if match not found, set to first entry
   If cbo.ListCount > 0 And cbo.ListIndex = -1 Then
      cbo.ListIndex = 0
   End If
End Sub

'Public Sub SetComboBoxByDataValue(cbo As ComboBox, sDataValue As String)
''this sets a Combobox where the list contains compressed values, but the key is compressed
''for instance, combobox might contain PARTNUM values, but search might be the compressed PARTREF value.
'   Dim I As Integer
'   For I = 0 To cbo.ListCount - 1
'      If Compress(cbo.ItemData(I)) = sDateValue Then
'         cbo.ListIndex = I
'         Exit Sub
'      End If
'   Next
'
'   'if match not found, set to no value (-1)
'   cbo.ListIndex = -1
'End Sub

Sub SetCompressedComboBox(cbo As ComboBox, Key As String)
   'use this to select key values in a combobox given that matches a compressed key
   'Example, select account 1210-000 when the key available is 1210000.
   'Note: this is really slow so just leave the dash numbers in wherever possible
   
   On Error Resume Next
   cbo = Key 'try the normal way
   
   'if failed, loop through all values
   If Err Then
      Dim keyComp As String
      Dim cboComp As String
      
      keyComp = Compress(Key)
      Dim I As Long
      For I = 0 To cbo.ListCount - 1
         cboComp = Compress(cbo.List(I))
         If StrComp(cboComp, keyComp, vbTextCompare) = 0 Then
            cbo.ListIndex = I
            Exit Sub
         End If
      Next
   End If
End Sub

Public Function IsInvoiceNumberAvailable(InvoiceNumber As Long) As Boolean
   'returns true if invoice number is available for use
   
   Dim rdo As ADODB.Recordset
   On Error GoTo whoops
   'Dim bSqlRows As Boolean
   
   IsInvoiceNumberAvailable = False
   sSql = "select InvNo from CihdTable where InvNo = " & InvoiceNumber
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo)
   rdo.Close
   If bSqlRows Then
      Dim sMsg As String
      sMsg = "Invoice number " & InvoiceNumber & " is already assigned."
      MsgBox sMsg, vbInformation, "Invoice number in use"
   Else
      IsInvoiceNumberAvailable = True
   End If
   Set rdo = Nothing
   Exit Function
   
whoops:
   Exit Function
End Function

Public Sub LoadComboWithVendors(cbo As ComboBox, Optional IncludeAll As Boolean = False)
   'load a combobox with vendor nicknames
   
   Dim rdo As ADODB.Recordset
   cbo.Clear
   If IncludeAll Then
      cbo.AddItem "<ALL>"
   End If
   
   sSql = "Select VENICKNAME from VndrTable " & vbCrLf _
          & "order by VENICKNAME"
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      With rdo
         While Not .EOF
            cbo.AddItem Trim(CStr(rdo(0)))
            .MoveNext
         Wend
      End With
   End If
   Set rdo = Nothing
   
   If cbo.ListCount > 0 Then
      cbo.ListIndex = 0
   End If
   
End Sub

Public Sub LoadComboWithSQL(cbo As ComboBox, sChkSQL As String, Optional IncludeAll As Boolean = False)
   'load a combobox with SQL string passed to it
   
    Dim rdo As ADODB.Recordset

    cbo.Clear
        
    If IncludeAll Then
       cbo.AddItem "<ALL>"
    End If
   
    If sChkSQL <> "" Then
        ' sSQL is defined globally; not good
        sSql = sChkSQL
    End If
   
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      With rdo
         While Not .EOF
            cbo.AddItem Trim(CStr(rdo(0)))
            .MoveNext
         Wend
      End With
      Set rdo = Nothing
   End If
   
   If cbo.ListCount > 0 Then
      cbo.ListIndex = 0
   End If
   
End Sub


Public Function GetVendorName(NickName As String) As String
   'returns the full vendor name if found
   'returns blank if vendor not found
   'returns "All Vendors" if nickname = "<ALL>"
   
   Dim nick As String
   nick = Trim(NickName)
   
   If nick = "<ALL>" Then
      GetVendorName = "All Vendors"
      Exit Function
   End If
   
   GetVendorName = ""
   Dim rdo As ADODB.Recordset
   On Error GoTo whoops
   'Dim bSqlRows As Boolean
   sSql = "select VEBNAME from VndrTable where VEREF = '" & Compress(nick) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo)
   If bSqlRows Then
      GetVendorName = Trim(rdo(0))
      rdo.Close
   Else
      rdo.Close
      Dim sMsg As String
      sMsg = NickName & " is not a valid vendor nickname"
      MsgBox sMsg, vbInformation, "No such Vendor"
   End If
   Set rdo = Nothing
   
   Exit Function
   
whoops:
   Exit Function
End Function

Public Sub LoadComboWithCustomers(cbo As ComboBox, Optional IncludeAll As Boolean = False)
   'load a combobox with Customer nicknames
   
   Dim rdo As ADODB.Recordset
   cbo.Clear
   If IncludeAll Then
      cbo.AddItem "<ALL>"
   End If
   
   sSql = "Select CUNICKNAME from CustTable " & vbCrLf _
          & "order by CUNICKNAME"
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      With rdo
         While Not .EOF
            cbo.AddItem Trim(CStr(rdo(0)))
            .MoveNext
         Wend
      End With
   End If
   Set rdo = Nothing
   
   If cbo.ListCount > 0 Then
      cbo.ListIndex = 0
   End If
   
End Sub

Public Function GetShowOpenOpOnly() As Boolean
    ' get COPOTIMESERVOP flag
    Dim RdoGet As ADODB.Recordset
    Dim bOpFlg As Boolean
   On Error GoTo whoops
    
    sSql = "SELECT ISNULL(COONLYOPENPO, 0) COONLYOPENPO FROM ComnTable WHERE COREF=1"
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet, ES_KEYSET)
    If bSqlRows Then
       With RdoGet
          bOpFlg = "" & Trim(!COONLYOPENPO)
          GetShowOpenOpOnly = bOpFlg
       End With
       RdoGet.Close
       Set RdoGet = Nothing
    Else
        ' record not found
        GetShowOpenOpOnly = True
    End If
   Exit Function
   
whoops:
   Exit Function
End Function

Public Function GetTCServiceOp() As Boolean
    ' get COPOTIMESERVOP flag
    Dim RdoGet As ADODB.Recordset
    Dim bSerFlg As Boolean
   On Error GoTo whoops
    
    sSql = "SELECT COPOTIMESERVOP FROM ComnTable WHERE COREF=1"
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet, ES_KEYSET)
    If bSqlRows Then
       With RdoGet
          bSerFlg = "" & Trim(!COPOTIMESERVOP)
          GetTCServiceOp = bSerFlg
       End With
       RdoGet.Close
       Set RdoGet = Nothing
    Else
        ' record not found
        GetTCServiceOp = True
    End If
   Exit Function
   
whoops:
   Exit Function
End Function
Public Function GetCustomerName(NickName As String) As String
   'returns the full Customer name if found
   'returns blank if Customer not found
   'returns "All Customers" if nickname = "<ALL>"
   
   Dim nick As String
   nick = Trim(NickName)
   
   If nick = "<ALL>" Then
      GetCustomerName = "All Customers"
      Exit Function
   End If
   
   GetCustomerName = ""
   Dim rdo As ADODB.Recordset
   On Error GoTo whoops
   'Dim bSqlRows As Boolean
   sSql = "select CUNAME from CustTable where CUREF = '" & Compress(nick) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo)
   If bSqlRows Then
      GetCustomerName = Trim(rdo(0))
      rdo.Close
      Set rdo = Nothing
   Else
      rdo.Close
      Dim sMsg As String
      sMsg = NickName & " is not a valid Customer nickname"
      MsgBox sMsg, vbInformation, "No such Customer"
   End If
   Exit Function
   
whoops:
   Exit Function
End Function

' the following 3 functions are used for sheet inventory
' implemented by Terry Oct 2016

Function UsingSheetInventory() As Boolean
    If Not UsingSheetInventory_Initialized Then
        Dim sql As String
        Dim rdo As ADODB.Recordset
        Dim result As Integer
        sql = "select COUSESHEETINVENTORY from ComnTable"
        If clsADOCon.GetDataSet(sql, rdo) Then
           result = rdo("COUSESHEETINVENTORY")
        End If
        Set rdo = Nothing
        UsingSheetInventory_Flag = IIf(result = 1, True, False)
        UsingSheetInventory_Initialized = True
    End If
    
    UsingSheetInventory = UsingSheetInventory_Flag
End Function

'Function GetPurchToInvConversionQty(sPartRef As String) As Currency
'    'return quantity to multiple purchase unit quantity to get inventory unit quantity
'    '   = 1 if no conversion required
'
'    If sPartRef = MostRecentSheetPartNo Then
'        GetPurchToInvConversionQty = MostRecentSheetConvQty
'    ElseIf UsingSheetInventory Then
'        Dim rdo As ADODB.Recordset
'        sSql = "select PAPURCONV from PartTable where PARTREF = '" + sPartRef + "'"
'        If clsADOCon.GetDataSet(sSql, rdo) Then
'            GetPurchToInvConversionQty = rdo("PAPURCONV")
'            Set rdo = Nothing
'            If GetPurchToInvConversionQty = 0 Then GetPurchToInvConversionQty = 1
'        Else
'            GetPurchToInvConversionQty = 1
'        End If
'    Else
'        GetPurchToInvConversionQty = 1
'    End If
'
'End Function

Function IsThisASheetPart(SPartRef As String) As Boolean
    If UsingSheetInventory Then
        If SPartRef <> MostRecentSheetPartNo Then
            Dim rdo As ADODB.Recordset
            sSql = "select PAUNITS, PAPUNITS from PartTable where PARTREF = '" + SPartRef + "' and PAPUNITS = 'SH' and PAUNITS <> PAPUNITS"
            MostRecentSheetStatus = clsADOCon.GetDataSet(sSql, rdo)
            MostRecentSheetPartNo = SPartRef
            Set rdo = Nothing
        End If
    End If
    IsThisASheetPart = MostRecentSheetStatus
End Function

Public Function TrimAddress(adr As String)
    'remove trailing blanks and trailing CRLFs
    Dim rtn As String
    
    rtn = Replace(adr, "'", "") 'remove single quotes if any
    Do While Len(rtn) > 0
        rtn = Trim(rtn)
        If Len(rtn) >= 2 Then
            If Right(Trim(rtn), 2) = vbCrLf Then
                rtn = Trim(Left(rtn, Len(rtn) - 2))
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    TrimAddress = rtn
End Function


Public Function SqlString(str As String) As String
   'prepares a string for a SQL statement
   'replaces a ' with '' (that's two single quotes, not one double quote)
   'removes trailing crlfs
   
   Dim rtn As String
   rtn = str
   rtn = Replace(rtn, "''", "'")    'avoid double-doubling apostrophes
   rtn = Replace(rtn, "'", "''")
   rtn = Replace(rtn, Chr(96), "''")
   rtn = Replace(rtn, Chr(146), "''")
   rtn = Replace(rtn, Chr(180), "''")
   Do While Len(rtn) > 0
      rtn = Trim(rtn)
      If Len(rtn) >= 2 Then
          If Right(Trim(rtn), 2) = vbCrLf Then
              rtn = Trim(Left(rtn, Len(rtn) - 2))
          Else
              Exit Do
          End If
      Else
          Exit Do
      End If
    Loop
    SqlString = rtn
End Function

Public Sub CreateDirectoryPath(fullPath As String)
   If Dir(fullPath, vbDirectory) <> "" Then
      Exit Sub
   End If
   
   'split off component paths
   Dim folder() As String, n As Integer, count As Integer, I As Integer
   count = 0
   Do While True
      n = InStr(fullPath, "\")
      If n = 0 Then Exit Do
      count = count + 1
      ReDim Preserve folder(count)
      folder(count) = Mid(fullPath, 1, n - 1)
      fullPath = Mid(fullPath, n + 1)
   Loop
'   For n = 1 To count
'      MsgBox folder(n)
'   Next n
'   MsgBox "end = " & fullPath
   
   If count > 1 Then
      Dim newPath As String
      newPath = folder(1)
      For n = 2 To count
         newPath = newPath & "\" & folder(n)
         If Dir(newPath, vbDirectory) = "" Then
            MkDir newPath
         End If
      Next n
   End If
End Sub

Public Function GetComnTableBit(ColumnName As String) As Boolean
   'Example:if  GetComnTableBit("CoUseAbbreviatedLotNumbers") ...
   Dim RdoCmn As ADODB.Recordset
   GetComnTableBit = False
   On Error Resume Next
   sSql = "select isnull(" & ColumnName & ", 0) Value from ComnTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmn, ES_FORWARD)
   If bSqlRows Then
      With RdoCmn
         If Val("" & !Value) = 1 Then GetComnTableBit = True
         ClearResultSet RdoCmn
      End With
   End If
   Set RdoCmn = Nothing
End Function

