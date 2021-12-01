Attribute VB_Name = "EsiEngr"
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'Customer Permissions 6/28/03
'10/13/04 revamped GetBidLabor
'03/29/05 Removed Jet references
'8/8/05 Fixed KeySet clearing
'10/31/05 Added Cur.CurrentGroup to OpenFavorite. Opens appropriate tab
'         when called from Recent/Favorites and closed.
'1/12/06 Completed renaming dialogs to be consistent with Fina
'1/16/06 See notes in ppiTabEsti
'2/14/06 Added GetEstimatingPermissions (See function)
'3/13/06 Added Customer Label to all Bid forms
'4/10/06 Added CheckBidEntries to trap Exit with no Part or Customer
'4/19/06 Added Stored Bid Queries and lightened the Escape parameters (no Part/Cust)
'4/20/06 Added Stored Procedures in all Groups
'5/30/06 BuildKeys
'6/23/06 Removed Threed32.OCX
'8/9/06 Removed SSTab32.OCX
'1/17/07 Added GetThisCustomer to Esti reports with cmbCst
Option Explicit

'Public Const ES_MoneyFormat = "0.00"
Public Const TTSAVEPRN = "_Printer"

Public Y As Byte
Public bFoundPart As Byte
Public bGoodBid As Byte

Public iAutoIncr As Integer
Public iSelected As Integer

Public sCurrForm As String
Public sCurrRout As String
Public sCurrEstimator As String
Public sDefaultShop As String
Public sLastDocClass As String
Public sPassedPart As String
Public sPassedRout As String
Public sSelected As String

Public sFavorites(13) As String
Public sRecent(10) As String
Public sSession(50) As String

'Rdo Connection for Full Bid Forms
Public RdoFull As ADODB.Recordset
Public RdoBid As ADODB.Recordset

'old help stuff for this module only
Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Any) As Long
Public Const HELP_CONTEXT = &H1 'Display topic in ulTopic
Public Const HELP_QUIT = &H2 'Terminate help
Public Const HELP_INDEX = &H3 'Display index
Public Const HELP_CONTENTS = &H3
Public Const HELP_HELPONHELP = &H4 'Display help on using help
Public Const HELP_SETINDEX = &H5 'Set the current Index for multi index help
Public Const HELP_SETCONTENTS = &H5
Public Const HELP_CONTEXTPOPUP = &H8
Public Const HELP_FORCEFILE = &H9
Public Const HELP_KEY = &H101 'Display topic for keyword in offabData
Public Const HELP_COMMAND = &H102
Public Const HELP_PARTIALKEY = &H105 'call the search engine in winhelp
Public Const HELP_SHOWTAB = &HB '(11) Show the tab

'Bid Bom Variables for GetBidLabor
Private cBomHours As Currency
Private cBomRate As Currency
Private cBomFoh As Currency
Private cBomLabor As Currency
Private cTotalHrs As Currency
Private cTotalRte As Currency
Private cTotalFoh As Currency
Private cTotalLabr As Currency
Dim cUnitHours As Currency
Dim cUnitOverhead As Currency
Dim cUnitCost As Currency

'Column updates
Private RdoCol As ADODB.Recordset
'Private ER As rdoError
Private ER As ADODB.Error

'Material Totals
'Dim cBidQuantity As Currency
Dim cBidBurden As Currency
Dim cBidMaterial As Currency
'Dim cBidTotMat As Currency


'Type tUser
'    Adduser As Integer
'    Level   As Integer
'    Group1  As Integer
'    Group2  As Integer
'    Group3  As Integer
'    Group4  As Integer
'    Group5  As Integer
'    Group6  As Integer
'    Group7  As Integer
'    Group8  As Integer
'End Type
'Public User As tUser

'1/11/07 Added for reports

Public Sub GetThisCustomer()
   Dim RdoRpt As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT CUNAME FROM CustTable WHERE CUREF='" _
          & Compress(MDISect.ActiveForm.cmbCst) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpt, ES_FORWARD)
   If bSqlRows Then
      MDISect.ActiveForm.lblCUName = "" & Trim(RdoRpt!CUNAME)
      ClearResultSet RdoRpt
   Else
      MDISect.ActiveForm.lblCUName = "*** A Range Of Customers Selected ***"
   End If
   Set RdoRpt = Nothing
   Exit Sub
   
modErr1:
   On Error GoTo 0
   
End Sub

Public Function GetBidPart(frm As Form) As Byte
   Dim RdoPrt As ADODB.Recordset
   On Error GoTo DiaErr1
   If (Trim(frm.txtPrt)) = "" Then
      GetBidPart = 0
      frm.txtPrt.ToolTipText = "No Part Number Has Been Entered."
      Exit Function
   End If
   sSql = "Qry_GetPartNumberBasics '" & Compress(frm.txtPrt) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
   If bSqlRows Then
      With RdoPrt
         frm.txtPrt.Text = "" & Trim(!PartNum)
         frm.txtPrt.ToolTipText = "" & Trim(!PADESC) & "     "
         GetBidPart = 1
         frm.txtPrt.ForeColor = vbBlack
         ClearResultSet RdoPrt
      End With
   Else
      frm.txtPrt.ForeColor = ES_RED
      frm.txtPrt.ToolTipText = "No Matching Part Found."
      GetBidPart = 0
   End If
   Set RdoPrt = Nothing
   Exit Function
   
DiaErr1:
   GetBidPart = 0
   sProcName = "getbidpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors frm
   
   
End Function

Public Function GetBidRouting() As Byte
   Dim RdoRte As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT BIDRTEREF FROM EsrtTable Where BIDRTEREF=" _
          & Val(MDISect.ActiveForm.cmbBid)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_FORWARD)
   If bSqlRows Then
      GetBidRouting = 1
      ClearResultSet RdoRte
   Else
      GetBidRouting = 0
   End If
   Set RdoRte = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getbidrout"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
End Function

Public Function GetBidPartsList() As Byte
   Dim RdoPls As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT BIDBOMREF FROM EsbmTable Where BIDBOMREF=" _
          & Val(MDISect.ActiveForm.cmbBid)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPls, ES_FORWARD)
   If bSqlRows Then
      GetBidPartsList = 1
      ClearResultSet RdoPls
   Else
      GetBidPartsList = 0
   End If
   Set RdoPls = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getbidparts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
End Function

'5/23/06 (PROPLA) Add the default formula (potential joins) leave in. some fool may delete it

Public Sub GetNextBid(frm As Form)
   Dim l As Long
   Dim RdoNxt As ADODB.Recordset
   'On Error Resume Next
   On Error GoTo DiaErr1
   If RunningBeta Then
      sSql = "SELECT FORMULA_REF FROM EsfrTable WHERE FORMULA_REF='NONE'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoNxt, ES_FORWARD)
      If Not bSqlRows Then
         sSql = "INSERT INTO EsfrTable (FORMULA_REF) VALUES('NONE')"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
      End If
   End If
   'On Error GoTo DiaErr1
   sSql = "SELECT MAX(BIDREF) FROM EstiTable "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoNxt, ES_FORWARD)
   If bSqlRows Then
      With RdoNxt
         If IsNull(.Fields(0)) Then
            l = 1
         Else
            l = .Fields(0) + 1
            
         End If
         ClearResultSet RdoNxt
      End With
      frm.lblNxt = Format(l, "000000")
   Else
      frm.lblNxt = Format(1, "000000")
   End If
   Set RdoNxt = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getnextbid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors frm
   
End Sub

Sub SelectHelpTopic(frm As Form, HelpTopic As String)
   Dim l&
   'change to esisale.hlp
   l& = WinHelp(frm.hwnd, sReportPath & "EsiFina.hlp", HELP_KEY, HelpTopic)
   
End Sub

Public Function GetBidCustomer(frm As Form, sCustomer) As Byte
   On Error GoTo DiaErr1
   Dim RdoBct As ADODB.Recordset
   If (Trim(sCustomer)) = "" Then
      GetBidCustomer = 0
      frm.cmbCst.ToolTipText = "No Customer Entered."
      Exit Function
   End If
   sSql = "Qry_GetCustomerBasics '" & Compress(sCustomer) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBct, ES_FORWARD)
   If bSqlRows Then
      With RdoBct
         frm.cmbCst.Text = "" & Trim(!CUNICKNAME)
         frm.cmbCst.ToolTipText = "" & Trim(!CUNAME)
         frm.cmbCst.ForeColor = vbBlack
         GetBidCustomer = 1
         ClearResultSet RdoBct
      End With
   Else
      frm.cmbCst.ForeColor = ES_RED
      frm.cmbCst.ToolTipText = "No Matching Customer Found."
      GetBidCustomer = 0
   End If
   Set RdoBct = Nothing
   Exit Function
   
DiaErr1:
   GetBidCustomer = 0
   sProcName = "getbidcustomer"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect
   
End Function


'Use installed query to find all part types 1 thru 3

Public Sub FillPartsBelow4(Cntrl As Control)
   Dim RdoFp4 As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "Qry_SortedPartTypesBelow4"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoFp4, ES_FORWARD)
   If bSqlRows Then
      With RdoFp4
         Do Until .EOF
            AddComboStr Cntrl.hwnd, "" & Trim(!PartNum)
            .MoveNext
         Loop
         ClearResultSet RdoFp4
      End With
   End If
   If Cntrl.ListCount <> 0 Then Cntrl = Cntrl.List(0)
   Set RdoFp4 = Nothing
   Exit Sub
   
modErr1:
   sProcName = "fillpartsbelow4"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
End Sub


Public Sub OpenFavorite(sSelected As String)
   CloseForms
   If LTrim$(sSelected) = "" Then Exit Sub
   MouseCursor 13
   On Error GoTo OpenFavErr1
   Select Case sSelected
      Case "Routing Report"
         cUR.CurrentGroup = "Rout"
         RoutRTp01a.Show
      Case "Routings By Routing Number"
         cUR.CurrentGroup = "Rout"
         RoutRTp02a.Show
      Case "Routings By Part Number"
         cUR.CurrentGroup = "Rout"
         RoutRTp03a.Show
      Case "Routings Used On Report"
         cUR.CurrentGroup = "Rout"
         RoutRTp04a.Show
      Case "Routings By Service Part Number"
         cUR.CurrentGroup = "Rout"
         RoutRTp05a.Show
      Case "Routing Assignments"
         cUR.CurrentGroup = "Rout"
         RoutRTe02a.Show
      Case "Copy A Routing"
         cUR.CurrentGroup = "Rout"
         RoutRTf02a.Show
      Case "Default Routing Assignments"
         cUR.CurrentGroup = "Rout"
         RoutRTe03a.Show
      Case "Delete A Routing"
         cUR.CurrentGroup = "Rout"
         RoutRTf01a.Show
      Case "Copy A Routing"
         cUR.CurrentGroup = "Rout"
         RoutRTf02a.Show
      Case "Routing Operation Library"
         cUR.CurrentGroup = "Rout"
         RoutRTe05a.Show
      Case "Merge Routings"
         cUR.CurrentGroup = "Rout"
         RoutRTf04a.Show
      Case "Reorganize Operations"
         cUR.CurrentGroup = "Rout"
         RoutRTf05a.Show
      Case "Revise A Routing Number"
         cUR.CurrentGroup = "Rout"
         RoutRTe04a.Show
      Case "Routings"
         cUR.CurrentGroup = "Rout"
         RoutRTe01a.Show
      Case "Reassign Shops And Work Centers"
         cUR.CurrentGroup = "Rout"
         RoutRTf03a.Show
      Case "Parts List"
         cUR.CurrentGroup = "Bomp"
         BompBMe02a.Show
         '8/16/04
         ' Case "Component References"
         '     diaBcref.Show
      Case "Assign A Parts List To A Part"
         cUR.CurrentGroup = "Bomp"
         BompBMe03a.Show
      Case "Create Parts List Revisions For Type 4's"
         cUR.CurrentGroup = "Bomp"
         BompBMe04a.Show
      Case "Copy A Parts List"
         cUR.CurrentGroup = "Bomp"
         BompBMf01a.Show
      Case "Copy A Bill Of Material"
         cUR.CurrentGroup = "Bomp"
         BompBMf01a.Show
      Case "Delete A Parts List"
         cUR.CurrentGroup = "Bomp"
         BompBMf02a.Show
      Case "Change A Parts List Revision"
         cUR.CurrentGroup = "Bomp"
         BompBMf03a.Show
      Case "Release A Parts List To Production"
         cUR.CurrentGroup = "Bomp"
         BompBMf04a.Show
      Case "Parts List Report"
         cUR.CurrentGroup = "Bomp"
         BompBMp01a.Show
      Case "Bills of Material (Report)"
         cUR.CurrentGroup = "Bomp"
         BompBMp02a.Show
      Case "Bills Of Material", "Bill Of Material"
         cUR.CurrentGroup = "Bomp"
         BompBMe01a.Show
      Case "Parts List Used On"
         cUR.CurrentGroup = "Bomp"
         BompBMp03a.Show
      Case "Document Classes"
         cUR.CurrentGroup = "Docu"
         DocuDCe03a.Show
      Case "Documents"
         cUR.CurrentGroup = "Docu"
         DocuDCe01a.Show
      Case "Document List"
         cUR.CurrentGroup = "Docu"
         DocuDCe02a.Show
      Case "Document Reference List"
         cUR.CurrentGroup = "Docu"
         DocuDCp01a.Show
      Case "Part Document List"
         cUR.CurrentGroup = "Docu"
         DocuDCp02a.Show
      Case "Copy A Document List"
         cUR.CurrentGroup = "Docu"
         DocuDCf01a.Show
      Case "Document Classes (Report)"
         cUR.CurrentGroup = "Docu"
         DocuDCp04a.Show
      Case "Document Used On"
         cUR.CurrentGroup = "Docu"
         DocuDCp05a.Show
      Case "List Of Parts Lists"
         cUR.CurrentGroup = "Docu"
         BompBMp04a.Show
      Case "Parts"
         cUR.CurrentGroup = "Bomp"
         InvcINe01a.Show
      Case "Delete A Routing Library Operation"
         cUR.CurrentGroup = "Rout"
         RoutRTf06a.Show
      Case "Qwik Bid"
         cUR.CurrentGroup = "Esti"
         If RunningBeta Then
            ' MM TODO: Load ppiESe01a
         Else
            Load EstiESe01a
         End If
      Case "Cancel An Estimate"
         cUR.CurrentGroup = "Esti"
         EstiESf02a.Show
      Case "Mark An Estimate As Not Accepted"
         cUR.CurrentGroup = "Esti"
         EstiESf01a.Show
      Case "Accept Estimates"
         cUR.CurrentGroup = "Esti"
         EstiESe04a.Show
      Case "Complete Estimates"
         cUR.CurrentGroup = "Esti"
         EstiESe03a.Show
      Case "Estimate (Report)"
         cUR.CurrentGroup = "Esti"
         If RunningBeta Then
            ' MM TODO: ppiESp01a.Show
         Else
            EstiESp01a.Show
         End If
      Case "Estimate Summary By Customer"
         If RunningBeta Then
            ' MM TODO: ppiESp02a.Show
         Else
            EstiESp01a.Show
         End If
         cUR.CurrentGroup = "Esti"
      Case "Estimate Summary By Part Number"
         cUR.CurrentGroup = "Esti"
         If RunningBeta Then
            ' MM TODO: ppiESp03a.Show
         Else
            EstiESp03a.Show
         End If
      Case "Estimates Not Completed"
         cUR.CurrentGroup = "Esti"
         EstiESp04a.Show
      Case "Requests For Quotation"
         cUR.CurrentGroup = "Esti"
         EstiESe05a.Show
      Case "Cancel An RFQ"
         cUR.CurrentGroup = "Esti"
         EstiESf03a.Show
      Case "Canceled Estimates"
         cUR.CurrentGroup = "Esti"
         EstiESp05a.Show
      Case "Requests For Quotation (Report)"
         cUR.CurrentGroup = "Esti"
         EstiESp06a.Show
      Case "Estimating Parts"
         cUR.CurrentGroup = "Esti"
         EstiESp07a.Show
      Case "Full Estimate"
         cUR.CurrentGroup = "Esti"
         If RunningBeta Then
            ' MM TODO: ppiESe02a.Show
         Else
            EstiESe02a.Show
         End If
      Case "Assign Documents To Parts"
         cUR.CurrentGroup = "Docu"
         DocuDCe04a.Show
      Case "Assign Pictures To Parts"
         cUR.CurrentGroup = "Docu"
         DocuDCe05a.Show
      Case "Update MO Document Lists"
         cUR.CurrentGroup = "Docu"
         DocuDCe06a.Show
      Case "Manufacturing Order Document List"
         cUR.CurrentGroup = "Docu"
         DocuDCp06a.Show
      Case "Delete A Document"
         cUR.CurrentGroup = "Docu"
         DocuDCf02a.Show
      Case "Documents Used On MO's"
         cUR.CurrentGroup = "Docu"
         DocuDCp07a.Show
      Case "Tools"
         cUR.CurrentGroup = "Tool"
         ToolTLe01a.Show
      Case "Tool Lists"
         cUR.CurrentGroup = "Tool"
         ToolTLe02a.Show
      Case "Tools By Tool Number"
         cUR.CurrentGroup = "Tool"
         ToolTLp01a.Show
      Case "Tools By Description"
         cUR.CurrentGroup = "Tool"
         ToolTLp02a.Show
      Case "Tool Lists (Report)"
         cUR.CurrentGroup = "Tool"
         ToolTLp03a.Show
      Case "Delete A Tool"
         cUR.CurrentGroup = "Tool"
         ToolTLf03a.Show
      Case "Delete A Tool List"
         cUR.CurrentGroup = "Tool"
         ToolTLf03a.Show
      Case "Copy A Tool List"
         cUR.CurrentGroup = "Tool"
         ToolTLf01a.Show
      Case "Assign A Tool List To A Routing"
         cUR.CurrentGroup = "Tool"
         ToolTLe03a.Show
      Case "Assign A Tool List To A Manufacturing Order"
         cUR.CurrentGroup = "Tool"
         ToolTLe04a.Show
      Case "Tool Lists Used On Routings"
         cUR.CurrentGroup = "Tool"
         ToolTLp04a.Show
      Case "Tool Lists Used On Manufacturing Orders"
         cUR.CurrentGroup = "Tool"
         ToolTLp05a.Show
      Case "List of Revised Parts Lists"
         cUR.CurrentGroup = "Bomp"
         BompBMp05a.Show
      Case "List of Released Parts Lists"
         cUR.CurrentGroup = "Bomp"
         BompBMp06a.Show
      Case "Routings By Engineer"
         cUR.CurrentGroup = "Rout"
         RoutRTp06a.Show
         '         Case "Test Parse"
         '            TestForm.Show
      Case "Add/Edit Formulae"
         ' MM TODO: ppiESf04a.Show
      Case "List Of Formulae"
         ' MM TODO: ppiESp04a.Show
      Case "Change An Estimating Formula"
         ' MM TODO: ppiESf05a.Show
      Case "Delete An Estimating Formula"
         ' MM TODO: ppiESf06a.Show
      Case "Estimate Status"
         EstiESp08a.Show
      Case "Formula Calculator"
         ' MM TODO: ppiESe03a.Show
      Case "Routing Operation Pictures"
         RoutRTe06a.Show
      Case "Pictures By Routing Operation"
         RoutRTp07a.Show
      Case "Copy An Estimate"
         EstiESf05a.Show
      Case "Copy An Estimate Routing"
         EstiESf06a.Show
      Case "Copy A Routing From A Manufacturing Order"
         RoutRTf07a.Show
      Case Else
         MouseCursor 0
   End Select
   
   On Error GoTo 0
   Exit Sub
   
OpenFavErr1:
   Resume OpenFavErr2
OpenFavErr2:
   MsgBox "ActiveX Error. Can't Load Form..", 48, "System    "
   On Error GoTo 0
   
End Sub

Public Sub FillBomhRev(sFillrevision As String)
   Dim BmhRes As ADODB.Recordset
   'MDISect.ActiveForm.cmbRev.Clear
   On Error GoTo modErr1
   MDISect.ActiveForm.cmbRev.Clear
   sSql = "Qry_FillBomRev '" & Compress(sFillrevision) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, BmhRes, ES_FORWARD)
   If bSqlRows Then
      With BmhRes
         Do Until .EOF
            AddComboStr MDISect.ActiveForm.cmbRev.hwnd, "" & Trim(!BMHREV)
            .MoveNext
         Loop
         ClearResultSet BmhRes
      End With
   End If
   Set BmhRes = Nothing
   MouseCursor 0
   Exit Sub
   
modErr1:
   sProcName = "fillbomrev"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
End Sub



'11/21/06 Changed Name from GetDefaults

Public Sub GetRoutingIncrementDefault()
   Dim RdoDef As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT RTEINCREMENT FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDef, ES_FORWARD)
   If bSqlRows Then
      iAutoIncr = RdoDef!RTEINCREMENT
   Else
      iAutoIncr = 10
   End If
   ClearResultSet RdoDef
   If iAutoIncr <= 0 Then iAutoIncr = 10
   RdoDef.Close
   Exit Sub
   
modErr1:
   sProcName = "getdefaults"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Forms(1)
   
End Sub

Sub Main()
   Dim sAppTitle As String
   sAppTitle = Trim(App.Title)
   If App.PrevInstance Then
      On Error Resume Next
      App.Title = "EsiEngr"
      SysMsgBox.Width = 3800
      SysMsgBox.msg.Width = 3200
      SysMsgBox.tmr1.Enabled = True
      SysMsgBox.msg = sAppTitle & " Is Already Open."
      SysMsgBox.Show
      Sleep 5000
   End
   Exit Sub
End If
' Set the Module name before loading the form
sProgName = "Engineering"
MainLoad "engr"
GetFavorites "EsiEngr"
' MM 9/10/2009
'sProgName = "Engineering"
MDISect.Show

End Sub

''Pick up permissions for this user
'Public Sub GetSectionPermissions()
'    Dim RdoUsr As ADODB.Recordset
'    '11/21/06
'    User.Group1 = 1
'    User.Group2 = 1
'    User.Group3 = 1
'    User.Group4 = 1
'    User.Group5 = 1
'    User.Group6 = 1
'
'    On Error GoTo ModErr1
'    sSql = "SELECT USERREF,USERADDUSER,USERLEVEL," _
'        & "USERENGRGR1,USERENGRGR2,USERENGRGR3," _
'        & "USERENGRGR4,USERENGRGR5,USERENGRGR6 " _
'        & "FROM UsscTable WHERE USERREF='" _
'        & UCase$(cur.CurrentUser) & "'"
'    bSqlRows = clsADOCon.GetDataSet(sSql,RdoUsr, ES_FORWARD)
'        If bSqlRows Then
'            With RdoUsr
'                User.Adduser = !UserAddUser
'                User.Level = !UserLevel
'                User.Group1 = !USERENGRGR1
'                User.Group2 = !USERENGRGR2
'                User.Group3 = !USERENGRGR3
'                User.Group4 = !USERENGRGR4
'                User.Group5 = !USERENGRGR5
'                User.Group6 = !USERENGRGR6
'                ClearResultSet RdoUsr
'            End With
'        End If
'    Set RdoUsr = Nothing
'    Exit Sub
'
'ModErr1:
'    Resume modErr2:
'modErr2:
'    On Error GoTo 0
'
'End Sub
'
'8/9/05 Close open sets

Public Sub FormUnload(Optional bDontShowForm As Byte)
   Dim iList As Integer
   Dim iResultSets As Integer
   On Error Resume Next
   MDISect.lblBotPanel.Caption = MDISect.Caption
   bGoodBid = 0
'   TODO: Not sure if we need this
'   If Forms.Count < 3 Then
'
'      iResultSets = RdoCon.ADODB.Recordsets.Count
'      For iList = iResultSets - 1 To 0 Step -1
'         RdoCon.ADODB.Recordsets(iList).Close
'      Next
'   End If
   If bDontShowForm = 0 Then
      Select Case cUR.CurrentGroup
         Case "Rout"
            zGr1Rout.Show
         Case "Bomp"
            zGr2Bomp.Show
         Case "Docu"
            zGr3Docu.Show
         Case "Esti"
            If RunningBeta Then
               zGr5EstiPPI.Show
            Else
               zGr5Esti.Show
            End If
         Case "Tool"
            zGr4Tool.Show
      End Select
      Erase bActiveTab
      cUR.CurrentGroup = ""
   End If
End Sub


Public Sub FindShop(sGetShop As String)
   Dim RdoShp As ADODB.Recordset
   On Error Resume Next
   If Len(sGetShop) > 0 Then
      sSql = "Qry_GetShop '" & Compress(sGetShop) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoShp, ES_FORWARD)
      If bSqlRows Then
         MDISect.ActiveForm.cmbShp = "" & Trim(RdoShp!SHPNUM)
      Else
         MsgBox "Shop Wasn't Found.", vbInformation, MDISect.ActiveForm.Caption
         MDISect.ActiveForm.cmbShp = ""
      End If
      ClearResultSet RdoShp
   End If
   Set RdoShp = Nothing
   
End Sub


Public Sub FindPart()
   Dim RdoPrt As ADODB.Recordset
   On Error Resume Next
   If Len(Trim(MDISect.ActiveForm.cmbPrt)) Then
      sSql = "Qry_GetINVCfindPart '" & Compress(MDISect.ActiveForm.cmbPrt) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
      If bSqlRows Then
         With RdoPrt
            MDISect.ActiveForm.cmbPrt = "" & Trim(!PartNum)
            MDISect.ActiveForm.lblDsc = "" & Trim(!PADESC)
            MDISect.ActiveForm.lblTyp = !PALEVEL
         End With
      Else
         MDISect.ActiveForm.lblDsc = "*** Part Number Wasn't Found ***"
         MDISect.ActiveForm.lblTyp = ""
      End If
   Else
      MDISect.ActiveForm.cmbPrt = "NONE"
      MDISect.ActiveForm.lblDsc = ""
      MDISect.ActiveForm.lblTyp = ""
   End If
   Set RdoPrt = Nothing
   
End Sub


Public Sub FillRoutings()
   Dim RdoRtg As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "Qry_FillRoutings "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRtg, ES_FORWARD)
   If bSqlRows Then
      With RdoRtg
         Do Until .EOF
            AddComboStr MDISect.ActiveForm.cmbRte.hwnd, "" & Trim(!RTNUM)
            .MoveNext
         Loop
         ClearResultSet RdoRtg
      End With
   End If
   If MDISect.ActiveForm.cmbRte.ListCount > 0 Then _
      MDISect.ActiveForm.cmbRte.Text = MDISect.ActiveForm.cmbRte.List(0)
   Set RdoRtg = Nothing
   Exit Sub
   
modErr1:
   sProcName = "fillroutings"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Set RdoRtg = Nothing
   DoModuleErrors MDISect.ActiveForm
   
End Sub


Public Sub UpdateTables()
   If MDISect.bUnloading = 1 Then Exit Sub
   Dim RdoTest As ADODB.Recordset

   SaveSetting "Esi2000", "AppTitle", "engr", "ESI Engineering"
   SysOpen.Show
   SysOpen.prg1.Visible = True
   SysOpen.pnl = "Configuration Settings."
   SysOpen.pnl.Refresh
   SysOpen.prg1.Value = 10
   On Error Resume Next
   'Some moved to OldUpdate 10/10/02, 5/17/05, 10/6/06

   '   2/6/07
   '    '5/26/06
   '    KeysEstTables
   '    '5/26/06
   '    ConvertEstimateTables
   '    '5/30/06
   '    ConvertPartsListTables
   '    ConvertRoutingTables
   '    ConvertRoutingLibrary
   '    BuildKeys
   '    KeysEstTables2

   '6/20/06 Routing Pictures (with folder for storing temp files
   '*** Leave the next line in:
   If Dir("c:\Program Files\ES2000\Temp") = "" Then _
          MkDir "c:\Program Files\ES2000\Temp"

   SysOpen.prg1.Value = 40
   Sleep 500
   Err.Clear
   clsADOCon.ADOErrNum = 0
   '7/21/06
   sSql = "SELECT OPREF FROM RtpcTable WHERE OPREF='fubar'"
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   If clsADOCon.ADOErrNum = 0 Then
      Err.Clear
      sSql = "CREATE TABLE RtpcTable (" _
             & "OPREF CHAR(30) NOT NULL," _
             & "OPNO INT NOT NULL," _
             & "OPDESC VARCHAR(80) NULL DEFAULT(''), " _
             & "OPPICTURE IMAGE NULL)"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      If clsADOCon.ADOErrNum = 0 Then
         sSql = "ALTER TABLE RtpcTable ADD Constraint PK_RtpcTable_OPREF PRIMARY KEY CLUSTERED (OPREF,OPNO) " _
                & "WITH FILLFACTOR=80 "
         clsADOCon.ExecuteSql sSql 'rdExecDirect

         sSql = "ALTER TABLE RtpcTable ADD CONSTRAINT FK_RtpcTable_RthdTable FOREIGN KEY (OPREF) References RthdTable ON DELETE CASCADE ON UPDATE CASCADE"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
      End If
   End If
   '*End Pictures
   SysOpen.prg1.Value = 70
   Sleep 500
   GoTo modErr2
   Exit Sub

modErr1:
   Resume modErr2
modErr2:
   Set RdoTest = Nothing
   On Error GoTo 0
   SysOpen.Timer1.Enabled = True
   SysOpen.prg1.Value = 100
   SysOpen.Refresh
   Sleep 500

End Sub

Public Sub FillCustomerRFQs(frm As Form, sCustomer As String, Optional bShowNone As Boolean)
   Dim RdoCrq As ADODB.Recordset
   
   sCustomer = Compress(sCustomer)
   On Error Resume Next
   frm.cmbRfq.Clear
   sSql = "SELECT RFQREF,RFQCUST FROM RfqsTable WHERE RFQCUST='" _
          & sCustomer & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCrq, ES_FORWARD)
   If bShowNone Then AddComboStr frm.cmbRfq.hwnd, "NONE"
   If bSqlRows Then
      With RdoCrq
         Do Until .EOF
            AddComboStr frm.cmbRfq.hwnd, "" & Trim(!RFQREF)
            .MoveNext
         Loop
         ClearResultSet RdoCrq
      End With
      If frm.cmbRfq.ListCount > 0 Then
         frm.cmbRfq = frm.cmbRfq.List(0)
      Else
         If Not bShowNone Then
            frm.txtDsc = ""
            frm.txtBuy = ""
            frm.txtDte = Format(ES_SYSDATE, "mm/dd/yy")
            frm.txtDue = Format(ES_SYSDATE, "mm/dd/yy")
            frm.optCom.Value = vbUnchecked
            If frm.cmbRfq.ListCount > 0 Then _
               frm.cmbRfq = frm.cmbRfq.List(0)
         End If
      End If
   End If
   If frm.cmbRfq.ListCount > 0 Then _
      frm.cmbRfq = frm.cmbRfq.List(0)
   Set RdoCrq = Nothing
   Exit Sub
   
modErr1:
   sProcName = "fillcustomerrf"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors frm
   
End Sub

Public Function GetBidLabor(BidPart As String, BidNumber As Long, BidQty As Currency) As Byte
   Dim RdoLbr As ADODB.Recordset
   Dim bByte As Byte
   'Dim iCounter As Integer
   'Dim cQuantity As Currency
   'Dim cSetup As Currency
   'Dim cUnit As Currency
   Dim cRate As Currency
   'Dim cFoh As Currency
   Dim cFohRate As Currency
   Dim cLabor As Currency
'   Dim cUnitHours As Currency
'   Dim cUnitOverhead As Currency
'   Dim cUnitCost As Currency
   
   cUnitHours = 0
   cUnitCost = 0
   cUnitOverhead = 0
   cTotalRte = 0
   'cQuantity = 1
   On Error GoTo modErr1
   sProcName = "getbidlabor(rte)"
   
   'Routing
'   sSql = "Qry_GetBidLabor " & BidNumber & " "
'   sSql = "SELECT BIDRTEREF,BIDRTEUNIT,BIDRTESETUP,BIDRTEHOURS," & vbCrLf _
'      & "BIDRTERATE,BIDRTEFOHRATE" & vbCrLf _
'      & "FROM EsrtTable WHERE BIDRTEREF=" & BidNumber
'
'   bSqlRows = clsADOCon.GetDataSet(sSql,RdoLbr, ES_FORWARD)
'   If bSqlRows Then
'      With RdoLbr
'         Do Until .EOF
'            cUnit = !BIDRTEUNIT
'            '''cHours = !BIDRTEHOURS
'            cUnit = !BIDRTEUNIT
'            cSetup = !BIDRTESETUP
'            cHours = !BIDRTEUNIT
'            cRate = !BIDRTERATE
'            cFohRate = !BIDRTEFOHRATE
'            cHours = (cUnit * cQuantity) + cSetup
'            cLabor = cHours * cRate
'            cFoh = cFohRate * cLabor
'            cUnitCost = cUnitCost + (cLabor + cFoh)
'            cUnitHours = cUnitHours + cHours
'            cUnitOverhead = cUnitOverhead + cFoh
'            .MoveNext
'         Loop
'         ClearResultSet RdoLbr
'      End With
'   End If
   
   Dim rdo As ADODB.Recordset
   sSql = "SELECT ISNULL(SUM(BIDRTEUNIT + BIDRTESETUP/" & BidQty & "),0) as UnitHours," & vbCrLf _
      & "ISNULL(SUM(BIDRTERATE * (BIDRTEUNIT + BIDRTESETUP/" & BidQty & ")),0) as UnitCost," & vbCrLf _
      & "ISNULL(SUM(BIDRTEFOHRATE * BIDRTERATE * (BIDRTEUNIT + BIDRTESETUP/" & BidQty & ")),0) as UnitOverhead" & vbCrLf _
      & "FROM EsrtTable WHERE BIDRTEREF=" & BidNumber
   If clsADOCon.GetDataSet(sSql, rdo) Then
      cUnitHours = rdo!UnitHours
      cUnitCost = rdo!UnitCost
      cUnitOverhead = rdo!UnitOverhead
   End If
   
   'Bills Of Material
   sProcName = "getbidlabor(bom)"
   'sSql = "Qry_GetBidBomLabor " & BidNumber & ",'" & BidPart & "'"
   sSql = "SELECT BIDBOMREF,BIDBOMASSYPART,BIDBOMPARTREF,BIDBOMQTYREQD," & vbCrLf _
      & "BIDBOMLABOR,BIDBOMLABOROH,BIDBOMLABORHRS FROM EsbmTable" & vbCrLf _
      & "WHERE BIDBOMREF=" & BidNumber & " AND BIDBOMASSYPART='" & BidPart & "'"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLbr, ES_FORWARD)
   If bSqlRows Then
      With RdoLbr
         Do Until .EOF
            sProcName = "getbidlabor(bom)"
            'iCounter = iCounter + 1
            cBomHours = !BIDBOMLABORHRS
            cBomRate = !BIDBOMLABOR
            cBomFoh = (!BIDBOMLABOROH / 100)
            cBomFoh = (cBomRate * cBomFoh) * cBomHours
'            cBomLabor = cBomFoh + (cBomHours * cBomRate)
'            cUnitCost = cUnitCost + cBomLabor
            cUnitCost = cUnitCost + cBomHours * cBomRate
            cUnitOverhead = cUnitOverhead + cBomFoh
            cUnitHours = cUnitHours + cBomHours
            bByte = GetBidLaborNext(Trim(!BIDBOMPARTREF), BidNumber, !BIDBOMQTYREQD, 2)
            .MoveNext
         Loop
         ClearResultSet RdoLbr
      End With
   End If
   'cQuantity = 1
   
   'Totals
   Dim cUnitTotalCost As Currency
   cUnitTotalCost = cUnitCost + cUnitOverhead
   If cUnitHours > 0 Then
      EstiESe02a.lblHours = Format(cUnitHours, "0.000") 'Format(cUnitHours / cQuantity, ES_QuantityDataFormat)
      EstiESe02a.lblUnitLaborCost = Format(cUnitCost, ES_MoneyFormat)
      EstiESe02a.lblUnitOverheadCost = Format(cUnitOverhead, ES_MoneyFormat) 'Format(cUnitOverhead / cQuantity, ES_QuantityDataFormat)
      EstiESe02a.LblUnitLabor = Format(cUnitTotalCost, ES_MoneyFormat)  'Format(cUnitCost / cQuantity, ES_QuantityDataFormat)
      'cRate = (cUnitCost - cUnitOverhead) / cUnitHours
      cRate = cUnitCost / cUnitHours
      EstiESe02a.lblRate = Format(cRate, ES_MoneyFormat)
      EstiESe02a.lblEstTotalLabor = Format(cUnitTotalCost * BidQty, ES_MoneyFormat)
   Else
      EstiESe02a.lblHours = ES_MoneyFormat
      EstiESe02a.lblUnitOverheadCost = ES_MoneyFormat
      EstiESe02a.LblUnitLabor = ES_MoneyFormat
      EstiESe02a.lblRate = ES_MoneyFormat
   End If
   
   'cLabor = cUnitCost - cUnitOverhead
   cLabor = cUnitCost
   If RunningBeta Then
'      If cLabor > 0 Then
'         cFohRate = (cUnitOverhead) / cLabor
'         ppiESe02a.lblFohRate = Format(cFohRate, ES_QuantityDataFormat)
'      Else
'         ppiESe02a.lblFohRate = ".000"
'      End If
   Else
      If cLabor > 0 Then
         cFohRate = (cUnitOverhead) / cLabor
         EstiESe02a.lblFohRate = Format(cFohRate, ES_QuantityDataFormat)
      Else
         EstiESe02a.lblFohRate = ".000"
      End If
   End If
   Set RdoLbr = Nothing
   MouseCursor 0
   Exit Function
   
modErr1:
   On Error Resume Next
   EstiESe02a.lblHours = ES_MoneyFormat
   EstiESe02a.lblUnitOverheadCost = ES_MoneyFormat
   EstiESe02a.LblUnitLabor = ES_MoneyFormat
   EstiESe02a.lblRate = ES_MoneyFormat
   EstiESe02a.lblFohRate = ES_MoneyFormat
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
End Function

Public Function GetBidServices(frm As Form, EstQty As Currency) As Byte
   Dim RdoOsp As ADODB.Recordset
   
   Dim cQuantity As Currency
   Dim cUnitOsp As Currency
   'Dim cTotalOsp As Currency
   
   cQuantity = Val(frm.txtQty)
   If cQuantity = 0 Then cQuantity = 1
   sSql = "SELECT BIDOSREF,BIDOSTOTALCOST,BIDOSLOT FROM " _
          & "EsosTable WHERE BIDOSREF=" & Val(frm.cmbBid) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoOsp, ES_FORWARD)
   If bSqlRows Then
      With RdoOsp
         Do Until .EOF
            If !BIDOSLOT = 1 Then
               cUnitOsp = cUnitOsp + (!BIDOSTOTALCOST / cQuantity)
            Else
               cUnitOsp = cUnitOsp + !BIDOSTOTALCOST
            End If
            'cTotalOsp = cTotalOsp + !BIDOSTOTALCOST
            .MoveNext
         Loop
         ClearResultSet RdoOsp
      End With
   End If
   
   frm.lblUnitServices = Format(cUnitOsp, ES_MoneyFormat)
   frm.lblTotServices = Format(cUnitOsp * cQuantity, ES_MoneyFormat)
   Set RdoOsp = Nothing
   Exit Function
   
modErr1:
   sProcName = "getbidserv"
   frm.lblUnitServices = "0.000"
   frm.lblTotServices = "0.000"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
End Function

'9/18/03

Public Sub GetEstimatingDefaults(MBurden As Currency, FOverHead As Currency, LRate As Currency)
   Dim RdoEst As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT EstMatlBurden, EstFactoryOverHead,EstLaborRate " _
          & "FROM Preferences Where PreRecord=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoEst, ES_FORWARD)
   If bSqlRows Then
      With RdoEst
         If Not IsNull(.Fields(0)) Then
            MBurden = .Fields(0)
         Else
            MBurden = 0
         End If
         If Not IsNull(.Fields(1)) Then
            FOverHead = .Fields(1)
         Else
            FOverHead = 0
         End If
         If Not IsNull(.Fields(2)) Then
            LRate = .Fields(2)
         Else
            LRate = 0
         End If
         ClearResultSet RdoEst
      End With
   Else
      MBurden = 0
      FOverHead = 0
      LRate = 0
   End If
   Set RdoEst = Nothing
   Exit Sub
   
modErr1:
   MBurden = 0
   FOverHead = 0
   LRate = 0
End Sub

'4/7/04

Public Function GetRoutShop(ShopRef As String) As String
   Dim RdoShp As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "Qry_GetShopBasics '" & ShopRef & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoShp, ES_FORWARD)
   If bSqlRows Then
      With RdoShp
         GetRoutShop = "" & Trim(!SHPNUM)
         ClearResultSet RdoShp
      End With
   Else
      GetRoutShop = ""
   End If
   Set RdoShp = Nothing
   Exit Function
   
modErr1:
   sProcName = "getroutshop"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
End Function

Public Function GetRoutCenter(CenterRef As String) As String
   Dim RdoShp As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "Qry_GetWorkCenter '" & CenterRef & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoShp, ES_FORWARD)
   If bSqlRows Then
      With RdoShp
         GetRoutCenter = "" & Trim(!WCNNUM)
         ClearResultSet RdoShp
      End With
   Else
      GetRoutCenter = ""
   End If
   Set RdoShp = Nothing
   Exit Function
   
modErr1:
   sProcName = "getroutcenter"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
End Function

Public Function FindToolList(ToolNumber As String, ToolDesc As String, Optional _
                             DontShow As Byte) As String
   Dim RdoTlst As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "Qry_GetToolList'" & Compress(ToolNumber) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTlst, ES_FORWARD)
   If bSqlRows Then
      With RdoTlst
         On Error Resume Next
         FindToolList = "" & Trim(!TOOLLIST_NUM)
         If DontShow = 0 Then MDISect.ActiveForm.lblLst = "" & Trim(!TOOLLIST_DESC)
         If Err > 0 Then _
            If DontShow = 0 Then MDISect.ActiveForm.txtDsc = "" & Trim(!TOOLLIST_DESC)
         ClearResultSet RdoTlst
      End With
   Else
      On Error Resume Next
      FindToolList = ""
      If DontShow = 0 Then MDISect.ActiveForm.lblLst = "*** Tool List Wasn't Found ***"
      If Err > 0 Then _
         If DontShow = 0 Then MDISect.ActiveForm.txtDsc = "*** Tool List Wasn't Found ***"
   End If
   Set RdoTlst = Nothing
   Exit Function
   
modErr1:
   sProcName = "findtoollist"
   FindToolList = ""
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
End Function

'10/13/04 See GetBidLabor. Recursive element to gather labor

Private Function GetBidLaborNext(AssemblyPart As String, BidNumber As Long, _
   BidQuantity As Currency, BomLevel As Integer) As Byte
   
Debug.Print "Level " & BomLevel & ": " & AssemblyPart
   
   'BomLevel = explosion level should be 2 or greater
   
   Dim RdoBlb As ADODB.Recordset
   Dim bByte As Byte
   
   sProcName = "GetBidLaborNext"
   
   'watch for infinite loops
   If BomLevel > 10 Then
      MsgBox "WARNING: More than 10 explosion levels." & vbCrLf _
      & "A part in the suspected loop is " & Trim(AssemblyPart) & "." & vbCrLf _
         & "Additional levels ignored."
      Exit Function
   End If
   
   '       "SELECT BIDBOMREF,BIDBOMASSYPART,BIDBOMPARTREF,BIDBOMQTYREQD," _
   '        & "BIDBOMLABOR,BIDBOMLABOROH,BIDBOMLABORHRS FROM EsbmTable WHERE " _
   '        & "(BIDBOMREF=" & BidNumber & " AND BIDBOMASSYPART='" & AssemblyPart & "')"
   
'SELECT BIDBOMREF,BIDBOMASSYPART,BIDBOMPARTREF,BIDBOMQTYREQD,
'BIDBOMLABOR,BIDBOMLABOROH,BIDBOMLABORHRS FROM EsbmTable
'WHERE (BIDBOMREF=@bidnumber AND BIDBOMASSYPART=@bidpart)
   
   sSql = "Qry_GetBidBomLabor " & BidNumber & ",'" & Trim(AssemblyPart) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBlb, ES_FORWARD)
   If bSqlRows Then
      With RdoBlb
         Do Until .EOF
'Debug.Print "      " & BomLevel & ": " & !BIDBOMPARTREF

            sProcName = "getbidlabornext"
            'iCounter = iCounter + 1
            'cBomHours = !BIDBOMLABORHRS
            'cBomRate = !BIDBOMLABOR
            'cBomFoh = (!BIDBOMLABOROH / 100)
            'cBomFoh = (cBomRate * cBomFoh) * cBomHours
            'cBomLabor = cBomFoh + (cBomHours * cBomRate)
Debug.Print "      " & BomLevel & ": " & !BIDBOMPARTREF & " " _
   & !BIDBOMLABORHRS & " HRS @ " & !BIDBOMLABOR
            'cTotalHrs = cTotalHrs + (BidQuantity * cBomHours)
            'cTotalLabr = cTotalLabr + (BidQuantity * cBomLabor)
            'cTotalFoh = cTotalFoh + (BidQuantity * cBomFoh)
            cUnitHours = cUnitHours + !BIDBOMLABORHRS * !BIDBOMQTYREQD * BidQuantity     'unit hours
            cUnitCost = cUnitCost + !BIDBOMLABORHRS * !BIDBOMLABOR * !BIDBOMQTYREQD * BidQuantity        'unit labor
            cUnitOverhead = cUnitOverhead + !BIDBOMLABORHRS * !BIDBOMLABOR _
               * !BIDBOMQTYREQD * BidQuantity * !BIDBOMLABOROH / 100   'unit overhead
            
            bByte = GetBidLaborNext(!BIDBOMPARTREF, BidNumber, !BIDBOMQTYREQD * BidQuantity, BomLevel + 1)
            .MoveNext
         Loop
         
         
         ClearResultSet RdoBlb
      End With
   End If
   Set RdoBlb = Nothing
End Function


'2/14/06 Added to find Allowed functions (Admin/Estimation parameters

Public Sub GetEstimatingPermissions(AllowScrap As Byte, AllowGna As Byte, AllowProfit As Byte)
   Dim RdoPermit As ADODB.Recordset
   On Error Resume Next
   sSql = "SELECT ESTOVERWRITESCRAP,ESTOVERWRITEGNA,ESTOVERWRITEPROFIT " _
          & "FROM Preferences WHERE PreRecord=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPermit, ES_FORWARD)
   If bSqlRows Then
      With RdoPermit
         AllowScrap = !ESTOVERWRITESCRAP
         AllowGna = !ESTOVERWRITEGNA
         AllowProfit = !ESTOVERWRITEPROFIT
         .Cancel
      End With
      ClearResultSet RdoPermit
   End If
   Set RdoPermit = Nothing
   
End Sub


Public Function CheckBidEntries(GoodPart As Byte, GoodCustomer As Byte) As Byte
   On Error Resume Next
   If Trim(MDISect.ActiveForm.txtPrt) = "" Or _
             Trim(MDISect.ActiveForm.cmbCst) = "" Or GoodPart = 0 Or GoodCustomer = 0 _
             Then CheckBidEntries = 1
      If CheckBidEntries = 1 Then
         MsgBox "All Estimates Require A Valid Part Number And " & vbCrLf _
            & "A Valid Customer Or They Will Be Deleted.", _
            vbExclamation, MDISect.ActiveForm.Caption
      End If
      
   End Function
   
   
   Public Sub FillEstimateCombo(frm As Form, BidClass As String)
      Dim RdoBids As ADODB.Recordset
      Dim iList As Integer
      
      On Error GoTo modErr1
      BidClass = UCase$(BidClass)
      If BidClass = "FULL" Then
         sSql = "Qry_EstimateFullDesc"
      ElseIf BidClass = "QWIK" Then
         sSql = "Qry_EstimateQwikDesc"
      Else
         sSql = "Qry_EstimateAllDesc"
      End If
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoBids, ES_FORWARD)
      If bSqlRows Then
         With RdoBids
            frm.cmbBid = "" & Trim(!BIDNUM)
            frm.cmbCls = "" & Trim(!BIDPRE)
            Do Until .EOF
               iList = iList + 1
               If iList > 500 Then Exit Do
               AddComboStr frm.cmbBid.hwnd, "" & Trim(!BIDNUM)
               .MoveNext
            Loop
            ClearResultSet RdoBids
         End With
      End If
      Set RdoBids = Nothing
      Exit Sub
      
modErr1:
      Err.Clear
      
   End Sub
   
   Private Sub KeysEstTables()
      On Error Resume Next
      Err.Clear
      sSql = "SELECT EstimatingKeys FROM Preferences WHERE PreRecord=1"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      If Err > 0 Then
         Err.Clear
         sSql = "DELETE FROM EsosTable " & vbCrLf _
                & "FROM EsosTable LEFT JOIN EstiTable ON EsosTable.BIDOSREF = EstiTable.BIDREF " & vbCrLf _
                & "WHERE (EstiTable.BIDREF Is Null)"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
         
         sSql = "DELETE FROM EsbmTable " & vbCrLf _
                & "FROM EsbmTable LEFT JOIN EstiTable ON EsbmTable.BIDBOMREF = EstiTable.BIDREF " & vbCrLf _
                & "WHERE (EstiTable.BIDREF Is Null)"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
         
         sSql = "DELETE FROM EsplTable " & vbCrLf _
                & "FROM EsplTable LEFT JOIN EstiTable ON EsplTable.BIDPLREF = EstiTable.BIDREF " & vbCrLf _
                & "WHERE (EstiTable.BIDREF Is Null)"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
         
         sSql = "DELETE FROM EsrtTable " & vbCrLf _
                & "FROM EsrtTable LEFT JOIN EstiTable ON EsrtTable.BIDRTEREF = EstiTable.BIDREF " & vbCrLf _
                & "WHERE (EstiTable.BIDREF Is Null)"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
         
         Err.Clear
         clsADOCon.ADOErrNum = 0
         clsADOCon.BeginTrans
         sSql = "ALTER TABLE Preferences ADD EstimatingKeys smalldatetime null"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
         
         sSql = "DROP INDEX EstiTable.BidRef"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
         
         sSql = "ALTER TABLE EstiTable ALTER COLUMN BIDREF INT NOT NULL"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
         
         sSql = "ALTER TABLE EstiTable ADD Constraint PK_EstiTable_BIDREF PRIMARY KEY CLUSTERED (BIDREF) " _
                & "WITH FILLFACTOR=80 "
         clsADOCon.ExecuteSql sSql 'rdExecDirect
         
         sSql = "ALTER TABLE EsbmTable ADD CONSTRAINT FK_EsbmTable_EstiTable FOREIGN KEY (BIDBOMREF) References EstiTable ON DELETE CASCADE ON UPDATE CASCADE"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
         
         sSql = "ALTER TABLE EsosTable ADD CONSTRAINT FK_EsosTable_EstiTable FOREIGN KEY (BIDOSREF) References EstiTable ON DELETE CASCADE ON UPDATE CASCADE"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
         
         sSql = "ALTER TABLE EsplTable ADD CONSTRAINT FK_EsplTable_EstiTable FOREIGN KEY (BIDPLREF) References EstiTable ON DELETE CASCADE ON UPDATE CASCADE"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
         
         sSql = "ALTER TABLE EsrtTable ADD CONSTRAINT FK_EsrtTable_EstiTable FOREIGN KEY (BIDRTEREF) References EstiTable ON DELETE CASCADE ON UPDATE CASCADE"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
         If clsADOCon.ADOErrNum = 0 Then
            clsADOCon.CommitTrans
            sSql = "UPDATE Preferences SET EstimatingKeys='" & Format(ES_SYSDATE, "mm/dd/yy") & "' WHERE PreRecord=1"
            clsADOCon.ExecuteSql sSql 'rdExecDirect
         Else
            clsADOCon.RollbackTrans
            clsADOCon.ADOErrNum = 0
         End If
      End If
      
      
   End Sub
   

   Private Function CheckConvErrors() As Byte
      Dim iColCounter As Integer
      CheckConvErrors = 0
      ' TODO: ADD ADO Error
      For Each ER In clsADOCon.ConnectionObject.Errors      'clsADOCon.ConnectionObject.Errors
         If Left(ER.Description, 5) = "22003" Then
            iColCounter = iColCounter + 1
            CheckConvErrors = 1
         End If
      Next ER
      
   End Function
   
   Private Function GetConstraint(sDescription As String) As String
      Dim bByte As Byte
      GetConstraint = ""
      bByte = InStr(1, sDescription, "DF_")
      If bByte > 0 Then
         GetConstraint = Mid$(sDescription, bByte, Len(sDescription))
         bByte = InStr(1, GetConstraint, "'")
         If bByte > 0 Then GetConstraint = Left$(GetConstraint, bByte - 1)
      Else
         bByte = InStr(1, sDescription, "DEFZERO")
         If bByte > 0 Then
            sSql = "sp_unbindefault '" & RdoCol.Fields(2) & "." & RdoCol.Fields(3) & "'"
            clsADOCon.ExecuteSql sSql 'rdExecDirect
         End If
      End If
      
      
   End Function
   
   
   '5/30/06 Includes keys (End)

   Private Sub ConvertPartsListTables()
      Dim bBadCol As Byte
      Dim sconstraint As String
      Err.Clear
      On Error Resume Next
      'Start BmplTable
      'BMQTYREQD
      sSql = "sp_columns @table_name=BmplTable,@column_name=BMQTYREQD"
      'Set RdoCol = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
      Set RdoCol = clsADOCon.GetRecordSet(sSql, adOpenForwardOnly)
      With RdoCol
         If Not .BOF And Not .EOF Then
            Err.Clear
            If Not IsNull(.Fields(5)) Then
               If .Fields(5) = "real" Then
                  sSql = "ALTER TABLE BmplTable ALTER COLUMN BMQTYREQD dec(12,4)"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  If Err > 0 Then
                     For Each ER In clsADOCon.ConnectionObject.Errors
                        sconstraint = GetConstraint(ER.Description)
                        If sconstraint <> "" Then Exit For
                     Next ER
                  End If
                  Err.Clear
                  If sconstraint <> "" Then
                     sSql = "ALTER TABLE BmplTable DROP " & sconstraint
                     clsADOCon.ExecuteSql sSql 'rdExecDirect
                  End If
                  sSql = "ALTER TABLE BmplTable ALTER COLUMN BMQTYREQD DEC(12,4) "
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  
                  bBadCol = CheckConvErrors()
                  sSql = "sp_bindefault DEFZERO, 'BmplTable.BMQTYREQD'"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
               Else
                  GoTo Keys
               End If
            End If
         End If
      End With
      'BMCONVERSION
      sSql = "sp_columns @table_name=BmplTable,@column_name=BMCONVERSION"
      'Set RdoCol = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
      Set RdoCol = clsADOCon.GetRecordSet(sSql, adOpenForwardOnly)
      With RdoCol
         If Not .BOF And Not .EOF Then
            Err.Clear
            If Not IsNull(.Fields(5)) Then
               If .Fields(5) = "real" Then
                  sSql = "ALTER TABLE BmplTable ALTER COLUMN BMCONVERSION dec(12,4)"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  If Err > 0 Then
                     For Each ER In clsADOCon.ConnectionObject.Errors
                        sconstraint = GetConstraint(ER.Description)
                        If sconstraint <> "" Then Exit For
                     Next ER
                  End If
                  Err.Clear
                  If sconstraint <> "" Then
                     sSql = "ALTER TABLE BmplTable DROP " & sconstraint
                     clsADOCon.ExecuteSql sSql 'rdExecDirect
                  End If
                  sSql = "ALTER TABLE BmplTable ALTER COLUMN BMCONVERSION DEC(12,4) "
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  
                  bBadCol = CheckConvErrors()
                  sSql = "sp_bindefault DEFZERO, 'BmplTable.BMCONVERSION'"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
               End If
            End If
         End If
      End With
      'BMESTCOST
      sSql = "sp_columns @table_name=BmplTable,@column_name=BMESTCOST"
      'Set RdoCol = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
      Set RdoCol = clsADOCon.GetRecordSet(sSql, adOpenForwardOnly)
      With RdoCol
         If Not .BOF And Not .EOF Then
            Err.Clear
            If Not IsNull(.Fields(5)) Then
               If .Fields(5) = "real" Then
                  sSql = "ALTER TABLE BmplTable ALTER COLUMN BMESTCOST dec(12,4)"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  If Err > 0 Then
                     For Each ER In clsADOCon.ConnectionObject.Errors
                        sconstraint = GetConstraint(ER.Description)
                        If sconstraint <> "" Then Exit For
                     Next ER
                  End If
                  Err.Clear
                  If sconstraint <> "" Then
                     sSql = "ALTER TABLE BmplTable DROP " & sconstraint
                     clsADOCon.ExecuteSql sSql 'rdExecDirect
                  End If
                  sSql = "ALTER TABLE BmplTable ALTER COLUMN BMESTCOST DEC(12,4) "
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  
                  bBadCol = CheckConvErrors()
                  sSql = "sp_bindefault DEFZERO, 'BmplTable.BMESTCOST'"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
               End If
            End If
         End If
      End With
      'BMADDER
      sSql = "sp_columns @table_name=BmplTable,@column_name=BMADDER"
      'Set RdoCol = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
      Set RdoCol = clsADOCon.GetRecordSet(sSql, adOpenForwardOnly)
      With RdoCol
         If Not .BOF And Not .EOF Then
            Err.Clear
            If Not IsNull(.Fields(5)) Then
               If .Fields(5) = "real" Then
                  sSql = "ALTER TABLE BmplTable ALTER COLUMN BMADDER dec(12,4)"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  If Err > 0 Then
                     For Each ER In clsADOCon.ConnectionObject.Errors
                        sconstraint = GetConstraint(ER.Description)
                        If sconstraint <> "" Then Exit For
                     Next ER
                  End If
                  Err.Clear
                  If sconstraint <> "" Then
                     sSql = "ALTER TABLE BmplTable DROP " & sconstraint
                     clsADOCon.ExecuteSql sSql 'rdExecDirect
                  End If
                  sSql = "ALTER TABLE BmplTable ALTER COLUMN BMADDER DEC(12,4) "
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  
                  bBadCol = CheckConvErrors()
                  sSql = "sp_bindefault DEFZERO, 'BmplTable.BMADDER'"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
               End If
            End If
         End If
      End With
      'BMPURCONV
      sSql = "sp_columns @table_name=BmplTable,@column_name=BMPURCONV"
      'Set RdoCol = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
      Set RdoCol = clsADOCon.GetRecordSet(sSql, adOpenForwardOnly)
      With RdoCol
         If Not .BOF And Not .EOF Then
            Err.Clear
            If Not IsNull(.Fields(5)) Then
               If .Fields(5) = "real" Then
                  sSql = "ALTER TABLE BmplTable ALTER COLUMN BMPURCONV dec(12,4)"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  If Err > 0 Then
                     For Each ER In clsADOCon.ConnectionObject.Errors
                        sconstraint = GetConstraint(ER.Description)
                        If sconstraint <> "" Then Exit For
                     Next ER
                  End If
                  Err.Clear
                  If sconstraint <> "" Then
                     sSql = "ALTER TABLE BmplTable DROP " & sconstraint
                     clsADOCon.ExecuteSql sSql 'rdExecDirect
                  End If
                  sSql = "ALTER TABLE BmplTable ALTER COLUMN BMPURCONV DEC(12,4) "
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  
                  bBadCol = CheckConvErrors()
                  sSql = "sp_bindefault DEFZERO, 'BmplTable.BMPURCONV'"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
               End If
            End If
         End If
      End With
      'BMSETUP
      sSql = "sp_columns @table_name=BmplTable,@column_name=BMSETUP"
      'Set RdoCol = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
      Set RdoCol = clsADOCon.GetRecordSet(sSql, adOpenForwardOnly)
      With RdoCol
         If Not .BOF And Not .EOF Then
            Err.Clear
            If Not IsNull(.Fields(5)) Then
               If .Fields(5) = "real" Then
                  sSql = "ALTER TABLE BmplTable ALTER COLUMN BMSETUP dec(12,4)"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  If Err > 0 Then
                     For Each ER In clsADOCon.ConnectionObject.Errors
                        sconstraint = GetConstraint(ER.Description)
                        If sconstraint <> "" Then Exit For
                     Next ER
                  End If
                  Err.Clear
                  If sconstraint <> "" Then
                     sSql = "ALTER TABLE BmplTable DROP " & sconstraint
                     clsADOCon.ExecuteSql sSql 'rdExecDirect
                  End If
                  sSql = "ALTER TABLE BmplTable ALTER COLUMN BMSETUP DEC(12,4) "
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  
                  bBadCol = CheckConvErrors()
                  sSql = "sp_bindefault DEFZERO, 'BmplTable.BMSETUP'"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
               End If
            End If
         End If
      End With
      'BMESTLABOR
      sSql = "sp_columns @table_name=BmplTable,@column_name=BMESTLABOR"
      'Set RdoCol = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
      Set RdoCol = clsADOCon.GetRecordSet(sSql, adOpenForwardOnly)
      With RdoCol
         If Not .BOF And Not .EOF Then
            Err.Clear
            If Not IsNull(.Fields(5)) Then
               If .Fields(5) = "real" Then
                  sSql = "ALTER TABLE BmplTable ALTER COLUMN BMESTLABOR dec(12,4)"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  If Err > 0 Then
                     For Each ER In clsADOCon.ConnectionObject.Errors
                        sconstraint = GetConstraint(ER.Description)
                        If sconstraint <> "" Then Exit For
                     Next ER
                  End If
                  Err.Clear
                  If sconstraint <> "" Then
                     sSql = "ALTER TABLE BmplTable DROP " & sconstraint
                     clsADOCon.ExecuteSql sSql 'rdExecDirect
                  End If
                  sSql = "ALTER TABLE BmplTable ALTER COLUMN BMESTLABOR DEC(12,4) "
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  
                  bBadCol = CheckConvErrors()
                  sSql = "sp_bindefault DEFZERO, 'BmplTable.BMESTLABOR'"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
               End If
            End If
         End If
      End With
      'BMESTLABOROH
      sSql = "sp_columns @table_name=BmplTable,@column_name=BMESTLABOROH"
      'Set RdoCol = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
      Set RdoCol = clsADOCon.GetRecordSet(sSql, adOpenForwardOnly)
      With RdoCol
         If Not .BOF And Not .EOF Then
            Err.Clear
            If Not IsNull(.Fields(5)) Then
               If .Fields(5) = "real" Then
                  sSql = "ALTER TABLE BmplTable ALTER COLUMN BMESTLABOROH dec(12,4)"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  If Err > 0 Then
                     For Each ER In clsADOCon.ConnectionObject.Errors
                        sconstraint = GetConstraint(ER.Description)
                        If sconstraint <> "" Then Exit For
                     Next ER
                  End If
                  Err.Clear
                  If sconstraint <> "" Then
                     sSql = "ALTER TABLE BmplTable DROP " & sconstraint
                     clsADOCon.ExecuteSql sSql 'rdExecDirect
                  End If
                  sSql = "ALTER TABLE BmplTable ALTER COLUMN BMESTLABOROH DEC(12,4) "
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  
                  bBadCol = CheckConvErrors()
                  sSql = "sp_bindefault DEFZERO, 'BmplTable.BMESTLABOROH'"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
               End If
            End If
         End If
      End With
      'BMESTMATERIAL
      sSql = "sp_columns @table_name=BmplTable,@column_name=BMESTMATERIAL"
      'Set RdoCol = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
      Set RdoCol = clsADOCon.GetRecordSet(sSql, adOpenForwardOnly)
      With RdoCol
         If Not .BOF And Not .EOF Then
            Err.Clear
            If Not IsNull(.Fields(5)) Then
               If .Fields(5) = "real" Then
                  sSql = "ALTER TABLE BmplTable ALTER COLUMN BMESTMATERIAL dec(12,4)"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  If Err > 0 Then
                     For Each ER In clsADOCon.ConnectionObject.Errors
                        sconstraint = GetConstraint(ER.Description)
                        If sconstraint <> "" Then Exit For
                     Next ER
                  End If
                  Err.Clear
                  If sconstraint <> "" Then
                     sSql = "ALTER TABLE BmplTable DROP " & sconstraint
                     clsADOCon.ExecuteSql sSql 'rdExecDirect
                  End If
                  sSql = "ALTER TABLE BmplTable ALTER COLUMN BMESTMATERIAL DEC(12,4) "
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  
                  bBadCol = CheckConvErrors()
                  sSql = "sp_bindefault DEFZERO, 'BmplTable.BMESTMATERIAL'"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
               End If
            End If
         End If
      End With
      'BMESTMATERIALBRD
      sSql = "sp_columns @table_name=BmplTable,@column_name=BMESTMATERIALBRD"
      'Set RdoCol = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
      Set RdoCol = clsADOCon.GetRecordSet(sSql, adOpenForwardOnly)
      With RdoCol
         If Not .BOF And Not .EOF Then
            Err.Clear
            If Not IsNull(.Fields(5)) Then
               If .Fields(5) = "real" Then
                  sSql = "ALTER TABLE BmplTable ALTER COLUMN BMESTMATERIALBRD dec(12,4)"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  If Err > 0 Then
                     For Each ER In clsADOCon.ConnectionObject.Errors
                        sconstraint = GetConstraint(ER.Description)
                        If sconstraint <> "" Then Exit For
                     Next ER
                  End If
                  Err.Clear
                  If sconstraint <> "" Then
                     sSql = "ALTER TABLE BmplTable DROP " & sconstraint
                     clsADOCon.ExecuteSql sSql 'rdExecDirect
                  End If
                  sSql = "ALTER TABLE BmplTable ALTER COLUMN BMESTMATERIALBRD DEC(12,4) "
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  
                  bBadCol = CheckConvErrors()
                  sSql = "sp_bindefault DEFZERO, 'BmplTable.BMESTMATERIALBRD'"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
               End If
            End If
         End If
      End With
      'End BmplTable
      
Keys:
      Err.Clear
      clsADOCon.ADOErrNum = 0
      sSql = "DROP INDEX BmhdTable.BmhRef"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      If clsADOCon.ADOErrNum = 0 Then
         sSql = "DROP INDEX BmhdTable.BmhRev"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
         
         sSql = "ALTER TABLE BmhdTable ALTER COLUMN BMHREF CHAR(30) NOT NULL"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
         
         sSql = "ALTER TABLE BmhdTable ALTER COLUMN BMHREV CHAR(4) NOT NULL"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
         
         sSql = "ALTER TABLE BmhdTable ADD Constraint PK_BmhdTable_BMHREF PRIMARY KEY CLUSTERED (BMHREF,BMHREV) " _
                & "WITH FILLFACTOR=80 "
         clsADOCon.ExecuteSql sSql 'rdExecDirect
         
         'No action on cascades SQL Server won't buy it. (possible circles)
         sSql = "ALTER TABLE BmplTable ADD CONSTRAINT FK_BmplTable_BmhdTable FOREIGN KEY (BMASSYPART,BMREV) References BmhdTable"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
      End If
      
   End Sub
   
   '5/30/06 Includes keys (End)

   Private Sub ConvertRoutingTables()
      Dim bBadCol As Byte
      Dim sconstraint As String
      Err.Clear
      'GoTo Keys
      On Error Resume Next
      'RTQUEUEHRS
      sSql = "sp_columns @table_name=RthdTable,@column_name=RTQUEUEHRS"
      'Set RdoCol = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
      Set RdoCol = clsADOCon.GetRecordSet(sSql, adOpenForwardOnly)
      With RdoCol
         If Not .BOF And Not .EOF Then
            Err.Clear
            If Not IsNull(.Fields(5)) Then
               If .Fields(5) = "real" Then 'See Else
                  sSql = "ALTER TABLE RthdTable ALTER COLUMN RTQUEUEHRS dec(9,4)"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  If Err > 0 Then
                     For Each ER In clsADOCon.ConnectionObject.Errors
                        sconstraint = GetConstraint(ER.Description)
                        If sconstraint <> "" Then Exit For
                     Next ER
                  End If
                  Err.Clear
                  If sconstraint <> "" Then
                     sSql = "ALTER TABLE RthdTable DROP " & sconstraint
                     clsADOCon.ExecuteSql sSql 'rdExecDirect
                  End If
                  sSql = "ALTER TABLE RthdTable ALTER COLUMN RTQUEUEHRS DEC(9,4) "
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  
                  bBadCol = CheckConvErrors()
                  sSql = "sp_bindefault DEFZERO, 'RthdTable.RTQUEUEHRS'"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
               End If
            Else
               GoTo Keys
            End If
         End If
      End With
      'RTMOVEHRS
      sSql = "sp_columns @table_name=RthdTable,@column_name=RTMOVEHRS"
      'Set RdoCol = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
      Set RdoCol = clsADOCon.GetRecordSet(sSql, adOpenForwardOnly)
      With RdoCol
         If Not .BOF And Not .EOF Then
            Err.Clear
            If Not IsNull(.Fields(5)) Then
               If .Fields(5) = "real" Then 'See Else
                  sSql = "ALTER TABLE RthdTable ALTER COLUMN RTMOVEHRS dec(9,4)"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  If Err > 0 Then
                     For Each ER In clsADOCon.ConnectionObject.Errors
                        sconstraint = GetConstraint(ER.Description)
                        If sconstraint <> "" Then Exit For
                     Next ER
                  End If
                  Err.Clear
                  If sconstraint <> "" Then
                     sSql = "ALTER TABLE RthdTable DROP " & sconstraint
                     clsADOCon.ExecuteSql sSql 'rdExecDirect
                  End If
                  sSql = "ALTER TABLE RthdTable ALTER COLUMN RTMOVEHRS DEC(9,4) "
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  
                  bBadCol = CheckConvErrors()
                  sSql = "sp_bindefault DEFZERO, 'RthdTable.RTMOVEHRS'"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
               End If
            End If
         End If
      End With
      'RTSETUPHRS
      sSql = "sp_columns @table_name=RthdTable,@column_name=RTSETUPHRS"
      'Set RdoCol = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
      Set RdoCol = clsADOCon.GetRecordSet(sSql, adOpenForwardOnly)
      With RdoCol
         If Not .BOF And Not .EOF Then
            Err.Clear
            If Not IsNull(.Fields(5)) Then
               If .Fields(5) = "real" Then
                  sSql = "ALTER TABLE RthdTable ALTER COLUMN RTSETUPHRS dec(9,4)"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  If Err > 0 Then
                     For Each ER In clsADOCon.ConnectionObject.Errors
                        sconstraint = GetConstraint(ER.Description)
                        If sconstraint <> "" Then Exit For
                     Next ER
                  End If
                  Err.Clear
                  If sconstraint <> "" Then
                     sSql = "ALTER TABLE RthdTable DROP " & sconstraint
                     clsADOCon.ExecuteSql sSql 'rdExecDirect
                  End If
                  sSql = "ALTER TABLE RthdTable ALTER COLUMN RTSETUPHRS DEC(9,4) "
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  
                  bBadCol = CheckConvErrors()
                  sSql = "sp_bindefault DEFZERO, 'RthdTable.RTSETUPHRS'"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
               End If
            End If
         End If
      End With
      'RTUNITHRS
      sSql = "sp_columns @table_name=RthdTable,@column_name=RTUNITHRS"
      'Set RdoCol = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
      Set RdoCol = clsADOCon.GetRecordSet(sSql, adOpenForwardOnly)
      With RdoCol
         If Not .BOF And Not .EOF Then
            Err.Clear
            If Not IsNull(.Fields(5)) Then
               If .Fields(5) = "real" Then
                  sSql = "ALTER TABLE RthdTable ALTER COLUMN RTUNITHRS dec(9,4)"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  If Err > 0 Then
                     For Each ER In clsADOCon.ConnectionObject.Errors
                        sconstraint = GetConstraint(ER.Description)
                        If sconstraint <> "" Then Exit For
                     Next ER
                  End If
                  Err.Clear
                  If sconstraint <> "" Then
                     sSql = "ALTER TABLE RthdTable DROP " & sconstraint
                     clsADOCon.ExecuteSql sSql 'rdExecDirect
                  End If
                  sSql = "ALTER TABLE RthdTable ALTER COLUMN RTUNITHRS DEC(9,4) "
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  
                  bBadCol = CheckConvErrors()
                  sSql = "sp_bindefault DEFZERO, 'RthdTable.RTUNITHRS'"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
               End If
            End If
         End If
      End With
      'End RthdTable
      Err.Clear
      'Start RtopTable
      'OPSETUP
      sSql = "sp_columns @table_name=RtopTable,@column_name=OPSETUP"
      'Set RdoCol = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
      Set RdoCol = clsADOCon.GetRecordSet(sSql, adOpenForwardOnly)
      With RdoCol
         If Not .BOF And Not .EOF Then
            Err.Clear
            If Not IsNull(.Fields(5)) Then
               If .Fields(5) = "real" Then
                  sSql = "ALTER TABLE RtopTable ALTER COLUMN OPSETUP dec(9,4)"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  If Err > 0 Then
                     For Each ER In clsADOCon.ConnectionObject.Errors
                        sconstraint = GetConstraint(ER.Description)
                        If sconstraint <> "" Then Exit For
                     Next ER
                  End If
                  Err.Clear
                  If sconstraint <> "" Then
                     sSql = "ALTER TABLE RtopTable DROP " & sconstraint
                     clsADOCon.ExecuteSql sSql 'rdExecDirect
                  End If
                  sSql = "ALTER TABLE RtopTable ALTER COLUMN OPSETUP DEC(9,4) "
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  
                  bBadCol = CheckConvErrors()
                  sSql = "sp_bindefault DEFZERO, 'RtopTable.OPSETUP'"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
               End If
            End If
         End If
      End With
      'OPUNIT
      sSql = "sp_columns @table_name=RtopTable,@column_name=OPUNIT"
      'Set RdoCol = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
      Set RdoCol = clsADOCon.GetRecordSet(sSql, adOpenForwardOnly)
      With RdoCol
         If Not .BOF And Not .EOF Then
            Err.Clear
            If Not IsNull(.Fields(5)) Then
               If .Fields(5) = "real" Then
                  sSql = "ALTER TABLE RtopTable ALTER COLUMN OPUNIT dec(9,4)"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  If Err > 0 Then
                     For Each ER In clsADOCon.ConnectionObject.Errors
                        sconstraint = GetConstraint(ER.Description)
                        If sconstraint <> "" Then Exit For
                     Next ER
                  End If
                  Err.Clear
                  If sconstraint <> "" Then
                     sSql = "ALTER TABLE RtopTable DROP " & sconstraint
                     clsADOCon.ExecuteSql sSql 'rdExecDirect
                  End If
                  sSql = "ALTER TABLE RtopTable ALTER COLUMN OPUNIT DEC(9,4) "
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  
                  bBadCol = CheckConvErrors()
                  sSql = "sp_bindefault DEFZERO, 'RtopTable.OPUNIT'"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
               End If
            End If
         End If
      End With
      'OPQHRS
      sSql = "sp_columns @table_name=RtopTable,@column_name=OPQHRS"
      'Set RdoCol = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
      Set RdoCol = clsADOCon.GetRecordSet(sSql, adOpenForwardOnly)
      With RdoCol
         If Not .BOF And Not .EOF Then
            Err.Clear
            If Not IsNull(.Fields(5)) Then
               If .Fields(5) = "real" Then
                  sSql = "ALTER TABLE RtopTable ALTER COLUMN OPQHRS dec(9,4)"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  If Err > 0 Then
                     For Each ER In clsADOCon.ConnectionObject.Errors
                        sconstraint = GetConstraint(ER.Description)
                        If sconstraint <> "" Then Exit For
                     Next ER
                  End If
                  Err.Clear
                  If sconstraint <> "" Then
                     sSql = "ALTER TABLE RtopTable DROP " & sconstraint
                     clsADOCon.ExecuteSql sSql 'rdExecDirect
                  End If
                  sSql = "ALTER TABLE RtopTable ALTER COLUMN OPQHRS DEC(9,4) "
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  
                  bBadCol = CheckConvErrors()
                  sSql = "sp_bindefault DEFZERO, 'RtopTable.OPQHRS'"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
               End If
            End If
         End If
      End With
      'OPMHRS
      sSql = "sp_columns @table_name=RtopTable,@column_name=OPMHRS"
      'Set RdoCol = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
      Set RdoCol = clsADOCon.GetRecordSet(sSql, adOpenForwardOnly)
      With RdoCol
         If Not .BOF And Not .EOF Then
            Err.Clear
            If Not IsNull(.Fields(5)) Then
               If .Fields(5) = "real" Then
                  sSql = "ALTER TABLE RtopTable ALTER COLUMN OPMHRS dec(9,4)"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  If Err > 0 Then
                     For Each ER In clsADOCon.ConnectionObject.Errors
                        sconstraint = GetConstraint(ER.Description)
                        If sconstraint <> "" Then Exit For
                     Next ER
                  End If
                  Err.Clear
                  If sconstraint <> "" Then
                     sSql = "ALTER TABLE RtopTable DROP " & sconstraint
                     clsADOCon.ExecuteSql sSql 'rdExecDirect
                  End If
                  sSql = "ALTER TABLE RtopTable ALTER COLUMN OPMHRS DEC(9,4) "
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  
                  bBadCol = CheckConvErrors()
                  sSql = "sp_bindefault DEFZERO, 'RtopTable.OPMHRS'"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
               End If
            End If
         End If
      End With
      'OPSVCUNIT
      sSql = "sp_columns @table_name=RtopTable,@column_name=OPSVCUNIT"
      'Set RdoCol = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
      Set RdoCol = clsADOCon.GetRecordSet(sSql, adOpenForwardOnly)
      With RdoCol
         If Not .BOF And Not .EOF Then
            Err.Clear
            If Not IsNull(.Fields(5)) Then
               If .Fields(5) = "real" Then
                  sSql = "ALTER TABLE RtopTable ALTER COLUMN OPSVCUNIT dec(12,4)"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  If Err > 0 Then
                     For Each ER In clsADOCon.ConnectionObject.Errors
                        sconstraint = GetConstraint(ER.Description)
                        If sconstraint <> "" Then Exit For
                     Next ER
                  End If
                  Err.Clear
                  If sconstraint <> "" Then
                     sSql = "ALTER TABLE RtopTable DROP " & sconstraint
                     clsADOCon.ExecuteSql sSql 'rdExecDirect
                  End If
                  sSql = "ALTER TABLE RtopTable ALTER COLUMN OPSVCUNIT DEC(12,4) "
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  
                  bBadCol = CheckConvErrors()
                  sSql = "sp_bindefault DEFZERO, 'RtopTable.OPSVCUNIT'"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
               End If
            End If
         End If
      End With
      'End RtopTable
Keys:
      Err.Clear
      'Dump bogus orphans (if any)
      sSql = "DELETE FROM RtopTable " & vbCrLf _
             & "FROM RtopTable LEFT JOIN RthdTable ON RtopTable.OPREF = RthdTable.RTREF " & vbCrLf _
             & "WHERE (RthdTable.RTREF Is Null)"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      
      Err.Clear
      clsADOCon.ADOErrNum = 0
      sSql = "DROP INDEX RthdTable.RteRef"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      If clsADOCon.ADOErrNum = 0 Then
         sSql = "ALTER TABLE RthdTable ALTER COLUMN RTREF CHAR(30) NOT NULL"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
         
         sSql = "ALTER TABLE RthdTable ADD Constraint PK_RthdTable_RTREF PRIMARY KEY CLUSTERED (RTREF) " _
                & "WITH FILLFACTOR=80 "
         clsADOCon.ExecuteSql sSql 'rdExecDirect
         
         sSql = "ALTER TABLE RtopTable ADD CONSTRAINT FK_RtopTable_RthdTable FOREIGN KEY (OPREF) References RthdTable ON DELETE CASCADE ON UPDATE CASCADE "
         clsADOCon.ExecuteSql sSql 'rdExecDirect
      End If
      
   End Sub
   
   Private Sub ConvertRoutingLibrary()
      Dim bBadCol As Byte
      Dim sconstraint As String
      Err.Clear
      'Start RlbrTable
      On Error Resume Next
      'LIBSETUP
      sSql = "sp_columns @table_name=RlbrTable,@column_name=LIBSETUP"
      'Set RdoCol = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
      Set RdoCol = clsADOCon.GetRecordSet(sSql, adOpenForwardOnly)
      With RdoCol
         If Not .BOF And Not .EOF Then
            Err.Clear
            If Not IsNull(.Fields(5)) Then
               If .Fields(5) = "real" Then 'See Else
                  sSql = "ALTER TABLE RlbrTable ALTER COLUMN LIBSETUP dec(9,4)"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  If Err > 0 Then
                     For Each ER In clsADOCon.ConnectionObject.Errors
                        sconstraint = GetConstraint(ER.Description)
                        If sconstraint <> "" Then Exit For
                     Next ER
                  End If
                  Err.Clear
                  If sconstraint <> "" Then
                     sSql = "ALTER TABLE RlbrTable DROP " & sconstraint
                     clsADOCon.ExecuteSql sSql 'rdExecDirect
                  End If
                  sSql = "ALTER TABLE RlbrTable ALTER COLUMN LIBSETUP DEC(9,4) "
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  
                  bBadCol = CheckConvErrors()
                  sSql = "sp_bindefault DEFZERO, 'RlbrTable.LIBSETUP'"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
               Else
                  GoTo Keys
               End If
            End If
         End If
      End With
      'LIBUNIT
      sSql = "sp_columns @table_name=RlbrTable,@column_name=LIBUNIT"
      'Set RdoCol = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
      Set RdoCol = clsADOCon.GetRecordSet(sSql, adOpenForwardOnly)
      With RdoCol
         If Not .BOF And Not .EOF Then
            Err.Clear
            If Not IsNull(.Fields(5)) Then
               If .Fields(5) = "real" Then
                  sSql = "ALTER TABLE RlbrTable ALTER COLUMN LIBUNIT DEC(9,4)"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  If Err > 0 Then
                     For Each ER In clsADOCon.ConnectionObject.Errors
                        sconstraint = GetConstraint(ER.Description)
                        If sconstraint <> "" Then Exit For
                     Next ER
                  End If
                  Err.Clear
                  If sconstraint <> "" Then
                     sSql = "ALTER TABLE RlbrTable DROP " & sconstraint
                     clsADOCon.ExecuteSql sSql 'rdExecDirect
                  End If
                  sSql = "ALTER TABLE RlbrTable ALTER COLUMN LIBUNIT DEC(9,4) "
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  
                  bBadCol = CheckConvErrors()
                  sSql = "sp_bindefault DEFZERO, 'RlbrTable.LIBUNIT'"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
               End If
            End If
         End If
      End With
      'LIBQHRS
      sSql = "sp_columns @table_name=RlbrTable,@column_name=LIBQHRS"
      'Set RdoCol = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
      Set RdoCol = clsADOCon.GetRecordSet(sSql, adOpenForwardOnly)
      With RdoCol
         If Not .BOF And Not .EOF Then
            Err.Clear
            If Not IsNull(.Fields(5)) Then
               If .Fields(5) = "real" Then
                  sSql = "ALTER TABLE RlbrTable ALTER COLUMN LIBQHRS DEC(9,4)"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  If Err > 0 Then
                     For Each ER In clsADOCon.ConnectionObject.Errors
                        sconstraint = GetConstraint(ER.Description)
                        If sconstraint <> "" Then Exit For
                     Next ER
                  End If
                  Err.Clear
                  If sconstraint <> "" Then
                     sSql = "ALTER TABLE RlbrTable DROP " & sconstraint
                     clsADOCon.ExecuteSql sSql 'rdExecDirect
                  End If
                  sSql = "ALTER TABLE RlbrTable ALTER COLUMN LIBQHRS DEC(9,4) "
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  
                  bBadCol = CheckConvErrors()
                  sSql = "sp_bindefault DEFZERO, 'RlbrTable.LIBQHRS'"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
               End If
            End If
         End If
      End With
      'LIBMHRS
      sSql = "sp_columns @table_name=RlbrTable,@column_name=LIBMHRS"
      'Set RdoCol = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
      Set RdoCol = clsADOCon.GetRecordSet(sSql, adOpenForwardOnly)
      With RdoCol
         If Not .BOF And Not .EOF Then
            Err.Clear
            If Not IsNull(.Fields(5)) Then
               If .Fields(5) = "real" Then
                  sSql = "ALTER TABLE RlbrTable ALTER COLUMN LIBMHRS DEC(9,4)"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  If Err > 0 Then
                     For Each ER In clsADOCon.ConnectionObject.Errors
                        sconstraint = GetConstraint(ER.Description)
                        If sconstraint <> "" Then Exit For
                     Next ER
                  End If
                  Err.Clear
                  If sconstraint <> "" Then
                     sSql = "ALTER TABLE RlbrTable DROP " & sconstraint
                     clsADOCon.ExecuteSql sSql 'rdExecDirect
                  End If
                  sSql = "ALTER TABLE RlbrTable ALTER COLUMN LIBMHRS DEC(9,4) "
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
                  
                  bBadCol = CheckConvErrors()
                  sSql = "sp_bindefault DEFZERO, 'RlbrTable.LIBMHRS'"
                  clsADOCon.ExecuteSql sSql 'rdExecDirect
               End If
            End If
         End If
      End With
Keys:
      Err.Clear
      clsADOCon.ADOErrNum = 0
      sSql = "DROP INDEX RlbrTable.LibRef"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      If clsADOCon.ADOErrNum = 0 Then
         sSql = "ALTER TABLE RlbrTable ALTER COLUMN LIBREF CHAR(12) NOT NULL"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
         
         sSql = "ALTER TABLE RlbrTable ADD Constraint PK_RlbrTable_LIBREF PRIMARY KEY CLUSTERED (LIBREF) " _
                & "WITH FILLFACTOR=80 "
         clsADOCon.ExecuteSql sSql 'rdExecDirect
      End If
      
   End Sub
   
   
   
   Public Sub FillDocumentRevisions(frm As Form)
      frm.cmbRev.Clear
      On Error GoTo DiaErr1
      sSql = "select DISTINCT DLSREV as Rev from DlstTable" & vbCrLf _
         & "where DLSREF='" & Compress(frm.cmbPrt) & "'" & vbCrLf '_
         '& "union" & vbCrLf _
         '& "select PADOCLISTREV as Rev from PartTable where PARTREF = '" & Compress(frm.cmbPrt) & "'" & vbCrLf _
         '& "order by Rev"
      LoadComboBox frm.cmbRev, -1
      'If frm.cmbRev.ListCount > 0 Then frm.cmbRev = frm.cmbRev.List(0)
      Exit Sub
      
DiaErr1:
      sProcName = "FillDocumentRev"
      CurrError.Number = Err.Number
      CurrError.Description = Err.Description
      DoModuleErrors frm
      
   End Sub
   
   '10/9/06
   '11/1/06 added On Error Resume Next

   Public Sub GetEstimatingRates(MaterialBurden As Currency, FactoryOverHead As Currency, _
                                 LaborRate As Currency, GenAdminExpense As Currency, ProfitOfSale As Currency, _
                                 ScrapRate As Currency)
      Dim RdoPar As ADODB.Recordset
      
      'Not all forms may comply with the TextBox names
      On Error Resume Next
      sSql = "Qry_EstParameters"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPar, ES_FORWARD)
      If bSqlRows Then
         With RdoPar
            MaterialBurden = Format(!EstMatlBurden / 100, ES_QuantityDataFormat)
            FactoryOverHead = Format(!EstFactoryOverHead / 100, ES_QuantityDataFormat)
            LaborRate = Format(!EstLaborRate, ES_QuantityDataFormat)
            GenAdminExpense = Format(!EstGenAdmnExp / 100, ES_QuantityDataFormat)
            ProfitOfSale = Format(!EstProfitOfSale / 100, ES_QuantityDataFormat)
            ScrapRate = Format(!EstScrapRate / 100, ES_QuantityDataFormat)
            MDISect.ActiveForm.txtGna = Format(!EstGenAdmnExp, ES_QuantityDataFormat)
            MDISect.ActiveForm.txtPrf = Format(ProfitOfSale * 100, ES_QuantityDataFormat)
            MDISect.ActiveForm.txtScr = Format(!EstScrapRate, ES_QuantityDataFormat)
            MDISect.ActiveForm.txtRte = Format(LaborRate, ES_QuantityDataFormat)
            ClearResultSet RdoPar
         End With
      Else
         MDISect.ActiveForm.txtRte = "0.000"
      End If
      Set RdoPar = Nothing
      Exit Sub
      
modErr1:
      sProcName = "getrates"
      CurrError.Number = Err.Number
      CurrError.Description = Err.Description
      DoModuleErrors MDISect.ActiveForm
      
   End Sub

Public Sub TotalBidMatl(EstNo As Long, EstPart As String, EstQty As Currency)
   'EstQty = the total quantity for the estimate
   
   Dim rdoMat As ADODB.Recordset
   Dim qtyAtThisLevel As Currency
   Dim EstqtyTopLevel As Currency
   'Dim iCounter As Integer
   
   'iCounter = 0
   cBidBurden = 0
   cBidMaterial = 0
   'cBidTotMat = 0
   On Error GoTo DiaErr1
   sSql = "SELECT BIDBOMREF,BIDBOMASSYPART,BIDBOMPARTREF,BIDBOMQTYREQD," & vbCrLf _
      & "BIDBOMCONVERSION,BIDBOMSETUP,BIDBOMADDER,BIDBOMMATERIAL," & vbCrLf _
      & "BIDBOMMATERIALBRD,BIDBOMESTUNITCOST" & vbCrLf _
      & "FROM EsbmTable" & vbCrLf _
      & "WHERE BIDBOMREF = " & EstNo & vbCrLf _
      & "AND BIDBOMASSYPART = '" & Compress(EstPart) & "'" & vbCrLf _
      & "AND BIDBOMLEVEL = 1"
   If clsADOCon.GetDataSet(sSql, rdoMat, ES_FORWARD) Then
      With rdoMat
         Do Until .EOF
            sProcName = "TotalBidMatl"
            'cBidQuantity = 1
            'iCounter = iCounter + 1
             
            'calculate quantity at first level of explosion
            'qtyAtThisLevel = IIf(!BIDBOMCONVERSION = 0, !BIDBOMQTYREQD, !BIDBOMQTYREQD / !BIDBOMCONVERSION) - EVALUATES BOTH = DIVIDE BY ZERO
            If !BIDBOMCONVERSION = 0 Then
                ' MM Add the Setup and adder quantity for calculating materail cost
                'qtyAtThisLevel = (!BIDBOMQTYREQD + !BIDBOMSETUP + !BIDBOMADDER)
                'MM qtyAtThisLevel = !BIDBOMQTYREQD
                ' MM 7/23/2009 Set has to be divided by the qty from parent level
                ' Set up issue
                If (CInt(EstQty) <> 0) Then
                    qtyAtThisLevel = (!BIDBOMQTYREQD) + (!BIDBOMSETUP + !BIDBOMADDER) / EstQty
                Else
                    qtyAtThisLevel = (!BIDBOMQTYREQD) + (!BIDBOMSETUP + !BIDBOMADDER)
                End If
            Else
            
                If (CInt(EstQty) <> 0) Then
                    qtyAtThisLevel = (!BIDBOMQTYREQD + ((!BIDBOMSETUP + !BIDBOMADDER) / EstQty)) / !BIDBOMCONVERSION
                Else
                    qtyAtThisLevel = (!BIDBOMQTYREQD + !BIDBOMSETUP + !BIDBOMADDER) / !BIDBOMCONVERSION
                End If
            End If
            
            'unit material without burden
            cBidMaterial = cBidMaterial + !BIDBOMMATERIAL * qtyAtThisLevel
            
            'calculate burden
            cBidBurden = cBidBurden + !BIDBOMMATERIAL * !BIDBOMMATERIALBRD * qtyAtThisLevel / 100
            
            'explode next level (2)
            EstqtyTopLevel = EstQty
            TotalBidNextMatl EstNo, Trim(!BIDBOMPARTREF), qtyAtThisLevel, 2, EstqtyTopLevel
            .MoveNext
         Loop
         ClearResultSet rdoMat
      End With
   End If
   Set rdoMat = Nothing
   EstiESe02a.lblMaterial = Format(cBidMaterial, ES_MoneyFormat)     'unit material
   EstiESe02a.lblBurden = Format(cBidBurden, ES_MoneyFormat)         'unit burden
   EstiESe02a.lblTotMat = Format(cBidMaterial + cBidBurden, ES_MoneyFormat)   'unit burdened matrial
   EstiESe02a.lblEstTotMatl = Format((cBidMaterial + cBidBurden) * EstQty, ES_MoneyFormat) 'grand total material
   Exit Sub
   
DiaErr1:
whoops:
   sProcName = "TotalBidMatl"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Dim frm As New ClassErrorForm
   DoModuleErrors frm
End Sub

Private Sub TotalBidNextMatl(EstNo As Long, AssemblyPart As String, _
   QtyAtPriorLevel As Currency, BomLevel As Integer, EstqtyTopLevel As Currency)
   
   'BomLevel should = 2 or greater
   
   If BomLevel > 10 Then
      MsgBox "WARNING: More than 10 explosion levels." & vbCrLf _
      & "A part in the suspected loop is " & Trim(AssemblyPart) & "." & vbCrLf _
         & "Additional levels ignored."
      Exit Sub
   End If
   
   Dim RdoNextMat As ADODB.Recordset
   Dim cBurden As Currency
   Dim cConvert As Currency
   Dim qtyAtThisLevel As Currency
   sSql = "SELECT BIDBOMREF,BIDBOMASSYPART,BIDBOMPARTREF,BIDBOMQTYREQD," & vbCrLf _
          & "BIDBOMCONVERSION,BIDBOMSETUP,BIDBOMADDER,BIDBOMMATERIAL," & vbCrLf _
          & "BIDBOMMATERIALBRD,BIDBOMESTUNITCOST" & vbCrLf _
          & "FROM EsbmTable" & vbCrLf _
          & "WHERE BIDBOMREF=" & EstNo & vbCrLf _
          & "AND BIDBOMASSYPART='" & AssemblyPart & "'" & vbCrLf _
          & "AND BIDBOMLEVEL = " & BomLevel
   If clsADOCon.GetDataSet(sSql, RdoNextMat, ES_FORWARD) Then
      With RdoNextMat
         Do Until .EOF
            sProcName = "TotalBidNextMatl"
            'iCounter = iCounter + 1
'            cConvert = Format(!BIDBOMCONVERSION, ES_QuantityDataFormat)
'            If cConvert = 0 Then cConvert = 1
'            cBurden = Format(!BIDBOMMATERIALBRD, ES_QuantityDataFormat)
'            If cBurden > 0 Then cBurden = cBurden / 100
'            cQuantity = cQuantity / cConvert
'            cQuantity = Format((!BIDBOMQTYREQD * cBidQuantity), ES_QuantityDataFormat)
'            cBurden = (cBurden * (!BIDBOMMATERIAL * !BIDBOMQTYREQD))
'            cBidBurden = cBidBurden + cBurden
'            cBidMaterial = cBidMaterial + ((!BIDBOMESTUNITCOST - cBurden) * cBidQuantity)
'            cBidTotMat = cBidTotMat + (!BIDBOMESTUNITCOST * cBidQuantity)
'            cBidQuantity = Format(!BIDBOMQTYREQD, ES_QuantityDataFormat)
            
            'calculate quantity at this level of explosion
'            qtyAtThisLevel = QtyAtPriorLevel * _
'               IIf(!BIDBOMCONVERSION = 0, !BIDBOMQTYREQD, !BIDBOMQTYREQD / !BIDBOMCONVERSION)
            'qtyAtThisLevel = IIf(!BIDBOMCONVERSION = 0, !BIDBOMQTYREQD, !BIDBOMQTYREQD / !BIDBOMCONVERSION) - EVALUATES BOTH = DIVIDE BY ZERO
            ' MM 7/23/2009 Set has to be divided by the qty from parent level
            If !BIDBOMCONVERSION = 0 Then
                If (CInt(EstqtyTopLevel) <> 0) Then
                    qtyAtThisLevel = QtyAtPriorLevel * (!BIDBOMQTYREQD + ((!BIDBOMSETUP + !BIDBOMADDER) / EstqtyTopLevel))
                Else
                    qtyAtThisLevel = QtyAtPriorLevel * (!BIDBOMQTYREQD + !BIDBOMSETUP + !BIDBOMADDER)
                End If
                
            Else
                If (CInt(EstqtyTopLevel) <> 0) Then
                    qtyAtThisLevel = QtyAtPriorLevel * (!BIDBOMQTYREQD + ((!BIDBOMSETUP + !BIDBOMADDER) / EstqtyTopLevel)) / !BIDBOMCONVERSION
                Else
                    qtyAtThisLevel = QtyAtPriorLevel * (!BIDBOMQTYREQD + !BIDBOMSETUP + !BIDBOMADDER) / !BIDBOMCONVERSION
                End If
            End If
            
            'unit material without burden
            cBidMaterial = cBidMaterial + !BIDBOMMATERIAL * qtyAtThisLevel
            
            'calculate burden
            cBidBurden = cBidBurden + !BIDBOMMATERIAL * !BIDBOMMATERIALBRD * qtyAtThisLevel / 100
            
            'explode next level
            TotalBidNextMatl EstNo, Trim(!BIDBOMPARTREF), qtyAtThisLevel, BomLevel + 1, EstqtyTopLevel
            .MoveNext
         Loop
         ClearResultSet RdoNextMat
      End With
   End If
   Set RdoNextMat = Nothing
   
End Sub



Public Function CurrentPartType(ByVal sPartNo) As Integer
    Dim rdoPart As ADODB.Recordset
    
   CurrentPartType = 0
   On Error Resume Next
   sSql = "SELECT PALEVEL FROM PartTable WHERE PartREF='" & Compress(sPartNo) & "'"
   If clsADOCon.GetDataSet(sSql, rdoPart, ES_FORWARD) Then
        If Not rdoPart.EOF Then CurrentPartType = Val("" & rdoPart!PALEVEL)
   End If
   ClearResultSet rdoPart
   Set rdoPart = Nothing
End Function
