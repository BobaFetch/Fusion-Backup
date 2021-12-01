Attribute VB_Name = "ESIFINA"

'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'*************************************************************************************
' ESIFINA - ES/2000 Finance Module Common Source.

' Common text
Public Const TTMARK = "Mark All"
Public Const TTUNMARK = "Unmark All"
Public Const TTPRINTER = "Show System Printers"
Public Const TTDEFAULT = "Default Printer"
Public Const TTSAVEPRN = "_Printer"

' JET 3.5
Public JetWkSpace As Workspace
Public JetDb As DAO.Database

Public Y As Byte
Public lCurrInvoice As Long
Public sCurrForm As String
Public sPassedPart As String
Public sSelected As String

Public bAccount As Byte

Public bFoundPart As Byte

Dim mzBuff As String

Public sFavorites(13) As String
Public sRecent(10) As String
Public sSession(50) As String
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

' Format Mask
Global Const CURRENCYMASK = "#,###,###,###,##0.00"
Global Const DATEMASK = "mm/dd/yy"

'old help stuff for this module only
Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Any) As Long
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

'7/25/07 TEL kludge to make menus (tab...frm) work with new common code
Public sActiveTab(8) As Integer


'11/14/02 to retrieve the next PKRECORD (index piece)

Public Function GetNextPickRecord(sMoPartRef As String, lRunno As Long) As Integer
   Dim RdoRec As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT MAX(PKRECORD) FROM MopkTable WHERE " _
          & "PKMOPART='" & sMoPartRef & "' AND " _
          & "PKMORUN=" & lRunno & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRec, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoRec.Fields(0)) Then
         GetNextPickRecord = RdoRec.Fields(0) + 1
      Else
         GetNextPickRecord = 1
      End If
   Else
      GetNextPickRecord = 1
   End If
   Set RdoRec = Nothing
   Exit Function
   
modErr1:
Resume modErr2:
modErr2:
   GetNextPickRecord = 1
   On Error GoTo 0
   
End Function

Public Sub FillBuyers()
   Dim RdoByr As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "Qry_GetBuyerList"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoByr, ES_FORWARD)
   If bSqlRows Then
      With RdoByr
         Do Until .EOF
            AddComboStr MdiSect.ActiveForm.cmbByr.hWnd, "" & Trim(!BYNUMBER)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoByr = Nothing
   Exit Sub
   
modErr1:
   sProcName = "fillbuyers"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Sub

'*************************************************************************************

Public Sub FillRuns(frm As Form, sSearchString As String, _
                    Optional sComboName As String)
   Dim RdoFrn As ADODB.Recordset
   On Error GoTo modErr1
   If sSearchString = "<> 'CA'" Then
      sSql = "Qry_RunsNotCanceled"
   Else
      sSql = "Qry_RunsNotLikeC"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoFrn, ES_FORWARD)
   If bSqlRows Then
      With RdoFrn
         On Error Resume Next
         If frm.ActiveControl.Name = "cmbPrt" Or sComboName = "cmbPrt" Then
            frm.cmbPrt = "" & Trim(!PARTNUM)
            Do Until .EOF
               AddComboStr frm.cmbPrt.hWnd, "" & Trim(!PARTNUM)
               .MoveNext
            Loop
         Else
            Do Until .EOF
               'frm.cmbMon.AddItem "" & Trim(!PARTNUM)
               AddComboStr frm.cmbMon.hWnd, "" & Trim(!PARTNUM)
               .MoveNext
            Loop
         End If
         .Cancel
      End With
   End If
   On Error Resume Next
   Set RdoFrn = Nothing
   Exit Sub
   
modErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume modErr2
modErr2:
   DoModuleErrors frm
   
End Sub

Public Function FindVendor( _
                           frm As Form, _
                           Optional bName As Boolean, _
                           Optional ByRef iNetDays As Integer, _
                           Optional bRemit As Byte) As Byte
   
   Dim RdoVnd As ADODB.Recordset
   Dim sVendRef As String
   
   sVendRef = Compress(frm.cmbVnd)
   If Len(sVendRef) = 0 Then Exit Function
   
   If bRemit Then
      sSql = "SELECT VEREF,VENICKNAME,VEBNAME,VENETDAYS,VEBADR,VEBCITY," _
             & "VEBSTATE,VEBZIP,VEBCOUNTRY,VECNAME,VECADR,VECCITY,VECSTATE," _
             & "VECZIP,VECCOUNTRY FROM VndrTable WHERE VEREF='" & sVendRef & "'"
   Else
      sSql = "SELECT VEREF,VENICKNAME,VEBNAME,VENETDAYS FROM " _
             & "VndrTable WHERE VEREF='" & sVendRef & "'"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVnd)
   On Error Resume Next
   If bSqlRows Then
      With RdoVnd
         frm.lblNme.ForeColor = frm.ForeColor
         frm.cmbVnd = "" & Trim(!VENICKNAME)
         If bRemit Then
            If Trim(!VECADR) <> "" Then
               frm.lblNme = Trim(!VECNAME) & vbCrLf _
                            & Trim(!VECADR) & vbCrLf _
                            & Trim(!VECCITY) & ", " _
                            & Trim(!VECSTATE) & " " _
                            & Trim(!VECZIP) & vbCrLf _
                            & Trim(!VECCOUNTRY)
            Else
               frm.lblNme = Trim(!VEBNAME) & vbCrLf _
                            & Trim(!VEBADR) & vbCrLf _
                            & Trim(!VEBCITY) & ", " _
                            & Trim(!VEBSTATE) & " " _
                            & Trim(!VEBZIP) & vbCrLf _
                            & Trim(!VEBCOUNTRY)
            End If
         Else
            frm.lblNme = "" & Trim(!VEBNAME)
         End If
         iNetDays = !VENETDAYS
         FindVendor = 1
         .Cancel
      End With
      Set RdoVnd = Nothing
   Else
      With frm
         .cmbVnd = ""
         If bName Then
            .lblNme.ForeColor = ES_RED
            .txtNme = "*** No Vendors Found ***"
            .lblNme = "*** No Vendors Found ***"
         Else
            .txtNme = ""
            .lblNme = ""
         End If
      End With
   End If
   Exit Function
End Function

Public Sub FormUnload(Optional bDontShowForm As Byte)
   On Error Resume Next
   MdiSect.BotPanel = MdiSect.Caption
   'RdoRes.Close
   If bDontShowForm = 0 Then
      Select Case cUR.CurrentGroup
         Case "Apay"
            tabApay.Show
         Case "Arec"
            tabArec.Show
         Case "Cred"
            tabClos.Show
         Case "Jorn"
            tabJorn.Show
         Case "Genl"
            tabGenl.Show
         Case "Scst"
            tabScst.Show
         Case "Jcst"
            tabJcst.Show
         Case "Lcst"
            tabLcst.Show
         Case "Slan"
            tabSalesAnalysis.Show
      End Select
      cUR.CurrentGroup = ""
   End If
End Sub

Public Sub OpenFavorite(sSelected)
   CloseForms
   If LTrim$(sSelected) = "" Then Exit Sub
   MouseCursor 13
   On Error GoTo OpenFavErr1
   Select Case Trim(sSelected)
      Case "Overhead Applied (Report)"
         diaCLp05a.Show
      Case "Material Purchase Price Variance (Report)"
         diaCLp01a.Show
      Case "B & O Tax Liability (Report)"
         diaARp15a.Show
      Case "Advance Payment Status (Report)"
         diaARp14a.Show
      Case "Tax Codes (Report)"
         diaARp13a.Show
      Case "Sales Tax Liabilty"
         diaARp11a.Show
      Case "View A Cash Receipt (Report)"
         diaARp09a.Show
      Case "Unprinted Invoices (Report)"
         diaARp08a.Show
      Case "Assign Customer Payers"
         diaARe12a.Show
      Case "Assign Customer Payers"
         diaARe12a.Show
      Case "Assign Tax Codes To Parts"
         diaARe11a.Show
      Case "Add Revise Tax Codes"
         diaARe10a.Show
      Case "Credit or Debit Memo Against Prior Invoice"
         diaARe06a.Show
      Case "Check Analysis (Report)"
         diaAPp19a.Show
      Case "Check Analysis (Report)"
         diaAPp19a.Show
      Case "Cleared Check Summary (Report)"
         diaAPp16a.Show
      Case "Received And Not Invoiced (Report)"
         diaAPp11a.Show
      Case "Purchases By GL Account (Report)"
         diaAPp10a.Show
      Case "Vendor Statements (Report)"
         diaAPp02a.Show
      Case "Vendor Invoice Register (Report)"
         diaAPp07a.Show
      Case "Average Age of Paid Invoices"
         diaAPp06a.Show
      Case "Material Movement To Project MO (Report)"
         diaCLp06a.Show
      Case "Sales By GL Account (Report)"
         diaARp07a.Show
      Case "Vendor Invoice"
         diaAPe01a.Show
      Case "Vendors"
        VendorEdit01.Tag = 2 'Set the tag for the calling procedure
        VendorEdit01.Show 'Show the new vendor form
      Case "Vendor Invoice (Report)"
         diaAPp01a.Show
      Case "Revise Invoice Due Dates/Comments"
         diaAPe04a.Show
      Case "Vendor Invoice Register"
         diaAPp07a.Show
      Case "Accounts Payable Aging (Report)"
         diaAPp08a.Show
      Case "Received And Not Invoiced"
         diaAPp11a.Show
      Case "Cancel An AP Invoice"
         diaAPf01a.Show
      Case "Change Invoice GL Distribution"
         diaAPe05a.Show
      Case "Cash Disbursements"
         diaAPe03a.Show
      Case "Customers"
         diaCcust.Show
      Case "Customer Invoice (Sales Order)"
         diaARe01a.Show
      Case "Customer Invoice (Packing Slip)"
         diaARe02a.Show
      Case "Customer Invoices (Report)"
         diaARp01a.Show
      Case "Customer Debit Or Credit Memo"
         diaARe03a.Show
      Case "Customer Statements"
         diaARp02a.Show
      Case "Customer Invoice Register"
         diaARp03a.Show
      Case "Customer Invoice Comments"
         diaARe05a.Show
      Case "Cash Receipts"
         diaARe04a.Show
      Case "Cancel A Cash Receipt"
         diaARf02a.Show
      Case "Cash Receipts Register (Report)"
         diaARp04a.Show
      Case "Cancel An Invoice"
         diaARf01a.Show
      Case "Accounts Receivable Aging (Report)"
         diaARp05a.Show
      Case "Financial Statement Structure"
         diaGLe03a.Show
      Case "Chart Of Accounts"
         diaGLe01a.Show
      Case "Chart Of Accounts (Report)"
         diaGLp01a.Show
      Case "Account Numbers For Parts"
         diaGLe07a.Show
      Case "Divisions"
         diaCdivs.Show
      Case "Accounts By Part Number"
         diaGLp11a.Show
      Case "Product Codes"
         diaPcode.Show
      Case "Close Journals"
         diaJRf02a.Show
      Case "Manufacturing Order Budgets"
         diaJcbud.Show
      Case "Charge Material To A Project"
         diaMproj.Show
         'Case "Sales Order Allocations"
         '    diaJcsoa.Show
      Case "Close A Manufacturing Order"
         diaScncl.Show
      Case "Manufacturing Order Cost Analysis"
         diaJCp01a.Show
         ' Case "Job Cost Summary Or Detail"
         '     diaPjc02.Show
      Case "Fiscal Years"
         diaGLe04a.Show
      Case "Journal Entry"
         diaGLe02a.Show
      Case "Journals (Report)"
         diaJRp01a.Show
      Case "Manufacturing Orders By Date"
         diaPsh03.Show
      Case "Manufacturing Orders By Part"
         diaPsh04.Show
      'Case "Manufacturing Orders"
      '   diaPsh01.Show
      Case "Standard Cost"
         diaIsstd.Show
      Case "Cost Information (Report)"
         diaSCp01a.Show
      Case "Change Invoice Customer"
         'tempbox.Show
      Case "Vendor Credit Or Debit Memo"
         diaAPe02a.Show
      Case "Void AP Check"
         diaAPf04a.Show
      Case "Edit Check Memos"
         diaAPe10a.Show
      Case "Computer Check Setup"
         diaAPe08a.Show
      Case "Check Setup (Report)"
         diaAPp14a.Show
      Case "Computer Check Summary"
         diaAPp15a.Show
      Case "Customer Statements (Report)"
         diaARp02a.Show
      Case "Sales Order Invoice"
         diaARe01a.Show
      Case "Work In Process (Report)"
         diaWip.Show
      Case "Update Sales Activity Standard Cost"
         diaCLf01a.Show
      Case "Update Sales Order Account Distributions"
         diaARf01a.Show
      Case "Cash Account Reconciliation (Report)"
         diaGLp15a.Show
      Case "General Journals(Report"
         diaGLp03a.Show
      Case "Opened Closed Journals"
         diaJRf05a.Show
      Case "Cash Account Reconciliation"
         diaGLe10a.Show
      Case "External Check (No Invoice)"
         diaAPe11a.Show
      Case "Proposed Vs. Current Standard Cost (Report)"
         diaSCp07a.Show
      Case "Proposed Vs. Current Standard Cost (Report)"
         diaSCp04a.Show
      Case "Cost Detail By Part (Report)"
         diaSCp03a.Show
      Case "Exploded Proposed Standard Cost Analysis (Report)"
         diaSCp02a.Show
      Case "Cost Information"
         diaSCe02a.Show
      Case "Raw Material/Finished Goods Inventory"
         diaRMFG.Show
      Case "Divisions (Report)"
         diaPco01.Show
      Case "Manufacturing Order Cost Analysis (Report)"
         diaJCp01a.Show
      Case "Cash Account Reconciliation (Report)"
         diaGLp15a.Show
         'Case "Cash Balance (Report)"
         '    diaGLp14a.Show
      Case "Income/Expense Comparison (Report)"
         diaGLp13a.Show
      Case "Pro Forma Income Statement"
         diaGLp12a.Show
      Case "View Budgets"
         diaGLp09a.Show
      Case "Balance Sheet (Report)"
         diaGLp08a.Show
   '      Case "Income Statment With Percentages (Report)"
   '         diaGLp07a.Show
      Case "Income Statement (Report)"
         diaGLp06a.Show
      Case "Detailed General Ledger (Report)"
         diaGLp04a.Show
      Case "General Journals (Report)"
         diaGLp03a.Show
      Case "Forecast report"
         diaForcastf01.Show
      Case "Journal Entry List (Report)"
         diaGLp02a.Show
      Case "Import Cash Receipt from Excel Sheet"
         diaARf09a.Show
      Case "Generate Invoices EDI file"
         PackPSf09a.Show
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

Public Sub FillAllRuns(Contrl As Control)
   Dim RdoRns As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,RUNREF FROM " _
          & "PartTable,RunsTable Where PARTREF=RUNREF ORDER BY PARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRns, ES_FORWARD)
   If bSqlRows Then
      With RdoRns
         Do Until .EOF
            AddComboStr Contrl.hWnd, "" & Trim(!PARTNUM)
            .MoveNext
         Loop
         .Cancel
      End With
      If Contrl.ListCount > 0 Then Contrl = Contrl.List(0)
   End If
   On Error Resume Next
   Set RdoRns = Nothing
   Exit Sub
   
modErr1:
   sProcName = "fillallruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Sub

Public Function SetRecent(frm As Form) As Integer
   Dim A As Integer
   Dim i As Integer
   Static iListcount As Integer
   Dim sTemp As String
   
   On Error GoTo modErr1
   If iListcount < 50 Then
      For i = iListcount To 0 Step -1
         sSession(i% + 1) = sSession(i)
      Next
      iListcount = iListcount + 1
      sSession(0) = frm.Caption
   End If
   
   Erase sRecent
   A = 0
   For i = 1 To 5
      sTemp = MdiSect.ActiveBar1.Bands("mnuFile").Tools("FileRecent" & Trim(str(i - 1))).Caption
      If sTemp = frm.Caption Then sTemp = ""
      If Len(Trim(sTemp)) < 3 Then sTemp = ""
      If sTemp <> "" Then
         A = A + 1
         sRecent(A) = sTemp
      End If
   Next
   If A > 4 Then A = 4
   sRecent(0) = frm.Caption
   For i = 0 To 4
      MdiSect.ActiveBar1.Bands("mnuFile").Tools("FileRecent" & Trim(str(i))).Visible = False
      MdiSect.ActiveBar1.Bands("mnuFile").Tools("FileRecent" & Trim(str(i))).Caption = Trim(str(i))
   Next
   For i = 0 To A
      MdiSect.ActiveBar1.Bands("mnuFile").Tools("FileRecent" & Trim(str(i))).Visible = True
      MdiSect.ActiveBar1.Bands("mnuFile").Tools("FileRecent" & Trim(str(i))).Caption = sRecent(i)
   Next
   Exit Function
   
modErr1:
   Resume modErr2
modErr2:
   'Ignore them
   On Error GoTo 0
   
End Function


Sub SelectHelpTopic(frm As Form, HelpTopic As String)
   Dim l&
   'change to esisale.hlp
   l& = WinHelp(frm.hWnd, sReportPath & "EsiFina.hlp", HELP_KEY, HelpTopic)
   
End Sub

Sub GetFormPos()
   
End Sub

Sub Main()
   Dim sAppTitle As String
   If App.PrevInstance Then
      On Error Resume Next
      sAppTitle = App.Title
      App.Title = "Esifina"
      SysMsg "Select Finance From The Task Bar.", True
      AppActivate sAppTitle
      End
      Exit Sub
   End If

   ' Set the Module name before loading the form
   sProgName = "Finance"
   MainLoad "fina"
   GetFavorites "EsiFina"
   ' save the setting in registry for the module
   SetRegistryAppTitle ("EsiFina")
   ' MM 9/10/2009
   'sProgName = "Finance"
   OpenJet
   MdiSect.Show

End Sub


''Pick up permissions for this user
'Public Sub GetSectionPermissions()
'    Dim RdoUsr As ADODB.RecordSet
'    On Error GoTo ModErr1
'    'Cur.CurrentUser = "ESI"
'    sSql = "SELECT USERREF,USERADDUSER,USERLEVEL," _
'        & "USERFINAGR1,USERFINAGR2,USERFINAGR3," _
'        & "USERFINAGR4,USERFINAGR5,USERFINAGR6 " _
'        & "FROM UsscTable WHERE USERREF='" _
'        & UCase$(cur.CurrentUser) & "'"
'    bSqlRows = clsAdoCon.GetDataSet(sSql,RdoUsr)
'        If bSqlRows Then
'            With RdoUsr
'                User.Adduser = !UserAddUser
'                User.Level = !UserLevel
'                User.Group1 = !USERFINAGR1
'                User.Group2 = !USERFINAGR2
'                User.Group3 = !USERFINAGR3
'                User.Group4 = !USERFINAGR4
'                User.Group5 = !USERFINAGR5
'                User.Group6 = !USERFINAGR6
'                User.Group7 = 1
'                If User.Group4 = 1 Then
'                    User.Group5 = 1
'                    User.Group6 = 1
'                    User.Group8 = 1
'                Else
'                    User.Group5 = 0
'                    User.Group6 = 0
'                    User.Group8 = 0
'                End If
'                .Cancel
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

Public Sub FindPart(frm As Form, Optional sGetPart As String)
   Dim RdoPrt As ADODB.Recordset
   If sGetPart = "" Then
      sGetPart = Compress(frm.cmbPrt)
   Else
      sGetPart = Compress(sGetPart)
   End If
   On Error Resume Next
   If Len(sGetPart) > 0 Then
      sSql = "SELECT PARTREF,PARTNUM,PADESC FROM PartTable " _
             & "WHERE PARTREF='" & sGetPart & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
      If bSqlRows Then
         With RdoPrt
            frm.cmbPrt = "" & Trim(!PARTNUM)
            frm.lblDsc.ForeColor = frm.ForeColor
            frm.lblDsc = "" & Trim(!PADESC)
         End With
      Else
         frm.lblDsc.ForeColor = ES_RED
         frm.cmbPrt = "NONE"
         frm.lblDsc = "*** Part Number Wasn't Found ***"
         
      End If
   Else
      frm.cmbPrt = "NONE"
   End If
   Set RdoPrt = Nothing
End Sub

Public Sub FillParts(frm As Form)
   Dim RdoCmb As ADODB.Recordset
   sSql = "Qry_FillSortedParts"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb)
   If bSqlRows Then
      With RdoCmb
         frm.cmbPrt = "" & Trim(!PARTNUM)
         Do Until .EOF
            'frm.cmbPrt.AddItem "" & Trim(!PARTNUM)
            AddComboStr frm.cmbPrt.hWnd, "" & Trim(!PARTNUM)
            .MoveNext
         Loop
      End With
   End If
   On Error Resume Next
   Set RdoCmb = Nothing
End Sub

'Create Tables, etc, here

Public Sub UpdateTables()
   Dim i As Integer
   Dim RdoNew As ADODB.Recordset
   Dim bByte As Byte
   Dim sCaption As String
   
   Dim sRelease As String
   Dim sTest As String
   Dim rdotest As ADODB.Recordset
   
   MouseCursor 13
   sCaption = MdiSect.Caption
   MdiSect.lblBotPanel.Caption = "Updating Tables."
   
   
   bByte = 1
   
   ' Want to show every time
   If bByte = 0 Then
      diaSql.Height = 1000
      diaSql.Show
      diaSql.Prg1.Visible = False
      diaSql.pnl = "Opening Page."
      diaSql.pnl.Refresh
   Else
      diaSql.Show
      diaSql.Prg1.Visible = True
      diaSql.pnl = "Configuring Settings."
      diaSql.pnl.Refresh
      diaSql.Prg1.Value = 100
   End If
   
   On Error Resume Next
   
   ' 11/09/04 Support Advance payments
   sSql = "ALTER TABLE CihdTable ADD INVORIGIN CHAR(12) NULL DEFAULT('')"
   clsADOCon.ExecuteSql sSql
   
'   ' 09/08/04 WIP reporting table
'   sSql = "CREATE TABLE EsReportWIP (" _
'          & "WIPRUNREF [char] (30) DEFAULT('') NULL," _
'          & "WIPRUNNO [smallint] DEFAULT(0) NULL," _
'          & "WIPLABOR [real] DEFAULT(0) NULL," _
'          & "WIPMATL [real] DEFAULT(0) NULL," _
'          & "WIPOH [real] DEFAULT(0) NULL," _
'          & "WIPEXP [real] DEFAULT(0) NULL," _
'          & "WIPMISSMATL [TINYINT] DEFAULT(0) NULL," _
'          & "WIPMISSTIME [TINYINT] DEFAULT(0) NULL," _
'          & "WIPMISSEXP [TINYINT] DEFAULT(0) NULL," _
'          & "WIPUNCOSTED [TINYINT] DEFAULT(0) NULL) "
'   clsAdoCon.ExecuteSQL sSql
'   sSql = "CREATE UNIQUE CLUSTERED INDEX [EsReportWIP_Unique] " _
'          & "ON EsReportWIP([WIPRUNREF],[WIPRUNNO]) WITH  FILLFACTOR = 80"
'   clsAdoCon.ExecuteSQL sSql
'
   ' Remember last check nuumber for cash account
   sSql = "ALTER TABLE GlacTable ADD GLLASTCHK CHAR (12) NULL DEFAULT('')"
   clsADOCon.ExecuteSql sSql
   
   ' Added assign other customer payers (new to ES/2000).
   sSql = "CREATE TABLE dbo.CpayTable (" _
          & "CPCUST [char] (10) DEFAULT('') NULL," _
          & "CPPAYER [char] (10) DEFAULT('') NULL)"
   clsADOCon.ExecuteSql sSql
   sSql = "CREATE UNIQUE CLUSTERED INDEX [CpayTable_Unique] " _
          & "ON [dbo].[CpayTable]([CPCUST],[CPPAYER]) WITH  FILLFACTOR = 80"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE GlacTable ADD GLRECDATE SMALLDATETIME NULL DEFAULT('')"
   clsADOCon.ExecuteSql sSql
   sSql = "ALTER TABLE GlacTable ADD GLRECBY CHAR (3) NULL DEFAULT('')"
   clsADOCon.ExecuteSql sSql
   sSql = "ALTER TABLE GlacTable ADD GLRECBAL MONEY NULL DEFAULT(0)"
   clsADOCon.ExecuteSql sSql
   sSql = "UPDATE GlacTable SET GLRECBAL = 0 WHERE GLRECBAL IS NULL"
   clsADOCon.ExecuteSql sSql
   
   ' Added save account reconcilation table (new to ES/2000).
   sSql = "CREATE TABLE dbo.ArecTable (" _
          & "RECITEM [char] (12) DEFAULT('') NULL," _
          & "RECITEMTYPE [tinyint] DEFAULT(0) NULL," _
          & "RECACCOUNT [char] (12) DEFAULT('') NULL," _
          & "RECDATE [smalldatetime] DEFAULT(GETDATE()) NULL," _
          & "RECBY [char] (3) DEFAULT('') NULL," _
          & "RECCUST [char] (10) DEFAULT('') NULL," _
          & "RECCLEARED [smalldatetime] NULL," _
          & "RECTRAN [int] DEFAULT(0) NULL," _
          & "RECREF [int] DEFAULT(0) NULL)"
   clsADOCon.ExecuteSql sSql
   
   sSql = "CREATE UNIQUE CLUSTERED INDEX [ArecTable_Unique] " _
          & "ON [dbo].[ArecTable]([RECITEM],[RECITEMTYPE],[RECACCOUNT]," _
          & "[RECCUST],[RECTRAN],[RECREF]) WITH  FILLFACTOR = 80"
   clsADOCon.ExecuteSql sSql
   
   
   ' Expand INVCOMMENTS from 255 to 2048
   sSql = "sp_columns @table_name=CihdTable,@column_name=INVCOMMENTS"
   Set rdotest = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
'   Set rdotest = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
   With rdotest
      If Not .BOF And Not .EOF Then
         If Not IsNull(.Fields(7)) Then
            If .Fields(7) = 255 Then
               sSql = "ALTER TABLE dbo.CihdTable ALTER COLUMN INVCOMMENTS VARCHAR(2048)"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
      .Cancel
   End With
   Set rdotest = Nothing
   
   ' Expand GJEXTDESC from 255 to 512
   sSql = "sp_columns @table_name=GjhdTable,@column_name=GJEXTDESC"
   Set rdotest = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
'   Set rdotest = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
   With rdotest
      If Not .BOF And Not .EOF Then
         If Not IsNull(.Fields(7)) Then
            If .Fields(7) = 255 Then
               sSql = "ALTER TABLE dbo.GjhdTable ALTER COLUMN GJEXTDESC VARCHAR(512)"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
      .Cancel
   End With
   Set rdotest = Nothing
   
   ' Added to support year end journal entries (nth) 06/03/04
   sSql = "ALTER TABLE GjhdTable ADD GJYEAREND tinyint default(0) NULL"
   clsADOCon.ExecuteSql sSql
   
   ' Added to support GL template journals (nth) 04/13/04
   sSql = "ALTER TABLE GjhdTable ADD GJTEMPLATE tinyint default(0) NULL"
   clsADOCon.ExecuteSql sSql
   
   ' Added to support closing GL fiscal periods (nth) 04/05/04.
   sSql = "ALTER TABLE GlfyTable ADD " _
          & "FYCLOSED1 tinyint default(0) NULL," _
          & "FYCLOSED2 tinyint default(0) NULL," _
          & "FYCLOSED3 tinyint default(0) NULL," _
          & "FYCLOSED4 tinyint default(0) NULL," _
          & "FYCLOSED5 tinyint default(0) NULL," _
          & "FYCLOSED6 tinyint default(0) NULL," _
          & "FYCLOSED7 tinyint default(0) NULL," _
          & "FYCLOSED8 tinyint default(0) NULL," _
          & "FYCLOSED9 tinyint default(0) NULL," _
          & "FYCLOSED10 tinyint default(0) NULL," _
          & "FYCLOSED11 tinyint default(0) NULL," _
          & "FYCLOSED12 tinyint default(0) NULL," _
          & "FYCLOSED13 tinyint default(0) NULL"
   clsADOCon.ExecuteSql sSql
   
   ' added for account balance (jcw) 03/15/04
   sSql = "ALTER TABLE Preferences ADD PrefAcctCutOff SMALLDATETIME NULL"
   clsADOCon.ExecuteSql sSql
   
   ' added for budgets (jcw) 02/20/04
   sSql = "CREATE TABLE dbo.BdgtTable (" _
          & "BUDACCT [char] (12) DEFAULT('') NULL," _
          & "BUDFY [int] DEFAULT(0) NULL," _
          & "BUDPER1 [money] DEFAULT(0) NULL ," _
          & "BUDPER2 [money] DEFAULT(0) NULL ," _
          & "BUDPER3 [money] DEFAULT(0) NULL ," _
          & "BUDPER4 [money] DEFAULT(0) NULL ," _
          & "BUDPER5 [money] DEFAULT(0) NULL ," _
          & "BUDPER6 [money] DEFAULT(0) NULL ," _
          & "BUDPER7 [money] DEFAULT(0) NULL ," _
          & "BUDPER8 [money] DEFAULT(0) NULL ," _
          & "BUDPER9 [money] DEFAULT(0) NULL ," _
          & "BUDPER10 [money] DEFAULT(0) NULL ," _
          & "BUDPER11 [money] DEFAULT(0) NULL ," _
          & "BUDPER12 [money] DEFAULT(0) NULL ," _
          & "BUDPER13 [money] DEFAULT(0) NULL)"
   clsADOCon.ExecuteSql sSql
   
   sSql = "CREATE  UNIQUE  CLUSTERED  INDEX [BdgtTable_Unique] ON [dbo].[BdgtTable]([BUDACCT],[BUDFY]) WITH  FILLFACTOR = 80"
   clsADOCon.ExecuteSql sSql
   
   
   ' added for budgets 02/20/04
   sSql = "ALTER TABLE dbo.ComnTable ADD COGLDIVISIONS TINYINT NULL DEFAULT(0)," _
          & "COGLDIVSTARTPOS TINYINT NULL DEFAULT(0)," _
          & "COGLDIVENDPOS TINYINT NULL DEFAULT(0)"
   clsADOCon.ExecuteSql sSql
   
   ' added for vendor 1099 1/26/04
   sSql = "ALTER TABLE VndrTable ADD VEATTORNEY TINYINT NULL DEFAULT(0)"
   clsADOCon.ExecuteSql sSql
   
   ' 2/26/06 AR Invoice Canceled date
   ' set cancelled date for existing cancelled invoices
   Err.Clear
   sSql = "ALTER TABLE CihdTable ADD INVCANCDATE DATETIME NULL"
   clsADOCon.ExecuteSql sSql
   Err.Clear
   sSql = "UPDATE CihdTable SET INVCANCDATE=INVDATE WHERE INVCANCELED=1 AND INVCANCDATE IS NULL"
   clsADOCon.ExecuteSql sSql
   
   '4/11/06 For storing prior invoice number in cm or dm
   Err.Clear
   sSql = "ALTER TABLE CihdTable ADD INVPRIORINV INT DEFAULT(0) NULL"
   clsADOCon.ExecuteSql sSql
   
   '6/7/06 tax changes for AuBeta
   'first make state large enough to accomodate country name
   UpdateTablesQuery "EXEC sp_unbindefault 'TxcdTable.TAXSTATE'"
   UpdateTablesQuery "alter table TxcdTable alter column TAXSTATE varchar(15)"
   UpdateTablesQuery "alter table TxcdTable add constraint DF_TAXSTATE DEFAULT( '' ) FOR TAXSTATE"
   UpdateTablesQuery "ALTER TABLE SoitTable ADD ITFEDTAXRATE decimal(6,4) NULL"
   UpdateTablesQuery "ALTER TABLE SoitTable ADD ITFEDTAXAMT decimal(10,2) NULL"
   UpdateTablesQuery "ALTER TABLE SoitTable ADD ITFEDTAXACCT varchar(12) NULL"
   UpdateTablesQuery "ALTER TABLE SoitTable ADD ITFEDTAXCODE varchar(8) NULL"
   'UpdateTablesQuery "ALTER TABLE SohdTable ADD FEDTAXCODE decimal(10,2) NULL"
   UpdateTablesQuery "ALTER TABLE ComnTable ADD TaxPerItem bit default(0) NULL"
   UpdateTablesQuery "Update ComnTable Set TaxPerItem = 0 where TaxPerItem is NULL"
   ' use COFEDTAXACCT: UpdateTablesQuery "ALTER TABLE ComnTable ADD COSJFEDTAXACCT varchar(12) NULL"
   UpdateTablesQuery "ALTER TABLE ComnTable DROP COLUMN COSJFEDTAXACCT"
   UpdateTablesQuery "ALTER TABLE CihdTable ADD INVFEDTAXACCT varchar(12) NULL"
   
   
   ''''''''''''''''''''''''''''''''''''''''''''''''
   
   'execute updates in UpdateEsiDatabase.sql in esierp directory
   'warning: the file must be ansi, not unicode!
   'don't care if script errs out after the first time.
   Err.Clear
   Dim S As String
   Dim hfile As Long
   hfile = FreeFile
   Open sReportPath + "UpdateEsiDatabase.sql" For Input As hfile
   If Err = 0 Then
      'Get #1, , s
      S = Input$(LOF(hfile), hfile)
      Close #1
      sSql = S
      clsADOCon.ExecuteSql sSql
   End If
   
   diaSql.Prg1.Value = 100
   Unload diaSql
End Sub

Private Sub UpdateTablesQuery(query As String)
   
   On Error Resume Next
   Err.Clear
   sSql = query
   clsADOCon.ExecuteSql sSql
   If Err.Number > 0 Then
      'Debug.Print "Error " & Err.Number & ": " & Err.description
   End If
   
End Sub


'Sets the Workspace and creates the Jet temp database
'Opens the Database in ReopenJet called in reports

Public Sub OpenJet()
   Dim sWindows As String
   On Error Resume Next
   sWindows = GetWindowsDir()
   If Dir(sWindows & "\temp\esifina.mdb") <> "" Then _
          Kill sWindows & "\temp\esifina.mdb"
   If Dir(sWindows & "\temp\") = "" Then _
          MkDir sWindows & "\temp"
   Set JetWkSpace = DBEngine.Workspaces(0)
   Set JetDb = JetWkSpace.CreateDatabase(sWindows & "\temp\" _
               & "Esifina.mdb", dbLangGeneral, dbVersion30)
   
   
End Sub

'Close(if necessary) and Reopen to resolve the
'20534 error in Crystal

Public Sub ReopenJet()
   Dim sWindows As String
   On Error Resume Next
   JetDb.Close
   sWindows = GetWindowsDir() & "\"
   Set JetDb = JetWkSpace.OpenDatabase(sWindows & "\temp\esifina.mdb")
   
End Sub


Public Function GetOldInvoice(sInvoice As String) As Boolean
   Dim RdoInv As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT INVNO FROM CihdTable WHERE " _
          & "INVNO=" & Val(sInvoice) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv)
   If bSqlRows Then
      GetOldInvoice = True
   Else
      GetOldInvoice = False
   End If
   Set RdoInv = Nothing
   Exit Function
   
modErr1:
   CurrError.Number = Err
   CurrError.Description = Err.Description
   Resume modErr2
modErr2:
   DoModuleErrors MdiSect.ActiveForm
   
End Function

Public Function CheckForCancelInv(sInvoice As String, ByRef sPackSlip As String) As Boolean
   Dim RdoInv As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT ISNULL(INVPACKSLIP, '') INVPACKSLIP FROM CihdTable WHERE " _
          & "INVNO=" & Val(sInvoice) & " AND INVCANCELED = 1"
          
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv)
   If bSqlRows Then
      CheckForCancelInv = True
      sPackSlip = Trim(RdoInv!INVPACKSLIP)
      'sCancelDt = Format(Trim(RdoInv!INVCHECKDATE), "mm/dd/yyyy")
   Else
      CheckForCancelInv = False
      sPackSlip = ""
      'sCancelDt = ""
   End If
   Set RdoInv = Nothing
   Exit Function
   
modErr1:
   CurrError.Number = Err
   CurrError.Description = Err.Description
   Resume modErr2
modErr2:
   DoModuleErrors MdiSect.ActiveForm
   
End Function

Public Function FindThisCustomer(frm As Form) As Byte
   Dim rdoCst As ADODB.Recordset
   Dim sCustRef As String
   sCustRef = Compress(frm.cmbCst)
   If Len(sCustRef) = 0 Then Exit Function
   
   On Error GoTo modErr1
   sSql = "SELECT CUREF,CUNICKNAME,CUBTNAME FROM CustTable " _
          & "WHERE CUREF='" & sCustRef & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCst)
   If bSqlRows Then
      On Error Resume Next
      frm.cmbCst = "" & Trim(rdoCst!CUNICKNAME)
      frm.lblNme = "" & Trim(rdoCst!CUBTNAME)
      FindThisCustomer = True
   Else
      On Error Resume Next
      frm.cmbCst = ""
      frm.lblNme = ""
      FindThisCustomer = False
   End If
   Set rdoCst = Nothing
   Exit Function
   
modErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume modErr2
modErr2:
   FindThisCustomer = False
   DoModuleErrors frm
   
End Function

Public Function GetMonthEnd(sNewMonth As String) As String
   Dim i As Integer
   Dim bDay As Byte
   Dim bMonth As Byte
   Dim bYear As Integer
   On Error Resume Next
   bYear = Format(sNewMonth, "yyyy")
   bMonth = Val(Left(sNewMonth, 2))
   Select Case bMonth
      Case 2
         bDay = 28
      Case 1, 3, 5, 7, 8, 10, 12
         bDay = 31
      Case Else
         bDay = 30
   End Select
   
   ' Check for leap year
   If bDay = 28 Then
      For i = 1984 To 2100 Step 4
         If bYear = i Then bDay = 29
      Next
   End If
   GetMonthEnd = Format(bMonth, "00") & "/" _
                 & Trim(str(bDay)) & "/" & Right(sNewMonth, 2)
End Function

Public Sub GetCurrentBuyer(sBuyer As String)
   Dim RdoByr As ADODB.Recordset
   sBuyer = UCase(Compress(sBuyer))
   sSql = "SELECT BYNUMBER,BYLSTNAME,BYFSTNAME,BYMIDINIT FROM " _
          & "BuyrTable WHERE BYREF='" & sBuyer & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoByr, ES_FORWARD)
   If bSqlRows Then
      With RdoByr
         MdiSect.ActiveForm.cmbByr = "" & Trim(!BYNUMBER)
         MdiSect.ActiveForm.lblByr = "" & Trim(!BYFSTNAME) _
                                     & " " & Trim(!BYMIDINIT) & " " & Trim(!BYLSTNAME)
         .Cancel
      End With
   Else
      If Len(Trim(sBuyer)) > 0 Then
         MdiSect.ActiveForm.lblByr = "*** Buyer Wasn't Found ***"
      Else
         MdiSect.ActiveForm.lblByr = ""
      End If
   End If
   Set RdoByr = Nothing
   Exit Sub
   
modErr1:
   sProcName = "getcurrentbuyer"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
End Sub

Public Function GetFYPeriodEnd(dNow As Date) As String
   Dim rdoYr As ADODB.Recordset
   Dim i As Integer
   sSql = "SELECT FYPERSTART1,FYPEREND1,FYPERSTART2,FYPEREND2,FYPERSTART3,FYPEREND3," & vbCrLf _
          & "FYPERSTART4,FYPEREND4,FYPERSTART5,FYPEREND5,FYPERSTART6,FYPEREND6,FYPERSTART7," & vbCrLf _
          & "FYPEREND7,FYPERSTART8,FYPEREND8,FYPERSTART9,FYPEREND9,FYPERSTART10,FYPEREND10," & vbCrLf _
          & "FYPERSTART11,FYPEREND11,FYPERSTART12,FYPEREND12,FYPERSTART13,FYPEREND13 " & vbCrLf _
          & "From GlfyTable Where DateDiff( day, FYSTART, '" & Format(dNow, "mm/dd/yyyy") & "' ) >= 0 " & vbCrLf _
          & "AND DateDiff( day, '" & Format(dNow, "mm/dd/yyyy") & "', FYEND  ) >= 0"
          '& "From GlfyTable Where FYYEAR = " & Format(dNow, "yyyy")
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoYr)
   If bSqlRows Then
      With rdoYr
         For i = 0 To 25 Step 2
            If dNow >= .Fields(i) And dNow <= .Fields(i + 1) Then
               GetFYPeriodEnd = Format(.Fields(i + 1), "mm/dd/yy")
               Exit For
            End If
         Next
         .Cancel
      End With
      Set rdoYr = Nothing
   End If
End Function

Public Sub GetCustBnO(sCust, nRate, sCode, sState, sType)
   ' Get B&O tax codes from customer
   ' Retail takes precidence over wholesale
   
   Dim rdoTx1 As ADODB.Recordset
   Dim rdoTx2 As ADODB.Recordset
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT TAXCODE,TAXSTATE,TAXRATE FROM TxcdTable,CustTable " _
          & "WHERE CUBORTAXCODE = TAXREF AND CUREF = '" & sCust _
          & "' AND TAXTYPE = 0"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoTx1)
   If bSqlRows Then
      With rdoTx1
         nRate = !TAXRATE
         sCode = "" & Trim(!taxCode)
         sState = "" & Trim(!taxState)
         sType = "R"
         .Cancel
      End With
   Else
      sSql = "SELECT TAXCODE,TAXSTATE,TAXRATE FROM TxcdTable,CustTable " _
             & "WHERE CUBORTAXCODE = TAXREF AND CUREF = '" & sCust _
             & "' AND TAXTYPE = 0"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoTx2)
      If bSqlRows Then
         With rdoTx2
            nRate = !TAXRATE
            sCode = "" & Trim(!taxCode)
            sState = "" & Trim(!taxState)
            sType = "W"
            .Cancel
         End With
      End If
   End If
   
   Set rdoTx1 = Nothing
   Set rdoTx2 = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getcustbno"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Sub

Public Sub GetPartBnO(sPart, nRate, sCode, sState, sType)
   ' Get B&O tax codes from part
   ' Retail takes precidence over wholesale
   
   Dim rdoTx1 As ADODB.Recordset
   Dim rdoTx2 As ADODB.Recordset
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT TAXCODE,TAXSTATE,TAXRATE FROM TxcdTable,PartTable WHERE " _
          & "PABORTAX = TAXREF AND TAXTYPE = 0 AND PARTREF = '" & sPart & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoTx1)
   If bSqlRows Then
      With rdoTx1
         nRate = !TAXRATE
         sCode = "" & Trim(!taxCode)
         sState = "" & Trim(!taxState)
         sType = "R"
         .Cancel
      End With
   Else
      sSql = "SELECT TAXCODE,TAXSTATE,TAXRATE FROM TxcdTable,PartTable WHERE " _
             & "PABOWTAX = TAXREF AND TAXTYPE = 0 AND PARTREF = '" & sPart & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoTx2)
      If bSqlRows Then
         With rdoTx2
            nRate = !TAXRATE
            sCode = "" & Trim(!taxCode)
            sState = "" & Trim(!taxState)
            sType = "W"
            .Cancel
         End With
      End If
   End If
   
   Set rdoTx1 = Nothing
   Set rdoTx2 = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getpartbno"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
End Sub

Public Sub GetSalesTaxInfo( _
                           sCust As String, _
                           nRate As Currency, _
                           sCode As String, _
                           sState As String, _
                           sAccount As String)
   
   On Error GoTo DiaErr1
   
   ' Load tax from customer.
   Dim RdoTax As ADODB.Recordset
   sSql = "SELECT TAXCODE,TAXSTATE,TAXRATE,TAXACCT FROM CustTable INNER JOIN " _
          & "TxcdTable ON CustTable.CUTAXCODE = TxcdTable.TAXREF " _
          & "WHERE CUREF = '" & sCust & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTax)
   If bSqlRows Then
      With RdoTax
         nRate = !TAXRATE
         sCode = "" & Trim(!taxCode)
         sState = "" & Trim(!taxState)
         sAccount = "" & Trim(!TAXACCT)
         .Cancel
      End With
   End If
   Set RdoTax = Nothing
   Exit Sub
DiaErr1:
   sProcName = "getsaletaxinfo"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Sub

Public Function FiscalPeriodOpen(dNow As String) As Byte
   ' Returns the number of the open fiscal period.
   ' Zero if no open fiscal period found...
   Dim iYear As Integer
   Dim rdoPer As ADODB.Recordset
   Dim b As Byte
   dNow = Format(dNow, "mm/dd/yyyy")
   iYear = CInt(Right(dNow, 4))
   sSql = "SELECT FYPERSTART1,FYPEREND1,FYCLOSED1,FYPERSTART2," _
          & "FYPEREND2,FYCLOSED2,FYPERSTART3,FYPEREND3,FYCLOSED3,FYPERSTART4," _
          & "FYPEREND4,FYCLOSED4,FYPERSTART5,FYPEREND5,FYCLOSED5,FYPERSTART6," _
          & "FYPEREND6,FYCLOSED6,FYPERSTART7,FYPEREND7,FYCLOSED7,FYPERSTART8," _
          & "FYPEREND8,FYCLOSED8,FYPERSTART9,FYPEREND9,FYCLOSED9,FYPERSTART10,FYPEREND10," _
          & "FYCLOSED10,FYPERSTART11,FYPEREND11,FYCLOSED11,FYPERSTART12,FYPEREND12,FYCLOSED12," _
          & "FYPERSTART13,FYPEREND13,FYCLOSED13 FROM GlfyTable WHERE FYYEAR = " & iYear
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoPer)
   With rdoPer
      For b = 0 To 38 Step 3
         If CDate(dNow) >= CDate(.Fields(b)) And _
                  CDate(dNow) <= CDate(.Fields(b + 1)) Then
            If IsNull(.Fields(b + 2)) Or .Fields(b + 2) = 0 Then
               FiscalPeriodOpen = (b + 1) / 3
               Exit For
            End If
         End If
      Next
   End With
   Set rdoPer = Nothing
   Exit Function
DiaErr1:
   sProcName = "fiscalperiodopen"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
End Function

Public Function GetInvoiceType(sType As String) As String
   Select Case Trim(sType)
      Case "PS"
         GetInvoiceType = "Packing Slip"
      Case "SO"
         GetInvoiceType = "Sales Order"
      Case "CM"
         GetInvoiceType = "Credit Memo"
      Case "DM"
         GetInvoiceType = "Debit Memo"
      Case "CA"
         GetInvoiceType = "Advance Payment"
      Case Else
         GetInvoiceType = ""
   End Select
End Function

Public Sub FillFiscalYears(frm As Form)
   Dim rdoYr As ADODB.Recordset
   On Error GoTo modErr1
   Dim curFy As String
   curFy = ""
   Dim fyStart As Date, fyEnd As Date, today As Date
   today = Now
   sSql = "SELECT FYYEAR, FYSTART, FYEND FROM GlfyTable "
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoYr, ES_FORWARD)
   If bSqlRows Then
      With rdoYr
         Do Until .EOF
            If Not IsNull(.Fields(0)) Then
               AddComboStr frm.cmbFyr.hWnd, Format(.Fields(0), "0000")
               fyStart = CDate(.Fields(1))
               fyEnd = CDate(.Fields(2))
               If DateDiff("d", fyStart, today) >= 0 And DateDiff("d", today, fyEnd) >= 0 Then
                  curFy = Format(.Fields(0), "0000")
               End If
            End If
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set rdoYr = Nothing
   'frm.cmbFyr = Format(ES_SYSDATE, "yyyy")
   If curFy <> "" Then
      frm.cmbFyr = curFy
   End If
   Exit Sub
modErr1:
   sProcName = "fillfisc"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors frm
End Sub


' Restrictions:
' Must not be partitally paid and not a vendor credit memo.

Public Function GetAPDiscount( _
                              sVendor As String, _
                              sInvoice As String, _
                              bDisFrt As Byte, _
                              sAsOf As String) As Currency
   
   Dim RdoInv As ADODB.Recordset
   Dim cTotal As Currency
   
   sProcName = "APDiscount"
   If CDate(sAsOf) > CDate(ES_SYSDATE) Then
      sAsOf = Format(sAsOf, "mm/dd/yy")
   End If
   
   sSql = "SELECT DISTINCT VIDUE,VIPAY,PODDAYS,PODISCOUNT,VEDDAYS,VEDISCOUNT,VIDATE," _
          & "PONUMBER,VIFREIGHT FROM ViitTable INNER JOIN VndrTable ON VITVENDOR=VEREF " _
          & "INNER JOIN VihdTable ON VITNO=VINO AND VITVENDOR=VIVENDOR " _
          & "LEFT OUTER JOIN PohdTable ON VITPORELEASE=PORELEASE AND VITPO=" _
          & "PONUMBER WHERE VINO='" & sInvoice & "' AND VIVENDOR='" & sVendor & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv)
   If bSqlRows Then
      With RdoInv
         If !VIPAY = 0 And !VIDUE > 0 Then
            cTotal = !VIDUE
            If Not bDisFrt Then
               cTotal = cTotal - !VIFREIGHT
            End If
            If IsNull(!PONumber) Or !PONumber = 0 Then
               ' No PO use vendor terms
               If CDate(sAsOf) <= DateAdd("d", !VEDDAYS, !VIDATE) Then
                  GetAPDiscount = (!VEDISCOUNT / 100) * cTotal
               End If
            Else
               If CDate(sAsOf) <= DateAdd("d", !PODDAYS, !VIDATE) Then
                  'GetAPDiscount = SARound((!PODISCOUNT / 100) * cTotal, 2)
                  GetAPDiscount = SARound(!PODISCOUNT * cTotal, 0) / 100
               End If
            End If
         End If
         .Cancel
      End With
   End If
   Set RdoInv = Nothing
End Function

' Returns company setup option Calculate AP discounts on freight

Public Function APDiscFreight() As Byte
   Dim rdoFrt As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT COAPDISC FROM ComnTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoFrt)
   If bSqlRows Then
      With rdoFrt
         APDiscFreight = CByte(.Fields(0))
         .Cancel
      End With
   End If
   Set rdoFrt = Nothing
   Exit Function
modErr1:
   sProcName = "apdiscfreight"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Function

Public Function GetTopSumAcctFlag() As Integer
    Dim RdoSum As ADODB.Recordset
    Dim bRows As Boolean
    Dim iSumAcct As Integer
    iSumAcct = 0
    
    sSql = "SELECT ISNULL(COTOPSUMACCT, 0) as COTOPSUMACCT FROM ComnTable WHERE COREF = 1"
    bRows = clsADOCon.GetDataSet(sSql, RdoSum, ES_FORWARD)

    If bRows Then
        With RdoSum
            iSumAcct = !COTOPSUMACCT
        End With
        'RdoLogo.Close
        ClearResultSet RdoSum
    End If
    
    GetTopSumAcctFlag = iSumAcct
    
End Function


