VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form diaAPf03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Computer Checks"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton optDis 
      Height          =   330
      Left            =   4800
      Picture         =   "diaAPf03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Display The Report"
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Tag             =   "4"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton optPrn 
      Height          =   315
      Left            =   5400
      Picture         =   "diaAPf03a.frx":017E
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Print The Check Run"
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   430
      Left            =   5400
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5880
      Top             =   1080
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4605
      FormDesignWidth =   6360
   End
   Begin VB.ComboBox cmbAct 
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Tag             =   "8"
      Text            =   "cmbAct"
      Top             =   960
      Width           =   1455
   End
   Begin VB.OptionButton optTst 
      Caption         =   "    "
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   3120
      Width           =   615
   End
   Begin VB.OptionButton optChk 
      Caption         =   "    "
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   3360
      Value           =   -1  'True
      Width           =   615
   End
   Begin ComctlLib.ProgressBar progress 
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   4200
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.TextBox txtChkNum 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Tag             =   "1"
      Top             =   1800
      Width           =   1155
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   6
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaAPf03a.frx":0308
      PictureDn       =   "diaAPf03a.frx":044E
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   7
      ToolTipText     =   "Show System Printers"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaAPf03a.frx":0594
      PictureDn       =   "diaAPf03a.frx":06DA
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4920
      TabIndex        =   27
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblTotalCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Total All Checks"
      Height          =   255
      Left            =   3660
      TabIndex        =   26
      Top             =   1860
      Width           =   1215
   End
   Begin VB.Label lblInvStub 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2280
      TabIndex        =   24
      Top             =   2700
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Date"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   23
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   21
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblRec 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   840
      TabIndex        =   20
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Check"
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   19
      Top             =   3840
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Of"
      Height          =   285
      Index           =   7
      Left            =   1560
      TabIndex        =   18
      Top             =   3840
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Actual Checks"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   16
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Test Pattern"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label lbldsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2280
      TabIndex        =   14
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Checking Account"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoices Per Check Stub"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label lblNumChks 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2280
      TabIndex        =   10
      Top             =   480
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Checks In Setup"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Beginning Check Number"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   2055
   End
End
Attribute VB_Name = "diaAPf03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'*************************************************************************************
' diaAPf03a - Print Computer Check
'
' Notes: Each check is created in the temporary jet database. Crystal Reports is
'        then triggered for every check.
'
' Created: (nth)
' Revisons:
' 02/01/02 (nth) Added multiple checking account support.
' 02/01/02 (nth) Fixed error with parital payments appearing as discounts.
' 12/20/02 (nth) Added bCancel.
' 01/08/03 (nth) Reversed CC journal Debit and Credit per WCK and JLH.
' 02/04/03 (nth) Added GetOptions and SaveOptions.
' 02/04/03 (nth) Use vendor address in check address is empty.
' 06/25/03 (nth) Added optional bReprint flag -
'                Transaction can now be queued by diaAPf05a (Reprint computer checks)
' 02/11/04 (nth) Added batch print date.
' 02/11/04 (nth) Added check stock question per Jevco.
' 03/08/04 (nth) Fixed error with invoiced being tied to void checks via JritTable.
' 05/03/04 (nth) Fixed last check in run not being recorded.
'
'*************************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim b As Byte
Dim bMaxInvoices As Byte

Dim lCheck As Double

Dim sApAcct As String
Dim sApDiscAcct As String
Dim sCcAcct As String
Dim sJournalID As String
Dim sMsg As String
Dim sReportName As String

' Are we reprinting checks ?
Public bReprint As Byte
Dim sCheck() As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Public Sub LoadCheckArray(sArray() As String)
   ' Used in reprint checks.  A little work around because
   ' arrays cannot be public varibles
   sCheck() = sArray()
   
End Sub

Private Function GetApAccounts() As Byte
   Dim RdoAps As ADODB.Recordset
   Dim i As Integer
   Dim b As Byte
   sProcName = "getapacco"
   On Error GoTo DiaErr1
   sSql = "SELECT COGLVERIFY,COAPACCT,COAPDISCACCT," _
          & "COCCCASHACCT FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAps, ES_FORWARD)
   If bSqlRows Then
      With RdoAps
         For i = 1 To 3
            If Not IsNull(.Fields(i)) Then
               If Trim(.Fields(i)) = "" Then b = 1
            Else
               b = 1
            End If
         Next
         sApAcct = "" & Trim(!COAPACCT)
         sApDiscAcct = "" & Trim(!COAPDISCACCT)
         sCcAcct = "" & Trim(!COCCCASHACCT)
         If b = 1 Then
            GetApAccounts = 3
         Else
            GetApAccounts = 2
         End If
      End With
   Else
      GetApAccounts = 1
   End If
   Set RdoAps = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "GetApAccounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Function SetCheck(RdoChk As ADODB.Recordset, lCurCheckNum As Long) As Currency
   
   ' Creates the check to be printed in a temp table
   ' Calculates the check totals and returns the value,
   ' which is later passed to postcheck.
   
   On Error GoTo whoops
   
   'Dim rdo As Recordset
   Dim rdo As ADODB.Recordset
   Dim RdoSum As ADODB.Recordset
   
   'On Error Resume Next
   'JetDb.Execute "DELETE * FROM ChHdrTable"
   sSql = "delete from ChHdrTable"
   clsADOCon.ExecuteSql sSql
   
   'Set rdo = JetDb.OpenRecordset("ChHdrTable", dbOpenDynaset)
   sSql = "select * from ChHdrTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo, ES_DYNAMIC)
   
   Dim strChkAdr As String
   
   With RdoChk
      rdo.AddNew
      rdo!Chknum = CStr(lCurCheckNum)
      
      If ("" & Trim(!VECNAME)) = "" Then
         rdo!ChkNme = "" & Trim(!VEBNAME)
      Else
         rdo!ChkNme = "" & Trim(!VECNAME)
      End If
      
      ' Check leaf address / and invoice address different ( removed the CRLF chars)
      strChkAdr = TrimAddress(!VECADR)
      If strChkAdr = "" Then
         rdo!ChkAdd = TrimAddress(!VEBADR) & vbCrLf & Trim(!VEBCITY)
         If Trim(!VEBSTATE) <> "" Then
            rdo!ChkAdd = rdo!ChkAdd & ", " & Compress(Trim(!VEBSTATE)) & " " & Trim(!VEBZIP)
         End If
         rdo!ChkLfAdd = rdo!ChkAdd
      Else
         rdo!ChkAdd = strChkAdr & vbCrLf & Trim(!VECCITY)
         rdo!ChkLfAdd = strChkAdr & vbCrLf & Trim(!VECCITY)
         If Trim(!VECSTATE) <> "" Then
            rdo!ChkAdd = rdo!ChkAdd & ", " & Compress(Trim(!VECSTATE)) & " " & Trim(!VECZIP)
            rdo!ChkLfAdd = rdo!ChkLfAdd & ", " & Compress(Trim(!VECSTATE)) & " " & Trim(!VECZIP)
         End If
      End If
      
      sSql = "SELECT SUM(CHKAMT) AS TotalAmt, SUM(CHKPAMT) AS TotalPAmt " _
             & "FROM ChseTable WHERE CHKVND = '" & Trim(!chkVnd) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoSum)
      
      'rdo!ChkTxt = ConvertCurrency(RdoSum!TotalPAmt, Me)
      rdo!CHKAMT = Format(RdoSum!TotalAmt, "#,###,##0.00")
      rdo!CHKPAMT = Format(RdoSum!TotalPAmt, "#,###,##0.00")
      
      If bReprint Then
         If IsNull(!CHKACCT) Then
            rdo!CHKACCT = Trim(!CHKACCT)
         Else
            rdo!CHKACCT = " "
         End If
      Else
         rdo!CHKACCT = Compress(cmbAct)
      End If
      
      SetCheck = CCur(RdoSum!TotalPAmt)
      
      Set RdoSum = Nothing
      
      If bReprint Then
         rdo!ChkDte = !CHKDATE
      Else
         rdo!ChkDte = txtDte
      End If
      
      rdo!ChkMem = "" & !chkMemo
      rdo.Update
      rdo.Close
      Set rdo = Nothing
   End With
   Exit Function
   
whoops:
   sProcName = "SetCheck"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub PrintReport()
   Dim sWindows As String
   
   On Error GoTo DiaErr1
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   On Error GoTo DiaErr1
   
   MouseCursor 13
   
   ' Report path based on detail or summary types of reports
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   
   sCustomReport = GetCustomReport("finCheck.rpt")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   aFormulaName.Add "CompanyName"
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   
   DoEvents
   sSql = ""
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
   MouseCursor 0
   
'   ReopenJet
'
'   sWindows = GetWindowsDir()
'   'SetMdiReportsize MdiSect
'
'   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   'MdiSect.crw.DataFiles(0) = sWindows & "\temp\esifina.mdb"
'   'MdiSect.crw.ReportFileName = sReportPath & (sReportName & ".rpt")
'   MdiSect.crw.ReportFileName = sReportPath & GetCustomReport("finCheck.rpt")
'
'   ' Turn off crystal's default page N out of N dialog
'   MdiSect.crw.ProgressDialog = False
'
'   ' Let system catch up prevents wierd error
'   ' on my development machine (nth)...
'   DoEvents
'
'   SetCrystalAction Me
   
   Exit Sub
   
DiaErr1:
   sProcName = "PrintReport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub PostCheck(sVend As String, lChkNum As Long, cAmount As Currency, bVoid As Byte)
   
   On Error GoTo whoops
   
   Dim rdo As ADODB.Recordset
   If bVoid = 0 Then
      sSql = "INSERT INTO ChksTable(CHKNUMBER,CHKAMOUNT,CHKVENDOR," & vbCrLf _
             & "CHKPOSTDATE,CHKPRINTDATE,CHKPRINTED,CHKTYPE,CHKACCT,CHKBY)" & vbCrLf _
             & " VALUES( '" & Val(lChkNum) & "'," & vbCrLf _
             & Format(cAmount, "0.00") & "," & vbCrLf _
             & "'" & sVend & "'," & vbCrLf _
             & "'" & txtDte & "'," & vbCrLf _
             & "'" & txtDte & "'," & vbCrLf _
             & "1,2,'" & Compress(cmbAct) & "','" & Secure.UserInitials & "')"
   Else
      ' Void check
      'sSql = "INSERT INTO ChksTable(CHKNUMBER,CHKVOID,CHKVOIDDATE,CHKPOSTDATE,CHKVENDOR)" & vbCrLf _
      '       & " VALUES( " & "'" & Val(lChkNum) & "',1,'" & ES_SYSDATE & "','" _
      '       & ES_SYSDATE & "','" & sVend & "')"
      
      'JetDb.Execute "DELETE * FROM ChHdrTable"
      'Set rdo = JetDb.OpenRecordset("ChHdrTable", dbOpenDynaset)
      clsADOCon.ExecuteSql "delete from ChHdrTable"
      sSql = "select * from ChHdrTable"
      clsADOCon.GetDataSet sSql, rdo, ES_DYNAMIC
            
      ' Update the jet report table - voiding out all the legal fields
      rdo.AddNew
      rdo!Chknum = CStr(lChkNum)
      rdo!CHKACCT = Compress(cmbAct)
      rdo!ChkMem = "VOID"
      rdo!ChkTxt = "VOID - AMOUNT - VOID - AMOUNT - VOID - AMOUNT - VOID"
      rdo!CHKAMT = 0
      rdo!CHKPAMT = 0
      rdo.Update
      rdo.Close
      Set rdo = Nothing
      
      sSql = "INSERT INTO ChksTable(CHKNUMBER,CHKVOID,CHKVOIDDATE,CHKPOSTDATE,CHKVENDOR,CHKACCT)" & vbCrLf _
             & " VALUES( " & "'" & Val(lChkNum) & "',1,'" & txtDte & "','" _
             & txtDte & "','" & sVend & "','" & Compress(cmbAct) & "')"
      
   End If
   
   ' If we are not reprinting them
   If Not bReprint Then
      clsADOCon.ExecuteSql sSql
      SaveLastCheck CStr(lChkNum), cmbAct
   End If
   Exit Sub
   
whoops:
   sProcName = "PostCheck"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub PayInvoice(RdoChk As ADODB.Recordset, lChkNum As Long, strChkAcct As String)
   Dim iRef As Integer
   Dim iTrans As Integer
   Dim cTotal As Currency
   Dim cPaid As Currency
   Dim cDiscount As Currency
   Dim sDebit As String
   Dim sCredit As String
   
   
   On Error GoTo DiaErr1
   ' Bail if we are just doing a reprint.
   If bReprint Then
      Exit Sub
   End If
   
   ' Apply payment to invoice
   With RdoChk
      
      Dim PaidInFull As Integer
      If !CHKPAMT + !CHKDIS = !CHKAMT Then
         PaidInFull = 1
      End If
      
      sSql = "UPDATE VihdTable " _
             & "SET VIPIF=" & PaidInFull & "," _
             & "VIDISCOUNT=" & !CHKDIS & "," _
             & "VIPAY=VIPAY+" & !CHKPAMT & "," _
             & "VICHECKNO='" & CStr(lChkNum) & "', " _
             & "VICHKACCT='" & CStr(strChkAcct) & "' " _
             & "WHERE VIVENDOR = '" & Trim(!chkVnd) & "' AND VINO = '" _
             & Trim(!VINO) & "'"
      cTotal = !CHKPAMT + !CHKDIS
      
      clsADOCon.ExecuteSql sSql
      
      ' Journal entries
      iTrans = GetNextTransaction(sJournalID)
      iRef = 0
      cPaid = Abs(!CHKPAMT)
      cDiscount = Abs(!CHKDIS)
      cTotal = Abs(cTotal)
      '*
      '*  AP = CC + DIS
      '*  or
      '*  CC = AP - DIS
      '*
      
      If !VIDUE < 0 Then
         sDebit = "DCCREDIT"
         sCredit = "DCDEBIT"
      Else
         sDebit = "DCDEBIT"
         sCredit = "DCCREDIT"
      End If
      
      ' 1 - DEBIT accounts payable
      iRef = iRef + 1
      sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," & sDebit & ",DCACCTNO," _
         & "DCDATE,DCCHECKNO,DCCHKACCT,DCVENDOR,DCVENDORINV,DCTYPE) " & vbCrLf _
         & "VALUES('" _
         & sJournalID & "'," _
         & iTrans & "," _
         & iRef & "," _
         & Format(cTotal, "0.00") & ",'" _
         & Compress(sApAcct) & "','" _
         & txtDte & "','" _
         & CStr(lChkNum) & "','" _
         & CStr(strChkAcct) & "','" _
         & Trim(!chkVnd) & "','" _
         & Trim(!VINO) & "'," _
         & DCTYPE_AP_ChkAPAcct & ")"
      clsADOCon.ExecuteSql sSql
      
      ' 2  - CREDIT Checking Account
      iRef = iRef + 1
      sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," & sCredit & ",DCACCTNO," _
         & "DCDATE,DCCHECKNO,DCCHKACCT, DCVENDOR,DCVENDORINV,DCTYPE) " & vbCrLf _
         & "VALUES('" _
         & sJournalID & "'," _
         & iTrans & "," _
         & iRef & "," _
         & Format(cPaid, "0.00") & ",'" _
         & Compress(sCcAcct) & "','" _
         & txtDte & "','" _
         & CStr(lChkNum) & "','" _
         & CStr(strChkAcct) & "','" _
         & Trim(!chkVnd) & "','" _
         & Trim(!VINO) & "'," _
         & DCTYPE_AP_ChkChkAcct & ")"
      clsADOCon.ExecuteSql sSql
      
      ' 3 - CREDIT discounts if applicable
      If cDiscount > 0 Then
         iRef = iRef + 1
         sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," & sCredit & ",DCACCTNO," _
            & "DCDATE,DCCHECKNO,DCCHKACCT, DCVENDOR,DCVENDORINV,DCTYPE) " & vbCrLf _
            & "VALUES('" _
            & sJournalID & "'," _
            & iTrans & "," _
            & iRef & "," _
            & Format(cDiscount, "0.00") & ",'" _
            & sApDiscAcct & "','" _
            & txtDte & "'," _
            & CStr(lChkNum) & ",'" _
            & CStr(strChkAcct) & "','" _
            & Trim(!chkVnd) & "','" _
            & Trim(!VINO) & "'," _
            & DCTYPE_AP_ChkDiscAcct & ")"
         clsADOCon.ExecuteSql sSql
      End If
   End With
   Exit Sub
   
DiaErr1:
   sProcName = "PayInvoice"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub PrintChecks()
   Dim RdoChk As ADODB.Recordset
   Dim rdoHeader As ADODB.Recordset 'Temp check table
   Dim rdoDetail As ADODB.Recordset 'Temp invoices (check stub) table
   Dim sField As String ' Used to fill the check stub table
   Dim sCurVnd As String ' Current Vendor
   Dim iCount As Integer
   Dim iInvCount As Integer ' Number of Invoices in Setup
   Dim lCurCheckNum As Long ' Check Number
   Dim strChkAcct As String
   
   Dim bSuccess As Byte
   Dim cAmount As Currency ' Amount of current check.  Values is retrived from SetCheck
   Dim i As Integer
   
   
   On Error GoTo DiaErr1
   sProcName = "printchecks"
   sSql = "SELECT CHKNUM,CHKVND,CHKAMT,CHKPAMT,CHKDIS,CHKDATE,CHKMEMO,CHKREPRINTNO,CHKACCT,VINO," & vbCrLf _
          & "VIVENDOR,VIDUE,VIPAY,VIDATE,VECNAME,VECADR,VECCITY,VECSTATE,VECZIP, " & vbCrLf _
          & "VEBNAME,VEBADR,VEBCITY,VEBSTATE,VEBZIP,VITYPE " & vbCrLf _
          & "FROM (VihdTable INNER JOIN VndrTable ON VihdTable.VIVENDOR = VndrTable.VEREF) " & vbCrLf _
          & "INNER JOIN ChseTable ON (VihdTable.VIVENDOR = ChseTable.CHKVND) " & vbCrLf _
          & "AND (VihdTable.VINO = ChseTable.CHKINV)"
   If bReprint Then
      sSql = sSql & " WHERE CHKREPRINTNO <> 0 ORDER BY CHKREPRINTNO"
   Else
      sSql = sSql & " ORDER BY CHKNUM"
   End If
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk, ES_STATIC)
   
   progress.Visible = True
   progress.max = RdoChk.RecordCount
   
   If bSqlRows Then
      
      'TODO Need the Transcation
      'clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      
      With RdoChk
         ' Get the first check and initialize it.
         If bReprint Then
            lCurCheckNum = CLng(!CHKREPRINTNO)
            strChkAcct = Trim(!CHKACCT)
         Else
            lCurCheckNum = CLng(txtChkNum)
            strChkAcct = Compress(cmbAct)
         End If
         
         sCurVnd = Trim(!VIVENDOR)
         
         cAmount = SetCheck(RdoChk, lCurCheckNum)
         
         'JetDb.Execute "DELETE * FROM ChDetTable"
         clsADOCon.ExecuteSql "delete from ChDetTable"
         'Set rdoDetail = JetDb.OpenRecordset("ChDetTable", dbOpenDynaset)
         sSql = "select * from ChDetTable"
         clsADOCon.GetDataSet sSql, rdoDetail, ES_DYNAMIC
         
         rdoDetail.AddNew
         rdoDetail!DetNum01 = CStr(lCurCheckNum)
         rdoDetail.Update
         Set rdoDetail = Nothing
         
         While Not .EOF
            iCount = iCount + 1
            progress.Value = iCount
            iInvCount = iInvCount + 1
            
            ' New check
            If sCurVnd <> Trim(!VIVENDOR) Then
               iInvCount = 1
               
               ' Post and print check
               sProcName = "postcheck"
               Debug.Print "print check # " & lCurCheckNum & " for " & sCurVnd
               PostCheck sCurVnd, lCurCheckNum, cAmount, 0
               PrintReport
               
               
               ' Reset for next vendor
               'JetDb.Execute "DELETE * FROM ChHdrTable"
               clsADOCon.ExecuteSql "delete from ChHdrTable"
               
               If bReprint Then
                  lCurCheckNum = CLng(!CHKREPRINTNO)
               Else
                  lCurCheckNum = lCurCheckNum + 1
               End If
               sCurVnd = Trim(!VIVENDOR)
               cAmount = SetCheck(RdoChk, lCurCheckNum)
               'JetDb.Execute "DELETE * FROM ChDetTable"
               clsADOCon.ExecuteSql "delete from ChDetTable"
               'Set rdoDetail = JetDb.OpenRecordset("ChDetTable", dbOpenDynaset)
               sSql = "select * from ChDetTable"
               clsADOCon.GetDataSet sSql, rdoDetail, ES_DYNAMIC
               rdoDetail.AddNew
               rdoDetail!DetNum01 = CStr(lCurCheckNum)
               rdoDetail.Update
               Set rdoDetail = Nothing
            End If
            
            ' Too many invoice for stub.  Void and move to next check
            If iInvCount > Val(lblInvStub) Then
'Debug.Print "print other than last page"
               iInvCount = 1
               ' Write a void check to the ChksTable
               sProcName = "postcheck"
               PostCheck sCurVnd, lCurCheckNum, cAmount, 1
               sProcName = "printchecks"
               PrintReport
               
               If bReprint Then
                  lCurCheckNum = CLng(!CHKREPRINTNO)
               Else
                  sSql = "UPDATE JritTable SET DCCHECKNO = '" _
                         & lCurCheckNum + 1 _
                         & "' WHERE DCCHECKNO = '" & lCurCheckNum & "'" & vbCrLf _
                         & "AND ( DCHEAD LIKE 'CC%' OR DCHEAD LIKE 'XC%' )" & vbCrLf _
                         & " AND DCCHKACCT = '" & CStr(strChkAcct) & "'"
                  clsADOCon.ExecuteSql sSql
                  '// Added code to reset the check number if th echeck count is more than 15
                  sSql = "UPDATE VihdTable SET " _
                         & "VICHECKNO='" & CStr(lCurCheckNum + 1) & "', " & vbCrLf _
                         & "VICHKACCT='" & CStr(strChkAcct) & "' " & vbCrLf _
                         & " WHERE VICHECKNO = '" & Trim(lCurCheckNum) & "' AND VICHKACCT = '" & vbCrLf _
                         & CStr(strChkAcct) & "'"
                  clsADOCon.ExecuteSql sSql
                  
                  lCurCheckNum = lCurCheckNum + 1
               End If
               
               SetCheck RdoChk, lCurCheckNum
               'JetDb.Execute "DELETE * FROM ChDetTable"
               clsADOCon.ExecuteSql "delete from ChDetTable"
               
               'Set rdoDetail = JetDb.OpenRecordset("ChDetTable", dbOpenDynaset)
               sSql = "select * from ChDetTable"
               clsADOCon.GetDataSet sSql, rdoDetail, ES_DYNAMIC
               rdoDetail.AddNew
               rdoDetail!DetNum01 = CStr(lCurCheckNum)
               rdoDetail.Update
               Set rdoDetail = Nothing
               'JetDb.Execute "UPDATE ChHdrTable SET ChkMem ='" & !chkMemo & "'"
               clsADOCon.ExecuteSql "UPDATE ChHdrTable SET ChkMem ='" & !chkMemo & "'"
            End If
            
Debug.Print "(" & iInvCount & ") " & !VINO
            sField = "DetInv" & Format(iInvCount, "00")
            'JetDb.Execute "UPDATE ChDetTable SET " & sField & " = '" & Trim(!VINO) & "'"
            clsADOCon.ExecuteSql "UPDATE ChDetTable SET " & sField & " = '" & Trim(!VINO) & "'"
            sField = "DetNum" & Format(iInvCount, "00")
            'JetDb.Execute "UPDATE ChDetTable SET " & sField & " = '" & CStr(lCurCheckNum) & "'"
            clsADOCon.ExecuteSql "UPDATE ChDetTable SET " & sField & " = '" & CStr(lCurCheckNum) & "'"
            sField = "DetDte" & Format(iInvCount, "00")
            'JetDb.Execute "UPDATE ChDetTable SET " & sField & " = '" & CStr(!VIDATE) & "'"
            clsADOCon.ExecuteSql "UPDATE ChDetTable SET " & sField & " = '" & CStr(!VIDATE) & "'"
            sField = "DetAmt" & Format(iInvCount, "00")
            'JetDb.Execute "UPDATE ChDetTable SET " & sField & " = " & CCur(!CHKAMT)
            clsADOCon.ExecuteSql "UPDATE ChDetTable SET " & sField & " = " & CCur(!CHKAMT)
            
            If Not IsNull(!CHKDIS) Then
               sField = "DetDis" & Format(iInvCount, "00")
               'JetDb.Execute "UPDATE ChDetTable SET " & sField & " = " & CCur(!CHKDIS)
               clsADOCon.ExecuteSql "UPDATE ChDetTable SET " & sField & " = " & CCur(!CHKDIS)
            End If
            
            sField = "DetPAmt" & Format(iInvCount, "00")
            'JetDb.Execute "UPDATE ChDetTable SET " & sField & " = " & CCur(!CHKPAMT)
            clsADOCon.ExecuteSql "UPDATE ChDetTable SET " & sField & " = " & CCur(!CHKPAMT)
            'ReopenJet
            
            PayInvoice RdoChk, lCurCheckNum, strChkAcct
            
            .MoveNext
         Wend
         
         ' Print the last one....
         Debug.Print "print last page - check # " & lCurCheckNum & " for " & sCurVnd
         PrintReport
         ' Backup one record
         .MovePrevious
         sProcName = "postcheck"
         PostCheck sCurVnd, lCurCheckNum, cAmount, 0
         sProcName = "printchecks"
      End With
   End If
   
   If clsADOCon.ADOErrNum = 0 Then
      'clsADOCon.CommitTrans
      
      If Val(lblNumChks) > 1 Then
         sMsg = Val(lblNumChks) & " Checks "
      Else
         sMsg = "1 Check "
      End If
      sMsg = sMsg & "Successfully Printed."
      
      If Not bReprint Then
         If DeleteCheckSetup Then
            sMsg = sMsg & vbCrLf & "Check Setup Cleared."
         Else
            sMsg = "Unable To Clear Check Setup. " & vbCrLf _
                   & "Please Contact System Administrator."
         End If
      Else
         ' Purge out reprint checks from check setup
         sSql = "DELETE FROM ChseTable WHERE CHKREPRINTNO IN("
         For i = 0 To UBound(sCheck)
            sSql = sSql & sCheck(i) & ","
         Next
         sSql = Left(sSql, Len(sSql) - 1) & ")"
         clsADOCon.ExecuteSql sSql
         Erase sCheck
      End If
      lblNumChks = "0"
      MsgBox sMsg, vbInformation, Caption
   Else
      'clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
      
      MsgBox "Transactions Were not Successful.  Error # " & Err _
         & "Please Contact System Administrator.", vbExclamation, Caption
   End If
   
   optPrn.enabled = True
   'MdiSect.crw.ProgressDialog = True
   progress.Visible = False
   Set RdoChk = Nothing
   Set rdoDetail = Nothing
   Set rdoHeader = Nothing
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

'''Private Sub BuildCheckTables()
'''   Dim NewTb1 As TableDef
'''   Dim NewTb2 As TableDef
'''   Dim NewFld As Field
'''   Dim NewIdx1 As Index
'''   Dim NewIdx2 As Index
'''
'''
'''   On Error Resume Next
'''
'''
'''   JetDb.Execute "DROP TABLE ChhdrTable"
'''   JetDb.Execute "DROP TABLE ChDetTable"
'''   'Fields. Note that we allow empties
'''   'Header
'''
'''
'''
'''   Set NewTb1 = JetDb.CreateTableDef("ChHdrTable")
'''   With NewTb1
'''      'Check
'''      .Fields.Append .CreateField("ChkNum", dbText, 12)
'''      .Fields(0).AllowZeroLength = True
'''      'Vendor
'''      .Fields.Append .CreateField("ChkVnd", dbText, 10)
'''      .Fields(1).AllowZeroLength = True
'''      'Vendor Name
'''      .Fields.Append .CreateField("ChkNme", dbText, 40)
'''      .Fields(2).AllowZeroLength = True
'''      'Vendor Address
'''      .Fields.Append .CreateField("ChkAdd", dbText, 160)
'''      .Fields(3).AllowZeroLength = True
'''      'Vendor City
'''      .Fields.Append .CreateField("ChkCty", dbText, 18)
'''      .Fields(4).AllowZeroLength = True
'''      'Vendor State
'''      .Fields.Append .CreateField("ChkSte", dbText, 4)
'''      .Fields(5).AllowZeroLength = True
'''      'Vendor Zip
'''      .Fields.Append .CreateField("ChkZip", dbText, 10)
'''      .Fields(6).AllowZeroLength = True
'''      'Check amount
'''      .Fields.Append .CreateField("ChkAmt", dbText)
'''      .Fields(7).DefaultValue = 0
'''      'Check Partial/discout total amount
'''      .Fields.Append .CreateField("ChkPAmt", dbText)
'''      .Fields(8).DefaultValue = 0
'''      'Check Date
'''      .Fields.Append .CreateField("ChkDte", dbDate)
'''      'Check Text
'''      .Fields.Append .CreateField("ChkTxt", dbText, 80)
'''      .Fields(9).AllowZeroLength = True
'''      'Check Memo
'''      .Fields.Append .CreateField("ChkMem", dbText, 40)
'''      .Fields(10).AllowZeroLength = True
'''      'Checking Account
'''      .Fields.Append .CreateField("ChkAcct", dbText, 12)
'''      .Fields(10).AllowZeroLength = True
'''   End With
'''
'''   'add the table and indexes to Jet.
'''   JetDb.TableDefs.Append NewTb1
'''   Set NewTb1 = JetDb!ChHdrTable
'''   With NewTb1
'''      Set NewIdx1 = .CreateIndex
'''      With NewIdx1
'''         .Name = "CheckIdx"
'''         .Fields.Append .CreateField("ChkNum")
'''      End With
'''      .Indexes.Append NewIdx1
'''   End With
'''
'''   'Details
'''   Set NewTb2 = JetDb.CreateTableDef("ChDetTable")
'''   With NewTb2
'''      '1
'''      'Check
'''      .Fields.Append .CreateField("DetNum01", dbText, 12)
'''      .Fields(0).AllowZeroLength = True
'''      'Inv Number
'''      .Fields.Append .CreateField("DetInv01", dbText, 20)
'''      .Fields(1).AllowZeroLength = True
'''      'Inv Date
'''      .Fields.Append .CreateField("DetDte01", dbDate)
'''      '.Fields(2)
'''      'Inv Notes
'''      .Fields.Append .CreateField("DetDis01", dbCurrency)
'''      .Fields(3).AllowZeroLength = 0
'''      'Inv Type
'''      .Fields.Append .CreateField("DetPAmt01", dbCurrency)
'''      .Fields(4).AllowZeroLength = 0
'''      'Inv Amt
'''      .Fields.Append .CreateField("DetAmt01", dbCurrency)
'''      .Fields(5).DefaultValue = 0
'''
'''      '2
'''      .Fields.Append .CreateField("DetNum02", dbText, 12)
'''      .Fields(6).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetInv02", dbText, 20)
'''      .Fields(7).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetDte02", dbDate)
'''      '.Fields(8)
'''
'''
'''      .Fields.Append .CreateField("DetDis02", dbCurrency)
'''      .Fields(3).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetPAmt02", dbCurrency)
'''      .Fields(4).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetAmt02", dbCurrency)
'''      .Fields(5).DefaultValue = 0
'''
'''      '3
'''      .Fields.Append .CreateField("DetNum03", dbText, 12)
'''      .Fields(12).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetInv03", dbText, 20)
'''      .Fields(13).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetDte03", dbDate)
'''      '.Fields(14)
'''
'''
'''      .Fields.Append .CreateField("DetDis03", dbCurrency)
'''      .Fields(3).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetPAmt03", dbCurrency)
'''      .Fields(4).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetAmt03", dbCurrency)
'''      .Fields(5).DefaultValue = 0
'''
'''      '4
'''      .Fields.Append .CreateField("DetNum04", dbText, 12)
'''      .Fields(18).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetInv04", dbText, 20)
'''      .Fields(19).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetDte04", dbDate)
'''      '.Fields(20)
'''
'''
'''      .Fields.Append .CreateField("DetDis04", dbCurrency)
'''      .Fields(3).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetPAmt04", dbCurrency)
'''      .Fields(4).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetAmt04", dbCurrency)
'''      .Fields(5).DefaultValue = 0
'''
'''      '5
'''      .Fields.Append .CreateField("DetNum05", dbText, 12)
'''      .Fields(24).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetInv05", dbText, 20)
'''      .Fields(25).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetDte05", dbDate)
'''      '.Fields(26)
'''
'''      'Inv Notes
'''      .Fields.Append .CreateField("DetDis05", dbCurrency)
'''      .Fields(3).AllowZeroLength = True
'''      'Inv Type
'''      .Fields.Append .CreateField("DetPAmt05", dbCurrency)
'''      .Fields(4).AllowZeroLength = True
'''      'Inv Amt
'''      .Fields.Append .CreateField("DetAmt05", dbCurrency)
'''      .Fields(5).DefaultValue = 0
'''
'''      '6
'''      .Fields.Append .CreateField("DetNum06", dbText, 12)
'''      .Fields(30).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetInv06", dbText, 20)
'''      .Fields(31).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetDte06", dbDate)
'''      '.Fields(32)
'''
'''      'Inv Notes
'''      .Fields.Append .CreateField("DetDis06", dbCurrency)
'''      .Fields(3).AllowZeroLength = True
'''      'Inv Type
'''      .Fields.Append .CreateField("DetPAmt06", dbCurrency)
'''      .Fields(4).AllowZeroLength = True
'''      'Inv Amt
'''      .Fields.Append .CreateField("DetAmt06", dbCurrency)
'''      .Fields(5).DefaultValue = 0
'''
'''      '7
'''      .Fields.Append .CreateField("DetNum07", dbText, 12)
'''      .Fields(36).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetInv07", dbText, 20)
'''      .Fields(37).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetDte07", dbDate)
'''      '.Fields(38)
'''
'''      'Inv Notes
'''      .Fields.Append .CreateField("DetDis07", dbCurrency)
'''      .Fields(3).AllowZeroLength = True
'''      'Inv Type
'''      .Fields.Append .CreateField("DetPAmt07", dbCurrency)
'''      .Fields(4).AllowZeroLength = True
'''      'Inv Amt
'''      .Fields.Append .CreateField("DetAmt07", dbCurrency)
'''      .Fields(5).DefaultValue = 0
'''
'''      '8
'''      .Fields.Append .CreateField("DetNum08", dbText, 12)
'''      .Fields(42).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetInv08", dbText, 20)
'''      .Fields(43).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetDte08", dbDate)
'''      '.Fields(44)
'''
'''      'Inv Notes
'''      .Fields.Append .CreateField("DetDis08", dbCurrency)
'''      .Fields(3).AllowZeroLength = True
'''      'Inv Type
'''      .Fields.Append .CreateField("DetPAmt08", dbCurrency)
'''      .Fields(4).AllowZeroLength = True
'''      'Inv Amt
'''      .Fields.Append .CreateField("DetAmt08", dbCurrency)
'''      .Fields(5).DefaultValue = 0
'''
'''      '9
'''      .Fields.Append .CreateField("DetNum09", dbText, 12)
'''      .Fields(48).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetInv09", dbText, 20)
'''      .Fields(49).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetDte09", dbDate)
'''      '.Fields(50)
'''
'''      'Inv Notes
'''      .Fields.Append .CreateField("DetDis09", dbCurrency)
'''      .Fields(3).AllowZeroLength = True
'''      'Inv Type
'''      .Fields.Append .CreateField("DetPAmt09", dbCurrency)
'''      .Fields(4).AllowZeroLength = True
'''      'Inv Amt
'''      .Fields.Append .CreateField("DetAmt09", dbCurrency)
'''      .Fields(5).DefaultValue = 0
'''
'''      '10
'''      .Fields.Append .CreateField("DetNum10", dbText, 12)
'''      .Fields(54).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetInv10", dbText, 20)
'''      .Fields(55).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetDte10", dbDate)
'''      '.Fields(56)
'''
'''      'Inv Notes
'''      .Fields.Append .CreateField("DetDis10", dbCurrency)
'''      .Fields(3).AllowZeroLength = True
'''      'Inv Type
'''      .Fields.Append .CreateField("DetPAmt10", dbCurrency)
'''      .Fields(4).AllowZeroLength = True
'''      'Inv Amt
'''      .Fields.Append .CreateField("DetAmt10", dbCurrency)
'''      .Fields(5).DefaultValue = 0
'''
'''      '11
'''      .Fields.Append .CreateField("DetNum11", dbText, 12)
'''      .Fields(60).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetInv11", dbText, 20)
'''      .Fields(61).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetDte11", dbDate)
'''      '.Fields(62)
'''
'''      'Inv Notes
'''      .Fields.Append .CreateField("DetDis11", dbCurrency)
'''      .Fields(3).AllowZeroLength = True
'''      'Inv Type
'''      .Fields.Append .CreateField("DetPAmt11", dbCurrency)
'''      .Fields(4).AllowZeroLength = True
'''      'Inv Amt
'''      .Fields.Append .CreateField("DetAmt11", dbCurrency)
'''      .Fields(5).DefaultValue = 0
'''
'''      '12
'''      .Fields.Append .CreateField("DetNum12", dbText, 12)
'''      .Fields(66).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetInv12", dbText, 20)
'''      .Fields(67).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetDte12", dbDate)
'''      '.Fields(68)
'''
'''      'Inv Notes
'''      .Fields.Append .CreateField("DetDis12", dbCurrency)
'''      .Fields(3).AllowZeroLength = True
'''      'Inv Type
'''      .Fields.Append .CreateField("DetPAmt12", dbCurrency)
'''      .Fields(4).AllowZeroLength = True
'''      'Inv Amt
'''      .Fields.Append .CreateField("DetAmt12", dbCurrency)
'''      .Fields(5).DefaultValue = 0
'''
'''      '13
'''      .Fields.Append .CreateField("DetNum13", dbText, 12)
'''      .Fields(72).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetInv13", dbText, 20)
'''      .Fields(73).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetDte13", dbDate)
'''      '.Fields(74)
'''
'''      'Inv Notes
'''      .Fields.Append .CreateField("DetDis13", dbCurrency)
'''      .Fields(3).AllowZeroLength = True
'''      'Inv Type
'''      .Fields.Append .CreateField("DetPAmt13", dbCurrency)
'''      .Fields(4).AllowZeroLength = True
'''      'Inv Amt
'''      .Fields.Append .CreateField("DetAmt13", dbCurrency)
'''      .Fields(5).DefaultValue = 0
'''
'''      '14
'''      .Fields.Append .CreateField("DetNum14", dbText, 12)
'''      .Fields(78).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetInv14", dbText, 20)
'''      .Fields(79).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetDte14", dbDate)
'''      '.Fields(80)
'''
'''      'Inv Notes
'''      .Fields.Append .CreateField("DetDis14", dbCurrency)
'''      .Fields(3).AllowZeroLength = True
'''      'Inv Type
'''      .Fields.Append .CreateField("DetPAmt14", dbCurrency)
'''      .Fields(4).AllowZeroLength = True
'''      'Inv Amt
'''      .Fields.Append .CreateField("DetAmt14", dbCurrency)
'''      .Fields(5).DefaultValue = 0
'''
'''      '15
'''      .Fields.Append .CreateField("DetNum15", dbText, 12)
'''      .Fields(84).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetInv15", dbText, 20)
'''      .Fields(85).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetDte15", dbDate)
'''      '.Fields(86)
'''
'''      'Inv Notes
'''      .Fields.Append .CreateField("DetDis15", dbCurrency)
'''      .Fields(3).AllowZeroLength = True
'''      'Inv Type
'''      .Fields.Append .CreateField("DetPAmt15", dbCurrency)
'''      .Fields(4).AllowZeroLength = True
'''      'Inv Amt
'''      .Fields.Append .CreateField("DetAmt15", dbCurrency)
'''      .Fields(5).DefaultValue = 0
'''
'''      '16
'''      .Fields.Append .CreateField("DetNum16", dbText, 12)
'''      .Fields(90).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetInv16", dbText, 20)
'''      .Fields(91).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetDte16", dbDate)
'''      '.Fields(92)
'''
'''      'Inv Notes
'''      .Fields.Append .CreateField("DetDis16", dbCurrency)
'''      .Fields(3).AllowZeroLength = True
'''      'Inv Type
'''      .Fields.Append .CreateField("DetPAmt16", dbCurrency)
'''      .Fields(4).AllowZeroLength = True
'''      'Inv Amt
'''      .Fields.Append .CreateField("DetAmt16", dbCurrency)
'''      .Fields(5).DefaultValue = 0
'''
'''      '17
'''      .Fields.Append .CreateField("DetNum17", dbText, 12)
'''      .Fields(96).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetInv17", dbText, 20)
'''      .Fields(97).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetDte17", dbDate)
'''      '.Fields(98)
'''
'''      'Inv Notes
'''      .Fields.Append .CreateField("DetDis17", dbCurrency)
'''      .Fields(3).AllowZeroLength = True
'''      'Inv Type
'''      .Fields.Append .CreateField("DetPAmt17", dbCurrency)
'''      .Fields(4).AllowZeroLength = True
'''      'Inv Amt
'''      .Fields.Append .CreateField("DetAmt17", dbCurrency)
'''      .Fields(5).DefaultValue = 0
'''
'''      '18
'''      .Fields.Append .CreateField("DetNum18", dbText, 12)
'''      .Fields(102).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetInv18", dbText, 20)
'''      .Fields(103).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetDte18", dbDate)
'''      '.Fields(104)
'''
'''      'Inv Notes
'''      .Fields.Append .CreateField("DetDis18", dbCurrency)
'''      .Fields(3).AllowZeroLength = True
'''      'Inv Type
'''      .Fields.Append .CreateField("DetPAmt18", dbCurrency)
'''      .Fields(4).AllowZeroLength = True
'''      'Inv Amt
'''      .Fields.Append .CreateField("DetAmt18", dbCurrency)
'''      .Fields(5).DefaultValue = 0
'''
'''      '19
'''      .Fields.Append .CreateField("DetNum19", dbText, 12)
'''      .Fields(108).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetInv19", dbText, 20)
'''      .Fields(109).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetDte19", dbDate)
'''      '.Fields(110)
'''
'''      'Inv Notes
'''      .Fields.Append .CreateField("DetDis19", dbCurrency)
'''      .Fields(3).AllowZeroLength = True
'''      'Inv Type
'''      .Fields.Append .CreateField("DetPAmt19", dbCurrency)
'''      .Fields(4).AllowZeroLength = True
'''      'Inv Amt
'''      .Fields.Append .CreateField("DetAmt19", dbCurrency)
'''      .Fields(5).DefaultValue = 0
'''
'''      '20
'''      .Fields.Append .CreateField("DetNum20", dbText, 12)
'''      .Fields(114).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetInv20", dbText, 20)
'''      .Fields(115).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetDte20", dbDate)
'''      '.Fields(116)
'''
'''      'Inv Notes
'''      .Fields.Append .CreateField("DetDis20", dbCurrency)
'''      .Fields(3).AllowZeroLength = True
'''      'Inv Type
'''      .Fields.Append .CreateField("DetPAmt20", dbCurrency)
'''      .Fields(4).AllowZeroLength = True
'''      'Inv Amt
'''      .Fields.Append .CreateField("DetAmt20", dbCurrency)
'''      .Fields(5).DefaultValue = 0
'''
'''      '21
'''      .Fields.Append .CreateField("DetNum21", dbText, 12)
'''      .Fields(120).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetInv21", dbText, 20)
'''      .Fields(121).AllowZeroLength = True
'''
'''      .Fields.Append .CreateField("DetDte21", dbDate)
'''      '.Fields(122)
'''
'''      'Inv Notes
'''      .Fields.Append .CreateField("DetDis21", dbCurrency)
'''      .Fields(3).AllowZeroLength = True
'''      'Inv Type
'''      .Fields.Append .CreateField("DetPAmt21", dbCurrency)
'''      .Fields(4).AllowZeroLength = True
'''      'Inv Amt
'''      .Fields.Append .CreateField("DetAmt21", dbCurrency)
'''      .Fields(5).DefaultValue = 0
'''
'''   End With
'''
'''   'add the table and indexes to Jet.
'''   JetDb.TableDefs.Append NewTb2
'''   Set NewTb2 = JetDb!ChDetTable
'''   With NewTb2
'''      Set NewIdx2 = .CreateIndex
'''      With NewIdx2
'''         .Name = "DetIdx"
'''         .Fields.Append .CreateField("DetNum01")
'''      End With
'''      .Indexes.Append NewIdx2
'''   End With
'''
'''End Sub

'Private Sub cmbAct_Click()
'   lbldsc = UpdateActDesc(cmbAct)
'   sCcAcct = cmbAct
'
'   lCheck = GetNextCheck(Compress(cmbAct))
'
'   txtChkNum = lCheck
'
'End Sub

Private Sub cmbAct_LostFocus()
    If Not ValidGLAccount Then
        MsgBox "Invalid GL Account", vbExclamation
        cmbAct = ""
    Else
      lbldsc = UpdateActDesc(cmbAct)
      sCcAcct = cmbAct
      
      lCheck = GetNextCheck(Compress(cmbAct))
      
      txtChkNum = lCheck
   End If
End Sub

Private Sub cmbAct_OLECompleteDrag(Effect As Long)
   lbldsc = UpdateActDesc(cmbAct)
   sCcAcct = cmbAct
   
   lCheck = GetNextCheck(Compress(cmbAct))
   txtChkNum = lCheck
End Sub

'Private Sub cmbFormat_Click()
'   If cmbFormat = "MCS" Then
'      bMaxInvoices = 13
'   Else
'      bMaxInvoices = 15
'   End If
'   If Val(txtInvStub) > bMaxInvoices Then
'      txtInvStub = bMaxInvoices
'   End If
'End Sub
'
Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "No Subject Help"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub Form_Activate()
   Dim b As Byte
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      If Not bReprint Then
         b = GetApAccounts
         If b = 3 Then
            MouseCursor 0
            MsgBox "One Or More Payables Accounts Are Not Active." & vbCr _
               & "Please Set All AP Accounts In The System." & vbCr _
               & "Company Setup, Administration Section.", _
               vbInformation, Caption
            Unload Me
            Exit Sub
         End If
         
         CurrentJournal "CC", ES_SYSDATE, sJournalID
         FillCombo
         bOnLoad = False
      End If
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   Dim rdoTotal As ADODB.Recordset

   FormLoad Me, ES_DONTLIST
   FormatControls
   If bReprint Then
      Me.Caption = "Reprint Computer Checks"
      lblNumChks = UBound(sCheck) + 1
      cmbAct.Visible = False
      lbldsc.Visible = False
      z1(4).Visible = False
      txtChkNum = sCheck(0)
      txtChkNum.enabled = False
      txtDte.enabled = False
   Else
      lblNumChks = NumberOfChecks()
      txtDte = Format(ES_SYSDATE, "mm/dd/yy")
   End If
'   AddComboStr cmbFormat.hWnd, "MCS"
'   AddComboStr cmbFormat.hWnd, "Custom"

    ' get total of all checks
    sSql = "select sum(isnull(ChseTable.CHKAMT,0.00)) from ChseTable"
    If clsADOCon.GetDataSet(sSql, rdoTotal) Then
        lblTotal = FormatCurrency(rdoTotal(0))
    End If
    Set rdoTotal = Nothing

    GetOptions
    bOnLoad = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaAPf03a = Nothing
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub


Private Sub optDis_Click()
    Dim isEnabled As Boolean
    isEnabled = cmbAct.enabled
    cmbAct.enabled = False  ' don't allow account to change during this process
    PrintReport
    cmbAct.enabled = isEnabled
End Sub

Private Sub optPrn_Click()
   
    Dim isEnabled As Boolean
    isEnabled = cmbAct.enabled
    cmbAct.enabled = False  ' don't allow account to change during this process
   
   Dim strChkAct As String
   
   If Val(lblNumChks) < 1 Then
      MsgBox "No Checks In Setup.", vbInformation, Caption
      cmbAct.enabled = isEnabled
      Exit Sub
   End If
   
   strChkAct = CStr(Compress(cmbAct))
   
   If (Trim(strChkAct) = "") Then
      sMsg = "The Check Account is Empty. Please select a valid Check Account."
      MsgBox sMsg, vbInformation, Caption
      cmbAct.SetFocus
      cmbAct.enabled = isEnabled
      Exit Sub
   End If
   
   If Not bReprint Then
      'make sure a check number is specified for new prints
      If Not IsNumeric(txtChkNum) Then
         MsgBox "A valid starting check number is required", vbInformation, Caption
         txtChkNum.SetFocus
          cmbAct.enabled = isEnabled
         Exit Sub
      ElseIf Not ValidateCheck(Trim(txtChkNum), strChkAct) Then
         sMsg = "Starting Check Number Is Already In Use."
         MsgBox sMsg, vbInformation, Caption
         txtChkNum.SetFocus
         cmbAct.enabled = isEnabled
         Exit Sub
      ElseIf Val(lblNumChks) = 0 And optChk.Value = vbChecked Then
         sMsg = "Check Setup Is Empty."
         MsgBox sMsg, vbInformation, Caption
         cmbAct.enabled = isEnabled
         Exit Sub
      End If
      
      sJournalID = GetOpenJournal("CC", Format(txtDte, "mm/dd/yy"))
      If sJournalID = "" Then
         sMsg = "No Open Computer Check Journal" & vbCrLf _
                & "For " & txtDte & " ."
         MsgBox sMsg, vbInformation, Caption
         txtDte.SetFocus
         cmbAct.enabled = isEnabled
         Exit Sub
      End If
   End If
   
   If Not optTst Then
      sMsg = "Is Check Stock In " & lblPrinter & " ?"
      Dim bResponse As Byte
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbNo Then
         cmbAct.enabled = isEnabled
         Exit Sub
      End If
   End If
   
   MouseCursor 13
   
   optPrn.enabled = False
   
'   Select Case cmbFormat
'      Case "MCS"
'         sReportName = "chkMCS"
'      Case "Custom"
'         sReportName = "chkCustom"
'   End Select
   
   'BuildCheckTables
   
   sSql = "delete from "
   
   If optTst Then
      ' Print test check pattern
      MouseCursor 13
      'SetMdiReportsize MdiSect
      'MdiSect.crw.ReportFileName = sReportPath & GetCustomReport("finTest.rpt")
      'SetCrystalAction Me
      MouseCursor 0
   Else
      ' Print the real thing
      
      PrintChecks
   End If
   
   'optPrn.Enabled = True
   
   MouseCursor 0
   cmbAct.enabled = isEnabled
   
End Sub

Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
End Sub

Private Sub txtChkNum_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtChkNum_LostFocus()
   If Not bCancel Then
      txtChkNum = CheckLen(txtChkNum, 12)
      If Not ValidateCheck(txtChkNum, Compress(cmbAct)) Then
         sMsg = "Check Already Exist For Account " & cmbAct & "."
         MsgBox sMsg, vbInformation, Caption
         txtChkNum.SetFocus
      End If
   End If
End Sub

Private Sub FillCombo()
   Dim rdoAct As ADODB.Recordset
   Dim b As Byte
   Dim lCheck As Double
   Dim sOptions As String
   
   MouseCursor 13
   
   On Error GoTo DiaErr1
   
   b = GetApAccounts
   If b = 3 Then
      MouseCursor 0
      MsgBox "One Or More Payables Accounts Are Not Active." & vbCr _
         & "Please Set All AP Accounts In The System." & vbCr _
         & "Company Setup, Administration Section.", _
         vbInformation, Caption
      Unload Me
      Exit Sub
   End If
   
   ' Accounts
   sSql = "SELECT GLACCTNO FROM GlacTable WHERE GLCASH = 1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct)
   
   If bSqlRows Then
      With rdoAct
         While Not .EOF
            AddComboStr cmbAct.hWnd, "" & Trim(!GLACCTNO)
            .MoveNext
         Wend
      End With
      cmbAct.ListIndex = 0
   Else
      ' Multiple checking accounts not setup default to company setup
      cmbAct = sCcAcct
      cmbAct.enabled = False
   End If
   
   'get next check number for account
   lCheck = GetNextCheck(Compress(cmbAct))
   txtChkNum = lCheck

   
   lbldsc = UpdateActDesc(cmbAct)
   Set rdoAct = Nothing
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "FillCombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub txtDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtDte_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtDte_LostFocus()
   txtDte = CheckDate(txtDte)
End Sub

'Private Sub txtInvStub_GotFocus()
'   SelectFormat Me
'End Sub
'
Private Sub SaveOptions()
   Dim sOptions As String
   
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
   
'   sSql = "UPDATE Preferences SET PreChkFormat = '" & cmbFormat _
'          & "',PreChkInvStub = " & Val(txtInvStub)
   
   'sSql = "UPDATE Preferences SET PreChkInvStub = " & Val(lblInvStub)
   'clsAdoCon.ExecuteSQL sSql
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   'Dim rdoFmt As ADODB.RecordSet
   
   'On Error Resume Next
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
   
   'Get invoices per stub
   Dim rdo As ADODB.Recordset
   sSql = "select COLinesPerCheckStub from ComnTable"
   If clsADOCon.GetDataSet(sSql, rdo) Then
      lblInvStub = rdo!COLinesPerCheckStub
   Else
      lblInvStub = 15
   End If
   
'   sSql = "SELECT PreChkFormat,PreChkInvStub FROM Preferences"
'   bSqlRows = clsAdoCon.GetDataSet(sSql,rdoFmt)
'   If bSqlRows Then
'      With rdoFmt
'         If IsNull(.Fields(1)) Or .Fields(1) = 0 Then
'            txtInvStub = "15" ' Default
'         Else
'            txtInvStub = .Fields(1)
'         End If
'      End With
'   End If
'
'   Set rdoFmt = Nothing
End Sub

'Private Sub txtInvStub_LostFocus()
'   txtInvStub = CheckLen(txtInvStub, 2)
'End Sub



Private Function ValidGLAccount() As Boolean
   Dim rdoAct As ADODB.Recordset
   
   sSql = "SELECT GLACCTNO FROM GlacTable WHERE GLCASH = 1 AND GLACCTREF = '" & Compress(cmbAct) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct)
   If bSqlRows And Not rdoAct.EOF Then ValidGLAccount = True Else ValidGLAccount = False
   Set rdoAct = Nothing
End Function

Private Function CompressLF(TestNo As Variant) As String
   Dim PartNo As String
   Dim NewPart As String
   
   On Error GoTo modErr1
   PartNo = Trim$(TestNo)
   If Len(PartNo) > 0 Then
      NewPart = Replace(PartNo, Chr$(10), "")  'lf
      NewPart = Replace(NewPart, Chr$(13), "")  'cr
      NewPart = Replace(NewPart, Chr$(39), "")  'single quote
   End If
   CompressLF = NewPart
   Exit Function
   
modErr1:
   Resume modErr2
modErr2:
   On Error Resume Next
   CompressLF = TestNo
   
End Function
