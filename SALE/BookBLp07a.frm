VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form BookBLp07a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backlog, 12 Month By Division"
   ClientHeight    =   2940
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2940
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox txtMonth 
      Height          =   315
      Left            =   1920
      TabIndex        =   15
      Tag             =   "4"
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "BookBLp07a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbCst 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Contains Customers With Sales Orders"
      Top             =   1560
      Width           =   1555
   End
   Begin VB.ComboBox cmbDiv 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Select Division From List Or Blank For All"
      Top             =   1200
      Width           =   860
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   3
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "BookBLp07a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "BookBLp07a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   2640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2940
      FormDesignWidth =   7260
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   1920
      TabIndex        =   14
      Top             =   2400
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   0
      Left            =   4080
      TabIndex        =   11
      Top             =   1560
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer(s)"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Division(s)"
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblMonth 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6240
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   2
      Left            =   4080
      TabIndex        =   7
      Top             =   1200
      Width           =   1425
   End
   Begin VB.Label cUR 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Month"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1425
   End
End
Attribute VB_Name = "BookBLp07a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'New 11/24/03
'2/25/05 Changed dates and Options
'4/26/05 Increased Customer array
Option Explicit
'Dim rdoQry As rdoQuery
Dim cmdObj As ADODB.Command
Dim bOnLoad As Byte
Dim iTotalCustomers As Integer

Dim dRptDates(14, 2) As Date
Dim sSoDates(14, 4) As String
Dim sCustomers(1300, 3) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FillCombo()
   On Error GoTo DiaErr1
   
   sSql = "Qry_GetCustomerSalesOrder"
   LoadComboBox cmbCst
   'cmbCst = "ALL"
   
   sSql = "SELECT DIVREF FROM CdivTable"
   LoadComboBox cmbCst, -1
   cmbCst = ""
   cmbDiv = "ALL"
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If cmbCst = "" Then cmbCst = "ALL"
   
End Sub


Private Sub cmbDiv_LostFocus()
   If Trim(cmbDiv) = "" Then cmbDiv = "ALL"
   If cmbDiv <> "ALL" Then cmbDiv = CheckLen(cmbDiv, 4)
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 907
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub



Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      CreateTable
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   sSql = "SELECT ITSO,ITSCHED,ITQTY,ITDOLLARS,ITPSNUMBER,ITCANCELED," _
          & "ITSCHED,SONUMBER,SOCUST,SODIVISION FROM SoitTable,SohdTable " _
          & "WHERE (ITSO=SONUMBER AND SOCUST= ? AND ITCANCELED=0 AND " _
          & "ITPSNUMBER='' AND ITPSSHIPPED=0 AND ITINVOICE=0 AND SODIVISION = ? ) " _
          & "AND ITSCHED BETWEEN ? AND ? "
   
   Set cmdObj = New ADODB.Command
   cmdObj.CommandText = sSql
   
   'Set rdoQry = RdoCon.CreateQuery("", sSql)
   Dim prmObj As ADODB.Parameter
   Dim prmObj1 As ADODB.Parameter
   Dim prmObj2 As ADODB.Parameter
   Dim prmObj3 As ADODB.Parameter
   
   Set prmObj = New ADODB.Parameter
   prmObj.Type = adChar
   prmObj.Size = 10
   cmdObj.parameters.Append prmObj

   Set prmObj1 = New ADODB.Parameter
   prmObj1.Type = adChar
   prmObj1.Size = 4
   cmdObj.parameters.Append prmObj1
   
   Set prmObj2 = New ADODB.Parameter
   prmObj2.Type = adDate
   cmdObj.parameters.Append prmObj2
   
   Set prmObj3 = New ADODB.Parameter
   prmObj3.Type = adDate
   cmdObj.parameters.Append prmObj3
   
   bOnLoad = 1
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set cmdObj = Nothing
   
   Set BookBLp07a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sCust As String
   Dim sDiv As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   
   On Error GoTo DiaErr1
   If Trim(Left$(cmbCst, 2)) <> "AL" Then sCust = Compress(cmbCst)
   If Trim(cmbDiv) <> "ALL" Then sDiv = Compress(cmbDiv)
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Customer(s) " & CStr(txtMonth) & "...'")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("slebl07")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport

   sSql = "{EsReportSale07.RPTCUSTOMER} LIKE '" & sCust & "*' " _
          & "AND {EsReportSale07.RPTDIVISION} LIKE '" & sDiv & "*' "
   'Header
   aFormulaName.Add "Period2"
   aFormulaName.Add "Period3"
   aFormulaName.Add "Period4"
   aFormulaName.Add "Period5"
   aFormulaName.Add "Period6"
   aFormulaName.Add "Period7"
   aFormulaName.Add "Period8"
   aFormulaName.Add "Period9"
   aFormulaName.Add "Period10"
   aFormulaName.Add "Period11"
   aFormulaName.Add "Period12"
   aFormulaName.Add "RequestBy"
   
   aFormulaValue.Add CStr("'" & CStr(sSoDates(2, 0) & " " & sSoDates(2, 3)) & "'")
   aFormulaValue.Add CStr("'" & CStr(sSoDates(3, 0) & " " & sSoDates(3, 3)) & "'")
   aFormulaValue.Add CStr("'" & CStr(sSoDates(4, 0) & " " & sSoDates(4, 3)) & "'")
   aFormulaValue.Add CStr("'" & CStr(sSoDates(5, 0) & " " & sSoDates(5, 3)) & "'")
   aFormulaValue.Add CStr("'" & CStr(sSoDates(6, 0) & " " & sSoDates(6, 3)) & "'")
   aFormulaValue.Add CStr("'" & CStr(sSoDates(7, 0) & " " & sSoDates(7, 3)) & "'")
   aFormulaValue.Add CStr("'" & CStr(sSoDates(8, 0) & " " & sSoDates(8, 3)) & "'")
   aFormulaValue.Add CStr("'" & CStr(sSoDates(9, 0) & " " & sSoDates(9, 3)) & "'")
   aFormulaValue.Add CStr("'" & CStr(sSoDates(10, 0) & " " & sSoDates(10, 3)) & "'")
   aFormulaValue.Add CStr("'" & CStr(sSoDates(11, 0) & " " & sSoDates(11, 3)) & "'")
   aFormulaValue.Add CStr("'" & CStr(sSoDates(12, 0) & " " & sSoDates(12, 3)) & "'")
   aFormulaValue.Add CStr("'Requested By: " & sInitials & "'")
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

   
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
 
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
   ' MM lblMonth = Format(ES_SYSDATE, "mmm yyyy")
   txtMonth = Format(ES_SYSDATE, "mmm yyyy")
End Sub

Private Sub SaveOptions()
   
End Sub

Private Sub GetOptions()
   
End Sub



Private Sub optDis_Click()
   BuildReport
   
End Sub


Private Sub optPrn_Click()
   BuildReport
   
End Sub



Private Sub GetDates()
   Dim b As Byte
   Dim iNext As Integer
   Dim iYear As Integer
   Dim iStatus As Integer
   Dim dDate As Date
   Dim sDate As String
   
   Erase sSoDates
   ' 11/13/2009 changed to selected date
   'MM sDate = Format(ES_SYSDATE, "mm/dd/yyyy")
   sDate = Format(txtMonth, "mm/dd/yyyy")
   sDate = Left(sDate, 3) & "01" & Right(sDate, 5)
   dDate = sDate
   iStatus = 20
   prg1.Value = iStatus
   iYear = Val(Right$(sDate, 4))
   sSoDates(0, 0) = Format$(dDate - 1, "mmm")
   sSoDates(0, 1) = Format$(1)
   'sSoDates(0, 3) = Format$(iYear)
   sSoDates(0, 3) = Format$(DateAdd("m", -1, dDate), "yyyy")
   sSoDates(1, 0) = Format$(dDate, "mmm")
   sSoDates(1, 1) = Format$(1)
   sSoDates(1, 3) = Format$(iYear)
   For b = 2 To 12
      iStatus = iStatus + 5
      prg1.Value = iStatus
      iNext = iNext + 32
      sSoDates(b, 0) = Format(dDate + iNext, "mmm")
      If Format(dDate + iNext, "mmm") = "Jan" Then iYear = iYear + 1
      sSoDates(b, 1) = Format$(1)
      sSoDates(b, 3) = Format$(iYear)
   Next
   For b = 0 To 12
      iStatus = iStatus + 5
      If iStatus > 90 Then iStatus = 90
      prg1.Value = iStatus
      ' MM Initilize the  Year - to get the number of days.
      iYear = sSoDates(b, 3)
      Select Case sSoDates(b, 0)
         Case "Jan", "Mar", "May", "Jul", "Aug", "Oct", "Dec"
            sSoDates(b, 2) = Format$(31)
         Case "Apr", "Jun", "Sep", "Nov"
            sSoDates(b, 2) = Format$(30)
         Case Else
            If iYear = 2004 Or iYear = 2008 Or iYear = 2012 Or iYear = 2016 _
                       Or iYear = 2020 Or iYear = 2024 Or iYear = 2028 Then
               sSoDates(b, 2) = Format$(29)
            Else
               sSoDates(b, 2) = Format$(28)
            End If
      End Select
      dRptDates(b, 0) = Format(sSoDates(b, 0) & " " & sSoDates(b, 1) _
                & " " & sSoDates(b, 3), "mm/dd/yy hh:nn")
      dRptDates(b, 1) = Format(sSoDates(b, 0) & " " & sSoDates(b, 2) _
                & " " & sSoDates(b, 3), "mm/dd/yy hh:nn")
   Next
   sDate = "01/01/" & Trim(str(iYear - 5))
   dRptDates(0, 0) = Format(sDate, "mm/dd/yy")
   prg1.Value = 100
   Sleep 500
   
End Sub

Private Sub GetCustomers()
   Dim RdoCst As ADODB.Recordset
   Dim iStatus As Integer
   
   Erase sCustomers
   iTotalCustomers = 0
   sSql = "SELECT DISTINCT SOCUST,SODIVISION,CUREF,CUNICKNAME " _
          & "FROM SohdTable,CustTable WHERE SOCUST=CUREF ORDER BY SOCUST"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCst, ES_FORWARD)
   If bSqlRows Then
      With RdoCst
         Do Until .EOF
            iStatus = iStatus + 10
            If iStatus > 90 Then iStatus = 90
            prg1.Value = iStatus
            iTotalCustomers = iTotalCustomers + 1
            sCustomers(iTotalCustomers, 0) = "" & Trim(!SOCUST)
            sCustomers(iTotalCustomers, 1) = "" & Trim(!SODIVISION)
            sCustomers(iTotalCustomers, 2) = "" & Trim(!CUNICKNAME)
            sSql = "INSERT INTO EsReportSale07 (" _
                   & "RPTCUSTOMER,RPTDIVISION,RPTNICKNAME) VALUES('" _
                   & Trim(!SOCUST) & "','" _
                   & Trim(!SODIVISION) & "','" _
                   & Trim(!CUNICKNAME) & "')"
            clsADOCon.ExecuteSql sSql ' rdExecDirect
            .MoveNext
         Loop
         ClearResultSet RdoCst
      End With
   End If
   Set RdoCst = Nothing
   
   prg1.Value = 100
   Sleep 500
   
End Sub

Private Sub BuildReport()
   sSql = "truncate table dbo.EsReportSale07"
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   prg1.Visible = True
   prg1.Value = 10
   lblStatus = "Getting Requirements"
   lblStatus.Visible = True
   lblStatus.Refresh
   
   On Error GoTo DiaErr1
   MouseCursor 11
   sProcName = "getcustomers"
   GetCustomers
   If iTotalCustomers > 0 Then
      prg1.Value = 10
      lblStatus = "Setting Parameters"
      lblStatus.Refresh
      sProcName = "getdates"
      GetDates
   Else
      MouseCursor 0
      MsgBox "There Are No Sales Orders That Meet The Criteria.", _
         vbInformation, Caption
      prg1.Visible = False
      lblStatus.Visible = False
      Exit Sub
   End If
   prg1.Value = 10
   lblStatus = "Building Report"
   lblStatus.Refresh
   sProcName = "getsalesord"
   GetSalesOrders
   PrintReport
   prg1.Visible = False
   lblStatus.Visible = False
   Exit Sub
   
DiaErr1:
   MouseCursor 0
   prg1.Visible = False
   lblStatus.Visible = False
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub CreateTable()
'   On Error Resume Next
'   sSql = "SELECT RPTRECORD FROM EsReportSale07"
'   clsADOCon.ExecuteSql sSql 'rdExecDirect
'   If Err = 40002 Then
'      On Error GoTo 0
'      Err = 0
'      clsADOCon.ADOErrNum = 0
   
   Dim rs As ADODB.Recordset
   sSql = "SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'EsReportSale07'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rs, ES_FORWARD)
   If Not bSqlRows Then
      sSql = "Create Table EsReportSale07 (" _
             & "RPTCUSTOMER CHAR(10) NULL DEFAULT('')," _
             & "RPTDIVISION CHAR(2) NULL DEFAULT('')," _
             & "RPTNICKNAME CHAR(10) NULL DEFAULT('')," _
             & "RPTPERIOD0 MONEY NULL DEFAULT(0)," _
             & "RPTPERIOD1 MONEY NULL DEFAULT(0)," _
             & "RPTPERIOD2 MONEY NULL DEFAULT(0)," _
             & "RPTPERIOD3 MONEY NULL DEFAULT(0)," _
             & "RPTPERIOD4 MONEY NULL DEFAULT(0)," _
             & "RPTPERIOD5 MONEY NULL DEFAULT(0)," _
             & "RPTPERIOD6 MONEY NULL DEFAULT(0)," _
             & "RPTPERIOD7 MONEY NULL DEFAULT(0)," _
             & "RPTPERIOD8 MONEY NULL DEFAULT(0)," _
             & "RPTPERIOD9 MONEY NULL DEFAULT(0)," _
             & "RPTPERIOD10 MONEY NULL DEFAULT(0)," _
             & "RPTPERIOD11 MONEY NULL DEFAULT(0)," _
             & "RPTPERIOD12 MONEY NULL DEFAULT(0))"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      If clsADOCon.ADOErrNum = 0 Then
         sSql = "CREATE UNIQUE CLUSTERED INDEX SalesRef ON " _
                & "dbo.EsReportSale07 (RPTCUSTOMER,RPTDIVISION) WITH FILLFACTOR = 80 "
         clsADOCon.ExecuteSql sSql 'rdExecDirect
      End If
      
   End If
   Set rs = Nothing
   
End Sub

Private Sub GetSalesOrders()
   Dim RdoSit As ADODB.Recordset
   Dim b As Byte
   Dim iRow As Integer
   Dim iStatus As Integer
   Dim cPeriodTotal As Currency
   Dim cLineTotal As Currency
   
   iStatus = prg1.Value
   For iRow = 1 To iTotalCustomers
      cLineTotal = 0
      For b = 0 To 12
         'rdoQry(0) = sCustomers(iRow, 0)
         'rdoQry(1) = sCustomers(iRow, 1)
         'rdoQry(2) = Format(dRptDates(b, 0), "mm/dd/yy")
         'rdoQry(3) = Format(dRptDates(b, 1), "mm/dd/yy 23:59")
         'bSqlRows = GetQuerySet(RdoSit, rdoQry)
        cmdObj.parameters(0).Value = sCustomers(iRow, 0)
        cmdObj.parameters(1).Value = sCustomers(iRow, 1)
        cmdObj.parameters(2).Value = Format(dRptDates(b, 0), "mm/dd/yy")
        cmdObj.parameters(3).Value = Format(dRptDates(b, 1), "mm/dd/yy 23:59")
        bSqlRows = clsADOCon.GetQuerySet(RdoSit, cmdObj, ES_FORWARD, True)
         If bSqlRows Then
            cPeriodTotal = 0
            With RdoSit
               Do Until .EOF
                  cPeriodTotal = cPeriodTotal + (!ITQty * !ITDOLLARS)
                  cLineTotal = cLineTotal + cPeriodTotal
                  .MoveNext
               Loop
               ClearResultSet RdoSit
               sSql = "UPDATE EsReportSale07 SET RPTPERIOD" & Trim(str(b)) & "=" & cPeriodTotal & " " _
                      & "WHERE RPTCUSTOMER='" & sCustomers(iRow, 0) & "' AND " _
                      & "RPTDIVISION='" & sCustomers(iRow, 1) & "'"
               clsADOCon.ExecuteSql sSql 'rdExecDirect
            End With
         End If
      Next
      If cLineTotal = 0 Then
         sSql = "Delete From EsReportSale07 WHERE RPTCUSTOMER='" & sCustomers(iRow, 0) & "' " _
                & "AND RPTDIVISION='" & sCustomers(iRow, 1) & "'"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
      Else
         iStatus = iStatus + 5
         If iStatus > 95 Then iStatus = 95
         prg1.Value = iStatus
      End If
   Next
   Set RdoSit = Nothing
   prg1.Value = 100
   Sleep 500
End Sub

Private Sub txtMonth_Change()
   txtMonth = Format(txtMonth, "mmm yyyy")
End Sub

Private Sub txtMonth_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtMonth_LostFocus()
   txtMonth = Format(txtMonth, "mmm yyyy")
End Sub

