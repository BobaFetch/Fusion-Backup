VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form BookBLp06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backlog By Part Number, Current Month"
   ClientHeight    =   2850
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2850
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   2040
      TabIndex        =   23
      Top             =   1080
      Width           =   2895
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "BookBLp06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Tag             =   "3"
      Top             =   1080
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.ComboBox cmbCde 
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   5400
      Sorted          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "3"
      ToolTipText     =   "Select Product Code From List"
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cmbCls 
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   5400
      Sorted          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Tag             =   "3"
      ToolTipText     =   "Select Product Class From List"
      Top             =   3120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   600
      TabIndex        =   15
      Tag             =   "4"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox optVew 
      Caption         =   "View"
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "BookBLp06a.frx":07AE
      Height          =   315
      Left            =   5040
      Picture         =   "BookBLp06a.frx":0AF0
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   1080
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.CheckBox optDet 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   4
      Top             =   1920
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   6
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "BookBLp06a.frx":0E32
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "BookBLp06a.frx":0FB0
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   3120
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2850
      FormDesignWidth =   7260
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   2040
      TabIndex        =   22
      Top             =   2280
      Width           =   4932
      _ExtentX        =   8705
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   2280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(End Of Period)"
      Height          =   285
      Index           =   6
      Left            =   3480
      TabIndex        =   19
      Top             =   1440
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Class"
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   3600
      TabIndex        =   18
      Top             =   3120
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   3600
      TabIndex        =   17
      Top             =   2880
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cut Off Date"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   16
      ToolTipText     =   "Contains Part Numbers Sales Order Items (Leading Chars Or Blank For All"
      Top             =   1440
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Week Starts On"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   14
      Tag             =   " "
      ToolTipText     =   "The Day That Your Week Starts In The Company Setup"
      Top             =   600
      Width           =   2145
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Prices"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   2
      Left            =   5640
      TabIndex        =   10
      Top             =   1080
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Numbers(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Contains Part Numbers Sales Order Items (Leading Chars Or Blank For All"
      Top             =   1080
      Width           =   1425
   End
End
Attribute VB_Name = "BookBLp06a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'New 11/14/03
'2/25/05 Changed dates and Options
'4/26/05 Added Null Date Trap GetDates
'        removed Combo, increased iRecord to long
'11/8/05 Added an additional Table In Use trap
Option Explicit
'Dim rdoQry As rdoQuery
Dim cmdObj As ADODB.Command
Dim bOnLoad As Byte
Dim bInUse As Byte
Dim sWkStarts As String
Dim sPeriodEnd As String

Dim dItemDates(6, 2) As Date
' 0 = Start Date
' 1 = End Date

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'Local Errors

Function GetStartDay() As String
   On Error Resume Next
   Dim RdoStr As ADODB.Recordset
   sSql = "SELECT WEEKENDS FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoStr, ES_FORWARD)
   If bSqlRows Then
      If IsNull(RdoStr.Fields(0)) Then _
                GetStartDay = "Sat" Else GetStartDay = RdoStr.Fields(0)
      ClearResultSet RdoStr
   End If
   If GetStartDay = "Sat" Then GetStartDay = "Sunday" Else _
                    GetStartDay = "Monday"
   sWkStarts = Left$(GetStartDay, 3)
   Set RdoStr = Nothing
   
End Function

Private Sub FillCombo()
   '    On Error GoTo DiaErr1
   '    sSql = "SELECT DISTINCT PARTREF,PARTNUM,ITPART FROM " _
   '        & "PartTable,SoitTable WHERE (PARTREF=ITPART AND " _
   '        & "ITACTUAL IS NULL AND ITCANCELED=0) ORDER BY PARTREF"
   '    LoadComboBox cmbPrt
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub cmbCde_LostFocus()
   cmbCde = CheckLen(cmbCde, 6)
   If Len(cmbCde) = 0 Then cmbCde = "ALL"
   
End Sub


Private Sub cmbCls_LostFocus()
   cmbCls = CheckLen(cmbCls, 4)
   If Len(cmbCls) = 0 Then cmbCls = "ALL"
   
End Sub


Private Sub cmbPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "CMBPRT"
      ViewParts.txtPrt = cmbPrt
      optVew.Value = vbChecked
      ViewParts.Show
   End If
   
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If cmbPrt = "" Then cmbPrt = "ALL"
   
End Sub

Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "TXTPRT"
      ViewParts.txtPrt = txtPrt
      optVew.Value = vbChecked
      ViewParts.Show
   End If
   
End Sub


Private Sub txtPrt_LostFocus()
   cmbPrt = txtPrt
   cmbPrt_LostFocus
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdFnd_Click()
   If txtPrt.Visible Then
      cmbPrt = txtPrt
      ViewParts.lblControl = "TXTPRT"
   Else
      ViewParts.lblControl = "CMBPRT"
   End If
   ViewParts.txtPrt = cmbPrt
   ViewParts.Show
   
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
      z1(4) = z1(4) & " " & GetStartDay()
      CreateTable
      FillProductCodes
      FillProductClasses
      
      Dim bPartSearch As Boolean
      
      bPartSearch = GetPartSearchOption
      SetPartSearchOption (bPartSearch)
      
      If (Not bPartSearch) Then FillPartCombo cmbPrt
      
      cmbPrt = "ALL"
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   
   sSql = "SELECT ITSO,ITNUMBER,ITREV,ITPART,ITQTY,ITDOLLARS,ITSCHED," _
          & "ITPSNUMBER,ITCANCELED,ITPSSHIPPED,SONUMBER,SOTYPE,SOCUST,SOTEXT," _
          & "CUREF,CUNICKNAME,PARTREF,PARTNUM,PADESC From SoitTable," _
          & "SohdTable, CustTable, PartTable WHERE (ITSO=SONUMBER AND " _
          & "ITCANCELED=0 AND ITPSNUMBER='' AND ITINVOICE=0 " _
          & "AND ITPSSHIPPED=0) AND ITPART=PARTREF AND ITSCHED " _
          & "BETWEEN ? AND ? ORDER BY ITPART,ITSCHED"
   
   Set cmdObj = New ADODB.Command
   cmdObj.CommandText = sSql
   Dim prmObj As ADODB.Parameter
   Set prmObj = New ADODB.Parameter
   prmObj.Type = adDate
   
   cmdObj.parameters.Append prmObj
   
   Dim prmObj1 As ADODB.Parameter
   Set prmObj1 = New ADODB.Parameter
   prmObj1.Type = adDate
   cmdObj.parameters.Append prmObj1
   'Set rdoQry = RdoCon.CreateQuery("", sSql)
   bOnLoad = 1
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   On Error Resume Next
   sSql = "TRUNCATE TABLE EsReportSale06"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   Set cmdObj = Nothing
   Set BookBLp06a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sParts As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   If bInUse = 1 Then Exit Sub
   
   MouseCursor 13
   On Error GoTo DiaErr1
   If cmbPrt = "" Then cmbPrt = "ALL"
   If cmbPrt <> "ALL" Then sParts = Compress(cmbPrt)
   lblStatus.Visible = False
   prg1.Visible = False
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "Week1"
   aFormulaName.Add "Week2"
   aFormulaName.Add "Week3"
   aFormulaName.Add "Week4"
   aFormulaName.Add "Week5"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowDetails"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & CStr(cmbPrt) & "...'")
   aFormulaValue.Add CStr("'PD To " & CStr(Format(dItemDates(0, 1), "mm/dd/yy")) & "'")
   aFormulaValue.Add CStr("'To" & CStr(Format(dItemDates(1, 1), "mm/dd/yy")) & "'")
   aFormulaValue.Add CStr("'To" & CStr(Format(dItemDates(2, 1), "mm/dd/yy")) & "'")
   aFormulaValue.Add CStr("'To" & CStr(Format(dItemDates(3, 1), "mm/dd/yy")) & "'")
   aFormulaValue.Add CStr("'Beyond To " & CStr(Format(dItemDates(4, 1), "mm/dd/yy")) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add optDet.Value
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("slebl06")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{EsReportSale06.RPTPARTREF} LIKE '" & sParts & "*' "
   
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
 
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
   MouseCursor 0
   Sleep 5000
   On Error Resume Next
   'sSql = "TRUNCATE TABLE EsReportSale06"
   'clsADOCon.ExecuteSQL sSql ' rdExecDirect
   Exit Sub
   
DiaErr1:
   sSql = "TRUNCATE TABLE EsReportSale06"
   clsADOCon.ExecuteSQL sSql ' rdExecDirect
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   Dim dCutOff As Date
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   cmbPrt = "ALL"
   txtBeg = Format(ES_SYSDATE, "mm/dd/yy")
   dCutOff = Format(ES_SYSDATE + 42, "mm/dd/yy")
   txtEnd = Format(dCutOff, "mm/dd/yy")
   sPeriodEnd = txtEnd
   txtEnd.ToolTipText = "Last Group Is Limited To This Date (Allow At Least 6 Weeks)"
   cmbPrt = "ALL"
   
End Sub

Private Sub SaveOptions()
   
End Sub

Private Sub GetOptions()
   
End Sub

Private Sub optDet_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   GetDates
   
End Sub


Private Sub optPrn_Click()
   GetDates
   
End Sub


Private Sub GetDates()
   Dim RdoItm As ADODB.Recordset
   Dim b As Integer
   Dim lRecord As Long
   Dim dDate As Date
   
   Dim cPrice As Currency
   Dim sBegMonth As String
   Dim sStartDate As String
   Dim sWkEnds As String
   Dim sSalesOrder As String
   
   Dim vNullDate As Variant
   optPrn.Enabled = False
   optDis.Enabled = False
   sSql = "SELECT RPTRECORD FROM EsReportSale06 WHERE RPTRECORD=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoItm, ES_FORWARD)
   If bSqlRows Then
      MsgBox "The Report Table Is In Use. Try Again In A Minute.", _
         vbInformation, Caption
      bInUse = 1
      Exit Sub
   Else
      bInUse = 0
   End If
   MouseCursor 13
   lRecord = 0
   lblStatus.Visible = True
   prg1.Visible = True
   lblStatus.Refresh
   
   On Error GoTo DiaErr1
   sSql = "TRUNCATE TABLE EsReportSale06"
   clsADOCon.ExecuteSQL sSql ' rdExecDirect
   
   sBegMonth = Format(txtBeg, "mmm") & "-" & "20" & Right(txtBeg, 2)
   sStartDate = Format(txtBeg, "ddd")
   dDate = Format(txtBeg, "mm/dd/yy 00:00")
   
   b = Format(dDate, "w")
   If sWkStarts = "Mon" Then
      sWkEnds = "Sun"
      b = 8 - b
   Else
      sWkEnds = "Sat"
      b = 9 - b
   End If
   dDate = Format(dDate + b, "mm/dd/yy")
   dItemDates(0, 0) = Format(dDate - 1800, "mm/dd/yy 00:00")
   dItemDates(0, 1) = Format(dDate, "mm/dd/yy 23:59")
   For b = 1 To 3
      dItemDates(b, 0) = Format(dItemDates(b - 1, 1) + 1, "mm/dd/yy 00:00")
      dItemDates(b, 1) = dItemDates(b - 1, 1) + 7
   Next
   dItemDates(b, 0) = Format(dItemDates(b - 1, 1) + 1, "mm/dd/yy 00:00")
   dItemDates(b, 1) = Format(txtEnd, "mm/dd/yy")
   prg1.Value = 20
   lRecord = 0
   For b = 0 To 4
      prg1.Value = prg1.Value + 10
      'rdoQry(0) = dItemDates(b, 0)
      'rdoQry(1) = dItemDates(b, 1)
       cmdObj.parameters(0).Value = dItemDates(b, 0)
       cmdObj.parameters(1).Value = dItemDates(b, 1)
      
      
      bSqlRows = clsADOCon.GetQuerySet(RdoItm, cmdObj, ES_FORWARD, True)
      If bSqlRows Then
         With RdoItm
            Do Until .EOF
               Err.Clear
               If Not IsNull(!itsched) Then
                  lRecord = lRecord + 1
                  If optDet.Value = vbChecked Then cPrice = !ITDOLLARS _
                                    Else cPrice = 0
                  sSalesOrder = Trim(!SOTYPE) & Format(!itso, SO_NUM_FORMAT) _
                                & "-" & Trim(str(!ITNUMBER)) & !itrev & "Qty: " & Format$(!ITQty, ES_QuantityDataFormat)
                  sSql = "INSERT INTO EsReportSale06 (RPTRECORD,RPTPARTREF,RPTPARTNUM," _
                         & "RPTPARTDESC," _
                         & "RPTCUSTOMER" & Trim(str(b + 1)) & "," _
                         & "RPTSALESORDER" & Trim(str(b + 1)) & "," _
                         & "RPTQUANTITY" & Trim(str(b + 1)) & "," _
                         & "RPTPRICE" & Trim(str(b + 1)) & ") " _
                         & "VALUES(" & lRecord & ",'" & Trim(!PartRef) & "','" _
                         & Trim(!PartNum) & "','" & Trim(!PADESC) & "','" _
                         & Format(!itsched, "mm/dd/yy") & " " & Trim(!CUNICKNAME) & "','" _
                         & sSalesOrder & "'," _
                         & Format(!ITQty, ES_QuantityDataFormat) & "," & Format(cPrice, ES_QuantityDataFormat) & ")"
'Debug.Print sSql
                 clsADOCon.ExecuteSQL sSql ' rdExecDirect
               End If
               .MoveNext
               DoEvents
            Loop
            ClearResultSet RdoItm
         End With
      End If
   Next
   prg1.Value = 100
   Sleep 1000
   PrintReport
   
   Exit Sub
   
DiaErr1:
   optPrn.Enabled = True
   optDis.Enabled = True
   lblStatus.Visible = False
   prg1.Visible = False
   If Left(CurrError.Description, 5) = "01000" Then
      MsgBox "The Report Table Is in Use. Try Again in Shortly..", _
         vbInformation, Caption
   Else
      sProcName = "GetDates"
      CurrError.Number = Err.Number
      CurrError.Description = Err.Description
      DoModuleErrors Me
   End If
End Sub

Private Sub txtend_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtEnd_LostFocus()
   txtEnd = CheckDate(txtEnd)
   If Format(txtEnd, "yyyy,mm,dd") < Format(sPeriodEnd, "yyyy,mm,dd") Then
      Beep
      txtEnd = sPeriodEnd
   End If
   
End Sub



Private Sub CreateTable()
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   sSql = "SELECT RPTRECORD FROM EsReportSale06"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   If clsADOCon.ADOErrNum = 40002 Then
      clsADOCon.ADOErrNum = 0
      sSql = "Create Table EsReportSale06 (" _
             & "RPTRECORD INT NULL DEFAULT(1)," _
             & "RPTPARTREF CHAR(30) NULL DEFAULT('')," _
             & "RPTPARTNUM CHAR(30) NULL DEFAULT('')," _
             & "RPTPARTDESC CHAR(30) NULL DEFAULT('')," _
             & "RPTCUSTOMER1 CHAR(20) NULL DEFAULT('')," _
             & "RPTSALESORDER1 CHAR(30) NULL DEFAULT('')," _
             & "RPTQUANTITY1 REAL NULL DEFAULT(0)," _
             & "RPTPRICE1 REAL NULL DEFAULT(0)," _
             & "RPTCUSTOMER2 CHAR(20) NULL DEFAULT('')," _
             & "RPTSALESORDER2 CHAR(30) NULL DEFAULT('')," _
             & "RPTQUANTITY2 REAL NULL DEFAULT(0)," _
             & "RPTPRICE2 REAL NULL DEFAULT(0)," _
             & "RPTCUSTOMER3 CHAR(20) NULL DEFAULT('')," _
             & "RPTSALESORDER3 CHAR(30) NULL DEFAULT('')," _
             & "RPTQUANTITY3 REAL NULL DEFAULT(0)," _
             & "RPTPRICE3 REAL NULL DEFAULT(0)," _
             & "RPTCUSTOMER4 CHAR(20) NULL DEFAULT('')," _
             & "RPTSALESORDER4 CHAR(30) NULL DEFAULT('')," _
             & "RPTQUANTITY4 REAL NULL DEFAULT(0)," _
             & "RPTPRICE4 REAL NULL DEFAULT(0)," _
             & "RPTCUSTOMER5 CHAR(20) NULL DEFAULT('')," _
             & "RPTSALESORDER5 CHAR(30) NULL DEFAULT('')," _
             & "RPTQUANTITY5 REAL NULL DEFAULT(0)," _
             & "RPTPRICE5 REAL NULL DEFAULT(0))"
      clsADOCon.ExecuteSQL sSql ' rdExecDirect
      If clsADOCon.ADOErrNum = 0 Then
         sSql = "CREATE UNIQUE CLUSTERED INDEX ReportRef ON dbo.EsReportSale06 (RPTRECORD) WITH  FILLFACTOR = 80"
         clsADOCon.ExecuteSQL sSql ' rdExecDirect
         
         sSql = "CREATE INDEX PartRef ON dbo.EsReportSale06 (RPTPARTREF) WITH  FILLFACTOR = 80"
         clsADOCon.ExecuteSQL sSql ' rdExecDirect
      End If
      clsADOCon.ADOErrNum = 0
   End If
   
End Sub

Function SetPartSearchOption(bPartSearch As Boolean)
   
   If (bPartSearch = True) Then
      cmbPrt.Visible = False
      txtPrt.Visible = True
      cmdFnd.Visible = True
   Else
      cmbPrt.Visible = True
      txtPrt.Visible = False
      cmdFnd.Visible = False
   End If
End Function

