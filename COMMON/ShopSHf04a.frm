VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ShopSHf04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Close A Manufacturing Order"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHf04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton optPrn 
      Height          =   330
      Left            =   5160
      Picture         =   "ShopSHf04a.frx":07AE
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Print The Report"
      Top             =   3360
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkIgnoreExpendables 
      Height          =   255
      Left            =   3000
      TabIndex        =   22
      ToolTipText     =   "Workstation Setting - Allow To Close Without Testing Expendables (Type 5) "
      Top             =   3120
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox chkIgnoreUnpicked 
      Height          =   255
      Left            =   3000
      TabIndex        =   23
      ToolTipText     =   "Workstation Setting - Allow To Close With Unpicked Items"
      Top             =   2880
      Width           =   495
   End
   Begin VB.CheckBox chkInvoices 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3360
      TabIndex        =   21
      Top             =   320
      Width           =   495
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "ShopSHf04a.frx":0938
      Height          =   350
      Left            =   6480
      Picture         =   "ShopSHf04a.frx":0E12
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "View Last Closed Run Log (Requires A Text Viewer) "
      Top             =   1920
      Width           =   360
   End
   Begin VB.ComboBox cmbCloseDate 
      Height          =   315
      Left            =   4200
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1575
      Width           =   1215
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Qualified Part Numbers (CO)"
      Top             =   720
      Width           =   3545
   End
   Begin VB.ComboBox cmbRun 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Contains Qualified Runs"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdCloseMO 
      Caption         =   "M&O Close"
      Height          =   315
      Left            =   6240
      TabIndex        =   3
      ToolTipText     =   " Close this Manufacturing Order"
      Top             =   1560
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6600
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5880
      Top             =   3360
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3525
      FormDesignWidth =   7665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity Completed"
      Height          =   252
      Index           =   10
      Left            =   2760
      TabIndex        =   29
      ToolTipText     =   "Total Quantity Completions"
      Top             =   2400
      Width           =   1692
   End
   Begin VB.Label lblComplete 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   4440
      TabIndex        =   28
      ToolTipText     =   "Total Quantity Completions"
      Top             =   2400
      Width           =   852
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ignore Part Type 5's (Expendables)"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   9
      Left            =   240
      TabIndex        =   25
      ToolTipText     =   "Workstation Setting - Allow To Close Without Testing Expendables (Type 5) "
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ignore Unpicked Items"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   8
      Left            =   240
      TabIndex        =   24
      ToolTipText     =   "Workstation Setting - Allow To Close With Unpicked Items"
      Top             =   2880
      Width           =   2715
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Verify Invoicing (PO) Before Closing MO"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   74
      Left            =   240
      TabIndex        =   20
      ToolTipText     =   "Test Allocated PO Items For Invoices (System Setting)"
      Top             =   315
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type"
      Height          =   255
      Index           =   7
      Left            =   5040
      TabIndex        =   18
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   17
      ToolTipText     =   "Manufacturing Order Yield"
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   16
      ToolTipText     =   "Manufacturing Order Yield"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   252
      Index           =   5
      Left            =   5040
      TabIndex        =   15
      Top             =   1080
      Width           =   1452
   End
   Begin VB.Label lblCode 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   6240
      TabIndex        =   14
      Top             =   1080
      Width           =   852
   End
   Begin VB.Label lblLvl 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   6240
      TabIndex        =   13
      ToolTipText     =   "Part Type"
      Top             =   720
      Width           =   492
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Closed"
      Height          =   255
      Index           =   4
      Left            =   2760
      TabIndex        =   12
      Top             =   1575
      Width           =   1095
   End
   Begin VB.Label lblDte 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4320
      TabIndex        =   11
      ToolTipText     =   "Last Quantity Completion Date"
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Completed"
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   10
      ToolTipText     =   "Last Quantity Completion Date"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblStat 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
End
Attribute VB_Name = "ShopSHf04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'3/26/00 Accounts
'1/2/03 Returned the code to trap the Close/Completion Dates
'3/18/03 Added updating LohdTable (lot costs)
'11/11/04 All areas revised and Log added
'
'*** 11/11/04 EMail to Larry/Nathan telling them to check the code and test it
'*** 11/29/04 Telecon Larry.  He has ignored the Email. Re-iterated the necessity
'*** 12/16/04 Tested at JEVCO and it is okay except won't close some
'12/17/04 Reset bCantClose flag for ensuing MO's
'1/6/05 Changed erroneous references to cmbPrt/cmbRun
'*** 1/6/05 apparently it has not been tested yet (see above)
'2/22/05 Reacted to a fax from JEVCO as a result of telecon with Larry
'     It is obvious that Larry hasn't tested the function.
'3/15/05 Added Unpicked switch (AWI)
'5/24/05 Added option to ignore Part Type 5 parts
'7/22/05 Corrected potential for creating second Activity row
'8/12/05 Changed Log to Crystal Reports
'9/9/05 Properly Writes log when not closing
'7/11/06 See GetExpenseCosts
'9/14/06 Added Returns to GetMaterialCosts
Option Explicit
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter

Dim bCantClose As Byte
Dim bOnLoad As Byte
Dim bGoodPrt As Byte
Dim bGoodRun As Byte
Dim bLotsOn As Byte

Dim iLogNumber As Integer
Dim iTotalPicks As Integer
Dim lRunno As Long
Dim sPartNumber As String

Dim cYield As Currency
Dim cRunExp As Currency
Dim cRunHours As Currency
Dim cRunLabor As Currency
Dim cRunMatl As Currency
Dim cRunOvHd As Currency
Dim cStdCost As Currency

Dim sLotNumber As String
Dim sLogNote(600, 2)
Private Const NOTE_Number = 0
Private Const NOTE_Description = 1

Dim sPartLots(100, 3) As String
'0=Part Number
'1=Lots 0/1
'2=Standard Cost
'3=Quantity Picked
Private Const PART_Number = 0
Private Const PART_IsLotTracked = 1
Private Const PART_StdCost = 2
Private Const PART_QtyPicked = 3

'WIP
Dim sInvLabAcct As String
Dim sInvMatAcct As String
Dim sInvExpAcct As String
Dim sInvOhdAcct As String
Dim sCgsAcct As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd
'9/19/06

Public Sub GetPreviousCompletions()
   Dim RdoPrev As ADODB.Recordset

   Dim cQtyIn As Currency 'Type 6
   Dim cQtyOut As Currency 'Type 38
   Dim cQtyBal As Currency 'Total Completed

   'Complete
   On Error GoTo DiaErr1
   MouseCursor 13
   sSql = "SELECT SUM(INAQTY) AS QtyComplete FROM InvaTable WHERE (INTYPE=6 " _
          & "AND INMOPART='" & Compress(cmbPrt) & "' AND INMORUN=" & Val(cmbRun) & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrev, ES_FORWARD)
   If bSqlRows Then
      With RdoPrev
         If Not IsNull(!QtyComplete) Then
            cQtyIn = !QtyComplete
         Else
            cQtyIn = 0
         End If
         ClearResultSet RdoPrev
      End With
   End If
   Set RdoPrev = Nothing
   
   If cQtyIn > 0 Then
      sSql = "SELECT SUM(INAQTY) AS QtyComplete FROM InvaTable WHERE (INTYPE=38 " _
             & "AND INMOPART='" & Compress(cmbPrt) & "' AND INMORUN=" & Val(cmbRun) & ")"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrev, ES_FORWARD)
      If bSqlRows Then
         With RdoPrev
            If Not IsNull(!QtyComplete) Then
               cQtyOut = !QtyComplete
            Else
               cQtyOut = 0
            End If
            ClearResultSet RdoPrev
         End With
      End If
   End If
   Set RdoPrev = Nothing
   cQtyBal = cQtyIn - Abs(cQtyOut)
   If cQtyBal < 0 Then cQtyBal = 0
   sSql = "UPDATE RunsTable SET RUNPARTIALQTY=" & cQtyBal & " WHERE (RUNREF='" _
          & sPartNumber & "' AND RUNNO=" & Val(cmbRun) & ")"
   clsADOCon.ExecuteSQL sSql
   lblComplete = Format(cQtyIn, ES_QuantityDataFormat)
   GoTo DiaErr2
   Exit Sub

DiaErr1:
   Resume DiaErr2
DiaErr2:
   MouseCursor 0
   Set RdoPrev = Nothing

End Sub

'Private Sub WriteReportLog()
'   Dim iList As Integer
'   On Error Resume Next
'   sSql = "TRUNCATE TABLE EsReportClosedRunsLog"
'   clsAdoCon.ExecuteSQL sSql
'   For iList = 1 To iLogNumber
'      If Len(sLogNote(iList, NOTE_Description)) > 80 Then sLogNote(iList, NOTE_Description) = Left$(sLogNote(iList, NOTE_Description), 80)
'      sSql = "INSERT INTO EsReportClosedRunsLog (LOG_NUMBER,LOG_TEXT) " _
'             & "VALUES(" & sLogNote(iList, NOTE_Number) & ",'" & sLogNote(iList, NOTE_Description) & "')"
'      clsAdoCon.ExecuteSQL sSql
'   Next
'
'End Sub
'
Public Sub CheckLogTable()
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   sSql = "SELECT LOG_NUMBER FROM EsReportClosedRunsLog"
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum > 0 Then
      clsADOCon.ADOErrNum = 0
      sSql = "CREATE TABLE EsReportClosedRunsLog (" _
             & "LOG_NUMBER SMALLINT NULL DEFAULT(0)," _
             & "LOG_TEXT VARCHAR(80) NULL DEFAULT('')," _
             & "LOG_PARTNO CHAR(30) NULL DEFAULT('')," _
             & "LOG_RUNNO INT NULL DEFAULT(0)," _
             & "LOG_CLOSED TINYINT NULL DEFAULT(0))"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then
         sSql = "CREATE UNIQUE CLUSTERED INDEX LogIndex ON " _
                & "EsReportClosedRunsLog(LOG_NUMBER) WITH FILLFACTOR = 80"
         clsADOCon.ExecuteSQL sSql
      End If
   End If
   
End Sub


Private Sub GetWipAccounts()
   Dim b As Byte
   sProcName = "getlaboracct"
   sInvLabAcct = GetLaborAcct(sPartNumber, lblCode, Val(lblLvl))
   sProcName = "getexpenseacct"
   sInvExpAcct = GetExpenseAcct(sPartNumber, lblCode, Val(lblLvl))
   sProcName = "getmaterialacct"
   sInvMatAcct = GetMaterialAcct(sPartNumber, lblCode, Val(lblLvl))
   sProcName = "getoverheadacct"
   sInvOhdAcct = GetOverHeadAcct(sPartNumber, lblCode, Val(lblLvl))
   sProcName = "getcgsaccount"
   b = GetCGSAccounts(sCgsAcct)
   
End Sub

Private Function GetCGSAccounts(CostOfGoods As String) As Byte
   Dim rdoAct As ADODB.Recordset
   Dim bType As Byte
   Dim sPcode As String
   
   'Use current Part
   bType = Val(lblLvl)
   sSql = "SELECT PAPRODCODE,PACGSMATACCT FROM PartTable WHERE PARTREF='" & sPartNumber & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         sPcode = "" & Trim(!PAPRODCODE)
         CostOfGoods = "" & Trim(!PACGSMATACCT)
         ClearResultSet rdoAct
      End With
   End If
   Set rdoAct = Nothing
   If CostOfGoods = "" Then
      'None in one or any, try Product code
      sSql = "SELECT PCCGSMATACCT FROM PcodTable WHERE PCREF='" _
             & sPcode & "' "
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
      If bSqlRows Then
         With rdoAct
            If CostOfGoods = "" Then CostOfGoods = "" & Trim(!PCCGSMATACCT)
            ClearResultSet rdoAct
         End With
      End If
   End If
   If CostOfGoods = "" Then
      'Still none, we'll check the common
      sSql = "SELECT COCGSMATACCT" & Trim(str(bType)) & " " _
             & "FROM ComnTable WHERE COREF=1"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
      If bSqlRows Then
         With rdoAct
            If CostOfGoods = "" Then CostOfGoods = "" & Trim(.Fields(0))
            ClearResultSet rdoAct
         End With
      End If
   End If
   'After this excercise, we'll give up if none are found
   Set rdoAct = Nothing
   
End Function

''Close MO
''9/14/06 Added Returns (Bottom)
'
Private Sub cmbPrt_Click()
   bGoodPrt = GetRunPart()
   GetRuns
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If Len(cmbPrt) Then
      bGoodPrt = GetRunPart()
      GetRuns
   End If
   
End Sub


Private Sub cmbRun_Click()
   bGoodRun = GetCurrRun()
   
End Sub


Private Sub cmbRun_LostFocus()
   cmbRun = CheckLen(cmbRun, 5)
   cmbRun = Format(Abs(Val(cmbRun)), "####0")
   If Val(cmbRun) > 32767 Then cmbRun = "32767"
   bGoodRun = GetCurrRun()
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   cmbPrt = ""
   
End Sub


Private Sub cmdCloseMO_Click()
   Dim RdoQty As ADODB.Recordset
   Dim bByte As Byte
   Dim lClose As Long
   Dim lComplete As Long
   
   bCantClose = 0
   sJournalID = GetOpenJournal("IJ", Format$(cmbCloseDate, "mm/dd/yyyy"))
   If Left(sJournalID, 4) = "None" Then
      sJournalID = ""
      bByte = 1
   Else
      If sJournalID = "" Then bByte = 0 Else bByte = 1
   End If
   If bByte = 0 Then
      MsgBox "There Is No Open Inventory Journal For The Period " & cmbCloseDate & ".", _
         vbInformation, Caption
      Exit Sub
   End If
   bGoodRun = GetCurrRun()
   If bGoodRun = 1 Then
      Dim sMsg As String, bResponse As Byte
      sMsg = "This Closes The MO To All Functions." & vbCrLf _
          & "Do You Really Want To Close This MO?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse <> vbYes Then
         Exit Sub
      End If
      
      cmdCloseMO.Enabled = False
      cmdCan.Enabled = False
      MouseCursor ccHourglass
      Dim mo As New ClassMO
      mo.LoggingEnabled = True
      mo.CloseRequiresInvoices = CBool(chkInvoices.Value)
      mo.CloseIgnoresUnpicked = CBool(chkIgnoreUnpicked.Value)
      mo.CloseIgnoresExpendables = CBool(chkIgnoreExpendables.Value)
      mo.CloseDate = CDate(cmbCloseDate.Text)
      mo.PartNumber = cmbPrt
      mo.RunNumber = cmbRun
      
      'determine whether picks for unclosed MOs
      If mo.AreTherePicksForUnclosedMos Then
         MouseCursor ccArrow
         cmdCloseMO.Enabled = True
         cmdCan.Enabled = True
         MsgBox "There are picks for unclosed MOs.  Can't close.  See log."
         Exit Sub
      End If
       
      If mo.CloseMO() Then
         MouseCursor ccArrow
         MsgBox "MO closed.  See log."
         FillCombo
      Else
         MouseCursor ccArrow
         MsgBox "Cannot close this MO.  See log."
      End If
      
'      lClose = DateValue(Format(cmbCloseDate, "yyyy,mm,dd"))
'      lComplete = DateValue(Format(lblDte, "yyyy,mm,dd"))
'      If lClose < lComplete Then
'         MsgBox "The Date of Closure Cannot Be Before The" & vbCr _
'            & "Completion Date.", _
'            vbInformation, Caption
'         Exit Sub
'      End If
'      iLogNumber = 1
'      sLogNote(iLogNumber, NOTE_Number) = Str(iLogNumber)
'      sLogNote(iLogNumber, NOTE_Description) = "Close MO " & cmbPrt & " Run " & cmbRun
'
'      iLogNumber = iLogNumber + 1
'      sLogNote(iLogNumber, NOTE_Number) = Str(iLogNumber)
'      sLogNote(iLogNumber, NOTE_Description) = " "
'      On Error GoTo DiaErr1
'      GetUnInvoicedPoItems ' L
'      GetPickList ' L
'      GetExpenseCosts ' L
'      GetLaborCosts ' L
'      GetMaterialCosts ' L
'      If bCantClose = 0 Then
'         CloseMo
'      Else
'         On Error Resume Next
'         WriteReportLog
'         MsgBox "Cannot Close This MO Run. See Log.", _
'            vbInformation, Caption
'      End If
   Else
      MsgBox "You Must Select A Valid Run.", _
         vbInformation, Caption
   End If
   Set RdoQty = Nothing
   MouseCursor ccArrow
   cmdCloseMO.Enabled = True
   cmdCan.Enabled = True
   Exit Sub
   
DiaErr1:
   'WriteReportLog
   MsgBox "Cannot Close This MO Run. See Log.", _
      vbInformation, Caption
   sProcName = "cmdCloseMO_Click"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4153
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdVew_Click()
   MouseCursor 13
   On Error GoTo DiaErr1
   
   
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection

   sCustomReport = GetCustomReport("closedruns")
'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "ShowUnclosedOnly"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
    aFormulaValue.Add CInt(0)
    
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.SetDbTableConnection
   
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName

   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
'   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MDISect.Crw.Formulas(1) = "RequestBy = 'Requested By: " & sInitials & "'"
'   MDISect.Crw.Formulas(2) = "ShowUnclosedOnly = 0 "
'
'   sCustomReport = GetCustomReport("closedruns")
'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   'SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Activate()
   Dim b As Byte
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      CheckLogTable
      CheckInvoicing
      b = CheckInvJournal()
      bLotsOn = CheckLotStatus
      If b = 1 Then FillCombo
      'bLotsOn = 1
      'If bLotsOn Then
      'Debug.Print "1"
      'End If
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   GetSettings
   sSql = "SELECT RUNREF,RUNNO FROM RunsTable WHERE " _
          & "RUNREF= ? AND RUNSTATUS='CO' "
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Type = adChar
   AdoParameter.SIZE = 30
   
   AdoQry.Parameters.Append AdoParameter
   
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveSettings
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   Set ShopSHf04a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   cmbCloseDate = Format(ES_SYSDATE, "mm/dd/yyyy")
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   cmbPrt.Clear
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,RUNREF FROM " _
          & "RunsTable,PartTable WHERE PARTREF=RUNREF AND " _
          & "RUNSTATUS='CO' ORDER BY PARTREF"
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then
      cmbPrt = cmbPrt.List(0)
      bGoodPrt = GetRunPart()
      GetRuns
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetRuns()
   Dim RdoRns As ADODB.Recordset
   cmbRun.Clear
   sPartNumber = Compress(cmbPrt)
   AdoQry.Parameters(0).Value = sPartNumber
   bSqlRows = clsADOCon.GetQuerySet(RdoRns, AdoQry, ES_FORWARD)
   If bSqlRows Then
      With RdoRns
         Do Until .EOF
            AddComboStr cmbRun.hwnd, Format$(!Runno, "####0")
            .MoveNext
         Loop
         ClearResultSet RdoRns
      End With
   End If
   Set RdoRns = Nothing
   If cmbRun.ListCount > 0 Then
      cmbRun = cmbRun.List(0)
      If GetPreferenceValue("AutoSelectLastRun") = "1" Then cmbRun = cmbRun.List(cmbRun.ListCount - 1)
      bGoodRun = GetCurrRun()
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetCurrRun() As Byte
   Dim RdoRun As ADODB.Recordset
   
   lRunno = Val(cmbRun)
   On Error GoTo DiaErr1
   sSql = "SELECT RUNREF,RUNNO,RUNSTATUS,RUNYIELD,RUNCOMPLETE,RUNPARTIALQTY,RUNLOTNUMBER FROM RunsTable " _
          & "WHERE RUNREF='" & Compress(cmbPrt) & "' AND RUNNO=" & lRunno & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRun, ES_FORWARD)
   If bSqlRows Then
      With RdoRun
         lblStat = "" & Trim(!RUNSTATUS)
         lblDte = Format(!RUNCOMPLETE, "mm/dd/yyyy")
         sLotNumber = "" & Trim(!RUNLOTNUMBER)
         lblQty = Format(!RUNYIELD, ES_QuantityDataFormat)
         cYield = !RUNYIELD
         lblComplete = Format(!RUNPARTIALQTY, ES_QuantityDataFormat)
         ClearResultSet RdoRun
      End With
   Else
      cYield = 0
      lblStat = "**"
      lblDte = ""
      sLotNumber = ""
   End If
   If lblStat = "CO" Then
      GetCurrRun = 1
   Else
      GetCurrRun = 0
      lblDte = ""
      lblStat = "**"
      sLotNumber = ""
   End If
   Set RdoRun = Nothing
   GetPreviousCompletions
   Exit Function
   
DiaErr1:
   sProcName = "getcurrrun"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function



Private Sub lblDsc_Change()
   If Left(lblDsc, 10) = "*** Part N" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = Es_TextForeColor
   End If
   
End Sub

Private Sub lblStat_Change()
   If lblStat = "**" Then
      lblStat.ForeColor = ES_RED
   Else
      lblStat.ForeColor = Es_TextForeColor
   End If
   
End Sub





'Private Sub CloseMo()
'   Dim RdoInv As ADODB.Recordset
'   Dim bResponse As Byte
'   Dim lRunno As Long
'   Dim lInRecord As Long
'   Dim cRunCost As Currency
'   Dim sMsg As String
'   Dim sPart As String
'   Dim vaDate As Variant
'
'   vaDate = GetServerDateTime()
'   sPart = Compress(cmbPrt)
'   lRunno = Val(cmbRun)
'
'
'
'   On Error GoTo DiaErr1
'   sMsg = "This Closes The MO To All Functions." & vbCrLf _
'          & "Do You Really Want To Close This MO?"
'   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
'   If bResponse = vbYes Then
'      cRunCost = cRunOvHd + cRunMatl + cRunExp + cRunLabor
'      If cYield = 0 Then cYield = 1
'      On Error Resume Next
'      clsAdoCon.begintrans
'      sSql = "UPDATE RunsTable SET RUNSTATUS='CL'," _
'             & "RUNCOST=" & cRunCost & "," _
'             & "RUNOHCOST=" & cRunOvHd & "," _
'             & "RUNCMATL=" & cRunMatl & "," _
'             & "RUNCEXP=" & cRunExp & "," _
'             & "RUNCHRS=" & cRunHours & "," _
'             & "RUNCLAB=" & cRunLabor & "," _
'             & "RUNCLOSED='" & cmbCloseDate & "'," _
'             & "RUNREVBY='" & sInitials & "' " _
'             & "WHERE (RUNREF='" & sPart & "' AND " _
'             & "RUNNO=" & lRunno & ")"
'      clsAdoCon.ExecuteSQL sSql
'
''      'LOTS - done in UpdateMoCosts
''      If sLotNumber <> "" Then
''         sSql = "UPDATE LohdTable SET " _
''                & "LOTDATECOSTED='" & vaDate & "'," _
''                & "LOTUNITCOST=" & cRunCost / cYield & "," _
''                & "LOTTOTMATL=" & cRunMatl & "," _
''                & "LOTTOTLABOR=" & cRunLabor & "," _
''                & "LOTTOTEXP=" & cRunExp & "," _
''                & "LOTTOTOH=" & cRunOvHd & "," _
''                & "LOTTOTHRS=" & cRunHours & " " _
''                & "WHERE LOTNUMBER='" & sLotNumber & "'"
''         clsAdoCon.ExecuteSQL sSql
''      End If
'      If Err = 0 Then
'         'clsAdoCon.CommitTrans
'
'         '                'determine whether to place standard cost in inventory activity record
'         '                ssql = "SELECT COALESCE(INLOTTRACK, PALOTTRACK, 0) as LotTrack," & vbCrLf _
'         '                    & "COALESCE(INUSEACTUALCOST, PAUSEACTUALCOST, 0) as UseActualCost, " & vbCrLf _
'         '                    & "PASTDCOST, PATOTMATL, PATOTLABOR, PATOTEXP, PATOTOH, PATOTHRS" & vbCrLf _
'         '                    & "from InvaTable join PartTable on INPART = PARTREF" & vbCrLf _
'         '                    & "WHERE INTYPE=6 " & vbCrLf _
'         '                    & "AND INPART='" & sPart & "' " _
'         '                    & "AND INMORUN=" & lRunno
'         '
'         '                Dim rdo As ADODB.Recordset
'         '                Dim cUnitCost As Currency
'         '                If GetDataSet(rdo, ES_FORWARD) <> 0 Then
'         '                    With rdo
'         '                        If !LotTrack = 0 Or (!LotTrack = 1 And !UseActualCost = 0) Then
'         '                            cUnitCost = !PASTDCOST
'         '                            cRunMatl = !PATOTMATL
'         '                            cRunLabor = !PATOTLABOR
'         '                            cRunExp = !PATOTEXP
'         '                            cRunOvHd = !PATOTOH
'         '                            cRunHrs = !PATOTHRS
'         '                        End If
'         '                    End With
'         '                Else
'         '                    cUnitCost = cRunCost / cYield
'         '                End If
'         '                rdo.Close
'         '                Set rdo = Nothing
'         '
'         '                ssql = "UPDATE InvaTable SET " _
'         '                    & "INREF1='CLOSED RUN'," _
'         '                    & "INADATE='" & vAdate & "'," _
'         '                    & "INAMT=" & cUnitCost & "," _
'         '                    & "INTOTMATL=" & cRunMatl & "," _
'         '                    & "INTOTLABOR=" & cRunLabor & "," _
'         '                    & "INTOTEXP=" & cRunExp & "," _
'         '                    & "INTOTOH=" & cRunOvHd & "," _
'         '                    & "INTOTHRS=" & cRunHours & " " _
'         '                    & "WHERE (INTYPE=6 AND INPART='" & sPart & "' " _
'         '                    & "AND INMORUN=" & lRunno & ")"
'         '                clsAdoCon.ExecuteSQL sSql
'
''         Dim unitCost As Currency
''         If cYield = 0 Then
''            unitCost = 0
''         Else
''            unitCost = cRunCost / cYield
''         End If
'
'         Dim mo As New ClassMO
'         iLogNumber = iLogNumber + 1
'         sLogNote(iLogNumber, NOTE_Number) = Str(iLogNumber)
''         sLogNote(iLogNumber, NOTE_Description) = mo.UpdateMOCosts(sPart, lRunno, "CLOSED RUN", _
''                  cYield, unitCost, _
''                  cRunMatl, cRunLabor, cRunExp, cRunOvHd, cRunHours)
'         sLogNote(iLogNumber, NOTE_Description) = mo.UpdateMOCosts(sPart, lRunno, "CLOSED RUN", _
'                  cRunMatl, cRunLabor, cRunExp, cRunOvHd, cRunHours)
'         clsAdoCon.CommitTrans
'
'         iLogNumber = iLogNumber + 1
'         sLogNote(iLogNumber, NOTE_Number) = Str(iLogNumber)
'         sLogNote(iLogNumber, NOTE_Description) = " "
'         iLogNumber = iLogNumber + 1
'         sLogNote(iLogNumber, NOTE_Number) = Str(iLogNumber)
'         sLogNote(iLogNumber, NOTE_Description) = "Manufacturing Order " & cmbPrt & " Run " & cmbRun & " Was Closed"
'         WriteReportLog
''         sMsg = "The Status Was Changed From CO To CL." & vbCrLf _
''                & "No Additional Action Can Be Executed."
''         MsgBox sMsg, vbInformation, Caption
''
''         sMsg = "Press The View Button To Study The Log."
''         MsgBox sMsg, vbInformation, Caption
'
'         sMsg = "The Status Was Changed From CO To CL." & vbCrLf _
'            & "No Additional Action Can Be Executed." & vbCrLf _
'            & "Press The View Button To Study The Log."
'         MsgBox sMsg, vbInformation, Caption
'
'         FillCombo
'      Else
'         clsAdoCon.RollbackTrans
'         iLogNumber = iLogNumber + 1
'         sLogNote(iLogNumber, NOTE_Number) = Str(iLogNumber)
'         sLogNote(iLogNumber, NOTE_Description) = " "
'         iLogNumber = iLogNumber + 1
'         sLogNote(iLogNumber, NOTE_Number) = Str(iLogNumber)
'         sLogNote(iLogNumber, NOTE_Description) = "* Manufacturing Order " & cmbPrt & " Run " & cmbRun & " Was Not Closed"
'         WriteReportLog
'         MsgBox "Couldn't Change The Run To Closed (CL).", vbExclamation, Caption
'         sMsg = "Press The View Button To Study The Log."
'         MsgBox sMsg, vbInformation, Caption
'      End If
'   Else
'      CancelTrans
'   End If
'   Exit Sub
'
'DiaErr1:
'   sProcName = "closemo"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub
'
Private Sub cmbCloseDate_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub cmbCloseDate_LostFocus()
   cmbCloseDate = CheckDateEx(cmbCloseDate)
   
End Sub




Private Function GetRunPart() As Byte
   Dim RdoPrt As ADODB.Recordset
   cmbRun.Clear
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PAPRODCODE," _
          & "PASTDCOST FROM PartTable WHERE PARTREF='" & Compress(cmbPrt) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
   If bSqlRows Then
      With RdoPrt
         cmbPrt = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         lblLvl = Format(!PALEVEL, "0")
         lblCode = "" & Trim(!PAPRODCODE)
         cStdCost = Format(!PASTDCOST, ES_QuantityDataFormat)
         ClearResultSet RdoPrt
         GetRunPart = 1
      End With
   Else
      GetRunPart = 0
      lblLvl = ""
      lblCode = ""
      lblDsc = "*** Part Number Wasn't Found ****"
   End If
   Set RdoPrt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getrunpart"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function CheckInvJournal() As Byte
   Dim b As Byte
   sJournalID = GetOpenJournal("IJ", Format$(cmbCloseDate, "mm/dd/yyyy"))
   If Left(sJournalID, 4) = "None" Then
      sJournalID = ""
      b = 1
   Else
      If sJournalID = "" Then b = 0 Else b = 1
   End If
   If b = 0 Then
      MsgBox "There Is No Open Inventory Journal For The Period.", _
         vbExclamation, Caption
      CheckInvJournal = 0
      Sleep 500
      Unload Me
   Else
      CheckInvJournal = 1
   End If
   
End Function


'Reserves the material for later use

Public Sub GetPartNumber(PartNumber As String, iRow As Integer)
   Dim RdoGet As ADODB.Recordset
   On Error Resume Next
   sSql = "SELECT PARTREF,PARTNUM,PALOTTRACK,PASTDCOST FROM PartTable " _
          & "WHERE PARTREF='" & PartNumber & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet, ES_FORWARD)
   If bSqlRows Then
      With RdoGet
         sPartLots(iRow, PART_Number) = "" & Trim(!PartNum)
         If bLotsOn Then
            sPartLots(iRow, PART_IsLotTracked) = Trim(str$(!PALOTTRACK))
         Else
            sPartLots(iRow, PART_IsLotTracked) = "0"
         End If
         sPartLots(iRow, PART_StdCost) = Trim(str$(!PASTDCOST))
      End With
      ClearResultSet RdoGet
   End If
   Set RdoGet = Nothing
End Sub

'Test Invoicing

Public Sub CheckInvoicing()
   Dim RdoInv As ADODB.Recordset
   On Error Resume Next
   sSql = "SELECT COVERIFYINVOICES FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_FORWARD)
   If bSqlRows Then chkInvoices.Value = RdoInv!COVERIFYINVOICES
   ClearResultSet RdoInv
   Set RdoInv = Nothing
   
End Sub

Private Sub SaveSettings()
   SaveSetting "Esi2000", "EsiProd", "ShopSHf04a", Trim(str$(chkIgnoreUnpicked.Value))
   SaveSetting "Esi2000", "EsiProd", "ShopSHf04aa", Trim(str$(chkIgnoreExpendables.Value))
   
End Sub

Private Sub GetSettings()
   chkIgnoreUnpicked.Value = GetSetting("Esi2000", "EsiProd", "ShopSHf04a", Trim(str$(chkIgnoreUnpicked.Value)))
   chkIgnoreExpendables.Value = GetSetting("Esi2000", "EsiProd", "ShopSHf04aa", Trim(str$(chkIgnoreExpendables.Value)))
   
End Sub

