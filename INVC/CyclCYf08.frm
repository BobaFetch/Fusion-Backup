VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form CyclCYf08 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update inventory based on physical counts"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   8460
   Begin VB.CommandButton optPrn 
      Height          =   330
      Left            =   7440
      Picture         =   "CyclCYf08.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Print The Report"
      Top             =   2400
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton optDis 
      Height          =   330
      Left            =   6840
      Picture         =   "CyclCYf08.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Display The Report"
      Top             =   2400
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "CyclCYf08.frx":0308
      Height          =   350
      Left            =   6300
      Picture         =   "CyclCYf08.frx":07E2
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "View Cycle Count Problems"
      Top             =   2340
      Width           =   360
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test Lot Allocations"
      Height          =   360
      Left            =   2340
      TabIndex        =   3
      ToolTipText     =   "Fill The Form With Qualifying Items"
      Top             =   2340
      Width           =   1770
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update Inventory"
      Height          =   360
      Left            =   4320
      TabIndex        =   4
      ToolTipText     =   "Fill The Form With Qualifying Items"
      Top             =   2340
      Width           =   1770
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CyclCYf08.frx":0CBC
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbCid 
      Height          =   315
      Left            =   1860
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "List Includes Cycle ID's Not Locked Or Completed"
      Top             =   120
      Width           =   2115
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1860
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   3375
   End
   Begin VB.ComboBox txtPlan 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6060
      Locked          =   -1  'True
      TabIndex        =   1
      Tag             =   "4"
      ToolTipText     =   "Planned Inventory Date"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   7320
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   60
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   60
      Top             =   360
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3015
      FormDesignWidth =   8460
   End
   Begin VB.Label lblCountsRequired 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   4080
      TabIndex        =   18
      ToolTipText     =   "Total Items Included"
      Top             =   1740
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Counts required"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   17
      Top             =   1740
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "No lots"
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   16
      Top             =   1500
      Width           =   1095
   End
   Begin VB.Label lblNoLots 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   4080
      TabIndex        =   15
      ToolTipText     =   "Total Items Included"
      Top             =   1500
      Width           =   855
   End
   Begin VB.Label lblCountsEntered 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   4080
      TabIndex        =   13
      ToolTipText     =   "Total Items Included"
      Top             =   1260
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Counts entered"
      Height          =   255
      Index           =   4
      Left            =   2880
      TabIndex        =   12
      Top             =   1260
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Items"
      Height          =   255
      Index           =   10
      Left            =   2880
      TabIndex        =   11
      Top             =   1020
      Width           =   495
   End
   Begin VB.Label lblTotalItems 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      ToolTipText     =   "Total Items Included"
      Top             =   1020
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cycle Count ID"
      Height          =   255
      Index           =   5
      Left            =   540
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   6
      Left            =   540
      TabIndex        =   8
      Top             =   525
      Width           =   1335
   End
   Begin VB.Label lblCabc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4140
      TabIndex        =   7
      ToolTipText     =   "ABC Code Selected"
      Top             =   120
      Width           =   405
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Planned Date"
      Height          =   255
      Index           =   7
      Left            =   4740
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "CyclCYf08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
Option Explicit
Dim bOnLoad As Byte
'Dim bGoodCount As Byte

Private TotalItemsFromGrid As Integer
Private CountsEnteredFromGrid As Integer
Private CountsRequiredFromGrid As Integer
Private NoLotsFromGrid As Integer

Private Const LOTREQUIREDMSG As String = "LOT REQUIRED"


'Dim iTotalLots As Integer
'Dim iIndex As Integer
'Dim lCOUNTER As Long

Dim sCreditAcct As String
Dim sDebitAcct As String

Dim vNextDate As Variant
Private editingRow As Integer
Private editingCol As Integer


'grid columns
Private Const COL_Location = 0
Private Const COL_PartRef = 1
Private Const COL_PartDescription = 2
Private Const COL_PartQty = 3
Private Const COL_UOM = 4
Private Const COL_PartCount = 5
Private Const COL_LotNo = 6
Private Const COL_UserLotNo = 7
Private Const COL_LotQty = 8
Private Const COL_LotCount = 9
Private Const COL_IsLotTracked = 10
Private Const COL_Count = 11      'number of columns

'grid cell colors
Private Const COLOR_NotEntered = &HC0C0FF          'red
Private Const COLOR_Entered = &HC0FFC0             'green
Private Const COLOR_NotEditable = &HE0E0E0         'grey
Private Const COLOR_ReadOnly = &HFFFFFF

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub GetAccounts(PartNumber As String)
   Dim rdoAct As ADODB.Recordset
   Dim bType As Byte
   Dim sPcode As String
   
   sDebitAcct = ""
   sCreditAcct = ""
   sSql = "SELECT COADJACCT FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         If Not IsNull(!COADJACCT) Then _
                       sDebitAcct = "" & Trim(!COADJACCT)
         ClearResultSet rdoAct
      End With
   End If
   Set rdoAct = Nothing
   'Use current Part
   sSql = "Qry_GetExtPartAccounts '" & PartNumber & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         sPcode = "" & Trim(!PAPRODCODE)
         If sDebitAcct = "" Then sDebitAcct = "" & Trim(!PACGSMATACCT)
         sCreditAcct = "" & Trim(!PAINVEXPACCT)
         ClearResultSet rdoAct
         Set rdoAct = Nothing
      End With
   Else
      sCreditAcct = ""
      sDebitAcct = ""
      Exit Sub
   End If
   If sDebitAcct = "" Or sCreditAcct = "" Then
      'None in one or both there, try Product code
      sSql = "Qry_GetPCodeAccounts '" & sPcode & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
      If bSqlRows Then
         With rdoAct
            If sDebitAcct = "" Then sCreditAcct = "" & Trim(!PCCGSMATACCT)
            If sCreditAcct = "" Then sDebitAcct = "" & Trim(!PCINVMATACCT)
            ClearResultSet rdoAct
            Set rdoAct = Nothing
         End With
      End If
      If sDebitAcct = "" Or sCreditAcct = "" Then
         'Still none, we'll check the common
         sSql = "SELECT COCGSMATACCT" & Trim(str(bType)) & "," _
                & "COINVMATACCT" & Trim(str(bType)) & " " _
                & "FROM ComnTable WHERE COREF=1"
         bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
         If bSqlRows Then
            With rdoAct
               If sCreditAcct = "" Then sDebitAcct = "" & Trim(.Fields(0))
               If sDebitAcct = "" Then sCreditAcct = "" & Trim(.Fields(1))
               ClearResultSet rdoAct
               Set rdoAct = Nothing
            End With
         End If
      End If
   End If
   Set rdoAct = Nothing
End Sub

Private Sub GetCycleCount()
   Dim RdoCid As ADODB.Recordset
   sSql = "SELECT *" & vbCrLf _
      & "FROM CchdTable" & vbCrLf _
      & "WHERE CCREF = '" & cmbCid & "'" & vbCrLf _
      & "AND CCCOUNTLOCKED = 1"
   If clsADOCon.GetDataSet(sSql, RdoCid, ES_FORWARD) Then
      With RdoCid
         lblCabc = "" & Trim(!CCABCCODE)
         txtDsc = "" & Trim(!CCDESC)
         txtPlan = Format(!CCPLANDATE, "mm/dd/yy")
         'cmdSel.Enabled = True
      End With
   End If
   Set RdoCid = Nothing
   Dim cc As New ClassCycleCount
   Dim ok As Boolean
   Dim TotalItems As Integer, CountsEntered As Integer, NoLots As Integer, CountsRequired As Integer
   Me.cmdUpdate.Enabled = cc.AnalyzeCounts(cmbCid, TotalItems, CountsEntered, NoLots, CountsRequired)
   Me.lblTotalItems = TotalItems
   Me.lblCountsEntered = CountsEntered
   Me.lblNoLots = NoLots
   Me.lblCountsRequired = CountsRequired
   If CountsRequired = 0 Then
      cmdTest.Enabled = True
      cmdUpdate.Enabled = True
   Else
      cmdTest.Enabled = False
      cmdUpdate.Enabled = False
   End If
   
End Sub

Private Sub cmbCid_Click()
   GetCycleCount
End Sub

Private Sub cmbCid_LostFocus()
   GetCycleCount
End Sub


Private Sub cmdCancel_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "5455"
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdTest_Click()
   OKToUpdate True
End Sub

Private Function OKToUpdate(Testing As Boolean) As Boolean
   'if testing = true, display a message
   'returns true if OK to perform the update
   'if true, the allocations are in table CCLotAlloc
   'if false, the reasons why are in table CCLog
   
   sSql = "exec AllocateCycleCountLots '" & cmbCid & "'"
   clsADOCon.ExecuteSQL sSql
   
   'see if any error messages
   sSql = "select count(*) from CCLog where CCREF = '" & cmbCid & "' and ERRORTYPE = 'FATAL'"
   Dim rdo As ADODB.Recordset
   Dim ct As Integer
   If clsADOCon.GetDataSet(sSql, rdo) Then
      ct = rdo.Fields(0)
   End If
   
   If ct <> 0 Then
      If Testing Then
         MsgBox "There were " & ct & " fatal errors." & vbCrLf _
            & "Cannot proceed.  See the log."
      End If
      OKToUpdate = False
   Else
      sSql = "select count(*) from CCLog where CCREF = '" & cmbCid & "' and ERRORTYPE = 'WARNING'"
      If clsADOCon.GetDataSet(sSql, rdo) Then
         ct = rdo.Fields(0)
      Else
         ct = 0
      End If
   
      If ct <> 0 Then
         If Testing Then
            MsgBox "There were " & ct & " warnings encoutered." & vbCrLf _
               & "OK to proceed but check log first."
         End If
      Else
         If Testing Then
            MsgBox "No errors or warnings encoutered." & vbCrLf _
               & "OK to proceed."
         End If
      End If

      OKToUpdate = True
   End If
   Set rdo = Nothing
End Function

Private Sub cmdUpdate_Click()
   PerformUpdate
End Sub

Private Function PerformUpdate() As Boolean
   'returns true if successful
   
   On Error GoTo whoops
   If OKToUpdate(True) Then
   
      If MsgBox("Proceed with update?", vbYesNo) <> vbYes Then
         Exit Function
      End If
      MouseCursor ccHourglass
      Me.cmdUpdate.Enabled = False
      Me.cmdTest.Enabled = False
      sSql = "exec UpdateCycleCount '" & Me.cmbCid & "', '" & sInitials & "'"
      clsADOCon.ExecuteSQL sSql
      MouseCursor ccDefault
      FillCombo               'remove completed cycle count
      MsgBox "Update completed"
   End If
   
   Exit Function
   
whoops:
   MouseCursor ccDefault
   sProcName = "PerformUpdate"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub cmdVew_Click()
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   On Error GoTo DiaErr1
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "CycleCountID"
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & CStr("Requested By: " & sInitials) & "'")
   aFormulaValue.Add CStr("'" & CStr(Me.cmbCid) & "'")
   sCustomReport = GetCustomReport("cclog")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.ShowGroupTree False
   
'   sSql = "{LohdTable.LOTNUMBER}='" & lblNumber & "'"
'   cCRViewer.SetReportSelectionFormula (sSql)
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   

   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   MouseCursor 0
   Exit Sub


'   MouseCursor 13
'   On Error GoTo DiaErr1
'   'S 'etMdiReportsize MdiSect
'   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MdiSect.Crw.Formulas(1) = "RequestBy = 'Requested By: " & sInitials & "'"
'   MdiSect.Crw.Formulas(2) = "CycleCountID = '" & Me.cmbCid & "'"
'   sCustomReport = GetCustomReport("cclog")
''   MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
 '  'SetCrystalAction Me
 '  MouseCursor 0
 '  Exit Sub
   
DiaErr1:
   sProcName = "cmdVew_Click"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      vNextDate = GetNextDate()
   End If
   bOnLoad = 0
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   bOnLoad = 1
   
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDsc.BackColor = Es_FormBackColor
   txtPlan.BackColor = Es_FormBackColor
   txtPlan.ToolTipText = "Planned Inventory Date"
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   cmbCid.Clear
   sSql = "SELECT CCREF FROM CchdTable WHERE (CCCOUNTLOCKED=1 AND " _
          & "CCUPDATED=0)"
   LoadComboBox cmbCid, -1
   If cmbCid.ListCount > 0 Then
      If Trim(cmbCid) = "" Then cmbCid = cmbCid.List(0)
   Else
      MsgBox "There Are No Locked And Not Reconciled Counts Recorded.", _
         vbInformation, Caption
      Unload Me
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetNextDate() As Variant
   Dim RdoDate As ADODB.Recordset
   Dim iFrequency As Integer
   Dim dDate As Date
   
   On Error Resume Next
   dDate = Format(txtPlan, "mm/dd/yy")
   sSql = "SELECT COABCROW,COABCCODE,COABCFREQUENCY " _
          & "FROM CabcTable WHERE COABCCODE='" & lblCabc & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDate, ES_FORWARD)
   If bSqlRows Then iFrequency = RdoDate!COABCFREQUENCY
   GetNextDate = Format(dDate + iFrequency, "mm/dd/yy")
   Set RdoDate = Nothing
   
End Function


'Private Sub MarkReconciled()
'   Dim rdoRec As rdoResultset
'   Dim bByte As Byte
'
'   sSql = "SELECT CIREF,CIPARTREF,CIRECONCILED FROM CcitTable WHERE " _
'          & "(CIREF='" & cmbCid & "' AND CIRECONCILED=0)"
'   bSqlRows = GetDataSet(rdoRec, ES_FORWARD)
'   If bSqlRows Then Exit Sub
'
'   sSql = "UPDATE CchdTable SET CCUPDATEDDATE='" & Format(ES_SYSDATE, "mm/dd/yy") _
'          & "',CCUPDATED=1 WHERE CCREF='" & cmbCid & "'"
'   RdoCon.Execute sSql, rdExecDirect
'   If Err = 0 Then MsgBox Trim(cmbCid) & " Has Been Reconciled.", _
'            vbInformation, Caption
'   FillCombo
'
'End Sub

Private Sub txtPlan_DropDown()
   ShowCalendar Me
   
End Sub

