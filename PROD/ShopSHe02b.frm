VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form ShopSHe02b 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revise Schedule Date, Quantity,Priority"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHe02b.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optChg 
      Enabled         =   0   'False
      Height          =   255
      Left            =   5340
      TabIndex        =   14
      Top             =   360
      Width           =   255
   End
   Begin VB.ComboBox txtCom 
      Height          =   315
      Left            =   4080
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ComboBox cmbDiv 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   4080
      Sorted          =   -1  'True
      TabIndex        =   3
      Tag             =   "8"
      ToolTipText     =   "Select Division From List"
      Top             =   1440
      Width           =   860
   End
   Begin VB.CommandButton cmdSch 
      Caption         =   "&Schedule"
      Height          =   315
      Left            =   5760
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Update Entries and Re-Schedule"
      Top             =   600
      Width           =   875
   End
   Begin VB.TextBox txtPri 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      ToolTipText     =   "MO Priority"
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      ToolTipText     =   "Run Quantity"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5760
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5760
      Top             =   1800
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2355
      FormDesignWidth =   6675
   End
   Begin MSComctlLib.ProgressBar PrgBar 
      Height          =   300
      Left            =   1320
      TabIndex        =   17
      Top             =   1920
      Visible         =   0   'False
      Width           =   3252
      _ExtentX        =   5741
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Allow Quantity Changed For PL, PP, PC Runs (System Setting)"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   360
      Width           =   4695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      Height          =   255
      Index           =   5
      Left            =   2640
      TabIndex        =   12
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Priority"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sched Completion"
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   10
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "  Run"
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   7
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblRun 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4980
      TabIndex        =   6
      Top             =   720
      Width           =   555
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   720
      Width           =   3075
   End
End
Attribute VB_Name = "ShopSHe02b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'6/4/04 Option to allow qty changes
'9/13/04 Corrected txtQty to allow SC and RL Qty Changes (checked or not)
'3/15/05 Added GetSchedDate to BackSchedule
'3/31/05 Revised BackSchedule queue to the same as New MO
'10/12/05 Added column OPMDATE and calculated OPQDATE AND OPMDATE
'10/12/05 Added GetWeekend
'11/3/05 Revise Q&M dates to correct Hours/Days difference
'11/3/05 Test Forward Schedule
'11/5/05 Added Saturday and Sunday test
'4/28/06 Added Time Conversion For BackSchedule Q&M
Option Explicit
Dim bOnLoad As Byte
Dim bGoodWCCal As Boolean
Dim bGoodCoCal As Boolean

Dim cOldQty As Currency
Dim sOLDDATE As String
Dim sPartNumber As String

Private txtKeyPress(4) As New EsiKeyBd
Private txtGotFocus(4) As New EsiKeyBd
Private txtKeyDown(2) As New EsiKeyBd

Private Sub FormatControls()
   On Error Resume Next
   Set txtGotFocus(0).esTxtGotFocus = txtQty
   Set txtGotFocus(1).esCmbGotfocus = txtCom
   Set txtGotFocus(2).esTxtGotFocus = txtPri
   Set txtGotFocus(3).esCmbGotfocus = cmbDiv
   
   Set txtKeyPress(0).esTxtKeyValue = txtQty
   Set txtKeyPress(1).esCmbKeyDate = txtCom
   Set txtKeyPress(2).esTxtKeyValue = txtPri
   Set txtKeyPress(3).esCmbKeylock = cmbDiv
   
   Set txtKeyDown(0).esTxtKeyDown = txtQty
   Set txtKeyDown(1).esTxtKeyDown = txtPri
   
End Sub


Private Sub cmbDiv_LostFocus()
   cmbDiv = CheckLen(cmbDiv, 4)
   On Error Resume Next
   If Len(cmbDiv) Then
      sSql = "UPDATE RunsTable SET RUNDIVISION='" & cmbDiv & "' " _
             & "WHERE RUNREF='" & sPartNumber & "' AND RUNNO=" & lblRun & " "
      ShopSHe02a.lblDiv = cmbDiv
      clsADOCon.ExecuteSql sSql
   End If
   
End Sub


Private Sub cmdCan_Click()
   Form_Deactivate
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4171
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdSch_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   bGoodWCCal = GetCenterCalendar(Me, txtCom)
   If Not bGoodWCCal Then
      sMsg = "Scheduling Will Be Based On Work Center Time." & vbCr _
             & "This May Not Be Accurate.  Continue Anyway?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbNo Then
         CancelTrans
         Exit Sub
      End If
   End If
   sMsg = "Reschedule MO Based On Latest Information?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      PrgBar.Visible = True
      BackSchedule
      'ForwardSchedule
   Else
      CancelTrans
   End If
   
End Sub


Private Sub Form_Activate()
   If bOnLoad Then
      GetSetting
      FillDivisions
      cOldQty = Val(txtQty)
      bOnLoad = 0
   End If
   
End Sub

Private Sub Form_Deactivate()
   Unload Me
   
End Sub


Private Sub Form_Load()
   SetFormSize Me
   If iBarOnTop Then
      Move MDISect.Left + 500, MDISect.Top + 3200
   Else
      Move 400, MDISect.Top + 2600
   End If
   FormatControls
   lblPrt = ShopSHe02a.cmbPrt
   lblRun = ShopSHe02a.cmbRun
   txtQty = ShopSHe02a.txtQty
   txtPri = ShopSHe02a.txtPri
   txtCom = ShopSHe02a.lblSch
   cmbDiv = ShopSHe02a.lblDiv
   bOnLoad = 1
   sOLDDATE = txtCom
   sPartNumber = Compress(lblPrt)
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   ShopSHe02a.optSrv.Value = vbUnchecked
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Set ShopSHe02b = Nothing
   
End Sub

Private Sub BackSchedule()

   MouseCursor ccHourglass
   
   Dim mo As New ClassMO
   mo.ScheduleOperations sPartNumber, lblRun, CCur(txtQty), txtCom, True
   
   MouseCursor ccDefault

   MsgBox "Operation scheduling complete"

'   Dim RdoOps As ADODB.Recordset
'   Dim bDays As Byte
'   Dim bFirstOp As Byte
'   Dim bQMConvert As Byte 'Hour conversion for Q&M
'   Dim bCounter As Byte 'Fall through for Q&M
'
'   Dim A As Integer
'   Dim iList As Integer
'   Dim n As Integer
'   Dim iDay As Integer
'   Dim iTotalOps As Integer
'   Dim iSunday As Integer
'   Dim iSaturday As Integer
'
'   Dim d As Long
'
'   Dim cOpHrs As Currency
'   Dim cMoveHrs As Currency
'   Dim cQueueHrs As Currency
'   Dim cRunqty As Currency
'   Dim cSetupHrs As Currency
'   Dim cUnitHrs As Currency
'   Dim cUnusedHrs As Currency
'   Dim cTestHours As Currency
'   Dim vSchedDate As Date
'   Dim vStartDate As Date
'   Dim vMoveDate As Date
'   Dim vQueueDate As Date
'
'   Dim sMonth As String
'   Dim sCenter As String
'   Dim sShop As String
'   Dim sOldCenter As String
'
'   Dim mo As New ClassMO
'
'   Dim vRunOps(300, 12) As Variant
'   '0 = Center
'   '1 = Shop
'   '2 = OPNO
'   '3 = QHRS
'   '4 = MHRS
'   '5 = OPSUHRS
'   '6 = OPUNITHRS
'   '7 = Sched Complete
'   '8 = Run Start Date
'   '9 = SU Start Date
'   '10 = Move Date
'   '11 = Queue Date
'   MouseCursor 13
'   'Initialize variant data
'
'   bQMConvert = mo.GetQNMConversion()
'   For iList = 1 To 298
'      vRunOps(iList, 1) = ""
'      vRunOps(iList, 2) = ""
'      vRunOps(iList, 3) = 0
'      vRunOps(iList, 4) = 0
'      vRunOps(iList, 5) = 0
'      vRunOps(iList, 6) = 0
'      vRunOps(iList, 7) = Format(ES_SYSDATE, "mm/dd/yy")
'      vRunOps(iList, 8) = Format(ES_SYSDATE, "mm/dd/yy")
'      vRunOps(iList, 9) = Format(ES_SYSDATE, "mm/dd/yy")
'   Next
'   vRunOps(iList, 1) = ""
'   vRunOps(iList, 2) = ""
'   vRunOps(iList, 3) = 0
'   vRunOps(iList, 4) = 0
'   vRunOps(iList, 5) = 0
'   vRunOps(iList, 6) = 0
'   vRunOps(iList, 7) = Format(ES_SYSDATE, "mm/dd/yy")
'   vRunOps(iList, 8) = Format(ES_SYSDATE, "mm/dd/yy")
'   vRunOps(iList, 9) = Format(ES_SYSDATE, "mm/dd/yy")
'
'   cRunqty = Val(txtQty)
'   'get distinct Shops,Work Centers
'   PrgBar.Value = 10
'   On Error GoTo SrvscErr1
'
'   PrgBar.Value = 20
'   'get the workcenters and ops in reverse order
'   sSql = "SELECT OPREF,OPRUN,OPNO,OPSHOP,OPCENTER,OPQHRS,OPMHRS,OPSUHRS,OPUNITHRS" _
'          & " FROM RnopTable WHERE OPREF='" & sPartNumber & "' AND OPRUN=" & lblRun _
'          & " ORDER BY OPNO DESC"
'   bsqlrows = clsadocon.getdataset(ssql, RdoOps, ES_STATIC)
'   If bSqlRows Then
'      With RdoOps
'         Do Until .EOF
'            iTotalOps = iTotalOps + 1
'            vRunOps(iTotalOps, 0) = "" & Trim(!OPSHOP)
'            vRunOps(iTotalOps, 1) = "" & Trim(!OPCENTER)
'            vRunOps(iTotalOps, 2) = !opNo
'            vRunOps(iTotalOps, 3) = !OPQHRS
'            vRunOps(iTotalOps, 4) = !OPMHRS
'            vRunOps(iTotalOps, 5) = !OPSUHRS
'            vRunOps(iTotalOps, 6) = !OPUNITHRS * cRunqty
'            .MoveNext
'         Loop
'         ClearResultSet RdoOps
'      End With
'      vRunOps(1, 7) = Format(txtCom, "mm/dd/yy")
'   Else
'      MouseCursor 0
'      MsgBox "No Operations To Schedule.", vbExclamation, Caption
'      Exit Sub
'   End If
'
'   PrgBar.Value = 40
'
'   'Get Workcenter Calendars and time
'   vRunOps(0, 4) = vRunOps(1, 4)
'   'Release the memory
'
'   PrgBar.Value = 50
'
'   'Schedule 'em
'   'We'll Allow a Calendar or Work Center Hours.
'   'If neither then they are bumming
'   On Error Resume Next
'   vSchedDate = Format(txtCom, "mm/dd/yy 16:00")
'   bDays = GetWeekEnd(Format(vSchedDate, "mmm-dd-yyyy"))
'   vSchedDate = vSchedDate - bDays
'   cTestHours = 0
'   sOldCenter = ""
'   cUnusedHrs = 0
'   bFirstOp = 1
'   For iList = 1 To iTotalOps 'one op at a time
'      ' Allow move 6/6/00
'      If cUnusedHrs < 0 Then cUnusedHrs = 0
'      If PrgBar.Value < 100 Then PrgBar.Value = PrgBar.Value + 10
'      cQueueHrs = Val(vRunOps(iList, 3))
'      cMoveHrs = Val(vRunOps(iList, 4))
'      cSetupHrs = Val(vRunOps(iList, 5))
'      cUnitHrs = vRunOps(iList, 6)
'      'new 6/2/00
'      sShop = vRunOps(iList, 0)
'      sCenter = vRunOps(iList, 1)
'      iDay = Format(vSchedDate, "d")
'      sMonth = vSchedDate
'      '11/5/05
'      iSunday = mo.TestWeekEnd(sMonth, "Sun", sShop, sCenter)
'      iSaturday = mo.TestWeekEnd(sMonth, "Sat", sShop, sCenter)
'
'      'Move Hrs
'      bCounter = 0
'      If cMoveHrs > 0 Then
'         'If bQMConvert = 12 Then cMoveHrs = cMoveHrs * 2
'         bGoodCoCal = mo.GetThisCoCalendar(vSchedDate)
'         If bGoodCoCal Then
'            cTestHours = GetQMCalHours(vSchedDate)
'            If cTestHours > cMoveHrs Then
'               vSchedDate = Format(vSchedDate - (cMoveHrs / bQMConvert), vTimeFormat)
'            Else
'               vSchedDate = Format(vSchedDate - (cTestHours / bQMConvert), vTimeFormat)
'               cMoveHrs = cMoveHrs - cTestHours
'               Do Until cMoveHrs <= 0
'                  cTestHours = GetQMCalHours(vSchedDate)
'                  If cTestHours = 0 Then bCounter = bCounter + 1
'                  If bCounter > 3 Then
'                     cTestHours = 8
'                     bCounter = 0
'                  End If
'                  If cTestHours <= cMoveHrs And cTestHours > 0 Then
'                     vSchedDate = Format(vSchedDate - (cTestHours / bQMConvert), vTimeFormat)
'                     cMoveHrs = cMoveHrs - cTestHours
'                  Else
'                     vSchedDate = Format(vSchedDate - (cMoveHrs / bQMConvert), vTimeFormat)
'                     cMoveHrs = 0
'                  End If
'               Loop
'            End If
'         Else
'            vSchedDate = Format(vSchedDate - (cMoveHrs / bQMConvert), vTimeFormat)
'         End If
'      End If
'      If Format(vSchedDate, "ddd") = "Sun" Then vSchedDate = vSchedDate - 2
'      If Format(vSchedDate, "ddd") = "Sat" Then vSchedDate = vSchedDate - 1
'
'      vMoveDate = vSchedDate
'      vRunOps(iList, 10) = vMoveDate
'      'end move
'
'      vRunOps(iList, 7) = vSchedDate
'      'run date
'      If sCenter <> sOldCenter Then
'         iDay = Format(vSchedDate, "d")
'         sMonth = vSchedDate
'         bGoodWCCal = GetThisCalendar(sMonth, sShop, sCenter)
'         If cUnusedHrs > 0 Then
'            If bGoodWCCal Then
'               cTestHours = GetCenterCalHours(sMonth, sShop, sCenter, iDay)
'            Else
'               cTestHours = GetCenterHours(sCenter, iDay)
'            End If
'            If cTestHours < cUnusedHrs Then cUnusedHrs = cTestHours
'            cTestHours = 0
'         End If
'      End If
'      sOldCenter = sCenter
'      If cUnitHrs > 0 Then
'         iDay = Format(vSchedDate, "d")
'         sMonth = vSchedDate
'         bGoodWCCal = GetThisCalendar(sMonth, sShop, sCenter)
'         If bGoodWCCal Then
'            sMonth = Format(vSchedDate, "mmm") & "-" & Format(vSchedDate, "yyyy")
'            If bFirstOp = 1 Then
'               If cUnusedHrs <= 0 Then cUnusedHrs = GetCenterCalHours(sMonth, sShop, sCenter, iDay)
'               bFirstOp = 0
'            End If
'            Do Until cUnitHrs <= 0
'               If cUnusedHrs >= cUnitHrs Then
'                  cUnusedHrs = cUnusedHrs - cUnitHrs
'                  vSchedDate = Format(vSchedDate - (cUnitHrs / 24), "mm/dd/yy hh:mm")
'                  vSchedDate = GetScheduledDate(Format(vSchedDate, "mm/dd/yy hh:mm"), 0)
'                  If iSunday = 0 Then
'                     If Format(vSchedDate, "ddd") = "Sun" Then
'                        If Format(vSchedDate, "hh:mm") < "07:00" Then
'                           vSchedDate = vSchedDate - 0.5
'                        Else
'                           vSchedDate = vSchedDate - 1
'                        End If
'                     End If
'                  End If
'                  If iSaturday = 0 Then
'                     If Format(vSchedDate, "ddd") = "Sat" Then
'                        If Format(vSchedDate, "hh:mm") < "07:00" Then
'                           vSchedDate = vSchedDate - 0.5
'                        Else
'                           vSchedDate = vSchedDate - 1
'                        End If
'                     End If
'                  End If
'                  cUnitHrs = 0
'                  Exit Do
'               Else
'                  cUnitHrs = cUnitHrs - cUnusedHrs
'                  If cUnitHrs < 0 Then cUnitHrs = 0
'                  vSchedDate = Format(vSchedDate - 1, vTimeFormat)
'                  vSchedDate = GetScheduledDate(Format(vSchedDate, "mm/dd/yy hh:mm"), 0)
'                  If iSunday = 0 Then
'                     If Format(vSchedDate, "ddd") = "Sun" Then
'                        If Format(vSchedDate, "hh:mm") < "07:00" Then
'                           vSchedDate = vSchedDate - 0.5
'                        Else
'                           vSchedDate = vSchedDate - 1
'                        End If
'                     End If
'                  End If
'                  If iSaturday = 0 Then
'                     If Format(vSchedDate, "ddd") = "Sat" Then
'                        If Format(vSchedDate, "hh:mm") < "07:00" Then
'                           vSchedDate = vSchedDate - 0.5
'                        Else
'                           vSchedDate = vSchedDate - 1
'                        End If
'                     End If
'                  End If
'                  cUnusedHrs = GetCenterCalHours(sMonth, sShop, sCenter, iDay)
'                  If cUnusedHrs = 0 Then cUnusedHrs = 8
'               End If
'            Loop
'         Else
'            'work center
'            Do Until cUnitHrs <= 0
'               iDay = Format(vSchedDate, "d")
'               cUnusedHrs = GetCenterHours(sCenter, iDay)
'               If cUnusedHrs = 0 Then
'                  For n = 0 To 2
'                     'back up up to 3
'                     vSchedDate = Format(vSchedDate - 1, vTimeFormat)
'                     vSchedDate = GetScheduledDate(Format(vSchedDate, "mm/dd/yy hh:mm"), 0)
'                     sMonth = vSchedDate
'                     iDay = Format(vSchedDate, "w")
'                     cUnusedHrs = GetCenterHours(sCenter, iDay)
'                     If cUnusedHrs > 0 Then
'                        cUnitHrs = cUnitHrs - cUnusedHrs
'                        If cUnitHrs > 0 Then vSchedDate = Format(vSchedDate - 1, vTimeFormat)
'                        Exit For
'                     Else
'                        If n = 2 Then
'                           cUnitHrs = cUnitHrs - 8
'                           If cUnitHrs > 0 Then vSchedDate = Format(vSchedDate - 1, vTimeFormat)
'                        End If
'                     End If
'                  Next
'               Else
'                  cUnitHrs = cUnitHrs - cUnusedHrs
'                  If cUnitHrs > 0 Then
'                     If cUnitHrs > cUnusedHrs Then
'                        vSchedDate = Format(vSchedDate - 1, vTimeFormat)
'                     Else
'                        vSchedDate = Format(vSchedDate - (cUnitHrs / 24), "mm/dd/yy hh:mm")
'                        Exit Do
'                     End If
'                     vSchedDate = GetScheduledDate(Format(vSchedDate, "mm/dd/yy hh:mm"), 0)
'                  End If
'               End If
'            Loop
'         End If
'      End If
'      bDays = GetWeekEnd(Format(vSchedDate, "mmm-dd-yyyy"))
'      vSchedDate = vSchedDate - bDays
'      vRunOps(iList, 8) = vSchedDate
'      If cUnusedHrs < 0 Then cUnusedHrs = 0
'
'      'Setup date
'      If cSetupHrs > 0 Then
'         If bGoodWCCal Then
'            If bFirstOp = 1 Then
'               sMonth = vSchedDate
'               iDay = Format(vSchedDate, "d")
'               If cUnusedHrs <= 0 Then cUnusedHrs = GetCenterCalHours(sMonth, sShop, sCenter, iDay)
'               bFirstOp = 0
'            End If
'
'            sMonth = Format(vSchedDate, "mmm") & "-" & Format(vSchedDate, "yyyy")
'            Do Until cSetupHrs <= 0
'               If cUnusedHrs >= cSetupHrs Then
'                  cUnusedHrs = cUnusedHrs - cSetupHrs
'                  vSchedDate = Format(vSchedDate - (cSetupHrs / 24), "mm/dd/yy hh:mm")
'                  If iSunday = 0 Then
'                     If Format(vSchedDate, "ddd") = "Sun" Then
'                        If Format(vSchedDate, "hh:mm") < "07:00" Then
'                           vSchedDate = vSchedDate - 0.5
'                        Else
'                           vSchedDate = vSchedDate - 1
'                        End If
'                     End If
'                  End If
'                  If iSaturday = 0 Then
'                     If Format(vSchedDate, "ddd") = "Sat" Then
'                        If Format(vSchedDate, "hh:mm") < "07:00" Then
'                           vSchedDate = vSchedDate - 0.5
'                        Else
'                           vSchedDate = vSchedDate - 1
'                        End If
'                     End If
'                  End If
'                  cSetupHrs = 0
'                  Exit Do
'               Else
'                  cSetupHrs = cSetupHrs - cUnusedHrs
'                  If cSetupHrs = 0 Then Exit Do
'                  vSchedDate = Format(vSchedDate - 1, vTimeFormat)
'                  If iSunday = 0 Then
'                     If Format(vSchedDate, "ddd") = "Sun" Then
'                        If Format(vSchedDate, "hh:mm") < "07:00" Then
'                           vSchedDate = vSchedDate - 0.5
'                        Else
'                           vSchedDate = vSchedDate - 1
'                        End If
'                     End If
'                  End If
'                  If iSaturday = 0 Then
'                     If Format(vSchedDate, "ddd") = "Sat" Then
'                        If Format(vSchedDate, "hh:mm") < "07:00" Then
'                           vSchedDate = vSchedDate - 0.5
'                        Else
'                           vSchedDate = vSchedDate - 1
'                        End If
'                     End If
'                  End If
'                  cUnusedHrs = GetCenterCalHours(sMonth, sShop, sCenter, iDay)
'                  If cUnusedHrs = 0 Then cUnusedHrs = 8
'               End If
'            Loop
'         Else
'            'work center
'            Do Until cSetupHrs <= 0
'               sMonth = vSchedDate
'               iDay = Format(vSchedDate, "w")
'               If cUnusedHrs = 0 Then
'                  For n = 0 To 2
'                     'back up up to 3
'                     vSchedDate = Format(vSchedDate - 1, vTimeFormat)
'                     vSchedDate = GetScheduledDate(Format(vSchedDate, "mm/dd/yy hh:mm"), 0)
'                     sMonth = vSchedDate
'                     iDay = Format(vSchedDate, "d")
'                     cUnusedHrs = GetCenterHours(sCenter, iDay)
'                     If cUnusedHrs > 0 Then
'                        cSetupHrs = cSetupHrs - cUnusedHrs
'                        If cSetupHrs > 0 Then vSchedDate = Format(vSchedDate - 1, vTimeFormat)
'                        Exit For
'                     Else
'                        cUnusedHrs = GetCenterHours(sCenter, iDay)
'                        If n = 2 Then
'                           cSetupHrs = cSetupHrs - 8
'                           If cSetupHrs > 0 Then vSchedDate = Format(vSchedDate - 1, vTimeFormat)
'                        End If
'                     End If
'                  Next
'               Else
'                  cSetupHrs = cSetupHrs - cUnusedHrs
'                  If cSetupHrs > 0 Then
'                     If cSetupHrs > cUnusedHrs Then
'                        vSchedDate = Format(vSchedDate - 1, vTimeFormat)
'                     Else
'                        vSchedDate = Format(vSchedDate - (cSetupHrs / 24), "mm/dd/yy hh:mm")
'                        Exit Do
'                     End If
'                  End If
'               End If
'            Loop
'         End If
'      End If
'      If cUnusedHrs < 0 Then cUnusedHrs = 0
'      bDays = GetWeekEnd(Format(vSchedDate, "mmm-dd-yyyy"))
'      vSchedDate = vSchedDate - bDays
'      vRunOps(iList, 9) = Format(vSchedDate, "mm/dd/yy hh:mm")
'      '       End Setup
'
'      bCounter = 0
'      '      'Queue Hours
'      If cQueueHrs > 0 Then
'         bGoodCoCal = mo.GetThisCoCalendar(vSchedDate)
'         If bGoodCoCal Then
'            cTestHours = GetQMCalHours(vSchedDate)
'            If cTestHours > cQueueHrs Then
'               vSchedDate = Format(vSchedDate - (cQueueHrs / bQMConvert), vTimeFormat)
'            Else
'               vSchedDate = Format(vSchedDate - (cTestHours / bQMConvert), vTimeFormat)
'               cQueueHrs = cQueueHrs - cTestHours
'               Do Until cQueueHrs <= 0
'                  cTestHours = GetQMCalHours(vSchedDate)
'                  If cTestHours = 0 Then bCounter = bCounter + 1
'                  If bCounter > 3 Then
'                     cTestHours = 8
'                     bCounter = 0
'                  End If
'                  If cTestHours <= cQueueHrs And cTestHours > 0 Then
'                     vSchedDate = Format(vSchedDate - (cTestHours / bQMConvert), vTimeFormat)
'                     cQueueHrs = cQueueHrs - cTestHours
'                  Else
'                     vSchedDate = Format(vSchedDate - (cTestHours / bQMConvert), vTimeFormat)
'                     cQueueHrs = 0
'                  End If
'               Loop
'            End If
'         Else
'            vSchedDate = Format(vSchedDate - (cQueueHrs / bQMConvert), vTimeFormat)
'         End If
'      End If
'      If Format(vSchedDate, "ddd") = "Sun" Then vSchedDate = vSchedDate - 2
'      If Format(vSchedDate, "ddd") = "Sat" Then vSchedDate = vSchedDate - 1
'      vQueueDate = vSchedDate
'      vRunOps(iList, 11) = vQueueDate
'      'end queue
'   Next
'   vStartDate = vSchedDate
'   'testing
'   On Error GoTo SrvscErr1
'   For iList = 1 To iTotalOps
'      sSql = "UPDATE RnopTable SET OPSCHEDDATE='" & vRunOps(iList, 7) & "'," _
'             & "OPRUNDATE='" & vRunOps(iList, 8) & "'," _
'             & "OPSUDATE='" & vRunOps(iList, 9) & "'," _
'             & "OPMDATE='" & vRunOps(iList, 10) & "', " _
'             & "OPQDATE='" & vRunOps(iList, 11) & "' " _
'             & "WHERE OPREF='" & sPartNumber & "' AND OPRUN=" & lblRun _
'             & " AND OPNO=" & vRunOps(iList, 2) & " "
'      clsAdoCon.ExecuteSQL sSql
'   Next
'   sSql = "UPDATE RunsTable SET RUNSCHED='" & Format(txtCom, "mm/dd/yy 16:00") & "'," _
'          & "RUNPKSTART='" & Format(vStartDate - 1, "mm/dd/yy hh:mm") & "'," _
'          & "RUNSTART='" & Format(vStartDate, "mm/dd/yy 07:00") & "'," _
'          & "RUNQTY=" & Val(txtQty) & ",RUNREMAININGQTY=" & Val(txtQty) & " " _
'          & "WHERE RUNREF='" & sPartNumber & "' " _
'          & "AND RUNNO=" & lblRun & " "
'   clsAdoCon.ExecuteSQL sSql
'
'   On Error Resume Next
'   'Update picks
'   sSql = "UPDATE MopkTable SET PKPDATE='" & vStartDate & "' " _
'          & "WHERE PKMOPART='" & sPartNumber & "' AND PKMORUN=" & lblRun & " " _
'          & "AND PKTYPE<>12 AND PKAQTY=0 "
'   clsAdoCon.ExecuteSQL sSql
'
'   Set RdoOps = Nothing
'   ShopSHe02a.txtQty = txtQty
'   ShopSHe02a.lblSch = txtCom
'
'   Erase vRunOps
'   cOldQty = Val(txtQty)
'   sOLDDATE = txtCom
'   PrgBar.Value = 100
'   MouseCursor 0
'   MsgBox "MO Was Successfully Rescheduled.", vbInformation, Caption
'   PrgBar.Visible = False
'   On Error GoTo 0
'   Exit Sub
'
'SrvscErr1:
'   Resume SrvscErr2
'SrvscErr2:
'   MouseCursor 0
'   MsgBox "Couldn't Reschedule MO.", vbExclamation, Caption
'   PrgBar.Visible = False
'   txtQty = Format(cOldQty, ES_QuantityDataFormat)
'   txtCom = sOLDDATE
'   On Error GoTo 0
'
End Sub


Private Sub txtCom_DropDown()
   ShowCalendarEx Me
     
End Sub

Private Sub txtCom_LostFocus()
   If txtCom = "" Then txtCom = sOLDDATE
   txtCom = CheckDateEx(txtCom)
   
End Sub


Private Sub txtPri_LostFocus()
   txtPri = CheckLen(txtPri, 2)
   txtPri = Abs(Val(txtPri))
   On Error Resume Next
   sSql = "UPDATE RunsTable SET RUNPRIORITY=" & txtPri & " " _
          & "WHERE RUNREF='" & sPartNumber & "' AND RUNNO=" & lblRun & " "
   ShopSHe02a.txtPri = txtPri
   clsADOCon.ExecuteSql sSql
   
End Sub

Private Sub txtQty_LostFocus()
   Dim bResponse As Byte
   Dim sMsg As String
   txtQty = CheckLen(txtQty, 9)
   If Val(txtQty) = 0 Then
      txtQty = Format(cOldQty, ES_QuantityDataFormat)
   Else
      txtQty = Format(Abs(Val(txtQty)), ES_QuantityDataFormat)
   End If
   If Left(ShopSHe02a.lblStat, 1) = "P" Then
      If cOldQty <> Val(txtQty) Then
         sMsg = "You Are Changing The Quantity On A Run That Might " & vbCr _
                & "Impact Picks And Schedules.  If You Continue, Those " & vbCr _
                & "Issues Should Be Addressed. Do You Wish To Continue?"
         Beep
         bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
         If bResponse = vbYes Then
            cOldQty = Val(txtQty)
         Else
            txtQty = Format(cOldQty, ES_QuantityDataFormat)
            CancelTrans
         End If
      End If
   Else
      cOldQty = Val(txtQty)
   End If
   ShopSHe02a.txtQty = txtQty
   sSql = "UPDATE RunsTable SET RUNQTY=" & cOldQty & ", RUNREMAININGQTY=" & cOldQty & _
          " WHERE RUNREF='" & sPartNumber & "' AND RUNNO=" & lblRun & " "
   ShopSHe02a.lblDiv = cmbDiv
   clsADOCon.ExecuteSql sSql
   
End Sub




Public Sub GetSetting()
   Dim RdoShp As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT AllowMOQuantityChanges FROM Preferences WHERE PreRecord=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoShp, ES_FORWARD)
   If bSqlRows Then
      With RdoShp
         optChg.Value = !AllowMOQuantityChanges
         ClearResultSet RdoShp
      End With
   End If
   '6/3/04
   On Error Resume Next
   If optChg.Value = vbChecked Then
      txtQty.Enabled = True
      txtQty.BackColor = Es_TextBackColor
      txtQty.SetFocus
   Else
      If Left(ShopSHe02a.lblStat, 1) = "P" Then
         txtQty.TabStop = False
         txtQty.Locked = True
         txtQty.BackColor = Es_TextDisabled
         txtCom.SetFocus
      End If
   End If
   Set RdoShp = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getsetting"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Function GetWeekEnd(TestDate As Variant) As Byte
   Dim RdoWe As ADODB.Recordset
   GetWeekEnd = 0
   If Left(Format(TestDate, "ddd"), 1) <> "S" Then Exit Function
   
   sSql = "SELECT SUM(COCSHT1+COCSHT2+COCSHT3+COCSHT4) AS AvailHours FROM " _
          & "CoclTable WHERE COCREF='" & Left(TestDate, 3) & "-" _
          & Right(TestDate, 4) & " ' AND COCDAY=" & Val(Mid(TestDate, 5, 2)) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoWe, ES_FORWARD)
   If Not IsNull(RdoWe!AvailHours) Then GetWeekEnd = RdoWe!AvailHours
   
   If GetWeekEnd = 0 Then
      If Format(TestDate, "ddd") = "Sun" Then GetWeekEnd = 2 _
                Else GetWeekEnd = 1
   End If
   Set RdoWe = Nothing
   
End Function

'11/3/05

''Private Sub ForwardSchedule()
''   Dim RdoOps As ADODB.Recordset
''   Dim bDays As Byte
''   Dim bFirstOp As Byte
''   Dim A As Integer
''   Dim iList As Integer
''   Dim n As Integer
''   Dim iDay As Integer
''   Dim iTotalOps As Integer
''   Dim d As Long
''
''   Dim cOpHrs As Currency
''   Dim cMoveHrs As Currency
''   Dim cQueueHrs As Currency
''   Dim cRunqty As Currency
''   Dim cSetupHrs As Currency
''   Dim cUnitHrs As Currency
''   Dim cUnusedHrs As Currency
''   Dim cTestHours As Currency
''
''   Dim vSchedDate As Date
''   Dim vStartDate As Date
''   Dim vMoveDate As Date
''   Dim vQueueDate As Date
''
''   Dim sMonth As String
''   Dim sCenter As String
''   Dim sShop As String
''   Dim sOldCenter As String
''
''   Dim vRunOps(300, 12) As Variant
''   '0 = Center
''   '1 = Shop
''   '2 = OPNO
''   '3 = QHRS
''   '4 = MHRS
''   '5 = OPSUHRS
''   '6 = OPUNITHRS
''   '7 = Sched Complete
''   '8 = Run Start Date
''   '9 = SU Start Date
''   '10 = Move Date
''   '11 = Queue Date
''   MouseCursor 13
''   'Initialize variant data
''   For iList = 1 To 298
''      vRunOps(iList, 1) = ""
''      vRunOps(iList, 2) = ""
''      vRunOps(iList, 3) = 0
''      vRunOps(iList, 4) = 0
''      vRunOps(iList, 5) = 0
''      vRunOps(iList, 6) = 0
''      vRunOps(iList, 7) = Format(ES_SYSDATE, "mm/dd/yy")
''      vRunOps(iList, 8) = Format(ES_SYSDATE, "mm/dd/yy")
''      vRunOps(iList, 9) = Format(ES_SYSDATE, "mm/dd/yy")
''   Next
''   vRunOps(iList, 1) = ""
''   vRunOps(iList, 2) = ""
''   vRunOps(iList, 3) = 0
''   vRunOps(iList, 4) = 0
''   vRunOps(iList, 5) = 0
''   vRunOps(iList, 6) = 0
''   vRunOps(iList, 7) = Format(ES_SYSDATE, "mm/dd/yy")
''   vRunOps(iList, 8) = Format(ES_SYSDATE, "mm/dd/yy")
''   vRunOps(iList, 9) = Format(ES_SYSDATE, "mm/dd/yy")
''
''   cRunqty = Val(txtQty)
''   'get distinct Shops,Work Centers
''   PrgBar.Value = 10
''   On Error GoTo SrvscErr1
''
''   PrgBar.Value = 20
''   'get the workcenters and ops in reverse order
''   sSql = "SELECT OPREF,OPRUN,OPNO,OPSHOP,OPCENTER,OPQHRS,OPMHRS,OPSUHRS,OPUNITHRS" _
''          & " FROM RnopTable WHERE OPREF='" & sPartNumber & "' AND OPRUN=" & lblRun _
''          & " ORDER BY OPNO"
''   bsqlrows = clsadocon.getdataset(ssql, RdoOps, ES_STATIC)
''   If bSqlRows Then
''      With RdoOps
''         Do Until .EOF
''            iTotalOps = iTotalOps + 1
''            vRunOps(iTotalOps, 0) = "" & Trim(!OPSHOP)
''            vRunOps(iTotalOps, 1) = "" & Trim(!OPCENTER)
''            vRunOps(iTotalOps, 2) = !opNo
''            vRunOps(iTotalOps, 3) = !OPQHRS
''            vRunOps(iTotalOps, 4) = !OPMHRS
''            vRunOps(iTotalOps, 5) = !OPSUHRS
''            vRunOps(iTotalOps, 6) = !OPUNITHRS * cRunqty
''            .MoveNext
''         Loop
''         ClearResultSet RdoOps
''      End With
''      vRunOps(1, 7) = Format(txtCom, "mm/dd/yy")
''   Else
''      MouseCursor 0
''      MsgBox "No Operations To Schedule.", vbExclamation, Caption
''      Exit Sub
''   End If
''
''   PrgBar.Value = 40
''
''   'Get Workcenter Calendars and time
''   vRunOps(0, 4) = vRunOps(1, 4)
''   'Release the memory
''
''   PrgBar.Value = 50
''
''   'Schedule 'em
''   'We'll Allow a Calendar or Work Center Hours.
''   'If neither then they are bumming
''   On Error Resume Next
''   vSchedDate = Format(txtCom, "mm/dd/yy 7:00")
''   bDays = GetWeekEnd(Format(vSchedDate, "mmm-dd-yyyy"))
''   vSchedDate = vSchedDate + bDays
''   cTestHours = 0
''   sOldCenter = ""
''   cUnusedHrs = 0
''   bFirstOp = 1
''   For iList = 1 To iTotalOps 'one op at a time
''      ' Allow move 6/6/00
''      If cUnusedHrs < 0 Then cUnusedHrs = 0
''      If PrgBar.Value < 100 Then PrgBar.Value = PrgBar.Value + 10
''      cQueueHrs = Val(vRunOps(iList, 3))
''      cMoveHrs = Val(vRunOps(iList, 4))
''      cSetupHrs = Val(vRunOps(iList, 5))
''      cUnitHrs = vRunOps(iList, 6)
''      'new 6/2/00
''      sShop = vRunOps(iList, 0)
''      sCenter = vRunOps(iList, 1)
''      iDay = Format(vSchedDate, "d")
''      sMonth = vSchedDate
''      '        'Queue Hours
''      If cQueueHrs > 0 Then
''         bGoodCoCal = GetThisCoCalendar(vSchedDate)
''         If bGoodCoCal Then
''            cTestHours = GetQMCalHours(vSchedDate)
''            If cTestHours > cQueueHrs Then
''               vSchedDate = Format(vSchedDate + (cQueueHrs / 24), vTimeFormat)
''            Else
''               vSchedDate = Format(vSchedDate + (cTestHours / 24), vTimeFormat)
''               cQueueHrs = cQueueHrs - cTestHours
''               Do Until cQueueHrs <= 0
''                  cTestHours = GetQMCalHours(vSchedDate)
''                  If cTestHours > cQueueHrs Then
''                     vSchedDate = Format(vSchedDate + (cQueueHrs / 24), vTimeFormat)
''                     cQueueHrs = cQueueHrs - cTestHours
''                  Else
''                     vSchedDate = Format(vSchedDate + (cTestHours / 24), vTimeFormat)
''                     cQueueHrs = 0
''                  End If
''               Loop
''            End If
''         Else
''            vSchedDate = Format(vSchedDate + (cQueueHrs / 24), vTimeFormat)
''         End If
''      End If
''      If Format(vSchedDate, "ddd") = "Sun" Then vSchedDate = vSchedDate + 1
''      If Format(vSchedDate, "ddd") = "Sat" Then vSchedDate = vSchedDate + 2
''      vQueueDate = vSchedDate
''      vRunOps(iList, 11) = vQueueDate
''      'end queue
''
''      'Setup date
''      vRunOps(iList, 7) = vSchedDate
''      vRunOps(iList, 9) = vSchedDate
''      If cSetupHrs > 0 Then
''         If bGoodWCCal Then
''            If bFirstOp = 1 Then
''               sMonth = vSchedDate
''               iDay = Format(vSchedDate, "d")
''               If cUnusedHrs <= 0 Then cUnusedHrs = GetCenterCalHours(sMonth, sShop, sCenter, iDay)
''               bFirstOp = 0
''            End If
''            Do Until cSetupHrs = 0
''               If cUnusedHrs >= cSetupHrs Then
''                  cUnusedHrs = cUnusedHrs + cSetupHrs
''                  vSchedDate = Format(vSchedDate + (cSetupHrs / 24), "mm/dd/yy hh:mm")
''                  vSchedDate = GetScheduledDate(vSchedDate, 0)
''                  sMonth = vSchedDate
''                  iDay = Format(vSchedDate, "d")
''                  If A > 0 Then vSchedDate = Format(vSchedDate + A, vTimeFormat)
''                  cSetupHrs = 0
''                  Exit Do
''               Else
''                  cSetupHrs = cSetupHrs + cUnusedHrs
''                  If cSetupHrs = 0 Then Exit Do
''                  vSchedDate = Format(vSchedDate + 1, vTimeFormat)
''                  vSchedDate = GetScheduledDate(vSchedDate, 0)
''                  sMonth = vSchedDate
''                  iDay = Format(vSchedDate, "d")
''                  If A > 0 Then vSchedDate = Format(vSchedDate + A, vTimeFormat)
''                  sMonth = vSchedDate
''                  iDay = Format(vSchedDate, "d")
''                  cUnusedHrs = GetCenterCalHours(sMonth, sShop, sCenter, iDay)
''                  If cUnusedHrs = 0 Then cUnusedHrs = 8
''               End If
''            Loop
''         Else
''            'work center
''            Do Until cSetupHrs <= 0
''               sMonth = vSchedDate
''               iDay = Format(vSchedDate, "w")
''               If cUnusedHrs = 0 Then
''                  For n = 0 To 2
''                     'back up up to 3
''                     vSchedDate = Format(vSchedDate + 1, vTimeFormat)
''                     vSchedDate = GetScheduledDate(vSchedDate, 0)
''                     sMonth = vSchedDate
''                     iDay = Format(vSchedDate, "d")
''                     cUnusedHrs = GetCenterHours(sCenter, iDay)
''                     If cUnusedHrs > 0 Then
''                        cSetupHrs = cSetupHrs + cUnusedHrs
''                        If cSetupHrs > 0 Then vSchedDate = Format(vSchedDate + 1, vTimeFormat)
''                        Exit For
''                     Else
''                        cUnusedHrs = GetCenterHours(sCenter, iDay)
''                        If n = 2 Then
''                           cSetupHrs = cSetupHrs + 8
''                           If cSetupHrs > 0 Then vSchedDate = Format(vSchedDate + 1, vTimeFormat)
''                        End If
''                     End If
''                  Next
''               Else
''                  cSetupHrs = cSetupHrs + cUnusedHrs
''                  If cSetupHrs > 0 Then
''                     If cSetupHrs > cUnusedHrs Then
''                        vSchedDate = Format(vSchedDate + 1, vTimeFormat)
''                     Else
''                        vSchedDate = Format(vSchedDate + (cSetupHrs / 24), "mm/dd/yy hh:mm")
''                        Exit Do
''                     End If
''                  End If
''               End If
''            Loop
''         End If
''      End If
''      If cUnusedHrs < 0 Then cUnusedHrs = 0
''      bDays = GetWeekEnd(Format(vSchedDate, "mmm-dd-yyyy"))
''      vSchedDate = vSchedDate + bDays
''      '       End Setup
''
''      'run date
''      vRunOps(iList, 8) = vSchedDate
''      If sCenter <> sOldCenter Then
''         iDay = Format(vSchedDate, "d")
''         sMonth = vSchedDate
''         bGoodWCCal = GetThisCalendar(sMonth, sShop, sCenter)
''         If cUnusedHrs > 0 Then
''            If bGoodWCCal Then
''               cTestHours = GetCenterCalHours(sMonth, sShop, sCenter, iDay)
''            Else
''               cTestHours = GetCenterHours(sCenter, iDay)
''            End If
''            If cTestHours < cUnusedHrs Then cUnusedHrs = cTestHours
''            cTestHours = 0
''         End If
''      End If
''      sOldCenter = sCenter
''      If cUnitHrs > 0 Then
''         iDay = Format(vSchedDate, "d")
''         sMonth = vSchedDate
''         bGoodWCCal = GetThisCalendar(sMonth, sShop, sCenter)
''         If bGoodWCCal Then
''            If bFirstOp = 1 Then
''               If cUnusedHrs <= 0 Then cUnusedHrs = GetCenterCalHours(sMonth, sShop, sCenter, iDay)
''               bFirstOp = 0
''            End If
''            Do Until cUnitHrs = 0
''               If cUnusedHrs >= cUnitHrs Then
''                  cUnusedHrs = cUnusedHrs - cUnitHrs
''                  vSchedDate = Format(vSchedDate + (cUnitHrs / 24), "mm/dd/yy hh:mm")
''                  vSchedDate = GetScheduledDate(vSchedDate, 0)
''                  sMonth = vSchedDate
''                  iDay = Format(vSchedDate, "d")
''                  If A > 0 Then vSchedDate = Format(vSchedDate + A, vTimeFormat)
''                  cUnitHrs = 0
''                  Exit Do
''               Else
''                  cUnitHrs = cUnitHrs - cUnusedHrs
''                  If cUnitHrs < 0 Then cUnitHrs = 0
''                  vSchedDate = Format(vSchedDate + 1, vTimeFormat)
''                  vSchedDate = GetScheduledDate(vSchedDate, 0)
''                  sMonth = vSchedDate
''                  iDay = Format(vSchedDate, "d")
''                  If A > 0 Then vSchedDate = Format(vSchedDate + A, vTimeFormat)
''                  sMonth = vSchedDate
''                  iDay = Format(vSchedDate, "d")
''                  cUnusedHrs = GetCenterCalHours(sMonth, sShop, sCenter, iDay)
''                  If cUnusedHrs = 0 Then cUnusedHrs = 8
''               End If
''            Loop
''         Else
''            'work center
''            Do Until cUnitHrs <= 0
''               iDay = Format(vSchedDate, "d")
''               cUnusedHrs = GetCenterHours(sCenter, iDay)
''               If cUnusedHrs = 0 Then
''                  For n = 0 To 2
''                     'back up up to 3
''                     vSchedDate = Format(vSchedDate + 1, vTimeFormat)
''                     vSchedDate = GetScheduledDate(vSchedDate, 0)
''                     sMonth = vSchedDate
''                     iDay = Format(vSchedDate, "w")
''                     cUnusedHrs = GetCenterHours(sCenter, iDay)
''                     If cUnusedHrs > 0 Then
''                        cUnitHrs = cUnitHrs - cUnusedHrs
''                        If cUnitHrs > 0 Then vSchedDate = Format(vSchedDate + 1, vTimeFormat)
''                        Exit For
''                     Else
''                        If n = 2 Then
''                           cUnitHrs = cUnitHrs - 8
''                           If cUnitHrs > 0 Then vSchedDate = Format(vSchedDate + 1, vTimeFormat)
''                        End If
''                     End If
''                  Next
''               Else
''                  cUnitHrs = cUnitHrs - cUnusedHrs
''                  If cUnitHrs > 0 Then
''                     If cUnitHrs > cUnusedHrs Then
''                        vSchedDate = Format(vSchedDate + 1, vTimeFormat)
''                     Else
''                        vSchedDate = Format(vSchedDate + (cUnitHrs / 24), "mm/dd/yy hh:mm")
''                        Exit Do
''                     End If
''                     vSchedDate = GetScheduledDate(vSchedDate, 0)
''                  End If
''               End If
''            Loop
''         End If
''      End If
''      bDays = GetWeekEnd(Format(vSchedDate, "mmm-dd-yyyy"))
''      vSchedDate = vSchedDate + bDays
''      If cUnusedHrs < 0 Then cUnusedHrs = 0
''      'End Run
''      'Move Hrs
''      If cMoveHrs > 0 Then
''         bGoodCoCal = GetThisCoCalendar(vSchedDate)
''         If bGoodCoCal Then
''            cTestHours = GetQMCalHours(vSchedDate)
''            If cTestHours > cMoveHrs Then
''               vSchedDate = Format(vSchedDate + (cMoveHrs / 24), vTimeFormat)
''            Else
''               vSchedDate = Format(vSchedDate + (cTestHours / 24), vTimeFormat)
''               cMoveHrs = cMoveHrs - cTestHours
''               Do Until cMoveHrs <= 0
''                  cTestHours = GetQMCalHours(vSchedDate)
''                  If cTestHours > cMoveHrs Then
''                     vSchedDate = Format(vSchedDate + (cMoveHrs / 24), vTimeFormat)
''                     cMoveHrs = cMoveHrs - cTestHours
''                  Else
''                     vSchedDate = Format(vSchedDate + (cTestHours / 24), vTimeFormat)
''                     cMoveHrs = 0
''                  End If
''               Loop
''            End If
''         Else
''            vSchedDate = Format(vSchedDate + (cMoveHrs / 24), vTimeFormat)
''         End If
''      End If
''      If Format(vSchedDate, "ddd") = "Sun" Then vSchedDate = vSchedDate + 1
''      If Format(vSchedDate, "ddd") = "Sat" Then vSchedDate = vSchedDate + 2
''      vMoveDate = vSchedDate
''      vRunOps(iList, 10) = vMoveDate
''      'end move
''   Next
''
''   vStartDate = vSchedDate
''   'testing
''   On Error GoTo SrvscErr1
''   For iList = 1 To iTotalOps
''      sSql = "UPDATE RnopTable SET OPSCHEDDATE='" & vRunOps(iList, 7) & "'," _
''             & "OPRUNDATE='" & vRunOps(iList, 8) & "'," _
''             & "OPSUDATE='" & vRunOps(iList, 9) & "'," _
''             & "OPMDATE='" & vRunOps(iList, 10) & "', " _
''             & "OPQDATE='" & vRunOps(iList, 11) & "' " _
''             & "WHERE OPREF='" & sPartNumber & "' AND OPRUN=" & lblRun _
''             & " AND OPNO=" & vRunOps(iList, 2) & " "
''      clsAdoCon.ExecuteSQL sSql
''   Next
''   txtCom = Format(vStartDate, "mm/dd/yy")
''   ShopSHe02a.lblSch = txtCom
''   sSql = "UPDATE RunsTable SET RUNSCHED='" & Format(vStartDate, "mm/dd/yy 16:00") & "'," _
''          & "RUNPKSTART='" & Format(vStartDate + 1, "mm/dd/yy hh:mm") & "'," _
''          & "RUNSTART='" & Format(txtCom, "mm/dd/yy 07:00") & "'," _
''          & "RUNQTY=" & Val(txtQty) & ",RUNREMAININGQTY=" & Val(txtQty) & " " _
''          & "WHERE RUNREF='" & sPartNumber & "' " _
''          & "AND RUNNO=" & lblRun & " "
''   clsAdoCon.ExecuteSQL sSql
''
''   On Error Resume Next
''   'Update picks
''   sSql = "UPDATE MopkTable SET PKPDATE='" & vStartDate & "' " _
''          & "WHERE PKMOPART='" & sPartNumber & "' AND PKMORUN=" & lblRun & " " _
''          & "AND PKTYPE<>12 AND PKAQTY=0 "
''   clsAdoCon.ExecuteSQL sSql
''
''   Set RdoOps = Nothing
''   ShopSHe02a.txtQty = txtQty
''   ShopSHe02a.lblSch = txtCom
''
''   Erase vRunOps
''   cOldQty = Val(txtQty)
''   sOLDDATE = txtCom
''   PrgBar.Value = 100
''   MouseCursor 0
''   MsgBox "MO Was Successfully Rescheduled.", vbInformation, Caption
''   PrgBar.Visible = False
''   On Error GoTo 0
''   Exit Sub
''
''SrvscErr1:
''   Resume SrvscErr2
''SrvscErr2:
''   MouseCursor 0
''   MsgBox "Couldn't Reschedule MO.", vbExclamation, Caption
''   PrgBar.Visible = False
''   txtQty = Format(cOldQty, ES_QuantityDataFormat)
''   txtCom = sOLDDATE
''   On Error GoTo 0
''End Sub
