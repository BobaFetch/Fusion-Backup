VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form ShopSHe02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revise a Manufacturing Order"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   4102
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdStatCode 
      DisabledPicture =   "ShopSHe02a.frx":0000
      DownPicture     =   "ShopSHe02a.frx":0972
      Height          =   315
      Left            =   5040
      MaskColor       =   &H8000000F&
      Picture         =   "ShopSHe02a.frx":0E01
      Style           =   1  'Graphical
      TabIndex        =   43
      TabStop         =   0   'False
      ToolTipText     =   "Add Internal Status Code"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   350
   End
   Begin VB.CommandButton cmdOpt 
      Caption         =   "Revise Inspection Document"
      Enabled         =   0   'False
      Height          =   300
      Index           =   12
      Left            =   1020
      TabIndex        =   39
      ToolTipText     =   "Foward Schedule"
      Top             =   6840
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.CommandButton cmdOpt 
      Caption         =   "Link to a Higher Level MOs"
      Enabled         =   0   'False
      Height          =   300
      Index           =   8
      Left            =   960
      TabIndex        =   38
      Top             =   5160
      Width           =   4500
   End
   Begin VB.CommandButton cmdOpt 
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   37
      Top             =   7320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox optNwr 
      Caption         =   "New Rte"
      Height          =   255
      Left            =   240
      TabIndex        =   36
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optfrom 
      Caption         =   "New"
      Height          =   255
      Left            =   5400
      TabIndex        =   35
      Top             =   6480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox optAll 
      Caption         =   " Allocations"
      Height          =   255
      Left            =   4440
      TabIndex        =   34
      Top             =   6480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox optSrv 
      Caption         =   "Srv"
      Height          =   255
      Left            =   3600
      TabIndex        =   33
      Top             =   6480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox optRte 
      Caption         =   "Rte"
      Height          =   195
      Left            =   2760
      TabIndex        =   32
      Top             =   6480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox optCmt 
      Caption         =   "cmt"
      Height          =   255
      Left            =   1920
      TabIndex        =   31
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optBud 
      Caption         =   "Budgets"
      Height          =   255
      Left            =   1080
      TabIndex        =   30
      Top             =   6480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdOpt 
      Caption         =   "Revise Schedule, Date, Quantity, Priority"
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      Left            =   960
      TabIndex        =   3
      ToolTipText     =   "Reschedule, Revise Quantity"
      Top             =   2280
      Width           =   4500
   End
   Begin VB.CommandButton cmdOpt 
      Caption         =   "Revise Routing Information, Comments"
      Enabled         =   0   'False
      Height          =   300
      Index           =   2
      Left            =   960
      TabIndex        =   4
      ToolTipText     =   "All Routing Information"
      Top             =   2640
      Width           =   4500
   End
   Begin VB.CommandButton cmdOpt 
      Caption         =   "Enter/Revise Sales Order Allocations"
      Enabled         =   0   'False
      Height          =   300
      Index           =   3
      Left            =   960
      TabIndex        =   5
      ToolTipText     =   "SO Allocations"
      Top             =   3000
      Width           =   4500
   End
   Begin VB.CommandButton cmdOpt 
      Caption         =   "Enter/Revise Manufacturing Order Comments"
      Enabled         =   0   'False
      Height          =   300
      Index           =   4
      Left            =   960
      TabIndex        =   6
      ToolTipText     =   "MO Comments"
      Top             =   3360
      Width           =   4500
   End
   Begin VB.CommandButton cmdOpt 
      Caption         =   "Print Or Display A Pick List"
      Enabled         =   0   'False
      Height          =   300
      Index           =   5
      Left            =   960
      TabIndex        =   9
      ToolTipText     =   "Print Or Display The Manufacturing Order Pick List"
      Top             =   4440
      Width           =   4500
   End
   Begin VB.CommandButton cmdOpt 
      Caption         =   "Print Or Display Manufacturing Order"
      Enabled         =   0   'False
      Height          =   300
      Index           =   6
      Left            =   960
      TabIndex        =   7
      ToolTipText     =   "Print Or Display The Manufacturing Order"
      Top             =   3720
      Width           =   4500
   End
   Begin VB.CommandButton cmdOpt 
      Caption         =   "Forward Schedule Manufacturing Order"
      Height          =   300
      Index           =   7
      Left            =   960
      TabIndex        =   23
      ToolTipText     =   "Operation Comments"
      Top             =   5520
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.CommandButton cmdOpt 
      Caption         =   "Release This MO To Production"
      Height          =   300
      Index           =   9
      Left            =   960
      TabIndex        =   10
      ToolTipText     =   "SC Manufacturing Orders Marke Released (RL)"
      Top             =   4800
      Width           =   4500
   End
   Begin VB.CommandButton cmdOpt 
      Caption         =   "Revise Inspection Document"
      Enabled         =   0   'False
      Height          =   300
      Index           =   10
      Left            =   960
      TabIndex        =   22
      Top             =   5880
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.CommandButton cmdOpt 
      Appearance      =   0  'Flat
      Caption         =   "Enter Manufacturing Order Budgets"
      Enabled         =   0   'False
      Height          =   300
      Index           =   11
      Left            =   960
      TabIndex        =   8
      ToolTipText     =   "Manufacturing Order Budgets"
      Top             =   4080
      Width           =   4500
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1800
      Width           =   1035
   End
   Begin VB.TextBox txtPri 
      Height          =   285
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1800
      Width           =   375
   End
   Begin VB.Frame z2 
      Height          =   40
      Left            =   240
      TabIndex        =   20
      Top             =   1680
      Width           =   6135
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHe02a.frx":1290
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optPick 
      Caption         =   "Picks"
      Height          =   255
      Left            =   4200
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbRun 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      ToolTipText     =   "Select Run Number"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      ToolTipText     =   "Select Part Number"
      Top             =   600
      Width           =   3545
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5520
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6000
      Top             =   1080
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5580
      FormDesignWidth =   6435
   End
   Begin VB.Label Label1 
      Caption         =   "Type"
      Height          =   255
      Left            =   5460
      TabIndex        =   42
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5880
      TabIndex        =   41
      Top             =   600
      Width           =   375
   End
   Begin VB.Label z13 
      Caption         =   "See Below"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   60
      TabIndex        =   40
      Top             =   5160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Mo Quantity"
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   29
      Top             =   1800
      Width           =   1092
   End
   Begin VB.Label lblUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   2400
      TabIndex        =   28
      Top             =   1800
      Width           =   372
   End
   Begin VB.Label lblSch 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   4200
      TabIndex        =   27
      Top             =   1800
      Width           =   950
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Sched Complete"
      Height          =   252
      Index           =   14
      Left            =   2880
      TabIndex        =   26
      Top             =   1800
      Width           =   1332
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Priority"
      Height          =   252
      Index           =   16
      Left            =   5280
      TabIndex        =   25
      Top             =   1800
      Width           =   612
   End
   Begin VB.Label lblDiv 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   5880
      TabIndex        =   24
      Top             =   2160
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Label lblFrom 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "(Excludes Status CA)"
      Height          =   255
      Index           =   15
      Left            =   3480
      TabIndex        =   16
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   15
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label lblStat 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2520
      TabIndex        =   14
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   13
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "ShopSHe02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'4/1/05 Trap to close if no Company Calendar
'10/6/06 Another shot and possible solution to passing to PickList
Option Explicit
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter

Dim bGoodCal As Boolean
Dim bGoodCoCal As Byte
Dim bCanceled As Boolean
Dim bGoodPart As Byte
Dim bGoodMo As Byte
Dim bOnLoad As Boolean
Dim bPrint As Byte
Dim bPickList As Byte
Dim bFromNew As Byte

Dim sPartNumber As String
Private cmdOpt_IndexClicked As Integer    'index of cmdOpt button clicked.  0 = none

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtPri = "99"
   lblSch.BackColor = Es_TextDisabled
   txtPri.BackColor = Es_TextDisabled
   txtQty.BackColor = Es_TextDisabled
   
End Sub



Private Sub cmbPrt_Click()
   bGoodPart = GetPart(True)
   
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If bCanceled Then Exit Sub
   If Len(Trim(cmbPrt)) > 0 Then bGoodPart = GetPart(False)
   
End Sub

Private Sub cmbRun_Click()
   bGoodMo = GetRun()
   
End Sub

Private Sub cmbRun_LostFocus()
   cmbRun = CheckLen(cmbRun, 5)
   cmbRun = Format(Abs(Val(cmbRun)), "####0")
   If Val(cmbRun) > 32767 Then cmbRun = "32767"
   bGoodMo = GetRun()
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub


Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCanceled = True
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4102
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdOpt_Click(Index As Integer)
   cmdOpt_IndexClicked = Index
   optNwr = vbUnchecked
   Select Case Index
      Case 1
         optSrv.Value = vbChecked
         '            6/3/04
         '            If Left(lblStat, 1) = "P" Or Left(lblStat, 1) = "C" Then
         '                ShopSHe02b.txtQty.Enabled = False
         '                ShopSHe02b.txtQty.BackColor = Es_TextDisabled
         '            End If
         ShopSHe02b.Show
      Case 2
         optRte.Value = vbChecked
         ShopSHe02c.Show
      Case 3
         optAll.Value = vbChecked
         ShopSHe02d.lblMon = cmbPrt
         ShopSHe02d.lblRun = cmbRun
         ShopSHe02d.lblRqty = txtQty
         ShopSHe02d.Show
      Case 4
         OptCmt.Value = vbChecked
         ShopSHe02e.Show
      Case 5
         bPickList = 1
         PickMCp01a.optFrom = vbChecked
         PickMCp01a.lblMon = cmbPrt
         PickMCp01a.lblRun = cmbRun
         PickMCp01a.Show
         Hide
      Case 6
         bPrint = 1
         If ES_CUSTOM = "WATERJET" Then
            awiShopSHp01a.cmbPrt = cmbPrt
            awiShopSHp01a.optFrom.Value = vbChecked
            awiShopSHp01a.cmbRun.SetFocus
            awiShopSHp01a.Show
'         ElseIf ES_CUSTOM = "JEVCO" Then
'            jevShopSHp01a.cmbPrt = cmbPrt
'            jevShopSHp01a.optFrom.Value = vbChecked
'            jevShopSHp01a.cmbRun.SetFocus
'            jevShopSHp01a.Show
         Else
            ShopSHp01a.cmbPrt = cmbPrt
            ShopSHp01a.optFrom.Value = vbChecked
            ShopSHp01a.cmbRun.SetFocus
            ShopSHp01a.Show
         End If
         Hide
      Case 8
      'pick to higher level MO not used
'         ShopSHe02h.lblLowerMoPart = cmbPrt
'         ShopSHe02h.lblLowerMoRun = cmbRun
'         ShopSHe02h.lblLowerMoQty = txtQty
'         ShopSHe02h.lblLowerMoDescription = Me.lblDsc
'         ShopSHe02h.lblLowerMoStatus = lblStat
'         ShopSHe02h.lblLowerType = lblType
'         ShopSHe02h.Show
         ShopSHe02i.lblLowerMoPart = cmbPrt
         ShopSHe02i.lblLowerMoRun = cmbRun
         ShopSHe02i.lblLowerMoQty = txtQty
         ShopSHe02i.lblLowerMoDescription = Me.lblDsc
         ShopSHe02i.lblLowerMoStatus = lblStat
         ShopSHe02i.lblLowerType = lblType
         ShopSHe02i.Show

      Case 9
         ReleaseMo
      Case 11
         optBud.Value = vbChecked
         ShopSHe02f.Show
      Case Else
         MsgBox "Function Is Not Enabled.", vbInformation, Caption
   End Select
   
End Sub

Private Sub cmdStatCode_Click()
 StatusCode.lblSCTypeRef = "MO PartNumber"
 StatusCode.txtSCTRef = Compress(cmbPrt.Text)
 StatusCode.LableRef1 = "Run"
 StatusCode.lblSCTRef1 = cmbRun.Text
 StatusCode.lblSCTRef2 = ""
 StatusCode.lblStatType = "MO"
 StatusCode.lblSysCommIndex = 8 ' The index in the Sys Comment "MO Comments"
 StatusCode.txtCurUser = cUR.CurrentUser

 StatusCode.Show

End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   
   Dim buttonIndex As Integer
   buttonIndex = cmdOpt_IndexClicked
   cmdOpt_IndexClicked = 0
   bFromNew = 0
   Select Case buttonIndex
   Case 8
      Unload ShopSHe02i
   End Select
   
   
   If optBud.Value = vbChecked Then
      Unload ShopSHe02f
      optBud.Value = vbUnchecked
   End If
   If OptCmt.Value = vbChecked Then
      Unload ShopSHe02e
      OptCmt.Value = vbUnchecked
   End If
   If optSrv.Value = vbChecked Then
      Unload ShopSHe02b
      optSrv.Value = vbUnchecked
   End If
   
   If optRte.Value = vbChecked Then
      Unload ShopSHe02c
      optRte.Value = vbUnchecked
   End If
   If optAll.Value = vbChecked Then
      Unload ShopSHe02d
      optAll.Value = vbUnchecked
   End If
   If optFrom.Value = vbChecked Then
      Caption = "New Manufacturing Order"
      lblFrom = ShopSHe01a.cmbPrt
      cmbPrt = ShopSHe01a.cmbPrt
      cmbRun = ShopSHe01a.lblRun
      cmbRun.Enabled = False
      cmbPrt.Enabled = False
      DoEvents
      bGoodPart = GetPart(True)
      bGoodMo = GetRun()
      optFrom = vbUnchecked
      bOnLoad = 0
      bFromNew = 1
      Unload ShopSHe01a
   Else
      If optNwr.Value = vbChecked Then
         lblFrom = ShopSHf03a.cmbPrt
         cmbPrt = ShopSHf03a.cmbPrt
         cmbRun = ShopSHf03a.cmbRun
         bGoodPart = GetPart(True)
         bGoodMo = GetRun()
         bOnLoad = 0
         Unload ShopSHf03a
         optNwr.Value = vbUnchecked
      Else
         If sPassedMo <> "" Then cmbPrt = sPassedMo
      End If
   End If
   If optPick.Value = vbChecked Then
      bOnLoad = 0
      optPick.Value = vbUnchecked
      '           lblFrom = PickMCp01a.cmbPrt
      '           cmbPrt = PickMCp01a.cmbPrt
      '           cmbRun = PickMCp01a.cmbRun
      '       Unload PickMCp01a
      PickMCp01a.Hide
   End If
   If bOnLoad Then
      'FillRuns Me, "NOT LIKE 'CA%'"
      FillRuns Me, "NOT LIKE 'C%'"  'changed 5/8/2019 this is what actually happened in the called subroutine
      bGoodPart = GetPart(True)
      bGoodCal = GetCenterCalendar(Me)
      bGoodCoCal = GetCompanyCalendar()
      If bGoodCoCal = 0 Then
         MsgBox "There Is No Company Calendar For The Period.", _
            vbInformation, Caption
         CapaCPe04a.Show
         Unload Me
         Exit Sub
      End If
      bOnLoad = 0
   
   'if returning from another form, update status
   Else
      UpdateStatus
   End If
   If bPickList = 1 Then
      cmbPrt = PickMCp01a.cmbPrt
      cmbRun = PickMCp01a.cmbRun
      bPickList = 0
   End If
   bPrint = 0
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   bOnLoad = 1
   FormatControls
   
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PAUNITS,PALEVEL,PARUN,RUNREF," _
          & "RUNNO,RUNSTATUS,RUNQTY,RUNSCHED," _
          & "RUNPRIORITY,RUNDIVISION " _
          & "FROM PartTable,RunsTable WHERE PARTREF= ? " _
          & "AND PARTREF=RUNREF AND RUNSTATUS NOT LIKE 'CA%'"   ' changed back 2/5/20
'          & "AND PARTREF=RUNREF AND RUNSTATUS NOT LIKE 'C%'"   ' changed from ca% 5/8/2019

   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Type = adChar
   AdoParameter.SIZE = 30
   
   AdoQry.Parameters.Append AdoParameter
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   
   'Dim I As Integer
   'For i = 1 To Forms.Count - 1
   '    If Forms(i).Name = "PickMCp01a" Then Unload Forms(i)
   'Next
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   If bGoodCoCal = 1 Then
      If bPickList = 0 Then
         If bFromNew = 1 Then ShopSHe01a.Show Else FormUnload bPrint
      End If
   Else
      If bPickList = 0 Then
         If bFromNew = 1 Then ShopSHe01a.Show Else FormUnload
      End If
   End If
   
   On Error Resume Next
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   Set ShopSHe02a = Nothing
   
End Sub


Private Sub lblFrom_Click()
   'Passed from ShopSHe01a or blank
   'tests Run number
   
End Sub

Private Sub optAll_Click()
   'Never visible
   'if True then allocations are visible
   
End Sub

Private Sub optBud_Click()
   'never visible
   
End Sub

Private Sub optCmt_Click()
   'Never visible - check for ShopSHe02e
   
End Sub

Private Sub optFrom_Click()
   'Never visible
   'if True then from New Mo
   'else from Menu. Never visible
   
End Sub

Private Sub optNwr_Click()
   'never visible - New routing
   
End Sub

Private Sub optRte_Click()
   'Never visible - check for ShopSHe02c
   
End Sub

Private Sub txtPri_LostFocus()
   txtPri = CheckLen(txtPri, 2)
   txtPri = Format(Abs(Val(txtPri)), "#0")
   
End Sub


Private Sub txtQty_LostFocus()
   txtQty = CheckLen(txtQty, 9)
   txtQty = Format(Abs(Val(txtQty)), ES_QuantityDataFormat)
   
End Sub



Private Function GetPart(OnLoad As Byte) As Byte
   Dim RdoMon As ADODB.Recordset
   Dim iList As Integer
   Dim C As Currency
   sPartNumber = Compress(cmbPrt)
   
   On Error GoTo DiaErr1
   MouseCursor 13
   GetPart = False
   If optFrom.Value = vbChecked Or optNwr.Value = vbChecked Or optPick.Value = vbChecked Then
      optPick.Value = vbUnchecked
      If cmbPrt = lblFrom Then
         bGoodMo = GetRun
         sPassedMo = cmbPrt
         GetPart = True
         Exit Function
      End If
   Else
      lblFrom = ""
      optFrom = False
   End If
   If cmbRun.Enabled = True Then cmbRun.Clear
   AdoParameter.Value = sPartNumber
   bSqlRows = clsADOCon.GetQuerySet(RdoMon, AdoQry)
   If bSqlRows Then
      With RdoMon
         cmbPrt = "" & Trim(!PartNum)
         sPassedMo = "" & Trim(!PartNum)
         If cmbRun.Enabled Then
            If optFrom.Value = vbUnchecked Or optNwr.Value = vbUnchecked Then
               cmbRun = Format(!Runno, "####0")
            End If
         End If
         txtQty = Format(!RUNQTY, "####0")
         lblStat = "" & !RUNSTATUS
         lblDsc = "" & Trim(!PADESC)
         lblUom = "" & Trim(!PAUNITS)
         lblSch = "" & Format(!runSched, "mm/dd/yyyy") 'RUNSCHED
         txtPri = 0 + Format(!RUNPRIORITY, "#0")
         lblDiv = "" & !RUNDIVISION
         lblType = !PALEVEL
         GetPart = True
         If cmbRun.Enabled Then
            Do Until .EOF
               AddComboStr cmbRun.hwnd, Format$(!Runno, "####0")
               .MoveNext
            Loop
            ClearResultSet RdoMon
            If cmbRun.ListCount > 0 Then
               cmbRun.ListIndex = cmbRun.ListCount - 1
            End If
         End If
      End With
      If OnLoad Then
      End If
   Else
      MsgBox "Part With MO (NOT CO,CL or CA) Wasn't Found.", vbExclamation, Caption
      cmbRun = "0"
      lblUom = ""
      lblDsc = ""
      lblSch = ""
      txtPri = "0"
      lblDiv = ""
   End If
   Set RdoMon = Nothing
   MouseCursor 0
   bGoodMo = GetRun()
   CloseUnused
   If GetPreferenceValue("AutoSelectLastRun") = "1" Then cmbRun = cmbRun.List(cmbRun.ListCount - 1)
   Exit Function
      
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function GetRun() As Byte
   Dim RdoRns As ADODB.Recordset
   Dim iList As Integer
   sPartNumber = Compress(cmbPrt)
   On Error Resume Next
   On Error GoTo DiaErr1
   
   If Len(Trim(cmbPrt)) = 0 Then Exit Function
   MouseCursor 13
   sSql = "SELECT RUNREF,RUNNO,RUNQTY,RUNSTATUS,RUNSCHED, " _
          & " RUNPRIORITY,RUNDIVISION" _
          & " FROM RunsTable WHERE RUNREF='" & sPartNumber & "' AND RUNNO=" & Val(cmbRun) _
          & " AND RUNSTATUS NOT LIKE 'CA%'"   ' changed back 2/5/20 because Imaginetics wants it - TEL
'          & " AND RUNSTATUS NOT LIKE 'C%'"   ' changed from CA% 5/8/2019 per discussion with Chuck - TEL
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRns, ES_FORWARD)
   If bSqlRows Then
      With RdoRns
         GetRun = True
         cmbRun = Format(!Runno, "####0")
         txtQty = Format(!RUNQTY, ES_QuantityDataFormat)
         lblStat = "" & !RUNSTATUS
         lblSch = "" & Format(!runSched, "mm/dd/yyyy") 'RUNSCHED
         txtPri = Format(!RUNPRIORITY, "#0")
         lblDiv = "" & !RUNDIVISION
         For iList = 1 To 12
            cmdOpt(iList).Enabled = True
         Next
         
         If Trim(lblStat) = "SC" Then
            cmdOpt(9).Enabled = True
         Else
            cmdOpt(9).Enabled = False
         End If
         ' 12/16/2009
         If Trim(lblStat) = "CO" Or Trim(lblStat) = "CL" Then
            For iList = 1 To 12
                If (iList <> 4) Then
                    cmdOpt(iList).Enabled = False
                End If
            Next
            ' just enable MO Comments
'            cmdOpt(4).Enabled = True
'            cmdOpt(4).SetFocus
         End If
         
         
      End With
      ClearResultSet RdoRns
   Else
      For iList = 1 To 12
         cmdOpt(iList).Enabled = False
      Next
      MsgBox "Run Wasn't Found. May Be CO,CL or CA..", vbInformation, Caption
      GetRun = False
   End If
   Set RdoRns = Nothing
   CloseUnused
   MouseCursor 0
   Exit Function
      
DiaErr1:
   sProcName = "getrun"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function



Private Sub CloseUnused()
   'Disables Functions Not formatted
   'z1(10).Enabled = False
   'z1(12).Enabled = False
   
End Sub


Private Sub ReleaseMo()
   Dim bResponse As Byte
   Dim sMsg As String
   sMsg = "Mark This MO As RL (Released To Production)?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      clsADOCon.ADOErrNum = 0
      sSql = "UPDATE RunsTable SET RUNSTATUS='RL',RUNRELEASED=1 " _
             & "WHERE RUNREF='" & Compress(cmbPrt) & "' " _
             & "AND RUNNO=" & Val(cmbRun) & " "
      clsADOCon.ExecuteSql sSql
      If clsADOCon.ADOErrNum = 0 Then
         lblStat = "RL"
         MsgBox "Run Status Updated.", vbInformation, Caption
      Else
         MsgBox "Run Status Could Not Be Updated.", vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   
End Sub

Private Sub UpdateStatus()
   Dim rdo As ADODB.Recordset
   sSql = "SELECT RUNSTATUS" & vbCrLf _
          & " FROM RunsTable WHERE RUNREF='" & Compress(cmbPrt) & "' AND RUNNO=" & Val(cmbRun)
          
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      lblStat = "" & rdo!RUNSTATUS
   End If
   
   If Trim(lblStat) = "SC" Then
      cmdOpt(9).Enabled = True
   Else
      cmdOpt(9).Enabled = False
   End If
   Set rdo = Nothing
End Sub
