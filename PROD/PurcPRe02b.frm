VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form PurcPRe02b 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Purchase Order Items"
   ClientHeight    =   6000
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   7050
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   4306
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAddStat 
      Height          =   350
      Left            =   6360
      Picture         =   "PurcPRe02b.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Status Code"
      Top             =   4320
      Width           =   350
   End
   Begin VB.ComboBox cmbOrigDueDte 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3600
      TabIndex        =   7
      Tag             =   "4"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CheckBox OptDockInsted 
      Caption         =   "Check1"
      Height          =   255
      Left            =   6720
      TabIndex        =   55
      Top             =   2640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox cmbOperation 
      Height          =   315
      Left            =   5640
      TabIndex        =   10
      Tag             =   "1"
      ToolTipText     =   "Select Run Number"
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRe02b.frx":048F
      Style           =   1  'Graphical
      TabIndex        =   54
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optVew 
      Caption         =   "View"
      Height          =   255
      Left            =   0
      TabIndex        =   53
      Top             =   5640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "PurcPRe02b.frx":0C3D
      Height          =   315
      Left            =   4560
      Picture         =   "PurcPRe02b.frx":0F7F
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   2160
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.TextBox txtPrt 
      Height          =   288
      Left            =   1440
      TabIndex        =   1
      Tag             =   "3"
      Top             =   2160
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmdLst 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6000
      Picture         =   "PurcPRe02b.frx":12C1
      Style           =   1  'Graphical
      TabIndex        =   49
      TabStop         =   0   'False
      ToolTipText     =   "Previous Entry"
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdNxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6520
      Picture         =   "PurcPRe02b.frx":13F7
      Style           =   1  'Graphical
      TabIndex        =   48
      TabStop         =   0   'False
      ToolTipText     =   "Next Entry"
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CheckBox optCom 
      Caption         =   "&Repeat Comments"
      Height          =   255
      Left            =   3720
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Carry Part Number, Quantity And Price Forward"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.ComboBox cmbJmp 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   720
      Sorted          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Tag             =   "8"
      ToolTipText     =   "Jump To Operation"
      Top             =   6000
      Width           =   2445
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "PurcPRe02b.frx":152D
      DownPicture     =   "PurcPRe02b.frx":1E9F
      Height          =   350
      Left            =   5880
      Picture         =   "PurcPRe02b.frx":2811
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Standard Comments"
      Top             =   4320
      Width           =   350
   End
   Begin VB.ComboBox cmbAccount 
      Height          =   315
      Left            =   4080
      Sorted          =   -1  'True
      TabIndex        =   14
      Tag             =   "3"
      ToolTipText     =   "Select Account From List"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CheckBox optRcd 
      Caption         =   "Received"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5760
      TabIndex        =   43
      Top             =   1560
      Width           =   1455
   End
   Begin VB.ComboBox txtDue 
      Height          =   315
      ItemData        =   "PurcPRe02b.frx":2E13
      Left            =   1200
      List            =   "PurcPRe02b.frx":2E15
      TabIndex        =   6
      Tag             =   "4"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "PurcPRe02b.frx":2E17
      Height          =   350
      Left            =   6600
      Picture         =   "PurcPRe02b.frx":32F1
      Style           =   1  'Graphical
      TabIndex        =   40
      TabStop         =   0   'False
      ToolTipText     =   "Show Purchase Order Items"
      Top             =   2160
      Width           =   360
   End
   Begin VB.CheckBox optRep 
      Caption         =   "&Repeat Purchased Items"
      Height          =   255
      Left            =   1440
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Carry Part Number, Quantity And Price Forward"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CheckBox optSrv 
      Caption         =   "Service?"
      Height          =   195
      Left            =   240
      TabIndex        =   20
      Top             =   4560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtCmt 
      Height          =   1095
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Tag             =   "9"
      ToolTipText     =   "Comments (2048 Chars Max"
      Top             =   4320
      Width           =   4335
   End
   Begin VB.CommandButton cmdTrm 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6120
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Cancel The Current PO Item"
      Top             =   850
      Width           =   915
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   315
      Left            =   6120
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Add A New PO Item"
      Top             =   520
      Width           =   915
   End
   Begin VB.ComboBox cmbRun 
      Height          =   315
      Left            =   5640
      TabIndex        =   9
      Tag             =   "1"
      ToolTipText     =   "Select Run Number"
      Top             =   3600
      Width           =   855
   End
   Begin VB.ComboBox cmbMon 
      Height          =   315
      ItemData        =   "PurcPRe02b.frx":37CB
      Left            =   1200
      List            =   "PurcPRe02b.frx":37CD
      TabIndex        =   8
      Tag             =   "3"
      ToolTipText     =   "Assign To MO"
      Top             =   3600
      Width           =   3475
   End
   Begin VB.CheckBox optIns 
      Alignment       =   1  'Right Justify
      Caption         =   "Inspection Required?"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4920
      TabIndex        =   17
      ToolTipText     =   "Inspect This Item?"
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox txtLot 
      Height          =   285
      Left            =   2640
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "Lot Quantity"
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox txtPrc 
      Height          =   285
      Left            =   5040
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "Price"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      ItemData        =   "PurcPRe02b.frx":37CF
      Left            =   1440
      List            =   "PurcPRe02b.frx":37D1
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Select From List"
      Top             =   2160
      Width           =   3075
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Enter Quantity"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6120
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   5520
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6000
      FormDesignWidth =   7050
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   1455
      Left            =   360
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Click To Select Or Scroll And Press Enter"
      Top             =   60
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2566
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.Label lblQoh 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4680
      TabIndex        =   58
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Z1 
      Caption         =   "QOH"
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   57
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label lblOrigShip 
      Caption         =   "Orig Due Date"
      Height          =   255
      Left            =   2520
      TabIndex        =   16
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label lblPrc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5040
      TabIndex        =   56
      ToolTipText     =   "Lot cost Price"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblPoQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   240
      TabIndex        =   52
      ToolTipText     =   "= PO Quantity/Purchasing Conversion"
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label lblPoUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   960
      TabIndex        =   51
      ToolTipText     =   "Purchasing Unit Of Measure"
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Dummy 
      Caption         =   "Label1"
      Height          =   135
      Left            =   0
      TabIndex        =   50
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number                                                          "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   46
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label lblPua 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Account"
      Height          =   255
      Left            =   2160
      TabIndex        =   45
      Top             =   5520
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Op No"
      Height          =   252
      Index           =   11
      Left            =   4920
      TabIndex        =   44
      Top             =   3960
      Width           =   612
   End
   Begin VB.Label lblMon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   42
      Top             =   3960
      Width           =   3475
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      Height          =   285
      Left            =   705
      TabIndex        =   41
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   6120
      TabIndex        =   39
      Top             =   2160
      Width           =   372
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type "
      Height          =   255
      Index           =   10
      Left            =   5640
      TabIndex        =   38
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lblTyp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6240
      TabIndex        =   37
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label lblDte 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6480
      TabIndex        =   36
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1080
      TabIndex        =   35
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label lblRel 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5880
      TabIndex        =   34
      Top             =   4800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblPon 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5880
      TabIndex        =   33
      Top             =   4560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   32
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   252
      Index           =   8
      Left            =   4920
      TabIndex        =   31
      Top             =   3600
      Width           =   612
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   30
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   29
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lot Quantity If Not A Unit Price"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   28
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   27
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price                           "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   5040
      TabIndex        =   26
      Top             =   1920
      Width           =   1572
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity           "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   25
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   24
      Top             =   1560
      Width           =   615
   End
End
Attribute VB_Name = "PurcPRe02b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'Corrected a problem with adding items for PO's with none
'5/19/04 Allow Reprice of Received Items
'7/8/04 Smoothed grid scrolling
'9/29/04 Reformatted values to prevent overflow
'10/7/04 ES_PurchasedDataFormat
'12/29/04 Added provision for Null date and shut off Last/Next buttons
'12/30/04 Added purchasing conversion to form
'3/7/05 Added PIENTERED to AddPOItem. Trap for Null Part
'5/9/05 Removed ValidateEdit (unnecessary Event logging)
'11/1/05 Removed RdoQry1 (attempt to resolve problems at PROPLA)
'11/9/05 Auto add a new PO Item
'11/10/05 Added Suppress Messages
'1/24/05 Added txtPrt for Part Count > 5000 (INTCOA)
'3/22/07 Operation checking added
Option Explicit
Dim AdoQry2 As ADODB.Command
Dim ADOParameter2 As ADODB.Parameter
Dim RdoPoi As ADODB.Recordset

Dim bAllowPrice As Byte 'Allow repricing after receipt
Dim bCantCancel As Byte
Dim bGoodItem As Byte
Dim bGoodMo As Byte
Dim bItemCan As Byte
Dim bNewItem As Byte
Dim bNewPo As Byte
Dim bOnLoad As Byte
Dim bPoRows As Byte
Dim bSuppress As Byte
Dim bView As Byte
Dim UseLastInvoicedPrice As Boolean
Dim bFieldChanged, bFromDropDown As Byte
   
   
Dim AdoQry1 As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter


'Dim FusionToolTip As New clsToolTip


Dim iIndex As Integer
Dim iMaxItem As Integer
Dim iNextItem As Integer
Dim iRows As Integer
Dim iTotalItems As Integer

Dim cConversion As Currency

Dim sPartNumber As String
Dim sOldMo As String
Dim PartNumOnComboEntry As String

Dim vItems(300, 2) As Variant

' purchase conversion info for current item (only relevant for UseLastInvoicedPrice = true
' and when quantity <> 0 and <> 1
Dim iPurchaseConversionQty


Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbJmp_Click()
   'On Error Resume Next
   iIndex = cmbJmp.ListIndex
   If iIndex < 0 Then iIndex = 0
   If iIndex > iMaxItem Then iIndex = iMaxItem
   bGoodItem = GetThisItem(False)
   
End Sub


Private Sub cmbJmp_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then KeyCode = 0
   
End Sub



Private Sub cmbMon_Click()
   
   If Compress(sOldMo) <> Compress(cmbMon) Then
      If cmbMon.ListCount > 0 Then _
         If Len(Trim(cmbMon)) Then GetRuns
   End If
   
End Sub

Private Sub cmbMon_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub cmbMon_LostFocus()
   cmbMon = CheckLen(cmbMon, 30)
     
   cmbMon_Click
   
   ' MM Ticket#43782
   ' Only if the Run number is greater than zero
   ' Commented the SetFocus because it will create a loop is then ISValidMORun fails.
   If (Val(cmbRun) > 0) Then
      If Not IsValidMORun(cmbMon, Val(cmbRun), True, True) Then
          'cmbMon.SetFocus
          Exit Sub
      End If
   End If
   
   FindMoPart
   If Trim(cmbMon) = "NONE" Then cmbMon = ""
   If Val(txtQty) > 0 Then
      'cmbJmp.Enabled = True
      cmdNew.Enabled = True
      If bCantCancel = 0 Then cmdTrm.Enabled = True
   End If
   If Trim(cmbMon) <> "" Then
      If Compress(sOldMo) <> Compress(cmbMon) Or Trim(cmbMon) <> "NONE" Then GetRuns
   Else
      lblMon = ""
      If cmbRun.ListCount > 0 Then cmbRun.Clear
   End If
   If lblMon.ForeColor <> ES_RED Then
      If bPoRows Then
         'On Error Resume Next
         RdoPoi!PIRUNPART = Compress(cmbMon)
         RdoPoi.Update
      End If
   End If
   If Len(Trim(cmbMon)) > 0 Then
      'txtOpn.Enabled = True
      cmbOperation.Enabled = True
      z1(11).Enabled = True
   Else
      'txtOpn.Enabled = False
      cmbOperation.Enabled = False
      z1(11).Enabled = False
   End If
   sOldMo = cmbMon
   
End Sub

Private Sub cmbMon_Validate(Cancel As Boolean)
    'If Not IsValidMORun(cmbMon, Val(cmbRun), True, True) Then Cancel = True
End Sub

Private Sub cmbOperation_LostFocus()
   Dim iOPNumber As Integer
   'txtOpn = CheckLen(txtOpn, 3)
   'txtOpn = Format(Abs(Val(txtOpn)), "000")
   If bPoRows Then
      iOPNumber = GetPoOp()
      'On Error Resume Next
      'RdoPoi.Edit
      If Trim(cmbMon) <> "" Then
         '            RdoPoi!PIRUNOPNO = Val(txtOpn)
         RdoPoi!PIRUNOPNO = Val(cmbOperation)
      Else
         RdoPoi!PIRUNOPNO = 0
      End If
      RdoPoi.Update
   End If
   If Val(cmbOperation) = 0 Then cmbOperation = ""
   cmbOperation = Format(cmbOperation, "000")

End Sub



Private Sub cmbOrigDueDte_DropDown() 'BBS Added on 03/09/2010 for Ticket #11364
  ShowCalendarEx Me
End Sub

Private Sub cmbOrigDueDte_KeyUp(KeyCode As Integer, Shift As Integer) 'BBS Added on 03/09/2010 for Ticket #11364
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
End Sub

Private Sub cmbOrigDueDte_LostFocus()  'BBS Added on 03/09/2010 for Ticket #11364
   cmbOrigDueDte = CheckDateEx(cmbOrigDueDte)
'   If Val(txtQty) > 0 Then
'      cmdNew.Enabled = True
'      If bCantCancel = 0 Then cmdTrm.Enabled = True
'   End If
   If bPoRows Then
      RdoPoi!PIPORIGDATE = cmbOrigDueDte
      RdoPoi.Update
   End If
   UpdateJump
End Sub


Private Sub cmbPrt_Click()
   Dim sDescription As String
   cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
   txtPrt = cmbPrt
   LoadExtendedPartDesc
   sOldMo = cmbMon
End Sub

Private Sub cmbPrt_GotFocus()
   PartNumOnComboEntry = Compress(cmbPrt)
End Sub

Private Sub cmbPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub cmbPrt_LostFocus()
   Dim sDescription As String
   Dim cPartCost As Currency
   Dim bPartVal As Boolean
   'On Error Resume Next
   If bView = 1 Then Exit Sub
   
   cmbPrt = CheckLen(cmbPrt, 30)
   cmbPrt = GetCurrentPart(cmbPrt, lblDsc, True)
   
   Dim partNumberChange As Boolean
'   If Compress(cmbPrt) <> PartNumOnComboEntry Then
'      partNumberChange = True
'      PartNumOnComboEntry = Compress(cmbPrt)
'   End If
   
   bPartVal = ValidPartNumber(cmbPrt.Text)
   If (Not bPartVal) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If
      
   'determine if part has changed
   Dim curPartInDb As String
   On Error Resume Next
   curPartInDb = Compress(RdoPoi!PIPART)
   Err.Clear
   On Error GoTo 0
   
   If Compress(cmbPrt) <> curPartInDb Then
      partNumberChange = True
   End If
   
   GetThisPart partNumberChange
   
   If Val(txtQty) > 0 Then
      If Len(cmbPrt) > 0 Then
         If bNewItem Then
            cPartCost = GetPartCost(cmbPrt, ES_STANDARDCOST)
            txtPrc = Format(cPartCost, ES_PurchasedDataFormat)
            ' MM lblPrc = txtPrc
         End If
         'cmbJmp.Enabled = True
         cmdNew.Enabled = True
         If bCantCancel = 0 Then cmdTrm.Enabled = True
         bNewItem = False
      End If
   End If
   If bPoRows Then
      If lblDsc.ForeColor <> ES_RED Then
         'On Error Resume Next
         RdoPoi!PIPART = Compress(cmbPrt)
         RdoPoi.Update
         Grd.Col = 2
         Grd.Text = cmbPrt
         Grd.Col = 0
      End If
   End If
   If Trim(cmbPrt) = "NONE" Then
      MsgBox "Please Select A Valid Part Number.", _
         vbInformation, Caption
   Else
      UpdateJump
   End If
   txtPrt = cmbPrt
   LoadExtendedPartDesc
End Sub

Private Sub cmbRun_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub cmbRun_LostFocus()
   cmbRun = CheckLen(cmbRun, 5)
   cmbRun = Format(Abs(Val(cmbRun)), "####0")
   If Val(cmbRun) > 32767 Then cmbRun = "32767"
   
   If Not IsValidMORun(cmbMon, Val(cmbRun), True, True) Then
       'MM comented as this will create a loop.
       'MM and crash the module
       'cmbRun.SetFocus
       Exit Sub
   End If

   If Val(txtQty) > 0 Then
      'cmbJmp.Enabled = True
      cmdNew.Enabled = True
      If bCantCancel = 0 Then cmdTrm.Enabled = True
   End If
   If lblMon.ForeColor <> ES_RED And Val(cmbRun) > 0 Then
      If bPoRows Then
         'On Error Resume Next
         RdoPoi!PIRUNNO = Val(cmbRun)
         RdoPoi.Update
      End If
   End If
   If Val(cmbRun) > 0 Then
      'txtOpn.Enabled = True
      cmbOperation.Enabled = True
      z1(11).Enabled = True
   Else
      'txtOpn.Enabled = False
      cmbOperation.Enabled = False
      z1(11).Enabled = False
   End If
   FillOps
End Sub

Private Sub cmbRun_Validate(Cancel As Boolean)
'    If Not IsValidMORun(cmbMon, Val(cmbRun), True, True) Then Cancel = True
End Sub


Private Sub cmdAddStat_Click()
    StatusCode.lblSCTypeRef = "PO"
    StatusCode.txtSCTRef = lblPon
    StatusCode.LableRef1 = "Item"
    StatusCode.lblSCTRef1 = lblItm
    StatusCode.lblSCTRef2 = lblRev
    StatusCode.lblStatType = "POI"
    StatusCode.lblSysCommIndex = 1 ' The index in the Sys Comment "MO Comments"
    StatusCode.txtCurUser = cUR.CurrentUser
    
    StatusCode.Show

End Sub

Private Sub cmdCan_Click()
   If txtPrt.Visible Then cmbPrt = txtPrt
   If Trim(cmbPrt) = "" Then bGoodItem = 0
   PurcPRe02a.optItm.Value = vbUnchecked
   If bGoodItem = 1 Then UpdateItem
   Unload Me
   
End Sub



Private Sub cmdComments_Click()
   If cmdComments Then
      'See List For Index
      If txtCmt.Enabled Then txtCmt.SetFocus
      SysComments.lblListIndex = 1
      SysComments.Show
      cmdComments = False
   End If
   
End Sub


Private Sub cmdFnd_Click()
   If txtPrt.Visible Then
      cmbPrt = txtPrt
      ViewParts.lblControl = "TXTPRT"
   Else
      ViewParts.lblControl = "CMBPRT"
   End If
   ViewParts.txtPrt = cmbPrt
   optVew.Value = vbChecked
   ViewParts.Show
   GetThisPart True
   bView = 0
   
End Sub

Private Sub cmdFnd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bView = 1
   
End Sub



Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4306
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub cmdLst_Click()
   Dim A As Integer
   iIndex = iIndex - 1
   If iIndex < 0 Then iIndex = 0
   'On Error Resume Next
   If bPoRows Then
      RdoPoi!PIPQTY = Val(txtQty)
      RdoPoi!PIESTUNIT = Val(txtPrc)
      RdoPoi!PICOMT = "" & Trim(txtCmt)
      RdoPoi.Update
   End If
   Grd.row = iIndex + 1
   bGoodItem = GetThisItem(False)
   If Grd.row < 6 Then
      Grd.TopRow = 1
   Else
      A = Grd.row Mod 4
      If A = 2 Then Grd.TopRow = Grd.row - 4
   End If
   
End Sub

Private Sub cmdNew_Click()
   If Val(lblItm) = 0 Then
      AddPOItem
   Else
      If Trim(cmbPrt) = "" Or Trim(cmbPrt) = "NONE" Or Val(txtQty) = 0 Then
         MsgBox "Invalid Part Number Or Quantity.", vbExclamation, Caption
         'On Error Resume Next
         cmbPrt.SetFocus
      Else
         UpdateItem
         AddPOItem
      End If
   End If
End Sub

Private Sub cmdNxt_Click()
   Dim A As Integer
   iIndex = iIndex + 1
   If iIndex > iMaxItem Then iIndex = iMaxItem
   'On Error Resume Next
   If bPoRows Then
      RdoPoi!PIPQTY = Val(txtQty)
      RdoPoi!PIESTUNIT = Val(txtPrc)
      RdoPoi.Update
   End If
   Grd.row = iIndex + 1
   bGoodItem = GetThisItem(False)
   If Grd.row < 6 Then
      Grd.TopRow = 1
   Else
      A = Grd.row Mod 4
      If A = 2 Then Grd.TopRow = Grd.row - 1
   End If
   
End Sub

Private Sub cmdTrm_Click()

   Dim bCancelType As Boolean
   If iTotalItems <= 1 Then
      MsgBox "You can't delete the last item on a PO.  " & vbCrLf _
         & "Either cancel the PO or revise the item.", vbInformation
      Exit Sub
   End If
   
   Dim bResponse As Byte
   Dim sMsg As String
   If optRcd.Value = vbChecked And optRcd.Caption <> "Cancelled" Then
      MsgBox "You May Not Cancel A Received Or Invoiced Item.", _
         vbInformation, Caption
      Exit Sub
   End If
   
    If optIns.Value = vbChecked Then
      If optDockInsted.Value = vbChecked Then
        MsgBox "You May Not Cancel A Dock Inspected Item.", _
            vbInformation, Caption
        Exit Sub
      End If
    End If
   
   
    If optRcd.Caption <> "Cancelled" Then
        sMsg = "You Are About To Cancel This Item, " & vbCr _
            & "Are You Sure That You Wish To Continue?"
      
      'MM Remove the PIPQTY as we can uncancel an item
'      sSql = "UPDATE PoitTable SET PITYPE=16,PIPQTY=0," _
'             & "PIESTUNIT=0 WHERE PINUMBER=" & lblPon & " " _
'             & "AND PIITEM=" & lblItm & " " _
'             & "AND PIREV='" & lblRev & "'"
      sSql = "UPDATE PoitTable SET PITYPE=16 " _
             & " WHERE PINUMBER=" & lblPon & " " _
             & "AND PIITEM=" & lblItm & " " _
             & "AND PIREV='" & lblRev & "' "
      bCancelType = True
    Else
        sMsg = "You Are About To UnCancel This Item, " & vbCr _
            & "Are You Sure That You Wish To Continue?"
      
      sSql = "UPDATE PoitTable SET PITYPE=14 " _
             & "WHERE PINUMBER=" & lblPon & " " _
             & "AND PIITEM=" & lblItm & " " _
             & "AND PIREV='" & lblRev & "' "
            
      bCancelType = False
    End If
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   'On Error Resume Next
   If bResponse = vbYes Then
      clsADOCon.ExecuteSql sSql
      If clsADOCon.RowsAffected > 0 Then
         If bCancelType = True Then
            bItemCan = 1
            MsgBox "Item " & lblItm & lblRev & " Cancelled.", vbInformation, Caption
         Else
            MsgBox "Item " & lblItm & lblRev & " UnCancelled.", vbInformation, Caption
         End If
      Else
         If bCancelType = True Then
            MsgBox "Couldn't Cancel Item.", vbInformation, Caption
         Else
            MsgBox "Couldn't UnCancel Item.", vbInformation, Caption
        End If
      End If
      GetItems
   Else
      CancelTrans
   End If
   
End Sub

Private Sub cmdVew_Click()
   Dim iList As Integer
   Dim iCol As Integer
   Dim GrdRows As Integer
   Dim RdoVew As ADODB.Recordset
   UpdateItem
   'On Error Resume Next
   'BBS added on 03/09/2010 for Ticket #11364
   sSql = "SELECT PINUMBER,PIITEM,PIREV,PITYPE,PIPDATE, PIPORIGDATE, PIADATE,PIPQTY,PIPART,PARTREF,PARTNUM," _
          & "PIESTUNIT,PILOT FROM PoitTable,PartTable WHERE (PIPART=PARTREF) AND " _
          & "PINUMBER=" & Val(lblPon) & " " _
          & "ORDER BY PIITEM"
   Set RdoVew = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   MouseCursor 13
   GrdRows = 10
   With PurcPRe01c.Grd
      .Rows = GrdRows
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      If Screen.Width > 9999 Then
         .ColWidth(0) = 350 * 1.25
         .ColWidth(1) = 1450 * 1.25
         .ColWidth(2) = 600 * 1.25
         .ColWidth(3) = 600 * 1.25
         .ColWidth(4) = 650 * 1.25
         .ColWidth(5) = 650 * 1.25
         .ColWidth(6) = 650 * 1.25
      Else
         .ColWidth(0) = 350
         .ColWidth(1) = 1450
         .ColWidth(2) = 600
         .ColWidth(3) = 600
         .ColWidth(4) = 650
         .ColWidth(5) = 650
         .ColWidth(6) = 650
      End If
   End With
   If Not RdoVew.BOF Then
      With RdoVew
         Do Until .EOF
            iList = iList + 1
            GrdRows = GrdRows + 1
            PurcPRe01c.Grd.Rows = GrdRows
            PurcPRe01c.Grd.row = GrdRows - 10
            
            PurcPRe01c.Grd.Col = 0
            PurcPRe01c.Grd = "" & str(!PIITEM) & Trim(!PIREV)
            
            PurcPRe01c.Grd.Col = 1
            PurcPRe01c.Grd = "" & Trim(!PartNum)
            
            PurcPRe01c.Grd.Col = 2
            PurcPRe01c.Grd = Format(!PIPQTY, ES_QuantityDataFormat)
            
            PurcPRe01c.Grd.Col = 3
            If !PILOT > 0 Then
               PurcPRe01c.Grd = Format(!PIESTUNIT / !PILOT, ES_QuantityDataFormat)
            Else
               PurcPRe01c.Grd = Format(!PIESTUNIT, ES_QuantityDataFormat)
            End If
            
            PurcPRe01c.Grd.Col = 4
            PurcPRe01c.Grd = Format(!PIPDATE, "mm/dd/yy")
            
            PurcPRe01c.Grd.Col = 5
            PurcPRe01c.Grd = Format(!PIADATE, "mm/dd/yy")
            
            PurcPRe01c.Grd.Col = 6
            If !PITYPE = 16 Then _
                         Set PurcPRe01c.Grd.CellPicture = PurcPRe01c.Chkyes.Picture
            .MoveNext
         Loop
         ClearResultSet RdoVew
      End With
      If iList > 9 Then PurcPRe01c.Grd.Rows = iList + 1
   End If
   MouseCursor 0
   Set RdoVew = Nothing
   PurcPRe01c.Show
   On Error GoTo 0
   
End Sub

Private Sub cmdVew_LostFocus()
   '
End Sub


Private Sub Form_Activate()
   Dim b As Byte
   Dim sRel As String
   
   cConversion = 1
   Caption = Caption & " - PO " & PurcPRe02a.cmbPon
   If Val(PurcPRe02a.txtRel) > 0 Then sRel = "" _
          Else sRel = PurcPRe02a.txtRel
   If bOnLoad Then
   
      'determine whether to use standard cost or most recently invoiced cost
      Dim rdo As ADODB.Recordset
      sSql = "select COPODEFAULTTOLASTPRICE from ComnTable where COREF=1 and COPODEFAULTTOLASTPRICE = 1"
      If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
         UseLastInvoicedPrice = True
      Else
         UseLastInvoicedPrice = False
      End If
      Set rdo = Nothing
      
      ES_PurchasedDataFormat = GetPODataFormat()
      b = CheckPoAccounts()
      FillRuns Me, "NOT LIKE 'C%'", cmbMon
      FillCombo
      FillAccounts
      If b = 1 Then
         lblPua.Visible = True
         cmbAccount.Visible = True
      End If
      GetItems
      GetRuns   'BBS Added this on 09/14/2010 for Ticket #18545
      bGoodItem = GetThisItem(False)
      bOnLoad = 0
   Else
      bGoodItem = GetThisItem(False)
   End If
   MouseCursor 0
    Call HookToolTips
    MVBBubble.MaxTipWidth = 600
    If (txtQty.Enabled = True) Then txtQty.SetFocus
End Sub

'Private Sub Form_Deactivate()
'   Unload Me
'
'End Sub

Private Sub Form_Load()
   bOnLoad = 1
   SetFormSize Me
   Move PurcPRe02a.Left + 300, PurcPRe02a.Top + 1000
   FormatControls
   MouseCursor 13
   optRep.Value = GetSetting("Esi2000", "EsiProd", "PoRepeat", optRep.Value)
   optCom.Value = GetSetting("Esi2000", "EsiProd", "ComRepeat", optCom.Value)
   lblPon = PurcPRe02a.cmbPon
   'lblDte = PurcPRe02a.txtSdt
   lblDte = Format(PurcPRe02a.txtSdt, "mm/dd/yy")
   If Trim(lblDte) = "" Then lblDte = Format(ES_SYSDATE, "mm/dd/yy")
   If Val(Right(lblDte, 2) & Left(lblDte, 2) & Mid(lblDte, 4, 2)) < Val(Format(ES_SYSDATE, "yymmdd")) Then lblDte = Format(ES_SYSDATE, "mm/dd/yy")
'   If lblDte < Format(ES_SYSDATE, "mm/dd/yy") Then _
'                      lblDte = Format(ES_SYSDATE, "mm/dd/yy")
   If PurcPRe02a.optNew.Value = vbChecked Then bNewPo = True Else bNewPo = False
   lblItm = ""
   optSrv.Value = PurcPRe02a.optSrv.Value
   '4/26/99 - Allow all PO items to allocate
   'If optSrv = vbUnchecked Then
   'z1(7).Enabled = False
   'z1(8).Enabled = False
   'cmbMon.Enabled = False
   'cmbRun.Enabled = False
   'End If
   'ES_PARTCOUNT = 5001
   If ES_PARTCOUNT > 5000 Then
     ' cmbPrt.Visible = False
     ' txtPrt.Visible = True
      txtPrt.TabIndex = 99
      'cmbPrt.TabIndex = 99
      'cmdFnd.Visible = True
   End If
      
   sSql = "SELECT PAEXTDESC FROM PartTable WHERE PARTREF = ? "
   Set AdoQry1 = New ADODB.Command
   AdoQry1.CommandText = sSql
   
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adChar
   AdoParameter1.SIZE = 30
   
   AdoQry1.Parameters.Append AdoParameter1
   
   
   
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PAONDOCK,PARUN,RUNREF," _
          & "RUNSTATUS,RUNNO FROM PartTable,RunsTable WHERE PARTREF= ? " _
          & "AND PARTREF=RUNREF AND RUNSTATUS NOT LIKE 'C%' "
   Set AdoQry2 = New ADODB.Command
   AdoQry2.CommandText = sSql
   
   Set ADOParameter2 = New ADODB.Parameter
   ADOParameter2.Type = adChar
   ADOParameter2.SIZE = 30
   
   AdoQry2.Parameters.Append ADOParameter2
   
   
   With Grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .ColAlignment(2) = 0
      .row = 0
      .Col = 0
      .Text = "Item"
      .ColWidth(0) = 650
      .Col = 1
      .Text = "Rev"
      .ColWidth(1) = 400
      .Col = 2
      .Text = "Part Number"
      .ColWidth(2) = 3350
      .Col = 3
      .Text = "Quantity"
      .ColWidth(3) = 950
      .Col = 0
   End With
'   bOnLoad = 1
'    With FusionToolTip
'        Call .Create(Me)
'        .MaxTipWidth = 240
'        .DelayTime(ttDelayShow) = 20000
'        Call .AddTool(txtPrt)
'    End With
End Sub


'Final check 2/21/02

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim bResponse As Byte
   Dim sMsg As String
   
   bResponse = FinalItemCheck()
   MouseCursor 0
   If bResponse = 1 Then
      sMsg = "System Check Has Determined That One Or More " & vbCr _
             & "Items Has An Invalid Part Number.  If You " & vbCr _
             & "Continue, The Item Will be Lost." & vbCr _
             & "Do You Wish To Close The Form Anyway?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbNo Then
         Cancel = True
         Exit Sub
      End If
   End If
   
   'if new po item and bad part then dump it
   SaveSetting "Esi2000", "EsiProd", "PoRepeat", Trim(str(optRep.Value))
   SaveSetting "Esi2000", "EsiProd", "ComRepeat", Trim(str(optCom.Value))
   'On Error Resume Next
   MouseCursor 13
   If bNewItem = 1 And bFoundPart = 0 Then
      sSql = "DELETE FROM PoitTable WHERE (PINUMBER=" & Val(lblPon) _
             & " AND PIITEM=" & lblItm & " AND PIREV='" & Trim(lblRev) & "')"
      clsADOCon.ExecuteSql sSql
   End If
   sSql = "DELETE FROM PoitTable WHERE PINUMBER=" & Val(lblPon) _
          & " AND (PIPART='NONE' OR PIPART='') AND PIAQTY=0"
   clsADOCon.ExecuteSql sSql
   
   sSql = "UPDATE PoitTable SET PIVENDOR='" & Compress(PurcPRe02a.cmbVnd) _
          & "' WHERE PINUMBER=" & lblPon & " "
   clsADOCon.ExecuteSql sSql
   PurcPRe02a.RefreshPoCursor
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   'On Error Resume Next
   Set AdoParameter1 = Nothing
   Set AdoQry1 = Nothing
   Set ADOParameter2 = Nothing
   Set AdoQry2 = Nothing
   
   MouseCursor 0
   UnhookToolTips
   Set RdoPoi = Nothing
   Set PurcPRe02b = Nothing
End Sub






'Add PON to Error trap (trapping and testing)

Private Sub GetItems()
   Dim RdoItm As ADODB.Recordset
   Dim iList As Integer
   Dim sRev As String * 2
   Dim sJpart As String * 12
   
   cmbJmp.Clear
   Erase vItems
   bGoodItem = 0
   bGoodMo = False
   iList = -1       '@@@
   GetLastItem
   iRows = 0
   Grd.Rows = 2
   Grd.row = 1
   On Error GoTo DiaErr1
' MM To populate the grid with the items which are canceled
'   sSql = "SELECT PINUMBER,PIITEM,PIREV,PITYPE,PIPART,PIPDATE," _
'          & "PIPQTY,PIACCOUNT FROM PoitTable WHERE (PITYPE<>16 AND PINUMBER=" _
'          & lblPon & ") ORDER BY PIITEM,PIREV "
   
   sSql = "SELECT PINUMBER,PIITEM,PIREV,PITYPE,PIPART,PIPDATE,PIPORIGDATE," _
          & "PIPQTY,PIACCOUNT FROM PoitTable WHERE (PINUMBER=" _
          & lblPon & ") ORDER BY PIITEM,PIREV "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoItm, ES_FORWARD)
   If bSqlRows Then
      With RdoItm
         Do Until .EOF
            iList = iList + 1
            vItems(iList, 0) = Format(!PIITEM, "###0")
            vItems(iList, 1) = "" & Trim(!PIREV)
            If Trim(!PIREV) = "" Then sRev = " " Else sRev = vItems(iList, 1)
            sJpart = Trim(!PIPART)
            If Val(vItems(iList, 0)) < 99 Then
               cmbJmp.AddItem Chr$(32) & Trim(str(!PIITEM)) & sRev & " " & sJpart & "... " & Format(!PIPDATE, "mm/dd/yy")
            Else
               cmbJmp.AddItem Trim(str(!PIITEM)) & sRev & " " & sJpart & "... " & Format(!PIPDATE, "mm/dd/yy")
            End If
            iRows = iRows + 1
            If iRows > 1 Then Grd.Rows = Grd.Rows + 1
            Grd.row = iRows
            Grd.Col = 0
            Grd.Text = Format(!PIITEM, "##0")
            Grd.Col = 1
            Grd.Text = "" & Trim(!PIREV)
            Grd.Col = 2
            If Not IsNull(!PIPART) Then
               Grd.Text = GetCurrentPart(!PIPART, Dummy)
            Else
               Grd.Text = ""
            End If
            Grd.Col = 3
            Grd.Text = Format(!PIPQTY, ES_QuantityDataFormat)
            .MoveNext
         Loop
         ClearResultSet RdoItm
      End With
      iTotalItems = iList + 1
      iIndex = 0
      iMaxItem = iList
      Grd.Col = 0
      Grd.row = 1
      bGoodItem = GetThisItem(False)
      sProcName = "getitems"
   Else
      iTotalItems = 0
      If Not bNewPo Then
         'Insert added 1/17/03 corrects problems
         MsgBox "No Open PO Items Were Found.", vbInformation, Caption
         AddPOItem True
         bItemCan = 0
      
      'if no items, add 1
      Else
         bSuppress = 1
         iMaxItem = -1     '@@@
         AddPOItem
      End If
   End If
   'On Error Resume Next
   Set RdoItm = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Function GetThisItem(NewItem As Boolean) As Byte
   '= true if new item was just added
   
   Dim A As Integer
   Dim iItem As Integer
   Dim sRev As String * 2
   Dim sJpart As String * 12
   Dim sDescription As String
   If bOnLoad = 0 Then UpdateItem
   
   On Error GoTo DiaErr1
   sProcName = "getthisit "
   sRev = "" & Trim(vItems(iIndex, 1))
   iItem = Val(vItems(iIndex, 0))
   'BBS Added PIPORIGDATE on 03/09/2010 for Ticket #11364
   sSql = "SELECT PINUMBER,PIITEM,PIREV,PIPART,PIPQTY,PIAMT," _
          & "PIESTUNIT,PILOT,PIONDOCK,PIONDOCKINSPECTED,PIRUNPART,PIRUNNO,PIRUNOPNO," _
          & "PIPDATE, PIPORIGDATE, PITYPE,PICOMT,PIACCOUNT" & vbCrLf _
          & "FROM PoitTable" & vbCrLf _
          & "WHERE " _
          & "(PINUMBER=" & Val(lblPon) & " AND PIITEM=" & iItem & " " _
          & "AND PIREV='" & sRev & "')"
   bPoRows = clsADOCon.GetDataSet(sSql, RdoPoi, ES_KEYSET)
   cmdNxt.Enabled = False
   cmdLst.Enabled = False
   If bPoRows Then
      MouseCursor 13
      With RdoPoi
         'On Error Resume Next
       
         cmbPrt = "" & Trim(!PIPART)
         cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
         txtPrt = cmbPrt
         LoadExtendedPartDesc
         bCantCancel = 0
         If lblDsc.ForeColor = ES_RED Then lblDsc = ""
         lblItm = Format(!PIITEM, "##0")
         lblRev = "" & Trim(!PIREV)
         txtQty = Format(!PIPQTY, ES_QuantityDataFormat)
         txtPrc = Format(!PIESTUNIT, ES_PurchasedDataFormat)
         txtLot = Format(!PILOT, ES_QuantityDataFormat)
         If (Val(txtLot) > 0) Then
            lblPrc = Format((Val(txtPrc) * Val(txtLot)), ES_PurchasedDataFormat)
         Else
            lblPrc = ""
         End If
         optIns.Value = !PIONDOCK
         optDockInsted.Value = !PIONDOCKINSPECTED
         If Not IsNull(!PIPDATE) Then
            txtDue = Format(!PIPDATE, "mm/dd/yyyy")
         Else
            txtDue = Format(Now, "mm/dd/yyyy")
         End If
         If Not IsNull(!PIPORIGDATE) Then
            cmbOrigDueDte = Format(!PIPORIGDATE, "mm/dd/yyyy")
        Else
            cmbOrigDueDte = ""
        End If
         Grd.Col = 1
         Grd.Text = lblRev
         Grd.Col = 2
         Grd.Text = cmbPrt
         Grd.Col = 3
         Grd.Text = txtQty
         lblMon = ""
         cmbMon = "" & Trim(!PIRUNPART)
         sOldMo = cmbMon
         
      'GetRuns   'BBS Added this on 09/14/2010 for Ticket #18545
      
         
         
         '            If !PIRUNOPNO > 0 Then
         '                txtOpn = Format(!PIRUNOPNO, "000")
         '            Else
         '                txtOpn = ""
         '            End If
'         MsgBox ("GetThisItem MO:" & cmbMon)
'         MsgBox ("GetThisItem RUN:" & !PIRUNNO)
         
         If !PIRUNNO > 0 Then
            cmbRun = Format(!PIRUNNO, "###0")
            'txtOpn.Enabled = True
            z1(11).Enabled = True
         Else
            cmbRun = ""
            'txtOpn.Enabled = False
            z1(11).Enabled = False
         End If
         Dim strRunOP As String
         strRunOP = Format(!PIRUNOPNO, "000")
         
         FillOps
         
         If !PIRUNOPNO > 0 Then
            'txtOpn = Format(!PIRUNOPNO, "000")
            cmbOperation.Text = strRunOP 'Format(!PIRUNOPNO, "000")
         Else
            'txtOpn = ""
            cmbOperation.Text = ""
         End If
         
         txtCmt = "" & Trim(!PICOMT)
         cmbAccount = "" & Trim(!PIACCOUNT)
         If Trim(cmbMon) <> "" Then FindMoPart
         If Trim(cmbMon) = "NONE" Then cmbMon = ""
         cmdTrm.Enabled = True
         GetThisItem = 1
         sPartNumber = cmbPrt
         'GetAccount
         GetThisPart NewItem
         sRev = "" & Trim(!PIREV)
         If sRev = "" Then sRev = " "
        cmdTrm.Caption = "&Cancel"
         If !PITYPE = 15 Or !PITYPE = 17 Or !PITYPE = 16 Then
            bCantCancel = 1
            cmdTrm.Enabled = False
            If !PITYPE = 17 Then
                optRcd.Caption = "Invoiced"
                cmdTrm.Enabled = False
            ElseIf (!PITYPE = 16) Then
                optRcd.Caption = "Cancelled"
                cmdTrm.Caption = "&Uncancel"
                cmdTrm.Enabled = True
                cmdFnd.Enabled = False
            Else
                optRcd.Caption = "Received"
                cmdTrm.Enabled = False
            End If
            
'         MsgBox ("ResetBoxes MO:" & cmbMon)
'         MsgBox ("ResetBoxes RUN:" & !PIRUNNO)
            ResetBoxes False
         Else
            If !PITYPE = 14 Then
               bCantCancel = 0
               cmdTrm.Enabled = True
               optRcd.Caption = "Received"
               ResetBoxes True
               cmdTrm.Enabled = True
               cmdFnd.Enabled = True
            End If
         End If
         '5/19/04
         If bAllowPrice = 1 And !PITYPE = 15 Then txtPrc.Enabled = True
         If Right(txtCmt, 1) = vbLf And Right(txtCmt, 1) = vbCr Then
            If Len(txtCmt) > 1 Then
               A = Len(txtCmt)
               txtCmt = (Left$(txtCmt, A - 1))
            Else
               txtCmt = ""
            End If
         End If
      End With
      bPoRows = 1
      Grd.Col = 0
      Grd.SetFocus
      MouseCursor 0
   Else
      MouseCursor 0
      bCantCancel = 1
      GetThisItem = 0
      cmdTrm.Enabled = False
      MsgBox "No Such Item.", vbInformation, Caption
   End If
   If Grd.Rows > 2 Then cmdNxt.Enabled = True
   cmdLst.Enabled = True
   Exit Function
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function





Private Sub Grd_Click()
   iIndex = Grd.row - 1
   GetThisItem False
   
End Sub

Private Sub grd_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      iIndex = Grd.row - 1
      GetThisItem False
   End If
   
End Sub


Private Sub grd_RowColChange()
'   If bOnLoad = 0 Then
'      iIndex = Grd.row - 1
'      GetThisItem False
'   End If
End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** Part" Then
      lblDsc.ForeColor = ES_RED
      bFoundPart = 0
   Else
      lblDsc.ForeColor = Es_TextForeColor
      bFoundPart = 1
   End If
   
End Sub



Private Sub lblItm_Change()
   If Val(lblItm) > 0 Then
      If bCantCancel = 0 Then cmdTrm.Enabled = True
   Else
      cmdTrm.Enabled = False
   End If
   
End Sub

Private Sub lblMon_Change()
   If Left(lblMon, 8) = "*** Part" Then
      lblMon.ForeColor = ES_RED
   Else
      lblMon.ForeColor = Es_TextForeColor
   End If
   
End Sub


Private Sub optIns_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optIns_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub optIns_LostFocus()
   If Val(txtQty) > 0 Then
      'cmbJmp.Enabled = True
      cmdNew.Enabled = True
      If bCantCancel = 0 Then cmdTrm.Enabled = True
   End If
   If bPoRows Then
      'On Error Resume Next
      RdoPoi!PIONDOCK = optIns.Value
      RdoPoi.Update
   End If
   
End Sub



Private Sub optRep_Click()
   If optRep.Value = vbChecked Then optCom.Value = vbChecked
   
End Sub

Private Sub optRep_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub



Private Sub txtCmt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub


Private Sub txtCmt_LostFocus()
   
   ' MM removes the double quote
   'txtCmt = CheckLen(txtCmt, 2048)
   
   txtCmt = CheckLenOnly(txtCmt, 2048)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   If Val(txtQty) > 0 Then
      'cmbJmp.Enabled = True
      cmdNew.Enabled = True
      If bCantCancel = 0 Then cmdTrm.Enabled = True
   End If
   If bPoRows Then
      'On Error Resume Next
      RdoPoi!PICOMT = txtCmt
      RdoPoi.Update
   End If
   
End Sub


Private Sub txtDue_DropDown()
   ShowCalendarEx Me
    bFromDropDown = 1   'BBS Added on 03/10/2010 for Ticket #11364
End Sub

Private Sub txtDue_GotFocus()
    bFromDropDown = 0   'BBS Added on 03/10/2010 for Ticket #11364
    bFieldChanged = 0   'BBS Added this to know when the field has changed (Ticket #11364)
End Sub

Private Sub txtDue_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtDue_LostFocus()
   'On Error Resume Next
   'If Len(txtDue) > 0 Then txtDue = CheckDate(txtDue)
   txtDue = CheckDateEx(txtDue)
   If Val(txtQty) > 0 Then
      'cmbJmp.Enabled = True
      cmdNew.Enabled = True
      If bCantCancel = 0 Then cmdTrm.Enabled = True
   End If
   If bPoRows Then
      'On Error Resume Next
      If (Format(RdoPoi!PIPDATE, "mm/dd/yyyy") <> txtDue) And (bFromDropDown <> 1) Then bFieldChanged = 1  'BBS Added on 03/10/2010 for Ticket #11364
      RdoPoi!PIPDATE = txtDue
      RdoPoi.Update
   End If
   UpdateJump
   
   
   If Len(cmbOrigDueDte) = 0 And bFieldChanged = 1 Then bFieldChanged = 0
   
   
   
   
   'BBS Added the logic below on 03/10/2010 for Ticket #11364
   If bFieldChanged = 1 Then
        If MsgBox("Would you also like to change the original due date to " & txtDue & " ?", vbYesNo, Caption) = vbYes Then
            cmbOrigDueDte = txtDue
            cmbOrigDueDte_LostFocus
        End If
   End If
   
   
End Sub

Private Sub txtlot_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtlot_LostFocus()
   txtLot = CheckLen(txtLot, 9)
   txtLot = Format(Abs(Val(txtLot)), ES_QuantityDataFormat)
   If Val(txtQty) > 0 Then
      'cmbJmp.Enabled = True
      cmdNew.Enabled = True
      If bCantCancel = 0 Then cmdTrm.Enabled = True
   End If
   If bPoRows Then
      If (txtLot <> "" And Val(txtLot) > 0) Then
        SetPriceAsUnitCost False
      Else
        lblPrc = ""
      End If
      
      'On Error Resume Next
      RdoPoi!PILOT = Val(txtLot)
      RdoPoi.Update
   End If
   
End Sub

'Private Sub txtOpn_LostFocus()
'    Dim iOPNumber As Integer
'    txtOpn = CheckLen(txtOpn, 3)
'    txtOpn = Format(Abs(Val(txtOpn)), "000")
'    If bPoRows Then
'        iOPNumber = GetPoOp()
'        On Error Resume Next
'        If Trim(cmbMon) <> "" Then
'            RdoPoi!PIRUNOPNO = Val(txtOpn)
'        Else
'            RdoPoi!PIRUNOPNO = 0
'        End If
'        RdoPoi.Update
'    End If
'    If Val(txtOpn) = 0 Then txtOpn = ""
'
'End Sub
'

Private Sub txtPrc_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtPrc_LostFocus()
   txtPrc = CheckLen(txtPrc, 9)
   txtPrc = Format(Abs(Val(txtPrc)), ES_PurchasedDataFormat)
   If Val(txtQty) > 0 Then
      'cmbJmp.Enabled = True
      cmdNew.Enabled = True
      If bCantCancel = 0 Then cmdTrm.Enabled = True
   End If
   If bPoRows Then
      'On Error Resume Next
      If (txtLot <> "" And Val(txtLot) > 0) Then
        SetPriceAsUnitCost True
      Else
        If RdoPoi!PIESTUNIT <> Val(txtPrc) Then
           RdoPoi!PIESTUNIT = Format(Val(txtPrc), ES_PurchasedDataFormat)
           If RdoPoi!PIAMT > 0 Or (RdoPoi!PITYPE >= 15) Then
                RdoPoi!PIAMT = Format(Val(txtPrc), ES_PurchasedDataFormat)
           End If
           
           RdoPoi.Update
           ' update the text
           lblPrc = ""
        End If
      End If
   End If
   
End Sub

Private Sub txtPrt_Change()
   cmbPrt = txtPrt
   LoadExtendedPartDesc
End Sub

Private Sub txtPrt_DblClick()
    'ShowExtendedPartDesc
End Sub

Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "CMBPRT"
      ViewParts.txtPrt = cmbPrt
      optVew.Value = vbChecked
      ViewParts.Show
   End If
   
End Sub


Private Sub txtPrt_LostFocus()
   Dim sDescription As String
   Dim cPartCost As Currency
   'On Error Resume Next
   If bView = 1 Then Exit Sub
   
   txtPrt = CheckLen(txtPrt, 30)
   txtPrt = GetCurrentPart(txtPrt, lblDsc, True)
   LoadExtendedPartDesc
   
   ' don't allow inactive or obsolete part
   If (Not ValidPartNumber(cmbPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      txtPrt = ""
      Exit Sub
   End If

   GetThisPart False
   If Val(txtQty) > 0 Then
      If Len(txtPrt) > 0 Then
         If bNewItem Or Compress(RdoPoi!PIPART) <> Compress(txtPrt) Then
            cPartCost = GetPartCost(txtPrt, ES_STANDARDCOST)
            txtPrc = Format(cPartCost, ES_PurchasedDataFormat)
         End If
         'cmbJmp.Enabled = True
         cmdNew.Enabled = True
         If bCantCancel = 0 Then cmdTrm.Enabled = True
         bNewItem = False
      End If
   End If
   If bPoRows Then
      If lblDsc.ForeColor <> ES_RED Then
         'On Error Resume Next
         If RdoPoi!PIPART <> Compress(txtPrt) Then
            RdoPoi!PIPART = Compress(txtPrt)
            RdoPoi.Update
         End If
         Grd.Col = 2
         Grd.Text = txtPrt
         Grd.Col = 0
      End If
   End If
   If Trim(txtPrt) = "NONE" Then
      MsgBox "Please Select A Valid Part Number.", _
         vbInformation, Caption
   Else
      UpdateJump
   End If
   cmbPrt = txtPrt
   
End Sub


Private Sub cmbAccount_Click()
   'GetAccount    'why?
   
End Sub


Private Sub cmbAccount_LostFocus()
   cmbAccount = CheckLen(cmbAccount, 12)
   'GetAccount
   sOldMo = cmbMon
   
End Sub


Private Sub txtQty_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub


Private Sub FillCombo()
   Dim RdoPon As ADODB.Recordset
   'On Error Resume Next
   'Double check the Service Status 11/7/00
   If optSrv.Value = vbUnchecked Then
      sSql = "SELECT PONUMBER,POSERVICE FROM PohdTable " _
             & "WHERE PONUMBER=" & Val(lblPon) & " "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPon, ES_FORWARD)
      If bSqlRows Then
         With RdoPon
            If Not IsNull(.Fields(1)) Then optSrv.Value = .Fields(1)
            ClearResultSet RdoPon
         End With
      End If
   End If
   
   On Error GoTo DiaErr1
   
   Dim bPartSearch As Boolean
   
   bPartSearch = GetPartSearchOption
   SetPartSearchOption (bPartSearch)
   
   If Not bPartSearch Then
      sSql = "SELECT PARTREF,PARTNUM,PALEVEL FROM PartTable "
      If optSrv.Value = vbChecked Then
         sSql = sSql & "WHERE PALEVEL=7 AND PAINACTIVE = 0 AND PAOBSOLETE = 0"
         'cmbPrt.ToolTipText = "Contains Service Part Type 7's"
      Else
         sSql = sSql & "WHERE PALEVEL<7 AND PAINACTIVE = 0 AND PAOBSOLETE = 0 "
         'cmbPrt.ToolTipText = "Contains Type 1,2,3,4,5,6"
      End If
      sSql = sSql & "ORDER BY PARTREF"
      LoadComboBox cmbPrt
   Else
      cmbPrt = txtPrt
   End If
   'FillMo's
   sSql = "SELECT DISTINCT RUNREF,PARTREF,PARTNUM FROM RunsTable,PartTable " _
          & "WHERE RUNREF=PARTREF AND RUNSTATUS NOT LIKE 'C%'"
   LoadComboBox cmbMon, 1
   
   'bAllowPrice
   sSql = "SELECT AllowPostReceiptPricing FROM Preferences WHERE PreRecord=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPon, ES_FORWARD)
   If bSqlRows Then
      With RdoPon
         If Not IsNull(!AllowPostReceiptPricing) Then _
                       bAllowPrice = !AllowPostReceiptPricing Else _
                       bAllowPrice = 0
         ClearResultSet RdoPon
      End With
   End If
   Set RdoPon = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtQty_LostFocus()
    txtQty = CheckLen(txtQty, 9)
    txtQty = Format(Abs(Val(txtQty)), ES_QuantityDataFormat)
    If cConversion = 0 Then cConversion = 1
    
    If UsingSheetInventory And iPurchaseConversionQty <> 0 And iPurchaseConversionQty <> 1 Then
        lblPoQty = iPurchaseConversionQty
    Else
      lblPoQty = Format((Val(txtQty) / cConversion), ES_QuantityDataFormat)
    End If
   
   If Val(txtQty) > 0 Then
      'cmbJmp.Enabled = True
      cmdNew.Enabled = True
      If bCantCancel = 0 Then cmdTrm.Enabled = True
   Else
      txtQty = Format(1, ES_QuantityDataFormat)
   End If
   If bPoRows Then
      'On Error Resume Next
      RdoPoi!PIPQTY = Format(Val(txtQty), ES_QuantityDataFormat)
      RdoPoi.Update
      Grd.Col = 3
      Grd.Text = Format(txtQty, ES_QuantityDataFormat)
      Grd.Col = 0
   End If
   
End Sub



Private Sub UpdateItem()
   Dim iOpno As Integer
   Dim iRunNo As Long
   Dim cQuantity As Currency
   Dim cPrice As Currency
   Dim sRunPart As String
   Dim strCmt As String
   
   If bOnLoad Then
      bOnLoad = 0
      Exit Sub
   End If
   On Error GoTo DiaErr1
   If txtPrt.Visible Then cmbPrt = txtPrt
   If optSrv.Value = vbChecked Then bGoodMo = GetMoNum()
   sPartNumber = Compress(cmbPrt)
   iRunNo = Val(cmbRun)
   sRunPart = Compress(cmbMon)
   If Len(Trim(cmbMon)) Then
      If bFoundPart = 0 Then
         MsgBox "Invalid MO Part Number .", vbExclamation, Caption
         Exit Sub
      End If
   End If
   'On Error Resume Next
   If Len(Trim(cmbMon)) > 0 Then
      iOpno = Val(cmbOperation)
      
   Else
      iOpno = 0
   End If
   
   'Added Operation 2/14/01
   If Trim(cmbPrt) <> "" Or cmbPrt <> "NONE" Then
      bFoundPart = 1
      cPrice = Format(Val(txtPrc), ES_PurchasedDataFormat)
      cQuantity = Format(Val(txtQty), ES_QuantityDataFormat)
'BBS added on 03/09/2010 for Ticket #11364

        
       If Len(cmbOrigDueDte) = 0 Then cmbOrigDueDte = txtDue
    
        sSql = "UPDATE PoitTable SET PIPART='" & sPartNumber & "'," _
             & "PIPDATE='" & txtDue & "'," _
             & "PIPORIGDATE='" & cmbOrigDueDte & "'," _
             & "PIPQTY=" & cQuantity & "," _
             & "PILOT=0" & txtLot & "," _
             & "PIESTUNIT=" & cPrice & "," _
             & "PIONDOCK=" & optIns.Value & "," _
             & "PIRUNPART='" & sRunPart & "'," _
             & "PIRUNNO=" & iRunNo & "," _
             & "PIRUNOPNO=" & iOpno & "," _
             & "PIACCOUNT='" & Compress(cmbAccount) & "' " _
             & "WHERE PINUMBER=" & lblPon & " " _
             & "AND PIITEM=" & lblItm & " " _
             & "AND PIREV='" & lblRev & "' "
   Else
      sSql = "UPDATE PoitTable SET " _
             & "PIPDATE='" & txtDue & "'," _
             & "PIPORIGDATE='" & cmbOrigDueDte & "'," _
             & "PIPQTY=" & txtQty & "," _
             & "PILOT=0" & txtLot & "," _
             & "PIESTUNIT=" & txtPrc & "," _
             & "PIONDOCK=" & optIns.Value & "," _
             & "PIRUNPART='" & sRunPart & "'," _
             & "PIRUNNO=" & iRunNo & "," _
             & "PIRUNOPNO=" & iOpno & "," _
             & "PIACCOUNT='" & Compress(cmbAccount) & "' " _
             & "WHERE PINUMBER=" & lblPon & " " _
             & "AND PIITEM=" & lblItm & " " _
             & "AND PIREV='" & lblRev & "' "
   End If
   clsADOCon.ExecuteSql sSql
   'Just the comment
   'txtCmt = ReplaceString(txtCmt)
   strCmt = Trim(txtCmt)
   strCmt = ReplaceSingleQuote(strCmt)
   
   txtCmt = ReplaceString(txtCmt)
   sSql = "UPDATE PoitTable SET PICOMT='" & Trim(strCmt) & "' " _
          & "WHERE PINUMBER=" & lblPon & " " _
          & "AND PIITEM=" & lblItm & " " _
          & "AND PIREV='" & lblRev & "' "
   clsADOCon.ExecuteSql sSql
   bNewItem = False
   Exit Sub
   
DiaErr1:
   sProcName = "updateitem"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetMoNum() As Byte
   Dim RdoItm As ADODB.Recordset
   sPartNumber = Compress(cmbMon)
   On Error GoTo DiaErr1
   If sPartNumber = "NONE" Then
      GetMoNum = False
      Set RdoItm = Nothing
      Exit Function
   End If
   If Len(sPartNumber) > 0 Then
      sSql = "SELECT RUNREF,RUNNO,RUNSTATUS FROM RunsTable WHERE " _
             & "RUNREF='" & sPartNumber & "' AND RUNNO=" & Val(cmbRun) & " "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoItm)
      If bSqlRows Then
         With RdoItm
            GetMoNum = True
            ClearResultSet RdoItm
         End With
      Else
         MsgBox "No Current MO Runs Found.", vbExclamation, Caption
         cmbMon = ""
         cmbRun = ""
         GetMoNum = False
      End If
      Set RdoItm = Nothing
   Else
      cmbRun = ""
      GetMoNum = False
   End If
   Set RdoItm = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getmonum"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

'1/17/03 Added option to procedure to fix No PO Items

Private Sub AddPOItem(Optional bNoneFound As Boolean)
   Dim iList As Integer
   Dim iNewItem As Integer
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sRev As String * 2
   Dim sJpart As String * 12
   Dim sNewPart As String
   GetLastItem
   iNewItem = iNextItem
   If txtPrt.Visible Then cmbPrt = txtPrt
   If Not bNoneFound Then
      If bSuppress = 0 Then
         sMsg = "Add An Item To This PO?"
         bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      Else
         bResponse = vbYes
      End If
   Else
      bResponse = vbYes
   End If
   If bItemCan = 0 Then
      sNewPart = sPartNumber
   Else
      cmbPrt = "NONE"
      sNewPart = "NONE"
   End If
   txtPrt = cmbPrt
   LoadExtendedPartDesc
   'On Error Resume Next
   If bResponse = vbYes Then
        'BBS Added on 03/10/2010 for Ticket #11364
        'I am defaulting the original ship date when inserting a PO item record
      sSql = "INSERT INTO PoitTable (PINUMBER," _
             & "PIITEM,PITYPE,PIPART,PIPDATE,PIPORIGDATE,PIVENDOR,PIUSER,PIENTERED,PIACCOUNT) " & vbCrLf _
             & "VALUES(" & lblPon & "," _
             & iNewItem & ",14,'" & sNewPart & "','" _
             & lblDte & "',Null,'" & Compress(PurcPRe02a.cmbVnd) & "'," & vbCrLf _
             & "'" & sInitials & "','" & Format(ES_SYSDATE, "mm/dd/yy") & "','" & cmbAccount.Text & "')"
      clsADOCon.ExecuteSql sSql
      If clsADOCon.RowsAffected > 0 Then
         bCantCancel = 0
         iList = Len(cmbPrt)
         '  sJpart = cmbPrt
         '  cmbJmp.AddItem Trim(Str(iNewItem)) & sRev & "  " & sJpart & "... " & txtDue
         '  cmbJmp = Trim(Str(iNewItem)) & sRev & "  " & sJpart & "... " & txtDue
         iMaxItem = iMaxItem + 1
         vItems(iMaxItem, 0) = iNewItem
         vItems(iMaxItem, 1) = ""
         iTotalItems = iTotalItems + 1
         iIndex = iMaxItem
         lblItm = iNewItem
         iRows = iRows + 1
         If iRows > 1 Then Grd.Rows = Grd.Rows + 1
         Grd.row = iRows
         If Grd.row > 5 Then Grd.TopRow = Grd.row - 4
         Grd.Col = 0
         Grd.Text = lblItm
         lblRev = ""
         txtDue = lblDte
         cmbOrigDueDte = lblDte     'BBS Added this on 03/10/2010 for Ticket #11364
         If optCom.Value = vbUnchecked Then txtCmt = ""
         cmbMon = ""
         cmbRun = ""
         If optRep.Value = vbUnchecked Then
            txtQty = "0.000"
            txtPrc = "0.000"
            lblPrc = ""
            txtLot = ""
            cmbPrt = ""
            txtPrt = ""
            lblDsc = ""
            lblTyp = ""
         End If
         Grd.Col = 2
         Grd.Text = cmbPrt
         Grd.Col = 3
         Grd.Text = txtQty
         Grd.Col = 0
         If bSuppress = 0 Then
            If bNoneFound Then SysMsg "Item Added.", True, Me
         End If
         bSuppress = 0
         bGoodItem = GetThisItem(True)
         cmbJmp.Enabled = False
         cmdNew.Enabled = False
         cmdTrm.Enabled = False
         bNewItem = True
         'On Error Resume Next
         sOldMo = ""
         GetThisItem False
         txtQty.SetFocus
      Else
         MsgBox "Couldn't Add Item.", vbExclamation, Caption
         bGoodItem = 0
      End If
   Else
      CancelTrans
   End If
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   lblPoQty.ToolTipText = "Purchased Quantity. Purchasing Conversion is 1"
   
End Sub

Private Sub GetRuns()
   Dim RdoRns As ADODB.Recordset
   On Error GoTo DiaErr1
   If Compress(sOldMo) = Compress(cmbMon) Then Exit Sub
   sOldMo = cmbMon
   cmbRun.Clear
   sPartNumber = Compress(cmbMon)
   ADOParameter2.Value = sPartNumber
   bSqlRows = clsADOCon.GetQuerySet(RdoRns, AdoQry2)
   If bSqlRows Then
      With RdoRns
         cmbRun = Format(!Runno, "####0")
         lblMon = "" & Trim(!PADESC)
         Do Until .EOF
            AddComboStr cmbRun.hwnd, Format$(!Runno, "####0")
            .MoveNext
         Loop
         ClearResultSet RdoRns
      End With
   Else
      sPartNumber = ""
   End If
   Set RdoRns = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'Find the next item (includes received and canceled)

Private Sub GetLastItem()
   Dim RdoNxt As ADODB.Recordset
   'On Error Resume Next
   sSql = "SELECT MAX(PIITEM) FROM PoitTable " _
          & "WHERE PINUMBER=" & lblPon & " "
   Set RdoNxt = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   If Not IsNull(RdoNxt.Fields(0)) Then
      If Not IsNull(RdoNxt.Fields(0)) Then
         If Val(RdoNxt.Fields(0)) > 0 Then iNextItem = RdoNxt.Fields(0) + 1
      Else
         iNextItem = 1
      End If
   Else
      iNextItem = 1
   End If
   Set RdoNxt = Nothing
   
End Sub



Private Sub ResetBoxes(bOpen As Boolean)
   If Not bOpen Then
      optRcd.Value = vbChecked
      txtQty.Enabled = False
      txtPrc.Enabled = False
      txtLot.Enabled = False
      optIns.Enabled = False
      txtDue.Enabled = False
      cmbPrt.Enabled = False
      txtPrt.Enabled = False
      cmbMon.Enabled = False
      cmbRun.Enabled = False
      cmbOperation.Enabled = False
      txtCmt.Enabled = True 'Fixed for Ticket #78386
   Else
      optRcd.Value = vbUnchecked
      txtQty.Enabled = True
      txtPrc.Enabled = True
      txtLot.Enabled = True
      'optIns.Enabled = True
      txtDue.Enabled = True
      cmbPrt.Enabled = True
      txtPrt.Enabled = True
      cmbMon.Enabled = True
      cmbRun.Enabled = True
      cmbOperation.Enabled = True
      txtCmt.Enabled = True
   End If
End Sub

Private Function FinalItemCheck() As Byte
   Dim RdoChk As ADODB.Recordset
   On Error GoTo DiaErr1
   MouseCursor 13
   sSql = "SELECT PINUMBER,PIPART FROM PoitTable WHERE " _
          & "(PIPART='NONE' OR PIPART='') AND PINUMBER=" _
          & Val(lblPon) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk, ES_FORWARD)
   If bSqlRows Then
      With RdoChk
         FinalItemCheck = 1
         ClearResultSet RdoChk
      End With
   Else
      FinalItemCheck = 0
   End If
   Set RdoChk = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "finalitemc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub GetThisPart(partNumberChange As Boolean)
   'PartNumberChange = True if new item was just added or changed
   '(don't update zero cost with std cost)
   Dim RdoPrt As ADODB.Recordset
   On Error GoTo DiaErr1
   sProcName = "getthispart"
   If txtPrt.Visible Then cmbPrt = txtPrt
   sSql = "SELECT PARTREF,PAONDOCK,PALEVEL,PAUNITS,PAPURCONV,PAPUNITS,PASTDCOST,PAQOH FROM " _
          & "PartTable WHERE PARTREF='" & Compress(cmbPrt) & "'"
          Debug.Print sSql
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt)
   If bSqlRows Then
      With RdoPrt
        iPurchaseConversionQty = !PAPURCONV
         If Not IsNull(!PAONDOCK) Then
            optIns.Value = !PAONDOCK
         Else
            optIns.Value = vbUnchecked
         End If
         
         If Not IsNull(!PAQOH) Then
            lblQoh = !PAQOH
         Else
            lblQoh = "0"
         End If
         'update zero cost with std cost or last invoiced cost if this is a new item
         'if using invoiced cost, use lowest cost on most recent day
         If partNumberChange Then
'            If Val(txtPrc) = 0 Then
               Dim cost As Currency
               If UseLastInvoicedPrice Then
                  Dim rdo As ADODB.Recordset
                  sSql = "select top 1 PIAMT from PoitTable" & vbCrLf _
                     & "where PIPART = '" & Compress(cmbPrt) & "' and PITYPE = 17" & vbCrLf _
                     & "order by PIADATE desc, PIAMT asc"
                  If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
                     cost = rdo!PIAMT
                  Else
                     cost = !PASTDCOST
                  End If
               Else
                  cost = !PASTDCOST
               End If
               txtPrc = Format(cost, ES_PurchasedDataFormat)
               txtPrc_LostFocus     'force write to PIESTUNIT
 '           End If
         End If
         
'<<<<< this was the code before sheet inventory
'         'just use the default setting
'         If !PAPURCONV = 0 Then cConversion = 0 Else cConversion = !PAPURCONV
'         lblUom = "" & Trim(!PAUNITS)
'         If cConversion = 0 Then cConversion = 1
'         lblPoQty = Format((Val(txtQty) / cConversion), ES_QuantityDataFormat)
'         lblPoQty.ToolTipText = "Purchased Quantity. Purchasing Conversion is " & cConversion
'         lblPoUom = "" & Trim(!PAPUNITS)
'         lblTyp = !PALEVEL
'>>>>>>

'<<<<< this was the first attempt at Sheet Inventory by Terry
'        If UsingSheetInventory And !PAPURCONV <> 0 And !PAPURCONV <> 1 Then
'            lblUom = !PAUNITS
'            lblPoQty = !PAPURCONV   ' * Val(txtQty)
'            lblPoQty.ToolTipText = "Purchase Conversion Quantity per " + !PAPUNITS + " = " + CStr(!PAPURCONV) + " " + !PAUNITS
'            lblPoUom = !PAUNITS
'            lblTyp = !PALEVEL
'         Else
'
'            'just use the default setting
'            If !PAPURCONV = 0 Then cConversion = 0 Else cConversion = !PAPURCONV
'            lblUom = "" & Trim(!PAUNITS)
'            If cConversion = 0 Then cConversion = 1
'            lblPoQty = Format((Val(txtQty) / cConversion), ES_QuantityDataFormat)
'            lblPoQty.ToolTipText = "Purchased Quantity. Purchasing Conversion is " & cConversion
'            lblPoUom = "" & Trim(!PAPUNITS)
'            lblTyp = !PALEVEL
'        End If
'>>>>>

'<<<<< this is how things ended up
        If IsThisASheetPart(Compress(txtPrt)) Then
            lblUom = !PAPUNITS
            lblPoQty = ""
            lblPoQty.ToolTipText = ""
            lblPoUom = ""
            lblTyp = !PALEVEL
         Else
            'just use the default setting
            If !PAPURCONV = 0 Then cConversion = 0 Else cConversion = !PAPURCONV
            lblUom = "" & Trim(!PAUNITS)
            If cConversion = 0 Then cConversion = 1
            lblPoQty = Format((Val(txtQty) / cConversion), ES_QuantityDataFormat)
            lblPoQty.ToolTipText = "Purchased Quantity. Purchasing Conversion is " & cConversion
            lblPoUom = "" & Trim(!PAPUNITS)
            lblTyp = !PALEVEL
        End If

'>>>>>
         
         ClearResultSet RdoPrt
      End With
   End If
   Set RdoPrt = Nothing
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub UpdateJump()
   
End Sub

Private Sub GetAccount()
   '    Dim rdoAct As ADODB.recordset
   '    If Trim(cmbAccount) <> "" Then
   '        sSql = "Qry_GetAccount '" & Compress(cmbAccount) & "'"
   '        bsqlrows = clsadocon.getdataset(ssql, rdoAct, ES_FORWARD)
   '            If bSqlRows Then
   '                With rdoAct
   '                    cmbAccount = "" & Trim(!GLACCTNO)
   '                    cmbAccount.ToolTipText = "" & Trim(!GLDESCR)
   '                    ClearResultSet rdoAct
   '                End With
   '            End If
   '    End If
   
   If bOnLoad = 0 Then
      Dim gl As New GLTransaction
      gl.FillComboWithAccounts cmbAccount
   End If
End Sub

Private Sub FillAccounts()
   Dim rdo As ADODB.Recordset
   sSql = "Qry_FillLowAccounts"
   LoadComboBox cmbAccount
   
   'set the po acct to the default
   sSql = "SELECT rtrim(COAPACCT) FROM ComnTable WHERE COREF=1"
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      If Len(rdo.Fields(0)) > 0 Then
         'On Error Resume Next
         cmbAccount = rdo.Fields(0)
         Err.Clear
      End If
   End If
   Set rdo = Nothing
End Sub

Private Function CheckPoAccounts() As Byte
   'On Error Resume Next
   Dim RdoChk As ADODB.Recordset
   CheckPoAccounts = 0
   sSql = "SELECT isnull(PurchaseAccount, 0) FROM Preferences WHERE " _
          & "PreRecord=1"
   If clsADOCon.GetDataSet(sSql, RdoChk, ES_FORWARD) Then
      CheckPoAccounts = RdoChk.Fields(0)
   End If
   
'   'set the po acct to the default
'   sSql = "SELECT COAPACCT FROM ComnTable WHERE COREF=1"
'   If GetDataSet(RdoChk, ES_FORWARD) Then
'      On Error Resume Next
'      Me.cmbAccount = RdoChk.fields(0)
'      Err.Clear
'   End If
   
End Function



'3/22/07 New

Private Function GetPoOp() As Integer
   Dim RdoOpn As ADODB.Recordset
   On Error GoTo DiaErr1
   If cmbMon <> "" And Val(cmbRun) > 0 Then
      sSql = "SELECT OPREF,OPRUN,OPNO FROM RnopTable WHERE " _
                & "(OPREF='" & Compress(cmbMon) & "' AND OPRUN=" _
                & Val(cmbRun) & " AND OPNO=" & Val(cmbOperation) & ")"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoOpn, ES_FORWARD)
      If bSqlRows Then
         GetPoOp = RdoOpn!opNo
      Else
         If (Val(cmbOperation) <> 0) Then
            MsgBox "MO Operation " & cmbOperation & " Wasn't Found.", _
               vbInformation, Caption
         End If
         GetPoOp = 0
      End If
   Else
      GetPoOp = 0
   End If
   Set RdoOpn = Nothing
   If GetPoOp = 0 Then
      Beep
      cmbOperation = ""
   Else
      cmbOperation = Format(GetPoOp, "000")
   End If
   Exit Function
   
DiaErr1:
   Beep
   GetPoOp = 0
   cmbOperation = ""
   
End Function

Private Sub FillOps()
   cmbOperation.Clear
   If cmbRun = "" Or lblTyp = "" Then Exit Sub
   
   sSql = "SELECT OPNO" & vbCrLf _
          & "FROM RnopTable" & vbCrLf _
          & "WHERE OPREF='" & Compress(cmbMon) & "' AND OPRUN=" & cmbRun & vbCrLf
   
   'service?
   If Val(lblTyp) = 7 Then
      Grd.Col = 2
      sSql = sSql & "AND OPSERVPART='" & Compress(Grd.Text) & "'" & vbCrLf
      
      'pick item
   Else
      sSql = sSql & "AND OPPICKOP>0" & vbCrLf
   End If
   
   sSql = sSql & "ORDER BY OPNO"
   Dim sSaveSql As String
   
   sSaveSql = sSql
   LoadComboBox Me.cmbOperation, -1
   
   'if no operation for service found, add an item
   If cmbOperation.ListCount = 0 And Val(lblTyp) = 7 Then
      
      
      sSql = sSaveSql
      LoadComboBox Me.cmbOperation
   End If
   
End Sub

Private Function SetPriceAsUnitCost(bFromPrice As Boolean)
    
    Dim strMsg As String
    Dim bResponse As Byte
    
    strMsg = "Is the Price a Lot Charge?"
    bResponse = MsgBox(strMsg, ES_YESQUESTION, Caption)
    
    If (bResponse = vbYes) Then
        lblPrc = txtPrc
        txtPrc = Format((Val(txtPrc) / Val(txtLot)), ES_PurchasedDataFormat)
        RdoPoi!PIESTUNIT = Format(Val(txtPrc), ES_PurchasedDataFormat)
        RdoPoi.Update
        txtPrc.ToolTipText = "Price as Unit cost"
    Else
        If (bFromPrice) Then
            lblPrc = ""
            txtLot = Format(0, ES_QuantityDataFormat)
            RdoPoi!PILOT = Val(txtLot)
            RdoPoi.Update
        End If
    End If

End Function


Private Sub LoadExtendedPartDesc()
   Dim RdoPrt As ADODB.Recordset
   Dim sExtendedInfo As String

   On Error GoTo SEPD1
   
   sExtendedInfo = ""
   AdoParameter1.Value = Compress(txtPrt)
   bSqlRows = clsADOCon.GetQuerySet(RdoPrt, AdoQry1, ES_KEYSET)
   If bSqlRows Then
      With RdoPrt
         sExtendedInfo = "" & Trim(!PAEXTDESC)
         ClearResultSet RdoPrt
      End With
   End If
   Set RdoPrt = Nothing
   If Len(sExtendedInfo) = 0 Then sExtendedInfo = "Enter Part Number" & vbCrLf & "(Extended Part Description will display here if available)"
   If (sExtendedInfo <> "") Then
      txtPrt.ToolTipText = sExtendedInfo
      cmbPrt.ToolTipText = sExtendedInfo
'      FusionToolTip.ToolText(txtPrt) = sExtendedInfo
   End If
   
   Exit Sub
   
SEPD1:
   'Ingore Errors on this routine
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


