VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form PurcPRe02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revise A Purchase Order"
   ClientHeight    =   5670
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   7950
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   4302
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox OptAutoPO 
      Caption         =   "FromAutoPO"
      Height          =   195
      Left            =   4440
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox cbTaxable 
      Height          =   255
      Left            =   4200
      TabIndex        =   56
      Top             =   360
      Width           =   375
   End
   Begin VB.Frame tabFrame 
      Height          =   3252
      Index           =   1
      Left            =   8040
      TabIndex        =   49
      Top             =   2280
      Width           =   7500
      Begin VB.CommandButton cmdComments 
         DisabledPicture =   "PurcPRe02a.frx":0000
         DownPicture     =   "PurcPRe02a.frx":0972
         Height          =   350
         Index           =   1
         Left            =   5280
         Picture         =   "PurcPRe02a.frx":12E4
         Style           =   1  'Graphical
         TabIndex        =   50
         TabStop         =   0   'False
         ToolTipText     =   "Standard Comments"
         Top             =   960
         Width           =   350
      End
      Begin VB.ComboBox txtVia 
         Height          =   288
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   13
         Tag             =   "2"
         Top             =   1800
         Width           =   3015
      End
      Begin VB.ComboBox cmbTrm 
         Height          =   288
         Left            =   1440
         TabIndex        =   10
         Tag             =   "3"
         ToolTipText     =   "Select Shipping Terms"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtShp 
         Height          =   735
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Tag             =   "9"
         Top             =   960
         Width           =   3735
      End
      Begin VB.TextBox txtFob 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Tag             =   "3"
         ToolTipText     =   "FOB"
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ship Via"
         Height          =   252
         Index           =   15
         Left            =   120
         TabIndex        =   54
         Top             =   1800
         Width           =   972
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Shipping Terms "
         Height          =   252
         Index           =   13
         Left            =   120
         TabIndex        =   53
         Top             =   240
         Width           =   1452
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ship To"
         Height          =   252
         Index           =   9
         Left            =   120
         TabIndex        =   52
         Top             =   960
         Width           =   852
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "FOB"
         Height          =   252
         Index           =   8
         Left            =   120
         TabIndex        =   51
         Top             =   600
         Width           =   732
      End
   End
   Begin VB.Frame tabFrame 
      Height          =   3372
      Index           =   0
      Left            =   120
      TabIndex        =   39
      Top             =   2280
      Width           =   7740
      Begin VB.TextBox txtPhone 
         Height          =   285
         Left            =   4680
         TabIndex        =   61
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton cmdStatCd 
         Height          =   350
         Left            =   7320
         Picture         =   "PurcPRe02a.frx":18E6
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Status Code"
         Top             =   1920
         Width           =   350
      End
      Begin VB.TextBox txtCmt 
         Height          =   1320
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Tag             =   "9"
         Top             =   1920
         Width           =   5295
      End
      Begin VB.CommandButton cmdComments 
         DisabledPicture =   "PurcPRe02a.frx":1D75
         DownPicture     =   "PurcPRe02a.frx":26E7
         Height          =   350
         Index           =   0
         Left            =   6840
         Picture         =   "PurcPRe02a.frx":3059
         Style           =   1  'Graphical
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   "Standard Comments"
         Top             =   1920
         Width           =   350
      End
      Begin VB.TextBox txtCnt 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Tag             =   "2"
         ToolTipText     =   "Vendor Contact"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.ComboBox cmbDiv 
         ForeColor       =   &H00800000&
         Height          =   288
         Left            =   5280
         Sorted          =   -1  'True
         TabIndex        =   5
         Tag             =   "8"
         ToolTipText     =   "Select Division From List"
         Top             =   240
         Width           =   780
      End
      Begin VB.ComboBox txtSdt 
         Height          =   315
         Left            =   3120
         TabIndex        =   4
         Tag             =   "4"
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox CmbByr 
         Height          =   288
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   6
         Tag             =   "2"
         Top             =   720
         Width           =   2775
      End
      Begin VB.ComboBox cmbReq 
         Height          =   288
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   8
         Tag             =   "2"
         ToolTipText     =   "Contains Previous Entries (20 Char Max)"
         Top             =   1440
         Width           =   2295
      End
      Begin VB.ComboBox cmbTyp 
         Height          =   288
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   3
         Tag             =   "3"
         ToolTipText     =   "(2) Character Purchase Order Type.  Searches Previous Entries"
         Top             =   240
         Width           =   732
      End
      Begin VB.Label z1 
         Caption         =   "Phone"
         Height          =   255
         Index           =   17
         Left            =   4080
         TabIndex        =   60
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   252
         Index           =   10
         Left            =   120
         TabIndex        =   48
         Top             =   1800
         Width           =   732
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   252
         Index           =   3
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   1092
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ship Date"
         Height          =   252
         Index           =   4
         Left            =   2280
         TabIndex        =   46
         Top             =   240
         Width           =   852
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Buyer"
         Height          =   252
         Index           =   5
         Left            =   120
         TabIndex        =   45
         Top             =   720
         Width           =   732
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         Height          =   255
         Index           =   6
         Left            =   4440
         TabIndex        =   44
         Top             =   240
         Width           =   975
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact"
         Height          =   252
         Index           =   7
         Left            =   120
         TabIndex        =   43
         Top             =   1080
         Width           =   732
      End
      Begin VB.Label lblByr 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4320
         TabIndex        =   42
         Top             =   720
         Width           =   2892
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Requested By"
         Height          =   252
         Index           =   16
         Left            =   120
         TabIndex        =   41
         Top             =   1440
         Width           =   1452
      End
   End
   Begin MSComctlLib.TabStrip tab1 
      Height          =   3975
      Left            =   0
      TabIndex        =   38
      Top             =   1920
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   7011
      TabWidthStyle   =   2
      TabFixedWidth   =   1409
      TabFixedHeight  =   473
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Shipping"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   37
      Top             =   1800
      Width           =   7572
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRe02a.frx":365B
      Style           =   1  'Graphical
      TabIndex        =   36
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtPdt 
      Height          =   315
      Left            =   4200
      TabIndex        =   1
      Tag             =   "4"
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton optPrn 
      DownPicture     =   "PurcPRe02a.frx":3E09
      Height          =   320
      Left            =   6120
      Picture         =   "PurcPRe02a.frx":439B
      Style           =   1  'Graphical
      TabIndex        =   34
      TabStop         =   0   'False
      ToolTipText     =   "Print This PO"
      Top             =   1320
      Width           =   350
   End
   Begin VB.CheckBox optPvw 
      Caption         =   "PoView"
      Height          =   195
      Left            =   360
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "PurcPRe02a.frx":492D
      Height          =   320
      Left            =   6120
      Picture         =   "PurcPRe02a.frx":529F
      Style           =   1  'Graphical
      TabIndex        =   32
      TabStop         =   0   'False
      ToolTipText     =   "Show Existing Purchase Orders"
      Top             =   960
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.CheckBox optNwl 
      Caption         =   "New Loaded"
      Height          =   255
      Left            =   0
      TabIndex        =   29
      Top             =   5880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   315
      Left            =   6600
      TabIndex        =   16
      ToolTipText     =   "Enter New Purchase Order"
      Top             =   1320
      Width           =   915
   End
   Begin VB.CheckBox optNew 
      Caption         =   "New Po"
      Height          =   255
      Left            =   2520
      TabIndex        =   28
      Top             =   5760
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.CheckBox optItm 
      Caption         =   "Items"
      Height          =   255
      Left            =   1800
      TabIndex        =   27
      Top             =   5760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdItm 
      Caption         =   "&Items"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6600
      TabIndex        =   14
      ToolTipText     =   "Add/Edit Po Items"
      Top             =   600
      Width           =   915
   End
   Begin VB.CommandButton cmdTrm 
      Caption         =   "&Terms"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6600
      TabIndex        =   15
      ToolTipText     =   "Edit Po Terms"
      Top             =   960
      Width           =   915
   End
   Begin VB.CheckBox optSrv 
      Enabled         =   0   'False
      Height          =   195
      Left            =   1560
      TabIndex        =   24
      Top             =   360
      Width           =   255
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Tag             =   "3"
      Top             =   1080
      Width           =   1555
   End
   Begin VB.ComboBox cmbPon 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Select Or Enter PO (Contains Last 300 Max)"
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6600
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6840
      Top             =   5760
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5670
      FormDesignWidth =   7950
   End
   Begin VB.Label Label2 
      Caption         =   "Taxable?"
      Height          =   255
      Left            =   3240
      TabIndex        =   57
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "More >>>"
      Height          =   252
      Left            =   5640
      TabIndex        =   55
      Top             =   600
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Label lblPO 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   2880
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Printed"
      Height          =   255
      Index           =   14
      Left            =   3240
      TabIndex        =   31
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblPrn 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4200
      TabIndex        =   30
      Top             =   1080
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Release"
      Height          =   252
      Index           =   1
      Left            =   5160
      TabIndex        =   26
      Top             =   1920
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Service PO"
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   25
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Date"
      Height          =   255
      Index           =   11
      Left            =   3240
      TabIndex        =   23
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblPdt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lblNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   21
      ToolTipText     =   "Vendor - Cannot Be Changed After PO Items Are Received"
      Top             =   1440
      Width           =   3555
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   20
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label txtRel 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   6360
      TabIndex        =   19
      Top             =   1920
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   18
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "PurcPRe02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'4/16/04 Added Services Items as separate dialog
'6/14/04 Added Requested By
'1/10/04 Added the Tab
'11/17/05 Added cmbTyp and RefreshPOCursor.  Removed RdoRes
'5/9/06   Fixed PO Swap between revise/print
'8/8/06 Replaced SSTab32
Option Explicit
Dim RdoPon As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter

Dim RdoRcd As ADODB.Command
Dim ADOParameter2 As ADODB.Parameter

Dim RdoPso As ADODB.Recordset

Dim bCancel As Byte
Dim bForceServ As Byte
Dim bAllowPrice As Byte
Dim bGoodPo As Byte
Dim bGoodVnd As Byte
Dim bNewLoaded As Byte
Dim bOnLoad As Byte
Dim bPrint As Byte

Dim sOrigVendor As String
Dim sPoRef As String

'Dim FusionToolTip As New clsToolTip

Dim cNetDays As Currency
Dim cDDays As Currency
Dim cDiscount As Currency
Dim cProxDt As Currency
Dim cProxdue As Currency

Dim sFob As String
Dim sBcontact As String


Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   cmbReq.Tag = 2
   
End Sub

Private Sub cbTaxable_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
End Sub

Private Sub cbTaxable_LostFocus()
    If bGoodPo Then
        On Error Resume Next
        If cbTaxable.Value = 1 Then RdoPso!POTAXABLE = 1 Else RdoPso!POTAXABLE = 0
        RdoPso.Update
        If Err > 0 Then ValidateEdit
    End If
End Sub


Private Sub cmbByr_Click()
   If Len(Trim(cmbByr)) > 0 Then GetCurrentBuyer cmbByr
   
End Sub

Private Sub CmbByr_DropDown()
   bDataHasChanged = True
   
End Sub


Private Sub cmbByr_LostFocus()
   cmbByr = CheckLen(cmbByr, 20)
   cmbByr = StrCase(cmbByr)
   If bGoodPo Then
      On Error Resume Next
      GetCurrentBuyer cmbByr
      RdoPso!POBUYER = "" & Compress(cmbByr)
      RdoPso.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub cmbDiv_DropDown()
   bDataHasChanged = True
   
End Sub

Private Sub cmbDiv_LostFocus()
   cmbDiv = CheckLen(cmbDiv, 4)
   If Len(cmbDiv) Then
      If bGoodPo Then
         On Error Resume Next
         RdoPso!PODIVISION = "" & cmbDiv
         RdoPso.Update
         If Err > 0 Then ValidateEdit
      End If
   End If
   
End Sub


Private Sub cmbPon_Click()
   bGoodPo = GetPurchaseOrder(0)
   LoadVendorAddress
End Sub


Private Sub cmbPon_GotFocus()
   bGoodPo = GetPurchaseOrder(0)
   
End Sub


Private Sub cmbPon_LostFocus()
   cmbPon = CheckLen(cmbPon, 6)
   cmbPon = Format(Abs(Val(cmbPon)), "000000")
   If bCancel Then Exit Sub
   bGoodPo = GetPurchaseOrder(1)
   
End Sub

Private Sub cmbReq_DropDown()
   bDataHasChanged = True
   
End Sub

Private Sub cmbReq_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   cmbReq = CheckLen(cmbReq, 20)
   On Error Resume Next
   If Len(cmbReq) < 4 Then cmbReq = UCase$(cmbReq) _
          Else cmbReq = StrCase(cmbReq)
   If bGoodPo Then
      RdoPso!POREQBY = Trim(cmbReq)
      RdoPso.Update
      If Err > 0 Then ValidateEdit
   End If
   For iList = 0 To cmbReq.ListCount - 1
      If cmbReq = cmbReq.List(iList) Then b = 1
   Next
   If b = 0 Then cmbReq.AddItem cmbReq
   
End Sub


Private Sub cmbTrm_DropDown()
   bDataHasChanged = True
   
End Sub

Private Sub cmbTrm_LostFocus()
   cmbTrm = CheckLen(cmbTrm, 2)
   If bGoodPo Then
      On Error Resume Next
      RdoPso!POSTERMS = "" & cmbTrm
      RdoPso.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub cmbTyp_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   cmbTyp = CheckLen(cmbTyp, 2)
   If bGoodPo Then
      On Error Resume Next
      RdoPso!POTYPE = "" & cmbTyp
      RdoPso.Update
      If Err > 0 Then ValidateEdit
   End If
   For iList = 0 To cmbTyp.ListCount - 1
      If cmbTyp = cmbTyp.List(iList) Then bByte = 1
   Next
   If bByte = 0 And cmbTyp <> "" Then cmbTyp.AddItem cmbTyp
   
End Sub



Private Sub cmbVnd_Click()
    Dim strReason As String
    bGoodVnd = FindVendor(cmbVnd, lblNme)
    LoadVendorAddress
    If bGoodVnd Then
        IsVendorApproved cmbVnd.Text, 1, strReason
    End If
End Sub

Private Sub cmbVnd_DropDown()
   bDataHasChanged = True
   
End Sub


Private Sub cmbVnd_GotFocus()
   'cmbVnd_Click
   
End Sub


Private Sub cmbVnd_LostFocus()
   Dim sRefVnd
   sRefVnd = Compress(cmbVnd)
   cmbVnd = CheckLen(cmbVnd, 10)
   
   bGoodVnd = FindVendor(cmbVnd, lblNme)
   
   If bGoodPo And bGoodVnd Then
      On Error Resume Next
      If sOrigVendor <> cmbVnd And Len(Trim(cmbVnd)) > 0 Then
         If MsgBox("Are you sure you want to change the Vendor?", vbYesNoCancel, "Change Vendor") = vbYes Then
            sOrigVendor = sRefVnd
            ' Get vendor terms
            GetVendorTerms (sRefVnd)
            
            RdoPso!POVENDOR = "" & sRefVnd
            RdoPso!PODDAYS = cDDays
            RdoPso!PODISCOUNT = cDiscount
            RdoPso!POPROXDT = cProxDt
            RdoPso!POPROXDUE = cProxdue
            RdoPso!POFOB = sFob
            RdoPso!POBCONTACT = sBcontact
            RdoPso.Update
            
            If Err > 0 Then ValidateEdit
            ' Update the PoitTable vendor field
            UpdatePoitVendor sRefVnd, Val(cmbPon)
         Else
            cmbVnd = sOrigVendor
         End If
      End If
   End If
   
End Sub


Private Sub cmdCan_Click()
   Dim sString As String
   Unload Me
   
End Sub



Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   bCancel = True
   
End Sub


Private Sub cmdComments_Click(Index As Integer)
   If Index = 0 Then
      If cmdComments(0) Then
         'See List For Index
         txtCmt.SetFocus
         SysComments.lblListIndex = 0
         SysComments.Show
         cmdComments(0) = False
      End If
   Else
      If cmdComments(1) Then
         'See List For Index
         txtShp.SetFocus
         SysComments.lblControl = "txtShp"
         SysComments.lblListIndex = 1
         SysComments.Show
         cmdComments(1) = False
      End If
   End If
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4302
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdItm_Click()
   Dim bByte As Byte
   If optSrv.Value = vbChecked Then
      If bForceServ = 0 Then
         bByte = GetSetting("Esi2000", "EsiProd", "ForceSevens", bByte)
         If bByte = 0 Then MsgBox "In Order To Force Routing Purchasing, A " & vbCr _
                    & "System Setting Is Required. See Administrator.", vbInformation, Caption
         SaveSetting "Esi2000", "EsiProd", "ForceSevens", 1
      End If
   End If
   If bGoodPo Then
      MouseCursor 13
      UpdatePurchaseOrder
      optItm.Value = vbChecked
      If optSrv.Value = vbChecked Then
         If bForceServ = 1 Then
            PurcPRe02c.lblPon = cmbPon
            PurcPRe02c.bAllowPrice = bAllowPrice
            PurcPRe02c.Show
         Else
            PurcPRe02b.Show
         End If
      Else
         PurcPRe02b.Show
      End If
   Else
      MsgBox "Requires A Valid PO.", vbInformation, Caption
   End If
   
End Sub

Private Sub cmdNew_Click()
   On Error Resume Next
   cmbPon.SetFocus
   PurcPRe01a.optRev = vbChecked
   optNwl.Value = vbChecked
   If optSrv.Value = vbChecked Then
      bPOCaption = 1
      PurcPRe01a.optSrv.Value = vbChecked
      PurcPRe01a.Caption = "New Services Purchase Order"
   End If
   PurcPRe01a.Show
   
End Sub

Private Sub cmdStatCd_Click()
    StatusCode.lblSCTypeRef = "Vendor"
    StatusCode.txtSCTRef = cmbVnd
    StatusCode.LableRef1 = "PO"
    StatusCode.lblSCTRef1 = cmbPon
    StatusCode.lblSCTRef2 = ""
    StatusCode.lblStatType = "PO"
    StatusCode.lblSysCommIndex = 14 ' The index in the Sys Comment "MO Comments"
    StatusCode.txtCurUser = cUR.CurrentUser
    
    StatusCode.Show

End Sub

Private Sub cmdTrm_Click()
   On Error Resume Next
   If bGoodPo Then
      'optTrm.Value = vbChecked
      PurcPRe01b.txtPdsc(0) = Format(RdoPso!PODISCOUNT, "#0.0")
      PurcPRe01b.txtPdsc(1) = Format(RdoPso!PODISCOUNT, "#0.0")
      PurcPRe01b.txtPnet = Format(RdoPso!PONETDAYS, "#0")
      PurcPRe01b.txtPday = Format(RdoPso!PODDAYS, "#0")
      PurcPRe01b.txtPdte = Format(RdoPso!POPROXDT, "#0")
      PurcPRe01b.txtPdue = Format(RdoPso!POPROXDUE, "#0")
      PurcPRe01b.Show
   Else
      MsgBox "Requires A Valid PO.", vbInformation, Caption
   End If
   
End Sub


Private Sub cmdVew_Click()
   '    If cmdVew.Value = True Then
   '        PurcPOtree.Show
   '        cmdVew.Value = False
   '    End If
   
End Sub







Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   bPrint = 0
   If optItm.Value = vbChecked Then
      Unload PurcPRe02b
      optItm.Value = vbUnchecked
   End If
   If optNwl.Value = vbChecked Then
      Unload PurcPRe01a
      optNwl.Value = vbUnchecked
   End If
   If bOnLoad Then
      FillBuyers
      FillVia
      FillTerms
      FillDivisions
      FillVendors
      FillCombo
      bOnLoad = 0
   Else
      MouseCursor 0
      bGoodPo = GetPurchaseOrder(0)
   End If

   If OptAutoPO = vbChecked Then
      OptAutoPO = vbUnchecked
      MrplMRe02.OptAutoPO = vbChecked
      Unload MrplMRe02
   End If

   Call HookToolTips
   MVBBubble.MaxTipWidth = 600
End Sub

Private Sub Form_GotFocus()
   MDISect.lblBotPanel.Caption = "Revise A Purchase Order"
   
End Sub


Private Sub Form_Load()
   FormLoad Me
   FormatControls
   'Tab1.Tab = 0
   MouseCursor 13
   tabFrame(0).BorderStyle = 0
   tabFrame(1).BorderStyle = 0
   tabFrame(0).Visible = True
   tabFrame(1).Visible = False
   tabFrame(0).Left = 10
   tabFrame(1).Left = 10
   sSql = "SELECT PONUMBER,POTYPE,POVENDOR,PODATE," _
          & "POBUYER,POREQBY,PODIVISION,POSHIP,POSERVICE,POSHIPTO," _
          & "POBCONTACT,POFOB,POPRINTED,POREMARKS,POSTERMS," _
          & "PONETDAYS,PODDAYS,PODISCOUNT,POPROXDT,POPROXDUE,POBUYER," _
          & "POVIA,POTAXABLE,VEREF,VENICKNAME,VEBNAME,VEBPHONE FROM PohdTable,VndrTable " _
          & "WHERE PONUMBER = ? AND POVENDOR=VEREF AND POCAN=0"
   Set RdoPon = New ADODB.Command
   RdoPon.CommandText = sSql
   
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adInteger
   RdoPon.Parameters.Append AdoParameter1
   
   
   sSql = "SELECT PINUMBER,PITYPE FROM PoitTable " _
          & "WHERE PINUMBER= ? AND PITYPE<>14"
   Set RdoRcd = New ADODB.Command
   Set ADOParameter2 = New ADODB.Parameter
   ADOParameter2.Type = adInteger
   RdoRcd.Parameters.Append ADOParameter2
   
  ' RdoRcd.MaxRows = 1
   bOnLoad = 1
'   With FusionToolTip
'        Call .Create(Me)
'        .MaxTipWidth = 240
'        .DelayTime(ttDelayShow) = 20000
'        Call .AddTool(cmbVnd)
'    End With
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If bGoodPo = 1 And bDataHasChanged Then UpdatePurchaseOrder
   SaveSetting "Esi2000", "EsiProd", "LastBuyer", Trim(cmbByr)
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload bPrint
   UnhookToolTips
   Set AdoParameter1 = Nothing
   Set ADOParameter2 = Nothing
   Set RdoPso = Nothing
   Set RdoPon = Nothing
   Set RdoRcd = Nothing
   Set PurcPRe02a = Nothing
End Sub

Private Sub FillCombo()
   Dim RdoRpo As ADODB.Recordset
   
'   Dim sYear As String
   On Error GoTo DiaErr1
   sSql = "SELECT ForceServices, AllowPostReceiptPricing  FROM Preferences WHERE PreRecord=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpo, ES_FORWARD)
   If bSqlRows Then
      With RdoRpo
         If Not IsNull(!ForceServices) Then bForceServ = !ForceServices _
                       Else bForceServ = 0
         If Not IsNull(!AllowPostReceiptPricing) Then bAllowPrice = !AllowPostReceiptPricing _
                       Else bAllowPrice = 0
         ClearResultSet RdoRpo
      End With
   End If
   Set RdoRpo = Nothing
   
'   sYear = Format(Now - 760, "yyyy-mm-dd")
'   sSql = "Qry_FillPurchaseOrders '" & sYear & "'"
   sSql = "Qry_FillPurchaseOrders '" & DateAdd("d", -760, Now) & "'"
   LoadNumComboBox cmbPon, "000000"
   
   sSql = "SELECT DISTINCT POREQBY FROM PohdTable ORDER BY POREQBY"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpo, ES_FORWARD)
   LoadComboBox cmbReq, -1
   If cmbPon.ListCount > 0 Then
      If lblPO = "" Then cmbPon = cmbPon.List(0) _
                 Else cmbPon = lblPO
      'bGoodPo = GetPurchaseOrder(0)
   Else
      MouseCursor 0
      If optNew = vbUnchecked Then MsgBox "There Are No Purchase Orders.", vbInformation, Caption
   End If
   
   sSql = "SELECT DISTINCT POTYPE FROM PohdTable WHERE POTYPE<>''"
   LoadComboBox cmbTyp, -1
   On Error Resume Next
   Set RdoRpo = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Function GetPurchaseOrder(Optional bMessage As Byte) As Byte
   Dim RdoVnd As ADODB.Recordset
   On Error GoTo DiaErr1
   If bGoodPo = 1 And bDataHasChanged Then UpdatePurchaseOrder
   MouseCursor 13
   ClearBoxes
   RdoPon.Parameters(0).Value = Val(cmbPon)
'   RdoPon.RowsetSize = 1
'   RdoPon(0) = Val(cmbPon)
   bSqlRows = clsADOCon.GetQuerySet(RdoPso, RdoPon, ES_KEYSET, True, 1)
   If bSqlRows Then
      With RdoPso
         cmbPon = Format(!PONUMBER, "000000")
         cmbVnd = "" & Trim(!VENICKNAME)
         sOrigVendor = cmbVnd
         lblNme = "" & !VEBNAME
         lblPdt = "" & Format(!PODATE, "mm/dd/yyyy")
         If !POSERVICE Then optSrv.Value = 1 Else optSrv.Value = vbUnchecked
         cmbTyp = "" & Trim(!POTYPE)
         txtSdt = "" & Format(!POSHIP, "mm/dd/yyyy")
         cmbDiv = "" & Trim(!PODIVISION)
         cmbByr = "" & Trim(!POBUYER)
         txtCnt = "" & Trim(!POBCONTACT)
         txtPhone = "" & Trim(!VEBPHONE)
         txtFob = "" & Trim(!POFOB)
         cmbTrm = "" & Trim(!POSTERMS)
         txtShp = "" & Trim(!POSHIPTO)
         txtVia = "" & Trim(!POVIA)
         txtCmt = "" & !POREMARKS
         If !POTAXABLE Then cbTaxable.Value = 1 Else cbTaxable.Value = vbUnchecked
         lblPrn = "" & Format(!POPRINTED, "mm/dd/yy")
         If txtSdt = "" Then txtSdt = Format(ES_SYSDATE, "mm/dd/yyyy")
         txtPdt = lblPdt
         cmbReq = "" & Trim(!POREQBY)
      End With
      If Len(Trim(cmbByr)) > 0 Then GetCurrentBuyer cmbByr
      GetPurchaseOrder = 1
      cmdTrm.Enabled = True
      cmdItm.Enabled = True
   Else
      If bMessage Then
         If Val(cmbPon) > 0 Then MsgBox "Purchase Order " & cmbPon & " Was     " & vbCr _
                & "Canceled Or Doesn't Exist.", _
                vbExclamation, Caption
         If cmbPon.ListCount > 0 Then cmbPon = cmbPon.List(0)
      End If
      GetPurchaseOrder = 0
      cmdTrm.Enabled = False
      cmdItm.Enabled = False
      On Error Resume Next
      cmbPon.SetFocus
   End If
   If GetPurchaseOrder = 1 Then
      'RdoRcd.RowsetSize = 1
      'RdoRcd(0) = Val(cmbPon)
      RdoRcd.Parameters(0).Value = cmbPon
      bSqlRows = clsADOCon.GetQuerySet(RdoVnd, RdoRcd, ES_KEYSET, True, 1)
      If bSqlRows Then
         'z1(2).Enabled = False
         cmbVnd.Enabled = False
      Else
         'z1(2).Enabled = True
         cmbVnd.Enabled = True
      End If
   End If
   bDataHasChanged = False
   MouseCursor 0
   Set RdoVnd = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpurcha"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function



Private Sub lblByr_Change()
   If Left(lblByr, 8) = "*** Buye" Then
      lblByr.ForeColor = ES_RED
   Else
      lblByr.ForeColor = Es_TextForeColor
   End If
   
End Sub

Private Sub optItm_Click()
   'never visible-used to alert that PurcPRe02b is loaded
   
End Sub

Private Sub optNew_Click()
   'never visible-used to PurcPRe02b of a new po
   
End Sub


Private Sub optPrn_Click()
   If Val(cmbPon) > 0 Then
      WindowState = vbMinimized
      bPrint = 1
      PurcPRp01a.lblPO = cmbPon
      PurcPRp01a.cboEndPo = cmbPon
      PurcPRp01a.Show
      Unload Me
   End If
   
End Sub

Private Sub optPvw_Click()
   If optPvw.Value = vbChecked Then
      cmbPon_Click
      optPvw.Value = vbUnchecked
      On Error Resume Next
      cmbPon.SetFocus
   End If
   
End Sub

Private Sub optSrv_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub



Private Sub tab1_Click()
   On Error Resume Next
   If tab1.SelectedItem.Index = 1 Then
      tabFrame(0).Visible = True
      tabFrame(1).Visible = False
      cmbTyp.SetFocus
   Else
      tabFrame(1).Visible = True
      tabFrame(0).Visible = False
      cmbTrm.SetFocus
   End If
   
End Sub

Private Sub txtCmt_LostFocus()
   ' Change the limit tio 6000
   'txtCmt = CheckLen(txtCmt, 2048)
   txtCmt = CheckLen(txtCmt, 6000)
   If bGoodPo Then
      On Error Resume Next
      RdoPso!POREMARKS = "" & txtCmt
      RdoPso.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtCnt_LostFocus()
   txtCnt = CheckLen(txtCnt, 20)
   txtCnt = StrCase(txtCnt)
   If bGoodPo Then
      On Error Resume Next
      RdoPso!POBCONTACT = "" & txtCnt
      RdoPso.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtFob_LostFocus()
   txtFob = CheckLen(txtFob, 12)
   If bGoodPo Then
      On Error Resume Next
      RdoPso!POFOB = "" & txtFob
      RdoPso.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtPdt_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtPdt_LostFocus()
   txtPdt = CheckLen(txtPdt, 10)
   If bGoodPo Then
      On Error Resume Next
      RdoPso!PODATE = "" & txtPdt
      RdoPso.Update
      lblPdt = txtPdt
      If Err > 0 Then ValidateEdit
   End If
   
End Sub




Private Sub txtSdt_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtSdt_LostFocus()
   txtSdt = CheckLen(txtSdt, 10)
   If bGoodPo Then
      On Error Resume Next
      If Trim(txtShp) <> "" Then
         RdoPso!POSHIP = txtSdt
      Else
         RdoPso!POSHIP = Null
      End If
      RdoPso.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtShp_LostFocus()
   txtShp = CheckLen(txtShp, 255)
   If bGoodPo Then
      On Error Resume Next
      RdoPso!POSHIPTO = "" & txtShp
      RdoPso.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub ClearBoxes()
   On Error Resume Next
   'tab1.Tab = 0
   'tab1.Enabled = False
   txtRel = ""
   lblNme = ""
   lblPdt = ""
   cmbTyp = ""
   txtSdt = Format(ES_SYSDATE, "mm/dd/yyyy")
   cmbByr = ""
   txtCnt = ""
   txtFob = ""
   cmbTrm = ""
   txtShp = ""
   txtCmt = ""
   
End Sub


Private Sub txtVia_DropDown()
   bDataHasChanged = True
   
End Sub

Private Sub txtVia_LostFocus()
   txtVia = CheckLen(txtVia, 24)
   txtVia = StrCase(txtVia)
   If bGoodPo Then
      On Error Resume Next
      RdoPso!POVIA = "" & txtVia
      RdoPso.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub



Private Sub FillVia()
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT POVIA FROM PohdTable WHERE " _
          & "POVIA<>''"
   LoadComboBox txtVia, -1
   Exit Sub
   
DiaErr1:
   sProcName = "fillvia"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub UpdatePoitVendor(ByVal strVendor As String, ByVal strPoNum As String)
   On Error GoTo DiaErr1
   
   sSql = "UPDATE PoitTable SET PIVENDOR='" & Compress(strVendor) _
            & "' WHERE PINUMBER=" & strPoNum
   clsADOCon.ExecuteSQL sSql

   Exit Sub
   
DiaErr1:
   sProcName = "UpdatePoitVendor"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub
Private Sub UpdatePurchaseOrder()
   On Error Resume Next
   With RdoPso

      If Trim(txtSdt) <> "" Then
         !POSHIP = Format(txtSdt, "mm/dd/yyyy")
      Else
         !POSHIP = Null
      End If
      !POVENDOR = Compress(cmbVnd)
      !PODIVISION = "" & cmbDiv
      !POBUYER = "" & Compress(cmbByr)
      !POREQBY = Trim(cmbReq)
      !POVIA = "" & txtVia
      !POTYPE = "" & Trim(cmbTyp)
      .Update
      .Close
   End With
   bDataHasChanged = False
   
End Sub


Public Sub RefreshPoCursor()
   'RdoPon.RowsetSize = 1
   'RdoPon(0) = Val(cmbPon)
   RdoPon.Parameters(0).Value = Val(cmbPon)
   bSqlRows = clsADOCon.GetQuerySet(RdoPso, RdoPon, ES_KEYSET, 1)
   
End Sub



Private Sub LoadVendorAddress()
    Dim RdoVndr As ADODB.Recordset
    
    Dim sAddress As String
    sAddress = ""
    Call HookToolTips
    sSql = "SELECT VEBADR, VEBCITY, VEBSTATE, VEBZIP,VEBCONTACT, VEBPHONE, VEBEXT FROM VndrTable WHERE VEREF = '" & Compress(cmbVnd) & "'"
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoVndr, ES_FORWARD)
    If bSqlRows Then
        If Len(Trim("" & RdoVndr!VEBCONTACT)) > 0 Then sAddress = sAddress & Trim(RdoVndr!VEBCONTACT) & vbCrLf
        If Len(Trim("" & RdoVndr!VEBADR)) > 0 Then sAddress = sAddress & Trim(RdoVndr!VEBADR) & vbCrLf
        If Len(Trim("" & RdoVndr!VEBCITY)) > 0 Then sAddress = sAddress & Trim(RdoVndr!VEBCITY) & " ," & Trim(RdoVndr!VEBSTATE) & " " & Trim(RdoVndr!VEBZIP) & vbCrLf
        If Len(Trim("" & RdoVndr!VEBPHONE)) > 0 And Trim("" & RdoVndr!VEBPHONE) <> "___-___-____" Then sAddress = sAddress & "" & Trim(RdoVndr!VEBPHONE)
        If Len(Trim("" & Trim(RdoVndr!VEBEXT))) > 0 And Val("" & Trim(RdoVndr!VEBEXT)) > 0 Then sAddress = sAddress & " Ext: " & Trim(RdoVndr!VEBEXT)
    End If
    Set RdoVndr = Nothing
    If Len(sAddress) > 0 Then
           cmbVnd.ToolTipText = sAddress
        'FusionToolTip.ToolText(cmbVnd) = sAddress
    Else
        cmbVnd.ToolTipText = "Enter the Vendor"
    End If

End Sub

Private Sub GetVendorTerms(sVendRef As String)
   Dim RdoTrm As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "SELECT VEREF,VENETDAYS,VEDDAYS,VEDISCOUNT," _
          & "VEPROXDT,VEPROXDUE,VEFOB,VEBNAME,VEBADR,VEBCITY," _
          & "VEBSTATE,VEBZIP,VEBCONTACT,VEBUYER FROM VndrTable WHERE " _
          & "VEREF='" & sVendRef & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTrm)
   If bSqlRows Then
      With RdoTrm
         On Error Resume Next
         cNetDays = 0 + !VENETDAYS
         cDDays = 0 + !VEDDAYS
         cDiscount = 0 + !VEDISCOUNT
         cProxDt = 0 + !VEPROXDT
         cProxdue = 0 + !VEPROXDUE
         sFob = "" & Trim(!VEFOB)
         sBcontact = "" & Trim(!VEBCONTACT)
      End With
   End If
   Set RdoTrm = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getvendorterms"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

