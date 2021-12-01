VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form SaleSLe02b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sales Order Items"
   ClientHeight    =   5145
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   9105
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   2170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrt 
      Height          =   315
      Left            =   6960
      Picture         =   "SaleSLe02b.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   63
      TabStop         =   0   'False
      ToolTipText     =   "New Part Numbers"
      Top             =   4560
      Width           =   350
   End
   Begin VB.CheckBox OptIgnQ 
      Caption         =   "&Ignore Qty "
      Height          =   255
      Left            =   4320
      TabIndex        =   62
      TabStop         =   0   'False
      ToolTipText     =   "Carry Part Number,Quantity And Price Forward"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdStatCode 
      DisabledPicture =   "SaleSLe02b.frx":049B
      DownPicture     =   "SaleSLe02b.frx":0E0D
      Height          =   350
      Left            =   6960
      MaskColor       =   &H8000000F&
      Picture         =   "SaleSLe02b.frx":177F
      Style           =   1  'Graphical
      TabIndex        =   61
      TabStop         =   0   'False
      ToolTipText     =   "Add Internal Status Code"
      Top             =   4140
      UseMaskColor    =   -1  'True
      Width           =   350
   End
   Begin VB.CommandButton cmdStatCd 
      DownPicture     =   "SaleSLe02b.frx":1C0E
      Height          =   350
      Left            =   6000
      Picture         =   "SaleSLe02b.frx":20E8
      Style           =   1  'Graphical
      TabIndex        =   60
      TabStop         =   0   'False
      ToolTipText     =   "Add Status Code for each Item"
      Top             =   5640
      Width           =   350
   End
   Begin VB.CommandButton cmdCom 
      DisabledPicture =   "SaleSLe02b.frx":25C2
      DownPicture     =   "SaleSLe02b.frx":2A04
      Enabled         =   0   'False
      Height          =   350
      Left            =   6960
      Picture         =   "SaleSLe02b.frx":2E46
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Commissions"
      Top             =   3780
      Width           =   350
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLe02b.frx":3288
      Style           =   1  'Graphical
      TabIndex        =   52
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtPrt 
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Tag             =   "3"
      Top             =   2340
      Width           =   3075
   End
   Begin VB.CheckBox optCreated 
      Caption         =   "Manfacturing Order Created"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   18
      ToolTipText     =   "If Checked, Ingnores Freight Days"
      Top             =   4860
      Visible         =   0   'False
      Width           =   2412
   End
   Begin VB.CommandButton cmdCmo 
      DownPicture     =   "SaleSLe02b.frx":3A36
      Height          =   350
      Left            =   6000
      Picture         =   "SaleSLe02b.frx":4248
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Create A Manufacturing Order For This Item"
      Top             =   4140
      Width           =   350
   End
   Begin VB.CheckBox optFreight 
      Height          =   255
      Left            =   1560
      TabIndex        =   17
      ToolTipText     =   "If Checked, Ingnores Freight Days"
      Top             =   4860
      Width           =   2172
   End
   Begin VB.ComboBox txtBook 
      Height          =   315
      Left            =   7920
      TabIndex        =   13
      Tag             =   "4"
      ToolTipText     =   "This Item Was Booked"
      Top             =   3420
      Width           =   1215
   End
   Begin VB.CommandButton cmdNxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8400
      Picture         =   "SaleSLe02b.frx":46D9
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Next Entry"
      Top             =   4200
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdLst 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7860
      Picture         =   "SaleSLe02b.frx":480F
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "Previous Entry"
      Top             =   4200
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CheckBox optVew 
      Caption         =   "View"
      Height          =   255
      Left            =   3960
      TabIndex        =   49
      Top             =   6000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "SaleSLe02b.frx":4945
      Height          =   315
      Left            =   4440
      Picture         =   "SaleSLe02b.frx":4C87
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   2340
      Width           =   350
   End
   Begin VB.TextBox txtCin 
      Height          =   285
      Left            =   3660
      TabIndex        =   15
      ToolTipText     =   "Customer's Item Number"
      Top             =   3420
      Width           =   735
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "SaleSLe02b.frx":4FC9
      DownPicture     =   "SaleSLe02b.frx":593B
      Height          =   350
      Left            =   6000
      Picture         =   "SaleSLe02b.frx":62AD
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Standard Comments"
      Top             =   3780
      Width           =   350
   End
   Begin VB.CheckBox optBok 
      Caption         =   "PriceBook"
      Height          =   255
      Left            =   360
      TabIndex        =   47
      Top             =   6000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdBok 
      DownPicture     =   "SaleSLe02b.frx":6705
      Height          =   350
      Left            =   6560
      Picture         =   "SaleSLe02b.frx":6BDF
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Show Price Book"
      Top             =   3780
      Width           =   350
   End
   Begin VB.CheckBox optInv 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6600
      TabIndex        =   3
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox optShp 
      Enabled         =   0   'False
      Height          =   255
      Left            =   5400
      TabIndex        =   2
      Top             =   1560
      Width           =   255
   End
   Begin VB.ComboBox txtDdt 
      Height          =   315
      Left            =   7920
      TabIndex        =   12
      Tag             =   "4"
      ToolTipText     =   "Expected Actual Delivery (With Freight Days)"
      Top             =   3060
      Width           =   1215
   End
   Begin VB.ComboBox txtRdt 
      Height          =   315
      Left            =   7920
      TabIndex        =   11
      Tag             =   "4"
      ToolTipText     =   "Customer Requested Date"
      Top             =   2700
      Width           =   1215
   End
   Begin VB.ComboBox txtSdt 
      Height          =   315
      Left            =   7920
      TabIndex        =   10
      Tag             =   "4"
      ToolTipText     =   "Expected Ship Date"
      Top             =   2340
      Width           =   1215
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "SaleSLe02b.frx":7012
      Height          =   350
      Left            =   6560
      Picture         =   "SaleSLe02b.frx":74EC
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Show Sales Order Items"
      Top             =   4140
      Width           =   350
   End
   Begin VB.CheckBox optRep 
      Caption         =   "&Repeat Sales Order Items"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Carry Part Number,Quantity And Price Forward"
      Top             =   1560
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.TextBox txtCmt 
      Height          =   975
      Left            =   1380
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Tag             =   "9"
      ToolTipText     =   "Comments (3072 Chars Max)"
      Top             =   3840
      Width           =   4335
   End
   Begin VB.TextBox txtDiscountPercent 
      Height          =   285
      Left            =   6000
      TabIndex        =   9
      Tag             =   "1"
      ToolTipText     =   "Discount Percentage"
      Top             =   2340
      Width           =   735
   End
   Begin VB.CheckBox optCom 
      Caption         =   "____"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   3120
      TabIndex        =   29
      Top             =   6120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdTrm 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   7980
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Cancel The Current PO Item"
      Top             =   900
      Width           =   915
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&New"
      Height          =   315
      Left            =   7980
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Add This Sales Order Item"
      Top             =   540
      Width           =   915
   End
   Begin VB.TextBox txtFrt 
      Height          =   285
      Left            =   1380
      TabIndex        =   14
      Tag             =   "1"
      ToolTipText     =   "Adder For Frieght"
      Top             =   3420
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   180
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "Enter Quantity"
      Top             =   2340
      Width           =   1095
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   6
      Tag             =   "3"
      ToolTipText     =   "Select From List"
      Top             =   2340
      Width           =   3075
   End
   Begin VB.TextBox txtListPrice 
      Height          =   285
      Left            =   4860
      TabIndex        =   8
      Tag             =   "1"
      ToolTipText     =   "Unit Price"
      Top             =   2340
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   7980
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   60
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6960
      Top             =   5760
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5145
      FormDesignWidth =   9105
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   1455
      Left            =   480
      TabIndex        =   0
      ToolTipText     =   "Click To Select Or Scroll And Press Enter (Also Page Up And Page Down)"
      Top             =   0
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   2566
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      FocusRect       =   2
      ScrollBars      =   2
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Internal Comments:"
      Height          =   495
      Index           =   10
      Left            =   240
      TabIndex        =   59
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label lblRow 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   8100
      TabIndex        =   58
      ToolTipText     =   "Our Item Number"
      Top             =   1500
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblFrd 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7620
      TabIndex        =   57
      ToolTipText     =   "Discount Days - not visible"
      Top             =   4980
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblDiscountAmount 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4860
      TabIndex        =   56
      ToolTipText     =   "Quantity On Hand"
      Top             =   2700
      Width           =   1095
   End
   Begin VB.Label lblNetPrice 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6780
      TabIndex        =   55
      ToolTipText     =   "Quantity On Hand"
      Top             =   2340
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Disc %"
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
      Index           =   17
      Left            =   6000
      TabIndex        =   54
      Top             =   2100
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price                 "
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
      Index           =   16
      Left            =   6840
      TabIndex        =   53
      Top             =   2100
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Booking Date"
      Height          =   255
      Index           =   0
      Left            =   6840
      TabIndex        =   51
      Top             =   3480
      Width           =   1050
   End
   Begin VB.Label lblQoh 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   180
      TabIndex        =   50
      ToolTipText     =   "Quantity On Hand"
      Top             =   2700
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Item"
      Height          =   255
      Index           =   15
      Left            =   2580
      TabIndex        =   48
      ToolTipText     =   "Customer's Item Number"
      Top             =   3420
      Width           =   1095
   End
   Begin VB.Label lblIfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Shipped"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   6000
      TabIndex        =   46
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblIfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Invoiced"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   7200
      TabIndex        =   45
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblPsNum 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   60
      TabIndex        =   44
      Top             =   4260
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments:"
      Height          =   255
      Index           =   13
      Left            =   180
      TabIndex        =   43
      Top             =   3780
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Commission?:"
      Height          =   255
      Index           =   9
      Left            =   1560
      TabIndex        =   42
      Top             =   6120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblSon 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   60
      TabIndex        =   41
      Top             =   4020
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Freight Allow:"
      Height          =   255
      Index           =   8
      Left            =   180
      TabIndex        =   40
      ToolTipText     =   "Adder For Frieght"
      Top             =   3420
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Date"
      Height          =   255
      Index           =   7
      Left            =   6840
      TabIndex        =   39
      Top             =   3120
      Width           =   1050
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Request Date"
      Height          =   255
      Index           =   6
      Left            =   6840
      TabIndex        =   38
      Top             =   2760
      Width           =   1050
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ship Date        "
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
      Index           =   5
      Left            =   7920
      TabIndex        =   37
      Top             =   2100
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   36
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   765
      TabIndex        =   35
      ToolTipText     =   "Our Item Number"
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity/Qoh            "
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
      Left            =   180
      TabIndex        =   34
      Top             =   2100
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number                                                           "
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
      Left            =   1380
      TabIndex        =   33
      Top             =   2100
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "List Price/Disc"
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
      Index           =   4
      Left            =   4860
      TabIndex        =   32
      Top             =   2100
      Width           =   1095
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   1380
      TabIndex        =   31
      ToolTipText     =   "Extended Description"
      Top             =   2700
      Width           =   3015
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   30
      ToolTipText     =   "Item Revision"
      Top             =   1560
      Width           =   495
   End
End
Attribute VB_Name = "SaleSLe02b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'3/26/02 Added Invoiced/shipped checks and ToolTipText
'9/27/03 Removed Commissions
'4/9/04 Grid
'5/13/04 added keyset cursor to columns
'7/8/04 Smoothed grid scrolling
'8/12/04 Added Book Date
'9/14/04 Added date trap for Ship on new item
'9/15/04 Changed to canceling rather than deleting an Item
'12/29/04 Provided for Nulls in UpdateItem and GetThisItem
'3/7/05 Corrected Grd.Rows
'4/13/05 Added CheckFreightDays and override switch
'12/03/05 Fixed Cust Requested Date (ITCUSTREQ)
'12/22/05 Added Create MO features (Production.ShopSHe01a)
'1/23/05 Added txtPrt for Part Count > 5000 (INTCOA)
'1/27/05 Enabled cmdCmo where disabled
'10/6/06 Added Run Allocations to Cancel An Item
'10/27/06 Fixed Cancel RnalTable Query.
'12/14/06 Unblocked the Cancel SO Item Query 7.2.5
Option Explicit
'Dim rdoQry As rdoQuery
Dim cmdObj As ADODB.Command
Dim RdoItm As ADODB.Recordset

Dim bOnLoad As Byte
Dim bAllowMo As Byte
Dim bFreightDays As Byte
Dim bGoodItem As Byte
Dim bGoodPart As Byte
Dim bNewItem As Byte
Dim bGoodBook As Byte
Dim bBookDated As Byte
Dim bView As Byte

Dim iIndex As Integer
Dim iTotalItems As Integer
Dim iLastitem As Integer

Dim lSalesOrder As Long

Dim cOldPrice As Currency
Dim cOldQty As Currency

Dim sBook As String
Dim sOldPart As String
Dim sPassedPart As String
Dim vOldShip As Variant
Dim vOldDel As Variant
Dim DefaultDiscountForCustomer As Currency

'grid columns
Private Const COL_ItemNo = 0
Private Const COL_ItemRev = 1
Private Const COL_Qty = 2
Private Const COL_PartNo = 3
Private Const COL_UnitPrice = 4
Private Const COL_Count = 5

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Function GetSellingPriceFormat() As String
   Dim RdoSpf As ADODB.Recordset
   On Error GoTo 0
   sSql = "SELECT SellingPriceFormat FROM Preferences WHERE PreRecord=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSpf, ES_FORWARD)
   If bSqlRows Then GetSellingPriceFormat = RdoSpf!SellingPriceFormat _
                                            Else GetSellingPriceFormat = "#######0.000"
   Set RdoSpf = Nothing
   Exit Function
modErr1:
   GetSellingPriceFormat = "#######0.000"
   
End Function

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtQty = "0.000"
   txtDdt.ToolTipText = "Includes Default Freight Days"
   
End Sub

Private Sub cmbPrt_Click()
   If txtPrt.Visible Then cmbPrt = txtPrt
   bGoodPart = GetPart(cmbPrt)
   Grd.Enabled = False
   If bGoodPart Then
      Grd.Col = COL_Qty
      Grd.Text = cmbPrt
      Grd.Col = COL_ItemNo
   End If
   
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   'On Error Resume Next
   If bView = 1 Then Exit Sub
   
   If (Not ValidPartNumber(cmbPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If
   
   bGoodPart = GetPart(cmbPrt)
   txtPrt = cmbPrt
   If bGoodPart = 1 Then
      GetBookPrice
      If Val(txtQty) > 0 Then
         cmdAdd.Enabled = True
         cmdTrm.Enabled = True
         Grd.Row = iIndex
         Grd.Col = COL_Qty
         Grd.Text = cmbPrt
         Grd.Col = COL_ItemNo
         If Len(Trim(cmbPrt)) > 0 And Val(txtQty) > 0 Then UpdateItem True, False
      End If
   End If
   Grd.Enabled = True
   
End Sub

Private Sub cmdAdd_Click()
    UpdateItem
    AddSoItem
   
End Sub

Private Sub cmdBok_Click()
   If bBookDated Then
      MsgBox "The Price Book Assigned Is Out Of Date.", _
         vbInformation, Caption
      Exit Sub
   End If
   If bGoodBook Then
      optBok.Value = vbChecked
      SaleSLe02c.Show
   Else
      MsgBox "There Isn't A Price Book Assigned For This Customer.", _
         vbInformation, Caption
   End If
   
End Sub

Private Sub cmdCan_Click()
   If bNewItem = 0 Then
      If Val(txtQty) > 0 Then UpdateItem
   End If
   SaleSLe02a.Enabled = True
   Unload Me
   
End Sub

Private Sub cmdCmo_Click()
   Dim bResponse As Byte
   bResponse = MsgBox("This Function Allows Creation And Allocation Of An MO." & vbCrLf _
               & "Once Created, You Cannot Change The Part Number For" & vbCrLf _
               & "The Item. Do You Wish To Continue To Create The MO?", _
               ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      cmdCmo.Enabled = False
      cUR.CurrentPart = cmbPrt
      ShopSHe01a.z1(11).Visible = True
      ShopSHe01a.optPrn.Visible = True
      ShopSHe01a.optRel = vbChecked
      ShopSHe01a.optAllocate.Visible = True
      ShopSHe01a.optAllocate = vbChecked
      ShopSHe01a.cmdAdd.Enabled = True
      ShopSHe01a.txtCmp = txtSdt
      ShopSHe01a.txtRun.Enabled = False
      ShopSHe01a.txtPrt = sPassedPart
      
      ShopSHe01a.lblSalesOrder = lblSon
      ShopSHe01a.lblSoItem = lblItm
      ShopSHe01a.lblSoItemRev = lblRev
      ShopSHe01a.cmbPrt = sPassedPart
      ShopSHe01a.GetNextRun
      ShopSHe01a.txtQty = txtQty
      ShopSHe01a.txtQty.Enabled = True

      ShopSHe01a.optDmy = vbChecked
      ShopSHe01a.Show
   Else
      CancelTrans
   End If
   
End Sub

Private Sub cmdCom_Click()
   If cmdCom Then
      CommCOe02a.optFrom.Value = vbChecked
      CommCOe02a.cmbSon = Format(lblSon, SO_NUM_FORMAT)
      CommCOe02a.bGoodSO = 1
      CommCOe02a.cmdSel.Value = True
      DoEvents
      CommCOe02a.cmbItm = Me.lblItm
      CommCOe02a.optSoItems.Value = vbChecked
      CommCOe02a.GetThisItem
      CommCOe02a.Show
      cmdCom = False
   End If
   
End Sub

Private Sub cmdComments_Click()
   If cmdComments Then
      'See List For Index
      txtCmt.SetFocus
      SysComments.lblListIndex = 3
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
   bView = 0
   
End Sub


Private Sub cmdFnd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bView = 1
   
End Sub



Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2170
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdLst_Click()
   Dim iGridRow As Integer
   UpdateItem
   If Grd.Row > 1 Then
      Grd.Row = Grd.Row - 1
   Else
      Grd.Row = 1
   End If
   Grd.Col = COL_ItemNo
   lblItm = Grd.Text
   Grd.Col = COL_ItemRev
   lblRev = Grd.Text
   iIndex = Grd.Row
   lblRow = iIndex
   bGoodItem = GetThisItem()
   If Grd.Row < 6 Then
      Grd.TopRow = 1
   Else
      iGridRow = Grd.Row Mod 4
      If iGridRow = 2 Then Grd.TopRow = Grd.Row - 4
   End If
   
End Sub

Private Sub cmdNxt_Click()
   Dim iGridRow As Integer
   UpdateItem
   If Grd.Row >= Grd.Rows - 1 Then
      Grd.Row = Grd.Rows - 1
   Else
      Grd.Row = Grd.Row + 1
   End If
   Grd.Col = COL_ItemNo
   lblItm = Grd.Text
   Grd.Col = COL_ItemRev
   lblRev = Grd.Text
   iIndex = Grd.Row
   lblRow = iIndex
   bGoodItem = GetThisItem()
   If Grd.Row < 6 Then
      Grd.TopRow = 1
   Else
      iGridRow = Grd.Row Mod 4
      If iGridRow = 2 Then Grd.TopRow = Grd.Row - 1
   End If
   
End Sub

Private Sub cmdPrt_Click()
   InvcINe01a.Show
   'cmdPrt.Value = False
   
End Sub

Private Sub cmdStatCode_Click()
    StatusCode.lblSCTypeRef = "SO Number"
    StatusCode.txtSCTRef = lblSon
    StatusCode.lblSCTRef1 = lblItm
    StatusCode.lblSCTRef2 = lblRev
    StatusCode.lblStatType = "SOI"
    StatusCode.lblSysCommIndex = 3 ' The index in the Sys Comment "SO"
    StatusCode.txtCurUser = cUR.CurrentUser
    
    StatusCode.Show

End Sub

Private Sub cmdTrm_Click()
   Dim bResponse As Integer
   Dim sMsg As String
   If Val(lblItm) = 0 Then Exit Sub
   'On Error Resume Next
   If optInv.Value = vbChecked Or optShp.Value = vbChecked Then
      MsgBox "Item Either Is Shipped Or Invoiced And Cannot Be Canceled.", _
         vbInformation, Caption
      Exit Sub
   End If
   
   If (Grd.Rows <= 2 Or iTotalItems <= 1) Then
        sMsg = "Can't delete the last item on a sales order."
        MsgBox sMsg, vbInformation, Caption
        Exit Sub
        
    End If
   sMsg = "Are You Sure That You Wish To Cancel " & vbCrLf _
          & "Item " & lblItm & "."
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      
      clsADOCon.ADOErrNum = 0
      
      sSql = "UPDATE SoitTable SET ITACTUAL=NULL,ITCANCELED=1,ITCANCELDATE='" _
             & Format(ES_SYSDATE, "mm/dd/yyyy") & "',ITUSER='" _
             & sInitials & "' WHERE (ITSO=" & lSalesOrder & " AND " _
             & "ITNUMBER=" & lblItm & " AND ITREV='" & lblRev & "')"
     clsADOCon.ExecuteSql sSql 'rdExecDirect
      
      'Allocations
      sSql = "DELETE FROM RnalTable WHERE (RASO=" & lSalesOrder & " " _
             & "AND RASOITEM=" & lblItm & " AND RASOREV='" & lblRev & "')"
     clsADOCon.ExecuteSql sSql 'rdExecDirect
      
      ' Delete Sales commission record
      sSql = "delete from SpcoTable" & vbCrLf _
              & "where SMCOSO = " & lSalesOrder & vbCrLf _
              & "and SMCOSOIT = " & lblItm & " and SMCOITREV = '" & lblRev & "'"
     clsADOCon.ExecuteSql sSql
      
      
      If clsADOCon.ADOErrNum = 0 Then
         MsgBox "Item Canceled.", vbInformation, Caption
         Grd.Row = 1
         iIndex = Grd.Row
         lblRow = iIndex
         GetItems
      Else
         MsgBox "Couldn't Cancel Item.", vbInformation, Caption
      End If
   Else
      CancelTrans
   End If
   
End Sub

Private Sub cmdVew_Click()
   Dim RdoVew As ADODB.Recordset
   Dim iList As Integer
   Dim iCol As Integer
   Dim iRows As Integer
   
   Dim cQuantity As Currency
   Dim cSoTotal As Currency
   
   UpdateItem
   'On Error Resume Next
   iRows = 10
   With SaleSLview.Grd
      .Rows = iRows
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      If Screen.Width > 9999 Then
         .ColWidth(0) = 350 * 1.25
         .ColWidth(1) = 1450 * 1.25
         .ColWidth(2) = 600 * 1.25
         .ColWidth(3) = 700 * 1.25
         .ColWidth(4) = 700 * 1.25
         .ColWidth(5) = 600 * 1.25
      Else
         .ColWidth(0) = 350
         .ColWidth(1) = 1450
         .ColWidth(2) = 600
         .ColWidth(3) = 700
         .ColWidth(4) = 700
         .ColWidth(4) = 600
      End If
   End With
   sSql = "SELECT ITSO,ITNUMBER,ITREV,ITSCHED,ITACTUAL,ITPART,ITQTY,ITDOLLARS," _
          & "ITCANCELED,PARTREF,PARTNUM FROM SoitTable,PartTable WHERE (ITPART=PARTREF) " _
          & "AND ITSO=" & Val(lblSon) & " ORDER BY ITNUMBER,ITREV"
   MouseCursor 13
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVew, ES_FORWARD)
   If bSqlRows Then
      With RdoVew
         Do Until .EOF
            iList = iList + 1
            cQuantity = cQuantity + !ITQty
            cSoTotal = cSoTotal + (!ITQty * !ITDOLLARS)
            iRows = iRows + 1
            SaleSLview.Grd.Rows = iRows
            SaleSLview.Grd.Row = iRows - 10
            
            SaleSLview.Grd.Col = 0
            SaleSLview.Grd.Text = "" & str(!ITNUMBER) & Trim(!itrev)
            
            SaleSLview.Grd.Col = 1
            SaleSLview.Grd.Text = "" & Trim(!PartNum)
            
            SaleSLview.Grd.Col = 2
            SaleSLview.Grd.Text = Format(!ITQty, ES_QuantityDataFormat)
            
            SaleSLview.Grd.Col = 3
            SaleSLview.Grd.Text = Format(!itsched, "mm/dd/yyyy")
            
            SaleSLview.Grd.Col = 4
            SaleSLview.Grd.Text = Format(!ITACTUAL, "mm/dd/yyyy")
            
            SaleSLview.Grd.Col = 5
            If !ITCANCELED = 1 Then
               Set SaleSLview.Grd.CellPicture = SaleSLview.Chkyes.Picture
            Else
               Set Grd.CellPicture = SaleSLview.Chkno.Picture
            End If
            .MoveNext
         Loop
         ClearResultSet RdoVew
      End With
      If iList > 9 Then SaleSLview.Grd.Rows = iList + 1
      SaleSLview.lblQty = Format(cQuantity, ES_QuantityDataFormat)
      SaleSLview.lblTot = Format(cSoTotal, "#,###,##0.00")
   End If
   MouseCursor 0
   Set RdoVew = Nothing
   On Error GoTo 0
   SaleSLview.Show
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   Dim bPartSearch As Boolean
   If bGoodSoMo = 1 Then
      txtPrt.Enabled = False
      cmbPrt.Enabled = False
      cmdCmo.Visible = False
      optCreated.Value = vbChecked
   Else
      cmdCmo.Enabled = True
   End If
   bGoodSoMo = 0
   If bOnLoad Then
      
      bAllowMo = CheckAllocations()
      If bAllowMo = 1 Then optCreated.Visible = True
      bFreightDays = Val(SaleSLe02a.txtFrd)
      ES_SellingPriceFormat = GetSellingPriceFormat()
      optFreight.Caption = "Override" & str$(bFreightDays) & " Freight Days "
      If bFreightDays = 0 Then
         optFreight.Value = vbChecked
         optFreight.Enabled = False
      End If
      cmdComments.Enabled = True
      If Trim(SaleSLe02a.txtBook) <> "" Then
         sBook = Compress(SaleSLe02a.txtBook)
         bGoodBook = 1
      Else
         sBook = ""
         bGoodBook = 0
      End If
      bBookDated = TestBookDates()
     
     ' If ES_PARTCOUNT < 5001 Then FillParts
      ' MM 4/13
      bPartSearch = GetPartSearchOption
      SetPartSearchOption (bPartSearch)
      If (Not bPartSearch) Then FillParts
      
      bOnLoad = 0
      GetAllItems
      GetItems
   End If
   If optBok.Value = vbChecked Then
      Unload SaleSLe02c
      optBok.Value = vbUnchecked
   End If
   lblFrd = SaleSLe02a.txtDay
   MouseCursor 0
   
End Sub

Private Sub Form_Load()

   'get default discount for customer
   Dim rdo As ADODB.Recordset
   
   sSql = "SELECT CUDISCOUNT from CustTable where CUREF = '" & Compress(SaleSLe02a.cmbCst) & "'"
   If clsADOCon.GetDataSet(sSql, rdo) Then
      DefaultDiscountForCustomer = "0" & rdo.Fields(0)
   Else
      DefaultDiscountForCustomer = 0
   End If
   
   ShowIgnoreQtyOnSO
   
   SetFormSize Me
   Move SaleSLe02a.Left + 200, SaleSLe02a.Top + 1100
   FormatControls
   optRep.Value = GetSetting("Esi2000", "EsiSale", "csitmrep", optRep.Value)
   lblSon = Format(0 + Val(SaleSLe02a.cmbSon), "####0")
   bGoodSoMo = 0
   lSalesOrder = Val(lblSon)
   sSql = "SELECT ITSO,ITNUMBER,ITREV,ITQTY," _
          & "ITPART,ITDOLLARS,ITDOLLORIG,ITSCHED,ITCUSTREQ,ITSCHEDDEL," _
          & "ITFRTALLOW,ITFRTALLOW,ITDISCRATE,ITDISCAMOUNT,ITBOOKDATE,ITCOMMISSION," _
          & "ITPSNUMBER,ITPSITEM,ITCANCELED,ITINVOICE,ITCUSTITEMNO,ITMOCREATED," _
          & "ITCOMMENTS FROM SoitTable WHERE ITSO=" & Val(SaleSLe02a.cmbSon) & " " _
          & "AND (ITNUMBER= ? AND ITREV= ? ) AND ITCANCELED=0"
   'Set rdoQry = RdoCon.CreateQuery("", sSql)
   'rdoQry.MaxRows = 1
   Set cmdObj = New ADODB.Command
   cmdObj.CommandText = sSql
   
   Dim prmObj As ADODB.Parameter
   Dim prmObj1 As ADODB.Parameter
   
   Set prmObj = New ADODB.Parameter
   prmObj.Type = adInteger
   cmdObj.parameters.Append prmObj

   Set prmObj1 = New ADODB.Parameter
   prmObj1.Type = adChar
   prmObj1.Size = 3
   cmdObj.parameters.Append prmObj1
   
  ' If ES_PARTCOUNT > 5000 Then
     ' cmbPrt.Visible = False
     ' txtPrt.Visible = True
      'txtPrt.TabIndex = 3
      'cmbPrt.TabIndex = 99
      'cmdFnd.Visible = True
  ' End If
   With Grd
      .Cols = COL_Count
      .ColAlignment(COL_ItemNo) = 0
      .ColAlignment(COL_ItemRev) = 0
      .ColAlignment(COL_Qty) = 0
      .Row = 0
      .Col = COL_ItemNo
      .Text = "Item"
      .ColWidth(COL_ItemNo) = 850
      .Col = COL_ItemRev
      .Text = "Rev"
      .ColWidth(COL_ItemRev) = 400
      .Col = COL_Qty
      .Text = "Part Number"
      .ColWidth(COL_Qty) = 3300
      .Col = COL_PartNo
      .Text = "Quantity"
      .ColWidth(COL_PartNo) = 950
      .Col = COL_UnitPrice
      .Text = "Unit Price"
      .ColWidth(COL_UnitPrice) = 1200
      .Col = COL_ItemNo
   End With

   SaleSLe02a.Enabled = False
   bOnLoad = 1
   
End Sub


Private Sub Form_LostFocus()
   SaleSLe02a.Enabled = True
   
End Sub

'10/27/04

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'On Error Resume Next
   If bNewItem = 1 Then
      If Val(txtQty) = 0 Or bGoodPart = 0 Then
         sSql = "DELETE FROM SoitTable WHERE ITSO=" & Val(lblSon) & " " _
                & "AND ITNUMBER=" & Val(lblItm) & " AND ITREV='" & lblRev & "' "
        clsADOCon.ExecuteSql sSql 'rdExecDirect
      End If
   End If
   sSql = "DELETE FROM SoitTable WHERE ITPART='' AND ITSO=" & Val(lblSon)
  clsADOCon.ExecuteSql sSql 'rdExecDirect
   
   SaveSetting "Esi2000", "EsiSale", "csitmrep", optRep.Value
   SaleSLe02a.RefreshSoCursor 1
   SaleSLe02a.optItm = vbUnchecked
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   MdiSect.lblBotPanel = MdiSect.Caption
   'On Error Resume Next
   Set cmdObj = Nothing
   Set RdoItm = Nothing
   Set SaleSLe02b = Nothing
   
End Sub

Private Sub GetItems()
   Dim iList As Integer
   Dim iRows As Integer
   
   On Error GoTo DiaErr1
'   grd.Rows = 2
'   grd.Row = 1
'   For iList = 0 To 3
'      grd.Col = iList
'      grd.Text = ""
'   Next
   Grd.Rows = 1
   iTotalItems = 0
   MouseCursor 13
   
'   sSql = "SELECT ITSO,ITNUMBER,ITREV,ITSCHED,ITPART,ITQTY,ITPSNUMBER," _
'          & "ITCANCELED,PARTREF,PARTNUM,PACOMMISSION FROM SoitTable,PartTable " _
'          & "WHERE (PARTREF=ITPART) AND (ITSO=" & str(lSalesOrder) & " " _
'          & "AND ITCANCELED=0) ORDER BY ITNUMBER,ITREV"
   sSql = "select ITNUMBER,ITREV,ITQTY,ITDOLLARS," & vbCrLf _
      & "isnull(PARTNUM,'') as PARTNUM" & vbCrLf _
      & "FROM SoitTable" & vbCrLf _
      & "left join PartTable on PARTREF=ITPART" & vbCrLf _
      & "where ITSO=" & str(lSalesOrder) & " " & vbCrLf _
      & "and ITCANCELED=0" & vbCrLf _
      & "order by ITNUMBER,ITREV"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoItm, ES_FORWARD)
   If bSqlRows Then
      With RdoItm
         lblItm = "" & str(!ITNUMBER)
         lblRev = "" & !itrev
         Do Until .EOF
            iList = iList + 1
            iRows = iRows + 1
            'If iRows > 1 Then grd.Rows = grd.Rows + 1
            Grd.Rows = Grd.Rows + 1
            Grd.Row = Grd.Rows - 1
            Grd.Col = COL_ItemNo
            Grd.Text = Format(!ITNUMBER, "##0")
            Grd.Col = COL_ItemRev
            Grd.Text = "" & Trim(!itrev)
            Grd.Col = COL_Qty
            Grd.Text = "" & Trim(!PartNum)
            Grd.Col = COL_PartNo
            Grd.Text = Format(!ITQty, ES_QuantityDataFormat)
            'Grd.TextMatrix(Grd.Row, COL_UnitPrice) = Format(!ITDOLLARS, "######0.000")
            Grd.TextMatrix(Grd.Row, COL_UnitPrice) = Format(!ITDOLLARS, ES_SellingPriceFormat)
            .MoveNext
         Loop
         ClearResultSet RdoItm
      End With
      iTotalItems = iList
      If bNewItem = 0 Then Grd.Row = 1
      Grd.Col = COL_ItemNo
      lblItm = Grd.Text
      lblRow = 1
      Grd.Col = COL_ItemRev
      lblRev = Grd.Text
      
      bGoodItem = GetThisItem()
   Else
      'Per Larry...6/15/00
      '           If SaleSLe02a.optNew.Value = vbUnchecked Then
      '               MsgBox "No Open SO Items Were Found.", vbInformation, Caption
      '               SaleSLe02a.optNew = vbUnchecked
      '           End If
      MouseCursor 0
      If iLastitem = 0 Then
         iIndex = 1
         lblRow = iIndex
         txtDdt = SaleSLe02a.txtDdt
         txtSdt = SaleSLe02a.txtShp
         txtFrt = SaleSLe02a.txtFra
         'txtDiscountPercent = SaleSLe02a.txtDis  -- This is the payment terms discount
         txtDiscountPercent = DefaultDiscountForCustomer
         txtQty = "0.000"
         If txtPrt.Visible Then
            bGoodPart = GetPart(txtPrt)
         Else
            bGoodPart = GetPart(cmbPrt)
         End If
         AddSoItem
      End If
   End If
   If iIndex = 0 Then iIndex = 1
   'On Error Resume Next
   If Grd.Rows = 1 Then Grd.Rows = iIndex
   Grd.Row = iIndex
   Grd.Col = COL_ItemNo
   lblRow = iIndex
   Exit Sub
   
DiaErr1:
   sProcName = "getitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetThisItem() As Byte
   Dim a As Integer
   MouseCursor 13
   On Error GoTo DiaErr1
'   rdoQry.RowsetSize = 1
'   rdoQry(0) = Val(lblItm)
'   rdoQry(1) = lblRev
   
   cmdObj.parameters(0).Value = Val(lblItm)
   cmdObj.parameters(1).Value = lblRev
   
   bSqlRows = clsADOCon.GetQuerySet(RdoItm, cmdObj, ES_FORWARD, True)
   
 '  bSqlRows = GetQuerySet(RdoItm, rdoQry, ES_KEYSET)
   
   If bSqlRows Then
      With RdoItm
         'lblItm = "" & str(!ITNUMBER)    'redundant
         lblRev = "" & Trim(!itrev)
         txtQty = Format(!ITQty, ES_QuantityDataFormat)
         If txtPrt.Visible Then
            cmbPrt = "" & Trim(!ITPART)
            txtPrt = "" & Trim(!ITPART)
         Else
            cmbPrt = "" & Trim(!ITPART)
         End If
         'txtListPrice = Format(!ITDOLLARS, ES_SellingPriceFormat)
         txtListPrice = Format(!ITDOLLORIG, ES_SellingPriceFormat)
         lblNetPrice = Format(!ITDOLLARS, ES_SellingPriceFormat)
         txtDiscountPercent = Format(!ITDISCRATE, "#0.000")
         lblDiscountAmount = Format(!ITDISCAMOUNT, ES_SellingPriceFormat)
         
         cOldPrice = Val(txtListPrice)
         If Not IsNull(!itsched) Then
            txtSdt = "" & Format(!itsched, "mm/dd/yyyy")
         Else
            txtSdt = ""
         End If
         If Not IsNull(!itscheddel) Then
            txtDdt = "" & Format(!itscheddel, "mm/dd/yyyy")
         Else
            txtDdt = txtSdt
         End If
         '12/03/05 Was pointed to the wrong column
         If Not IsNull(!itcustreq) Then
            txtRdt = "" & Format(!itcustreq, "mm/dd/yyyy")
         Else
            txtRdt = txtSdt
         End If
         txtFrt = Format(!ITFRTALLOW, ES_QuantityDataFormat)
         
         'txtCin = Format(!ITCUSTITEMNO, "#####0")
         txtCin = "" & !ITCUSTITEMNO
         
         If !ITCOMMISSION Then optCom.Value = vbChecked Else optCom.Value = vbUnchecked
         txtCmt = "" & Trim(!ITCOMMENTS)
         If Right(txtCmt, 1) = vbLf And Right(txtCmt, 1) = vbCrLf Then
            If Len(txtCmt) > 1 Then
               a = Len(txtCmt)
               txtCmt = (Left$(txtCmt, a - 1))
            Else
               txtCmt = ""
            End If
         End If
         lblPsNum = "" & Trim(!ITPSNUMBER)
         If !ITINVOICE > 0 Then
            optInv.Value = vbChecked
            lblIfo(0).ToolTipText = "Invoice " & Format$(!ITINVOICE, "000000")
         Else
            optInv.Value = vbUnchecked
            lblIfo(0).ToolTipText = "Not Invoiced"
         End If
         If "" & Trim(!ITPSNUMBER) = "" Or optInv.Value = vbChecked Then
            lblIfo(1).ToolTipText = "No Pack Slip"
            ResetBoxes True
         Else
            lblIfo(1).ToolTipText = "Pack Slip " & Trim(!ITPSNUMBER)
            ResetBoxes False
         End If
         '8/12/04
         txtBook = Format(!ITBOOKDATE, "mm/dd/yyyy")
         If Trim(txtBook) = "" Then txtBook = Format(ES_SYSDATE, "mm/dd/yyyy")
         sOldPart = cmbPrt
         ClearResultSet RdoItm
         txtCmt.Enabled = True
         '12/22/05
         If bAllowMo = 1 Then
            cUR.CurrentPart = cmbPrt
            If Not IsNull(!ITMOCREATED) Then
               If !ITMOCREATED = 0 Then
                  optCreated.Value = vbUnchecked
                  cmdCmo.Enabled = True
                  cmdCmo.Visible = True
                  cmbPrt.Enabled = True
               Else
                  optCreated.Value = vbChecked
                  cmdCmo.Visible = False
                  cmbPrt.Enabled = False
               End If
            Else
               optCreated.Value = vbUnchecked
               cmdCmo.Enabled = True
               cmdCmo.Visible = True
               cmbPrt.Enabled = True
            End If
         End If
      End With
      vOldShip = txtSdt
      vOldDel = txtDdt
      'grd.Row = iIndex
      cOldQty = Val(txtQty)
      GetThisItem = 1
      'On Error Resume Next
      Grd.SetFocus
   Else
      sOldPart = ""
      GetThisItem = 0
   End If
   If GetThisItem = 1 Then
      bGoodPart = GetPart(cmbPrt)
      Grd.Col = COL_ItemRev
      Grd.Text = lblRev
      Grd.Col = COL_Qty
      Grd.Text = cmbPrt
      Grd.Col = COL_PartNo
      Grd.Text = txtQty
      Grd.TextMatrix(Grd.Row, COL_UnitPrice) = Format(Me.lblNetPrice, ES_SellingPriceFormat)
   End If
   Grd.Col = COL_ItemNo
   MouseCursor 0
   Exit Function
   
DiaErr1:
   sProcName = "getthisit"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub UpdateItem(Optional bUpdtComm As Boolean = False, Optional bUserMsg As Boolean = True)
   Dim sPartNumber As String
   Dim sComments As String
   
   Dim sSchedDte As String
   Dim sCustrDte As String
   Dim sSchelDte As String
   Dim sBookDte As String
   
   If txtPrt.Visible Then cmbPrt = txtPrt
   If txtSdt = "" Then sSchedDte = "Null" _
               Else sSchedDte = "'" & txtSdt & "'"
   
   If txtRdt = "" Then sCustrDte = "Null" _
               Else sCustrDte = "'" & txtRdt & "'"
   
   If txtDdt = "" Then sSchelDte = "Null" _
               Else sSchelDte = "'" & txtDdt & "'"
   
   If txtBook = "" Then sBookDte = "Null" _
                Else sBookDte = "'" & txtBook & "'"
   
   If (Trim(txtListPrice) = "") Then txtListPrice = "0.0000"
   
   If Val(txtQty) = 0 Then
      If Len(lblItm) > 0 Then
         MsgBox "Requires A Valid Quantity." & vbCrLf _
            & "Item Couldn't Be Updated.", vbExclamation, Caption
      End If
      Exit Sub
   End If
   If bGoodPart = 0 Then
      MsgBox "Requires A Valid Part Number." & vbCrLf _
         & "Item Couldn't Be Updated.", vbExclamation, Caption
      Exit Sub
   End If
   If Val(Trim(lblItm)) = 0 Then Exit Sub
   If Trim(txtBook) = "" Then txtBook = Format(ES_SYSDATE, "mm/dd/yyyy")
   sPartNumber = Compress(cmbPrt)
   If Len(Trim(txtDdt)) = 0 Then txtDdt = txtSdt
   sComments = Trim(txtCmt)
   If Len(sComments) > 0 Then sComments = ReplaceSingleQuote(sComments)
   
   CalculateDiscount
   
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   sSql = "UPDATE SoitTable" & vbCrLf _
      & "SET ITPART='" & sPartNumber & "'," & vbCrLf _
      & "ITQTY=" & CCur(txtQty) & "," & vbCrLf _
      & "ITDOLLORIG=" & CCur(txtListPrice) & "," & vbCrLf _
      & "ITDOLLARS=" & CCur(lblNetPrice) & "," & vbCrLf _
      & "ITDISCRATE=" & CCur(txtDiscountPercent) & "," & vbCrLf _
      & "ITDISCAMOUNT=" & CCur(lblDiscountAmount) & "," & vbCrLf _
      & "ITSCHED=" & sSchedDte & ",ITCUSTREQ=" & sCustrDte & "," & vbCrLf _
      & "ITSCHEDDEL=" & sSchelDte & ",ITFRTALLOW=" & Val(txtFrt) & "," & vbCrLf _
      & "ITCUSTITEMNO='" & txtCin & "'," & vbCrLf _
      & "ITBOOKDATE=" & sBookDte & "," & vbCrLf _
      & "ITCOMMENTS='" & sComments & "'" & vbCrLf _
      & "WHERE ITSO=" & Val(lblSon) & " AND ITNUMBER=" & Val(lblItm) & vbCrLf _
      & "AND ITREV='" & lblRev & "' "
      
  Debug.Print sSql
  
  clsADOCon.ExecuteSql sSql 'rdExecDirect
   
   'update commission info
   
   If (bUpdtComm = True) Then
     Dim Item As New ClassSoItem
     Item.InsertCommission lblSon, lblItm, lblRev, SaleSLe02a.cmbSlp
     Item.UpdateCommissions lblSon, lblItm, lblRev, bUserMsg
   End If
   clsADOCon.CommitTrans
   
   bNewItem = 0
   cmdAdd.Enabled = True
   cmdTrm.Enabled = True
   Exit Sub
   
DiaErr1:
   sProcName = "updateitem"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub AddSoItem()
   Dim iNewItem As Integer
   Dim iRows As Integer
   Dim iDelDays As Integer
   
   Dim lCurr As Long
   Dim lNew As Long
   Dim sNewDate As String
   Dim sNewPart As String
   
   Dim vDelDate As Date
   
   If Len(txtSdt) = 0 Then txtSdt = SaleSLe02a.txtDdt
   If Len(txtSdt) = 0 Then txtSdt = Format(ES_SYSDATE, "mm/dd/yyyy")
   
   'if not repeating row
   'If optRep.Value = vbUnchecked Or (Grd.Rows <= 2 And Trim(Grd.TextMatrix(0, 2)) = "") Then
   If optRep.Value = vbUnchecked Or (Grd.Rows = 1) Then
      cmbPrt = ""
      txtPrt = ""
      txtListPrice = Format("0", ES_SellingPriceFormat)
      txtQty = "0.000"
      'txtDiscountPercent = Val("0" & SaleSLe02a.txtDsc)
      txtDiscountPercent = DefaultDiscountForCustomer
   
   'if repeating an item
   Else
      Grd.Row = Grd.Rows - 1
      Grd.Col = COL_ItemNo
      lblItm = Grd.Text
      Grd.Col = COL_ItemRev
      lblRev = Grd.Text
      lblRow = Grd.Rows - 1
      bGoodItem = GetThisItem()
      
      If (OptIgnQ.Value = vbChecked) Then
         txtQty = "0.000"
      End If
      
   End If
   iDelDays = Val("0" & lblFrd)
   If txtPrt.Visible Then cmbPrt = txtPrt
   sNewPart = Compress(cmbPrt)
   sNewDate = Format(ES_SYSDATE, "mm/dd/yy")
   iNewItem = GetLastItem()
   lCurr = DateValue(Format(txtSdt, "mm/dd/yy"))
   lNew = DateValue(Format(ES_SYSDATE, "mm/dd/yy"))
   
   If lCurr < lNew Then _
      txtSdt = Format(ES_SYSDATE, "mm/dd/yyyy")
      
   On Error GoTo DiaErr1
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   sSql = "INSERT SoitTable (ITSO,ITNUMBER,ITPART,ITQTY,ITSCHED,ITBOOKDATE,ITUSER) " _
          & "VALUES(" & lblSon & "," & iNewItem & ",'" _
          & sNewPart & "'," & Val(txtQty) & ",'" & txtSdt & "','" _
          & Format(ES_SYSDATE, "mm/dd/yy") & "','" & sInitials & "')"
  clsADOCon.ExecuteSql sSql 'rdExecDirect
   
''   'if commissionable, add commission info
'   If Me.cmdCom.Enabled Then
'      Dim item As New ClassSoItem
'      item.UpdateCommission lblSon, iNewItem, "", SaleSLe02a.cmbSlp
'   End If
   
   clsADOCon.CommitTrans
   
   If clsADOCon.RowsAffected > 0 Then
      iLastitem = iNewItem
      iTotalItems = iTotalItems + 1
      lblItm = "" & str(iLastitem)
      txtCin = Format(iLastitem, "#####0")
      lblRev = ""
      txtCmt = ""
      txtRdt = txtSdt
      vDelDate = Format(txtSdt, "mm/dd/yyyy")
      vDelDate = vDelDate + iDelDays
      txtDdt = Format(vDelDate, "mm/dd/yyyy")
      txtBook = Format(ES_SYSDATE, "mm/dd/yyyy")
      'If optRep.Value = vbUnchecked Then
      '   txtDiscountPercent = SaleSLe02a.txtDis
      'End If
      
      'If grd.Rows > 2 Then grd.Rows = grd.Rows + 1
      Grd.Rows = Grd.Rows + 1
      Grd.Row = Grd.Rows - 1
      If Grd.Row > 5 Then Grd.TopRow = Grd.Row - 4
      iIndex = Grd.Row
      lblRow = iIndex
      If optRep.Value = vbChecked Then
         Grd.Col = COL_Qty
         Grd.Text = cmbPrt
         Grd.Col = COL_PartNo
         Grd.Text = txtQty
      Else
         Grd.Col = COL_Qty
         Grd.Text = "No Part Selected"
         Grd.Col = COL_PartNo
         Grd.Text = "0.000"
      End If
      Grd.Col = COL_ItemNo
      Grd.Text = Format(iLastitem, "##0")
      
      ' Reset the Invoice and Shipped check boxes
      optShp.Value = vbUnchecked
      optInv.Value = vbUnchecked
      ResetBoxes True
      bNewItem = 1
      If bAllowMo = 1 Then
         cmdCmo.Enabled = True
         cmdCmo.Visible = True
      End If
      optCreated.Value = vbUnchecked
      cmbPrt.Enabled = True

      'Add commission if applicable.
      If cmdCom.Enabled Then
        Dim Item As New ClassSoItem
        Dim bUserMsg As Boolean
        bUserMsg = False
        Item.InsertCommission lblSon, lblItm, lblRev, SaleSLe02a.cmbSlp
        Item.UpdateCommissions lblSon, lblItm, lblRev, bUserMsg
      End If
      
      
      bGoodSoMo = 0
      MouseCursor 0
      SysMsg "Item " & str(iNewItem) & " Added.", True
      'On Error Resume Next
      txtQty.SetFocus
   Else
      MsgBox "Couldn't Add Item.", vbInformation, Caption
      'On Error Resume Next
      bNewItem = 0
      Grd.Row = 1
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "addsoitem"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub grd_Click()
   UpdateItem
   DoEvents
   Grd.Col = COL_ItemNo
   lblItm = Grd.Text
   Grd.Col = COL_ItemRev
   lblRev = Grd.Text
   Grd.Col = COL_ItemNo
   iIndex = Grd.Row
   lblRow = iIndex
   bGoodItem = GetThisItem()
   
End Sub

Private Sub grd_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      UpdateItem
      iIndex = Grd.Row
      lblRow = iIndex
      Grd.Col = COL_ItemNo
      lblItm = Grd.Text
      Grd.Col = COL_ItemRev
      lblRev = Grd.Text
      Grd.Col = COL_ItemNo
      bGoodItem = GetThisItem()
   End If
   
End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** Part" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = Es_TextForeColor
   End If
   
End Sub


Private Sub lblSon_Click()
   'hold so number
   
End Sub



Private Function GetPart(sGetPart As String) As Byte
   Dim RdoPrt As ADODB.Recordset
   Dim sComment As String
   
   On Error GoTo DiaErr1
   sGetPart = Compress(sGetPart)
   If Len(sGetPart) > 0 Then
      sSql = "SELECT PARTREF,PARTNUM,PADESC,PAEXTDESC,PAPRICE,PAQOH," _
             & "PACOMMISSION FROM PartTable WHERE PARTREF='" & sGetPart & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_STATIC)
      If bSqlRows Then
         With RdoPrt
            cmbPrt = "" & Trim(!PartNum)
            txtPrt = "" & Trim(!PartNum)
            lblDsc = "" & Trim(!PADESC) & vbCrLf
            sComment = "" & Trim(!PAEXTDESC)
            If !PACOMMISSION = 1 Then cmdCom.Enabled = True _
                               Else cmdCom.Enabled = False
            If bNewItem = 1 Then
               If sOldPart <> cmbPrt Then txtListPrice = Format(!PAPRICE, ES_QuantityDataFormat)
            End If
            lblQoh = Format(!PAQOH, ES_QuantityDataFormat)
            sOldPart = cmbPrt
            sPassedPart = "" & Trim(!PartNum)
            cUR.CurrentPart = sPassedPart
            GetPart = 1
            ClearResultSet RdoPrt
         End With
      Else
         GetPart = 0
         cmbPrt = ""
         txtPrt = ""
         lblDsc = "*** Part Wasn't Found ***"
      End If
      'On Error Resume Next
      Set RdoPrt = Nothing
   Else
      txtPrt = ""
      cmbPrt = ""
      lblDsc = ""
   End If
   lblDsc = lblDsc & sComment
   Set RdoPrt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub optBok_Click()
   'never visible - Price Book is showing
   
End Sub

Private Sub optCom_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optCom_LostFocus()
   If bGoodPart = 1 Then
      If Val(txtQty) > 0 Then
         cmdAdd.Enabled = True
         cmdTrm.Enabled = True
      End If
   End If

End Sub


Private Sub optRep_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub txtBook_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtBook_LostFocus()
   txtBook = CheckDateEx(txtBook)
   If bGoodPart = 1 Then
      If Val(txtQty) > 0 Then
         cmdAdd.Enabled = True
         cmdTrm.Enabled = True
      End If
   End If
   
End Sub


Private Sub txtCin_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtCin_LostFocus()
   txtCin = CheckLen(txtCin, 6)
   'txtCin = Format(Abs(Val(txtCin)), "#####0")
   If txtPrt.Visible Then cmbPrt = txtPrt
   If Len(Trim(cmbPrt)) > 0 And Val(txtQty) > 0 Then UpdateItem
   '        With RdoItm
   '            .Edit
   '            !ITCUSTITEMNO = Val(txtCin)
   '            .Update
   '        End With
   
End Sub


Private Sub txtCmt_LostFocus()
   
   Dim strCmt As String
   
   txtCmt = CheckLenOnly(txtCmt, 3072)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   ' MM If Len(txtCmt) > 0 Then txtCmt = ReplaceString(txtCmt)
   If bGoodPart = 1 Then
      If Val(txtQty) > 0 Then
         cmdAdd.Enabled = True
         cmdTrm.Enabled = True
      End If
   End If
   strCmt = Trim(txtCmt)
   strCmt = ReplaceSingleQuote(strCmt)
   'On Error Resume Next
   sSql = "UPDATE SoitTable SET ITCOMMENTS='" & Trim(strCmt) & "' WHERE " _
          & "ITSO=" & Val(lblSon) & " AND ITNUMBER=" & Val(lblItm) _
          & " AND ITREV='" & lblRev & "' "
  clsADOCon.ExecuteSql sSql 'rdExecDirect
   If txtPrt.Visible Then cmbPrt = txtPrt
   If Len(Trim(cmbPrt)) > 0 And Val(txtQty) > 0 Then UpdateItem
   
End Sub


Private Sub txtDdt_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtDdt_LostFocus()
   txtDdt = CheckDateEx(txtDdt)
   If bGoodPart = 1 Then
      If Val(txtQty) > 0 Then
         cmdAdd.Enabled = True
         cmdTrm.Enabled = True
      End If
   End If
   If vOldDel <> txtDdt Then CheckFreightDays "DEL"
   If txtPrt.Visible Then cmbPrt = txtPrt
   If Len(Trim(cmbPrt)) > 0 And Val(txtQty) > 0 Then UpdateItem
   
End Sub

Private Sub txtDiscountPercent_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtDiscountPercent_LostFocus()
   txtDiscountPercent = CheckLen(txtDiscountPercent, 6)
   txtDiscountPercent = Format(Abs(Val(txtDiscountPercent)), "#0.000")
   If bGoodPart = 1 Then
      If Val(txtQty) > 0 Then
         cmdAdd.Enabled = True
         cmdTrm.Enabled = True
      End If
   End If
   If txtPrt.Visible Then cmbPrt = txtPrt
   If Len(Trim(cmbPrt)) > 0 And Val(txtQty) > 0 Then UpdateItem True
   '    On Error Resume Next
   '        With RdoItm
   '            .Edit
   '            !ITDISCRATE = Val(txtDiscountPercent)
   '            .Update
   '        End With
   
End Sub

Private Sub txtFrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtFrt_LostFocus()
   txtFrt = CheckLen(txtFrt, 9)
   txtFrt = Format(Abs(Val(txtFrt)), ES_QuantityDataFormat)
   If bGoodPart = 1 Then
      If Val(txtQty) > 0 Then
         cmdAdd.Enabled = True
         cmdTrm.Enabled = True
      End If
   End If
   If txtPrt.Visible Then cmbPrt = txtPrt
   If Len(Trim(cmbPrt)) > 0 And Val(txtQty) > 0 Then UpdateItem True
   '    On Error Resume Next
   '        With RdoItm
   '            .Edit
   '            !ITFRTALLOW = Val(txtFrt)
   '            .Update
   '        End With
   
End Sub

Private Sub txtListPrice_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtListPrice_LostFocus()
   Dim bResponse As Byte
   Dim sMsg As String
   txtListPrice = CheckLen(txtListPrice, 12)
   'txtListPrice = Format(Abs(Val(txtListPrice)), ES_SellingPriceFormat)
   txtListPrice = Format(Val(txtListPrice), ES_SellingPriceFormat)      'negative OK
   If bGoodPart = 1 Then
      If Val(txtQty) > 0 Then
         cmdAdd.Enabled = True
         cmdTrm.Enabled = True
         If txtPrt.Visible Then cmbPrt = txtPrt
         If Len(Trim(cmbPrt)) > 0 And Val(txtQty) > 0 Then UpdateItem True
      End If
   End If
   If optShp.Value = vbChecked Then
      If cOldPrice <> Val(txtListPrice) Then
         sMsg = "This Item Has A Packing Slip.  Do You " & vbCrLf _
                & "Want To Update The Price For The Item?" & vbCrLf _
                & "Note: Items Not Invoiced."
         bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
         If bResponse = vbYes Then
            UpdatePackSlip
         Else
            CancelTrans
            txtListPrice = Format(cOldPrice, ES_SellingPriceFormat)
         End If
      End If
   End If
   
End Sub

Private Sub txtPrt_Change()
   cmbPrt = txtPrt
   
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
   'always allow search box
   If cmdFnd.Value <> 0 Then
      Exit Sub
   End If
   txtPrt = CheckLen(txtPrt, 30)
   If bView = 1 Then Exit Sub
   'On Error Resume Next
   bGoodPart = GetPart(txtPrt)
   cmbPrt = txtPrt
   If bGoodPart = 1 Then
   
      ' don't allow inactive or obsolete part
      If (Not ValidPartNumber(cmbPrt.Text)) Then
         MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
            vbInformation, Caption
         cmbPrt = ""
         txtPrt = ""
         Exit Sub
      End If
      
      GetBookPrice
      If Val(txtQty) > 0 Then
         cmdAdd.Enabled = True
         cmdTrm.Enabled = True
         Grd.Row = iIndex
         Grd.Col = COL_Qty
         Grd.Text = txtPrt
         Grd.Col = COL_ItemNo
         If Len(Trim(txtPrt)) > 0 And Val(txtQty) > 0 Then UpdateItem True, False
      End If
   End If
   Grd.Enabled = True
   
End Sub


Private Sub txtQty_Click()
   Grd.Enabled = False
   
End Sub

Private Sub txtQty_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub


Private Sub txtQty_LostFocus()
   txtQty = CheckLen(txtQty, 9)
   txtQty = Format(Abs(Val(txtQty)), ES_QuantityDataFormat)
   If Val(txtQty) = 0 Then txtQty = Format(cOldQty, ES_QuantityDataFormat)
   cOldQty = Val(txtQty)
   Grd.Col = COL_PartNo
   Grd.Text = txtQty
   Grd.Col = COL_ItemNo
   If Len(Trim(cmbPrt)) > 0 And Val(txtQty) > 0 Then UpdateItem True
   '   On Error Resume Next
   '       With RdoItm
   '           .Edit
   '          '!ITQTY = Val(txtQty)
   '           .Update
   '       End With
   Grd.Enabled = True
   
End Sub

Private Sub txtRdt_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtRdt_LostFocus()
   txtRdt = CheckDateEx(txtRdt)
   If bGoodPart = 1 Then
      If Val(txtQty) > 0 Then
         cmdAdd.Enabled = True
         cmdTrm.Enabled = True
      End If
   End If
   If txtPrt.Visible Then cmbPrt = txtPrt
   If Len(Trim(cmbPrt)) > 0 And Val(txtQty) > 0 Then UpdateItem
   
End Sub

Private Sub txtSdt_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtSdt_LostFocus()
   txtSdt = CheckDateEx(txtSdt)
   'On Error Resume Next
   If bGoodPart = 1 Then
      If Val(txtQty) > 0 Then
         cmdAdd.Enabled = True
         cmdTrm.Enabled = True
      End If
   End If
   If bNewItem = 1 Then
      txtRdt = txtSdt
      txtDdt = txtSdt
      bNewItem = 0
   End If
   If vOldShip <> txtSdt Then CheckFreightDays "SCHED"
   If txtPrt.Visible Then cmbPrt = txtPrt
   If Len(Trim(cmbPrt)) > 0 And Val(txtQty) > 0 Then UpdateItem
   
End Sub



Private Sub GetAllItems()
   Dim RdoGet As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT ITSO,ITNUMBER,ITCANCELED FROM SoitTable " _
          & "WHERE (ITSO=" & str(lSalesOrder) & " AND ITCANCELED=0) " _
          & "ORDER BY ITNUMBER DESC"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet, ES_FORWARD)
   If bSqlRows Then iLastitem = Format(RdoGet!ITNUMBER, "##0")
   Set RdoGet = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getallitem"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub ResetBoxes(bOpen As Boolean)
   If Not bOpen Then
      optShp.Value = vbChecked
      optCom.Enabled = False
      txtQty.Enabled = False
      txtPrt.Enabled = False
      cmbPrt.Enabled = False
      txtListPrice.Enabled = False
      txtSdt.Enabled = False
      txtRdt.Enabled = False
      txtDdt.Enabled = False
      txtFrt.Enabled = False
      txtDiscountPercent.Enabled = False
      txtCmt.Enabled = False
      cmdFnd.Enabled = False
      cmdBok.Enabled = False
   Else
      optShp.Value = vbUnchecked
      optCom.Enabled = True
      cmbPrt.Enabled = True
      txtPrt.Enabled = True
      txtListPrice.Enabled = True
      txtQty.Enabled = True
      txtListPrice.Enabled = True
      txtSdt.Enabled = True
      txtRdt.Enabled = True
      txtDdt.Enabled = True
      txtFrt.Enabled = True
      txtDiscountPercent.Enabled = True
      txtCmt.Enabled = True
      cmdBok.Enabled = True
      cmdFnd.Enabled = True
   End If
   If optInv.Value = vbChecked Then
      optCom.Enabled = False
      txtQty.Enabled = False
      cmbPrt.Enabled = False
      txtListPrice.Enabled = False
      txtSdt.Enabled = False
      txtRdt.Enabled = False
      txtDdt.Enabled = False
      txtFrt.Enabled = False
      txtDiscountPercent.Enabled = False
      txtCmt.Enabled = False
      cmdBok.Enabled = False
      cmdFnd.Enabled = False
   Else
      txtListPrice.Enabled = True
   End If
   
End Sub

Private Sub UpdatePackSlip()
   'On Error Resume Next
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   sSql = "UPDATE SoitTable SET ITDOLLARS=" & Me.lblNetPrice & "," & vbCrLf _
      & "ITDOLLORIG=" & Me.txtListPrice & "," & vbCrLf _
      & "ITDISCRATE=" & Me.txtDiscountPercent & "," & vbCrLf _
      & "ITDISCAMOUNT=" & Me.lblDiscountAmount & vbCrLf _
      & "WHERE ITSO=" & Val(lblSon) & " AND ITNUMBER=" & Val(lblItm) & " " & vbCrLf _
      & "AND ITREV='" & lblRev & "' "
  clsADOCon.ExecuteSql sSql 'rdExecDirect
   sSql = "UPDATE PsitTable SET PISELLPRICE=" & Me.lblNetPrice & " " _
          & "WHERE (PIPACKSLIP='" & lblPsNum & "' AND PISONUMBER=" _
          & Val(lblSon) & " AND PISOITEM=" & Val(lblItm) _
          & " AND PISOREV='" & lblRev & "')"
  clsADOCon.ExecuteSql sSql 'rdExecDirect
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      MsgBox "Update Was Successful.", _
         vbInformation, Caption
      cOldPrice = Val(txtListPrice)
   Else
      clsADOCon.RollbackTrans
      MsgBox "Update Was Not Successful.", _
         vbExclamation, Caption
      txtListPrice = Format(cOldPrice, ES_SellingPriceFormat)
   End If
   
End Sub

Private Sub GetBookPrice()
   Dim RdoBok As ADODB.Recordset
   Dim bResponse As Byte
   Dim sMsg As String
   If txtPrt.Visible Then cmbPrt = txtPrt
   If bGoodBook = 1 Then
      sSql = "SELECT PARTREF,PARTNUM,PBIREF,PBIPARTREF,PBIPRICE " _
             & "FROM PartTable,PbitTable WHERE (PARTREF=PBIPARTREF) AND " _
             & "(PBIREF='" & sBook & "') AND (PARTREF='" & Compress(cmbPrt) & "')"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoBok, ES_FORWARD)
      If bSqlRows Then
         With RdoBok
            If Val(txtListPrice) <> !PBIPRICE Then
               sMsg = "Overwrite Price With Current Book Price?"
               bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
               If bResponse = vbYes Then _
                              txtListPrice = Format(!PBIPRICE, ES_SellingPriceFormat)
                              
            End If
            ClearResultSet RdoBok
         End With
      End If
   End If
   Set RdoBok = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getbookpr"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function TestBookDates() As Byte
   Dim RdoBok As ADODB.Recordset
   Dim bDate As Byte
   Dim bResponse As Byte
   Dim sMsg As String
   Dim vDate As Variant
   
   vDate = Format(GetServerDateTime, "mm/dd/yyyy")
   sBook = Trim(sBook)
   If bGoodBook = 1 Then
      bDate = 1
      'Trap Dates
      sSql = "SELECT PBHREF,PBHSTARTDATE,PBHENDDATE FROM PbhdTable " _
             & "WHERE PBHREF='" & sBook & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoBok, ES_FORWARD)
      With RdoBok
         If IsNull(.Fields(0)) Then bDate = 0
         If IsNull(.Fields(1)) Then bDate = 0
         ClearResultSet RdoBok
      End With
      If bDate = 1 Then
         sSql = "SELECT PBHREF,PBHSTARTDATE,PBHENDDATE FROM PbhdTable " _
                & "WHERE ('" & vDate & "' BETWEEN " _
                & "PBHSTARTDATE AND PBHENDDATE) AND PBHREF='" & sBook & "'"
         bSqlRows = clsADOCon.GetDataSet(sSql, RdoBok, ES_FORWARD)
         If bSqlRows Then
            TestBookDates = 0
         Else
            TestBookDates = 1
            MsgBox "Your Assigned Price Book, " & Trim(SaleSLe02a.txtBook) & ", Is Out Of Date.", _
               vbInformation, Caption
         End If
      Else
         TestBookDates = 0
      End If
      
   Else
      TestBookDates = 1
   End If
   Set RdoBok = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "testbookda"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function GetLastItem() As Long
   Dim RdoLst As ADODB.Recordset
   'On Error Resume Next
   sSql = "SELECT MAX(ITNUMBER) FROM SoitTable WHERE " _
          & "ITSO=" & Val(lblSon) & ""
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_FORWARD)
   If bSqlRows Then
      With RdoLst
         If Not IsNull(.Fields(0)) Then
            GetLastItem = .Fields(0) + 1
         Else
            GetLastItem = 1
         End If
         ClearResultSet RdoLst
      End With
   Else
      GetLastItem = 1
   End If
   Set RdoLst = Nothing
   If Err > 0 Then GetLastItem = 1
   
End Function

Private Sub CheckFreightDays(TypeOfDays As String)
   Dim dNewDate As Date
   If optFreight.Value = vbUnchecked Then
      'On Error Resume Next
      If TypeOfDays = "DEL" Then
         dNewDate = Format(txtDdt, "mm/dd/yyyy")
         txtSdt = Format(dNewDate - bFreightDays, "mm/dd/yyyy")
      Else
         dNewDate = Format(txtSdt, "mm/dd/yyyy")
         txtDdt = Format(dNewDate + bFreightDays, "mm/dd/yyyy")
      End If
   End If
   vOldShip = txtSdt
   vOldDel = txtDdt
   
End Sub


Private Function CheckAllocations()
   Dim RdoAll As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT AllowMOCreationFromSO From Preferences WHERE PreRecord=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAll, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoAll.Fields(0)) Then
         CheckAllocations = RdoAll.Fields(0)
      Else
         CheckAllocations = 0
      End If
   End If
   Set RdoAll = Nothing
   Exit Function
DiaErr1:
   CheckAllocations = 0
   
End Function

Public Sub CalculateDiscount()
   Dim listPrice As Currency, discRate As Currency, discAmount As Currency
   'listPrice = "0" & Me.txtListPrice
   listPrice = IIf(IsNumeric(txtListPrice), txtListPrice, 0)
   'discRate = "0" & Me.txtDiscountPercent
   discRate = IIf(IsNumeric(txtDiscountPercent), txtDiscountPercent, 0)
   discAmount = listPrice * discRate / 100
   'Me.lblDiscountAmount = Format(discAmount, "######0.000")
   'Me.lblNetPrice = Format(CCur(listPrice) - discAmount, "######0.0000")
   Me.lblDiscountAmount = Format(discAmount, ES_SellingPriceFormat)
   Me.lblNetPrice = Format(CCur(listPrice) - discAmount, ES_SellingPriceFormat)
   'grd.TextMatrix(grd.Row, COL_UnitPrice) = Format(Me.lblNetPrice, "######0.000")
   'iindex is the row we are leaving.  grd.Row is the row we are entering
   Grd.TextMatrix(Val(iIndex), COL_UnitPrice) = Format(Me.lblNetPrice, ES_SellingPriceFormat)
End Sub

Private Function ShowIgnoreQtyOnSO()

   Dim rdo As ADODB.Recordset
   sSql = "SELECT ISNULL(COIGNQYTRPTSO,0) from ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD)
   
   If bSqlRows Then
      If rdo.Fields(0) = 1 Then
         OptIgnQ.Visible = True
         OptIgnQ.Value = vbUnchecked
      Else
         OptIgnQ.Visible = False
         OptIgnQ.Value = vbUnchecked
      End If
   End If
   Set rdo = Nothing

End Function

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


