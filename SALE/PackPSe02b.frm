VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form PackPSe02b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Packing Slip Items"
   ClientHeight    =   5640
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   7110
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   2270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAddS 
      Caption         =   "Add New &Sales Item"
      Height          =   315
      Left            =   1080
      TabIndex        =   64
      TabStop         =   0   'False
      ToolTipText     =   "Add new Sales Item"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CommandButton cmdUp 
      DisabledPicture =   "PackPSe02b.frx":0000
      DownPicture     =   "PackPSe02b.frx":04F2
      Enabled         =   0   'False
      Height          =   372
      Left            =   6600
      MaskColor       =   &H00000000&
      Picture         =   "PackPSe02b.frx":09E4
      Style           =   1  'Graphical
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   4800
      Width           =   400
   End
   Begin VB.CommandButton cmdDn 
      DisabledPicture =   "PackPSe02b.frx":0ED6
      DownPicture     =   "PackPSe02b.frx":13C8
      Enabled         =   0   'False
      Height          =   372
      Left            =   6600
      MaskColor       =   &H00000000&
      Picture         =   "PackPSe02b.frx":18BA
      Style           =   1  'Graphical
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   5184
      Width           =   400
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PackPSe02b.frx":1DAC
      Style           =   1  'Graphical
      TabIndex        =   61
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "PackPSe02b.frx":255A
      DownPicture     =   "PackPSe02b.frx":2ECC
      Height          =   350
      Index           =   3
      Left            =   5520
      Picture         =   "PackPSe02b.frx":383E
      Style           =   1  'Graphical
      TabIndex        =   59
      ToolTipText     =   "Standard Comments"
      Top             =   4440
      Width           =   350
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "PackPSe02b.frx":3E40
      DownPicture     =   "PackPSe02b.frx":47B2
      Height          =   350
      Index           =   2
      Left            =   5520
      Picture         =   "PackPSe02b.frx":5124
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "Standard Comments"
      Top             =   3240
      Width           =   350
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "PackPSe02b.frx":5726
      DownPicture     =   "PackPSe02b.frx":6098
      Height          =   350
      Index           =   1
      Left            =   5520
      Picture         =   "PackPSe02b.frx":6A0A
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "Standard Comments"
      Top             =   2040
      Width           =   350
   End
   Begin VB.TextBox lblCmt 
      Height          =   495
      Index           =   3
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   4440
      Width           =   4335
   End
   Begin VB.TextBox lblCmt 
      Height          =   495
      Index           =   2
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   3240
      Width           =   4335
   End
   Begin VB.TextBox lblCmt 
      Height          =   495
      Index           =   1
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Tag             =   "9"
      Top             =   2040
      Width           =   4335
   End
   Begin VB.CheckBox optVew 
      Caption         =   "Check1"
      Height          =   255
      Left            =   4680
      TabIndex        =   53
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtBox 
      Height          =   285
      Index           =   3
      Left            =   5520
      TabIndex        =   10
      Top             =   4125
      Width           =   975
   End
   Begin VB.TextBox txtBox 
      Height          =   285
      Index           =   2
      Left            =   5520
      TabIndex        =   6
      Top             =   2925
      Width           =   975
   End
   Begin VB.TextBox txtBox 
      Height          =   285
      Index           =   1
      Left            =   5520
      TabIndex        =   2
      ToolTipText     =   "Enter Box Label (If one)"
      Top             =   1725
      Width           =   975
   End
   Begin VB.CheckBox optCmt 
      Caption         =   "Include SO Item Comments"
      Height          =   195
      Left            =   4200
      TabIndex        =   52
      ToolTipText     =   "Include SO Item Comments On Pack Slip"
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "PackPSe02b.frx":700C
      Height          =   315
      Left            =   3600
      Picture         =   "PackPSe02b.frx":74E6
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "View Selected Items"
      Top             =   840
      Width           =   375
   End
   Begin VB.CheckBox optShp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   1
      Left            =   6600
      TabIndex        =   1
      ToolTipText     =   "Mark To Ship "
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtShp 
      Height          =   285
      Index           =   1
      Left            =   5520
      TabIndex        =   0
      ToolTipText     =   "Quantity To Ship"
      Top             =   1440
      Width           =   975
   End
   Begin VB.CheckBox optShp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   2
      Left            =   6600
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtShp 
      Height          =   285
      Index           =   2
      Left            =   5520
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtShp 
      Height          =   285
      Index           =   3
      Left            =   5520
      TabIndex        =   8
      Top             =   3840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox optShp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   3
      Left            =   6600
      TabIndex        =   9
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   315
      Left            =   6120
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Add Selections To Pack Slip (When Finished)"
      Top             =   520
      Width           =   915
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   1200
      Top             =   5520
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5640
      FormDesignWidth =   7110
   End
   Begin VB.Label lblSon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1500
      TabIndex        =   68
      Top             =   840
      Width           =   885
   End
   Begin VB.Label lblSchDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   67
      Top             =   4125
      Width           =   855
   End
   Begin VB.Label lblSchDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   66
      Top             =   2925
      Width           =   855
   End
   Begin VB.Label lblSchDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   65
      Top             =   1725
      Width           =   855
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1080
      TabIndex        =   60
      Top             =   840
      Width           =   350
   End
   Begin VB.Image Dsup 
      Height          =   300
      Left            =   360
      Picture         =   "PackPSe02b.frx":79C0
      Top             =   5280
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Enup 
      Height          =   300
      Left            =   120
      Picture         =   "PackPSe02b.frx":7EB2
      Top             =   5280
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Endn 
      Height          =   300
      Left            =   720
      Picture         =   "PackPSe02b.frx":83A4
      Top             =   5280
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Dsdn 
      Height          =   300
      Left            =   960
      Picture         =   "PackPSe02b.frx":8896
      Top             =   5280
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label lblQoh 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   4200
      TabIndex        =   56
      ToolTipText     =   "Quantity On Hand"
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label lblQoh 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   55
      ToolTipText     =   "Quantity On Hand"
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label lblQoh 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   54
      ToolTipText     =   "Quantity On Hand"
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pack Slip"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   51
      Top             =   520
      Width           =   975
   End
   Begin VB.Label lblPsl 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1080
      TabIndex        =   50
      Top             =   520
      Width           =   975
   End
   Begin VB.Label lblUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   5040
      TabIndex        =   48
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label lblUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   47
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lblUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   46
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "So Itm"
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
      Left            =   120
      TabIndex        =   45
      Top             =   1240
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev  "
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
      Left            =   600
      TabIndex        =   44
      Top             =   1240
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number                                               "
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
      Left            =   1080
      TabIndex        =   43
      Top             =   1240
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Order Qty/Qoh            "
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
      Left            =   4200
      TabIndex        =   42
      Top             =   1240
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ship Qty/Box          "
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
      Index           =   6
      Left            =   5520
      TabIndex        =   41
      Top             =   1240
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ship   "
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
      Index           =   7
      Left            =   6600
      TabIndex        =   40
      Top             =   1240
      Width           =   615
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   39
      Top             =   1725
      Width           =   3015
   End
   Begin VB.Label lblOrd 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   38
      ToolTipText     =   "Click To Insert Ship Quantity"
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   37
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label lblRev 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   36
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblitm 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   35
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label lblOrd 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   34
      ToolTipText     =   "Click To Insert Ship Quantity"
      Top             =   2640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   33
      Top             =   2640
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblRev 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   32
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblitm 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   31
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   30
      Top             =   2925
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label z2 
      BackStyle       =   0  'Transparent
      Caption         =   "PS Item Comments"
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   29
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label z2 
      Caption         =   "PS Item Comments"
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   28
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label z2 
      Caption         =   "PS Item Comments"
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   27
      Top             =   4440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   26
      Top             =   4125
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblitm 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   25
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblRev 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   24
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   23
      Top             =   3840
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblOrd 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   4200
      TabIndex        =   22
      ToolTipText     =   "Click To Insert Ship Quantity"
      Top             =   3840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblPge 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5280
      TabIndex        =   21
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label lblLst 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6120
      TabIndex        =   20
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Page"
      Height          =   255
      Index           =   9
      Left            =   4680
      TabIndex        =   19
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "of"
      Height          =   255
      Index           =   8
      Left            =   5760
      TabIndex        =   18
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2280
      TabIndex        =   16
      Top             =   220
      Width           =   3135
   End
   Begin VB.Label lblCst 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1080
      TabIndex        =   15
      Top             =   220
      Width           =   1200
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   220
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "PackPSe02b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'PS Item Comments 10/22/03
'2/11/02 added Primary Sales Order for Ship To
'11/8/05 Reformatted the SO Selection
'12/15/05 Changed GetLastItem() skips a row on revised PackSlips
'11/10/06 Corrected CreatePsTable - decimal to dec(9,4)
Option Explicit
'Dim rdoQry As rdoQuery
Dim cmdObj As ADODB.Command

Dim bGoodItems As Byte
Dim bGoodSO As Byte
Dim bOnLoad As Byte
Dim bPsSaved As Byte

Dim iBookMark As Integer
Dim iCurrPage As Integer
Dim iMaxPages As Integer
Dim iNewBookmark As Integer
'Dim lSalesOrders As Integer
Dim iSoItems As Integer
Dim iTotalItems As Integer

Dim sStName As String
Dim sStAdr As String
Dim sTableDef As String

Dim sPSComments(1200) As String 'Item Comments
'Dim vSalesOrders(90, 3) As Variant    'CAN ONLY HANDLE ONE SO PER INSTANCE ANYWAY
'      vSalesOrders(lSalesOrders, 0) = cmbSon
'      vSalesOrders(lSalesOrders, 1) = vBookmark
'      vSalesOrders(lSalesOrders, 2) = iSoItems
Dim vItems(1200, 15) As Variant
'   vItems(iTotalitems, 0) = Right(cmbSon, 5)
'   vItems(iTotalitems, 1) = Format(!ITNUMBER, "#0")
'   vItems(iTotalitems, 2) = Trim(!ITREV)
'   vItems(iTotalitems, 3) = Trim(!PARTNUM)
'   vItems(iTotalitems, 4) = Trim(!PADESC)
'   vItems(iTotalitems, 5) = Trim(!ITQTY)
Private Const PS_SOQTY = 5
'   vItems(iTotalitems, 6) = Trim(!ITCOMMENTS)
'   vItems(iTotalitems, 7) = Trim(!PAUNITS)
'   vItems(iTotalitems, 8) = Ship Qty
Private Const PS_QTYTOSHIP = 8
'   vItems(iTotalitems, 9) = Ship? (T/F)
'   vItems(iTotalitems, 10) = Compressed Part
'   vItems(iTotalitems, 11) = Selling Price
'   vItems(iTotalitems, 12) = Box
'   vItems(iTotalitems, 13) = Qoh
Private Const PS_QTYONHAND = 13
Private Const PS_ITSCHED = 15

Dim sPartNumber(1200) As String

Private Sub CreatePsTable()
   Dim b As Byte
   Dim iRows As Integer
   Dim RdoCols As ADODB.Recordset
   Dim sTableN As String
   Dim sCol1 As Variant
   Dim sCol2 As Variant
   Dim sCol3 As Variant
   Dim sCol4 As Variant
   Dim sTable(100) As Variant
   MouseCursor 13
   On Error GoTo DiaErr1
   Err = 0
   clsADOCon.ADOErrNum = 0
   sSql = "sp_columns 'SoitTable'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCols, ES_FORWARD)
   If bSqlRows Then
      On Error GoTo 0
      With RdoCols
         sTableN = UCase(Format(Now, "ddd")) & Right(Compress(GetNextLotNumber()), 8)
         sTableDef = Trim$(sTableN)
         Do Until .EOF
            iRows = iRows + 1
            sCol1 = Trim(.Fields(3))
            sCol2 = Trim(.Fields(5))
            sCol3 = Trim(.Fields(7))
            If sCol1 = "" Then Exit Do
            sCol4 = ""
            If iRows > 1 Then sCol1 = "," & sCol1
            If LCase$(sCol2) = "decimal" Then
               sCol2 = "dec(12,4) "
            End If
            If sCol2 = "char" Or sCol2 = "varchar" Then
               sCol3 = "(" & sCol3 & ") Null "
               sCol4 = "default('')"
            Else
               sCol3 = " Null "
               If sCol2 <> "smalldatetime" Then sCol4 = "default(0)"
            End If
            sTable(iRows) = sCol1 & " " & sCol2 & sCol3 & sCol4
            .MoveNext
         Loop
         sTable(iRows) = sTable(iRows) & ")"
         ClearResultSet RdoCols
      End With
   End If
   If iRows > 0 Then
      sTableN = "create table " & sTableDef & " ("
      For b = 1 To iRows
         sTableN = sTableN & sTable(b)
      Next
      clsADOCon.ExecuteSql sTableN
      sSql = "create unique clustered index WorkRef on " & sTableDef & " " _
             & "(ITSO,ITNUMBER,ITREV) WITH FILLFACTOR=80"
      clsADOCon.ExecuteSql sSql '
   Else
      sTableDef = "PswkTable"
   End If
   Set RdoCols = Nothing
   Exit Sub
   
DiaErr1:
   sTableDef = "PswkTable"
   
End Sub












Private Sub cmbSon_Click()
   bGoodSO = CheckSalesOrder()
   CloseBoxes 'False
   
End Sub

Private Sub cmbSon_GotFocus()
   SelectFormat Me
   
End Sub


Private Sub cmbSon_KeyPress(KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub


'Private Sub cmbSon_LostFocus()
'   cmbSon = Format(Abs(Val(cmbSon)), SO_NUM_FORMAT)
'   bGoodSO = CheckSalesOrder()
'
'End Sub

Private Sub cmdAdd_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   sMsg = "Are You Finished With Your Selections?" & vbCrLf _
          & "Do You Wish to Record The Items?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      CreatePackSlip
   Else
      On Error Resume Next
      txtShp(1).SetFocus
   End If
   
End Sub

Private Sub cmdAddS_Click()
   Dim bRet As Boolean
   
   PackPSe02c.lblCst = lblCst
   PackPSe02c.lblNme = lblNme
   PackPSe02c.lblSon = lblSon
   PackPSe02c.lblType = lblType
   
   PackPSe02c.Show
   
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdComments_Click(Index As Integer)
   'Add one of these to the form and go
   If cmdComments(Index) Then
      SysComments.lblControl = "lblCmt(" & Trim(str(Index)) & ")"
      SysComments.lblListIndex = 3
      SysComments.Show
      cmdComments(Index) = False
   End If
   
End Sub

Private Sub cmdDn_Click()
   iCurrPage = iCurrPage + 1
   GetNextGroup
   
End Sub

Private Sub cmdDn_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2270
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


'Private Sub cmdSel_Click()
'   If bGoodSO = 0 Then
'      If bGoodSO = 0 Then MsgBox "That Sales Order Does Not Belong To " & lblCst & ".", _
'                   vbInformation, Caption
'   Else
'      bGoodItems = GetItems()
'   End If
'
'End Sub
'

Private Sub cmdUp_Click()
   iCurrPage = iCurrPage - 1
   GetNextGroup
   
End Sub

Private Sub cmdUp_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub


Private Sub cmdVew_Click()
   Dim iList As Integer
   Dim a As Integer
   Dim sItem As String * 11
   Dim sString As String * 20
   If iTotalItems > 0 Then
      For iList = 1 To iTotalItems
         If vItems(iList, 9) Then
            sItem = vItems(iList, 0) & "-" & vItems(iList, 1) & vItems(iList, 2)
            a = Len(Trim(vItems(iList, PS_SOQTY)))
            a = 9 - a
            sString = Left(vItems(iList, 3), 20)
            PackPSIview.lstItm.AddItem sItem _
               & " " & sString _
               & String(a, Chr(160)) _
               & vItems(iList, PS_QTYTOSHIP) & vItems(iList, 7)
         End If
      Next
   Else
      MsgBox "No Items Selected For This Revision.", _
         vbInformation
      Exit Sub
   End If
   optVew.Value = vbChecked
   PackPSIview.Show
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      CreatePsTable
      'FillCombo
      If Val(Right(PackPSe02a.lblPson, SO_NUM_SIZE)) > 0 Then lblSon = Right(PackPSe02a.lblPson, SO_NUM_SIZE)
      bGoodSO = CheckSalesOrder()
      

      ' Check the capability to add SO items.
      CheckAlloNewSOit
      CloseBoxes 'False
      
      bGoodItems = GetItems()
      
      bOnLoad = 0
   End If
   If optVew.Value = vbChecked Then
      Unload PackPSIview
   End If
   MouseCursor 0
   
End Sub




Private Sub Form_Load()
   SetFormSize Me
   Move PackPSe02a.Top + 200, PackPSe02a.Left + 200
   lblPsl = PackPSe02a.cmbPsl
   sCurrForm = ""
   sSql = "SELECT ITSO,ITNUMBER,ITREV,ITQTY,ITPART,ITDOLLARS,ITSCHED," _
          & "ITCANCELED,ITCOMMENTS,PARTREF,PARTNUM,PADESC,PAUNITS," _
          & "PAQOH From SoitTable,PartTable WHERE ITSO= ? AND " _
          & "(ITPART=PARTREF AND ITACTUAL IS NULL AND ITCANCELED=0 " _
          & "AND ITQTY>0 AND ITPSNUMBER='' AND ITINVOICE=0)" & vbCrLf _
          & "ORDER BY ITSO, ITNUMBER, ITREV"
   
   'Set rdoQry = RdoCon.CreateQuery("", sSql)
   Set cmdObj = New ADODB.Command
   cmdObj.CommandText = sSql
   
   Dim prmObj As ADODB.Parameter
   Set prmObj = New ADODB.Parameter
   prmObj.Type = adInteger
   cmdObj.parameters.Append prmObj
   
   bPsSaved = False
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim bResponse As Byte
   Dim sMsg As String
   On Error Resume Next
   If Not bPsSaved Then
      sMsg = "You Have Not Added Any Items." & vbCrLf _
             & "Do You Really Want Close?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbNo Then
         Cancel = True
         On Error Resume Next
         SetFocus
         'cmbSon.SetFocus
         Exit Sub
      End If
   End If
   If UCase$(Left$(sTableDef, 2)) <> "PS" Then
      sSql = "DROP TABLE " & sTableDef
      clsADOCon.ExecuteSql sSql ', rdExecDirect
   End If
   PackPSe02a.optItm.Value = vbUnchecked
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   MdiSect.lblBotPanel = PackPSe02a.Caption
   Erase vItems
   On Error Resume Next
   Set cmdObj = Nothing
   Set PackPSe02b = Nothing
   
End Sub



'Private Sub FillCombo()
'   Dim sCust As String
'   cmbSon.Clear
'   sCust = Compress(lblCst)
'   On Error GoTo DiaErr1
'   sSql = "SELECT DISTINCT SONUMBER,ITSO FROM SohdTable,SoitTable " _
'          & "WHERE SONUMBER=ITSO AND (SOCUST='" & sCust & "' AND ITQTY>0 " _
'          & "AND ITPSNUMBER='' AND ITINVOICE=0 AND ITCANCELED=0) " _
'          & "ORDER BY SONUMBER DESC"
'   LoadNumComboBox cmbSon, SO_NUM_FORMAT
'   If Val(Right(PackPSe02a.lblPson, SO_NUM_SIZE)) > 0 Then cmbSon = Right(PackPSe02a.lblPson, SO_NUM_SIZE)
'   bGoodSO = CheckSalesOrder()
'   Exit Sub
'
'DiaErr1:
'   On Error Resume Next
'   If UCase$(Left$(sTableDef, 2)) <> "PS" Then
'      sSql = "DROP TABLE " & sTableDef
'      clsADOCon.ExecuteSql sSql ', rdExecDirect
'   End If
'   sProcName = "fillcombo"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub



Private Function GetItems() As Byte
   Dim RdoItm As ADODB.Recordset
   Dim bSoIncluded As Boolean
   Dim iRow As Integer
   Dim vBookmark As Variant
   
   cmdUp.Picture = Dsup
   cmdDn.Picture = Dsdn
   cmdUp.Enabled = False
   cmdDn.Enabled = False
   lblLst = 1
   lblPge = 1
   iTotalItems = 0
   On Error GoTo DiaErr1
   bSoIncluded = False
'   For iRow = 1 To lSalesOrders
'      If vSalesOrders(iRow, 0) = cmbSon Then
'         iBookMark = vSalesOrders(iRow, 1)
'         bSoIncluded = True
'         Exit For
'      End If
'   Next
   If bSoIncluded Then
      GetItems = 1
      'GetThisSalesOrder
      Exit Function
   Else
      For iRow = 1 To 3
         txtShp(iRow) = ""
         txtBox(iRow) = ""
      Next
   End If
   iRow = 0
   'rdoQry(0) = Val(cmbSon)
   cmdObj.parameters(0).Value = Val(lblSon)
   bSqlRows = clsADOCon.GetQuerySet(RdoItm, cmdObj, ES_STATIC, True)
   If bSqlRows Then
      With RdoItm
         Do Until .EOF
            iRow = iRow + 1
            If iRow < 4 Then
               lblitm(iRow).Visible = True
               lblRev(iRow).Visible = True
               lblSchDte(iRow).Visible = True
               lblPrt(iRow).Visible = True
               lblDsc(iRow).Visible = True
               lblOrd(iRow).Visible = True
               txtShp(iRow).Visible = True
               txtBox(iRow).Visible = True
               lblCmt(iRow).Visible = True
               optShp(iRow).Visible = True
               lblUom(iRow).Visible = True
               lblQoh(iRow).Visible = True
               cmdComments(iRow).Visible = True
               z2(iRow).Visible = True
               lblitm(iRow) = Format(!ITNUMBER, "#0")
               lblRev(iRow) = "" & Trim(!itrev)
               lblSchDte(iRow) = Format(!itsched, "MM/dd/yy")
               lblPrt(iRow) = "" & Trim(!PartNum)
               lblDsc(iRow) = "" & Trim(!PADESC)
               lblOrd(iRow) = Format(!ITQty, ES_QuantityDataFormat)
               txtShp(iRow) = "0.000"
               lblCmt(iRow) = "" & Trim(!ITCOMMENTS)
               lblUom(iRow) = "" & Trim(!PAUNITS)
               lblQoh(iRow) = "" & Trim(!PAQOH)
            End If
            iTotalItems = iTotalItems + 1
            If iRow = 1 Then vBookmark = iTotalItems
            vItems(iTotalItems, 0) = "" & lblSon
            vItems(iTotalItems, 1) = Format(!ITNUMBER, "#0")
            vItems(iTotalItems, 2) = "" & Trim(!itrev)
            vItems(iTotalItems, 3) = "" & Trim(!PartNum)
            vItems(iTotalItems, 4) = "" & Trim(!PADESC)
            vItems(iTotalItems, PS_SOQTY) = Format(!ITQty, ES_QuantityDataFormat)
            ' vItems(iTotalItems, 6) = "" & Trim(!ITCOMMENTS)
            vItems(iTotalItems, 7) = "" & Trim(!PAUNITS)
            vItems(iTotalItems, PS_QTYTOSHIP) = "0.000"
            vItems(iTotalItems, 9) = 0
            vItems(iTotalItems, 10) = "" & Trim(!PartRef)
            sPartNumber(iTotalItems) = "" & Trim(!PartRef)
            vItems(iTotalItems, 11) = Format(!ITDOLLARS, ES_QuantityDataFormat)
            vItems(iTotalItems, 12) = ""
            vItems(iTotalItems, PS_QTYONHAND) = Format(!PAQOH, ES_QuantityDataFormat)
            vItems(iTotalItems, PS_ITSCHED) = Format(!itsched, "MM/dd/yy")
            .MoveNext
         Loop
         ClearResultSet RdoItm
         GetItems = 1
      End With
      iSoItems = iRow
      iCurrPage = 1
      'iMaxPages = (iSoItems / 3) + 0.45
      iMaxPages = (iSoItems + 2) \ 3
      lblLst = iCurrPage
      lblLst = iMaxPages
      If iSoItems > 3 Then
         cmdDn.Picture = Endn
         cmdDn.Enabled = True
      End If
'      lSalesOrders = lSalesOrders + 1
'      vSalesOrders(lSalesOrders, 0) = cmbSon
'      vSalesOrders(lSalesOrders, 1) = vBookmark
'      vSalesOrders(lSalesOrders, 2) = iSoItems
      iBookMark = vBookmark
      iNewBookmark = vBookmark
      On Error Resume Next
      txtShp(1).SetFocus
   Else
      GetItems = 0
      MsgBox "No Items Available To Ship For That Sales Order.", _
         vbInformation, Caption
   End If
   Set RdoItm = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub CloseBoxes()   '(bNotFirstBox As Byte)
   'bNotFirstBox is not used
'   On Error GoTo DiaErr1
   Dim iList As Integer
   'Dim A As Integer
   
   'If bNotFirstBox Then A = 1 Else A = 2
   'For iList = A To 2
   For iList = 1 To 3
      lblitm(iList).Visible = False
      lblRev(iList).Visible = False
      lblSchDte(iList).Visible = False
      lblPrt(iList).Visible = False
      lblDsc(iList).Visible = False
      lblOrd(iList).Visible = False
      txtShp(iList).Visible = False
      lblCmt(iList).Visible = False
      lblUom(iList).Visible = False
      optShp(iList).Visible = False
      txtBox(iList).Visible = False
      lblQoh(iList).Visible = False
      cmdComments(iList).Visible = False
      z2(iList).Visible = False
   Next
'   lblitm(iList).Visible = False
'   lblRev(iList).Visible = False
'   lblPrt(iList).Visible = False
'   lblDsc(iList).Visible = False
'   lblOrd(iList).Visible = False
'   txtShp(iList).Visible = False
'   lblCmt(iList).Visible = False
'   lblUom(iList).Visible = False
'   optShp(iList).Visible = False
'   txtBox(iList).Visible = False
'   lblQoh(iList).Visible = False
'   cmdComments(iList).Visible = False
'   z2(iList).Visible = False
   
'   For iList = 1 To 2
   For iList = 1 To 3
      lblitm(iList) = ""
      lblRev(iList) = ""
      lblSchDte(iList) = ""
      lblPrt(iList) = ""
      lblDsc(iList) = ""
      lblOrd(iList) = ""
      txtShp(iList) = ""
      lblCmt(iList) = ""
      lblUom(iList) = ""
      txtBox(iList) = ""
      lblQoh(iList) = ""
      optShp(iList) = vbUnchecked
   Next
'   lblitm(iList) = ""
'   lblRev(iList) = ""
'   lblPrt(iList) = ""
'   lblDsc(iList) = ""
'   lblOrd(iList) = ""
'   txtShp(iList) = ""
'   lblCmt(iList) = ""
'   lblUom(iList) = ""
'   txtBox(iList) = ""
'   lblQoh(iList) = ""
'   optShp(iList) = vbUnchecked
   Exit Sub
'DiaErr1:
'   On Error GoTo 0
   
End Sub


Private Sub GetNextGroup()
   Dim iRow As Integer, a
   CloseBoxes 'False
   On Error Resume Next
   If iCurrPage < 1 Then iCurrPage = 1
   If iCurrPage > iMaxPages Then
      iCurrPage = iMaxPages
      cmdDn.Picture = Dsdn
      cmdDn.Enabled = False
   Else
      cmdDn.Picture = Endn
      cmdDn.Enabled = True
   End If
   If iCurrPage > 1 Then
      cmdUp.Picture = Enup
      cmdUp.Enabled = True
   Else
      cmdUp.Picture = Dsup
      cmdUp.Enabled = False
   End If
   lblPge = iCurrPage
   iNewBookmark = iBookMark + ((3 * iCurrPage) - 3)
   For iRow = iNewBookmark To iTotalItems
      a = a + 1
      If a > 3 Then Exit For
      lblitm(a).Visible = True
      lblRev(a).Visible = True
      lblSchDte(a).Visible = True
      lblPrt(a).Visible = True
      lblDsc(a).Visible = True
      lblOrd(a).Visible = True
      txtShp(a).Visible = True
      lblCmt(a).Visible = True
      optShp(a).Visible = True
      txtBox(a).Visible = True
      lblUom(a).Visible = True
      lblQoh(a).Visible = True
      cmdComments(iRow).Visible = True
      z2(a).Visible = True
      lblitm(a) = Format(vItems(iRow, 1), "#0")
      lblRev(a) = vItems(iRow, 2)
      lblSchDte(a) = vItems(iRow, PS_ITSCHED)
      lblPrt(a) = vItems(iRow, 3)
      lblDsc(a) = vItems(iRow, 4)
      lblOrd(a) = vItems(iRow, PS_SOQTY)
      lblCmt(a) = sPSComments(iRow)
      lblUom(a) = vItems(iRow, 7)
      txtShp(a) = vItems(iRow, PS_QTYTOSHIP)
      optShp(a) = vItems(iRow, 9)
      txtBox(a) = vItems(iRow, 12)
      lblQoh(a) = vItems(iRow, PS_QTYONHAND)
      lblSchDte(a) = vItems(iRow, PS_ITSCHED)
   Next
   txtShp(1).SetFocus
   
End Sub

'Private Sub GetThisSalesOrder()
'   Dim iRow As Integer
'   On Error Resume Next
'   Do Until vSalesOrders(iRow, 0) = cmbSon
'      iRow = iRow + 1
'   Loop
'   If vSalesOrders(iRow, 1) > 0 Then
'      iBookMark = vSalesOrders(iRow, 1)
'      iCurrPage = 1
'      iMaxPages = (vSalesOrders(iRow, 2) / 3) + 0.45
'      lblLst = iCurrPage
'      lblLst = iMaxPages
'      GetNextGroup
'   End If
'
'End Sub





Private Sub lblCmt_LostFocus(Index As Integer)
   lblCmt(Index) = CheckLen(lblCmt(Index), 255)
   sPSComments(iNewBookmark + (Index - 1)) = lblCmt(Index)
   
End Sub


Private Sub lblOrd_Click(Index As Integer)
   txtShp(Index) = lblOrd(Index)
   optShp(Index).Value = vbChecked
   vItems(iNewBookmark + (Index - 1), PS_QTYTOSHIP) = Val(txtShp(Index))
   vItems(iNewBookmark + (Index - 1), 9) = 1
   
End Sub


Private Sub optShp_Click(Index As Integer)
   If optShp(Index).Value = vbChecked Then
      If Val(txtShp(Index)) = 0 Then
         txtShp(Index) = lblOrd(Index)
         vItems(iNewBookmark + (Index - 1), PS_QTYTOSHIP) = Val(txtShp(Index))
         vItems(iNewBookmark + (Index - 1), 9) = 1
      End If
   Else
      If Val(txtShp(Index)) > 0 Then
         txtShp(Index) = "0.000"
         vItems(iNewBookmark + (Index - 1), PS_QTYTOSHIP) = "0.000"
         vItems(iNewBookmark + (Index - 1), 9) = 0
      End If
   End If
   
End Sub

Private Sub optShp_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optShp_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub


Private Sub optShp_LostFocus(Index As Integer)
   vItems(iNewBookmark + (Index - 1), 9) = optShp(Index).Value
   
End Sub


Private Sub optVew_Click()
   'never visible-view is showing
   
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtBox_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyCase KeyAscii
   
End Sub

Private Sub txtBox_LostFocus(Index As Integer)
   txtBox(Index) = CheckLen(txtBox(Index), 8)
   vItems(iNewBookmark + (Index - 1), 12) = txtBox(Index)
   
   
End Sub

Private Sub txtShp_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtShp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtShp_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub


Private Sub txtShp_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub


Private Sub txtShp_LostFocus(Index As Integer)
   txtShp(Index) = CheckLen(txtShp(Index), 9)
   ' Wouldn't allow overshipment
   ' If Val(txtShp(Index)) > Val(lblOrd(Index)) Then txtShp(Index) = lblOrd(Index)
   txtShp(Index) = Format(Abs(Val(txtShp(Index))), ES_QuantityDataFormat)
   vItems(iNewBookmark + (Index - 1), PS_QTYTOSHIP) = Val(txtShp(Index))
   If Val(txtShp(Index)) > 0 Then
      optShp(Index) = vbChecked
      vItems(iNewBookmark + (Index - 1), 9) = 1
   Else
      optShp(Index) = vbUnchecked
      vItems(iNewBookmark + (Index - 1), 9) = 0
   End If
   
End Sub

Private Sub CreatePackSlip()
   Dim RdoRev As ADODB.Recordset
   Dim iItems() As Integer
   Dim iCurrentItem As Integer
   
   
   Dim bResponse As Byte
   Dim bByte As Byte
   Dim iRow As Integer
   Dim a As Integer
   Dim iItemNumber As Integer
   
   Dim lSalesOrder As Long
   
   Dim sComments As String
   Dim sMsg As String
   Dim sRev As String
   
   Dim cSoPrice As Currency
   Dim cSoQty As Currency
   Dim cPsQty As Currency
   Dim cPrice As Currency
   Dim cQuantity As Currency
   'Dim bKanBanResponse As Byte
   
   
   bByte = 0
   
   'GetSoShipTo 2/13/02
   'Any selected?
   For iRow = 1 To iTotalItems
      If vItems(iRow, 9) = 1 Then
         bByte = 1
      End If
   Next
   'None found, bail out
   'On Error Resume Next
   On Error GoTo DiaErr2
   If bByte = 0 Then
      MsgBox "There Are No Items Selected.", vbInformation, Caption
      If (iTotalItems <> 0) Then txtShp(1).SetFocus
      Exit Sub
   End If
   iCurrentItem = 0
   bByte = 0
   For iRow = 1 To iTotalItems
      If vItems(iRow, 9) = 1 Then
         cPsQty = Val(vItems(iRow, PS_QTYTOSHIP))
         cSoQty = Val(vItems(iRow, PS_SOQTY))
         If cPsQty > cSoQty Then bByte = 1
      End If
   Next
   If bByte = 1 Then
      
      If (Not CheckAlloOverShip) Then
         sMsg = "There is At Least One Item Marked" & vbCrLf _
                & "For Overshipment. Cann't Continue?"
         MsgBox sMsg, vbCritical, Caption
         Exit Sub
      End If
      
      sMsg = "There is At Least One Item Marked" & vbCrLf _
             & "For Overshipment. Continue?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbNo Then
         CancelTrans
         Exit Sub
      End If
   End If
   
   Dim bOverQOH  As Boolean
   Dim ht As New HashTable
   Dim strErrCurPart As String
   bOverQOH = CheckOfQOH(ht, strErrCurPart)
   If bOverQOH = True Then
         MsgBox "Packslip Partnumber - " & strErrCurPart & " is picked over Quantity on Hand. Couldn't Create The Packing Slip.", vbExclamation, Caption
      Exit Sub
   End If
   
   
   ' Check pre Package allowed
   Dim RdoCom As ADODB.Recordset
   sSql = "SELECT * FROM ComnTable WHERE COALLOWPSPREPICKS = 1"
   
   strErrCurPart = ""
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCom, ES_STATIC)
   If bSqlRows Then
      Dim bOverPSQOH  As Boolean
      bOverPSQOH = CheckOfUnPrintedPS(lblSon, strErrCurPart)
      If bOverPSQOH = True Then
      
         MsgBox "Packslip Partnumber - " & strErrCurPart & " is picked over Quantity on Hand. " & vbCrLf _
               & "Have some Un-Printed Packslips. Couldn't Create The Packing Slip.", vbExclamation, Caption
         Exit Sub
      End If
   End If
   Set RdoCom = Nothing
   
   
    'If PrintingKanBanLabels(Compress(lblCst)) Then
    '    bKanBanResponse = MsgBox("Do You Want to Print the KanBan Label Now?", vbYesNoCancel, Caption)
    'End If
 
   
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   iItemNumber = GetLastItem()
   'Mark records as having a Pack Slip
   For iRow = 1 To iTotalItems
      If vItems(iRow, 9) = 1 Then
         'lSalesOrder = Val(Right(vItems(iRow, 0), PS_SOQTY))
         lSalesOrder = Val(vItems(iRow, 0))
         iItemNumber = iItemNumber + 1
         If optCmt.Value = vbChecked Then sComments = "" & Trim(vItems(iRow, 6))
         cPsQty = Val(vItems(iRow, PS_QTYTOSHIP))
         cSoQty = Val(vItems(iRow, PS_SOQTY))
         
         'Update SoitTable
         Err = 0
         'RdoCon.BeginTrans
         sSql = "UPDATE SoitTable SET ITPSNUMBER='" & lblPsl & "', ITPSITEM=" _
                & iItemNumber & ",ITQTY=" & cPsQty & " " _
                & "WHERE ITSO=" & lSalesOrder & " AND ITNUMBER=" _
                & vItems(iRow, 1) & " AND ITREV='" & vItems(iRow, 2) & "' "
         clsADOCon.ExecuteSql sSql ', rdExecDirect
         
         'Create PsitTable record
         cQuantity = Format(Val(vItems(iRow, PS_QTYTOSHIP)), ES_QuantityDataFormat)
         cPrice = Format(Val(vItems(iRow, 11)), ES_QuantityDataFormat)
         sSql = "INSERT PsitTable (PIPACKSLIP,PIITNO,PITYPE,PIQTY,PIPART," _
                & "PISONUMBER,PISOITEM,PISOREV,PISELLPRICE,PIBOX,PICOMMENTS) " _
                & "VALUES('" & lblPsl & "'," & iItemNumber & ",1," & cQuantity _
                & ",'" & sPartNumber(iRow) & "'," & lSalesOrder & "," & Val(vItems(iRow, 1)) _
                & ",'" & vItems(iRow, 2) & "'," & cPrice & ",'" & vItems(iRow, 12) _
                & "','" & sPSComments(iRow) & "')"
         clsADOCon.ExecuteSql sSql ', rdExecDirect
         
         iCurrentItem = iCurrentItem + 1
         ReDim Preserve iItems(1 To iCurrentItem) As Integer
         iItems(iCurrentItem) = iItemNumber
         'ReDim Preserve iItems(1 To iRow) As Integer
         'iItems(iRow) = iItemNumber
         
         
'         If Err = 0 Then
'            RdoCon.CommitTrans
'         Else
'            RdoCon.RollbackTrans
'            MsgBox "Couldn't Compete Pack Slip.", _
'               vbExclamation, Caption
'            Exit Sub
'         End If
         
         'If we need to split it, do it
         cPsQty = Val(vItems(iRow, PS_QTYTOSHIP))
         cSoQty = Val(vItems(iRow, PS_SOQTY))
         If cPsQty < cSoQty Then
            'sRev = Trim(vItems(iRow, 2))
            ' MM 5/16/2010 compiler error: Type mismatch in long to string.
            'sRev = GetNextSORevision(GetLastRevForSalesOrderItem(lSalesOrder, vItems(iRow, 1)))
            sRev = GetNextSORevision(GetLastRevForSalesOrderItem(CStr(lSalesOrder), CStr(vItems(iRow, 1))))
            
    
            sSql = "INSERT " & sTableDef & " SELECT * FROM SoitTable WHERE " _
                   & "ITSO=" & lSalesOrder & " AND ITNUMBER=" & Val(vItems(iRow, 1)) & " AND " _
                   & "ITREV='" & vItems(iRow, 2) & "' "
            clsADOCon.ExecuteSql sSql ', rdExecDirect

'combine with SoitTable update below
'            sSql = "UPDATE " & sTableDef & " SET ITQTY=" & cSoQty - cPsQty _
'                   & ",ITREV='" & sRev & "' WHERE ITSO=" & lSalesOrder & " And " _
'                   & "ITNUMBER=" & vItems(iRow, 1) & " AND ITREV='" & vItems(iRow, 2) & "' "
'            clsADOCon.ExecuteSQL sSql ', rdExecDirect
'
                        
            sSql = "UPDATE " & sTableDef & vbCrLf _
               & "SET ITQTY=" & cSoQty - cPsQty & "," & vbCrLf _
               & "ITREV='" & sRev & "'," & vbCrLf _
               & "ITACTUAL=NULL," & vbCrLf _
               & "ITPSNUMBER=''," & vbCrLf _
               & "ITPSITEM=0" & vbCrLf _
               & "WHERE ITSO=" & lSalesOrder & vbCrLf _
               & "AND ITNUMBER=" & vItems(iRow, 1) & vbCrLf _
               & "AND ITREV='" & vItems(iRow, 2) & "'"
            clsADOCon.ExecuteSql sSql ', rdExecDirect
            
            sSql = "INSERT SoitTable SELECT * FROM " & sTableDef & " WHERE " _
                   & "ITSO=" & lSalesOrder & " AND ITNUMBER=" & vItems(iRow, 1) & " AND " _
                   & "ITREV='" & sRev & "' "
            clsADOCon.ExecuteSql sSql ', rdExecDirect
            
'combine with above update.  Don't revise ITDOLLARS
'            sSql = "UPDATE SoitTable SET ITACTUAL=NULL,ITPSNUMBER=''," _
'                   & "ITDOLLARS=" & vItems(iRow, 11) & "," _
'                   & "ITPSITEM=0 WHERE ITSO=" & lSalesOrder & " AND " _
'                   & "ITNUMBER=" & vItems(iRow, 1) & " AND ITREV='" & sRev & "' "
'            clsADOCon.ExecuteSQL sSql ', rdExecDirect
            
            sSql = "DELETE FROM " & sTableDef & " WHERE " _
                   & "ITSO=" & lSalesOrder & " AND ITNUMBER=" & vItems(iRow, 1) & " AND " _
                   & "ITREV='" & sRev & "' "
            clsADOCon.ExecuteSql sSql ', rdExecDirect
         End If
      End If
   Next
   If iTotalItems > 0 Then
      sSql = "UPDATE PsitTable SET PISELLPRICE=" _
             & "ITDOLLARS FROM PsitTable,SoitTable WHERE " _
             & "(PISONUMBER=ITSO AND PISOITEM=ITNUMBER " _
             & "AND PISOREV=ITREV) AND PIPACKSLIP='" & Trim(lblPsl) & "'"
      clsADOCon.ExecuteSql sSql ', rdExecDirect
   End If
   
   clsADOCon.CommitTrans
   
   bPsSaved = True
   'If bKanBanResponse = vbYes Then
   '     For iRow = 1 To UBound(iItems)
   '        PrintKanBanLabel lblPsl, iItems(iRow)
   '     Next iRow
   'End If

   MsgBox "Packing Slip Items Were Recorded.", vbInformation, Caption
   Set RdoRev = Nothing
   
   Erase iItems
   
   Unload Me
   Exit Sub
   
DiaErr1:
   Resume DiaErr2
DiaErr2:
   On Error Resume Next
   bPsSaved = False
   MouseCursor 0
   clsADOCon.RollbackTrans
   MsgBox "Couldn't Create The Packing Slip.", vbExclamation, Caption
   
End Sub

Private Function CheckAlloOverShip() As Boolean
   Dim rdoPS As ADODB.Recordset
   
   sSql = "SELECT * FROM ComnTable WHERE COPSOVERSHIP = 1"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoPS, ES_KEYSET)
   If bSqlRows Then
      CheckAlloOverShip = True
      ClearResultSet rdoPS
   Else
      CheckAlloOverShip = False
   End If
   Set rdoPS = Nothing
   
End Function

Private Function GetLastItem() As Integer
   Dim RdoNxt As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT MAX(PIITNO) FROM PsitTable WHERE " _
          & "PIPACKSLIP='" & lblPsl & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoNxt, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoNxt.Fields(0)) Then
         If Val(RdoNxt.Fields(0)) > 0 Then _
                GetLastItem = RdoNxt.Fields(0)
      Else
         GetLastItem = 0
      End If
   Else
      GetLastItem = 0
   End If
   Set RdoNxt = Nothing
   Exit Function
   
DiaErr1:
   GetLastItem = 0
End Function

'Added 1/11/02 to correct Ship to
'Note that if there are more than one Sales Order...well

Private Sub GetSoShipTo()
   On Error GoTo DiaErr1
   Dim RdoSti As ADODB.Recordset
   sSql = "SELECT SONUMBER,SOSTNAME,SOSTADR WHERE " _
          & "SONUMBER=" & Val(Right(lblSon, SO_NUM_SIZE))
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSti, ES_FORWARD)
   If bSqlRows Then
      With RdoSti
         sStName = "" & Trim(!SOSTNAME)
         sStAdr = "" & Trim(!SOSTADR)
         ClearResultSet RdoSti
      End With
   End If
   Exit Sub
   
DiaErr1:
   On Error GoTo 0
   
End Sub

Private Function CheckSalesOrder() As Byte
   Dim RdoSon As ADODB.Recordset
   sSql = "SELECT SONUMBER,SOCUST,SOTYPE FROM SohdTable WHERE SOCUST='" _
          & Compress(lblCst) & "' AND SONUMBER=" & Val(lblSon)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSon, ES_FORWARD)
   If bSqlRows Then
      With RdoSon
         lblType = "" & Trim(!SOTYPE)
         CheckSalesOrder = 1
      End With
      ClearResultSet RdoSon
   Else
      CheckSalesOrder = 0
   End If
   Set RdoSon = Nothing
   
   
End Function

Public Function AddNewSOtoPS(ItemNum As Integer)
   
   Dim bRet As Boolean
   ' Add the Item to the end of the list
   bRet = AddNewSoItem(ItemNum)

   If (bRet = True) Then
      iTotalItems = iTotalItems + 1
      iMaxPages = iTotalItems \ 3 ' 3 items per page
      If (iTotalItems Mod 3 > 0) Then
         iMaxPages = iMaxPages + 1
         If (iMaxPages = iCurrPage) Then cmdDn.Enabled = True
      End If
      lblLst = iMaxPages
   End If
   
End Function

Private Function AddNewSoItem(ItemNum As Integer)
   
   Dim RdoSoit As ADODB.Recordset
   Dim TotItems As Integer
   
   TotItems = iTotalItems + 1
   
   sSql = "SELECT ITSO,ITNUMBER,ITREV,ITQTY,ITPART,ITDOLLARS,ITSCHED, " _
          & "ITCANCELED,ITCOMMENTS,PARTREF,PARTNUM,PADESC,PAUNITS," _
          & "PAQOH From SoitTable,PartTable WHERE ITSO=" & Val(lblSon) & " AND " _
          & "ITPART=PARTREF AND ITNUMBER = " & ItemNum
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSoit, ES_KEYSET)
   If bSqlRows Then
      With RdoSoit
         vItems(TotItems, 0) = "" & lblSon
         vItems(TotItems, 1) = Format(!ITNUMBER, "#0")
         vItems(TotItems, 2) = "" & Trim(!itrev)
         vItems(TotItems, 3) = "" & Trim(!PartNum)
         vItems(TotItems, 4) = "" & Trim(!PADESC)
         vItems(TotItems, PS_SOQTY) = Format(!ITQty, ES_QuantityDataFormat)
         ' vItems(iTotalItems, 6) = "" & Trim(!ITCOMMENTS)
         vItems(TotItems, 7) = "" & Trim(!PAUNITS)
         vItems(TotItems, PS_QTYTOSHIP) = "0.000"
         vItems(TotItems, 9) = 0
         vItems(TotItems, 10) = "" & Trim(!PartRef)
         sPartNumber(TotItems) = "" & Trim(!PartRef)
         vItems(TotItems, 11) = Format(!ITDOLLARS, ES_QuantityDataFormat)
         vItems(TotItems, 12) = ""
         vItems(TotItems, PS_QTYONHAND) = Format(!PAQOH, ES_QuantityDataFormat)
         vItems(TotItems, PS_ITSCHED) = Format(!itsched, "MM/dd/yy")
         ' Show if is the current window
         Dim iPage As Integer
         Dim iItemMod As Integer
         iPage = iTotalItems \ 3 ' 3 items per page
         iItemMod = iTotalItems Mod 3
         iItemMod = iItemMod + 1 ' next items
'         If (iItemMod > 0) Then
'            iPage = iPage + 1
'         End If
         
         If ((iPage + 1) = iCurrPage) Then
            lblitm(iItemMod).Visible = True
            lblRev(iItemMod).Visible = True
            lblSchDte(iItemMod).Visible = True
            lblPrt(iItemMod).Visible = True
            lblDsc(iItemMod).Visible = True
            lblOrd(iItemMod).Visible = True
            txtShp(iItemMod).Visible = True
            txtBox(iItemMod).Visible = True
            lblCmt(iItemMod).Visible = True
            optShp(iItemMod).Visible = True
            lblUom(iItemMod).Visible = True
            lblQoh(iItemMod).Visible = True
            cmdComments(iItemMod).Visible = True
            z2(iItemMod).Visible = True
            lblitm(iItemMod) = Format(!ITNUMBER, "#0")
            lblRev(iItemMod) = "" & Trim(!itrev)
            lblSchDte(iItemMod) = Format(!itsched, "MM/dd/yy")
            lblPrt(iItemMod) = "" & Trim(!PartNum)
            lblDsc(iItemMod) = "" & Trim(!PADESC)
            lblOrd(iItemMod) = Format(!ITQty, ES_QuantityDataFormat)
            txtShp(iItemMod) = "0.000"
            lblCmt(iItemMod) = "" & Trim(!ITCOMMENTS)
            lblUom(iItemMod) = "" & Trim(!PAUNITS)
            lblQoh(iItemMod) = "" & Trim(!PAQOH)
         End If
      End With
      ClearResultSet RdoSoit
      AddNewSoItem = True
   Else
      AddNewSoItem = False
   End If
   Set RdoSoit = Nothing
   
End Function


Private Function CheckAlloNewSOit()
   Dim RdoSoit As ADODB.Recordset
   
   sSql = "SELECT * FROM ComnTable WHERE COALLOWPSNEWSOITEM = 1"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSoit, ES_KEYSET)
   If bSqlRows Then
      cmdAddS.Visible = True
      ClearResultSet RdoSoit
   Else
      cmdAddS.Visible = False
   End If
   Set RdoSoit = Nothing
   
End Function

Private Function CheckOfUnPrintedPS(strSoNum As String, ByRef strErrCurPart As String) As Boolean

   Dim RdoChkPS As ADODB.Recordset
   Dim ht As New HashTable
   Dim bOverQOH  As Boolean
   Dim cPsQty As Currency
   Dim cTotPsQty As Currency
   Dim strCurPart As String
   
   On Error GoTo DiaErr
          
   bOverQOH = False
   
   sSql = "SELECT PIPART, SUM(PIQTY) TotQty FROM PSitTable, SoitTable,pshdtable where ITPART = PIPART " _
            & "AND PSNUMBER = PIPACKSLIP AND ITSO = PISONUMBER AND ITNUMBER = PISOITEM " _
            & "AND PISOREV = ITREV AND ITSO = '" & strSoNum & "' AND PSPRINTED = NULL " _
         & " GROUP BY PIPART"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChkPS, ES_KEYSET)
   If bSqlRows Then
   
      With RdoChkPS
         strCurPart = Trim(!PIPART)
         cPsQty = Val(!TotQty)
         
         If ht.Exists(strCurPart) Then
             cTotPsQty = CCur(ht(strCurPart)) + cPsQty
             ht(strCurPart) = CVar(cTotPsQty)
         Else
             ht.Add strCurPart, CVar(cPsQty)
         End If
      End With
      ClearResultSet RdoChkPS
      
      bOverQOH = CheckOfQOH(ht, strErrCurPart)
   
   End If
   
   CheckOfUnPrintedPS = bOverQOH
   Set RdoChkPS = Nothing
   Exit Function
DiaErr:
   sProcName = "CheckOfUnPrintedPS"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function

Private Function CheckOfQOH(ByRef ht As HashTable, ByRef strErrCurPart As String) As Boolean

   On Error GoTo DiaErr
   
   Dim bOverOH As Boolean
   Dim cQOH As Currency
   Dim cPsQty As Currency
   Dim cTotPsQty As Currency
   Dim iRow As Integer
   Dim strCurPart As String
   
   bOverOH = False
   
   For iRow = 1 To iTotalItems
      If vItems(iRow, 9) = 1 Then
         
         strCurPart = vItems(iRow, 10)
         cQOH = Val(vItems(iRow, PS_QTYONHAND))
         cPsQty = Val(vItems(iRow, PS_QTYTOSHIP))
         
         If ht.Exists(strCurPart) Then
             cTotPsQty = CCur(ht(strCurPart)) + cPsQty
             If (cTotPsQty > cQOH) Then
               strErrCurPart = strErrCurPart + " " + strCurPart
               bOverOH = True
             End If
             ht(strCurPart) = CVar(cTotPsQty)
         Else
             If (cPsQty > cQOH) Then
               strErrCurPart = strErrCurPart + " " + strCurPart
               bOverOH = True
             End If
             ht.Add strCurPart, CVar(cPsQty)
         End If
         
      End If
   Next
   
   ht.RemoveAll
   
   CheckOfQOH = bOverOH
   Exit Function
DiaErr:
   sProcName = "CheckOfQOH"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function
