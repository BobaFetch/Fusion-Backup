VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form LotEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit New Lots"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCert 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "User Lot (40)"
      Top             =   3240
      Width           =   3615
   End
   Begin VB.TextBox txtUnitCost 
      Height          =   285
      Left            =   1920
      TabIndex        =   74
      Tag             =   "1"
      Top             =   1800
      Width           =   795
   End
   Begin VB.CommandButton cmdDockLok 
      Height          =   495
      Left            =   6000
      Picture         =   "LotEdit.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   73
      ToolTipText     =   "Push Info to M-Files"
      Top             =   4440
      Width           =   615
   End
   Begin VB.ComboBox cboLoc 
      Height          =   315
      Left            =   1920
      TabIndex        =   5
      Top             =   3600
      Width           =   915
   End
   Begin VB.ComboBox cboExpirationDate 
      Height          =   315
      Left            =   4500
      TabIndex        =   6
      Tag             =   "4"
      Top             =   3600
      Width           =   1275
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "LotEdit.frx":05DA
      Style           =   1  'Graphical
      TabIndex        =   71
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "&Edit Org"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5880
      TabIndex        =   70
      ToolTipText     =   "Edit The First Lot (Original Lot)"
      Top             =   2880
      Width           =   875
   End
   Begin VB.TextBox txtThisLot 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   67
      TabStop         =   0   'False
      Tag             =   "1"
      ToolTipText     =   "Current Lot Quantity"
      Top             =   2220
      Width           =   975
   End
   Begin VB.CheckBox optMo 
      Caption         =   "From Mo (hidden)"
      Height          =   195
      Left            =   4440
      TabIndex        =   53
      Top             =   5520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CheckBox optPo 
      Caption         =   "From Po (hidden)"
      Height          =   195
      Left            =   2040
      TabIndex        =   39
      Top             =   5520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CheckBox optCreate 
      Height          =   255
      Left            =   5760
      TabIndex        =   38
      TabStop         =   0   'False
      ToolTipText     =   "Check To Create New Lots (Restricted)"
      Top             =   640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtLotQty 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5760
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Enter Quantity of New Lot"
      Top             =   1000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create"
      Height          =   315
      Left            =   5880
      TabIndex        =   2
      ToolTipText     =   "Create Another New Lot "
      Top             =   1380
      Visible         =   0   'False
      Width           =   875
   End
   Begin VB.TextBox txtCmt 
      Height          =   285
      Left            =   1440
      TabIndex        =   30
      Tag             =   "3"
      ToolTipText     =   "User Lot (40)"
      Top             =   5760
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox txtLong 
      Height          =   615
      Left            =   1920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Tag             =   "9"
      ToolTipText     =   "Comments (2048)"
      Top             =   2580
      Width           =   3615
   End
   Begin VB.CommandButton cmdSug 
      Cancel          =   -1  'True
      Caption         =   "S&uggest"
      Height          =   315
      Left            =   4800
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Suggest A Lot Number"
      Top             =   6240
      Visible         =   0   'False
      Width           =   875
   End
   Begin VB.TextBox txtHgt 
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Tag             =   "1"
      ToolTipText     =   "Mat Heght"
      Top             =   3960
      Width           =   915
   End
   Begin VB.TextBox txtLng 
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Tag             =   "1"
      ToolTipText     =   "Mat Length"
      Top             =   4320
      Width           =   915
   End
   Begin VB.TextBox txtWid 
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Tag             =   "1"
      ToolTipText     =   "Mat Width"
      Top             =   4680
      Width           =   915
   End
   Begin VB.TextBox txtHum 
      Height          =   285
      Left            =   4515
      TabIndex        =   8
      Tag             =   "3"
      ToolTipText     =   "Unit Of Measure (2)"
      Top             =   3960
      Width           =   435
   End
   Begin VB.TextBox txtLum 
      Height          =   285
      Left            =   4515
      TabIndex        =   10
      Tag             =   "3"
      ToolTipText     =   "Unit Of Measure (2)"
      Top             =   4320
      Width           =   435
   End
   Begin VB.TextBox txtWum 
      Height          =   285
      Left            =   4515
      TabIndex        =   12
      Tag             =   "3"
      ToolTipText     =   "Unit Of Measure (2)"
      Top             =   4680
      Width           =   435
   End
   Begin VB.TextBox txtLot 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "User Lot (40)"
      Top             =   2220
      Width           =   3615
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   5880
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6480
      Top             =   5640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5370
      FormDesignWidth =   7125
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cert / HT #"
      Height          =   255
      Index           =   21
      Left            =   285
      TabIndex        =   77
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Cost"
      Height          =   255
      Index           =   20
      Left            =   1080
      TabIndex        =   76
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label lblNewLot 
      Caption         =   "0"
      Height          =   255
      Left            =   3960
      TabIndex        =   75
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblExpirationDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Lot Expiration Date"
      Height          =   255
      Left            =   3000
      TabIndex        =   72
      Top             =   3660
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Original Lot"
      Height          =   255
      Index           =   19
      Left            =   480
      TabIndex        =   69
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label lblOrginalLot 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   68
      Top             =   7680
      Width           =   2175
   End
   Begin VB.Label lblLotRemaining 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4800
      TabIndex        =   66
      ToolTipText     =   "Remainder Of Original Lot"
      Top             =   280
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remains Of Orig Qty"
      Height          =   255
      Index           =   18
      Left            =   3120
      TabIndex        =   65
      Top             =   280
      Width           =   1815
   End
   Begin VB.Label lblEdit 
      BackStyle       =   0  'Transparent
      Caption         =   "Editing Lot 1 (Original Lot)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   285
      TabIndex        =   64
      Top             =   5040
      Width           =   4215
   End
   Begin VB.Label INVOHDACCT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4560
      TabIndex        =   63
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label INVEXPACCT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3480
      TabIndex        =   62
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label INVMATACCT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      TabIndex        =   61
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label INVLABACCT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   60
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label RUNHRS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4800
      TabIndex        =   59
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label RUNLBR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5640
      TabIndex        =   58
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label RUNEXP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3840
      TabIndex        =   57
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label RUNMATL 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   56
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label RUNOVHD 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      TabIndex        =   55
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label RUNCOST 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   54
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label MORUN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3600
      TabIndex        =   52
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label MONUMBER 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   51
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Stuff"
      Height          =   255
      Index           =   17
      Left            =   480
      TabIndex        =   50
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      Height          =   255
      Index           =   16
      Left            =   285
      TabIndex        =   49
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label VENDOR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   48
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label STDCOST 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3840
      TabIndex        =   47
      Top             =   6120
      Width           =   615
   End
   Begin VB.Label CreditAccount 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2640
      TabIndex        =   46
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label DebitAccount 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   45
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label INVACTIVITY 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   44
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Stuff"
      Height          =   255
      Index           =   15
      Left            =   480
      TabIndex        =   43
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label POREV 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      TabIndex        =   42
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label POITEM 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2640
      TabIndex        =   41
      Top             =   6120
      Width           =   615
   End
   Begin VB.Label PONUMBER 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   40
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Create Another Lot"
      Height          =   255
      Index           =   4
      Left            =   4080
      TabIndex        =   37
      Top             =   640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "This Lot Quantity"
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   36
      Top             =   1000
      Width           =   1335
   End
   Begin VB.Label lblQtyRemaining 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   35
      ToolTipText     =   "Remains In Original Lot"
      Top             =   640
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity Remaining"
      Height          =   255
      Index           =   1
      Left            =   280
      TabIndex        =   34
      Top             =   640
      Width           =   1695
   End
   Begin VB.Label lblOrigQty 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   33
      ToolTipText     =   "Beginning Quantity"
      Top             =   280
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Original Quantity"
      Height          =   255
      Index           =   0
      Left            =   280
      TabIndex        =   32
      Top             =   280
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Stuff under here V"
      Height          =   255
      Left            =   360
      TabIndex        =   31
      Top             =   5520
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5760
      TabIndex        =   29
      Top             =   6240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Height"
      Height          =   255
      Index           =   7
      Left            =   285
      TabIndex        =   27
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Of Measure"
      Height          =   255
      Index           =   8
      Left            =   3000
      TabIndex        =   26
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Length"
      Height          =   255
      Index           =   9
      Left            =   285
      TabIndex        =   25
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Of Measure"
      Height          =   255
      Index           =   10
      Left            =   3000
      TabIndex        =   24
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Width"
      Height          =   255
      Index           =   11
      Left            =   285
      TabIndex        =   23
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Of Measure"
      Height          =   255
      Index           =   12
      Left            =   3000
      TabIndex        =   22
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "System Lot Number"
      Height          =   255
      Index           =   2
      Left            =   280
      TabIndex        =   21
      Top             =   1000
      Width           =   1575
   End
   Begin VB.Label lblNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   20
      ToolTipText     =   "System Produced Lot Number Click To Set User Lot The Same"
      Top             =   1000
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Lot Number"
      Height          =   255
      Index           =   5
      Left            =   285
      TabIndex        =   19
      Top             =   2220
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lot Comments"
      Height          =   255
      Index           =   6
      Left            =   285
      TabIndex        =   18
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Index           =   14
      Left            =   4920
      TabIndex        =   17
      Top             =   7680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5760
      TabIndex        =   16
      Top             =   7680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   13
      Left            =   280
      TabIndex        =   15
      Top             =   1380
      Width           =   1575
   End
   Begin VB.Label lblPart 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   14
      Top             =   1380
      Width           =   2775
   End
End
Attribute VB_Name = "LotEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'4/28/05 Added Lot Locations
'5/9/05 Multiple lots (Purchase Orders)
'5/17/05 Added MO Runs (Multiple lots)
'6/15/05 Clarified creating new lots
'7/11/05 Added cmdFirst and ability to edit the first Lot
'4/17/06 Add MsgBox traps at beginning of Create New Lot
Option Explicit
Dim bOnLoad As Byte
Dim bFirstLot As Byte

Dim lCOUNTER As Long
Dim cOriginalQty As Currency
Dim cSelectedQty As Currency
Dim cRemainingQty As Currency
Dim sOldLot As String
Dim sNextLot As String
Dim sOriginalLot As String
Private LotsExpire As Boolean

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub GetFirstLot()
   Dim RdoFirst As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT LOTNUMBER,LOTUSERLOTID,LOTPARTREF," _
          & "LOTUNITCOST,LOTDATECOSTED,LOTADATE,LOTMATLENGTH," _
          & "LOTMATLENGTHUM,LOTMATHEIGHT,LOTMATHEIGHTHUM,LOTREMAININGQTY," _
          & "LOTMATWIDTH,LOTMATWIDTHHUM,LOTLOCATION,LOTCOMMENTS,LOINUMBER,LOIRECORD," _
          & "LOITYPE,LOTCERT FROM LohdTable,LoitTable WHERE " _
          & "(LOTNUMBER='" & Trim(lblOrginalLot) & " ' AND LOTNUMBER=LOINUMBER " _
          & "AND LOIRECORD=1)"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoFirst, ES_FORWARD)
   If bSqlRows Then
      With RdoFirst
         lblNumber = "" & Trim(!lotNumber)
         txtlot = "" & Trim(!LOTUSERLOTID)
         txtLong = "" & Trim(!LOTCOMMENTS)
         txtCert = "" & Trim(!LOTCERT)
'         txtLoc = "" & Trim(!LOTLOCATION)
         If Not IsNull(.Fields(6)) Then
            lblDate = "" & Format(!LotADate, "mm/dd/yy")
         Else
            lblDate = Format(GetServerDateTime, "mm/dd/yy")
         End If
         txtHgt = Format(!LOTMATHEIGHT, ES_QuantityDataFormat)
         txtLng = Format(!LOTMATLENGTH, ES_QuantityDataFormat)
         txtWid = Format(!LOTMATWIDTH, ES_QuantityDataFormat)
         If Val(txtHgt) > 0 Then txtHum = "" & Trim(!LOTMATHEIGHTHUM)
         If Val(txtLng) > 0 Then txtLum = "" & Trim(!LOTMATLENGTHUM)
         If Val(txtWid) > 0 Then txtWum = "" & Trim(!LOTMATWIDTHHUM)
      End With
   End If
   lblEdit = "Editing Lot 1 (Orginal Lot)"
   cmdFirst.Enabled = False
   bFirstLot = 1
   Set RdoFirst = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getfirstlot"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cboExpirationDate_DropDown()
   ShowCalendar Me
   
   'to avoid showing dropdown and selection ( this happens with a modal form)
   cboExpirationDate.ListIndex = 0
   'cboExpirationDate.SelLength = 0
   'cboExpirationDate.SelStart = 0
   txtHgt.SetFocus
   cboExpirationDate.SetFocus
   DoEvents
   cboExpirationDate.SelLength = 0
   'cboExpirationDate.SelLength = 0
   'cboExpirationDate.SelStart = 0
End Sub

Private Sub cboExpirationDate_LostFocus()
   cboExpirationDate = CheckDateYYYY(cboExpirationDate)
End Sub


Private Sub cboLoc_LostFocus()
    cboLoc = CheckLen(cboLoc, 4)
End Sub

Private Sub cmdCan_Click()
    'prevent duplicate clicks
    Dim priorCreateEnabled As Boolean
    priorCreateEnabled = cmdCreate.Enabled
    cmdCan.Enabled = False
    cmdCreate.Enabled = False
   
    'if this is a sheet part require dimensions and qty 1
    If IsThisASheetPart(Compress(lblPart)) Then
         If Val(txtLotQty) <> 1 Or Val(txtLng) <= 1 Or Val(txtHgt) <= 1 Then
         MsgBox "Sheet lots require quantity 1, length and width.", _
            vbOKOnly, Caption
         cmdCan.Enabled = True
         cmdCreate.Enabled = priorCreateEnabled
         Exit Sub
         End If
    End If

  
   If LotsExpire Then
      If Not IsDate(cboExpirationDate.Text) Then
         MsgBox "Valid lot expiration date required"
         cmdCan.Enabled = True
         cmdCreate.Enabled = priorCreateEnabled
         Exit Sub
      End If
      
   
      Dim daysInPast As Long
      Dim dt As String
      dt = cboExpirationDate.Text
      
      daysInPast = DateDiff("d", CDate(cboExpirationDate.Text), Now)
      If daysInPast > 30 Then
         MsgBox "You can't enter an expiration dates more than 30 days in the past"
         cmdCan.Enabled = True
         Exit Sub
      ElseIf daysInPast > 0 Then
         Select Case MsgBox("The expiration date is " & daysInPast & " days in the past.  Is that correct?", vbYesNo)
         Case vbYes
         Case Else
            cmdCan.Enabled = True
            cmdCreate.Enabled = priorCreateEnabled
            Exit Sub
         End Select
      End If
      
   End If
   
   If bFirstLot = 1 Then
      UpdateLot 1
   Else
      If cRemainingQty > Val(txtLotQty) Then
         MsgBox "Please Complete Lot Selection.", _
            vbInformation, Caption
      Else
         UpdateLot 1
      End If
   End If
   cmdCan.Enabled = True
   cmdCreate.Enabled = priorCreateEnabled
End Sub


Private Sub cmdCreate_Click()
   cmdCreate.Enabled = False
   cmdCan.Enabled = False
   Dim bResponse As Byte
   Dim sMsg As String
   
   If Val(txtLotQty) >= Val(lblQtyRemaining) Then
      MsgBox "The New Lot Quantity Must Be Less Than The" & vbCr _
         & "Remaining Quantity (Positive Original Value).", _
         vbOKOnly, Caption
      cmdCreate.Enabled = True
      cmdCan.Enabled = True
      Exit Sub
   End If
   If Val(txtLotQty) = 0 Then
      MsgBox "A New Lot Cannot Be Created With A Zero Quantity.", _
         vbOKOnly, Caption
      cmdCreate.Enabled = True
      cmdCan.Enabled = True
      Exit Sub
   End If
   
   'if this is a sheet part require dimensions and qty 1
   If IsThisASheetPart(Compress(lblPart)) Then
        If Val(txtLotQty) <> 1 Or Val(txtLng) <= 1 Or Val(txtHgt) <= 1 Then
        MsgBox "Sheet lots require quantity 1, length and width.", _
           vbOKOnly, Caption
        cmdCreate.Enabled = True
        cmdCan.Enabled = True
        Exit Sub
        End If
   End If
   
   sMsg = "Have You Finished Editing the Current Lot And " & vbCr _
          & "Do You Want To Create Another Lot ?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      If optPo.Value = vbChecked Then
         BuildLotsPoReceipts
      Else
         BuildLotsMORuns
      End If
   Else
      CancelTrans
   End If
   
   cmdCreate.Enabled = True
   cmdCan.Enabled = True
   
End Sub

Private Sub cmdDockLok_Click()
   
   Dim strSysLot As String
   Dim strUserLot As String
   Dim strPartNum As String
   Dim strVendor As String
   Dim strPONum As String
   
   Dim mfileInt As MfileIntegrator
   Set mfileInt = New MfileIntegrator

   Dim strIniPath As String
   strIniPath = App.Path & "\" & "MFileInit.ini"

   strSysLot = GetSectionEntry("MFILES_METADATA", "SYSLOT", strIniPath)
   strUserLot = GetSectionEntry("MFILES_METADATA", "USERLOT", strIniPath)
   strPartNum = GetSectionEntry("MFILES_METADATA", "PARTNUM", strIniPath)
   strVendor = GetSectionEntry("MFILES_METADATA", "VENDOR", strIniPath)
   strPONum = GetSectionEntry("MFILES_METADATA", "PONUM", strIniPath)

   If (sProgName = "Production") Then
      mfileInt.OpenXMLFile "Insert", "MO", "Manufacturing Orders", mfileInt.gsMOClassID, "Scan", lblNumber
      mfileInt.AddXMLMetaData "SystemLotNumber", strSysLot, lblNumber
      mfileInt.AddXMLMetaData "UserLotNumber", strUserLot, txtlot
      mfileInt.AddXMLMetaData "PartNumber", strPartNum, lblPart
   Else
      mfileInt.OpenXMLFile "Insert", "MaterialPO", "Material Purchase Orders", mfileInt.gsMOPOClassID, "Scan", lblNumber
      mfileInt.AddXMLMetaData "SystemLotNumber", strSysLot, lblNumber
      mfileInt.AddXMLMetaData "UserLotNumber", strUserLot, txtlot
      mfileInt.AddXMLMetaData "PartNumber", strPartNum, lblPart
      mfileInt.AddXMLMetaData "Vendor", strVendor, VENDOR
      mfileInt.AddXMLMetaData "PONumber", strPONum, PONUMBER
   End If
   
   mfileInt.CloseXMLFile
'   If mfileInt.SendXMLFileToMFile Then SysMsg "Record Successfully Indexed", True
   If Not mfileInt.SendXMLFileToMFile Then SysMsg "Record Indexing Failed", True

   Set mfileInt = Nothing
   
End Sub

Private Function GetSectionEntry(ByVal strSectionName As String, ByVal strEntry As String, ByVal strIniPath As String) As String
   
   Dim X As Long
   Dim sSection As String, sEntry As String, sDefault As String
   Dim sRetBuf As String, iLenBuf As Integer, sFileName As String
   Dim sValue As String

   On Error GoTo modErr1
   
   sSection = strSectionName
   sEntry = strEntry
   sDefault = ""
   sRetBuf = String(256, vbNull) '256 null characters
   iLenBuf = Len(sRetBuf)
   sFileName = strIniPath
   X = GetPrivateProfileString(sSection, sEntry, _
                     "", sRetBuf, iLenBuf, sFileName)
   sValue = Trim(Left$(sRetBuf, X))
   
   If sValue <> "" Then
      GetSectionEntry = sValue
   Else
      GetSectionEntry = ""
   End If
   
   Exit Function
   
modErr1:
   sProcName = "GetSectionEntry"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
End Function

Private Sub cmdFirst_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   sMsg = "Edits The Intitial Lot-" & vbCr _
          & "Use This Function Only When Finished Creating Lots.  " & vbCr _
          & "You Will Not Be Able To Create Any Additional Lots   " & vbCr _
          & "After Selecting This Feature. Do You Wish To Continue?."
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbNo Then
      'do it
      optCreate.Enabled = False
      txtLotQty.Enabled = False
      cmdCreate.Visible = False
      GetFirstLot
   Else
      CancelTrans
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 933
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdSug_Click()
   Dim sUserLot As String
   If Len(Trim(lblPart)) > 26 Then _
          sUserLot = Left$(lblPart, 26) _
          Else sUserLot = Trim(lblPart)
   txtlot = sUserLot & "-" & Format(ES_SYSDATE, "yymmdd") _
            & "-" & lblTime & GetLotTrailer()
   sOldLot = txtlot
   
End Sub




Private Sub Command1_Click()
    Dim doclok As DocLokIntegrator
    
    Set doclok = New DocLokIntegrator
    
    If doclok.Installed Then
        doclok.OpenXMLFile "Attach", "Material Purchase Order"
        doclok.AddXMLIndex "System Lot Number", lblNumber
        doclok.AddXMLIndex "User Lot Number", txtlot
    
        doclok.CloseXMLFile
        If doclok.SendXMLFileToDocLok Then MsgBox "record indexed"
        
    Else
        MsgBox "DocumentLok is not installed"
    End If
    Set doclok = Nothing
End Sub

Private Sub Form_Activate()
   If bOnLoad = 1 Then
      bOnLoad = 0
      GetLotParameters
      MdiSect.Enabled = False
      sOriginalLot = lblNumber
      lblOrginalLot = lblNumber
      lblQtyRemaining = lblOrigQty
      cOriginalQty = Val(lblOrigQty)
      cRemainingQty = cOriginalQty
      txtThisLot = lblOrigQty
      lblLotRemaining = Format(cRemainingQty, ES_PurchasedDataFormat)
      txtLotQty = Format(cOriginalQty, ES_PurchasedDataFormat)
      sOldLot = Trim(txtlot)
      lCOUNTER = Val(INVACTIVITY)
      If cOriginalQty > 1 Then
         If optPo.Value = vbChecked Or optMo.Value = vbChecked Then
            optCreate.Visible = True
            txtLotQty.Visible = True
            'cmdCreate.Visible = True
            Z1(4).Visible = True
         End If
      End If
      bFirstLot = 1
      If DocumentLokEnabled Then cmdDockLok.Visible = True Else cmdDockLok.Visible = False
      If (lblNewLot = "1") Then txtUnitCost.Enabled = True
'      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   Move 1000, 1500
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   MdiSect.Enabled = True
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set LotEdit = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtLong.Visible = True
   txtThisLot.BackColor = Es_TextDisabled
   
End Sub




Private Sub lblNumber_Change()
   'Me.ReadLotData
End Sub

Private Sub optCreate_Click()
   On Error Resume Next
   If Val(lblQtyRemaining) > 0 Then
      If optCreate.Value = vbChecked Then
         txtLotQty.Enabled = True
         txtLotQty.SetFocus
         cmdCreate.Visible = True
      Else
         txtLotQty.Enabled = False
         cmdCreate.Visible = False
      End If
   Else
      cmdCreate.Visible = False
   End If
   
End Sub

Private Sub optCreate_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub



Private Sub OptPo_Click()
   'Never visible - from po
   
End Sub

Private Sub txtHgt_GotFocus()
   cboExpirationDate.SelLength = 0
End Sub

Private Sub txtHgt_LostFocus()
   txtHgt = CheckLen(txtHgt, 8)
   txtHgt = Format(Abs(Val(txtHgt)), ES_QuantityDataFormat)
   
End Sub


Private Sub txtHum_LostFocus()
   txtHum = CheckLen(txtHum, 2)
   If Trim(txtLum) = "" Then txtLum = txtHum
   If Trim(txtWum) = "" Then txtWum = txtHum
   
End Sub


Private Sub txtLng_LostFocus()
   txtLng = CheckLen(txtLng, 8)
   txtLng = Format(Abs(Val(txtLng)), ES_QuantityDataFormat)
   
End Sub


'Private Sub txtLoc_LostFocus()
'   txtLoc = CheckLen(txtLoc, 4)
'
'End Sub


Private Sub txtLong_LostFocus()
   txtLong = CheckLen(txtLong, 2048)
   txtLong = StrCase(txtLong, ES_FIRSTWORD)
   
End Sub


Private Sub txtlot_LostFocus()
   'Dim bByte As Byte
   txtlot = CheckLen(txtlot, 40)
   If Len(Trim(txtlot)) < 5 Then
      txtlot = "" 'sOldLot
      MsgBox "Requires At Least (5 chars).", _
         vbInformation
   Else
'      bByte = GetUserLotID(Trim(txtlot))
'      If bByte = 1 Then txtlot = sOldLot
      If IsUserLotIDInUse() Then
         txtlot = sOldLot
      End If
   End If
   sOldLot = txtlot
   
End Sub


Private Sub txtLotQty_LostFocus()
   txtLotQty = Format(Abs(Val(txtLotQty)), ES_PurchasedDataFormat)
   If Val(txtLotQty) > Val(lblQtyRemaining) Or Val(txtLotQty) = 0 Then
      Beep
      txtLotQty = lblQtyRemaining
   End If
   
End Sub


Private Sub txtLum_LostFocus()
   txtLum = CheckLen(txtLum, 2)
   
End Sub


Private Sub txtWid_LostFocus()
   txtWid = CheckLen(txtWid, 8)
   txtWid = Format(Abs(Val(txtWid)), ES_QuantityDataFormat)
   
End Sub


Private Sub txtWum_LostFocus()
   txtWum = CheckLen(txtWum, 2)
   
End Sub



Private Sub UpdateLot(bUnLoad As Byte)
   MouseCursor 13
'   Dim sComment As String
'   sComment = Trim(txtLong)
   'On Error Resume Next
'   If Trim(txtlot) <> "" Then
'      sSql = "UPDATE LohdTable SET LOTUSERLOTID='" & Trim(txtlot) & "'," _
'             & "LOTLOCATION='" & Trim(txtLoc) & "'," _
'             & "LOTCOMMENTS='" & Trim(txtLong) & "'," _
'             & "LOTMATLENGTH=" & Val(txtLng) & "," _
'             & "LOTMATLENGTHUM='" & txtLum & "'," _
'             & "LOTMATHEIGHT=" & Val(txtHgt) & "," _
'             & "LOTMATHEIGHTHUM='" & txtHum & "'," _
'             & "LOTMATWIDTH=" & Val(txtWid) & "," _
'             & "LOTMATWIDTHHUM='" & txtWum & "' " & vbCrLf
'   Else
'      sSql = "UPDATE LohdTable SET " _
'             & "LOTLOCATION='" & Trim(txtLoc) & "'," _
'             & "LOTCOMMENTS='" & sComment & "'," _
'             & "LOTMATLENGTH=" & Val(txtLng) & "," _
'             & "LOTMATLENGTHUM='" & txtLum & "'," _
'             & "LOTMATHEIGHT=" & Val(txtHgt) & "," _
'             & "LOTMATHEIGHTHUM='" & txtHum & "'," _
'             & "LOTMATWIDTH=" & Val(txtWid) & "," _
'             & "LOTMATWIDTHHUM='" & txtWum & "' " & vbCrLf
'   End If
   
   clsADOCon.ADOErrNum = 0
   sSql = "UPDATE LohdTable" & vbCrLf _
      & "SET "
   If Trim(txtlot) <> "" Then
      sSql = sSql & "LOTUSERLOTID = '" & Trim(txtlot) & "'," & vbCrLf
   End If
   If LotsExpire And Trim(cboExpirationDate.Text) <> "" Then
      sSql = sSql & "LOTEXPIRESON = '" & cboExpirationDate.Text & "'," & vbCrLf
   End If
   
'    ' if sheet inventory, convert from sheets to square inches
'    Dim units As Currency
'    Dim unitCost As Currency
'    If IsThisASheetPart(Compress(lblPart)) And IsNumeric(txtLong.Text) And IsNumeric(txtHgt.Text) Then
'        units = CCur(txtLong.Text) * CCur(txtHgt.Text)
'        If units > 1 Then
'            unitCost = Round(CCur(txtLotQty) / units, 4)
'            sSql = sSql & "LOTUNITCOST = " & CStr(unitCost) & "," & vbCrLf _
'                & "LOTORIGINALQTY = " & CStr(units) & "," & vbCrLf _
'                & "LOTREMAININGQTY = " & CStr(units) & "," & vbCrLf
'        End If
'    End If
'
   
   sSql = sSql & "LOTLOCATION = '" & Trim(cboLoc) & "'," & vbCrLf _
      & "LOTCOMMENTS = '" & Trim(txtLong) & "'," & vbCrLf _
      & "LOTCERT = '" & Trim(txtCert) & "'," & vbCrLf _
      & "LOTMATLENGTH = " & Val(txtLng) & "," & vbCrLf _
      & "LOTMATLENGTHUM = '" & txtLum & "'," & vbCrLf _
      & "LOTMATHEIGHT = " & Val(txtHgt) & "," & vbCrLf _
      & "LOTMATHEIGHTHUM = '" & txtHum & "'," & vbCrLf _
      & "LOTMATWIDTH = " & Val(txtWid) & "," & vbCrLf _
      & "LOTMATWIDTHHUM = '" & txtWum & "' " & vbCrLf

   
   sSql = sSql & "WHERE LOTNUMBER='" & lblNumber & "'"
   clsADOCon.ExecuteSql sSql
   
    'add height, width
    clsADOCon.ADOErrNum = 0
    sSql = "UPDATE LoitTable" & vbCrLf _
      & "SET " _
      & "LOILENGTH = " & Val(txtLng) & "," & vbCrLf _
      & "LOIHEIGHT = " & Val(txtHgt) & vbCrLf _
      & "WHERE LOINUMBER='" & lblNumber & "' AND LOIRECORD=1"
   clsADOCon.ExecuteSql sSql
   
   If (lblNewLot = "1") Then UpdateUnitCost
   
   MouseCursor 0
   If bUnLoad = 1 Then
      If clsADOCon.ADOErrNum = 0 Then
         'MsgBox "The Lot Was Successfully Edited.", _
         '   vbInformation, Caption
      Else
         MsgBox "The Lot Was Not Successfully Edited.", _
            vbExclamation, Caption
      End If
      
    ' if this is a sheet part and using sheet inventory mgmt, then update data to reflect
    ' inventory units rather than purchasing units
     If IsThisASheetPart(Compress(lblPart)) Then
         sSql = "exec SheetUnitConversion '" & lblNumber & "'"
         clsADOCon.ExecuteSql sSql
     End If
      Unload Me
   End If
   
End Sub

Private Sub BuildLotsPoReceipts()
   Static bedit As Byte
   Static bLot As Byte
   Dim bLotType As Byte
   Static lNextActivity As Long
   
   Dim sPon As String
   'UpdateLot 0     'redundant
   
   On Error Resume Next
   UpdateLot 0
   cSelectedQty = Val(txtLotQty)
   cRemainingQty = cRemainingQty - cSelectedQty
   lblQtyRemaining = Format(cRemainingQty, ES_PurchasedDataFormat)
   txtLotQty = lblQtyRemaining
   If lNextActivity = 0 Then lNextActivity = Val(INVACTIVITY)
   lNextActivity = lNextActivity + 1
   'Update old Rows
   ' Change the qty from cRemainingQty -> cSelectedQty
   sSql = "UPDATE LohdTable SET LOTORIGINALQTY=" & cSelectedQty & "," _
          & "LOTREMAININGQTY=" & cSelectedQty & ", LOTTOTMATL=" & (Val(STDCOST) * cSelectedQty) & " WHERE LOTNUMBER='" _
          & sOriginalLot & "'"
   clsADOCon.ExecuteSql sSql
   
   sSql = "UPDATE LoitTable SET LOIQUANTITY=" & cSelectedQty & " " _
          & "WHERE (LOINUMBER='" & sOriginalLot & "' AND LOIRECORD=1)"
   clsADOCon.ExecuteSql sSql
      
   sSql = "UPDATE InvaTable SET INAQTY=" & cSelectedQty & "," _
          & "INPQTY=" & cSelectedQty & ", INTOTMATL=" & (Val(STDCOST) * cSelectedQty) & " WHERE (INPART='" _
          & Compress(lblPart) & "' AND INNUMBER=" & lCOUNTER & ")"
   clsADOCon.ExecuteSql sSql
   
    ' if using sheet inventory, update the original lot, which is done, to use inventory units
    If IsThisASheetPart(Compress(lblPart)) Then
        sSql = "exec SheetUnitConversion '" & sOriginalLot & "'"
        clsADOCon.ExecuteSql sSql
    End If

   txtlot = Trim(txtlot)
   If bLot = 0 Then bLot = 64
   If bLot > 64 Then
      If Right$(txtlot, 2) = "-" & Chr$(64) Then txtlot = Left$(txtlot, Len(txtlot) - 2)
   End If
   bLot = bLot + 1
   sNextLot = GetNextLotNumber()
   lblNumber = sNextLot
   txtlot = sOldLot & "-" & Chr$(bLot)
   Dim totalCost As Currency
'   If IsThisASheetPart(Compress(lblPart)) Then
'      totalCost = Val(STDCOST) * cRemainingQty
'   Else
'      totalCost = Val(STDCOST) * cRemainingQty
'   End If
   totalCost = Val(STDCOST) * cRemainingQty

   'Build New Rows
   sSql = "INSERT INTO LohdTable (LOTNUMBER,LOTUSERLOTID," _
          & "LOTPARTREF,LOTPDATE,LOTADATE,LOTORIGINALQTY,LOTREMAININGQTY," _
          & "LOTPO,LOTPOITEM,LOTPOITEMREV,LOTUNITCOST,LOTTOTMATL,LOTUSER) " _
          & "VALUES('" _
          & lblNumber & "','" & txtlot & "','" & Compress(lblPart) _
          & "','" & lblDate & "','" & lblDate & "'," & cRemainingQty & "," & cRemainingQty _
          & "," & Val(PONUMBER) & "," & Val(POITEM) & ",'" _
          & POREV & "'," & Val(STDCOST) & ",cast(" & totalCost & " as decimal(12,4))" & ",'" & sInitials & "')"
   clsADOCon.ExecuteSql sSql
   
   sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
          & "LOITYPE,LOIPARTREF,LOIPDATE,LOIADATE,LOIQUANTITY," _
          & "LOIPONUMBER,LOIPOITEM,LOIPOREV,LOIVENDOR," _
          & "LOIACTIVITY,LOICOMMENT,LOIUSER) " _
          & "VALUES('" _
          & lblNumber & "',1,15,'" & Compress(lblPart) _
          & "','" & lblDate & "','" & lblDate & "'," & cRemainingQty _
          & "," & Val(PONUMBER) & "," & Val(POITEM) & ",'" _
          & POREV & "','" & VENDOR & "'," _
          & lNextActivity & ",'Purchase Order Receipt','" & sInitials & "')"
   clsADOCon.ExecuteSql sSql
   
   'Inva
   sPon = "PO " & PONUMBER & "-Item " & POITEM & POREV
   sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2," _
          & "INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT,INPONUMBER," _
          & "INPOITEM,INPOREV,INPDATE,INADATE,INNUMBER,INLOTNUMBER,INUSER, INTOTMATL) " _
          & "VALUES(15,'" & Compress(lblPart) & "'," & "'PO RECEIPT','" _
          & sPon & "'," & cRemainingQty & "," & cRemainingQty & "," _
          & Val(STDCOST) & ",'" & CreditAccount & "','" & DebitAccount _
          & "'," & Val(PONUMBER) & "," & Val(POITEM) _
          & ",'" & POREV & "','" & lblDate & "','" & lblDate & "'," & lNextActivity & ",'" _
          & lblNumber & "','" & sInitials & "', " & (Val(STDCOST) * cRemainingQty) & ")"
   clsADOCon.ExecuteSql sSql
   UpdateWipColumns lNextActivity
   
'don't do this yet.  it happens when you hit the done button
'    ' if using sheet inventory, update to use inventory units
'    If IsThisASheetPart(Compress(lblPart)) Then
'        sSql = "exec SheetUnitConversion '" & lblNumber & "'"
'        clsADOCon.ExecuteSQL sSql
'    End If
'
   
   If cRemainingQty < 2 Then
      optCreate.Value = vbUnchecked
      optCreate.Enabled = False
      txtLotQty.Enabled = False
      cmdCreate.Visible = False
   End If
   If bedit = 0 Then bedit = 1
   bFirstLot = 0
   optCreate.Value = vbUnchecked
   txtLotQty.Enabled = False
   cmdCreate.Visible = False
   cmdFirst.Enabled = True
   bedit = bedit + 1
   lblEdit = "Editing Lot " & Trim$(bedit) & " Qty: " & Trim$(cSelectedQty)
   lblLotRemaining = Format(cRemainingQty, ES_PurchasedDataFormat)
   txtThisLot = Format(cSelectedQty, ES_PurchasedDataFormat)
   
   lCOUNTER = lNextActivity ' we need to set this becuase otherwise it will update the wrong one when the lot splits
   
   ' 10/29/2009 - Change to the current lot
   sOriginalLot = lblNumber
End Sub

Private Sub BuildLotsMORuns()
   Static bLot As Byte
   Static bedit As Byte
   Dim bLotType As Byte
   Static lNextActivity As Long
   Dim sPon As String
   UpdateLot 0
   
   On Error Resume Next
   UpdateLot 0
   cSelectedQty = Val(txtLotQty)
   cRemainingQty = cRemainingQty - cSelectedQty
   lblQtyRemaining = Format(cRemainingQty, ES_PurchasedDataFormat)
   txtLotQty = lblQtyRemaining
   If lNextActivity = 0 Then lNextActivity = Val(INVACTIVITY)
   lNextActivity = lNextActivity + 1
   'Update old Rows
   sSql = "UPDATE LohdTable SET LOTORIGINALQTY=" & cRemainingQty & "," _
          & "LOTREMAININGQTY=" & cRemainingQty & " WHERE LOTNUMBER='" _
          & sOriginalLot & "'"
   clsADOCon.ExecuteSql sSql
   
   sSql = "UPDATE LoitTable SET LOIQUANTITY=" & cRemainingQty & " " _
          & "WHERE (LOINUMBER='" & sOriginalLot & "' AND LOIRECORD=1)"
   clsADOCon.ExecuteSql sSql
   
   sSql = "UPDATE InvaTable SET INAQTY=" & cRemainingQty & "," _
          & "INPQTY=" & cRemainingQty & " WHERE (INPART='" _
          & Compress(lblPart) & "' AND INNUMBER=" & lCOUNTER & ")"
   clsADOCon.ExecuteSql sSql
   
   txtlot = Trim(txtlot)
   If bLot = 0 Then bLot = 64
   If bLot > 64 Then
      If Right$(txtlot, 2) = "-" & Chr$(64) Then txtlot = Left$(txtlot, Len(txtlot) - 2)
   End If
   bLot = bLot + 1
   sNextLot = GetNextLotNumber()
   lblNumber = sNextLot
   txtlot = sOldLot & "-" & Chr$(bLot)
   
   'Build New Rows
   
   sSql = "INSERT INTO LohdTable (LOTNUMBER,LOTUSERLOTID," _
          & "LOTPARTREF,LOTPDATE,LOTORIGINALQTY,LOTREMAININGQTY," _
          & "LOTUNITCOST,LOTMOPARTREF,LOTMORUNNO," _
          & "LOTTOTMATL,LOTTOTLABOR,LOTTOTEXP,LOTTOTOH,LOTTOTHRS) " _
          & "VALUES('" _
          & lblNumber & "','" & txtlot _
          & "','" & Compress(lblPart) & "','" & lblDate & "'," & cSelectedQty & "," & cSelectedQty _
          & "," & Val(RUNCOST) & ",'" & Compress(lblPart) & "'," _
          & Val(MORUN) & "," & Val(RUNMATL) & "," & Val(RUNLBR) & "," _
          & Val(RUNEXP) & "," & Val(RUNOVHD) & "," & Val(RUNHRS) & ")"
   clsADOCon.ExecuteSql sSql
   
   sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
          & "LOITYPE,LOIPARTREF,LOIPDATE,LOIQUANTITY," _
          & "LOIMOPARTREF,LOIMORUNNO,LOIACTIVITY,LOICOMMENT) " _
          & "VALUES('" _
          & lblNumber & "',1,6,'" & Compress(lblPart) _
          & "','" & lblDate & "'," & cSelectedQty _
          & ",'" & Trim(MONUMBER) & "'," & Val(MORUN) & "," _
          & lNextActivity & ",'MO Run Completion')"
   clsADOCon.ExecuteSql sSql
   
   'Inva
   sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2," _
          & "INPDATE,INADATE,INAQTY,INAMT,INCREDITACCT,INDEBITACCT," _
          & "INMOPART,INMORUN,INTOTMATL,INTOTLABOR,INTOTEXP," _
          & "INTOTOH,INTOTHRS,INWIPLABACCT,INWIPMATACCT," _
          & "INWIPOHDACCT,INWIPEXPACCT,INNUMBER,INLOTNUMBER,INUSER) " _
          & "VALUES(6,'" & Compress(lblPart) & "','COMPLETED RUN'," _
          & "'RUN " & MORUN & "'," _
          & "'" & lblDate & "','" & lblDate & "'," & cSelectedQty & "," _
          & Val(RUNCOST) & ",'" & CreditAccount & "','" & DebitAccount & "','" _
          & Compress(lblPart) & "'," & Val(MORUN) & ","
   sSql = sSql & Val(RUNMATL) & "," & Val(RUNLBR) & "," _
          & Val(RUNEXP) & "," & Val(RUNOVHD) & "," & Val(RUNHRS) & ",'" _
          & INVLABACCT & "','" & INVMATACCT & "','" _
          & INVEXPACCT & "','" & INVOHDACCT & "'," & lNextActivity & ",'" _
          & lblNumber & "','" & sInitials & "')"
   clsADOCon.ExecuteSql sSql
   UpdateWipColumns lNextActivity
   
   If cRemainingQty <= 1 Then
      optCreate.Value = vbUnchecked
      optCreate.Enabled = False
      txtLotQty.Enabled = False
      cmdCreate.Visible = False
   End If
   bFirstLot = 0
   optCreate.Value = vbUnchecked
   txtLotQty.Enabled = False
   cmdCreate.Visible = False
   cmdFirst.Enabled = True
   If bedit = 0 Then bedit = 1
   bedit = bedit + 1
   lblEdit = "Editing Lot " & Trim$(bedit) & " Qty: " & Trim$(cSelectedQty)
   lblLotRemaining = Format(cRemainingQty, ES_PurchasedDataFormat)
   txtThisLot = Format(cSelectedQty, ES_PurchasedDataFormat)
   
   lCOUNTER = lNextActivity ' we need to set this becuase otherwise it will update the wrong one when the lot splits
   
End Sub

Private Sub GetLotParameters()
   Dim rdo As ADODB.Recordset
   Dim sDefaultLoc As String
   
   sSql = "SELECT PALOTSEXPIRE, PALOCATION from PartTable where PARTREF = '" & Compress(lblPart) & "'"
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      With rdo
         LotsExpire = IIf(!PALOTSEXPIRE = 0, False, True)
         sDefaultLoc = Trim("" & !PALOCATION)
      End With
   End If
   rdo.Close
   lblExpirationDate.Visible = LotsExpire
   cboExpirationDate.Visible = LotsExpire
   
   'if this is a new lot, set location = part default location
   If Me.Caption = "Edit New Lots" Then
      sSql = "update LohdTable set LOTLOCATION = '" & sDefaultLoc & "'" & vbCrLf _
         & "where  LOTNUMBER = '" & lblNumber & "'"
      clsADOCon.ExecuteSql sSql
   End If
   Set rdo = Nothing
   
   'get information from lot
   Dim S As String
   Dim sLotLoc As String
   sSql = "SELECT * from LohdTable where LOTNUMBER = '" & lblNumber & "'"
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      S = "" & rdo!LOTEXPIRESON
      sLotLoc = Trim("" & rdo!LOTLOCATION)
   End If
   If LotsExpire Then
      If Len(S) > 0 Then
         cboExpirationDate = S
      End If
   End If
   Set rdo = Nothing
   
   'populate location dropdown
   sSql = "select '" & sDefaultLoc & "' as Location" & vbCrLf _
      & "union" & vbCrLf _
      & "select distinct rtrim(LOTLOCATION) as location" & vbCrLf _
      & "from LohdTable where LOTPARTREF = '" & Compress(lblPart) & "'"
   LoadComboBox cboLoc, -1
   cboLoc = sLotLoc
   
End Sub

Public Sub ReadExistingMoData()
   'call this to load controls with data from current lot (number in lblNumber)
         
   Dim rdo As ADODB.Recordset
   sSql = "select *" & vbCrLf _
      & "from LohdTable" & vbCrLf _
      & "join InvaTable on INLOTNUMBER = LOTNUMBER and INTYPE = " & IATYPE_MoCompletion & vbCrLf _
      & "where LOTNUMBER = '" & lblNumber & "'"
      
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      With rdo
'         LotEdit.optMo.Value = vbChecked
'         LotEdit.DebitAccount = sDebitAcct
'         LotEdit.CreditAccount = sDebitAcct
'         LotEdit.MONUMBER = Compress(cmbPrt)
'         LotEdit.MORUN = cmbRun
'         LotEdit.INVACTIVITY = lCOUNTER
'         LotEdit.lblOrigQty = cComQty
   
         optMo.Value = vbChecked
         DebitAccount = Trim(!INDEBITACCT)
         CreditAccount = Trim(!INCREDITACCT)
         MONUMBER = Trim(!LOTMOPARTREF)
         MORUN = !LOTMORUNNO
         INVACTIVITY = !innumber
         lblOrigQty = !INAQTY
   
'         LotEdit.RUNCOST = cRunCost
'         LotEdit.RUNOVHD = cRunOvHd
'         LotEdit.RUNMATL = cRunMatl
'         LotEdit.RUNEXP = cRunExp
'         LotEdit.RUNHRS = cRunHours
'         LotEdit.RUNLBR = cRunLabor
   
         RUNCOST = !INAMT
         RUNOVHD = !INTOTOH
         RUNMATL = !INTOTMATL
         RUNEXP = !INTOTEXP
         RUNHRS = !INTOTHRS
         RUNLBR = !INTOTLABOR
   
'         LotEdit.txtLong = "MO Completion"
'         LotEdit.txtlot = "MO CO-" & cmbPrt & " Run " & Trim(Val(cmbRun))
         txtLong = "MO Completion"
         txtlot = "" & Trim(!LOTUSERLOTID)  '"MO CO-" & Trim(!LOTMOPARTREF) & " Run " & !LOTMORUNNO
'
'         LotEdit.INVLABACCT = sInvLabAcct
'         LotEdit.INVMATACCT = sInvMatAcct
'         LotEdit.INVEXPACCT = sInvExpAcct
'         LotEdit.INVOHDACCT = sInvOhdAcct

         INVLABACCT = Trim(!INWIPLABACCT)
         INVMATACCT = Trim(!INWIPMATACCT)
         INVEXPACCT = Trim(!INWIPEXPACCT)
         INVOHDACCT = Trim(!INWIPOHDACCT)

'         LotEdit.lblPart = cmbPrt
'         LotEdit.lblDate = Format(ES_SYSDATE, "mm/dd/yy")
'         LotEdit.lblTime = Format(ES_SYSDATE, "hh:mm")
'         LotEdit.lblNumber = sLotNumber
         
         LotEdit.lblPart = Trim(!LOTMOPARTREF)
         LotEdit.lblDate = Format(GetServerDateTime, "mm/dd/yy")
         LotEdit.lblTime = Format(GetServerDateTime, "hh:mm")
      End With
   End If
   Set rdo = Nothing
   
End Sub

Private Function IsUserLotIDInUse() As Boolean
   Dim lot As New ClassLot
   IsUserLotIDInUse = lot.IsUserLotIdInUseForAnotherLot(txtlot, lblNumber, True)
End Function


Private Function DocumentLokEnabled() As Boolean
    Dim AdoDocLok As ADODB.Recordset
    Dim iDL As Integer
    
    DocumentLokEnabled = False
    sSql = "SELECT CODOCLOK FROM ComnTable WHERE COREF=1"
    bSqlRows = clsADOCon.GetDataSet(sSql, AdoDocLok, ES_FORWARD)
    If bSqlRows Then
        iDL = 0 & AdoDocLok!CODOCLOK
        If iDL = 1 Then DocumentLokEnabled = True
    End If
    Set AdoDocLok = Nothing
End Function

Private Function UpdateUnitCost()

   Dim cCost As Double
   If Trim(txtUnitCost) <> "" Then
      cCost = Val(txtUnitCost)
   
      sSql = "UPDATE LohdTable SET LOTUNITCOST = " & cCost & vbCrLf _
               & "WHERE LOTNUMBER='" & lblNumber & "'"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      
      sSql = "UPDATE InvaTable SET INAMT = " & cCost & " FROM InvaTable, LoitTable " & vbCrLf _
            & " WHERE loiactivity = INNUMBER AND LOINUMBER = '" & lblNumber & "'"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
   End If
   

End Function



