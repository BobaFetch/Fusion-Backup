VERSION 5.00
Begin VB.Form RecvRVe01b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Receive Purchase Order Items"
   ClientHeight    =   6930
   ClientLeft      =   1845
   ClientTop       =   330
   ClientWidth     =   7770
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   Icon            =   "RecvRVe01b.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPS 
      Height          =   285
      Index           =   5
      Left            =   5520
      TabIndex        =   95
      ToolTipText     =   "Amount To Pick"
      Top             =   4440
      Width           =   2055
   End
   Begin VB.TextBox txtPS 
      Height          =   285
      Index           =   4
      Left            =   5520
      TabIndex        =   94
      ToolTipText     =   "Amount To Pick"
      Top             =   3720
      Width           =   2055
   End
   Begin VB.TextBox txtPS 
      Height          =   285
      Index           =   3
      Left            =   5520
      TabIndex        =   93
      ToolTipText     =   "Amount To Pick"
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox txtPS 
      Height          =   285
      Index           =   6
      Left            =   5520
      TabIndex        =   92
      ToolTipText     =   "Amount To Pick"
      Top             =   5160
      Width           =   2055
   End
   Begin VB.TextBox txtPS 
      Height          =   285
      Index           =   2
      Left            =   5520
      TabIndex        =   91
      ToolTipText     =   "Amount To Pick"
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox txtPS 
      Height          =   285
      Index           =   1
      Left            =   5520
      TabIndex        =   90
      ToolTipText     =   "Amount To Pick"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton optDis 
      Height          =   330
      Left            =   6540
      Picture         =   "RecvRVe01b.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   82
      ToolTipText     =   "Display The Report"
      Top             =   6540
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton optPrnNotUsed 
      Height          =   330
      Left            =   7095
      Picture         =   "RecvRVe01b.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   81
      ToolTipText     =   "Print The Report"
      Top             =   6540
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox lblPrinter 
      Height          =   315
      Left            =   1680
      TabIndex        =   78
      Top             =   6240
      Width           =   4695
   End
   Begin VB.CheckBox optLbl 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   6
      Left            =   7080
      TabIndex        =   77
      Top             =   4800
      Width           =   495
   End
   Begin VB.CheckBox optLbl 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   5
      Left            =   7080
      TabIndex        =   76
      Top             =   4080
      Width           =   495
   End
   Begin VB.CheckBox optLbl 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   4
      Left            =   7080
      TabIndex        =   75
      Top             =   3360
      Width           =   495
   End
   Begin VB.CheckBox optLbl 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   3
      Left            =   7080
      TabIndex        =   74
      Top             =   2640
      Width           =   495
   End
   Begin VB.CheckBox optLbl 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   2
      Left            =   7080
      TabIndex        =   73
      Top             =   1920
      Width           =   495
   End
   Begin VB.CheckBox optLbl 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   1
      Left            =   7080
      TabIndex        =   72
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdUp 
      DisabledPicture =   "RecvRVe01b.frx":0314
      DownPicture     =   "RecvRVe01b.frx":0806
      Enabled         =   0   'False
      Height          =   372
      Left            =   6600
      MaskColor       =   &H00000000&
      Picture         =   "RecvRVe01b.frx":0CF8
      Style           =   1  'Graphical
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   5580
      Width           =   400
   End
   Begin VB.CommandButton cmdDn 
      DisabledPicture =   "RecvRVe01b.frx":11EA
      DownPicture     =   "RecvRVe01b.frx":16DC
      Enabled         =   0   'False
      Height          =   372
      Left            =   6600
      MaskColor       =   &H00000000&
      Picture         =   "RecvRVe01b.frx":1BCE
      Style           =   1  'Graphical
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   5970
      Width           =   400
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "RecvRVe01b.frx":20C0
      Style           =   1  'Graphical
      TabIndex        =   65
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdCmt 
      DownPicture     =   "RecvRVe01b.frx":286E
      Height          =   320
      Index           =   6
      Left            =   120
      Picture         =   "RecvRVe01b.frx":2D48
      Style           =   1  'Graphical
      TabIndex        =   63
      TabStop         =   0   'False
      ToolTipText     =   "Edit PO Item Comments"
      Top             =   5160
      Width           =   350
   End
   Begin VB.CommandButton cmdCmt 
      DownPicture     =   "RecvRVe01b.frx":334A
      Height          =   320
      Index           =   5
      Left            =   120
      Picture         =   "RecvRVe01b.frx":3824
      Style           =   1  'Graphical
      TabIndex        =   62
      TabStop         =   0   'False
      ToolTipText     =   "Edit PO Item Comments"
      Top             =   4440
      Width           =   350
   End
   Begin VB.CommandButton cmdCmt 
      DownPicture     =   "RecvRVe01b.frx":3E26
      Height          =   320
      Index           =   4
      Left            =   120
      Picture         =   "RecvRVe01b.frx":4300
      Style           =   1  'Graphical
      TabIndex        =   61
      TabStop         =   0   'False
      ToolTipText     =   "Edit PO Item Comments"
      Top             =   3720
      Width           =   350
   End
   Begin VB.CommandButton cmdCmt 
      DownPicture     =   "RecvRVe01b.frx":4902
      Height          =   320
      Index           =   3
      Left            =   120
      Picture         =   "RecvRVe01b.frx":4DDC
      Style           =   1  'Graphical
      TabIndex        =   60
      TabStop         =   0   'False
      ToolTipText     =   "Edit PO Item Comments"
      Top             =   3000
      Width           =   350
   End
   Begin VB.CommandButton cmdCmt 
      DownPicture     =   "RecvRVe01b.frx":53DE
      Height          =   320
      Index           =   2
      Left            =   120
      Picture         =   "RecvRVe01b.frx":58B8
      Style           =   1  'Graphical
      TabIndex        =   59
      TabStop         =   0   'False
      ToolTipText     =   "Edit PO Item Comments"
      Top             =   2280
      Width           =   350
   End
   Begin VB.CommandButton cmdCmt 
      DownPicture     =   "RecvRVe01b.frx":5EBA
      Height          =   320
      Index           =   1
      Left            =   120
      Picture         =   "RecvRVe01b.frx":6394
      Style           =   1  'Graphical
      TabIndex        =   58
      TabStop         =   0   'False
      ToolTipText     =   "Edit PO Item Comments"
      Top             =   1560
      Width           =   350
   End
   Begin VB.CommandButton optPrn 
      Caption         =   "&Receive"
      Height          =   315
      Left            =   6060
      TabIndex        =   50
      ToolTipText     =   "Receive Selected Items"
      Top             =   520
      Width           =   915
   End
   Begin VB.TextBox txtAqt 
      Height          =   285
      Index           =   6
      Left            =   5170
      TabIndex        =   10
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CheckBox optPck 
      Alignment       =   1  'Right Justify
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   315
      Index           =   6
      Left            =   6405
      TabIndex        =   11
      Top             =   4800
      Width           =   420
   End
   Begin VB.TextBox txtAqt 
      Height          =   285
      Index           =   5
      Left            =   5170
      TabIndex        =   8
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CheckBox optPck 
      Alignment       =   1  'Right Justify
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   315
      Index           =   5
      Left            =   6405
      TabIndex        =   9
      Top             =   4080
      Width           =   420
   End
   Begin VB.TextBox txtAqt 
      Height          =   285
      Index           =   4
      Left            =   5170
      TabIndex        =   6
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CheckBox optPck 
      Alignment       =   1  'Right Justify
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   315
      Index           =   4
      Left            =   6405
      TabIndex        =   7
      Top             =   3360
      Width           =   420
   End
   Begin VB.TextBox txtAqt 
      Height          =   285
      Index           =   3
      Left            =   5170
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CheckBox optPck 
      Alignment       =   1  'Right Justify
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   315
      Index           =   3
      Left            =   6405
      TabIndex        =   5
      Top             =   2640
      Width           =   420
   End
   Begin VB.TextBox txtAqt 
      Height          =   285
      Index           =   2
      Left            =   5170
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CheckBox optPck 
      Alignment       =   1  'Right Justify
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   315
      Index           =   2
      Left            =   6405
      TabIndex        =   3
      Top             =   1920
      Width           =   420
   End
   Begin VB.TextBox txtAqt 
      Height          =   285
      Index           =   1
      Left            =   5170
      TabIndex        =   0
      ToolTipText     =   "Amount To Pick"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CheckBox optPck 
      Alignment       =   1  'Right Justify
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   315
      Index           =   1
      Left            =   6360
      TabIndex        =   1
      ToolTipText     =   "Check If Complete (Checked if More Than Rq'd)"
      Top             =   1200
      Width           =   420
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6060
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin VB.PictureBox ReSize1 
      Height          =   480
      Left            =   4200
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   80
      Top             =   6960
      Width           =   1200
   End
   Begin VB.Label Label3 
      Caption         =   "Due Date        "
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
      Left            =   4080
      TabIndex        =   89
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblDueDate 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   6
      Left            =   4080
      TabIndex        =   88
      Top             =   5160
      Width           =   1080
   End
   Begin VB.Label lblDueDate 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   5
      Left            =   4080
      TabIndex        =   87
      Top             =   4440
      Width           =   1080
   End
   Begin VB.Label lblDueDate 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   4
      Left            =   4080
      TabIndex        =   86
      Top             =   3720
      Width           =   1080
   End
   Begin VB.Label lblDueDate 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   85
      Top             =   3000
      Width           =   1080
   End
   Begin VB.Label lblDueDate 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   84
      Top             =   2280
      Width           =   1080
   End
   Begin VB.Label lblDueDate 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   83
      Top             =   1560
      Width           =   1080
   End
   Begin VB.Label Label2 
      Caption         =   "Label Printer"
      Height          =   255
      Left            =   240
      TabIndex        =   79
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label?"
      Height          =   255
      Left            =   7080
      TabIndex        =   71
      Top             =   960
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Legend:"
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   70
      ToolTipText     =   "Red = Not Inspected, Green = Inspected And Approved, Gray = Inspection Not Required"
      Top             =   5640
      Width           =   800
   End
   Begin VB.Label cmbVnd 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   69
      ToolTipText     =   "Purchase Order Vendor"
      Top             =   120
      Width           =   1320
   End
   Begin VB.Label txtNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   68
      ToolTipText     =   "Purchase Order Vendor"
      Top             =   480
      Width           =   3720
   End
   Begin VB.Image OnDockOff 
      Height          =   300
      Left            =   1440
      Picture         =   "RecvRVe01b.frx":6996
      ToolTipText     =   "Item Does Not Require Inspection"
      Top             =   5640
      Width           =   240
   End
   Begin VB.Image OnDock 
      Height          =   300
      Index           =   6
      Left            =   5220
      Picture         =   "RecvRVe01b.frx":6D98
      ToolTipText     =   "Red = Not Inspected, Green = Inspected And Approved, Gray = Inspection Not Required"
      Top             =   5160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image OnDock 
      Height          =   300
      Index           =   5
      Left            =   5220
      Picture         =   "RecvRVe01b.frx":719A
      ToolTipText     =   "Red = Not Inspected, Green = Inspected And Approved, Gray = Inspection Not Required"
      Top             =   4440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image OnDock 
      Height          =   300
      Index           =   4
      Left            =   5220
      Picture         =   "RecvRVe01b.frx":759C
      ToolTipText     =   "Red = Not Inspected, Green = Inspected And Approved, Gray = Inspection Not Required"
      Top             =   3720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image OnDock 
      Height          =   300
      Index           =   3
      Left            =   5220
      Picture         =   "RecvRVe01b.frx":799E
      ToolTipText     =   "Red = Not Inspected, Green = Inspected And Approved, Gray = Inspection Not Required"
      Top             =   3000
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image OnDock 
      Height          =   300
      Index           =   2
      Left            =   5220
      Picture         =   "RecvRVe01b.frx":7DA0
      ToolTipText     =   "Red = Not Inspected, Green = Inspected And Approved, Gray = Inspection Not Required"
      Top             =   2280
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image OnDock 
      Height          =   300
      Index           =   1
      Left            =   5220
      Picture         =   "RecvRVe01b.frx":81A2
      ToolTipText     =   "Red = Not Inspected, Green = Inspected And Approved, Gray = Inspection Not Required"
      Top             =   1560
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image OnDockNo 
      Height          =   300
      Left            =   1200
      Picture         =   "RecvRVe01b.frx":85A4
      ToolTipText     =   "Not Yet Inspected"
      Top             =   5640
      Width           =   240
   End
   Begin VB.Image OnDockYes 
      Height          =   300
      Left            =   960
      Picture         =   "RecvRVe01b.frx":89A6
      ToolTipText     =   "Inspected And Approved"
      Top             =   5640
      Width           =   240
   End
   Begin VB.Label lblNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3600
      TabIndex        =   64
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "                                "
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2400
      TabIndex        =   57
      Top             =   7320
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   6
      Left            =   480
      TabIndex        =   56
      Top             =   4800
      Width           =   276
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   5
      Left            =   480
      TabIndex        =   55
      Top             =   4080
      Width           =   276
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   4
      Left            =   480
      TabIndex        =   54
      Top             =   3360
      Width           =   276
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   3
      Left            =   480
      TabIndex        =   53
      Top             =   2640
      Width           =   276
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   2
      Left            =   480
      TabIndex        =   52
      Top             =   1920
      Width           =   276
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   1
      Left            =   480
      TabIndex        =   51
      Top             =   1200
      Width           =   276
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rel"
      Height          =   255
      Index           =   10
      Left            =   1320
      TabIndex        =   49
      Top             =   6960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   48
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblRel 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3360
      TabIndex        =   47
      Top             =   6960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblPon 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   480
      TabIndex        =   46
      ToolTipText     =   "Purchase Order"
      Top             =   120
      Width           =   660
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "of"
      Height          =   252
      Index           =   8
      Left            =   5640
      TabIndex        =   45
      Top             =   5640
      Width           =   372
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Page"
      Height          =   252
      Index           =   7
      Left            =   4560
      TabIndex        =   44
      Top             =   5640
      Width           =   612
   End
   Begin VB.Label lblLst 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   6000
      TabIndex        =   43
      Top             =   5640
      Width           =   372
   End
   Begin VB.Label lblPge 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   5160
      TabIndex        =   42
      Top             =   5640
      Width           =   372
   End
   Begin VB.Image Dsdn 
      Height          =   300
      Left            =   2040
      Picture         =   "RecvRVe01b.frx":8DA8
      Top             =   5760
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Endn 
      Height          =   300
      Left            =   2280
      Picture         =   "RecvRVe01b.frx":929A
      Top             =   5760
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Enup 
      Height          =   300
      Left            =   2760
      Picture         =   "RecvRVe01b.frx":978C
      Top             =   5760
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Dsup 
      Height          =   300
      Left            =   2520
      Picture         =   "RecvRVe01b.frx":9C7E
      Top             =   5760
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   6
      Left            =   120
      TabIndex        =   41
      Top             =   4800
      Width           =   372
   End
   Begin VB.Label lblPrt 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   6
      Left            =   840
      TabIndex        =   40
      ToolTipText     =   "Click For Extended Information"
      Top             =   4800
      Width           =   3108
   End
   Begin VB.Label lblPqt 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   6
      Left            =   4080
      TabIndex        =   39
      ToolTipText     =   "Click To Update The Received Quantity"
      Top             =   4800
      Width           =   1080
   End
   Begin VB.Label lblDsc 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   6
      Left            =   840
      TabIndex        =   38
      ToolTipText     =   "Click For Extended Information"
      Top             =   5112
      Width           =   3108
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   5
      Left            =   120
      TabIndex        =   37
      Top             =   4080
      Width           =   372
   End
   Begin VB.Label lblPrt 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   5
      Left            =   840
      TabIndex        =   36
      ToolTipText     =   "Click For Extended Information"
      Top             =   4080
      Width           =   3108
   End
   Begin VB.Label lblPqt 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   5
      Left            =   4080
      TabIndex        =   35
      ToolTipText     =   "Click To Update The Received Quantity"
      Top             =   4080
      Width           =   1080
   End
   Begin VB.Label lblDsc 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   5
      Left            =   840
      TabIndex        =   34
      ToolTipText     =   "Click For Extended Information"
      Top             =   4440
      Width           =   3108
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   4
      Left            =   120
      TabIndex        =   33
      Top             =   3360
      Width           =   372
   End
   Begin VB.Label lblPrt 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   4
      Left            =   840
      TabIndex        =   32
      ToolTipText     =   "Click For Extended Information"
      Top             =   3360
      Width           =   3108
   End
   Begin VB.Label lblPqt 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   4
      Left            =   4080
      TabIndex        =   31
      ToolTipText     =   "Click To Update The Received Quantity"
      Top             =   3360
      Width           =   1080
   End
   Begin VB.Label lblDsc 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   4
      Left            =   840
      TabIndex        =   30
      ToolTipText     =   "Click For Extended Information"
      Top             =   3720
      Width           =   3108
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   3
      Left            =   120
      TabIndex        =   29
      Top             =   2640
      Width           =   372
   End
   Begin VB.Label lblPrt 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   3
      Left            =   840
      TabIndex        =   28
      ToolTipText     =   "Click For Extended Information"
      Top             =   2640
      Width           =   3108
   End
   Begin VB.Label lblPqt 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   3
      Left            =   4080
      TabIndex        =   27
      ToolTipText     =   "Click To Update The Received Quantity"
      Top             =   2640
      Width           =   1080
   End
   Begin VB.Label lblDsc 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   3
      Left            =   840
      TabIndex        =   26
      ToolTipText     =   "Click For Extended Information"
      Top             =   2952
      Width           =   3108
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   2
      Left            =   120
      TabIndex        =   25
      Top             =   1920
      Width           =   372
   End
   Begin VB.Label lblPrt 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   2
      Left            =   840
      TabIndex        =   24
      ToolTipText     =   "Click For Extended Information"
      Top             =   1920
      Width           =   3108
   End
   Begin VB.Label lblPqt 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   2
      Left            =   4080
      TabIndex        =   23
      ToolTipText     =   "Click To Update The Received Quantity"
      Top             =   1920
      Width           =   1080
   End
   Begin VB.Label lblDsc 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   2
      Left            =   840
      TabIndex        =   22
      ToolTipText     =   "Click For Extended Information"
      Top             =   2232
      Width           =   3108
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   1
      Left            =   120
      TabIndex        =   21
      Top             =   1200
      Width           =   372
   End
   Begin VB.Label lblPrt 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   1
      Left            =   840
      TabIndex        =   20
      ToolTipText     =   "Click For Extended Information"
      Top             =   1200
      Width           =   3108
   End
   Begin VB.Label lblPqt 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   1
      Left            =   4080
      TabIndex        =   19
      ToolTipText     =   "Click To Update The Received Quantity"
      Top             =   1200
      Width           =   1080
   End
   Begin VB.Label lblDsc 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   1
      Left            =   840
      TabIndex        =   18
      ToolTipText     =   "Click For Extended Information"
      Top             =   1512
      Width           =   3108
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comp?    "
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
      Index           =   6
      Left            =   6360
      TabIndex        =   17
      Top             =   960
      Width           =   624
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Recv Qty / PS      "
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
      Left            =   5160
      TabIndex        =   16
      Top             =   960
      Width           =   1092
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Qty (Avail) /"
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   15
      Top             =   750
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number/Description                              "
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
      Index           =   1
      Left            =   840
      TabIndex        =   14
      Top             =   960
      Width           =   3132
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item/Rev"
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
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   732
   End
End
Attribute VB_Name = "RecvRVe01b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'1/15/03 On Dock
'2/4/04 Auto Pick Purchased Material
'9/29/04 Reformatted values to prevent overflow
'10/29/04 Added delete an unpicked item that is the same as received
'1/18/05 Corrected InvaTable.INMOPART & INMORUN in ReceiveItems
'1/24/05 Added PKUNITS to Insert MopkTable
'12/14/05 Changed position of transactions to beginning and end if each Item
'4/17/06 Revisited receiving and moved some SQL calls to
'        before the transaction begins. If Lots and no error, commit (see GNext1)
'        Added LOIUNITS, INUNITS to Insert query
'9/21/06 Revised layout, fixed hide Stoplights, added legend and ToolTips
'2/1/07 Replaced the ONDOCK columns for split PO quantities 7.1.8
'2/28/07 Added update to RunsTable.RUNSTATUS in ReceiveItems (PP if not PC)
Option Explicit
Dim AdoQry1 As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter

Dim AdoQry2 As ADODB.Command
Dim AdoParameter2 As ADODB.Parameter
Dim AdoParameter3 As ADODB.Parameter

Dim bOnLoad As Byte
Dim iCurrPage As Integer
Dim iIndex As Integer
Dim iMaxPages As Integer
Dim iTotalItems As Integer

Dim sCreditAcct As String
Dim sDebitAcct As String
Dim sLabelPrinter As String

Dim sPartNumber(150) As String '9/23/04 PARTREF
Dim vItems(150, 24) As Variant
' 0 = item
' 1 = rev
' 2 = compressed part
' 3 = Part Number
' 4 = desc
' 5 = standard cost
' 6 = planned qty
Const VITEMS_PlannedQty = 6
' 7 = actual qty
Const VITEMS_ActualQty = 7
Const VITEMS_DueDate = 9
' 8 = complete?
' 9 = PIPDATE
'10 = PIESTUNIT
Const VITEMS_EstUnitCost = 10
'11 = PIRUNPART
'12 = PIRUNNO
'13 = PIRUNOPNO
'14 = PICOMT
'15 = PIONDOCK (1 or 0)
Const VITEMS_ItemRequiresInspection = 15
'16 = PIONDOCKINSPECTED (1 or 0)
Const VITEMS_ItemHasBeenInspected = 16
'17 = PIONDOCKQTYACC
'18 = PIONDOCKQTYREJ
'19 = Lot Tracked
'20 = Part Type (level)
Const VITEMS_PARTTYPE = 20
'21 = PAUNITS
'22 = PrinttLabels
Const VITEMS_PRINTLABELS = 22
Const VITEMS_PS = 24

Public quantityPerLabel As Currency
Public totalQuantity As Currency

'BBS Added this function on 03/17/2010 for Ticket #27859
Private Function FormCount(ByVal frmName As String) As Long
    Dim frm As Form
    For Each frm In Forms
        If StrComp(frm.Name, frmName, vbTextCompare) = 0 Then
            FormCount = FormCount + 1
        End If
    Next
End Function


Private Sub cmdCan_Click()
   Dim iList As Integer
   Dim A As Integer
   Dim bByte As Byte
   Dim bResponse As Byte
   Dim sMsg As String
   
   bByte = False
   For iList = 1 To iTotalItems
      If vItems(iList, VITEMS_ActualQty) > 0 Then bByte = True
   Next
   If bByte Then
      sMsg = "There Are Items Marked For Receipt." & vbCr _
             & "Do You Still Wish To Quit?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbYes Then Unload Me
   Else
      Unload Me
   End If
   
End Sub



Private Sub cmdCmt_Click(Index As Integer)
   If cmdCmt(Index) Then
      MouseCursor 13
      Load RecvRVe01c
      RecvRVe01c.lblPon = lblPon
      RecvRVe01c.lblItem = lblItm(Index)
      RecvRVe01c.lblRev = lblRev(Index)
      RecvRVe01c.Show
      cmdCmt(Index) = False
   End If
   
End Sub

Private Sub cmdDn_Click()
   iCurrPage = iCurrPage + 1
   If iCurrPage > iMaxPages Then iCurrPage = iMaxPages
   GetThisGroup
   
End Sub

Private Sub cmdDn_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "5301"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub optPrn_Click()
   optPrn.Enabled = False  'prevent duplicate clicks
   Dim iList As Integer
   Dim bByte As Byte
   bByte = False
   
   For iList = 1 To iTotalItems
      If vItems(iList, VITEMS_ActualQty) > 0 Then
         bByte = True
      End If
   Next
   
   If Not bByte Then
      MsgBox "There Are No Items Marked For Receipt.", vbInformation, Caption
   Else
      ReceiveItems
      Dim nCount As Integer
      For iList = 1 To iTotalItems
         If vItems(iList, VITEMS_PRINTLABELS) > 0 Then
            Load RecvRVe01d
            RecvRVe01d.lblItem = vItems(iList, 0)
            RecvRVe01d.lblRev = vItems(iList, 1)
            RecvRVe01d.lblPartNo = vItems(iList, 3)
            RecvRVe01d.lblQty = vItems(iList, VITEMS_ActualQty)
            RecvRVe01d.txtQtyPerLabel = vItems(iList, VITEMS_ActualQty)
            'RecvRVe01d.Cancelled = False
            Set RecvRVe01d.ParentForm = Me
            RecvRVe01d.Show vbModal
            If Me.quantityPerLabel > 0 Then
               'MsgBox "print " & Me.quantityPerLabel & " per label"
               PrintLabels iList, Me.totalQuantity, Me.quantityPerLabel
            End If
         End If
      Next
      Unload Me
   End If
   optPrn.Enabled = True
End Sub

Private Sub cmdUp_Click()
   iCurrPage = iCurrPage - 1
   If iCurrPage < 1 Then iCurrPage = 1
   GetThisGroup
   
End Sub

Private Sub cmdUp_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub


Private Sub Form_Activate()
   Dim X As Printer
   MdiSect.lblBotPanel = Caption
   If bOnLoad = 1 Then
      ES_PurchasedDataFormat = GetPODataFormat()
      cmbVnd = RecvRVe01a.cmbVnd
      txtNme = RecvRVe01a.txtNme
      
      'populate printer selection combo
      For Each X In Printers
         If Left(X.DeviceName, 9) <> "Rendering" Then
            lblPrinter.AddItem X.DeviceName
         End If
      Next
      
      On Error Resume Next
      
      Dim sDefaultPrinter As String
      If lblPrinter.ListCount > 0 Then
         sDefaultPrinter = lblPrinter.List(0)
      End If
      
      lblPrinter.Text = GetSetting("Esi2000", "EsiInvc", "ReceivingLabelPrinter", sDefaultPrinter)
      
   End If
   bOnLoad = 0
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   On Error Resume Next
   MdiSect.BotPanel = Caption
   
   
   FormLoad Me, ES_DONTLIST
   lblPon = RecvRVe01a.cmbPon
   lblRel = RecvRVe01a.txtRel
   lblInfo.BackColor = ES_ViewBackColor
   sSql = "SELECT PINUMBER,PIITEM,PIREV,PITYPE,PIPART," _
          & "PIPQTY,PIPDATE,PIESTUNIT,PIRUNPART,PIRUNNO,PIRUNOPNO," _
          & "PIONDOCK,PIONDOCKINSPECTED,PIONDOCKQTYACC,PIONDOCKQTYREJ,PIODDELPSNUMBER," _
          & "PICOMT,PARTREF,PARTNUM,PADESC,PALEVEL,PAUNITS,PASTDCOST,PALOTTRACK FROM " _
          & "PoitTable,PartTable WHERE PIPART=PARTREF AND " _
          & "(PINUMBER= ? AND PITYPE=14) ORDER by PIITEM "
   Set AdoQry1 = New ADODB.Command
   AdoQry1.CommandText = sSql
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adInteger
   AdoQry1.Parameters.Append AdoParameter1
   
   'Set RdoQry1 = RdoCon.CreateQuery("", sSql)
   
   sSql = "SELECT PINUMBER,PIITEM,PIREV,PIPART,PICOMT,PARTREF," _
          & "PARTNUM,PAEXTDESC FROM PoitTable,PartTable WHERE (PIPART=PARTREF AND " _
          & "PINUMBER=" & Val(lblPon) & ") AND PIITEM= ? AND PIREV= ? "
   Set AdoQry2 = New ADODB.Command
   AdoQry2.CommandText = sSql
   Set AdoParameter2 = New ADODB.Parameter
   AdoParameter2.Type = adSmallInt
   AdoQry2.Parameters.Append AdoParameter2
   
   Set AdoParameter3 = New ADODB.Parameter
   AdoParameter3.Type = adChar
   AdoParameter3.Size = 2
   AdoQry2.Parameters.Append AdoParameter3
   
   
   'Set RdoQry2 = RdoCon.CreateQuery("", sSql)
   'RdoQry2.MaxRows = 1
   cmdUp.Picture = Enup
   cmdUp.Enabled = True
   cmdDn.Picture = Endn
   cmdDn.Enabled = True
   bOnLoad = 1
   GetItems
   'If iTotalItems = 0 Then Unload Me
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   RecvRVe01a.optItm.Value = vbUnchecked
   
End Sub

Private Sub PrintLabels(subscript As Integer, totalQuantity As Currency, quantityPerLabel As Currency)
   Dim b As Byte
   Dim FormDriver As String
   Dim FormPort As String
   Dim FormPrinter As String
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   On Error GoTo DiaErr1
   MouseCursor 13
   FormPrinter = Trim(lblPrinter)
'   If Len(Trim(FormPrinter)) > 0 Then
'      b = GetPrinterPort(FormPrinter, FormDriver, FormPort)
'   Else
'      FormPrinter = ""
'      FormDriver = ""
'      FormPort = ""
'   End If
   
   Dim quantityLeft As Currency
   quantityLeft = totalQuantity
   
   Do
      
      sCustomReport = GetCustomReport("Invre01")
      Set cCRViewer = New EsCrystalRptViewer
      cCRViewer.Init
      cCRViewer.SetReportFileName sCustomReport, sReportPath
      cCRViewer.SetReportTitle = sCustomReport
      aFormulaName.Add "PONumber"
      aFormulaName.Add "POItemNumber"
      aFormulaName.Add "POItemRevision"
      aFormulaName.Add "Quantity"
      
      
      aFormulaValue.Add CStr("'" & lblPon.Caption & "'")
      aFormulaValue.Add CStr("'" & vItems(subscript, 0) & "'")
      aFormulaValue.Add CStr("'" & vItems(subscript, 1) & "'")
       
      If quantityLeft > quantityPerLabel Then
         aFormulaValue.Add CStr("'" & quantityPerLabel & "'")
      Else
         aFormulaValue.Add CStr("'" & quantityLeft & "'")
      End If
       
      cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
         
         
      cCRViewer.CRViewerSize Me
      cCRViewer.SetDbTableConnection
      cCRViewer.OpenCrystalReportObject Me, aFormulaName
      cCRViewer.ShowGroupTree False
      
      cCRViewer.ClearFieldCollection aRptPara
      cCRViewer.ClearFieldCollection aFormulaName
      cCRViewer.ClearFieldCollection aFormulaValue
      Set cCRViewer = Nothing
      
   ''      MdiSect.Crw.PrinterName = FormPrinter
   ''      MdiSect.Crw.PrinterDriver = FormDriver
   ''      MdiSect.Crw.PrinterPort = FormPort
   ''      MdiSect.Crw.Destination = crptToPrinter
   '      MdiSect.Crw.ReportFileName = sReportPath & GetCustomReport("Invre01")
   '
   '      MdiSect.Crw.Formulas(0) = "PONumber='" & lblPon.Caption & "'"
   '      MdiSect.Crw.Formulas(1) = "POItemNumber='" & vItems(subscript, 0) & "'"
   '      MdiSect.Crw.Formulas(2) = "POItemRevision='" & vItems(subscript, 1) & "'"
   '      If quantityLeft > quantityPerLabel Then
   '         MdiSect.Crw.Formulas(3) = "Quantity='" & quantityPerLabel & "'"
   '      Else
   '         MdiSect.Crw.Formulas(3) = "Quantity='" & quantityLeft & "'"
   '      End If
   '
   '      SetCrystalAction Me
         
      quantityLeft = quantityLeft - quantityPerLabel
   
   Loop While quantityLeft > 0
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "PrintLabels"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   
   SaveSetting "Esi2000", "EsiInvc", "ReceivingLabelPrinter", lblPrinter.Text
   
   Set AdoParameter1 = Nothing
   Set AdoParameter2 = Nothing
   Set AdoParameter3 = Nothing
   Set AdoQry1 = Nothing
   Set AdoQry2 = Nothing
   Set RecvRVe01b = Nothing
   
   If FormCount("frmTip") > 0 Then Unload frmTip 'BBS Added this on 03/17/2010 for Ticket #27859

End Sub





Private Sub GetItems()
   MouseCursor 13
   Dim RdoItm As ADODB.Recordset
   Dim iRow As Integer
   ClearBoxes
   iRow = 0
   iTotalItems = 0
   On Error GoTo DiaErr1
   AdoQry1.Parameters(0).Value = Val(lblPon)
   bSqlRows = clsADOCon.GetQuerySet(RdoItm, AdoQry1, ES_STATIC, False, 1)
   If bSqlRows Then
      With RdoItm
         Do Until .EOF
                 
            ' the arry size is set to 150 and the array starts at 1 - 149
            If (iRow = 149) Then
               Exit Do
            End If
                 
            iRow = iRow + 1
            vItems(iRow, 0) = Format(!PIITEM, "##0")
            vItems(iRow, 1) = "" & Trim(!PIREV)
            vItems(iRow, 2) = "" & Trim(!PartRef)
            sPartNumber(iRow) = "" & Trim(!PartRef)
            vItems(iRow, 3) = "" & Trim(!PartNum)
            vItems(iRow, 4) = "" & Trim(!PADESC)
            vItems(iRow, 5) = Format(!PASTDCOST, ES_QuantityDataFormat)
            vItems(iRow, VITEMS_PlannedQty) = Format(!PIPQTY, ES_QuantityDataFormat)
            vItems(iRow, VITEMS_ActualQty) = Format(0, ES_QuantityDataFormat)
            vItems(iRow, 8) = 0
            
            vItems(iRow, 9) = Format(!PIPDATE, "mm/dd/yy")
            vItems(iRow, VITEMS_EstUnitCost) = Format(!PIESTUNIT, ES_PurchasedDataFormat)
            vItems(iRow, 11) = "" & Trim(!PIRUNPART)
            vItems(iRow, 12) = Format(!PIRUNNO, "####0")
            vItems(iRow, 13) = Format(!PIRUNOPNO, "##0")
            vItems(iRow, 14) = "" & Trim(!PICOMT)
            'On Dock
            vItems(iRow, VITEMS_ItemRequiresInspection) = Format(!PIONDOCK, "#0")
            vItems(iRow, VITEMS_ItemHasBeenInspected) = Format(!PIONDOCKINSPECTED, "#0")
            vItems(iRow, 17) = Format(!PIONDOCKQTYACC, ES_QuantityDataFormat)
            vItems(iRow, 18) = Format(!PIONDOCKQTYREJ, ES_QuantityDataFormat)
            vItems(iRow, 19) = Format(!PALOTTRACK, "0")
            vItems(iRow, VITEMS_PARTTYPE) = Format(!PALEVEL, "0")
            vItems(iRow, 21) = "" & Trim(!PAUNITS)
            If Val(vItems(iRow, VITEMS_ItemRequiresInspection)) = 1 Then
               If Val(vItems(iRow, VITEMS_ItemHasBeenInspected)) = 1 Then
                  vItems(iRow, VITEMS_PlannedQty) = vItems(iRow, 17)
               End If
            End If
            vItems(iRow, VITEMS_PS) = Trim(!PIODDELPSNUMBER)
            
            .MoveNext
         Loop
         ClearResultSet RdoItm
      End With
   Else
      optPrn.Visible = False
      'MsgBox "No Items Found.", vbInformation, Caption
      Exit Sub
   End If
   
   optPrn.Visible = True
   iTotalItems = iRow
   
   If iTotalItems > 0 Then
      iCurrPage = 1
      iIndex = 0
      'iMaxPages = 1 + (iTotalItems \ 6)
      iMaxPages = (iTotalItems + 5) \ 6
      lblPge = "1"
      lblLst = str(iMaxPages)
   Else
      lblPge = "0"
      lblLst = "0"
   End If
   
   GetThisGroup
   
   'On Error Resume Next
   Set RdoItm = Nothing
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "getitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub GetThisGroup()
   Dim A As Integer
   Dim b
   Dim iRow As Integer
   A = (iCurrPage * 6) - 5
   iIndex = A - 1
   lblPge = LTrim(str(iCurrPage))
   ClearBoxes
   On Error GoTo DiaErr1
   For iRow = A To A + 5
      If iRow > iTotalItems Then Exit For
      b = (iRow + 1) - A
'      OnDock(b).Picture = OnDockOff.Picture
'      OnDock(b).Visible = True
      txtAqt(b).Enabled = True
      optPck(b).Enabled = True
      lblItm(b).Visible = True
      lblRev(b).Visible = True
      lblPrt(b).Visible = True
      lblDsc(b).Visible = True
      lblPqt(b).Visible = True
      lblDueDate(b).Visible = True
      txtAqt(b).Visible = True
      optPck(b).Visible = True
      optLbl(b).Visible = True
      cmdCmt(b).Visible = True
      lblItm(b) = vItems(iRow, 0)
      lblRev(b) = vItems(iRow, 1)
      lblPrt(b) = vItems(iRow, 3)
      lblDsc(b) = vItems(iRow, 4)
      lblPqt(b) = vItems(iRow, VITEMS_PlannedQty)
      lblDueDate(b) = vItems(iRow, VITEMS_DueDate)
      
      txtAqt(b) = vItems(iRow, VITEMS_ActualQty)
      optPck(b).Value = vItems(iRow, 8)
      
      'On Dock
      txtPS(b).Visible = True
      txtPS(b) = vItems(iRow, VITEMS_PS)
      
      OnDock(b).Picture = OnDockOff.Picture
      OnDock(b).Visible = True
      lblPqt(b).ToolTipText = "Click To Update The Received Quantity"
      lblDueDate(b).ToolTipText = "Due Date"
      If Val(vItems(iRow, VITEMS_ItemRequiresInspection)) = 1 Then
         If Val(vItems(iRow, VITEMS_ItemHasBeenInspected)) = 1 Then
            OnDock(b).Picture = OnDockYes.Picture
         Else
            OnDock(b).Picture = OnDockNo.Picture
            txtAqt(b).Enabled = False
            optPck(b).Enabled = False
            lblPqt(b).ToolTipText = ""
            lblDueDate(b).ToolTipText = ""
         End If
      End If
   Next
   On Error Resume Next
   If iCurrPage = 1 Then
      cmdUp.Picture = Dsup
      cmdUp.Enabled = False
   Else
      cmdUp.Picture = Enup
      cmdUp.Enabled = True
   End If
   If iCurrPage = iMaxPages Then
      cmdDn.Picture = Dsdn
      cmdDn.Enabled = False
   Else
      cmdDn.Picture = Endn
      cmdDn.Enabled = True
   End If
   txtAqt(1).SetFocus
   Exit Sub
   
DiaErr1:
   sProcName = "getthisgr"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub lblDsc_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Or Button = 2 Then
      If Len(Trim(lblPrt(Index))) Then GetExtendedInfo (Index)
   End If
   
End Sub


Private Sub lblItm_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
      If Len(Trim(lblPrt(Index))) Then GetExtendedInfo (Index)
   End If
   
End Sub



Private Sub lblPqt_Click(Index As Integer)
   
   'set received quantity = total quantity except where an inspection is required
   If Val(vItems(Index + iIndex, VITEMS_ItemRequiresInspection)) = 0 _
      Or Val(vItems(Index + iIndex, VITEMS_ItemHasBeenInspected)) = 1 Then
   
      txtAqt(Index) = lblPqt(Index)
      vItems(Index + iIndex, VITEMS_ActualQty) = txtAqt(Index)
      optPck(Index).Value = vbChecked
      vItems(Index + iIndex, 8) = optPck(Index).Value
   
   End If
   
End Sub

Private Sub lblPrt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Or Button = 2 Then
      If Len(Trim(lblPrt(Index))) Then GetExtendedInfo (Index)
   End If
   
End Sub


Private Sub lblRev_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
      If Len(Trim(lblPrt(Index))) Then GetExtendedInfo (Index)
   End If
   
End Sub



Private Sub optLbl_Click(Index As Integer)
   vItems(Index + iIndex, VITEMS_PRINTLABELS) = optLbl(Index).Value
End Sub

Private Sub optPck_Click(Index As Integer)
   If optPck(Index).Value = vbChecked Then
      If vItems(Index + iIndex, VITEMS_ActualQty) = 0 Then
         txtAqt(Index) = lblPqt(Index)
         vItems(Index + iIndex, VITEMS_ActualQty) = lblPqt(Index)
      End If
   End If
   vItems(Index + iIndex, 8) = optPck(Index).Value
   
End Sub


Private Sub optPck_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optPck_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub


Private Sub optPck_LostFocus(Index As Integer)
   If Val(vItems(Index + iIndex, VITEMS_PlannedQty)) <= Val(vItems(Index + iIndex, VITEMS_ActualQty)) Then optPck(Index).Value = vbChecked
   If vItems(Index + iIndex, VITEMS_ActualQty) = 0 Then optPck(Index).Value = vbUnchecked
   
End Sub


Private Sub txtAqt_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtAqt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtAqt_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub


Private Sub txtAqt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub

Private Sub txtAqt_LostFocus(Index As Integer)
   txtAqt(Index) = CheckLen(txtAqt(Index), 9)
   txtAqt(Index) = Format(Abs(Val(txtAqt(Index))), ES_QuantityDataFormat)
   If OnDock(Index).Picture = OnDockYes.Picture And Val(txtAqt(Index)) > 0 Then
      If Val(txtAqt(Index)) <> Val(lblPqt(Index)) Then
         Beep
         MsgBox "Inspected Quantity And Received Quantity Must Match.", _
            vbInformation, Caption
         txtAqt(Index) = lblPqt(Index)
      End If
   End If
   vItems(Index + iIndex, VITEMS_ActualQty) = txtAqt(Index)
   If Val(txtAqt(Index)) >= Val(lblPqt(Index)) Then
      If Val(txtAqt(Index)) <> vItems(Index + iIndex, VITEMS_PlannedQty) Then optPck(Index).Value = vbChecked
   Else
      If Val(txtAqt(Index)) <> vItems(Index + iIndex, VITEMS_ActualQty) Then optPck(Index).Value = vbUnchecked
   End If
   vItems(Index + iIndex, VITEMS_ActualQty) = txtAqt(Index)
   If Val(txtAqt(Index)) = 0 Then optPck(Index).Value = vbUnchecked
   vItems(Index + iIndex, 8) = optPck(Index).Value
   
End Sub

'Beginning of lot tracking 2/27/02
'Added code and tested lots 3/12/02

Private Sub ReceiveItems()
   Dim rdoRev As ADODB.Recordset
   
   Dim bByte As Byte
   Dim bLots As Byte
   Dim bResponse As Byte
   
   Dim A As Integer
   Dim iRow As Integer
   Dim iPkRecord As Integer
   Dim iFailed As Integer
   
   Dim lCOUNTER As Long
   Dim lSysCount As Long
   
   Dim cQuantity As Currency
   Dim cPickQty As Currency
   Dim cPlannedQty As Currency
   Dim cRejQty As Currency
   Dim cCost As Currency
   
   Dim sMsg As String
   Dim sPon As String
   Dim sRev As String
   Dim sMon As String
   Dim sDate As String
   Dim sLotCmt As String
   Dim sLotNumber As String
   Dim sVendor As String
   
   ' On Dock
   Dim sPIONDOCKINSPDATE As String
   Dim sPIONDOCKREJTAG As String
   Dim sPIONDOCKINSPECTOR As String
   Dim sPIONDOCKCOMMENT As String
   
   sDate = RecvRVe01a.txtRcd
   sVendor = Compress(RecvRVe01a.cmbVnd)
   
   sMsg = "Are You Ready To Receive The Selected Items?.."
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbNo Then Exit Sub
   
   For iRow = 1 To iTotalItems
      If Val(vItems(iRow, VITEMS_ActualQty)) > Val(vItems(iRow, VITEMS_PlannedQty)) Then bByte = True
   Next
   
   sMsg = "You Have Items That The Received Quantity Is Greater" & vbCr _
          & "Than The Ordered Quantity. Do You Wish To Continue?"
   If bByte Then
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbNo Then Exit Sub
   End If
   bByte = False
   For iRow = 1 To iTotalItems
      If Val(vItems(iRow, VITEMS_ActualQty)) > 0 Then
         If Val(vItems(iRow, VITEMS_PlannedQty)) > Val(vItems(iRow, VITEMS_ActualQty)) Then bByte = True
      End If
   Next
   If bByte Then
      sMsg = "You Have Items That The Received Quantity Is Less" & vbCr _
             & "Than The Ordered Quantity. Do You Wish To Continue?.."
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbNo Then Exit Sub
   End If
   
   'All clear, let's receive some
   
   bLots = CheckLotTracking()
   MouseCursor 13
   
   'On Error Resume Next
   On Error GoTo RitmsPi1
   'lCOUNTER = GetLastActivity()    'move inside of transaction
   lSysCount = lCOUNTER + 1
   iFailed = 0
   For iRow = 1 To iTotalItems
      If Val(vItems(iRow, VITEMS_ActualQty)) > 0 Then
         'See revision notes
         '7/8/99 (accounts)
         bByte = GetPartAccounts(sPartNumber(iRow), sDebitAcct, sCreditAcct)
         '2/27/02 lots
         sLotNumber = GetNextLotNumber()
         cQuantity = Val(vItems(iRow, VITEMS_ActualQty))
         cPlannedQty = Val(vItems(iRow, VITEMS_PlannedQty))
         cCost = Val(vItems(iRow, VITEMS_EstUnitCost))
         
         ' adjust qty and cost to inventory units if they are different from purchase units
         Dim conversionQty As Currency
         Dim cAdjPlannedQty As Currency, cAdjActualQty As Currency, cAdjCost As Currency
         cAdjPlannedQty = cPlannedQty
         cAdjActualQty = cQuantity
         cAdjCost = cCost
         
         ' sheet inventory no longer handled this way
         ' conversionQty = AdjustForPurchaseUnits(sPartNumber(iRow), cAdjPlannedQty, cAdjActualQty, cAdjCost)
         
         Err.Clear
         clsADOCon.ADOErrNum = 0
         clsADOCon.BeginTrans       '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
         
         lCOUNTER = GetLastActivity() + 1
         
         ' if using sheet inventory, use abbreviated user lot number
         sPon = "PO " & lblPon & "-Item " & vItems(iRow, 0) & vItems(iRow, 1)
         Dim userLotNo As String
         If UsingSheetInventory Then
            userLotNo = lblPon & "-" & vItems(iRow, 0) & vItems(iRow, 1)
         ElseIf GetComnTableBit("CoUseAbbreviatedLotNumbers") Then
            'userLotNo = "PO" & lblPon & Format(vItems(iRow, 0), String(3, "0")) & vItems(iRow, 1) & "-" & Format(ES_SYSDATE, "mm/dd/yyyy")
            userLotNo = lblPon & Format(vItems(iRow, 0), String(3, "0")) & vItems(iRow, 1)
         Else
            userLotNo = sPon & "-" & Format(ES_SYSDATE, "mm/dd/yyyy")
         End If
         sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2," _
                & "INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT,INPONUMBER," _
                & "INPORELEASE,INPOITEM,INPOREV,INPDATE,INADATE,INNUMBER,INLOTNUMBER,INUSER," _
                & "INUNITS) VALUES(" & IATYPE_PoReceipt & ", '" & sPartNumber(iRow) & "'," & "'PO RECEIPT','" _
                & sPon & "'," & cAdjPlannedQty & "," & cAdjActualQty & "," _
                & cAdjCost & ",'" & sCreditAcct & "','" & sDebitAcct _
                & "'," & Val(lblPon) & "," & Val(lblRel) & "," & Val(vItems(iRow, 0)) _
                & ",'" & vItems(iRow, 1) & "','" & sDate & "','" & sDate & "'," & lCOUNTER & ",'" _
                & sLotNumber & "','" & sInitials & "','" & vItems(iRow, 21) & "')"
                
         clsADOCon.ExecuteSql sSql
         
         'Used on an MO?
         If Len(Trim(vItems(iRow, 11))) > 0 And Val(vItems(iRow, 12)) > 0 Then
            'Should be nothing if it's on an MO (See below)
'            sSql = "INSERT INTO LohdTable (LOTNUMBER,LOTUSERLOTID," _
'                   & "LOTPARTREF,LOTPDATE,LOTORIGINALQTY,LOTREMAININGQTY," _
'                   & "LOTPO,LOTPOITEM,LOTPOITEMREV) VALUES('" & sLotNumber _
'                   & "','" & sPon & "-" & Format(ES_SYSDATE, "mm/dd/yyy") _
'                   & "','" & sPartNumber(iRow) & "','" & sDate & "'," _
'                   & cAdjActualQty & ",0," & Val(lblPon) & "," _
'                   & Val(vItems(iRow, 0)) & ",'" & vItems(iRow, 1) & "')"
            sSql = "INSERT INTO LohdTable (LOTNUMBER,LOTUSERLOTID," _
                   & "LOTPARTREF,LOTPDATE,LOTORIGINALQTY,LOTREMAININGQTY," _
                   & "LOTPO,LOTPOITEM,LOTPOITEMREV) VALUES('" & sLotNumber _
                   & "','" & userLotNo _
                   & "','" & sPartNumber(iRow) & "','" & sDate & "'," _
                   & cAdjActualQty & ",0," & Val(lblPon) & "," _
                   & Val(vItems(iRow, 0)) & ",'" & vItems(iRow, 1) & "')"
                   
            clsADOCon.ExecuteSql sSql
            
            'If it's not a Type 7 Part then create a Pick Record
            If Val(vItems(iRow, VITEMS_PARTTYPE)) <> 7 Then
               sSql = ""
               iPkRecord = GetNextPickRecord(Trim$(vItems(iRow, 11)), Val(vItems(iRow, 12)))
               sSql = "INSERT INTO MopkTable (PKPARTREF," _
                      & "PKMOPART,PKMORUN,PKTYPE,PKPDATE,PKADATE," _
                      & "PKPQTY,PKAQTY,PKRECORD,PKUNITS,PKCOMT) VALUES('" _
                      & sPartNumber(iRow) & "','" & Trim(vItems(iRow, 11)) & "'," _
                      & Trim(vItems(iRow, 12)) & ",10,'" & Format(ES_SYSDATE, "mm/dd/yy") _
                      & "','" & Format(ES_SYSDATE, "mm/dd/yy") _
                      & "'," & cAdjActualQty & "," & cAdjActualQty & "," _
                      & iPkRecord & ",'" & vItems(iRow, 21) & "','')"
                 
               clsADOCon.ExecuteSql sSql
               
               'update runstatus to partial pick (PP) or pick complete (PC)
               Dim sMOStatus As String
               sSql = "select count(*) from MopkTable" & vbCrLf _
                  & "where PKAQTY=0 AND PKTYPE=" & IATYPE_PickOpenItem & vbCrLf _
                  & "AND PKMOPART='" & Trim(vItems(iRow, 11)) & "' AND PKMORUN=" & Trim(vItems(iRow, 12)) & vbCrLf _
                  & " AND RTRIM(PKPARTREF) <> '" & sPartNumber(iRow) & "'"
               
               sMOStatus = "PC"
               Dim rdo As ADODB.Recordset
               If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
                  If CLng(rdo.Fields(0)) > 0 Then
                     sMOStatus = "PP"
                  End If
               End If
               rdo.Close
               Set rdo = Nothing
               
               '2/28/07
               sSql = "UPDATE RunsTable SET RUNSTATUS='" & sMOStatus & "' WHERE (RUNREF='" _
                      & Trim(vItems(iRow, 11)) & "' AND RUNNO=" & vItems(iRow, 12) _
                      & " AND RUNSTATUS<>'PC')"
               clsADOCon.ExecuteSql sSql
               
               '10/29/04 if it is on a PickList
               sSql = "DELETE FROM MopkTable WHERE (PKPARTREF='" & sPartNumber(iRow) _
                      & "' AND PKMOPART='" & Trim(vItems(iRow, 11)) & "' AND " _
                      & "PKMORUN=" & Trim(vItems(iRow, 12)) & " AND PKAQTY=0 AND " _
                      & "PKRECORD<" & iPkRecord & ")"
               clsADOCon.ExecuteSql sSql
            End If
         Else
'            Dim lotUserLotID As String
'            If Not UsingSheetInventory Then
'                lotUserLotID = lotUserLotID & "-" & Format(ES_SYSDATE, "mm/dd/yyy")
'            End If
                
            
'            sSql = "INSERT INTO LohdTable (LOTNUMBER,LOTUSERLOTID," _
'                   & "LOTPARTREF,LOTPDATE,LOTORIGINALQTY,LOTREMAININGQTY," _
'                   & "LOTPO,LOTPOITEM,LOTPOITEMREV,LOTUSER,LOTUNITCOST) " _
'                   & "VALUES('" _
'                   & sLotNumber & "','" & sPon & "-" & Format(ES_SYSDATE, "mm/dd/yyy") & "','" & sPartNumber(iRow) _
'                   & "','" & sDate & "'," & cAdjActualQty & "," & cAdjActualQty _
'                   & "," & Val(lblPon) & "," & Val(vItems(iRow, 0)) & ",'" _
'                   & vItems(iRow, 1) & "','" & sInitials & "'," & cAdjCost & ")"
            sSql = "INSERT INTO LohdTable (LOTNUMBER,LOTUSERLOTID," _
                   & "LOTPARTREF,LOTPDATE,LOTADATE,LOTORIGINALQTY,LOTREMAININGQTY," _
                   & "LOTPO,LOTPOITEM,LOTPOITEMREV,LOTUSER,LOTUNITCOST) " _
                   & "VALUES('" _
                   & sLotNumber & "','" & userLotNo & "','" & sPartNumber(iRow) _
                   & "','" & sDate & "','" & sDate & "'," & cAdjActualQty & "," & cAdjActualQty _
                   & "," & Val(lblPon) & "," & Val(vItems(iRow, 0)) & ",'" _
                   & vItems(iRow, 1) & "','" & sInitials & "'," & cAdjCost & ")"
            clsADOCon.ExecuteSql sSql
         End If
         sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
                & "LOITYPE,LOIPARTREF,LOIPDATE,LOIADATE,LOIQUANTITY," _
                & "LOIPONUMBER,LOIPOITEM,LOIPOREV,LOIVENDOR," _
                & "LOIACTIVITY,LOICOMMENT,LOIUNITS,LOIUSER) " _
                & "VALUES('" _
                & sLotNumber & "',1,15,'" & sPartNumber(iRow) _
                & "','" & sDate & "','" & sDate & "'," & cAdjActualQty _
                & "," & Val(lblPon) & "," & vItems(iRow, 0) & ",'" _
                & vItems(iRow, 1) & "','" & sVendor & "'," _
                & lCOUNTER & ",'" _
                & "Purchase Order Receipt" & "','" & vItems(iRow, 21) & "','" & sInitials & "')"
         clsADOCon.ExecuteSql sSql
         
         'Service PO allocated to a Manufacturing Order Or Part Purchased to MO
         If Len(Trim(vItems(iRow, 11))) > 0 And Val(vItems(iRow, 12)) > 0 Then
            'Allocated to a run then create another activity
            'Added 12/13/01
            '1/18/04 added INMOPART
            lCOUNTER = lCOUNTER + 1
            cPickQty = cAdjActualQty
            sMon = "MO Num " & vItems(iRow, 11) & " Run " & vItems(iRow, 12)
            sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2," _
                   & "INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT,INPONUMBER," _
                   & "INPORELEASE,INPOITEM,INPOREV,INPDATE,INNUMBER,INLOTNUMBER," _
                   & "INUSER,INMOPART,INMORUN) VALUES(10,'" & sPartNumber(iRow) & "'," _
                   & "'MO ALLOC','" & sMon & "',-" & cPickQty & ",-" & cPickQty & "," _
                   & cAdjCost & ",'" & sCreditAcct & "','" & sDebitAcct & "'," _
                   & Val(lblPon) & "," & Val(lblRel) & "," & Val(vItems(iRow, 0)) _
                   & ",'" & vItems(iRow, 1) & "','" & sDate & "'," & lCOUNTER & ",'" _
                   & sLotNumber & "','" & sInitials & "','" & vItems(iRow, 11) & "'," _
                   & vItems(iRow, 12) & ")"
            clsADOCon.ExecuteSql sSql
            
            'Need another Lot Transaction to Auto Pick
            sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
                   & "LOITYPE,LOIPARTREF,LOIPDATE,LOIQUANTITY," _
                   & "LOIMOPARTREF,LOIMORUNNO,LOIPOREV,LOIVENDOR," _
                   & "LOIACTIVITY,LOICOMMENT,LOIUNITS) " _
                   & "VALUES('" _
                   & sLotNumber & "',2,10,'" & sPartNumber(iRow) _
                   & "','" & sDate & "',-" & cAdjActualQty _
                   & ",'" & vItems(iRow, 11) & "'," & vItems(iRow, 12) & "," _
                   & "'','" & sVendor & "'," & lCOUNTER & ",'" & "MO Auto Pick" & "','" _
                   & vItems(iRow, 21) & "')"
            clsADOCon.ExecuteSql sSql
         
         
         Else 'Corrected 2/28/01
            sSql = "UPDATE PartTable SET PAQOH=PAQOH+" & Abs(cAdjActualQty) & "," _
                   & "PALOTQTYREMAINING=PALOTQTYREMAINING+" & Abs(cAdjActualQty) & " " _
                   & "WHERE PARTREF='" & sPartNumber(iRow) & "'"
            clsADOCon.ExecuteSql sSql
         End If
         sSql = "UPDATE PoitTable SET PITYPE = " & IATYPE_PoReceipt & "," & vbCrLf _
            & "PIADATE = '" & sDate & "',PIAMT = " & cCost & ", PIAQTY=" & cQuantity & "," & vbCrLf _
            & "PILOTNUMBER = '" & sLotNumber & "'," & vbCrLf _
            & "PIODDELPSNUMBER = '" & vItems(iRow, VITEMS_PS) & "'" & vbCrLf _
            & "WHERE PINUMBER=" & Val(lblPon) & vbCrLf _
            & "AND PIRELEASE=" & Val(lblRel) & " AND PIITEM=" & vItems(iRow, 0) & vbCrLf _
            & "AND PIREV='" & vItems(iRow, 1) & "'"
         clsADOCon.ExecuteSql sSql
         
         'Is there a rejected Quantity? Need To Split that
         cRejQty = Val(vItems(iRow, 18))
         
         'Don't split it if 0, even if not checked
         If (cPlannedQty - cQuantity) <= 0 Then _
             vItems(iRow, 8) = 1
         If cQuantity >= cPlannedQty Then
            vItems(iRow, 8) = 1              'part is complete
         End If
         
             
         'Need a split?
         If Val(vItems(iRow, 8)) = 0 Then
            'if it was received and canceled
            sSql = "SELECT MAX(PIREV) FROM PoitTable WHERE " _
                   & "PINUMBER=" & Val(lblPon) & " And PIITEM= " _
                   & vItems(iRow, 0) & " "
            If clsADOCon.GetDataSet(sSql, rdoRev, ES_FORWARD) Then
               sRev = Trim(Left(rdoRev.Fields(0), 1))
               If sRev = "" Then
                  sRev = "A"
               Else
                  sRev = Chr(Asc(sRev) + 1)
               End If
            Else
               sRev = "A"
            End If
            sSql = "UPDATE PoitTable SET PIPQTY=" & str(cQuantity) & " " _
                   & "WHERE PINUMBER=" & Val(lblPon) & " AND PIRELEASE=" & Val(lblRel) & " " _
                   & "AND (PIITEM=" & vItems(iRow, 0) & " AND PIREV='" & vItems(iRow, 1) & "')"
            clsADOCon.ExecuteSql sSql
            
            'cQuantity = cPlannedQty - cQuantity
            Dim remainingQty As Currency
            remainingQty = cPlannedQty - cQuantity
            
            sSql = "INSERT INTO PoitTable (PINUMBER,PIRELEASE,PIITEM," _
                   & "PIREV,PITYPE,PIPART,PIPDATE,PIPQTY,PIESTUNIT," _
                   & "PIRUNPART,PIRUNNO,PIRUNOPNO,PICOMT,PIVENDOR) " _
                   & "VALUES(" & Val(lblPon) & "," & Val(lblRel) & "," & vItems(iRow, 0) _
                   & ",'" & sRev & "'," & "14,'" & sPartNumber(iRow) & "','" _
                   & vItems(iRow, 9) & "'," & remainingQty & "," & cCost _
                   & ",'" & vItems(iRow, 11) & "'," & vItems(iRow, 12) & "," _
                   & vItems(iRow, 13) & ",'" & vItems(iRow, 14) & "','" & sVendor & "')"
            clsADOCon.ExecuteSql sSql
         End If
         
         'if standard cost is being used, update inventory activity to that now
         'in any case, update lot and ia costs
         Dim ia As New ClassInventoryActivity
         Dim sReturnMsg As String
         sReturnMsg = ia.UpdateReceiptCosts(sPartNumber(iRow), CLng(lblPon), _
                      Val(vItems(iRow, 0)), CStr(vItems(iRow, 1)), cAdjActualQty, _
                      cAdjCost, False)
         If InStr(1, sReturnMsg, "failed") > 0 Then
            clsADOCon.RollbackTrans
            Exit Sub
         End If
         
         clsADOCon.CommitTrans      ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        
        'lot tracked?  if so, allow user to edit lot
         If Val(vItems(iRow, 19)) = 1 Then
            '4/17/06 The following to prevent LotExit from locking the system out
            'If Err = 0 Then
               
               'RdoCon.CommitTrans         ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
               
               MsgBox "Item " & vItems(iRow, 0) & vItems(iRow, 1) & " for part " & vItems(iRow, 3) _
                  & " is lot tracked.  You may edit the lot now.", _
                  vbInformation, Caption
               LotEdit.optPo.Value = vbChecked
               LotEdit.VENDOR = sVendor
               LotEdit.DebitAccount = sDebitAcct
               LotEdit.CreditAccount = sDebitAcct
               LotEdit.poNumber = lblPon
               'LotEdit.STDCOST = cCost
               LotEdit.STDCOST = cAdjCost
               LotEdit.poItem = vItems(iRow, 0)
               LotEdit.poRev = vItems(iRow, 1)
               LotEdit.INVACTIVITY = lCOUNTER
               LotEdit.txtLong = "PO Receipt"
               'LotEdit.lblOrigQty = Format(vItems(iRow, VITEMS_ActualQty), ES_PurchasedDataFormat)
               LotEdit.lblOrigQty = Format(cAdjActualQty, ES_PurchasedDataFormat)
'               LotEdit.txtlot = sPon & "-" & Format(ES_SYSDATE, "mm/dd/yyyy")
               LotEdit.txtlot = userLotNo
               LotEdit.lblPart = vItems(iRow, 3)
               'LotEdit.lblDate = Format(ES_SYSDATE, "mm/dd/yy")
               LotEdit.lblDate = RecvRVe01a.txtRcd
               LotEdit.lblTime = Format(ES_SYSDATE, "hh:mm")
               LotEdit.lblNumber = sLotNumber
               LotEdit.Show vbModal
               GoTo GNext1
            'End If
         End If
         'If Err = 0 Then
'            RdoCon.CommitTrans      ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
         'Else
'            iFailed = iFailed + 1
'            RdoCon.RollbackTrans
'            MsgBox "Receipt Of " & lblPon & "-" & vItems(iRow, 0) & vItems(iRow, 1) _
'               & " Was Not Successful.", vbInformation, "PO Receipt"
'         End If
GNext1:  'Err.Clear
      End If
   Next
   MouseCursor 0
   'If Err = 0 Then
      UpdateWipColumns lSysCount
      If iFailed > 0 Then
         MsgBox iFailed & " Items Were Not Received..", _
            vbInformation, Caption
      End If
      
      MsgBox "Purchase Order Items Have Been Received.", vbInformation, Caption
      MouseCursor 13
      If RecvRVe01a.optRcv.Value = vbChecked Then
         RecvRVe01a.optRcv.Value = vbUnchecked
      End If
'   Else
'      sProcName = "receiveit"
'      CurrError.Number = Err.Number
'      CurrError.Description = Err.Description
'      sMsg = "Could Not Successfully Receive Some Or All Of The PO Items."
'      MsgBox sMsg, vbExclamation, Caption
'   End If
   Set rdoRev = Nothing
   
   'Unload Me
   Exit Sub
   
RitmsPi1:
   sProcName = "receiveit"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   On Error Resume Next
   clsADOCon.RollbackTrans
   MouseCursor 0
   DoModuleErrors Me
   Unload Me
   
End Sub

Private Sub ClearBoxes()
   Dim iList As Integer
   For iList = 1 To 6
      lblItm(iList).Visible = False
      lblRev(iList).Visible = False
      lblPrt(iList).Visible = False
      lblDsc(iList).Visible = False
      lblPqt(iList).Visible = False
      lblDueDate(iList).Visible = False
      txtAqt(iList).Visible = False
      cmdCmt(iList).Visible = False
      optPck(iList).Visible = False
      optLbl(iList).Visible = False
      OnDock(iList).Visible = False
      txtPS(iList).Visible = False
   Next
   
End Sub

Private Sub GetExtendedInfo(iThisindex As Integer)
   Dim RdoPrt As ADODB.Recordset
   Dim sExtendedInfo As String  'BBS Added this on 03/17/2010 for Ticket #27859

   On Error GoTo DiaErr1
   If FormCount("frmTip") > 0 Then Unload frmTip 'BBS Added for Ticket #80670
   sExtendedInfo = ""   'BBS Added this on 03/17/2010 for Ticket #27859
   
   AdoQry2.Parameters(0).Value = Val(lblItm(iThisindex))
   AdoQry2.Parameters(1).Value = Trim(lblRev(iThisindex))
   bSqlRows = clsADOCon.GetQuerySet(RdoPrt, AdoQry2, ES_KEYSET, False, 1)
   If bSqlRows Then
      With RdoPrt
        'BBS 03/17/2010 for Ticket #27859
        ' I removed all the references to the label and am now assigning this to a string that\
        ' will be passed into a fake/pseudo tool tip function
         'lblInfo.Top = lblPrt(iThisindex).Top + 300
         'lblInfo.Visible = True
         sExtendedInfo = "Extended Part Description: " & vbCr _
                   & Trim(!PAEXTDESC) & vbCr _
                   & vbCr & "Item Comments: " & vbCr _
                   & Trim(!PICOMT)
         'lblInfo.Refresh
         'Sleep 2000
         'lblInfo.Visible = False
         'lblInfo.Top = 0
         ClearResultSet RdoPrt
      End With
   End If
   Set RdoPrt = Nothing
   
   'BBS Added this on 03/17/2010 for Ticket #27859
   If (sExtendedInfo <> "") Then
        If FormCount("frmTip") = 0 Then Load frmTip
        frmTip.ShowMoreInfo sExtendedInfo
   End If
   
   Exit Sub
   
DiaErr1:
   'Ingore Errors on this routine
   
End Sub



'1/30/03

Private Function GetRejInfo(iList As Integer, sODDate As String, sODRejTag As String, _
                            sODInspector As String, sODComment As String) As Byte
   Dim RdoRej As ADODB.Recordset
   On Error Resume Next
   sSql = "SELECT PINUMBER,PIITEM,PIREV,PIONDOCKINSPDATE,PIONDOCKREJTAG," _
          & "PIONDOCKINSPECTOR,PIONDOCKCOMMENT FROM PoitTable WHERE " _
          & "(PINUMBER=" & Val(lblPon) & " And PIITEM=" & vItems(iList, 0) & " AND " _
          & "PIREV='" & Trim(vItems(iList, 1)) & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRej, ES_FORWARD)
   If bSqlRows Then
      With RdoRej
         sODDate = "" & Trim(Format$(!PIONDOCKINSPDATE))
         sODRejTag = "" & Trim(!PIONDOCKREJTAG)
         sODInspector = "" & Trim(!PIONDOCKINSPECTOR)
         sODComment = "" & Trim(!PIONDOCKCOMMENT)
         GetRejInfo = 1
         ClearResultSet RdoRej
      End With
   Else
      sODDate = ""
      sODRejTag = ""
      sODInspector = ""
      sODComment = ""
      GetRejInfo = 0
   End If
    Set RdoRej = Nothing
End Function

'Private Function AdjustForPurchaseUnits(sPartRef As String, cPlanQty As Currency, cActQty As Currency, cUnitCost As Currency) As Integer
'    'adjusts quantities and cost for PO receipt where purchase and inventory units are different
'    'return = conversion quantity if it is used
'    'return = 1 if no conversion required
'
'    'Dim conversion  As Integer
'    'AdjustForPurchaseUnits = 0
'
''    If UsingSheetInventory Then
''        Dim rdo As ADODB.Recordset
''        sSql = "select PAPURCONV from PartTable where PARTREF = '" + sPartRef + "'"
''        bSqlRows = clsADOCon.GetDataSet(sSql, rdo)
''        If bSqlRows Then
''            conversion = rdo("PAPURCONV")
''            If conversion <> 0 And conversion <> 1 Then
''                cPlanQty = cPlanQty * conversion
''                cActQty = cActQty * conversion
''                cUnitCost = CCur(Format(cUnitCost / conversion, ES_QuantityDataFormat))
''                'cUnitCost = cUnitCost / conversion
''                AdjustForPurchaseUnits = conversion
''            End If
''        End If
''        Set rdo = Nothing
''    End If
'
'    AdjustForPurchaseUnits = GetPurchToInvConversionQty(sPartRef)
'    If AdjustForPurchaseUnits <> 1 Then
'        cPlanQty = cPlanQty * AdjustForPurchaseUnits
'        cActQty = cActQty * AdjustForPurchaseUnits
'        cUnitCost = CCur(Format(cUnitCost / AdjustForPurchaseUnits, ES_QuantityDataFormat))
'    End If
'
'End Function

Private Sub txtPS_LostFocus(Index As Integer)
   txtPS(Index) = CheckLen(txtPS(Index), 20)
   vItems(Index + iIndex, VITEMS_PS) = txtPS(Index).Text
End Sub
