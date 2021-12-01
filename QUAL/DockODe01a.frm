VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form DockODe01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "On Dock Inspection"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCmt 
      DownPicture     =   "DockODe01a.frx":0000
      Height          =   320
      Index           =   4
      Left            =   3060
      Picture         =   "DockODe01a.frx":04DA
      Style           =   1  'Graphical
      TabIndex        =   77
      TabStop         =   0   'False
      ToolTipText     =   "PO Item Comments"
      Top             =   5400
      Width           =   350
   End
   Begin VB.CommandButton cmdCmt 
      DownPicture     =   "DockODe01a.frx":0ADC
      Height          =   320
      Index           =   3
      Left            =   3060
      Picture         =   "DockODe01a.frx":0FB6
      Style           =   1  'Graphical
      TabIndex        =   76
      TabStop         =   0   'False
      ToolTipText     =   "PO Item Comments"
      Top             =   4320
      Width           =   350
   End
   Begin VB.CommandButton cmdCmt 
      DownPicture     =   "DockODe01a.frx":15B8
      Height          =   320
      Index           =   2
      Left            =   3060
      Picture         =   "DockODe01a.frx":1A92
      Style           =   1  'Graphical
      TabIndex        =   75
      TabStop         =   0   'False
      ToolTipText     =   "PO Item Comments"
      Top             =   3240
      Width           =   350
   End
   Begin VB.CommandButton cmdCmt 
      DownPicture     =   "DockODe01a.frx":2094
      Height          =   320
      Index           =   1
      Left            =   3060
      Picture         =   "DockODe01a.frx":256E
      Style           =   1  'Graphical
      TabIndex        =   74
      TabStop         =   0   'False
      ToolTipText     =   "PO Item Comments"
      Top             =   2160
      Width           =   350
   End
   Begin VB.TextBox txtWaste 
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   8100
      TabIndex        =   24
      Tag             =   "1"
      ToolTipText     =   "Rejected Quantity"
      Top             =   5055
      Width           =   1095
   End
   Begin VB.TextBox txtWaste 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   8100
      TabIndex        =   18
      Tag             =   "1"
      ToolTipText     =   "Rejected Quantity"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox txtWaste 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   8100
      TabIndex        =   12
      Tag             =   "1"
      ToolTipText     =   "Rejected Quantity"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtWaste 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   8100
      TabIndex        =   6
      Tag             =   "1"
      ToolTipText     =   "Rejected Quantity"
      Top             =   1785
      Width           =   1095
   End
   Begin VB.CommandButton cmdDn 
      DisabledPicture =   "DockODe01a.frx":2B70
      DownPicture     =   "DockODe01a.frx":3062
      Height          =   372
      Left            =   7890
      MaskColor       =   &H00000000&
      Picture         =   "DockODe01a.frx":3554
      Style           =   1  'Graphical
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   6120
      Width           =   400
   End
   Begin VB.CommandButton cmdUp 
      DisabledPicture =   "DockODe01a.frx":3A46
      DownPicture     =   "DockODe01a.frx":3F38
      Height          =   372
      Left            =   7440
      MaskColor       =   &H00000000&
      Picture         =   "DockODe01a.frx":442A
      Style           =   1  'Graphical
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   6120
      Width           =   400
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "DockODe01a.frx":491C
      Style           =   1  'Graphical
      TabIndex        =   70
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optInc 
      Alignment       =   1  'Right Justify
      Caption         =   "Show Completed "
      Height          =   195
      Left            =   4320
      TabIndex        =   2
      Top             =   400
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "C&ancel"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7680
      TabIndex        =   31
      ToolTipText     =   "Cancel This Operation"
      Top             =   1200
      Width           =   875
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   3120
      TabIndex        =   1
      Tag             =   "4"
      Top             =   360
      Width           =   1095
   End
   Begin VB.CheckBox chkComplete 
      Caption         =   "    "
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   315
      Index           =   4
      Left            =   9300
      TabIndex        =   25
      ToolTipText     =   "Check If Completed"
      Top             =   5040
      Width           =   735
   End
   Begin VB.CheckBox chkComplete 
      Caption         =   "    "
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   315
      Index           =   3
      Left            =   9300
      TabIndex        =   19
      ToolTipText     =   "Check If Completed"
      Top             =   3960
      Width           =   735
   End
   Begin VB.CheckBox chkComplete 
      Caption         =   "    "
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   315
      Index           =   2
      Left            =   9300
      TabIndex        =   13
      ToolTipText     =   "Check If Completed"
      Top             =   2880
      Width           =   735
   End
   Begin VB.CheckBox chkComplete 
      Caption         =   "    "
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   315
      Index           =   1
      Left            =   9300
      TabIndex        =   7
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdItm 
      Cancel          =   -1  'True
      Caption         =   "&Select"
      Height          =   315
      Left            =   6240
      TabIndex        =   3
      ToolTipText     =   "Fill With PO Items"
      Top             =   360
      Width           =   875
   End
   Begin VB.TextBox txtAccepted 
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   5640
      TabIndex        =   22
      Tag             =   "1"
      ToolTipText     =   "Accepted Quantity"
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox txtRejected 
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   6960
      TabIndex        =   23
      Tag             =   "1"
      ToolTipText     =   "Rejected Quantity"
      Top             =   5040
      Width           =   1095
   End
   Begin VB.ComboBox cmbIns 
      Enabled         =   0   'False
      Height          =   315
      Index           =   4
      Left            =   1080
      Sorted          =   -1  'True
      TabIndex        =   26
      Tag             =   "3"
      ToolTipText     =   "Select Inspector ID From List"
      Top             =   5400
      Width           =   1665
   End
   Begin VB.TextBox txtCmt 
      Enabled         =   0   'False
      Height          =   615
      Index           =   4
      Left            =   4560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   27
      Tag             =   "9"
      ToolTipText     =   "Comments.  Multiple Line"
      Top             =   5400
      Width           =   4635
   End
   Begin VB.TextBox txtAccepted 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   5640
      TabIndex        =   16
      Tag             =   "1"
      ToolTipText     =   "Accepted Quantity"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtRejected 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   6960
      TabIndex        =   17
      Tag             =   "1"
      ToolTipText     =   "Rejected Quantity"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.ComboBox cmbIns 
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   1080
      Sorted          =   -1  'True
      TabIndex        =   20
      Tag             =   "3"
      ToolTipText     =   "Select Inspector ID From List"
      Top             =   4320
      Width           =   1665
   End
   Begin VB.TextBox txtCmt 
      Enabled         =   0   'False
      Height          =   615
      Index           =   3
      Left            =   4560
      MultiLine       =   -1  'True
      TabIndex        =   21
      Tag             =   "9"
      ToolTipText     =   "Comments.  Multiple Line"
      Top             =   4320
      Width           =   4635
   End
   Begin VB.TextBox txtAccepted 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   5640
      TabIndex        =   10
      Tag             =   "1"
      ToolTipText     =   "Accepted Quantity"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtRejected 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   6960
      TabIndex        =   11
      Tag             =   "1"
      ToolTipText     =   "Rejected Quantity"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.ComboBox cmbIns 
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   1080
      Sorted          =   -1  'True
      TabIndex        =   14
      Tag             =   "3"
      ToolTipText     =   "Select Inspector ID From List"
      Top             =   3240
      Width           =   1665
   End
   Begin VB.TextBox txtCmt 
      Enabled         =   0   'False
      Height          =   615
      Index           =   2
      Left            =   4560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Tag             =   "9"
      ToolTipText     =   "Comments.  Multiple Line"
      Top             =   3240
      Width           =   4635
   End
   Begin VB.TextBox txtAccepted 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   5640
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "Accepted Quantity"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtRejected 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   6960
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "Rejected Quantity"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.ComboBox cmbIns 
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   1080
      Sorted          =   -1  'True
      TabIndex        =   8
      Tag             =   "3"
      ToolTipText     =   "Select Inspector ID From List"
      Top             =   2160
      Width           =   1665
   End
   Begin VB.TextBox txtCmt 
      Enabled         =   0   'False
      Height          =   615
      Index           =   1
      Left            =   4560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Tag             =   "9"
      ToolTipText     =   "Comments.  Multiple Line"
      Top             =   2160
      Width           =   4635
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   8640
      TabIndex        =   30
      ToolTipText     =   "Update Current List Of Items"
      Top             =   1200
      Width           =   875
   End
   Begin VB.Frame z2 
      Height          =   60
      Left            =   240
      TabIndex        =   39
      Top             =   1080
      Width           =   9285
   End
   Begin VB.TextBox lblVendor 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   720
      Width           =   1695
   End
   Begin VB.ComboBox cmbPon 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Contains Only PO's With On Dock Requirements"
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   8160
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   600
      Top             =   6600
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6660
      FormDesignWidth =   10200
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Waste             "
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
      Index           =   8
      Left            =   8100
      TabIndex        =   73
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Image Dsup 
      Height          =   300
      Left            =   3240
      Picture         =   "DockODe01a.frx":50CA
      Top             =   6360
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Enup 
      Height          =   300
      Left            =   3720
      Picture         =   "DockODe01a.frx":55BC
      Top             =   6360
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Endn 
      Height          =   300
      Left            =   3480
      Picture         =   "DockODe01a.frx":5AAE
      Top             =   6360
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Dsdn 
      Height          =   300
      Left            =   3000
      Picture         =   "DockODe01a.frx":5FA0
      Top             =   6360
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label lblUom 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   5280
      TabIndex        =   69
      ToolTipText     =   "Click To Insert Quantity"
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label lblUom 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   5280
      TabIndex        =   68
      ToolTipText     =   "Click To Insert Quantity"
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label lblUom 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   5280
      TabIndex        =   67
      ToolTipText     =   "Click To Insert Quantity"
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label lblUom 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   5280
      TabIndex        =   66
      ToolTipText     =   "Click To Insert Quantity"
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   285
      Index           =   7
      Left            =   2520
      TabIndex        =   65
      Top             =   400
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comp?"
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
      Left            =   9225
      TabIndex        =   28
      Top             =   1575
      Width           =   615
   End
   Begin VB.Label lblCmt 
      BackStyle       =   0  'Transparent
      Caption         =   "On Dock Comment"
      Height          =   435
      Index           =   4
      Left            =   3780
      TabIndex        =   64
      Top             =   5460
      Width           =   735
   End
   Begin VB.Label lblCmt 
      BackStyle       =   0  'Transparent
      Caption         =   "On Dock Comment"
      Height          =   555
      Index           =   3
      Left            =   3780
      TabIndex        =   63
      Top             =   4380
      Width           =   795
   End
   Begin VB.Label lblCmt 
      BackStyle       =   0  'Transparent
      Caption         =   "On Dock Comment"
      Height          =   495
      Index           =   2
      Left            =   3780
      TabIndex        =   62
      Top             =   3300
      Width           =   795
   End
   Begin VB.Label lblCmt 
      BackStyle       =   0  'Transparent
      Caption         =   "On Dock Comment"
      Height          =   555
      Index           =   1
      Left            =   3780
      TabIndex        =   61
      Top             =   2220
      Width           =   735
   End
   Begin VB.Label lblIns 
      BackStyle       =   0  'Transparent
      Caption         =   "Inspector"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   60
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label lblIns 
      BackStyle       =   0  'Transparent
      Caption         =   "Inspector"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   59
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label lblIns 
      BackStyle       =   0  'Transparent
      Caption         =   "Inspector"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   58
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label lblIns 
      BackStyle       =   0  'Transparent
      Caption         =   "Inspector"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   57
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblPqt 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   4095
      TabIndex        =   56
      ToolTipText     =   "Click To Insert Quantity"
      Top             =   5040
      Width           =   1065
   End
   Begin VB.Label lblPrt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   960
      TabIndex        =   55
      ToolTipText     =   "Part Description"
      Top             =   5040
      Width           =   3105
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   54
      Top             =   5040
      Width           =   375
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   53
      Top             =   5040
      Width           =   285
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   600
      TabIndex        =   52
      Top             =   3960
      Width           =   285
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   51
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label lblPrt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   960
      TabIndex        =   50
      ToolTipText     =   "Part Description"
      Top             =   3960
      Width           =   3105
   End
   Begin VB.Label lblPqt 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   4095
      TabIndex        =   49
      ToolTipText     =   "Click To Insert Quantity"
      Top             =   3960
      Width           =   1065
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   600
      TabIndex        =   48
      Top             =   2880
      Width           =   285
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   47
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label lblPrt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   960
      TabIndex        =   46
      ToolTipText     =   "Part Description"
      Top             =   2880
      Width           =   3105
   End
   Begin VB.Label lblPqt 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   4095
      TabIndex        =   45
      ToolTipText     =   "Click To Insert Quantity"
      Top             =   2880
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rejected          "
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
      Left            =   6960
      TabIndex        =   44
      Top             =   1575
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Accepted Qty"
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
      Left            =   5640
      TabIndex        =   43
      Top             =   1575
      Width           =   1095
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
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   42
      Top             =   1575
      Width           =   735
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
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   41
      Top             =   1575
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Qty             "
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
      Left            =   4080
      TabIndex        =   40
      Top             =   1575
      Width           =   1095
   End
   Begin VB.Label lblPqt 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   4095
      TabIndex        =   38
      ToolTipText     =   "Click To Insert Quantity"
      Top             =   1800
      Width           =   1065
   End
   Begin VB.Label lblPrt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   960
      TabIndex        =   37
      ToolTipText     =   "Part Description"
      Top             =   1800
      Width           =   3105
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   36
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   35
      Top             =   1800
      Width           =   285
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3120
      TabIndex        =   34
      Top             =   720
      Width           =   3915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   33
      Top             =   400
      Width           =   1095
   End
End
Attribute VB_Name = "DockODe01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'1/03   New
'8/9/04 Added Uom
'10/18/04 Corrected error in array and reformated Quantities
Option Explicit
Dim bCancel As Byte
Dim bOnLoad As Byte

Dim iCurrPage As Integer
Dim iIndex As Integer
Dim iLastPage As Integer
Dim iTotalItems As Integer

Dim vItems(100, 18) As Variant

' 1 = PO
' 2 = Item
' 3 = Item Rev
' 4 = Part Number
' 5 = PO Qty
Private Const POITEM_OrderedQty = 5
' 6 = Acc Qty
Private Const POITEM_AcceptedQty = 6
' 7 = Rej Qty
Private Const POITEM_RejectedQty = 7
' 8 = Rej Insp
' 9 = Rej Comment
'10 = Complete 0 or 1
Private Const POITEM_Complete = 10
'11 = ToolTipText
'12 = Del date
'13 = uom
Private Const POITEM_WastedQty = 14
Private Const POITEM_SplitForReject = 15

Private Const POITEM_DEL = 16
Private Const POITEM_DELQty = 17
Private Const POITEM_Comment = 18

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub FillInspectors()
   Dim RdoCmb As ADODB.Recordset
   Dim b As Byte
   On Error GoTo DiaErr1
   sSql = "Qry_FillInspectorsActive"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb)
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            For b = 1 To 4
               AddComboStr cmbIns(b).hwnd, "" & Trim(!INSID)
            Next
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   End If
   
   If cmbIns(1).ListCount = 0 Then
      MsgBox "You Are Required To Have A Least One " & vbCr _
         & "Inspector To Use This Function.", vbInformation, Caption
      bCancel = 1
   Else
      For b = 1 To 4
         cmbIns(b) = cmbIns(b).List(0)
      Next
   End If
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillinspe"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbIns_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub

Private Sub cmbIns_KeyPress(index As Integer, KeyAscii As Integer)
   KeyCase KeyAscii
   
End Sub


Private Sub cmbIns_LostFocus(index As Integer)
   Dim iList As Integer
   
   For iList = 0 To cmbIns(index).ListCount - 1
      If cmbIns(index) = cmbIns(index).List(iList) Then
         vItems(index + iIndex, 8) = cmbIns(index)
         Exit Sub
      End If
   Next
   MsgBox "That Inspector Was Not Found.", vbInformation, Caption
   cmbIns(index) = cmbIns(index).List(0)
   
End Sub

Private Sub cmbPon_Click()
   GetThisVendor
   
End Sub


Private Sub cmbPon_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   cmbPon = CheckLen(cmbPon, 6)
   cmbPon = Format(Abs(Val(cmbPon)), "000000")
   If bCancel Then Exit Sub
   For iList = 0 To cmbPon.ListCount - 1
      If cmbPon = cmbPon.List(iList) Then b = 1
   Next
   If b = 1 Then
      GetThisVendor
   Else
      Beep
      MsgBox "The Requested PO Does Not Exist Or Is Not Listed", _
         vbInformation, Caption
      If cmbPon.ListCount > 0 Then cmbPon = cmbPon.List(0)
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = 1
   
End Sub



Private Sub cmdCmt_Click(index As Integer)
   If cmdCmt(index) Then
      MouseCursor 13
      Load PoItemComments
      PoItemComments.lblPon = cmbPon
      PoItemComments.lblItem = lblItm(index)
      PoItemComments.lblRev = lblRev(index)
      PoItemComments.txtCmt = vItems(index, POITEM_Comment)
      PoItemComments.show
      PoItemComments.Left = (Me.Width - PoItemComments.Width) / 2
      PoItemComments.Top = (Me.Height - PoItemComments.Height) / 2
      
      cmdCmt(index) = False
   End If
   
End Sub

Private Sub showItem(indexItem As Integer, showThisItem As Boolean)
      lblItm(indexItem).Visible = showThisItem
      lblRev(indexItem).Visible = showThisItem
      lblPrt(indexItem).Visible = showThisItem
      lblPqt(indexItem).Visible = showThisItem
      txtAccepted(indexItem).Visible = showThisItem
      txtRejected(indexItem).Visible = showThisItem
      txtWaste(indexItem).Visible = showThisItem
      cmbIns(indexItem).Visible = showThisItem
      cmdCmt(indexItem).Visible = showThisItem
      txtCmt(indexItem).Visible = showThisItem
      lblIns(indexItem).Visible = showThisItem
      lblCmt(indexItem).Visible = showThisItem
      chkComplete(indexItem).Visible = showThisItem
      lblUom(indexItem).Visible = showThisItem
End Sub

Private Sub cmdDn_Click()
   'Next
   iCurrPage = iCurrPage + 1
   If iCurrPage > iLastPage Then iCurrPage = iLastPage
   GetTheNextGroup
   
End Sub

Private Sub cmdEnd_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   sMsg = "Are You Sure That You Want To Cancel Without" & vbCr _
          & "Saving Any Changes To The Data?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      ManageBoxes
      cmdUpd.Enabled = False
      cmdEnd.Enabled = False
      cmbPon.Enabled = True
      txtDte.Enabled = True
      cmdItm.Enabled = True
      On Error Resume Next
      cmbPon.SetFocus
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 6401
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdItm_Click()
   GetPoItems
   
End Sub


Private Sub cmdUp_Click()
   'last
   iCurrPage = iCurrPage - 1
   If iCurrPage < 1 Then iCurrPage = 1
   GetTheNextGroup
   
End Sub

Private Sub cmdUpd_Click()
   UpdateList
   
End Sub


Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillInspectors
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   If bCancel = 1 Then
      InspRTe06a.show
      Unload Me
   End If
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   bOnLoad = 1
   Dim i As Integer
   For i = 1 To 4
      showItem i, False
   Next
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set DockODe01a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   lblVendor.BackColor = Es_FormBackColor
   txtDte = Format(ES_SYSDATE, "mm/dd/yy")
   For b = 1 To 4
      lblPrt(b).ToolTipText = "ToolTip = Part Description"
   Next
   
End Sub

Private Sub FillCombo()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT PINUMBER,PIONDOCK,PONUMBER FROM PoitTable," _
          & "PohdTable Where(PINUMBER = PONUMBER AND PIONDOCK = 1 " _
          & "And PITYPE = 14) "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            AddComboStr cmbPon.hwnd, "" & Format(0 + !poNumber, "000000")
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
      If cmbPon.ListCount > 0 Then
         cmbPon = cmbPon.List(0)
         GetThisVendor
      End If
   End If
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetThisVendor()
   Dim RdoVnd As ADODB.Recordset
   sSql = "SELECT PONUMBER,POVENDOR,VEREF,VENICKNAME," _
          & "VEBNAME FROM PohdTable,VndrTable WHERE (VEREF=" _
          & "POVENDOR AND PONUMBER=" & Val(cmbPon) & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVnd, ES_FORWARD)
   If bSqlRows Then
      With RdoVnd
         lblVendor = "" & Trim(!VENICKNAME)
         lblName = "" & Trim(!VEBNAME)
         ClearResultSet RdoVnd
      End With
   Else
      lblVendor = ""
      lblName = "No Such PO Or PO Doesn't Qualify"
   End If
   Set RdoVnd = Nothing
   
End Sub


' 1 = PO
' 2 = Item
' 3 = Item Rev
' 4 = Part Number
' 5 = PO Qty
' 6 = Acc Qty
' 7 = Rej Qty
' 8 = Rej Insp
' 9 = Rej Comment
'10 = Complete 0 or 1
'11 = ToolTipText

Private Sub GetPoItems()
   Dim RdoGpi As ADODB.Recordset
   Dim iRow As Integer
   Dim rPages As Single
   
   iIndex = -1
   iTotalItems = 0
   On Error GoTo DiaErr1
   ManageBoxes
   iRow = 0
   sSql = "SELECT PINUMBER,PIITEM,PIREV,PITYPE,PIPART,PIPQTY," & vbCrLf _
          & "PIONDOC,PIONDOCKINSPECTED,PIONDOCKQTYACC,PIONDOCKQTYREJ,PIONDOCKQTYWASTE," & vbCrLf _
          & "PIONDOCKINSPECTOR,PIONDOCKCOMMENT,PIODDELDATE,PARTREF,PARTNUM," & vbCrLf _
          & "ISNULL(PIODDELIVERED, 0) PIODDELIVERED, ISNULL(PIODDELQTY, 0) PIODDELQTY, " & vbCrLf _
          & "PADESC,PAUNITS, PICOMT FROM PoitTable,PartTable" & vbCrLf _
          & "WHERE PIPART=PARTREF AND PITYPE=14 AND PIONDOCK=1" & vbCrLf _
          & "AND PINUMBER=" & Val(cmbPon) & " AND PIONDOCKINSPECTED = 0"
   'If optInc.Value = vbUnchecked Then sSql = sSql & "AND PIONDOCKINSPECTED=0"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGpi, ES_FORWARD)
   If bSqlRows Then
      With RdoGpi
         Do Until .EOF
            iRow = iRow + 1
            iTotalItems = iTotalItems + 1
            vItems(iRow, 1) = cmbPon
            vItems(iRow, 2) = !PIITEM
            vItems(iRow, 3) = Trim(!PIREV)
            vItems(iRow, 4) = Trim(!PartNum)
            vItems(iRow, POITEM_OrderedQty) = Format(!PIPQTY, ES_QuantityDataFormat)
            vItems(iRow, POITEM_AcceptedQty) = Format(!PIONDOCKQTYACC, ES_QuantityDataFormat)
            vItems(iRow, POITEM_RejectedQty) = Format(!PIONDOCKQTYREJ, ES_QuantityDataFormat)
            vItems(iRow, POITEM_WastedQty) = Format(!PIONDOCKQTYWASTE, ES_QuantityDataFormat)
            vItems(iRow, 8) = "" & Trim(!PIONDOCKINSPECTOR)
            If Trim(vItems(iRow, 8)) = "" Then vItems(iRow, 8) = cmbIns(1)
            vItems(iRow, 9) = "" & Trim(!PIONDOCKCOMMENT)
            vItems(iRow, POITEM_Complete) = !PIONDOCKINSPECTED
            vItems(iRow, 11) = "" & Trim(!PADESC)
            vItems(iRow, 13) = "" & Trim(!PAUNITS)
            If Not IsNull(!PIODDELDATE) Then
               vItems(iRow, 12) = "'" & Format(!PIODDELDATE, "mm/dd/yy") & "'"
            Else
               vItems(iRow, 12) = "Null"
            End If

            vItems(iRow, POITEM_DEL) = Trim(!PIODDELIVERED)
            vItems(iRow, POITEM_DELQty) = Format(!PIODDELQTY, ES_QuantityDataFormat)
            vItems(iRow, POITEM_Comment) = !PICOMT
                        
'            If iRow < 5 Then
'               lblItm(iRow).Visible = True
'               lblRev(iRow).Visible = True
'               lblPrt(iRow).Visible = True
'               lblPqt(iRow).Visible = True
'               txtAccepted(iRow).Visible = True
'               txtRejected(iRow).Visible = True
'               txtWaste(iRow).Visible = True
'               cmbIns(iRow).Visible = True
'               txtCmt(iRow).Visible = True
'               lblIns(iRow).Visible = True
'               lblCmt(iRow).Visible = True
'               chkComplete(iRow).Visible = True
'               lblUom(iRow).Visible = True
'            End If
            .MoveNext
         Loop
         ClearResultSet RdoGpi
      End With
   End If
   
   If iTotalItems > 4 Then
      cmdUp.Enabled = True
      cmdUp.Picture = Enup.Picture
      cmdDn.Enabled = True
      cmdDn.Picture = Endn.Picture
   End If
   If iTotalItems > 0 Then
      iLastPage = 0.4 + (iTotalItems / 4)
'      txtAccepted(1).Enabled = True
      cmdItm.Enabled = False
      cmbPon.Enabled = False
      txtDte.Enabled = False
      cmdUpd.Enabled = True
      cmdEnd.Enabled = True
   
      For iRow = 1 To 4
         txtAccepted(iRow).Enabled = True
         txtRejected(iRow).Enabled = True
         txtWaste(iRow).Enabled = True
         cmbIns(iRow).Enabled = True
         cmdCmt(iRow).Enabled = True
         txtCmt(iRow).Enabled = True
         lblUom(iRow).Enabled = True
      Next
   
   End If
   Set RdoGpi = Nothing
   
   iIndex = 0
   iCurrPage = 1
   GetTheNextGroup
   
   Exit Sub
   
DiaErr1:
   sProcName = "getpoitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub GetTheNextGroup()
   Dim i As Integer
   Dim iRow As Integer
   iIndex = (iCurrPage - 1) * 4
   ManageBoxes
   
   For iRow = iIndex + 1 To iTotalItems
      i = i + 1
      If i > 4 Then Exit For
'      lblItm(I).Visible = True
'      lblRev(I).Visible = True
'      'lblItm(i).Visible = True
'      lblPrt(I).Visible = True
'      lblPqt(I).Visible = True
'      txtAccepted(I).Visible = True
'      txtRejected(I).Visible = True
'      txtWaste(I).Visible = True
'      cmbIns(I).Visible = True
'      cmdCmt(I).Visible = True
'      txtCmt(I).Visible = True
'      lblIns(I).Visible = True
'      lblCmt(I).Visible = True
'      chkComplete(I).Visible = True
'      lblUom(I).Visible = True

      showItem i, True
      lblItm(i) = vItems(iRow, 2)
      lblRev(i) = vItems(iRow, 3)
      lblPrt(i) = vItems(iRow, 4)
      lblPrt(i).ToolTipText = vItems(iRow, 11)
      lblPqt(i) = Format(vItems(iRow, POITEM_OrderedQty), ES_QuantityDataFormat)
      txtAccepted(i) = Format(vItems(iRow, POITEM_AcceptedQty), ES_QuantityDataFormat)
      txtRejected(i) = Format(vItems(iRow, POITEM_RejectedQty), ES_QuantityDataFormat)
      txtWaste(i) = Format(vItems(iRow, POITEM_WastedQty), ES_QuantityDataFormat)
      cmbIns(i) = vItems(iRow, 8)
      txtCmt(i) = vItems(iRow, 9)
      chkComplete(i).Value = vItems(iRow, POITEM_Complete)
      lblUom(i) = vItems(iRow, 13)
      
      CheckBalance i
   Next
   
End Sub





Private Sub lblPqt_Click(index As Integer)
   txtAccepted(index) = lblPqt(index)
   If Val(txtAccepted(index)) >= Val(lblPqt(index)) Then chkComplete(index).Value = vbChecked
   vItems(index + iIndex, POITEM_AcceptedQty) = Val(txtAccepted(index))
   vItems(index + iIndex, POITEM_Complete) = chkComplete(index).Value
   
End Sub


Private Sub txtAccepted_GotFocus(index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtAccepted_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtAccepted_KeyPress(index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub

Private Sub txtAccepted_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub

Private Sub txtAccepted_Validate(index As Integer, Cancel As Boolean)
   txtAccepted(index) = CheckLen(txtAccepted(index), 9)
   txtAccepted(index) = Format(Abs(Val(txtAccepted(index))), ES_QuantityDataFormat)
   If Val(txtAccepted(index)) >= Val(lblPqt(index)) Then chkComplete(index).Value = vbChecked
   vItems(index + iIndex, POITEM_AcceptedQty) = Val(txtAccepted(index))
   vItems(index + iIndex, POITEM_Complete) = chkComplete(index).Value
   CheckBalance index
End Sub


Private Sub txtCmt_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
End Sub


Private Sub txtCmt_KeyPress(index As Integer, KeyAscii As Integer)
   KeyMemo KeyAscii
   
End Sub


Private Sub txtCmt_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub

Private Sub txtCmt_Validate(index As Integer, Cancel As Boolean)
   txtCmt(index) = CheckLen(txtCmt(index), 255)
   vItems(iIndex + index, 9) = txtCmt(index)
   
End Sub

Private Sub txtDte_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtDte_LostFocus()
   txtDte = CheckDate(txtDte)
   
End Sub

Private Sub txtRejected_GotFocus(index As Integer)
   SelectFormat Me
End Sub

Private Sub txtRejected_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
End Sub

Private Sub txtRejected_KeyPress(index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
End Sub

Private Sub txtRejected_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
End Sub

Private Sub txtRejected_Validate(index As Integer, Cancel As Boolean)
   txtRejected(index) = CheckLen(txtRejected(index), 9)
   txtRejected(index) = Format(Abs(Val(txtRejected(index))), ES_QuantityDataFormat)
   vItems(index + iIndex, POITEM_RejectedQty) = Val(txtRejected(index))
   CheckBalance index
End Sub

Private Sub txtWaste_GotFocus(index As Integer)
   SelectFormat Me
End Sub

Private Sub txtWaste_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
End Sub

Private Sub txtWaste_KeyPress(index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
End Sub

Private Sub txtWaste_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
End Sub

Private Sub txtWaste_Validate(index As Integer, Cancel As Boolean)
   txtWaste(index) = CheckLen(txtWaste(index), 9)
   txtWaste(index) = Format(Abs(Val(txtWaste(index))), ES_QuantityDataFormat)
   vItems(index + iIndex, POITEM_WastedQty) = Val(txtWaste(index))
   CheckBalance index
End Sub



Private Sub UpdateList()
   Dim bResponse As Byte
   Dim iRow As Integer
   Dim cTotalQty As Currency, cOrderedQty As Currency, cRejectedQty As Currency
   Dim sMsg As String
   Dim bOverQty As Boolean
   
   On Error GoTo DiaErr1
   
   bOverQty = False
   bOverQty = GetPOOverQtyFlag
   
   Dim completedCount As Integer
   Dim zeroCount As Integer
   Dim partialCount As Integer
   For iRow = 1 To iTotalItems
      cTotalQty = Val(vItems(iRow, POITEM_AcceptedQty)) _
         + Val(vItems(iRow, POITEM_RejectedQty)) + Val(vItems(iRow, POITEM_WastedQty))
      cOrderedQty = Val(vItems(iRow, POITEM_OrderedQty))
      cRejectedQty = Val(vItems(iRow, POITEM_RejectedQty))
      vItems(iRow, POITEM_SplitForReject) = 0
      
      'if no quantities, leave item as is
      If cTotalQty = 0 Then
         vItems(iRow, POITEM_Complete) = 0
         zeroCount = zeroCount + 1
      ElseIf cTotalQty = cOrderedQty Then
         vItems(iRow, POITEM_Complete) = 1
         completedCount = completedCount + 1
      ElseIf (cTotalQty <> cOrderedQty) Or ((cTotalQty > cOrderedQty)) Then
         If (bOverQty) Then
            vItems(iRow, POITEM_Complete) = 1
            completedCount = completedCount + 1
         Else
            vItems(iRow, POITEM_Complete) = 0
            partialCount = partialCount + 1
         End If
          Else
         vItems(iRow, POITEM_Complete) = 0
         partialCount = partialCount + 1
      End If
      

      'if there are rejections, ask whether to create a split item
      If cRejectedQty > 0 Then
         Select Case MsgBox("Split PO item " & vItems(iRow, 2) & vItems(iRow, 3) & " to replenish the rejected quantity " _
            & cRejectedQty & " of part " & vItems(iRow, 4) & "?", vbYesNo + vbQuestion)
            Case vbYes:
               vItems(iRow, POITEM_SplitForReject) = 1
         End Select
      End If
   Next
   

   If completedCount = 0 Then
      MsgBox "No items selected.  Nothing to do."
      Exit Sub
   End If
   
   If partialCount > 0 Then
      MsgBox partialCount & " items with unbalanced quantities.  Cannot proceed."
      Exit Sub
   End If
   
   'get confirmation before proceeding
   sMsg = "You have chosen to update inspection data for " & completedCount & " items." & vbCrLf _
      & "Do you wish to proceed?"
   If MsgBox(sMsg, ES_YESQUESTION, Caption) <> vbYes Then
      Exit Sub
   End If

   'Passes the tests -- update the database
   MouseCursor 13
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   
   Dim poItem As New ClassPoItem
   Dim cAcceptedQty As Currency, cWastedQty As Currency, cItemQty As Currency, cSplitQty As Currency
   Dim cTotalDelQty As Currency
   Dim splitForRejected As Boolean, itemComplete As Boolean
   Dim poNumber As Long, poItemNo As Integer, poRev As String
   
   poNumber = Val(cmbPon)
   
   Dim bContinue As Boolean
   '  On Error Resume Next
   For iRow = 1 To iTotalItems
      cAcceptedQty = Val(vItems(iRow, POITEM_AcceptedQty))
      cRejectedQty = Val(vItems(iRow, POITEM_RejectedQty))
      cWastedQty = Val(vItems(iRow, POITEM_WastedQty))
      cTotalQty = cAcceptedQty + cRejectedQty + cWastedQty
      cOrderedQty = Val(vItems(iRow, POITEM_OrderedQty))
      splitForRejected = IIf(vItems(iRow, POITEM_SplitForReject) = 0, False, True)
      itemComplete = IIf(vItems(iRow, POITEM_Complete) = 0, False, True)
      poItemNo = Val(vItems(iRow, 2))
      poRev = vItems(iRow, 3)
      
      'If cTotalQty > 0 Then
      If itemComplete Then
         
         'if item is split, update PO item quantity
'         If cTotalQty < cOrderedQty And Not itemComplete Then
'            cItemQty = cTotalQty
'            cSplitQty = cOrderedQty - cTotalQty
'         Else
            cItemQty = cOrderedQty
            cSplitQty = 0
'         End If
         
         If vItems(iRow, 12) = "Null" Then vItems(iRow, 12) = "'" & txtDte & "'"

         
         bContinue = True
         If (Val(vItems(iRow, POITEM_DEL)) = 1) Then
            cTotalDelQty = Val(vItems(iRow, POITEM_DELQty))
            If (cTotalDelQty <> cTotalQty) Then
               
               sMsg = "There is a difference in the PO Delevered quantity " _
                      & "to the Inspected quantity. " & vbCrLf _
                      & "Would you like to continue anyway?"
               If MsgBox(sMsg, ES_NOQUESTION, Caption) = vbNo Then
                  bContinue = False
               End If
               
            End If
         End If
         
         If (bContinue) Then
            
            sSql = "UPDATE PoitTable SET PIONDOCKINSPECTED=1," & vbCrLf _
               & "PIODDELIVERED=1," & vbCrLf _
               & "PIODDELDATE=" & vItems(iRow, 12) & "," & vbCrLf _
               & "PIODDELQTY=" & cTotalQty & "," & vbCrLf _
               & "PIONDOCKINSPDATE='" & txtDte & "'," & vbCrLf _
               & "PIONDOCKQTYACC=" & cAcceptedQty & "," & vbCrLf _
               & "PIONDOCKQTYREJ=" & cRejectedQty & "," & vbCrLf _
               & "PIONDOCKQTYWASTE=" & cWastedQty & "," & vbCrLf _
               & "PIONDOCKINSPECTOR='" & vItems(iRow, 8) & "'," & vbCrLf _
               & "PIONDOCKCOMMENT='" & vItems(iRow, 9) & "'," & vbCrLf _
               & "PIPQTY=" & cItemQty
               
            'if no quantity left, set status = received
            If cAcceptedQty = 0 Then
               ' if Purchase Qty is same Rejected Qty - mark as IATYPE_PoRejectItem
               If (cItemQty = cRejectedQty) Then
                   sSql = sSql & "," & vbCrLf & "PITYPE = " & IATYPE_PoRejectItem
               Else
                   sSql = sSql & "," & vbCrLf & "PITYPE = " & IATYPE_PoReceipt
               End If
            End If
            
            sSql = sSql & vbCrLf & "WHERE PINUMBER=" & poNumber & vbCrLf _
               & "AND PIITEM=" & poItemNo & vbCrLf _
               & "AND PIREV='" & poRev & "'"
            clsADOCon.ExecuteSql sSql
            
            'if a split for rejected quantity, create that
            If splitForRejected Then
               poItem.CreateSplit poNumber, 0, poItemNo, poRev, cRejectedQty, False
               
'               ' point operation records tied to this po item to the new po item
'               sSql = "update RnopTable" & vbCrLf _
'                  & "set rop.OPPOREV = " & vbCrLf _
'                  & "isnull((select max(PIREV) from PoitTable where PINUMBER = OPPONUMBER and PIRELEASE = 0 and PIITEM = OPPOITEM),'')" & vbCrLf _
'                  & "where OPPONUMBER = 49106 and OPPOITEM = 1 and OPPOREV = ''"
'               Dim success As Boolean
'               success = clsADOCon.ExecuteSql(sSql)
            End If
         End If
         
      End If
   Next
   MouseCursor 0
   clsADOCon.CommitTrans
   MsgBox "Inspection information has been updated.", vbInformation, Caption
   ManageBoxes
   cmdUpd.Enabled = False
   cmdEnd.Enabled = False
   cmbPon.Enabled = True
   txtDte.Enabled = True
   cmdItm.Enabled = True
'  On Error Resume Next
   cmbPon.SetFocus
   Exit Sub
   
DiaErr1:
   clsADOCon.RollbackTrans
   sProcName = "updatelist"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub





Private Sub ManageBoxes()
   Dim iList As Integer
   For iList = 1 To 4
'      lblItm(iList).Visible = False
'      lblRev(iList).Visible = False
      lblPrt(iList).ToolTipText = "Part Description"
'      lblPrt(iList).Visible = False
'      lblPqt(iList).Visible = False
'      txtAccepted(iList).Visible = False
'      txtRejected(iList).Visible = False
'      txtWaste(iList).Visible = False
'      cmbIns(iList).Visible = False
'      cmdCmt(iList).Visible = False
'      txtCmt(iList).Visible = False
'      lblIns(iList).Visible = False
'      lblCmt(iList).Visible = False
'      chkComplete(iList).Visible = False
'      lblUom(iList).Visible = False
      showItem iList, False
      
   Next
   
End Sub

Private Sub CheckBalance(index As Integer)
   Dim totalQty As Currency
   If IsNumeric("0" & txtAccepted(index)) And IsNumeric("0" & txtRejected(index)) And IsNumeric("0" & txtWaste(index)) Then
      totalQty = CCur("0" & txtAccepted(index)) + CCur("0" & txtRejected(index)) + CCur("0" & txtWaste(index))
      If totalQty > 0 Then
         If CCur(Me.lblPqt(index)) = totalQty Then
            chkComplete(index) = vbChecked
            SetLineColor index, ES_BLACK
         Else
            chkComplete(index) = vbUnchecked
            SetLineColor index, ES_RED
         End If
      Else
         chkComplete(index) = vbUnchecked
         SetLineColor index, ES_BLACK
      End If
   Else
      chkComplete(index) = vbUnchecked
      SetLineColor index, ES_RED
   End If
End Sub

Private Sub SetLineColor(index As Integer, color As Long)
   txtAccepted(index).ForeColor = color
   txtRejected(index).ForeColor = color
   txtWaste(index).ForeColor = color
End Sub

Private Function GetPOOverQtyFlag()
   
   Dim RdoPOqty As ADODB.Recordset
   sSql = "SELECT ISNULL(COALLOWDELOVERQTY, 0) COALLOWDELOVERQTY FROM ComnTable WHERE COREF=1"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPOqty, ES_FORWARD)
   If bSqlRows Then
      With RdoPOqty
         GetPOOverQtyFlag = IIf(!COALLOWDELOVERQTY = 0, False, True)
         ClearResultSet RdoPOqty
      End With
   Else
      GetPOOverQtyFlag = False
   End If
   Set RdoPOqty = Nothing

End Function
