VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaAPe08a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Computer Check Setup"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkInvoiceNumber 
      Caption         =   "Order vendor by invoice #"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6420
      TabIndex        =   130
      ToolTipText     =   "Leave unchecked to order vendor by invoice date"
      Top             =   1080
      Width           =   2775
   End
   Begin VB.CommandButton cmdLastPage 
      Height          =   495
      Left            =   10200
      Picture         =   "diaAPe08a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   129
      ToolTipText     =   "Last Page"
      Top             =   5520
      Width           =   495
   End
   Begin VB.CommandButton cmdNextPage 
      Height          =   495
      Left            =   9720
      Picture         =   "diaAPe08a.frx":0578
      Style           =   1  'Graphical
      TabIndex        =   128
      ToolTipText     =   "Next Page"
      Top             =   5520
      Width           =   495
   End
   Begin VB.CommandButton cmdPrevPage 
      Height          =   495
      Left            =   9240
      Picture         =   "diaAPe08a.frx":0AEB
      Style           =   1  'Graphical
      TabIndex        =   127
      ToolTipText     =   "Previous Page"
      Top             =   5520
      Width           =   495
   End
   Begin VB.CommandButton cmdFirstPage 
      Height          =   495
      Left            =   8760
      Picture         =   "diaAPe08a.frx":105D
      Style           =   1  'Graphical
      TabIndex        =   126
      ToolTipText     =   "First Page"
      Top             =   5520
      Width           =   495
   End
   Begin VB.ComboBox cboCheckDate 
      Height          =   315
      Left            =   4200
      TabIndex        =   1
      Tag             =   "4"
      ToolTipText     =   "Show discounts available through this date"
      Top             =   360
      Width           =   1095
   End
   Begin VB.CheckBox chkFullName 
      Caption         =   "Order by full vendor name"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3720
      TabIndex        =   114
      ToolTipText     =   "Leave unchecked to order by vendor nickname"
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox txtDis 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   9
      Left            =   9000
      TabIndex        =   40
      Tag             =   "1"
      Top             =   5160
      Width           =   735
   End
   Begin VB.TextBox txtDis 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   8
      Left            =   9000
      TabIndex        =   36
      Tag             =   "1"
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox txtDis 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   7
      Left            =   9000
      TabIndex        =   32
      Tag             =   "1"
      Top             =   4440
      Width           =   735
   End
   Begin VB.TextBox txtDis 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   6
      Left            =   9000
      TabIndex        =   28
      Tag             =   "1"
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox txtDis 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   5
      Left            =   9000
      TabIndex        =   24
      Tag             =   "1"
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox txtDis 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   9000
      TabIndex        =   20
      Tag             =   "1"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox txtDis 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   9000
      TabIndex        =   16
      Tag             =   "1"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox txtDis 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   9000
      TabIndex        =   12
      Tag             =   "1"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox txtDis 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   9000
      TabIndex        =   8
      Tag             =   "1"
      Top             =   2280
      Width           =   735
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CheckBox optPar 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   5
      Left            =   8640
      TabIndex        =   23
      ToolTipText     =   "Short Pay"
      Top             =   3720
      Width           =   255
   End
   Begin VB.CheckBox optPar 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   4
      Left            =   8640
      TabIndex        =   19
      ToolTipText     =   "Short Pay"
      Top             =   3360
      Width           =   255
   End
   Begin VB.CheckBox optPar 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   3
      Left            =   8640
      TabIndex        =   15
      ToolTipText     =   "Short Pay"
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox optPar 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   2
      Left            =   8640
      TabIndex        =   11
      ToolTipText     =   "Short Pay"
      Top             =   2640
      Width           =   255
   End
   Begin VB.CheckBox optPar 
      Height          =   255
      Index           =   1
      Left            =   8640
      TabIndex        =   7
      ToolTipText     =   "Short Pay Invoice But Mark As PIF"
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox optPar 
      ForeColor       =   &H8000000F&
      Height          =   315
      Index           =   7
      Left            =   8640
      TabIndex        =   31
      ToolTipText     =   "Short Pay"
      Top             =   4440
      Width           =   255
   End
   Begin VB.CheckBox optPar 
      ForeColor       =   &H8000000F&
      Height          =   315
      Index           =   8
      Left            =   8640
      TabIndex        =   35
      ToolTipText     =   "Short Pay"
      Top             =   4800
      Width           =   255
   End
   Begin VB.CheckBox optPar 
      ForeColor       =   &H8000000F&
      Height          =   375
      Index           =   9
      Left            =   8640
      TabIndex        =   39
      ToolTipText     =   "Short Pay"
      Top             =   5160
      Width           =   255
   End
   Begin VB.CheckBox optPar 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   6
      Left            =   8640
      TabIndex        =   27
      ToolTipText     =   "Short Pay"
      Top             =   4080
      Width           =   255
   End
   Begin VB.CheckBox optPad 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   25
      ToolTipText     =   "Apply Payment"
      Top             =   4080
      Width           =   255
   End
   Begin VB.TextBox txtPad 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   7440
      TabIndex        =   30
      Tag             =   "1"
      ToolTipText     =   "Amount To Be Paid"
      Top             =   4440
      Width           =   1100
   End
   Begin VB.CheckBox optPad 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   37
      ToolTipText     =   "Apply Payment"
      Top             =   5160
      Width           =   255
   End
   Begin VB.TextBox txtPad 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   7440
      TabIndex        =   38
      Tag             =   "1"
      ToolTipText     =   "Amount To Be Paid"
      Top             =   5160
      Width           =   1100
   End
   Begin VB.TextBox txtPad 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   7440
      TabIndex        =   34
      Tag             =   "1"
      ToolTipText     =   "Amount To Be Paid"
      Top             =   4800
      Width           =   1100
   End
   Begin VB.CheckBox optPad 
      ForeColor       =   &H8000000F&
      Height          =   315
      Index           =   8
      Left            =   120
      TabIndex        =   33
      ToolTipText     =   "Apply Payment"
      Top             =   4800
      Width           =   255
   End
   Begin VB.TextBox txtPad 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   7440
      TabIndex        =   26
      Tag             =   "1"
      ToolTipText     =   "Amount To Be Paid"
      Top             =   4080
      Width           =   1100
   End
   Begin VB.CheckBox optPad 
      ForeColor       =   &H8000000F&
      Height          =   315
      Index           =   7
      Left            =   120
      TabIndex        =   29
      ToolTipText     =   "Apply Payment"
      Top             =   4440
      Width           =   255
   End
   Begin VB.CheckBox optPad 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Apply Payment"
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox optPad 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Apply Payment"
      Top             =   2640
      Width           =   255
   End
   Begin VB.CheckBox optPad 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "Apply Payment"
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox optPad 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   17
      ToolTipText     =   "Apply Payment"
      Top             =   3360
      Width           =   255
   End
   Begin VB.CheckBox optPad 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   21
      ToolTipText     =   "Apply Payment"
      Top             =   3720
      Width           =   255
   End
   Begin VB.TextBox txtDummy 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   285
      Left            =   6840
      TabIndex        =   80
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "&Select"
      Height          =   315
      Left            =   9840
      TabIndex        =   4
      ToolTipText     =   "Select Invoices For This Vendor"
      Top             =   720
      Width           =   875
   End
   Begin VB.CommandButton cmdCnl 
      Caption         =   " &Reselect"
      Height          =   315
      Left            =   9840
      TabIndex        =   42
      ToolTipText     =   "Cancel Operation"
      Top             =   1560
      Width           =   875
   End
   Begin VB.CommandButton cmdPst 
      Caption         =   "&Post"
      Height          =   315
      Left            =   9840
      TabIndex        =   41
      ToolTipText     =   "Post This Check"
      Top             =   1080
      Width           =   875
   End
   Begin Threed.SSFrame z2 
      Height          =   30
      Left            =   240
      TabIndex        =   75
      Top             =   1440
      Width           =   10455
      _Version        =   65536
      _ExtentX        =   18441
      _ExtentY        =   53
      _StockProps     =   14
      Caption         =   "SSFrame1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.26
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtPad 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   7440
      TabIndex        =   22
      Tag             =   "1"
      ToolTipText     =   "Amount To Be Paid"
      Top             =   3720
      Width           =   1100
   End
   Begin VB.TextBox txtPad 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   7440
      TabIndex        =   18
      Tag             =   "1"
      ToolTipText     =   "Amount To Be Paid"
      Top             =   3360
      Width           =   1100
   End
   Begin VB.TextBox txtPad 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   7440
      TabIndex        =   14
      Tag             =   "1"
      ToolTipText     =   "Amount To Be Paid"
      Top             =   3000
      Width           =   1100
   End
   Begin VB.TextBox txtPad 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   7440
      TabIndex        =   10
      Tag             =   "1"
      ToolTipText     =   "Amount To Be Paid"
      Top             =   2640
      Width           =   1100
   End
   Begin VB.TextBox txtPad 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   7440
      TabIndex        =   6
      Tag             =   "1"
      ToolTipText     =   "Amount To Be Paid"
      Top             =   2280
      Width           =   1100
   End
   Begin VB.ComboBox cboDueDate 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Tag             =   "4"
      ToolTipText     =   "Show invoices due by this date"
      Top             =   360
      Width           =   1095
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   6480
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Contains Vendors With PO's Unpaid Invoices"
      Top             =   360
      Width           =   1555
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   9840
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   2820
      Top             =   780
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6105
      FormDesignWidth =   10845
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   83
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaAPe08a.frx":15DD
      PictureDn       =   "diaAPe08a.frx":1723
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Date"
      Height          =   285
      Index           =   18
      Left            =   3180
      TabIndex        =   125
      Top             =   420
      Width           =   1005
   End
   Begin VB.Label lblDiscDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   9840
      TabIndex        =   124
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label lblDiscDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   9840
      TabIndex        =   123
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label lblDiscDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   9840
      TabIndex        =   122
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label lblDiscDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   9840
      TabIndex        =   121
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label lblDiscDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   9840
      TabIndex        =   120
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label lblDiscDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   9840
      TabIndex        =   119
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label lblDiscDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   9840
      TabIndex        =   118
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label lblDiscDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   9840
      TabIndex        =   117
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label lblDiscDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   9840
      TabIndex        =   116
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Disc Date   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   9840
      TabIndex        =   115
      Top             =   2040
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pay All on This Page"
      Height          =   285
      Index           =   16
      Left            =   240
      TabIndex        =   113
      Top             =   1560
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor               "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   480
      TabIndex        =   112
      Top             =   2040
      Width           =   1185
   End
   Begin VB.Label lblVnd 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   480
      TabIndex        =   111
      Top             =   5160
      Width           =   1150
   End
   Begin VB.Label lblVnd 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   480
      TabIndex        =   110
      Top             =   4800
      Width           =   1150
   End
   Begin VB.Label lblVnd 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   480
      TabIndex        =   109
      Top             =   4440
      Width           =   1150
   End
   Begin VB.Label lblVnd 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   480
      TabIndex        =   108
      Top             =   4080
      Width           =   1150
   End
   Begin VB.Label lblVnd 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   480
      TabIndex        =   107
      Top             =   3720
      Width           =   1150
   End
   Begin VB.Label lblVnd 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   480
      TabIndex        =   106
      Top             =   3360
      Width           =   1150
   End
   Begin VB.Label lblVnd 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   480
      TabIndex        =   105
      Top             =   3000
      Width           =   1150
   End
   Begin VB.Label lblVnd 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   480
      TabIndex        =   104
      Top             =   2640
      Width           =   1150
   End
   Begin VB.Label lblVnd 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   480
      TabIndex        =   103
      Top             =   2280
      Width           =   1150
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   1
      Left            =   8160
      TabIndex        =   102
      Top             =   360
      Width           =   1185
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   1680
      TabIndex        =   101
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   2640
      TabIndex        =   100
      Top             =   4440
      Width           =   2400
   End
   Begin VB.Label lblPon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   5160
      TabIndex        =   99
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label lblDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   6240
      TabIndex        =   98
      ToolTipText     =   "Includes Tax And Freight"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   1680
      TabIndex        =   97
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   2640
      TabIndex        =   96
      Top             =   5160
      Width           =   2400
   End
   Begin VB.Label lblPon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   5160
      TabIndex        =   95
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label lblDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   6240
      TabIndex        =   94
      ToolTipText     =   "Includes Tax And Freight"
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label lblDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   6240
      TabIndex        =   93
      ToolTipText     =   "Includes Tax And Freight"
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label lblPon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   5160
      TabIndex        =   92
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   2640
      TabIndex        =   91
      Top             =   4800
      Width           =   2400
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   1680
      TabIndex        =   90
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label lblDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   6240
      TabIndex        =   89
      ToolTipText     =   "Includes Tax And Freight"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label lblPon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   5160
      TabIndex        =   88
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   2640
      TabIndex        =   87
      Top             =   4080
      Width           =   2400
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   1680
      TabIndex        =   86
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pay"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   120
      TabIndex        =   85
      ToolTipText     =   "Apply Payment"
      Top             =   2040
      Width           =   435
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "SP "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   17
      Left            =   8640
      TabIndex        =   84
      ToolTipText     =   "Short Pay Invoice But Mark As PIF"
      Top             =   2040
      Width           =   315
   End
   Begin VB.Label lblTot 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   82
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Selected"
      Height          =   285
      Index           =   15
      Left            =   120
      TabIndex        =   81
      Top             =   5640
      Width           =   1305
   End
   Begin VB.Label lblPge 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6960
      TabIndex        =   79
      Top             =   5760
      Width           =   375
   End
   Begin VB.Label lblLst 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7680
      TabIndex        =   78
      Top             =   5760
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Page"
      Height          =   255
      Index           =   14
      Left            =   6480
      TabIndex        =   77
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "of"
      Height          =   255
      Index           =   13
      Left            =   7440
      TabIndex        =   76
      Top             =   5760
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Discount    "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   9000
      TabIndex        =   74
      Top             =   2040
      Width           =   705
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Paid          "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   7440
      TabIndex        =   73
      Top             =   2040
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Due          "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   6240
      TabIndex        =   72
      Top             =   2040
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Number             "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   5160
      TabIndex        =   71
      Top             =   2040
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Number                            "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   2640
      TabIndex        =   70
      Top             =   2040
      Width           =   2475
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inv Date                 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   1680
      TabIndex        =   69
      Top             =   2040
      Width           =   825
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   1680
      TabIndex        =   68
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   2640
      TabIndex        =   67
      Top             =   3720
      Width           =   2400
   End
   Begin VB.Label lblPon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   5160
      TabIndex        =   66
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label lblDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   6240
      TabIndex        =   65
      ToolTipText     =   "Includes Tax And Freight"
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   1680
      TabIndex        =   64
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   2640
      TabIndex        =   63
      Top             =   3360
      Width           =   2400
   End
   Begin VB.Label lblPon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   5160
      TabIndex        =   62
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   6240
      TabIndex        =   61
      ToolTipText     =   "Includes Tax And Freight"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   1680
      TabIndex        =   60
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   2640
      TabIndex        =   59
      Top             =   3000
      Width           =   2400
   End
   Begin VB.Label lblPon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   5160
      TabIndex        =   58
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   6240
      TabIndex        =   57
      ToolTipText     =   "Includes Tax And Freight"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   1680
      TabIndex        =   56
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   2640
      TabIndex        =   55
      Top             =   2640
      Width           =   2400
   End
   Begin VB.Label lblPon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   5160
      TabIndex        =   54
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   6240
      TabIndex        =   53
      ToolTipText     =   "Includes Tax And Freight"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   52
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   2640
      TabIndex        =   51
      Top             =   2280
      Width           =   2400
   End
   Begin VB.Label lblPon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   5160
      TabIndex        =   50
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   6240
      TabIndex        =   49
      ToolTipText     =   "Includes Tax And Freight"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblCnt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   285
      Left            =   1560
      TabIndex        =   48
      Top             =   720
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoices Found"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   47
      Top             =   720
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   46
      Top             =   360
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   285
      Index           =   0
      Left            =   5760
      TabIndex        =   45
      Top             =   420
      Width           =   705
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3720
      TabIndex        =   44
      Top             =   720
      Width           =   2775
   End
End
Attribute VB_Name = "diaAPe08a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'*****************************************************************************************
' diaAPe08a - Computer Check Setup
'
' Created: (nth)
' Revisions:
' 12/26/01 (nth) Added a ledger like look and feel.
' 12/27/01 (nth) Reworked GetVendorInvoices
' 12/27/01 (nth) Display AP discounts.
' 09/09/02 (nth) Added LoadChkSetup.
' 12/10/02 (nth) Fixed error in PostCheck (loading all invoices into check setup).
' 01/24/03 (nth) Load a prior check setup for revision.
' 05/12/04 (nth) Trap for negative check amount.
' 07/27/04 (nth) Add compress when building computer check setup.
' 12/08/04 (nth) Select invoices not yet due but with discounts about to expire.
' 12/13/04 (nth) To allow discount override.
' 12/13/04 (nth) Added tooltips vendor full name and invoice due date.
'
'*****************************************************************************************

Dim bOnLoad As Byte
Dim bGoodVendor As Byte
Dim bCancel As Byte
Dim bLoadSetup As Byte
Dim bPage As Byte

Dim iIndex As Integer
Dim iCurrPage As Integer
Dim iTotalItems As Integer

'Dim lCheck      As Long

Dim sApAcct As String
Dim sApDiscAcct As String
Dim sXcAcct As String
Dim sCcAcct As String
Dim sCgAcct As String
Dim sJournalID As String
Dim sMsg As String

' vInvoice Documentation
' 0  = Invoice Date
Private Const VI_DATE = 0
' 1  = Invoice Number
Private Const VI_INVNO = 1
' 2  = PO
Private Const VI_PO = 2
' 3  = Invoice Total
Private Const VI_INVOICETOTAL = 3
' 4  = Amount Paid
Private Const VI_AMOUNTPAID = 4
' 5  = Calculated Discount
Private Const VI_CalcDiscount = 5
' 6  = Freight
' 7  = Tax
' 8  = Discount Percentage -- NOT USED
' 9  = Marked Paid ? (vbchecked or vbunchecked)
Private Const VI_PayThisInvoice = 9
' 10 = Partial (0 or 1)
Private Const VI_SHORTPAY = 10
' 11 = Item Number
' 12 = Vendor Nickname
' 13 = Due Date (ToolTip)
' 14 = Vendor Name
' 15 = Discount Amount
Private Const VI_DiscountTaken = 15
' 16 = due date info
Private Const VI_DUEDATEINFO = 16
Private Const VI_DISCOUNTCUTOFFDATE = 17

Dim vInvoice(10000, 17) As Variant

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cboCheckDate_DropDown()
   ShowCalendar Me
End Sub

Private Sub chkAll_Click()
   
   Dim Index As Integer
   
   If Not bPage Then
      iIndex = (Val(lblPge) - 1) * 9
      For Index = 1 To 9
         If Index + iIndex <= iTotalItems Then
            If optPad(Index) <> vbChecked Then
               vInvoice(Index + iIndex, VI_AMOUNTPAID) = (vInvoice(Index + iIndex, VI_INVOICETOTAL) _
                        - vInvoice(Index + iIndex, VI_CalcDiscount))
               txtPad(Index) = Format(vInvoice(Index + iIndex, VI_AMOUNTPAID), CURRENCYMASK)
               txtPad(Index).enabled = True
               If vInvoice(Index + iIndex, VI_INVOICETOTAL) > 0 _
                           And vInvoice(Index + iIndex, VI_CalcDiscount) > 0 Then
                  txtDis(Index).enabled = True
               End If
               optPad(Index) = vbChecked

               vInvoice(Index + iIndex, VI_PayThisInvoice) = optPad(Index)
               UpdateTotals
            End If
         End If
      Next
      chkAll.Value = False
   End If
End Sub

Private Sub cboCheckDate_LostFocus()
   If Not bCancel Then
      cboCheckDate = CheckDate(cboCheckDate)
      GetInvoices
   End If
End Sub

Private Sub chkFullName_Click()
   SaveSetting "Esi2000", "ComputerChecks", "OrderByFullName", chkFullName.Value
End Sub

'*****************************************************************************************

Private Sub cmbVnd_Click()
   bGoodVendor = FindVendor(Me)
End Sub

Private Sub cmbVnd_GotFocus()
   SelectFormat Me
End Sub

Private Sub cmbVnd_LostFocus()
   If Not bCancel Then
      cmbVnd = CheckLen(cmbVnd, 10)
      If cmbVnd = "" Or UCase(cmbVnd) = "ALL" Then
         bGoodVendor = True
         cmbVnd = "ALL"
         lblNme = "*** Multiple Vendors Selected ***"
      Else
         bGoodVendor = FindVendor(Me)
      End If
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub cmdCnl_Click()
   Dim bByte As Boolean
   Dim i As Integer
   Dim bResponse As Byte
   
   For i = 1 To iTotalItems
      If vInvoice(i, VI_PayThisInvoice) = vbChecked Then
         bByte = True
         Exit For
      End If
   Next
   If bByte Then
      sMsg = "You Have Items Selected. Are You" _
             & vbCrLf & "Sure That You Want To Cancel?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbYes Then
         bByte = False
      Else
         bByte = True
      End If
   End If
   If Not bByte Then
      cboDueDate.enabled = True
      cboCheckDate.enabled = True
      cmbVnd.enabled = True
      CloseBoxes
      txtDummy.enabled = False
      On Error Resume Next
      cboDueDate.SetFocus
      If bGoodVendor Then
         cmdSel.enabled = True
      End If
   End If
End Sub




Private Sub cmdFirstPage_Click()
    iCurrPage = 1
    lblPge = iCurrPage
    GetNextGroup
End Sub

Private Sub cmdLastPage_Click()
    iCurrPage = Val(lblLst)
    lblPge = iCurrPage
    GetNextGroup
End Sub

Private Sub cmdNextPage_Click()
    If iCurrPage >= Val(lblLst) Then Exit Sub
    iCurrPage = iCurrPage + 1
    lblPge = iCurrPage
    GetNextGroup
End Sub

Private Sub cmdPrevPage_Click()
    If iCurrPage = 1 Then Exit Sub
    iCurrPage = iCurrPage - 1
    lblPge = iCurrPage
    GetNextGroup
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, Caption
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub cmdPst_Click()
   Dim b As Byte
   Dim i As Integer
   For i = 1 To iTotalItems
      If vInvoice(i, VI_PayThisInvoice) = vbChecked Then
         b = 1
         Exit For
      End If
   Next
   If b = 0 Then
      sMsg = "You Have Nothing Selected To Pay."
      MsgBox sMsg, vbInformation, Caption
   Else
      PostChecks
      MouseCursor 0
   End If
End Sub

Private Sub cmdSel_Click()
   bCancel = False
   If Trim(cmbVnd) = "" Then
      cmbVnd = "ALL"
      lblNme = "*** Multiple Vendors Selected ***"
   End If
   chkFullName.enabled = False
   'bLoadSetup = True
   FillGrid
End Sub

Private Sub cmdSel_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
End Sub


Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      CurrentJournal "CC", ES_SYSDATE, sJournalID
      LoadCheckSetup
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   CloseBoxes
   cboDueDate = Format(ES_SYSDATE, DATEMASK)
   cboCheckDate = cboDueDate
   txtDummy.BackColor = Me.BackColor
   cmdSel.enabled = False
   bOnLoad = True
   chkFullName.Value = GetSetting("Esi2000", "ComputerChecks", "OrderByFullName", 0)
   
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaAPe08a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub optPad_Click(Index As Integer)
   If Not bPage Then
      If bLoadSetup Then
         txtPad(Index).enabled = True
      Else
         If optPad(Index) = vbChecked Then
            vInvoice(Index + iIndex, VI_AMOUNTPAID) = (vInvoice(Index + iIndex, VI_INVOICETOTAL) _
                     - vInvoice(Index + iIndex, VI_CalcDiscount))
            txtPad(Index) = Format(vInvoice(Index + iIndex, VI_AMOUNTPAID), CURRENCYMASK)
            txtPad(Index).enabled = True
            If vInvoice(Index + iIndex, VI_INVOICETOTAL) > 0 _
                        And vInvoice(Index + iIndex, VI_CalcDiscount) > 0 Then
               txtDis(Index).enabled = True
            End If
         Else
            vInvoice(Index + iIndex, VI_AMOUNTPAID) = 0
            txtPad(Index) = "0.00"
            txtPad(Index).enabled = False
            
            vInvoice(Index + iIndex, VI_SHORTPAY) = vbUnchecked
            optPar(Index) = vbUnchecked
            
            txtDis(Index).enabled = False
            vInvoice(Index + iIndex, VI_DiscountTaken) = vInvoice(Index + iIndex, VI_CalcDiscount)
            'txtDis(Index) = Format(vInvoice(Index + iIndex, VI_DiscountTaken), CURRENCYMASK)
            txtDis(Index) = Format(vInvoice(Index + iIndex, VI_CalcDiscount), CURRENCYMASK)
         End If
         vInvoice(Index + iIndex, VI_PayThisInvoice) = optPad(Index)
         UpdateTotals
      End If
   End If
End Sub

Private Sub optPad_KeyPress(Index As Integer, _
                            KeyAscii As Integer)
   KeyLock KeyAscii
End Sub

Private Sub optPad_KeyUp(Index As Integer, _
                         KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdPrevPage_Click
   If KeyCode = vbKeyPageDown Then cmdNextPage_Click
End Sub

Private Sub optPad_LostFocus(Index As Integer)
   vInvoice(iIndex + Index, VI_PayThisInvoice) = optPad(Index).Value
End Sub

Private Sub optPar_Click(Index As Integer)
   If Not bPage Then
      If optPar(Index).Value = vbChecked Then
         vInvoice(Index + iIndex, VI_DiscountTaken) = 0
         vInvoice(Index + iIndex, VI_SHORTPAY) = 1
         txtDis(Index) = "0.00"
         optPad(Index) = vbChecked
         txtDis(Index).enabled = False
      Else
         vInvoice(Index + iIndex, VI_SHORTPAY) = 0
         txtDis(Index).enabled = True
      End If
      UpdateTotals
   End If
End Sub

Private Sub optPar_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyLock KeyAscii
End Sub

Private Sub optPar_KeyUp(Index As Integer, _
                         KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdPrevPage_Click
   If KeyCode = vbKeyPageDown Then cmdNextPage_Click
End Sub

Private Sub txtDis_GotFocus(Index As Integer)
   SelectFormat Me
End Sub

Private Sub txtDis_KeyDown(Index As Integer, _
                           KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
End Sub

Private Sub txtDis_KeyPress(Index As Integer, _
                            KeyAscii As Integer)
   KeyValue KeyAscii
End Sub

Private Sub txtDis_KeyUp(Index As Integer, _
                         KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then
      cmdPrevPage_Click
   ElseIf KeyCode = vbKeyPageDown Then
      cmdNextPage_Click
   End If
End Sub

Private Sub txtDis_LostFocus(Index As Integer)
   Dim cDis As Currency
   Dim cPay As Currency
   Dim cDue As Currency
   Dim cFin As Currency
   
   txtDis(Index) = CheckLen(txtDis(Index), 10)
   If ("" & Trim(txtDis(Index))) = "" Then
      txtDis(Index) = "0.00"
   End If
   
   cDis = CCur(txtDis(Index))
   cPay = CCur(txtPad(Index))
   cDue = CCur(lblDue(Index))
   
   If cDis <> (cDue - cPay) Or cDis < 0 Then
      Beep
      'cFin = vInvoice(Index + iIndex, VI_DiscountTaken)
      cPay = cDue - cDis
      vInvoice(Index + iIndex, VI_AMOUNTPAID) = cPay
      txtPad(Index) = Format(cPay, CURRENCYMASK)
   ElseIf cDis <> 0 Then
      cFin = cDue - cPay
   End If
   
   vInvoice(Index + iIndex, VI_DiscountTaken) = cFin
   txtDis(Index) = Format(cFin, CURRENCYMASK)
   UpdateTotals
End Sub

Private Sub cboDueDate_DropDown()
   ShowCalendar Me
End Sub

Private Sub GetInvoices()
   Dim RdoInv As ADODB.Recordset
   Dim sVendor As String
   Dim iCount As Integer
   
   On Error GoTo DiaErr1
   
   CloseBoxes
   txtDummy.enabled = False
   lblTot = "0.00"
   cmbVnd.Clear
   
   sSql = "select fn.VIVENDOR" & vbCrLf _
      & "from dbo.fnGetOpenAP('" & cboCheckDate & "', '" & cboDueDate & "') fn" & vbCrLf _
      & "order by VIVENDOR"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv)
   If bSqlRows Then
      With RdoInv
         While Not .EOF
            If sVendor <> "" & Trim(!VIVENDOR) Then
               sVendor = "" & Trim(!VIVENDOR)
               cmbVnd.AddItem sVendor
            End If
            iCount = iCount + 1
            .MoveNext
         Wend
         .Cancel
      End With
      lblCnt = iCount     'should be updated for individual vendor
      cmdSel.enabled = True
   Else
      cmdSel.enabled = False
   End If
   Set RdoInv = Nothing
   cmdPst.enabled = False
   cmdCnl.enabled = False
   Exit Sub
DiaErr1:
   sProcName = "getinvoi"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cboDueDate_GotFocus()
   SelectFormat Me
End Sub

Private Sub cboDueDate_LostFocus()
   If Not bCancel Then
      cboDueDate = CheckDate(cboDueDate)
      GetInvoices
   End If
End Sub

Private Sub txtPad_GotFocus(Index As Integer)
   SelectFormat Me
End Sub

Private Sub txtPad_KeyDown(Index As Integer, _
                           KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
End Sub

Private Sub txtPad_KeyPress(Index As Integer, _
                            KeyAscii As Integer)
   KeyValue KeyAscii
End Sub

Private Sub txtPad_KeyUp(Index As Integer, _
                         KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdPrevPage_Click
   If KeyCode = vbKeyPageDown Then cmdNextPage_Click
End Sub

Private Sub CloseBoxes()
   Dim i As Integer
   For i = 1 To 9
      lblVnd(i) = ""
      lblDte(i) = ""
      lblInv(i) = ""
      lblPon(i) = ""
      lblDue(i) = ""
      txtDis(i) = ""
      txtDis(i).enabled = False
      txtPad(i) = ""
      txtPad(i).enabled = False
      optPad(i).Caption = ""
      optPad(i).enabled = False
      optPad(i).Value = vbUnchecked
      optPar(i).Caption = ""
      optPar(i).enabled = False
      optPar(i).Value = vbUnchecked
   Next
   
   cmdFirstPage.enabled = False
   cmdLastPage.enabled = False
   cmdNextPage.enabled = False
   cmdPrevPage.enabled = False
   cmdCnl.enabled = False
   cmdPst.enabled = False
End Sub

'Private Sub FillGrid()
'   Dim rdoInv As ADODB.RecordSet
'   Dim RdoTol As ADODB.RecordSet
'   Dim rdoPo As ADODB.RecordSet
'   Dim i As Integer
'   Dim cTotal As Currency
'   Dim sVendor As String
'   Dim sInvoice As String
'   Dim bDisFrt As Byte
'
'   On Error GoTo DiaErr1
'
'   Erase vInvoice
'   lblTot = "0.00"
'   iTotalItems = 0
'   sVendor = cmbVnd
'   If sVendor <> "ALL" Then
'      sVendor = Compress(sVendor)
'   End If
'   bDisFrt = APDiscFreight
'
'   'standard fields regardless of query
'   sSql = "SELECT VINO,VIVENDOR,VIDATE,VIFREIGHT,VITAX,VIPAY,VENICKNAME," & vbCrLf _
'          & "VIDUE,VIDUEDATE,VEBNAME,VEDISCOUNT,VEDDAYS," & vbCrLf _
'          & "CASE WHEN VEDISCOUNT > 0.0" & vbCrLf _
'          & "THEN CONVERT( varchar(10), DATEADD(day,VEDDAYS,VIDATE), 101 )" & vbCrLf _
'          & "ELSE '' END as DISCDUEDATE" & vbCrLf
'
'   If bLoadSetup Then
'      sSql = sSql & ",CHKPAMT,CHKDIS FROM VihdTable INNER JOIN " _
'             & "VndrTable ON VIVENDOR = VEREF INNER JOIN " _
'             & "ChseTable ON VIVENDOR = CHKVND AND VINO = CHKINV " _
'             & "WHERE (NOT (VIPIF=1)) AND (VIDUEDATE<='" & cboDueDate & "') " & vbCrLf
'      If chkFullName Then
'         sSql = sSql & "ORDER BY VEBNAME,VENICKNAME"
'      Else
'         sSql = sSql & "ORDER BY VENICKNAME"
'      End If
'   Else
'      sSql = sSql & "FROM VihdTable JOIN VndrTable on VIVENDOR=VEREF" & vbCrLf _
'             & "AND VIPIF = 0" & vbCrLf _
'             & "WHERE (VIDUEDATE<='" & cboDueDate & "'" & vbCrLf _
'             & "OR (VEDISCOUNT > 0.0 AND DATEADD(day,VEDDAYS,VIDATE)>='" & cboDueDate & "'))" & vbCrLf
'      If sVendor = "ALL" Then
'         '            sSql = "SELECT VINO,VIVENDOR,VIDATE,VIFREIGHT,VITAX,VIPAY,VENICKNAME," _
'         '                & "VIDUE,VIDUEDATE,VEBNAME FROM VihdTable,VndrTable WHERE " _
'         '                & "(VIVENDOR=VEREF AND VIPIF=0 AND VIDUEDATE<='" & cboDueDate & "') " _
'         '                & "OR (VIVENDOR=VEREF AND VIPIF=0 AND DATEADD(day,VEDDAYS,VIDATE)" _
'         '                & ">='" & cboDueDate & "') ORDER BY VENICKNAME"
'         'sSql = "SELECT VINO,VIVENDOR,VIDATE,VIFREIGHT,VITAX,VIPAY,VENICKNAME," & vbCrLf _
'         '    & "VIDUE,VIDUEDATE,VEBNAME" & vbCrLf
'         '            sSql = sSql & "FROM VihdTable JOIN VndrTable on VIVENDOR=VEREF" & vbCrLf _
'         '                & "AND VIPIF = 0" & vbCrLf _
'         '                & "WHERE VIDUEDATE<='" & cboDueDate & "'" & vbCrLf _
'         '                & "OR (VEDISCOUNT > 0.0 AND DATEADD(day,VEDDAYS,VIDATE)>='" & cboDueDate & "')" & vbCrLf _
'         '                & "ORDER BY VENICKNAME" & vbCrLf
'
'      Else
'         'sSql = "SELECT VINO,VIVENDOR,VIDATE,VIFREIGHT,VITAX,VIPAY,VENICKNAME," _
'         '    & "VIDUE,VIDUEDATE,VEBNAME"
'         '            sSql = sSql & "FROM VihdTable,VndrTable WHERE " _
'         '                & "(VIVENDOR=VEREF AND VIPIF=0 AND VIDUEDATE<='" & cboDueDate _
'         '                & "' AND VEREF='" & sVendor & "') OR (VIVENDOR=VEREF AND VIPIF=0 " _
'         '                & "AND DATEADD(day,VEDDAYS,VIDATE)>='" & cboDueDate & "' AND VEREF = '" _
'         '                & sVendor & "') ORDER BY VENICKNAME"
'         sSql = sSql & "AND VEREF = '" & sVendor & "'" & vbCrLf
'      End If
'      If chkFullName Then
'         sSql = sSql & "ORDER BY VEBNAME,VENICKNAME,VIDATE"
'      Else
'         sSql = sSql & "ORDER BY VENICKNAME,VIDATE"
'      End If
'      'sSql = sSql & "ORDER BY VENICKNAME"
'   End If
'   bSqlRows = clsAdoCon.GetDataSet(sSql,rdoInv)
'
'   With rdoInv
'      While Not .EOF
'         iTotalItems = iTotalItems + 1
'         sVendor = Trim(!VIVENDOR)
'         sInvoice = "" & Trim(!VINO)
'
'         ' Calculate the invoice total
'         'Old: 5011 * .125 + 1.38 = 624.99000000000001
'         'SELECT SUM(ROUND((VITQTY*VITCOST)+VITADDERS,2)) FROM ViitTable WHERE (VITVENDOR='ABCOR') AND (VITNO='12345')
'
'         'New: 5011 * .125 + 1.38 = 625.00
'         'SELECT CAST(SUM(ROUND(cast(VITQTY as decimal(12,3)) * cast(VITCOST as decimal(12,4))
'         '+ cast(VITADDERS as decimal(12,2)), 2)) as decimal(12,2)) FROM ViitTable WHERE (VITVENDOR='ABCOR') AND (VITNO='12345')
'
'
'         sSql = "SELECT CAST(SUM(ROUND(cast(VITQTY as decimal(12,3)) " & vbCrLf _
'                & "* cast(VITCOST as decimal(12,4)) " & vbCrLf _
'                & "+ cast(VITADDERS as decimal(12,2)), 2)) as decimal(12,2))" & vbCrLf _
'                & "FROM ViitTable " _
'                & "WHERE (VITVENDOR='" & sVendor & "') AND (VITNO='" _
'                & Trim(!VINO) & "')"
'         bSqlRows = clsAdoCon.GetDataSet(sSql,RdoTol)
'         If IsNull(RdoTol.Fields(0)) Then
'            cTotal = 0
'         Else
'            cTotal = RdoTol.Fields(0)
'         End If
'         RdoTol.Cancel
'         Set RdoTol = Nothing
'
'         ' Get PO number
'         ' Supports multiple PO's per invoice (not yet enabled).
'         sSql = "SELECT DISTINCT VITPO,VITPORELEASE FROM ViitTable " _
'                & "WHERE (VITVENDOR = '" & sVendor & " ') AND (VITNO = '" _
'                & sInvoice & "')"
'         bSqlRows = clsAdoCon.GetDataSet(sSql,rdoPo)
'         If bSqlRows Then
'            vInvoice(iTotalItems, VI_PO) = Format(rdoPo.Fields(0), "000000") _
'                     & "-" & Format(rdoPo.Fields(1), "##0")
'         End If
'         rdoPo.Cancel
'         Set rdoPo = Nothing
'
'         ' Assign to grid array
'         vInvoice(iTotalItems, VI_DATE) = Format(!VIDATE, DATEMASK)
'         vInvoice(iTotalItems, VI_INVNO) = sInvoice
'         vInvoice(iTotalItems, VI_INVOICETOTAL) = Format((cTotal + !VIFREIGHT _
'                  + !VITAX) - !VIPAY, CURRENCYMASK)
'         vInvoice(iTotalItems, VI_SHORTPAY) = 0
'         If bLoadSetup Then
'            vInvoice(iTotalItems, VI_AMOUNTPAID) = Format(!CHKPAMT, CURRENCYMASK)
'            If !CHKDIS = 0 And !CHKPAMT < !VIDUE Then
'               vInvoice(iTotalItems, VI_SHORTPAY) = 1
'            End If
'         Else
'            vInvoice(iTotalItems, VI_AMOUNTPAID) = "0.00"
'         End If
'
'         If bLoadSetup Then
'            vInvoice(iTotalItems, VI_CalcDiscount) = !CHKDIS
'         Else
'            vInvoice(iTotalItems, VI_CalcDiscount) = _
'                     GetAPDiscount(sVendor, sInvoice, bDisFrt, cboDueDate)
'         End If
'
'         vInvoice(iTotalItems, 6) = !VIFREIGHT
'         vInvoice(iTotalItems, 7) = !VITAX
'         vInvoice(iTotalItems, VI_PayThisInvoice) = chkAll.Value
'         vInvoice(iTotalItems, 12) = "" & Trim(!VENICKNAME)
'         vInvoice(iTotalItems, 13) = Format(!VIDUEDATE, DATEMASK)
'         vInvoice(iTotalItems, 14) = "" & Trim(!VEBNAME)
'         vInvoice(iTotalItems, VI_DiscountTaken) = vInvoice(iTotalItems, VI_CalcDiscount)
'         vInvoice(iTotalItems, VI_DUEDATEINFO) = "due date = " & Format(!VIDUEDATE, "mm/dd/yyyy")
'         If !DISCDUEDATE <> "" Then
'            vInvoice(iTotalItems, VI_DUEDATEINFO) = vInvoice(iTotalItems, VI_DUEDATEINFO) _
'                     & " Discount due date = " & !DISCDUEDATE
'         End If
'
'         .MoveNext
'      Wend
'   End With
'   Set rdoInv = Nothing
'
'   iIndex = 0
'   If iTotalItems > 0 Then
'      iCurrPage = 1
'      lblLst = (iTotalItems - 1) \ 9 + 1 'calculate # of pages
'
'      For i = 1 To 9
'
'         If i > iTotalItems Then Exit For
'
'         lblDte(i) = vInvoice(i, VI_DATE) 'Format(vInvoice(i, VI_DATE), DATEMASK)
'         'lbldte(i).ToolTipText = "Due Date: " & vInvoice(i, 13)
'         lblDte(i).ToolTipText = vInvoice(i + iIndex, VI_DUEDATEINFO)
'
'         lblInv(i) = Trim(vInvoice(i, VI_INVNO))
'         lblPon(i) = Trim(vInvoice(i, VI_PO))
'         lblDue(i) = Format(vInvoice(i, VI_INVOICETOTAL), CURRENCYMASK)
'
'         If vInvoice(i, VI_PayThisInvoice) = vbChecked Then
'            txtDis(i) = Format(vInvoice(i, VI_DiscountTaken), CURRENCYMASK)
'         Else
'            txtDis(i) = Format(vInvoice(i, VI_CalcDiscount), CURRENCYMASK)
'         End If
'
'         txtPad(i) = Format(vInvoice(i, VI_AMOUNTPAID), CURRENCYMASK)
'
'         lblVnd(i) = Trim(vInvoice(i, 12))
'         lblVnd(i).ToolTipText = vInvoice(i, 14)
'
'         optPad(i).enabled = True
'         optPad(i).Value = vInvoice(i, VI_PayThisInvoice)
'         optPar(i).enabled = True
'         optPar(i).Value = vInvoice(i, VI_SHORTPAY)
'
'      Next
'
'      On Error Resume Next
'
'      txtDummy.enabled = True
'      cmdPst.enabled = True
'      cmdCnl.enabled = True
'      cmdSel.enabled = False
'
'      If Val(lblLst) > 1 Then
'         cmdDn.enabled = True
'         cmdDn.Picture = Endn
'      End If
'
'      cboDueDate.enabled = False
'      cmbVnd.enabled = False
'      optPad(1).SetFocus
'   End If
'
'   bLoadSetup = False
'   MouseCursor 0
'   Exit Sub
'
'DiaErr1:
'   sProcName = "getvendorinv"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'End Sub
'

Private Sub FillGrid()
   Dim RdoInv As ADODB.Recordset
   'Dim RdoTol As ADODB.RecordSet
   'Dim rdoPo As ADODB.RecordSet
   Dim i As Integer
   Dim cTotal As Currency
   Dim sVendor As String
   Dim sInvoice As String
   Dim bDisFrt As Byte
   Dim iCnt As Integer
   
   On Error GoTo DiaErr1
   
   Erase vInvoice
   lblTot = "0.00"
   iTotalItems = 0
   sVendor = cmbVnd
   If sVendor <> "ALL" Then
      sVendor = Compress(sVendor)
   End If
   bDisFrt = APDiscFreight
   
' 9/18/08 changed to show existing setup and all open invoices if re-entered
''   If bLoadSetup Then
''      sSql = "select fn.*, CHKPAMT, CHKDIS" & vbCrLf _
''         & "from dbo.fnGetOpenAP('" & cboCheckDate & "', '" & cboDueDate & "') fn" & vbCrLf _
''         & "join ChseTable ON VIVENDOR = CHKVND AND VINO = CHKINV " & vbCrLf
''
''      sSql = "select fn.*, case when CHKNUM is null then 0 else 1 end as Selected" & vbCrLf _
''         & "from dbo.fnGetOpenAP('" & GetServerDate & "', '" & cboDueDate & "') fn" & vbCrLf _
''         & "left join ChseTable ON VIVENDOR = CHKVND AND VINO = CHKINV " & vbCrLf
''      If sVendor <> "ALL" Then
''         sSql = sSql & "where VIVENDOR = '" & sVendor & "'" & vbCrLf
''      End If
''
''
''   Else
''      sSql = "select fn.*" & vbCrLf _
''         & "from dbo.fnGetOpenAP('" & GetServerDate & "', '" & cboDueDate & "') fn" & vbCrLf
''      If sVendor <> "ALL" Then
''         sSql = sSql & "where VIVENDOR = '" & sVendor & "'" & vbCrLf
''      End If
''   End If
   
   sSql = "select fn.*, case when CHKNUM is null then 0 else 1 end as Selected," & vbCrLf _
      & "CHKPAMT, CHKDIS" & vbCrLf _
      & "from dbo.fnGetOpenAP('" & GetServerDate & "', '" & cboDueDate & "') fn" & vbCrLf _
      & "left join ChseTable ON VIVENDOR = CHKVND AND VINO = CHKINV " & vbCrLf
   If sVendor <> "ALL" Then
      sSql = sSql & "where VIVENDOR = '" & sVendor & "'" & vbCrLf
      'sSql = sSql & "where VIVENDOR like '" & sVendor & "%'" & vbCrLf  ' leading chars for testing
   End If
   
'   If chkFullName Then
'      sSql = sSql & "ORDER BY VEBNAME,VIVENDOR,VIDATE"
'   Else
'      sSql = sSql & "ORDER BY VIVENDOR,VIDATE"
'   End If
      
   If chkFullName Then
      sSql = sSql & "ORDER BY VEBNAME, VIVENDOR"
   Else
      sSql = sSql & "ORDER BY VIVENDOR"
   End If
      
   If chkInvoiceNumber Then
      sSql = sSql & ", VINO"
   Else
      sSql = sSql & ", VIDATE"
   End If
      
   iCnt = 0
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv)
   With RdoInv
      While Not .EOF
         iTotalItems = iTotalItems + 1
If iTotalItems >= 30 Then
Debug.Print iTotalItems
End If
         sVendor = Trim(!VIVENDOR)
         sInvoice = "" & Trim(!VINO)
         
         cTotal = !CalcTotal

         vInvoice(iTotalItems, VI_PO) = Format(!VITPO, "000000") _
            & "-" & Format(!VITPORELEASE, "##0")
         
         ' Assign to grid array
         vInvoice(iTotalItems, VI_DATE) = Format(!VIDATE, DATEMASK)
         vInvoice(iTotalItems, VI_INVNO) = sInvoice
         vInvoice(iTotalItems, VI_INVOICETOTAL) = Format(cTotal - !VIPAY, CURRENCYMASK)
         vInvoice(iTotalItems, VI_SHORTPAY) = 0
'         If bLoadSetup Then
'            vInvoice(iTotalItems, VI_AMOUNTPAID) = Format(!CHKPAMT, CURRENCYMASK)
'            If !CHKDIS + !CHKPAMT < !INVTOTAL - !VIPAY Then
'               vInvoice(iTotalItems, VI_SHORTPAY) = 1
'            End If
'         Else
'            vInvoice(iTotalItems, VI_AMOUNTPAID) = "0.00"
'         End If
         
         If !Selected Then
            vInvoice(iTotalItems, VI_AMOUNTPAID) = Format(!CHKPAMT, CURRENCYMASK)
            If !CHKDIS + !CHKPAMT < !INVTOTAL - !VIPAY Then
               vInvoice(iTotalItems, VI_SHORTPAY) = 1
            End If
            vInvoice(iTotalItems, VI_CalcDiscount) = !CHKDIS
         Else
            vInvoice(iTotalItems, VI_AMOUNTPAID) = "0.00"
            vInvoice(iTotalItems, VI_CalcDiscount) = !DiscAvail
         End If
         
         
'         If bLoadSetup Then
'            vInvoice(iTotalItems, VI_CalcDiscount) = !CHKDIS
'         Else
'            vInvoice(iTotalItems, VI_CalcDiscount) = !DiscAvail
'         End If
         If IsNull(!TakeDiscByDate) Then
            vInvoice(iTotalItems, VI_DISCOUNTCUTOFFDATE) = ""
         Else
            vInvoice(iTotalItems, VI_DISCOUNTCUTOFFDATE) = Format(!TakeDiscByDate, "mm/dd/yy")
         End If
         
         vInvoice(iTotalItems, 6) = !VIFREIGHT
         vInvoice(iTotalItems, 7) = !VITAX
         'vInvoice(iTotalItems, VI_PayThisInvoice) = chkAll.Value
         vInvoice(iTotalItems, VI_PayThisInvoice) = !Selected
         vInvoice(iTotalItems, 12) = "" & Trim(!VIVENDOR)
         vInvoice(iTotalItems, 13) = Format(!VIDUEDATE, DATEMASK)
         vInvoice(iTotalItems, 14) = "" & Trim(!VEBNAME)
         vInvoice(iTotalItems, VI_DiscountTaken) = vInvoice(iTotalItems, VI_CalcDiscount)
         vInvoice(iTotalItems, VI_DUEDATEINFO) = "due date = " & Format(!VIDUEDATE, "mm/dd/yy")
         If IsNull(!TakeDiscByDate) Then
            vInvoice(iTotalItems, VI_DUEDATEINFO) = vInvoice(iTotalItems, VI_DUEDATEINFO) _
                     & " Discount due date = " & Format(!TakeDiscByDate, "mm/dd/yy")
         End If
         
         If (iCnt = 5000 - 1) Then
            
            sMsg = "There Are More That 5000 Invoices Returned For The Selected Search." & _
                  vbCrLf & "Please Narrow the Search."
            MsgBox sMsg, vbInformation, Caption
            GoTo ExitWhile
         End If
         iCnt = iCnt + 1
         .MoveNext
      Wend
   End With
   
ExitWhile:
   Set RdoInv = Nothing
   
   iIndex = 0
   If iTotalItems > 0 Then
      iCurrPage = 1
      lblLst = (iTotalItems - 1) \ 9 + 1 'calculate # of pages
      
      For i = 1 To 9
         
         If i > iTotalItems Then Exit For
         
         lblDte(i) = vInvoice(i, VI_DATE) 'Format(vInvoice(i, VI_DATE), DATEMASK)
         lblDte(i).ToolTipText = vInvoice(i + iIndex, VI_DUEDATEINFO)
         
         lblInv(i) = Trim(vInvoice(i, VI_INVNO))
         lblPon(i) = Trim(vInvoice(i, VI_PO))
         lblDue(i) = Format(vInvoice(i, VI_INVOICETOTAL), CURRENCYMASK)
         
         If vInvoice(i, VI_PayThisInvoice) = vbChecked Then
            txtDis(i) = Format(vInvoice(i, VI_DiscountTaken), CURRENCYMASK)
         Else
            txtDis(i) = Format(vInvoice(i, VI_CalcDiscount), CURRENCYMASK)
         End If
         lblDiscDate(i) = vInvoice(i, VI_DISCOUNTCUTOFFDATE)
         
         txtPad(i) = Format(vInvoice(i, VI_AMOUNTPAID), CURRENCYMASK)
         
         lblVnd(i) = Trim(vInvoice(i, 12))
         lblVnd(i).ToolTipText = vInvoice(i, 14)
         
         optPad(i).enabled = True
         optPad(i).Value = vInvoice(i, VI_PayThisInvoice)
         optPar(i).enabled = True
         optPar(i).Value = vInvoice(i, VI_SHORTPAY)
         
      Next
      
      'On Error Resume Next
      
      txtDummy.enabled = True
      cmdPst.enabled = True
      cmdCnl.enabled = True
      cmdSel.enabled = False
      
      'If Val(lblLst) > 1 Then
      '   cmdDn.enabled = True
      '   cmdDn.Picture = Endn
      'End If
      
      
      
      cboDueDate.enabled = False
      cboCheckDate.enabled = False
      cmbVnd.enabled = False
      optPad(1).SetFocus
   End If
   
   cmdFirstPage.enabled = True
   cmdLastPage.enabled = True
   cmdNextPage.enabled = True
   cmdPrevPage.enabled = True

   bLoadSetup = False
   MouseCursor 0
   UpdateTotals
   Exit Sub
   
DiaErr1:
   sProcName = "getvendorinv"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Sub GetNextGroup()
   Dim i As Integer
   
   On Error GoTo DiaErr1
   
   bPage = True
   iIndex = (Val(lblPge) - 1) * 9
   
   For i = 1 To 9
      If i + iIndex > iTotalItems Then
         lblDte(i) = ""
         lblDte(i).ToolTipText = ""
         
         lblInv(i) = ""
         lblPon(i) = ""
         
         lblVnd(i) = ""
         lblVnd(i).ToolTipText = ""
         
         lblDue(i) = ""
         lblDue(i).enabled = False
         
         txtPad(i) = ""
         txtPad(i).enabled = False
         
         txtDis(i) = ""
         txtDis(i).enabled = False
         Me.lblDiscDate(i) = ""
         
         optPar(i).Caption = ""
         optPar(i).enabled = False
         optPar(i).Value = vbUnchecked
         
         optPad(i).Caption = ""
         optPad(i).enabled = False
         optPad(i).Value = vbUnchecked
      Else
         lblVnd(i) = vInvoice(i + iIndex, 12)
         lblVnd(i).ToolTipText = vInvoice(i + iIndex, 14)
         
         lblDte(i) = vInvoice(i + iIndex, VI_DATE)
         'lblDte(i).ToolTipText = "Due Date: " & vInvoice(i + iIndex, 13)
         lblDte(i).ToolTipText = vInvoice(i + iIndex, VI_DUEDATEINFO)
         
         lblInv(i) = vInvoice(i + iIndex, VI_INVNO)
         lblPon(i) = vInvoice(i + iIndex, VI_PO)
         
         lblDue(i) = Format(vInvoice(i + iIndex, VI_INVOICETOTAL), CURRENCYMASK)
         lblDue(i).enabled = True
         
         txtPad(i) = Format(vInvoice(i + iIndex, VI_AMOUNTPAID), CURRENCYMASK)
         
         txtDis(i) = Format(vInvoice(i + iIndex, VI_DiscountTaken), CURRENCYMASK)
         lblDiscDate(i) = vInvoice(i + iIndex, VI_DISCOUNTCUTOFFDATE)
         
         optPar(i).Caption = "___"
         optPar(i).Value = Val(vInvoice(i + iIndex, VI_SHORTPAY))
         optPar(i).enabled = True
         
         optPad(i).Caption = "___"
         optPad(i).Value = Val(vInvoice(i + iIndex, VI_PayThisInvoice))
         optPad(i).enabled = True
         
         If optPad(i) Then
            txtPad(i).enabled = True
            txtDis(i).enabled = True
         Else
            txtPad(i).enabled = False
            txtDis(i).enabled = False
         End If
      End If
   Next
   'On Error Resume Next
   'txtPad(1).SetFocus
   bPage = False
   Exit Sub
DiaErr1:
   bPage = False
   sProcName = "getnextg"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub UpdateTotals()
   Dim i As Integer
   On Error Resume Next
   Dim cTotal As Currency
   For i = 1 To iTotalItems
      If vInvoice(i, VI_PayThisInvoice) = vbChecked Then
         cTotal = cTotal + vInvoice(i, VI_AMOUNTPAID)
      End If
   Next
   lblTot = Format(cTotal, CURRENCYMASK)
   
End Sub

Private Sub txtPad_LostFocus(Index As Integer)
   Dim cTotalItem As Currency
   Dim cDue As Currency
   Dim cPay As Currency
   Dim cDis As Currency
   Dim cFin As Currency
   
   txtPad(Index) = CheckLen(txtPad(Index), 10)
   If ("" & Trim(txtPad(Index))) = "" Then
      txtPad(Index) = "0.00"
   End If
   cDue = CCur(lblDue(Index))
   cPay = CCur(txtPad(Index))
   cDis = CCur(txtDis(Index))
   
   If cPay = 0 And cDue <> 0 Then
      optPad(Index).Value = vbUnchecked
   End If
   
    ' Adding if the Total is not Same as the
   If (cPay < cDue) Then
      txtDis(Index).enabled = True
   End If
      
   If optPad(Index).Value = vbChecked Then
      cFin = cPay
      If cDue >= 0 Then
         If (cPay + cDis) > cDue Or cPay < 0 Then
            cFin = cDue - cDis
         End If
      Else
         If cPay >= 0 Or cPay < cDue Then
            cFin = cDue
         End If
      End If
      If Not optPar(Index) And cDue >= 0 _
                    And vInvoice(Index + iIndex, VI_CalcDiscount) > 0 Then
         cDis = cDue - cPay
      End If
      txtDis(Index) = Format(cDis, CURRENCYMASK)
      txtPad(Index) = Format(cFin, CURRENCYMASK)
      vInvoice(Index + iIndex, VI_AMOUNTPAID) = cFin
      vInvoice(Index + iIndex, VI_DiscountTaken) = cDis
   End If
   UpdateTotals
End Sub

Private Sub PostChecks()
   Dim i As Integer
   Dim iChkNum As Integer ' Tempory place holder in check setup
   Dim iResponse As Integer
   Dim sVendor As String
   Dim sOldVendor As String
   Dim cDiscount As Currency
   Dim cPaid As Currency
   Dim cTotal As Currency
   Dim cCheckTotal As Currency
   
   MouseCursor 13
   On Error GoTo DiaErr1
   sSql = "DELETE FROM ChseTable"
   clsADOCon.ExecuteSql sSql
   
   On Error Resume Next
   iChkNum = 0
   Err = 0
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0

   'sOldVendor = vInvoice(1, 12)
   
   For i = 1 To iTotalItems
      If vInvoice(i, VI_PayThisInvoice) = vbChecked Then
         sVendor = vInvoice(i, 12)
         If sVendor <> sOldVendor Then
            If cCheckTotal < 0 Then
               sMsg = "Check For " & sOldVendor & " Is Negative" _
                      & vbCrLf & "Cannot Post Check Setup."
               MsgBox sMsg, vbInformation, Caption
               clsADOCon.RollbackTrans
               Exit Sub
            End If
            iChkNum = iChkNum + 1
            cCheckTotal = 0
         End If
         sOldVendor = sVendor
         
         ' vInvoice() doc
         ' 3 = Invoice Total
         ' 4 = Amount Paid
         ' 5 = Discount
         If vInvoice(i, VI_SHORTPAY) = 1 Then ' Short Pay
            cDiscount = 0
            cTotal = vInvoice(i, VI_INVOICETOTAL)
            cPaid = vInvoice(i, VI_AMOUNTPAID)
         Else ' PIF with or without discount
            cDiscount = vInvoice(i, VI_DiscountTaken)
            cPaid = vInvoice(i, VI_AMOUNTPAID)
            cTotal = vInvoice(i, VI_INVOICETOTAL)
         End If
         Debug.Print "setup check # " & Format(iChkNum, "000#") & " for " & sVendor
         sSql = "INSERT INTO ChseTable(CHKNUM,CHKVND,CHKINV,CHKAMT," _
                & "CHKPAMT,CHKDIS,CHKDATE,CHKBY)" _
                & " VALUES( " _
                & "'" & Format(iChkNum, "000#") & "'," _
                & "'" & Compress(sVendor) & "'," _
                & "'" & vInvoice(i, VI_INVNO) & "'," _
                & cTotal & "," _
                & cPaid & "," _
                & cDiscount & "," _
                & "'" & Format(ES_SYSDATE, "mm/dd/yy") & "'," _
                & "'" & sInitials & "')"
         clsADOCon.ExecuteSql sSql
         cCheckTotal = cCheckTotal + cPaid
      End If
   Next
   If cCheckTotal < 0 Then
      sMsg = "Check For " & sOldVendor & " Is Negative" _
             & vbCrLf & "Cannot Post Check Setup."
      MsgBox sMsg, vbInformation, Caption
      clsADOCon.RollbackTrans
      Exit Sub
   End If
   
   ' Record setup as of date.
   sSql = "UPDATE Preferences SET PreChkSetDte = '" & cboDueDate & "'"
   clsADOCon.ExecuteSql sSql
   
         
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
   Else
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
      MsgBox "Couldn't Successfully Record Check Setup.", _
         vbInformation, Caption
      Exit Sub
   End If
   
   iResponse = MsgBox("Checks Successfully Posted." & vbCrLf _
               & "Do You Wish To Print Them Now?", ES_YESQUESTION, Caption)
   If iResponse = vbYes Then
      Unload Me
      diaAPf03a.Show
   End If
   MouseCursor 0
   Unload Me
   Exit Sub
   
DiaErr1:
   sProcName = "postchec"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub LoadCheckSetup()
   Dim RdoSet As ADODB.Recordset
   Dim rdoTmp As ADODB.Recordset
   Dim sAsOf As String
   Dim bResponse As Byte
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT PreChkSetDte FROM Preferences"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSet)
   
   If bSqlRows Then
      
      If IsNull(RdoSet!PreChkSetDte) Then Exit Sub
      
      sAsOf = Format(RdoSet!PreChkSetDte, "mm/dd/yy")
      sMsg = "Existing Check Setup As Of " & sAsOf & vbCrLf _
             & "Found. Clear Existing Setup?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      
      If bResponse = vbNo Then
         
         ' Found existing check setup load it.
         bLoadSetup = True
         cboDueDate = sAsOf
         GetInvoices
         
         sSql = "SELECT DISTINCT COUNT(CHKVND) FROM ChseTable"
         bSqlRows = clsADOCon.GetDataSet(sSql, rdoTmp, ES_FORWARD)
         If bSqlRows Then
            If rdoTmp.Fields(0) > 1 Then
               cmbVnd.enabled = True
               cmbVnd = "ALL"
               lblNme = "*** Multiple Vendors Selected ***"
               Refresh
            Else
               Set rdoTmp = Nothing
               sSql = "SELECT VENICKNAME FROM ChseTable INNER JOIN " _
                      & "VndrTable ON ChseTable.CHKVND = VndrTable.VEREF"
               bSqlRows = clsADOCon.GetDataSet(sSql, rdoTmp, ES_FORWARD)
               If bSqlRows Then
                  cmbVnd = rdoTmp.Fields(0)
                  bGoodVendor = FindVendor(Me)
               End If
            End If
            
         End If
         Set rdoTmp = Nothing
         'chkAll.Value = vbChecked
         ' Trigger the "select" button
         cmdSel_Click
      Else
         bLoadSetup = False
         sSql = "UPDATE Preferences SET PreChkSetDte = null"
         clsADOCon.ExecuteSql sSql
         sSql = "DELETE FROM ChseTable"
         clsADOCon.ExecuteSql sSql
      End If
   End If
   Set RdoSet = Nothing
   Exit Sub
DiaErr1:
   sProcName = "loadchecksetup"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub
