VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PackPSe07a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revise Request Date For Unshipped PS Items"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdUp 
      DisabledPicture =   "PackPSe07a.frx":0000
      DownPicture     =   "PackPSe07a.frx":04F2
      Enabled         =   0   'False
      Height          =   372
      Left            =   6840
      MaskColor       =   &H00000000&
      Picture         =   "PackPSe07a.frx":09E4
      Style           =   1  'Graphical
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   5400
      Width           =   400
   End
   Begin VB.CommandButton cmdDn 
      DisabledPicture =   "PackPSe07a.frx":0ED6
      DownPicture     =   "PackPSe07a.frx":13C8
      Enabled         =   0   'False
      Height          =   372
      Left            =   6840
      MaskColor       =   &H00000000&
      Picture         =   "PackPSe07a.frx":18BA
      Style           =   1  'Graphical
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   5784
      Width           =   400
   End
   Begin VB.Frame z2 
      Height          =   40
      Left            =   240
      TabIndex        =   87
      Top             =   1750
      Width           =   6972
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PackPSe07a.frx":1DAC
      Style           =   1  'Graphical
      TabIndex        =   86
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtCpo 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Customer Purchase Orders"
      Top             =   1320
      Width           =   2415
   End
   Begin VB.ComboBox txtReq 
      Enabled         =   0   'False
      Height          =   315
      Index           =   10
      Left            =   6120
      TabIndex        =   12
      Tag             =   "4"
      Top             =   4960
      Width           =   1095
   End
   Begin VB.ComboBox txtReq 
      Enabled         =   0   'False
      Height          =   315
      Index           =   9
      Left            =   6120
      TabIndex        =   11
      Tag             =   "4"
      Top             =   4640
      Width           =   1095
   End
   Begin VB.ComboBox txtReq 
      Enabled         =   0   'False
      Height          =   315
      Index           =   8
      Left            =   6120
      TabIndex        =   10
      Tag             =   "4"
      Top             =   4320
      Width           =   1095
   End
   Begin VB.ComboBox txtReq 
      Enabled         =   0   'False
      Height          =   315
      Index           =   7
      Left            =   6120
      TabIndex        =   9
      Tag             =   "4"
      Top             =   4000
      Width           =   1095
   End
   Begin VB.ComboBox txtReq 
      Enabled         =   0   'False
      Height          =   315
      Index           =   6
      Left            =   6120
      TabIndex        =   8
      Tag             =   "4"
      Top             =   3680
      Width           =   1095
   End
   Begin VB.ComboBox txtReq 
      Enabled         =   0   'False
      Height          =   315
      Index           =   5
      Left            =   6120
      TabIndex        =   7
      Tag             =   "4"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.ComboBox txtReq 
      Enabled         =   0   'False
      Height          =   315
      Index           =   4
      Left            =   6120
      TabIndex        =   5
      Tag             =   "4"
      Top             =   3040
      Width           =   1095
   End
   Begin VB.ComboBox txtReq 
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   6120
      TabIndex        =   6
      Tag             =   "4"
      Top             =   2720
      Width           =   1095
   End
   Begin VB.ComboBox txtReq 
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   6120
      TabIndex        =   4
      Tag             =   "4"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.ComboBox txtReq 
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   6120
      TabIndex        =   3
      Tag             =   "4"
      Top             =   2080
      Width           =   1095
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Tag             =   "3"
      Top             =   600
      Width           =   1555
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "&Select"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6240
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Retrieve List Of Items"
      Top             =   1320
      Width           =   915
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6240
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   6480
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6225
      FormDesignWidth =   7365
   End
   Begin VB.Image Dsup 
      Height          =   300
      Left            =   720
      Picture         =   "PackPSe07a.frx":255A
      Top             =   5520
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Enup 
      Height          =   300
      Left            =   480
      Picture         =   "PackPSe07a.frx":2A4C
      Top             =   5520
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Endn 
      Height          =   300
      Left            =   1080
      Picture         =   "PackPSe07a.frx":2F3E
      Top             =   5520
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Dsdn 
      Height          =   300
      Left            =   1320
      Picture         =   "PackPSe07a.frx":3430
      Top             =   5520
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "of"
      Height          =   255
      Index           =   8
      Left            =   6000
      TabIndex        =   85
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Page"
      Height          =   255
      Index           =   9
      Left            =   4920
      TabIndex        =   84
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label lblLst 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6360
      TabIndex        =   83
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label lblPge 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5520
      TabIndex        =   82
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label lblPart 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S00005"
      Height          =   285
      Index           =   10
      Left            =   3000
      TabIndex        =   81
      ToolTipText     =   "Part Number"
      Top             =   4960
      Width           =   3000
   End
   Begin VB.Label lblPsItm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "389"
      Height          =   285
      Index           =   10
      Left            =   2520
      TabIndex        =   80
      ToolTipText     =   "Packing Slip Item"
      Top             =   4960
      Width           =   405
   End
   Begin VB.Label lblSon 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S00005"
      Height          =   285
      Index           =   10
      Left            =   240
      TabIndex        =   79
      ToolTipText     =   "Sales Order"
      Top             =   4960
      Width           =   645
   End
   Begin VB.Label lblSitm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "389"
      Height          =   285
      Index           =   10
      Left            =   915
      TabIndex        =   78
      ToolTipText     =   "Sales Order Item"
      Top             =   4960
      Width           =   405
   End
   Begin VB.Label lblItRev 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AB"
      Height          =   285
      Index           =   10
      Left            =   1335
      TabIndex        =   77
      Top             =   4960
      Width           =   285
   End
   Begin VB.Label lblPsnum 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PS000005"
      Height          =   285
      Index           =   10
      Left            =   1700
      TabIndex        =   76
      ToolTipText     =   "Packing Slip"
      Top             =   4960
      Width           =   820
   End
   Begin VB.Label lblPart 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S00005"
      Height          =   285
      Index           =   9
      Left            =   3000
      TabIndex        =   75
      ToolTipText     =   "Part Number"
      Top             =   4640
      Width           =   3000
   End
   Begin VB.Label lblPsItm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "389"
      Height          =   285
      Index           =   9
      Left            =   2520
      TabIndex        =   74
      ToolTipText     =   "Packing Slip Item"
      Top             =   4640
      Width           =   405
   End
   Begin VB.Label lblSon 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S00005"
      Height          =   285
      Index           =   9
      Left            =   240
      TabIndex        =   73
      ToolTipText     =   "Sales Order"
      Top             =   4640
      Width           =   645
   End
   Begin VB.Label lblSitm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "389"
      Height          =   285
      Index           =   9
      Left            =   915
      TabIndex        =   72
      ToolTipText     =   "Sales Order Item"
      Top             =   4640
      Width           =   405
   End
   Begin VB.Label lblItRev 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AB"
      Height          =   285
      Index           =   9
      Left            =   1335
      TabIndex        =   71
      Top             =   4640
      Width           =   285
   End
   Begin VB.Label lblPsnum 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PS000005"
      Height          =   285
      Index           =   9
      Left            =   1700
      TabIndex        =   70
      ToolTipText     =   "Packing Slip"
      Top             =   4640
      Width           =   820
   End
   Begin VB.Label lblPart 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S00005"
      Height          =   285
      Index           =   8
      Left            =   3000
      TabIndex        =   69
      ToolTipText     =   "Part Number"
      Top             =   4320
      Width           =   3000
   End
   Begin VB.Label lblPsItm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "389"
      Height          =   285
      Index           =   8
      Left            =   2520
      TabIndex        =   68
      ToolTipText     =   "Packing Slip Item"
      Top             =   4320
      Width           =   405
   End
   Begin VB.Label lblSon 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S00005"
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   67
      ToolTipText     =   "Sales Order"
      Top             =   4320
      Width           =   645
   End
   Begin VB.Label lblSitm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "389"
      Height          =   285
      Index           =   8
      Left            =   915
      TabIndex        =   66
      ToolTipText     =   "Sales Order Item"
      Top             =   4320
      Width           =   405
   End
   Begin VB.Label lblItRev 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AB"
      Height          =   285
      Index           =   8
      Left            =   1335
      TabIndex        =   65
      Top             =   4320
      Width           =   285
   End
   Begin VB.Label lblPsnum 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PS000005"
      Height          =   285
      Index           =   8
      Left            =   1700
      TabIndex        =   64
      ToolTipText     =   "Packing Slip"
      Top             =   4320
      Width           =   820
   End
   Begin VB.Label lblPart 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S00005"
      Height          =   285
      Index           =   7
      Left            =   3000
      TabIndex        =   63
      ToolTipText     =   "Part Number"
      Top             =   4000
      Width           =   3000
   End
   Begin VB.Label lblPsItm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "389"
      Height          =   285
      Index           =   7
      Left            =   2520
      TabIndex        =   62
      ToolTipText     =   "Packing Slip Item"
      Top             =   4000
      Width           =   405
   End
   Begin VB.Label lblSon 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S00005"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   61
      ToolTipText     =   "Sales Order"
      Top             =   4000
      Width           =   645
   End
   Begin VB.Label lblSitm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "389"
      Height          =   285
      Index           =   7
      Left            =   915
      TabIndex        =   60
      ToolTipText     =   "Sales Order Item"
      Top             =   4000
      Width           =   405
   End
   Begin VB.Label lblItRev 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AB"
      Height          =   285
      Index           =   7
      Left            =   1335
      TabIndex        =   59
      Top             =   4000
      Width           =   285
   End
   Begin VB.Label lblPsnum 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PS000005"
      Height          =   285
      Index           =   7
      Left            =   1700
      TabIndex        =   58
      ToolTipText     =   "Packing Slip"
      Top             =   4000
      Width           =   820
   End
   Begin VB.Label lblPart 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S00005"
      Height          =   285
      Index           =   6
      Left            =   3000
      TabIndex        =   57
      ToolTipText     =   "Part Number"
      Top             =   3680
      Width           =   3000
   End
   Begin VB.Label lblPsItm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "389"
      Height          =   285
      Index           =   6
      Left            =   2520
      TabIndex        =   56
      ToolTipText     =   "Packing Slip Item"
      Top             =   3680
      Width           =   405
   End
   Begin VB.Label lblSon 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S00005"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   55
      ToolTipText     =   "Sales Order"
      Top             =   3680
      Width           =   645
   End
   Begin VB.Label lblSitm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "389"
      Height          =   285
      Index           =   6
      Left            =   915
      TabIndex        =   54
      ToolTipText     =   "Sales Order Item"
      Top             =   3680
      Width           =   405
   End
   Begin VB.Label lblItRev 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AB"
      Height          =   285
      Index           =   6
      Left            =   1335
      TabIndex        =   53
      Top             =   3680
      Width           =   285
   End
   Begin VB.Label lblPsnum 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PS000005"
      Height          =   285
      Index           =   6
      Left            =   1700
      TabIndex        =   52
      ToolTipText     =   "Packing Slip"
      Top             =   3680
      Width           =   820
   End
   Begin VB.Label lblPart 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S00005"
      Height          =   285
      Index           =   5
      Left            =   3000
      TabIndex        =   51
      ToolTipText     =   "Part Number"
      Top             =   3360
      Width           =   3000
   End
   Begin VB.Label lblPsItm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "389"
      Height          =   285
      Index           =   5
      Left            =   2520
      TabIndex        =   50
      ToolTipText     =   "Packing Slip Item"
      Top             =   3360
      Width           =   405
   End
   Begin VB.Label lblSon 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S00005"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   49
      ToolTipText     =   "Sales Order"
      Top             =   3360
      Width           =   645
   End
   Begin VB.Label lblSitm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "389"
      Height          =   285
      Index           =   5
      Left            =   915
      TabIndex        =   48
      ToolTipText     =   "Sales Order Item"
      Top             =   3360
      Width           =   405
   End
   Begin VB.Label lblItRev 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AB"
      Height          =   285
      Index           =   5
      Left            =   1335
      TabIndex        =   47
      Top             =   3360
      Width           =   285
   End
   Begin VB.Label lblPsnum 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PS000005"
      Height          =   285
      Index           =   5
      Left            =   1700
      TabIndex        =   46
      ToolTipText     =   "Packing Slip"
      Top             =   3360
      Width           =   820
   End
   Begin VB.Label lblPart 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S00005"
      Height          =   285
      Index           =   4
      Left            =   3000
      TabIndex        =   45
      ToolTipText     =   "Part Number"
      Top             =   3040
      Width           =   3000
   End
   Begin VB.Label lblPsItm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "389"
      Height          =   285
      Index           =   4
      Left            =   2520
      TabIndex        =   44
      ToolTipText     =   "Packing Slip Item"
      Top             =   3040
      Width           =   405
   End
   Begin VB.Label lblSon 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S00005"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   43
      ToolTipText     =   "Sales Order"
      Top             =   3040
      Width           =   645
   End
   Begin VB.Label lblSitm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "389"
      Height          =   285
      Index           =   4
      Left            =   915
      TabIndex        =   42
      ToolTipText     =   "Sales Order Item"
      Top             =   3040
      Width           =   405
   End
   Begin VB.Label lblItRev 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AB"
      Height          =   285
      Index           =   4
      Left            =   1335
      TabIndex        =   41
      Top             =   3040
      Width           =   285
   End
   Begin VB.Label lblPsnum 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PS000005"
      Height          =   285
      Index           =   4
      Left            =   1700
      TabIndex        =   40
      ToolTipText     =   "Packing Slip"
      Top             =   3040
      Width           =   820
   End
   Begin VB.Label lblPart 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S00005"
      Height          =   285
      Index           =   3
      Left            =   3000
      TabIndex        =   39
      ToolTipText     =   "Part Number"
      Top             =   2720
      Width           =   3000
   End
   Begin VB.Label lblPsItm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "389"
      Height          =   285
      Index           =   3
      Left            =   2520
      TabIndex        =   38
      ToolTipText     =   "Packing Slip Item"
      Top             =   2720
      Width           =   405
   End
   Begin VB.Label lblSon 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S00005"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   37
      ToolTipText     =   "Sales Order"
      Top             =   2720
      Width           =   645
   End
   Begin VB.Label lblSitm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "389"
      Height          =   285
      Index           =   3
      Left            =   915
      TabIndex        =   36
      ToolTipText     =   "Sales Order Item"
      Top             =   2720
      Width           =   405
   End
   Begin VB.Label lblItRev 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AB"
      Height          =   285
      Index           =   3
      Left            =   1335
      TabIndex        =   35
      Top             =   2720
      Width           =   285
   End
   Begin VB.Label lblPsnum 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PS000005"
      Height          =   285
      Index           =   3
      Left            =   1700
      TabIndex        =   34
      ToolTipText     =   "Packing Slip"
      Top             =   2720
      Width           =   820
   End
   Begin VB.Label lblPart 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S00005"
      Height          =   285
      Index           =   2
      Left            =   3000
      TabIndex        =   33
      ToolTipText     =   "Part Number"
      Top             =   2400
      Width           =   3000
   End
   Begin VB.Label lblPsItm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "389"
      Height          =   285
      Index           =   2
      Left            =   2520
      TabIndex        =   32
      ToolTipText     =   "Packing Slip Item"
      Top             =   2400
      Width           =   405
   End
   Begin VB.Label lblSon 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S00005"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   31
      ToolTipText     =   "Sales Order"
      Top             =   2400
      Width           =   645
   End
   Begin VB.Label lblSitm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "389"
      Height          =   285
      Index           =   2
      Left            =   915
      TabIndex        =   30
      ToolTipText     =   "Sales Order Item"
      Top             =   2400
      Width           =   405
   End
   Begin VB.Label lblItRev 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AB"
      Height          =   285
      Index           =   2
      Left            =   1335
      TabIndex        =   29
      Top             =   2400
      Width           =   285
   End
   Begin VB.Label lblPsnum 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PS000005"
      Height          =   285
      Index           =   2
      Left            =   1700
      TabIndex        =   28
      ToolTipText     =   "Packing Slip"
      Top             =   2400
      Width           =   820
   End
   Begin VB.Label lblPart 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S00005"
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   27
      ToolTipText     =   "Part Number"
      Top             =   2080
      Width           =   3000
   End
   Begin VB.Label lblPsItm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "389"
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   26
      ToolTipText     =   "Packing Slip Item"
      Top             =   2080
      Width           =   405
   End
   Begin VB.Label lblSon 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S00005"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   25
      ToolTipText     =   "Sales Order"
      Top             =   2080
      Width           =   645
   End
   Begin VB.Label lblSitm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "389"
      Height          =   285
      Index           =   1
      Left            =   915
      TabIndex        =   24
      ToolTipText     =   "Sales Order Item"
      Top             =   2080
      Width           =   405
   End
   Begin VB.Label lblItRev 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AB"
      Height          =   285
      Index           =   1
      Left            =   1335
      TabIndex        =   23
      Top             =   2080
      Width           =   285
   End
   Begin VB.Label lblPsnum 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PS000005"
      Height          =   285
      Index           =   1
      Left            =   1700
      TabIndex        =   22
      ToolTipText     =   "Packing Slip"
      Top             =   2080
      Width           =   820
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Requested        "
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
      Index           =   5
      Left            =   6120
      TabIndex        =   21
      Top             =   1850
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number                                              "
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
      Left            =   3000
      TabIndex        =   20
      Top             =   1850
      Width           =   2985
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Packing Slip       "
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
      Left            =   1700
      TabIndex        =   19
      Top             =   1850
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order            "
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
      Index           =   2
      Left            =   240
      TabIndex        =   18
      Top             =   1850
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer "
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   17
      Top             =   600
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO's Beginning With"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   16
      Top             =   1320
      Width           =   1785
   End
   Begin VB.Label txtNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   15
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   6
      Left            =   4680
      TabIndex        =   14
      Top             =   1320
      Width           =   1785
   End
End
Attribute VB_Name = "PackPSe07a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'7/14/05 New (INTCOA custom)
Option Explicit
Dim bOnLoad As Byte
Dim iIndex As Integer
Dim iLastIndex As Integer
Dim iCurrPage As Integer
Dim iTotalItems As Integer

Dim sSales(801, 7) As String
Private CurrentCustomer As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCst_Click()
   'CloseBoxes
   'FindCustomer Me, cmbCst, False
   'FillPurchaseOrders
   cmbCst_LostFocus
End Sub


Private Sub cmbCst_GotFocus()
   'Debug.Print "CMBCST_GOTFOCUS"
End Sub

Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If cmbCst <> CurrentCustomer Then
      FindCustomer Me, cmbCst, False
      FillPurchaseOrders
      CurrentCustomer = cmbCst
   End If
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdDn_Click()
   iIndex = iIndex + 10
   If iIndex > iLastIndex Then iIndex = iLastIndex
   iCurrPage = iCurrPage + 1
   If iCurrPage > Val(lblLst) Then Val (lblLst)
   lblPge = iCurrPage
   GetNextGroup
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2207
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdSel_Click()
   If txtNme.ForeColor = ES_RED Then
      MsgBox "Please Select A Valid Customer.", _
         vbInformation, Caption
   Else
      SelectSalesOrders
   End If
   
End Sub

Private Sub cmdUp_Click()
   iIndex = iIndex - 10
   If iIndex < 0 Then iIndex = 0
   iCurrPage = iCurrPage - 1
   If iCurrPage < 1 Then iCurrPage = 1
   lblPge = iCurrPage
   GetNextGroup
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCustomers
      If cmbCst.ListCount > 0 Then
         cmbCst = cmbCst.List(0)
         FindCustomer Me, cmbCst, False
      End If
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set PackPSe07a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   CloseBoxes
   txtCpo.ToolTipText = "Select Customer PO, Leading Characters Or Blank For All"
   
End Sub


Public Sub SelectSalesOrders()
   Dim RdoItm As ADODB.Recordset
   Dim iRow As Integer
   Erase sSales
   CloseBoxes
   On Error GoTo DiaErr1
   sSql = "SELECT SONUMBER,SOTYPE,SOCUST,SOPO,ITSO,ITNUMBER," _
          & "ITREV,ITPART,ITCUSTREQ,ITPSNUMBER,ITPSITEM,PARTREF,PARTNUM" & vbCrLf _
          & "FROM SohdTable,SoitTable,PartTable" & vbCrLf _
          & "WHERE (SONUMBER=ITSO AND SOCUST='" _
          & Compress(cmbCst) & "' AND SOPO LIKE '" & Trim(txtCpo) & "%' " _
          & "AND ITPSNUMBER <> '' AND ITPSSHIPPED=0 AND PARTREF=ITPART)"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoItm, ES_FORWARD)
   If bSqlRows Then
      With RdoItm
         Do Until .EOF
            iRow = iRow + 1
            If iRow < 11 Then
               lblSon(iRow) = Trim(!SOTYPE) & Format(!SoNumber, SO_NUM_FORMAT)
               lblSitm(iRow) = Format(!ITNUMBER, "##0")
               lblItRev(iRow) = "" & Trim(!ITREV)
               lblPsNum(iRow) = "" & Trim(!ITPSNUMBER)
               lblPsItm(iRow) = "" & Trim(!ITPSITEM)
               lblPart(iRow) = "" & Trim(!PartNum)
               If Not IsNull(!ITCUSTREQ) Then _
                             txtReq(iRow) = Format(!ITCUSTREQ, "mm/dd/yy")
               txtReq(iRow).Enabled = True
            End If
            sSales(iRow, 0) = str$(!SoNumber)
            sSales(iRow, 1) = "" & Trim(!SOTYPE) & Format$(!SoNumber, SO_NUM_FORMAT)
            sSales(iRow, 2) = str$(!ITNUMBER)
            sSales(iRow, 3) = "" & Trim(!ITREV)
            sSales(iRow, 4) = "" & Trim(!ITPSNUMBER)
            sSales(iRow, 5) = str$(!ITPSITEM)
            sSales(iRow, 6) = "" & Trim(!PartNum)
            If Not IsNull(!ITCUSTREQ) Then
               sSales(iRow, 7) = Format$(!ITCUSTREQ, "mm/dd/yy")
            Else
               sSales(iRow, 7) = ""
            End If
            If iRow > 799 Then Exit Do
            .MoveNext
         Loop
         ClearResultSet RdoItm
      End With
   Else
      MsgBox "No Matching Items Were Found.", _
         vbInformation, Caption
   End If
   Set RdoItm = Nothing
   iTotalItems = iRow
   iIndex = 0
   lblLst = 0
   If iTotalItems > 0 Then
      iRow = iTotalItems Mod 10
      iLastIndex = iTotalItems
      txtReq(1).SetFocus
      If iRow > 0 Then
         lblLst = Int(iTotalItems / 10) + 1
      Else
         lblLst = Int(iTotalItems / 10)
      End If
      iCurrPage = 1
   End If
   If Val(lblLst) > 1 Then
      cmdUp.Picture = Enup
      cmdDn.Picture = Endn
      cmdUp.Enabled = True
      cmdDn.Enabled = True
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "selectsaleso"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtNme_Change()
   If Left(txtNme, 6) = "*** Cu" Then
      txtNme.ForeColor = ES_RED
      cmdSel.Enabled = False
   Else
      txtNme.ForeColor = vbBlack
      cmdSel.Enabled = True
   End If
   
End Sub

Private Sub txtReq_DropDown(Index As Integer)
   ShowCalendar Me
   
End Sub



Private Sub CloseBoxes(Optional CloseButtons As Byte)
   Dim b As Byte
   For b = 1 To 10
      lblSon(b) = ""
      lblSitm(b) = ""
      lblItRev(b) = ""
      lblPsNum(b) = ""
      lblPsItm(b) = ""
      lblPart(b) = ""
      txtReq(b) = ""
      txtReq(b) = ""
      txtReq(b).Enabled = False
   Next
   If CloseButtons = 0 Then
      cmdUp.Picture = Dsup
      cmdDn.Picture = Dsdn
      cmdUp.Enabled = False
      cmdDn.Enabled = False
   End If
   
End Sub

Private Sub txtReq_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   KeyCheck KeyCode
   
End Sub

Private Sub txtReq_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyDate KeyAscii
   
End Sub

Private Sub txtReq_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub


Private Sub txtReq_LostFocus(Index As Integer)
   txtReq(Index) = CheckDate(txtReq(Index))
   If txtReq(Index) <> sSales(iIndex + Index, 7) Then
      sSales(iIndex + Index, 7) = txtReq(Index)
      sSql = "UPDATE SoitTable SET ITCUSTREQ='" & txtReq(Index) & "' " _
             & "WHERE (ITSO=" & sSales(iIndex + Index, 0) & " AND " _
             & "ITNUMBER=" & sSales(iIndex + Index, 2) & " AND " _
             & "ITREV='" & sSales(iIndex + Index, 3) & "')"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
   End If
   
End Sub



Public Sub GetNextGroup()
   Dim iList As Integer
   Dim iRow As Integer
   CloseBoxes 1
   On Error Resume Next
   For iList = iIndex To iTotalItems
      iRow = iRow + 1
      If iRow > 10 Then Exit For
      If sSales(iIndex + iRow, 1) = "" Then Exit For
      lblSon(iRow) = sSales(iIndex + iRow, 1)
      lblSitm(iRow) = sSales(iIndex + iRow, 2)
      lblItRev(iRow) = sSales(iIndex + iRow, 3)
      lblPsNum(iRow) = sSales(iIndex + iRow, 4)
      lblPsItm(iRow) = sSales(iIndex + iRow, 5)
      lblPart(iRow) = sSales(iIndex + iRow, 6)
      txtReq(iRow) = sSales(iIndex + iRow, 7)
      txtReq(iRow).Enabled = True
   Next
   If lblPge = lblLst Then
      cmdDn.Picture = Dsdn
      cmdDn.Enabled = False
   Else
      cmdDn.Picture = Endn
      cmdDn.Enabled = True
   End If
   txtReq(1).SetFocus
   
End Sub

Public Sub FillPurchaseOrders()
   txtCpo.Clear
   sSql = "SELECT DISTINCT SOPO FROM SohdTable WHERE (SOCUST='" _
          & Compress(cmbCst) & "' AND SOPO<>'')"
   LoadComboBox txtCpo, -1
   If txtCpo.ListCount > 0 Then txtCpo = txtCpo.List(0)
   
End Sub
