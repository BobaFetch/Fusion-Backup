VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FainFAe01b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "First Artical Inspection Detail"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdUp 
      DisabledPicture =   "FainFAe01b.frx":0000
      DownPicture     =   "FainFAe01b.frx":04F2
      Enabled         =   0   'False
      Height          =   372
      Left            =   7800
      MaskColor       =   &H00000000&
      Picture         =   "FainFAe01b.frx":09E4
      Style           =   1  'Graphical
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   5760
      Width           =   400
   End
   Begin VB.CommandButton cmdDn 
      DisabledPicture =   "FainFAe01b.frx":0ED6
      DownPicture     =   "FainFAe01b.frx":13C8
      Enabled         =   0   'False
      Height          =   372
      Left            =   7800
      MaskColor       =   &H00000000&
      Picture         =   "FainFAe01b.frx":18BA
      Style           =   1  'Graphical
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   6144
      Width           =   400
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "FainFAe01b.frx":1DAC
      Style           =   1  'Graphical
      TabIndex        =   83
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdCmt 
      Height          =   350
      Index           =   6
      Left            =   6480
      Picture         =   "FainFAe01b.frx":255A
      Style           =   1  'Graphical
      TabIndex        =   82
      ToolTipText     =   "Additional Inspection Description/Comments"
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   350
   End
   Begin VB.CommandButton cmdCmt 
      Height          =   350
      Index           =   5
      Left            =   6480
      Picture         =   "FainFAe01b.frx":2B5C
      Style           =   1  'Graphical
      TabIndex        =   81
      ToolTipText     =   "Additional Inspection Description/Comments"
      Top             =   4560
      UseMaskColor    =   -1  'True
      Width           =   350
   End
   Begin VB.CommandButton cmdCmt 
      Height          =   350
      Index           =   4
      Left            =   6480
      Picture         =   "FainFAe01b.frx":315E
      Style           =   1  'Graphical
      TabIndex        =   80
      ToolTipText     =   "Additional Inspection Description/Comments"
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   350
   End
   Begin VB.CommandButton cmdCmt 
      Height          =   350
      Index           =   3
      Left            =   6480
      Picture         =   "FainFAe01b.frx":3760
      Style           =   1  'Graphical
      TabIndex        =   79
      ToolTipText     =   "Additional Inspection Description/Comments"
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   350
   End
   Begin VB.CommandButton cmdCmt 
      Height          =   350
      Index           =   2
      Left            =   6480
      Picture         =   "FainFAe01b.frx":3D62
      Style           =   1  'Graphical
      TabIndex        =   78
      ToolTipText     =   "Additional Inspection Description/Comments"
      Top             =   2040
      UseMaskColor    =   -1  'True
      Width           =   350
   End
   Begin VB.CommandButton cmdCmt 
      Height          =   350
      Index           =   1
      Left            =   6480
      Picture         =   "FainFAe01b.frx":4364
      Style           =   1  'Graphical
      TabIndex        =   77
      ToolTipText     =   "Additional Inspection Description/Comments"
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   350
   End
   Begin VB.CheckBox optSeq 
      Alignment       =   1  'Right Justify
      Caption         =   "Resequence On Exit"
      Height          =   195
      Left            =   3960
      TabIndex        =   76
      TabStop         =   0   'False
      ToolTipText     =   "Resequences When The Data Is Saved On Exit (Recommended)"
      Top             =   480
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   315
      Left            =   6360
      TabIndex        =   75
      ToolTipText     =   "Adds An Item To The Report"
      Top             =   480
      Width           =   875
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Resort"
      Height          =   315
      Left            =   7320
      TabIndex        =   74
      ToolTipText     =   "Resorts And Saves List"
      Top             =   480
      Width           =   875
   End
   Begin VB.CheckBox optAct 
      Alignment       =   1  'Right Justify
      Caption         =   "Accepted "
      Height          =   255
      Index           =   6
      Left            =   7080
      TabIndex        =   53
      ToolTipText     =   "Accept This Feature"
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CheckBox optAct 
      Alignment       =   1  'Right Justify
      Caption         =   "Accepted "
      Height          =   255
      Index           =   5
      Left            =   7080
      TabIndex        =   44
      ToolTipText     =   "Accept This Feature"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CheckBox optAct 
      Alignment       =   1  'Right Justify
      Caption         =   "Accepted "
      Height          =   255
      Index           =   4
      Left            =   7080
      TabIndex        =   35
      ToolTipText     =   "Accept This Feature"
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CheckBox optAct 
      Alignment       =   1  'Right Justify
      Caption         =   "Accepted "
      Height          =   255
      Index           =   3
      Left            =   7080
      TabIndex        =   26
      ToolTipText     =   "Accept This Feature"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CheckBox optAct 
      Alignment       =   1  'Right Justify
      Caption         =   "Accepted "
      Height          =   255
      Index           =   2
      Left            =   7080
      TabIndex        =   17
      ToolTipText     =   "Accept This Feature"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CheckBox optAct 
      Alignment       =   1  'Right Justify
      Caption         =   "Accepted "
      Height          =   255
      Index           =   1
      Left            =   7080
      TabIndex        =   8
      ToolTipText     =   "Accept This Feature"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtInsp 
      Height          =   285
      Index           =   1
      Left            =   2280
      TabIndex        =   6
      Tag             =   "2"
      ToolTipText     =   "Actual Inspected"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Index           =   6
      Left            =   1200
      TabIndex        =   47
      Tag             =   "1"
      ToolTipText     =   "Location On The Drawing"
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox txtSeq 
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   45
      Tag             =   "1"
      ToolTipText     =   "Sequence (Reorder)"
      Top             =   5040
      Width           =   375
   End
   Begin VB.TextBox txtOpn 
      Height          =   285
      Index           =   6
      Left            =   600
      TabIndex        =   46
      Tag             =   "1"
      ToolTipText     =   "MO Operation"
      Top             =   5040
      Width           =   495
   End
   Begin VB.TextBox txtDim 
      Height          =   285
      Index           =   6
      Left            =   2280
      TabIndex        =   48
      Tag             =   "2"
      ToolTipText     =   "The Feature (Up To 40 Chars)"
      Top             =   5040
      Width           =   4695
   End
   Begin VB.TextBox txtTol 
      Height          =   285
      Index           =   6
      Left            =   600
      TabIndex        =   50
      Tag             =   "2"
      ToolTipText     =   "Dim Tolerance"
      Top             =   5400
      Width           =   1575
   End
   Begin VB.TextBox txtInsp 
      Height          =   285
      Index           =   6
      Left            =   2280
      TabIndex        =   51
      Tag             =   "2"
      ToolTipText     =   "Actual Inspected"
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete"
      Height          =   285
      Index           =   6
      Left            =   7320
      TabIndex        =   49
      TabStop         =   0   'False
      ToolTipText     =   "Deletes The Current Feature"
      Top             =   5040
      Width           =   855
   End
   Begin VB.ComboBox cmbMet 
      Height          =   315
      Index           =   6
      Left            =   3960
      Sorted          =   -1  'True
      TabIndex        =   52
      Tag             =   "2"
      ToolTipText     =   "Method Of Inspection - Records For Future Use"
      Top             =   5400
      Width           =   2415
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Index           =   5
      Left            =   1200
      TabIndex        =   38
      Tag             =   "1"
      ToolTipText     =   "Location On The Drawing"
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox txtSeq 
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   36
      Tag             =   "1"
      ToolTipText     =   "Sequence (Reorder)"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox txtOpn 
      Height          =   285
      Index           =   5
      Left            =   600
      TabIndex        =   37
      Tag             =   "1"
      ToolTipText     =   "MO Operation"
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox txtDim 
      Height          =   285
      Index           =   5
      Left            =   2280
      TabIndex        =   39
      Tag             =   "2"
      ToolTipText     =   "The Feature (Up To 40 Chars)"
      Top             =   4200
      Width           =   4695
   End
   Begin VB.TextBox txtTol 
      Height          =   285
      Index           =   5
      Left            =   600
      TabIndex        =   41
      Tag             =   "2"
      ToolTipText     =   "Dim Tolerance"
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox txtInsp 
      Height          =   285
      Index           =   5
      Left            =   2280
      TabIndex        =   42
      Tag             =   "2"
      ToolTipText     =   "Actual Inspected"
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete"
      Height          =   285
      Index           =   5
      Left            =   7320
      TabIndex        =   40
      TabStop         =   0   'False
      ToolTipText     =   "Deletes The Current Feature"
      Top             =   4200
      Width           =   855
   End
   Begin VB.ComboBox cmbMet 
      Height          =   315
      Index           =   5
      Left            =   3960
      Sorted          =   -1  'True
      TabIndex        =   43
      Tag             =   "2"
      ToolTipText     =   "Method Of Inspection - Records For Future Use"
      Top             =   4560
      Width           =   2415
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Index           =   4
      Left            =   1200
      TabIndex        =   29
      Tag             =   "1"
      ToolTipText     =   "Location On The Drawing"
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtSeq 
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   27
      Tag             =   "1"
      ToolTipText     =   "Sequence (Reorder)"
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox txtOpn 
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   28
      Tag             =   "1"
      ToolTipText     =   "MO Operation"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtDim 
      Height          =   285
      Index           =   4
      Left            =   2280
      TabIndex        =   30
      Tag             =   "2"
      ToolTipText     =   "The Feature (Up To 40 Chars)"
      Top             =   3360
      Width           =   4695
   End
   Begin VB.TextBox txtTol 
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   32
      Tag             =   "2"
      ToolTipText     =   "Dim Tolerance"
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox txtInsp 
      Height          =   285
      Index           =   4
      Left            =   2280
      TabIndex        =   33
      Tag             =   "2"
      ToolTipText     =   "Actual Inspected"
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete"
      Height          =   285
      Index           =   4
      Left            =   7320
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Deletes The Current Feature"
      Top             =   3360
      Width           =   855
   End
   Begin VB.ComboBox cmbMet 
      Height          =   315
      Index           =   4
      Left            =   3960
      Sorted          =   -1  'True
      TabIndex        =   34
      Tag             =   "2"
      ToolTipText     =   "Method Of Inspection - Records For Future Use"
      Top             =   3720
      Width           =   2415
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Index           =   3
      Left            =   1200
      TabIndex        =   20
      Tag             =   "1"
      ToolTipText     =   "Location On The Drawing"
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtSeq 
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   18
      Tag             =   "1"
      ToolTipText     =   "Sequence (Reorder)"
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox txtOpn 
      Height          =   285
      Index           =   3
      Left            =   600
      TabIndex        =   19
      Tag             =   "1"
      ToolTipText     =   "MO Operation"
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox txtDim 
      Height          =   285
      Index           =   3
      Left            =   2280
      TabIndex        =   21
      Tag             =   "2"
      ToolTipText     =   "The Feature (Up To 40 Chars)"
      Top             =   2520
      Width           =   4695
   End
   Begin VB.TextBox txtTol 
      Height          =   285
      Index           =   3
      Left            =   600
      TabIndex        =   23
      Tag             =   "2"
      ToolTipText     =   "Dim Tolerance"
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox txtInsp 
      Height          =   285
      Index           =   3
      Left            =   2280
      TabIndex        =   24
      Tag             =   "2"
      ToolTipText     =   "Actual Inspected"
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete"
      Height          =   285
      Index           =   3
      Left            =   7320
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Deletes The Current Feature"
      Top             =   2520
      Width           =   855
   End
   Begin VB.ComboBox cmbMet 
      Height          =   315
      Index           =   3
      Left            =   3960
      Sorted          =   -1  'True
      TabIndex        =   25
      Tag             =   "2"
      ToolTipText     =   "Method Of Inspection - Records For Future Use"
      Top             =   2880
      Width           =   2415
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Index           =   2
      Left            =   1200
      TabIndex        =   11
      Tag             =   "1"
      ToolTipText     =   "Location On The Drawing"
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtSeq 
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Tag             =   "1"
      ToolTipText     =   "Sequence (Reorder)"
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox txtOpn 
      Height          =   285
      Index           =   2
      Left            =   600
      TabIndex        =   10
      Tag             =   "1"
      ToolTipText     =   "MO Operation"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtDim 
      Height          =   285
      Index           =   2
      Left            =   2280
      TabIndex        =   12
      Tag             =   "2"
      ToolTipText     =   "The Feature (Up To 40 Chars)"
      Top             =   1680
      Width           =   4695
   End
   Begin VB.TextBox txtTol 
      Height          =   285
      Index           =   2
      Left            =   600
      TabIndex        =   14
      Tag             =   "2"
      ToolTipText     =   "Dim Tolerance"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox txtInsp 
      Height          =   285
      Index           =   2
      Left            =   2280
      TabIndex        =   15
      Tag             =   "2"
      ToolTipText     =   "Actual Inspected"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete"
      Height          =   285
      Index           =   2
      Left            =   7320
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Deletes The Current Feature"
      Top             =   1680
      Width           =   855
   End
   Begin VB.ComboBox cmbMet 
      Height          =   315
      Index           =   2
      Left            =   3960
      Sorted          =   -1  'True
      TabIndex        =   16
      Tag             =   "2"
      ToolTipText     =   "Method Of Inspection - Records For Future Use"
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "Location On The Drawing"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtSeq 
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Sequence (Reorder)"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox txtOpn 
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "MO Operation"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtDim 
      Height          =   285
      Index           =   1
      Left            =   2280
      TabIndex        =   3
      Tag             =   "2"
      ToolTipText     =   "The Feature (Up To 40 Chars)"
      Top             =   960
      Width           =   4695
   End
   Begin VB.TextBox txtTol 
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   5
      Tag             =   "2"
      ToolTipText     =   "Dim Tolerance"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete"
      Height          =   285
      Index           =   1
      Left            =   7320
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Deletes The Current Feature"
      Top             =   960
      Width           =   855
   End
   Begin VB.ComboBox cmbMet 
      Height          =   315
      Index           =   1
      Left            =   3960
      Sorted          =   -1  'True
      TabIndex        =   7
      Tag             =   "2"
      ToolTipText     =   "Method Of Inspection - Records For Future Use"
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox txtSeq 
      Height          =   285
      Index           =   0
      Left            =   -600
      TabIndex        =   54
      Tag             =   "1"
      ToolTipText     =   "Sequence (Reorder)"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   7320
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   1680
      Top             =   6600
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6660
      FormDesignWidth =   8280
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   2880
      TabIndex        =   86
      Top             =   6240
      Width           =   4092
      _ExtentX        =   7223
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Image Dsup 
      Height          =   300
      Left            =   960
      Picture         =   "FainFAe01b.frx":4966
      Top             =   6360
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Enup 
      Height          =   300
      Left            =   1440
      Picture         =   "FainFAe01b.frx":4E58
      Top             =   6360
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Endn 
      Height          =   300
      Left            =   1200
      Picture         =   "FainFAe01b.frx":534A
      Top             =   6360
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Dsdn 
      Height          =   300
      Left            =   720
      Picture         =   "FainFAe01b.frx":583C
      Top             =   6360
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   1200
      TabIndex        =   73
      Top             =   6240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblPges 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7080
      TabIndex        =   72
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label lblPge 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6120
      TabIndex        =   71
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Of"
      Height          =   285
      Index           =   11
      Left            =   6720
      TabIndex        =   70
      Top             =   5880
      Width           =   705
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Page"
      Height          =   285
      Index           =   10
      Left            =   5520
      TabIndex        =   69
      Top             =   5880
      Width           =   705
   End
   Begin VB.Label lblRows 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1200
      TabIndex        =   68
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rows Found:"
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   67
      Top             =   5880
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Method Of Inspection                                "
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
      Left            =   3960
      TabIndex        =   66
      Top             =   720
      Width           =   2985
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inspected                   "
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
      Left            =   2280
      TabIndex        =   65
      Top             =   720
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tolerance                  "
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
      Left            =   600
      TabIndex        =   64
      Top             =   720
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Feature description/"
      Height          =   285
      Index           =   5
      Left            =   2280
      TabIndex        =   63
      Top             =   480
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dwg Loc/"
      Height          =   285
      Index           =   4
      Left            =   1200
      TabIndex        =   62
      Top             =   480
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Op No "
      Height          =   285
      Index           =   3
      Left            =   600
      TabIndex        =   61
      Top             =   480
      Width           =   465
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seq  "
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   60
      Top             =   480
      Width           =   465
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Report Number"
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   59
      Top             =   120
      Width           =   1545
   End
   Begin VB.Label txtPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   58
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Report Revision"
      Height          =   285
      Index           =   0
      Left            =   4680
      TabIndex        =   57
      Top             =   120
      Width           =   1545
   End
   Begin VB.Label txtRev 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6000
      TabIndex        =   56
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "FainFAe01b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'new May 2003
Option Explicit

Dim bCanceled As Byte
Dim bDataChg As Byte
Dim bOnLoad As Byte
Dim bPages As Byte

Dim iCurRow As Integer
Dim iIndex As Integer
Dim iItems As Integer
Dim iCurrPage As Integer

Dim sRptNo As String
Dim sRevNo As String
Dim vFaItems(300, 9) As Variant

Dim StartTime As Date

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub GetOptions()
   Dim sOptions As String
   'Get By Menu Option
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiQual", "faitm", sOptions)
   If Len(sOptions) > 0 Then optSeq.Value = Val(sOptions) _
          Else optSeq.Value = vbChecked
   
End Sub


Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim(str(optSeq.Value))
   SaveSetting "Esi2000", "EsiQual", "faitm", Trim(sOptions)
   
End Sub


Private Sub cmbMet_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub

Private Sub cmbMet_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyCheck KeyAscii
   
End Sub

Private Sub cmbMet_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then
      cmdUp_Click
   End If
   If KeyCode = vbKeyPageDown Then
      cmdDn_Click
   
   End If
   
End Sub

Private Sub cmbMet_Validate(Index As Integer, Cancel As Boolean)
   cmbMet(Index) = CheckLen(cmbMet(Index), 20)
   cmbMet(Index) = StrCase(cmbMet(Index))
   vFaItems(iIndex + Index, 7) = cmbMet(Index)
   
End Sub

Private Sub cmdAdd_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   sMsg = "This Funtion Adds An Item To The End Of The " & vbCr _
          & "Current Row But May The Sequence May Be Changed." & vbCr _
          & "Do You Wish To Continue?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then AddItem Else CancelTrans
   
   
End Sub

Private Sub cmdCan_Click()
   Dim iList As Integer
   For iList = 1 To 6
      If vFaItems(iIndex + iList, 8) <> optAct(iList).Value Then bDataChg = 1
      vFaItems(iIndex + iList, 8) = optAct(iList).Value
   Next
   bCanceled = 1
   If bDataChg = 1 Then UpDateItems 1 Else Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCanceled = 1
   
End Sub


Private Sub cmdCmt_Click(Index As Integer)
   FainFAe02c.lblItem = vFaItems(Index + iIndex, 0)
   FainFAe02c.txtPrt = txtPrt
   FainFAe02c.txtRev = txtRev
   FainFAe02c.Show
   
End Sub

Private Sub cmdDel_Click(Index As Integer)
   Dim bResponse As Byte
   Dim sMsg As String
   sMsg = "This Funtion Deletes The Current Item." & vbCr _
          & "Do You Wish To Continue?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   If bResponse = vbYes Then
      sSql = ""
      sSql = "DELETE FROM FaitTable WHERE " _
             & "FA_ITNUMBER='" & sRptNo & "' AND " _
             & "FA_ITREVISION='" & sRevNo & "' AND " _
             & "FA_ITFEATURENUM=" & vFaItems(iIndex + Index, 0) _
             & " "
      clsADOCon.ExecuteSql sSql
      If clsADOCon.ADOErrNum = 0 Then
         SysMsg "Item Deleted.", True
         FillItems
      Else
         MsgBox "Couldn't Delete The Item.", _
            vbInformation, Caption
      End If
   Else
      CancelTrans
   End If
End Sub

Private Sub cmdDn_Click()
   iCurrPage = iCurrPage + 1
   If iCurrPage > bPages Then iCurrPage = bPages
   GetNextGroup
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 6202
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdUp_Click()
   iCurrPage = iCurrPage - 1
   If iCurrPage < 1 Then iCurrPage = 1
   GetNextGroup
   
End Sub

Private Sub cmdUpd_Click()
   UpDateItems
   
End Sub

Private Sub Form_Activate()
   If bOnLoad = 1 Then
      MouseCursor 13
      FainFAe02b.Enabled = False
      txtPrt = FainFAe02b.txtPrt
      txtRev = FainFAe02b.txtRev
      sRptNo = Compress(txtPrt)
      sRevNo = Trim(txtRev)
      iCurRow = 1
      FillMethods
      FillItems
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Deactivate()
   'Unload Me
   
End Sub

Private Sub Form_Load()
   Move 200, 500
   FormatControls
   bOnLoad = 1
   GetOptions
End Sub


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   cmdCan.ToolTipText = "Save Changes And Exit"
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   FainFAe02b.Enabled = True
   
End Sub



Private Sub FillMethods()
   Dim RdoMet As ADODB.Recordset
   Dim b As Byte
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT FA_ITMETHOD FROM FaitTable WHERE " _
          & "FA_ITMETHOD<>''"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoMet, ES_FORWARD)
   If bSqlRows Then
      With RdoMet
         Do Until .EOF
            For b = 1 To 6
               AddComboStr cmbMet(b).hwnd, "" & Trim(!FA_ITMETHOD)
            Next
            .MoveNext
         Loop
         ClearResultSet RdoMet
      End With
   End If
   Set RdoMet = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillmethods"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillItems()
   Dim RdoRpt As ADODB.Recordset
   Dim iList As Integer
   Dim a As Integer
   Dim cPages As Currency
   
   Erase vFaItems
   CloseBoxes
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM FaitTable WHERE (FA_ITNUMBER='" _
          & sRptNo & "' AND FA_ITREVISION='" & sRevNo & "') " _
          & "ORDER BY FA_ITSEQUENCE,FA_ITFEATURENUM"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpt, ES_FORWARD)
   If bSqlRows Then
      a = a + 5
      If a > 95 Then a = 95
      lblStatus = "Loading"
      prg1.Value = a
      lblStatus.Visible = True
      prg1.Visible = True
      With RdoRpt
         Do Until .EOF
            iList = iList + 1
            If iList < 7 Then
               txtSeq(iList) = !FA_ITSEQUENCE
               If Val(!FA_ITOPNO) > 0 Then txtOpn(iList) = Format(!FA_ITOPNO, "000")
               txtLoc(iList) = "" & Trim(!FA_ITDRAWINGLOC)
               If !FA_ITOPNO > 0 Then txtOpn(iList) = Format(!FA_ITOPNO, "000")
               txtDim(iList) = "" & Trim(!FA_ITDIMENSION)
               txtTol(iList) = "" & Trim(!FA_ITDIMTOL)
               txtInsp(iList) = "" & Trim(!FA_ITDIMACT)
               cmbMet(iList) = "" & Trim(!FA_ITMETHOD)
               If Trim(!FA_ITACCEPTED) = "Y" Then
                  optAct(iList).Value = 1
               Else
                  optAct(iList).Value = 0
               End If
               txtSeq(iList).Visible = True
               txtOpn(iList).Visible = True
               txtLoc(iList).Visible = True
               txtDim(iList).Visible = True
               txtTol(iList).Visible = True
               txtInsp(iList).Visible = True
               cmbMet(iList).Visible = True
               cmdDel(iList).Visible = True
               optAct(iList).Visible = True
            End If
            vFaItems(iList, 0) = !FA_ITFEATURENUM
            vFaItems(iList, 1) = !FA_ITSEQUENCE
            If !FA_ITOPNO > 0 Then
               vFaItems(iList, 2) = Format(!FA_ITOPNO, "000")
            Else
               vFaItems(iList, 2) = ""
            End If
            vFaItems(iList, 3) = "" & Trim(!FA_ITDRAWINGLOC)
            vFaItems(iList, 4) = "" & Trim(!FA_ITDIMENSION)
            vFaItems(iList, 5) = "" & Trim(!FA_ITDIMTOL)
            vFaItems(iList, 6) = "" & Trim(!FA_ITDIMACT)
            vFaItems(iList, 7) = "" & Trim(!FA_ITMETHOD)
            If "" & Trim(!FA_ITACCEPTED) = "Y" Then
               vFaItems(iList, 8) = "1"
            Else
               vFaItems(iList, 8) = "0"
            End If
            .MoveNext
         Loop
         ClearResultSet RdoRpt
      End With
      iIndex = 0
      iItems = iList
      lblRows = iItems
      iCurrPage = 1
      lblPge = iCurrPage
      cPages = iItems / 6
      lblPges = Format(cPages, "#0.0")
      If Val(Right(lblPges, 1)) > 0 Then cPages = cPages + 1
      bPages = cPages
      lblPges = bPages
      bDataChg = 0
      cmdUp.Enabled = True
      cmdDn.Enabled = True
      cmdUp.Picture = Enup
      cmdDn.Picture = Endn
      On Error Resume Next
      MouseCursor 0
      lblStatus.Visible = False
      prg1.Visible = False
      txtSeq(1).SetFocus
   End If
   
   Set RdoRpt = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetNextGroup()
   Dim a As Integer
   Dim iList As Integer
   For iList = 1 To 6
      If vFaItems(iIndex + iList, 8) <> str(optAct(iList).Value) Then bDataChg = 1
      vFaItems(iIndex + iList, 8) = str(optAct(iList).Value)
   Next
   
   lblPge = iCurrPage
   CloseBoxes
   iIndex = (iCurrPage - 1) * 6
   On Error Resume Next
   For iList = iIndex To iItems
      If iList + 1 > iItems Then Exit For
      a = a + 1
      If a > 6 Then Exit For
      txtSeq(a).Visible = True
      txtOpn(a).Visible = True
      txtLoc(a).Visible = True
      txtDim(a).Visible = True
      txtTol(a).Visible = True
      txtInsp(a).Visible = True
      cmbMet(a).Visible = True
      cmdDel(a).Visible = True
      optAct(a).Visible = True
      
      txtSeq(a) = vFaItems(iList + 1, 1)
      txtOpn(a) = vFaItems(iList + 1, 2)
      txtLoc(a) = vFaItems(iList + 1, 3)
      txtDim(a) = vFaItems(iList + 1, 4)
      txtTol(a) = vFaItems(iList + 1, 5)
      txtInsp(a) = vFaItems(iList + 1, 6)
      cmbMet(a) = vFaItems(iList + 1, 7)
      optAct(a).Value = Val(vFaItems(iList + 1, 8))
   Next
   On Error Resume Next
   txtSeq(1).SetFocus
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set FainFAe01b = Nothing
   
End Sub

Private Sub optAct_Click(Index As Integer)
   '  If vFaItems(iIndex + Index, 8) <> Str(optAct(Index).Value) Then bDataChg = 1
   '  vFaItems(iIndex + Index, 8) = Str(optAct(Index).Value)
   
End Sub

Private Sub optAct_GotFocus(Index As Integer)
   iCurRow = iIndex + Index
   
End Sub


Private Sub optAct_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optAct_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub

Private Sub optSeq_Click()
   If Not bOnLoad Then bDataChg = 1
   
End Sub

Private Sub txtDim_GotFocus(Index As Integer)
   SelectFormat Me
   iCurRow = iIndex + Index
   
End Sub


Private Sub txtDim_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtDim_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyCheck KeyAscii
   
End Sub

Private Sub txtDim_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub

Private Sub txtDim_Validate(Index As Integer, Cancel As Boolean)
   txtDim(Index) = CheckLen(txtDim(Index), 40)
   txtDim(Index) = StrCase(txtDim(Index))
   If vFaItems(iIndex + Index, 4) <> txtDim(Index) Then bDataChg = 1
   vFaItems(iIndex + Index, 4) = txtDim(Index)
   
End Sub

Private Sub txtInsp_GotFocus(Index As Integer)
   SelectFormat Me
   iCurRow = iIndex + Index
   
End Sub


Private Sub txtInsp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtInsp_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyCheck KeyAscii
   
End Sub

Private Sub txtInsp_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub

Private Sub txtInsp_Validate(Index As Integer, Cancel As Boolean)
   txtInsp(Index) = CheckLen(txtInsp(Index), 12)
   If vFaItems(iIndex + Index, 6) <> txtInsp(Index) Then bDataChg = 1
   vFaItems(iIndex + Index, 6) = txtInsp(Index)
   
End Sub

Private Sub txtLoc_GotFocus(Index As Integer)
   SelectFormat Me
   iCurRow = iIndex + Index
   
End Sub


Private Sub txtLoc_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtLoc_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyCase KeyAscii
   
End Sub


Private Sub txtLoc_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub

Private Sub txtLoc_Validate(Index As Integer, Cancel As Boolean)
   txtLoc(Index) = CheckLen(txtLoc(Index), 8)
   If vFaItems(iIndex + Index, 3) <> txtLoc(Index) Then bDataChg = 1
   vFaItems(iIndex + Index, 3) = txtLoc(Index)
   
End Sub

Private Sub txtOpn_GotFocus(Index As Integer)
   SelectFormat Me
   iCurRow = iIndex + Index
   
End Sub


Private Sub txtOpn_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtOpn_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub

Private Sub txtOpn_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub

Private Sub txtOpn_Validate(Index As Integer, Cancel As Boolean)
   If Val(txtOpn(Index)) > 0 Then
      txtOpn(Index) = CheckLen(txtOpn(Index), 3)
      txtOpn(Index) = Format(Abs(Val(txtOpn(Index))), "000")
   End If
   If Val(vFaItems(iIndex + Index, 2)) <> Val(txtOpn(Index)) Then bDataChg = 1
   vFaItems(iIndex + Index, 2) = txtOpn(Index)
   
   
End Sub

Private Sub txtSeq_GotFocus(Index As Integer)
   SelectFormat Me
   iCurRow = iIndex + Index
   
End Sub


Private Sub txtSeq_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtSeq_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub

Private Sub txtSeq_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then
      'MsgBox (vFaItems(1, 1))
      cmdUp_Click
   End If
   If KeyCode = vbKeyPageDown Then
      'MsgBox (vFaItems(1, 1))
      cmdDn_Click
   End If
   
End Sub

Private Sub txtSeq_Validate(Index As Integer, Cancel As Boolean)
   If Val(txtSeq(Index)) = 0 Then txtSeq(Index) = vFaItems(iIndex + Index, 1)
   txtSeq(Index) = CheckLen(txtSeq(Index), 3)
   txtSeq(Index) = Format(Abs(Val(txtSeq(Index))), "##0")
   If vFaItems(iIndex + Index, 1) <> txtSeq(Index) Then bDataChg = 1
   vFaItems(iIndex + Index, 1) = txtSeq(Index)
   
End Sub

Private Sub txtTol_GotFocus(Index As Integer)
   SelectFormat Me
   iCurRow = iIndex + Index
   
End Sub


Private Sub txtTol_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtTol_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyCheck KeyAscii
   
End Sub


Private Sub txtTol_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub



Private Sub CloseBoxes()
   On Error Resume Next
   Dim iList As Integer
   For iList = 1 To 5
      
      txtSeq(iList).Visible = False
      txtOpn(iList).Visible = False
      txtLoc(iList).Visible = False
      txtDim(iList).Visible = False
      txtTol(iList).Visible = False
      txtInsp(iList).Visible = False
      cmbMet(iList).Visible = False
      cmdDel(iList).Visible = False
      optAct(iList).Visible = False
   
      txtSeq(iList) = ""
      txtOpn(iList) = ""
      txtLoc(iList) = ""
      txtDim(iList) = ""
      txtTol(iList) = ""
      txtInsp(iList) = ""
      cmbMet(iList) = ""
      cmdDel(iList).Visible = False
      optAct(iList).Value = vbUnchecked
   
   Next
   
   txtSeq(iList).Visible = False
   txtOpn(iList).Visible = False
   txtLoc(iList).Visible = False
   txtDim(iList).Visible = False
   txtTol(iList).Visible = False
   txtInsp(iList).Visible = False
   cmbMet(iList).Visible = False
   cmdDel(iList).Visible = False
   optAct(iList).Visible = False
   
   txtSeq(iList) = ""
   txtOpn(iList) = ""
   txtLoc(iList) = ""
   txtDim(iList) = ""
   txtTol(iList) = ""
   txtInsp(iList) = ""
   cmbMet(iList) = ""
   cmdDel(iList).Visible = False
   ' optAct(iList).Value = vbUnchecked
   
End Sub

Private Sub txtTol_Validate(Index As Integer, Cancel As Boolean)
   txtTol(Index) = CheckLen(txtTol(Index), 12)
   If vFaItems(iIndex + Index, 5) <> txtTol(Index) Then bDataChg = 1
   vFaItems(iIndex + Index, 5) = txtTol(Index)
   
End Sub



Private Sub UpDateItems(Optional bUnload As Byte)
   Dim RdoUpd As ADODB.Recordset
   Dim iList As Integer
   Dim a As Integer
   Dim b As Integer
   
   b = Val(lblRows)
   Dim sAcct(300) As String
   
   MouseCursor 13
   a = a + 5
   If a > 95 Then a = 95
   lblStatus = "Updating"
   prg1.Value = a
   lblStatus.Visible = True
   prg1.Visible = True
   On Error Resume Next
   'Items
   For iList = 1 To b
      a = a + 5
      If a > 65 Then a = 65
      prg1.Value = a
      If Val(vFaItems(iList, 8)) = 1 Then sAcct(iList) = "Y" _
             Else sAcct(iList) = "N"
      sSql = "SELECT * FROM FaitTable WHERE (FA_ITNUMBER='" _
             & sRptNo & "' AND FA_ITREVISION='" & sRevNo & "' " _
             & "AND FA_ITFEATURENUM=" & vFaItems(iList, 0) & ")"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoUpd, ES_DYNAMIC)
      If bSqlRows Then
         With RdoUpd
            !FA_ITSEQUENCE = vFaItems(iList, 1)
            !FA_ITOPNO = Val(vFaItems(iList, 2))
            !FA_ITDRAWINGLOC = vFaItems(iList, 3)
            !FA_ITDIMENSION = vFaItems(iList, 4)
            !FA_ITDIMTOL = vFaItems(iList, 5)
            !FA_ITDIMACT = vFaItems(iList, 6)
            !FA_ITMETHOD = vFaItems(iList, 7)
            !FA_ITACCEPTED = sAcct(iList)
            .Update
            ClearResultSet RdoUpd
         End With
      End If
   Next
   Set RdoUpd = Nothing
   'Sequence
   iList = 0
   'Resequence the items
   If bUnload < 3 And optSeq.Value = vbChecked Then
      sSql = "SELECT FA_ITNUMBER,FA_ITREVISION,FA_ITSEQUENCE " _
             & "FROM FaitTable WHERE (FA_ITNUMBER='" _
             & sRptNo & "' AND FA_ITREVISION='" & sRevNo & "') " _
             & "ORDER BY FA_ITSEQUENCE,FA_ITFEATURENUM"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoUpd, ES_DYNAMIC)
      If bSqlRows Then
         With RdoUpd
            Do Until .EOF
               iList = iList + 1
               a = a + 5
               If a > 95 Then a = 95
               prg1.Value = a
               !FA_ITSEQUENCE = iList
               .Update
               .MoveNext
            Loop
            ClearResultSet RdoUpd
         End With
      End If
   End If
   Set RdoUpd = Nothing
   prg1.Value = 100
   If bUnload = 0 Or bUnload = 3 Then
      FillItems
   Else
      SysMsg "The Data Was Updated.", True
      Unload Me
   End If
   
End Sub

Private Sub AddItem()
   Dim iNewRow As Integer
   Dim RdoAdd As ADODB.Recordset
   
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   sSql = "SELECT MAX(FA_ITFEATURENUM) FROM FaitTable WHERE " _
          & "(FA_ITNUMBER='" & sRptNo & "' AND FA_ITREVISION='" _
          & sRevNo & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAdd, ES_FORWARD)
   If bSqlRows Then
      With RdoAdd
         If Not IsNull(.Fields(0)) Then iNewRow = _
                       .Fields(0) + 1
         ClearResultSet RdoAdd
      End With
   End If
   If iNewRow > 1 And clsADOCon.ADOErrNum = 0 Then
      
      clsADOCon.ADOErrNum = 0
      
      sSql = "INSERT INTO FaitTable (FA_ITNUMBER,FA_ITREVISION," _
             & "FA_ITFEATURENUM,FA_ITSEQUENCE) VALUES('" & sRptNo _
             & "','" & sRevNo & "'," & iNewRow & "," & iCurRow & ")"
      clsADOCon.ExecuteSql sSql
      If clsADOCon.ADOErrNum = 0 Then
         SysMsg "Feature Added.", True
         UpDateItems 3
         bDataChg = 1
      Else
         MsgBox "Couldn't Add The New Feature.", _
            vbInformation, Caption
      End If
   Else
      MsgBox "Couldn't Add The New Feature.", _
         vbInformation, Caption
   End If
   Set RdoAdd = Nothing
End Sub


