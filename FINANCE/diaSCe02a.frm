VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaSCe02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cost Information"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8265
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optVew 
      Height          =   255
      Left            =   360
      TabIndex        =   69
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox cmbPrt 
      Height          =   285
      Left            =   1560
      TabIndex        =   68
      Tag             =   "3"
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "diaSCe02a.frx":0000
      Height          =   320
      Left            =   5160
      Picture         =   "diaSCe02a.frx":0972
      Style           =   1  'Graphical
      TabIndex        =   67
      TabStop         =   0   'False
      ToolTipText     =   "Show Parts List"
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   350
   End
   Begin VB.CommandButton cmdFnd 
      Height          =   315
      Left            =   4680
      Picture         =   "diaSCe02a.frx":12E4
      Style           =   1  'Graphical
      TabIndex        =   66
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part"
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   350
   End
   Begin VB.TextBox txtCst 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   63
      TabStop         =   0   'False
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   5400
      Width           =   1035
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7320
      TabIndex        =   60
      TabStop         =   0   'False
      ToolTipText     =   "Update Cost"
      Top             =   600
      Width           =   875
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "___"
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4440
      TabIndex        =   58
      Top             =   6120
      Width           =   975
   End
   Begin VB.CheckBox chkPur 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   5760
      Width           =   735
   End
   Begin VB.CheckBox chkStd 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   6120
      Width           =   735
   End
   Begin VB.TextBox txtOhd 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   5
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   55
      TabStop         =   0   'False
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   4680
      Width           =   1035
   End
   Begin VB.TextBox txtMat 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   5
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   54
      TabStop         =   0   'False
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   3960
      Width           =   1035
   End
   Begin VB.TextBox txtExp 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   5
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   53
      TabStop         =   0   'False
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   4320
      Width           =   1035
   End
   Begin VB.TextBox txtLab 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   5
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   52
      TabStop         =   0   'False
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   3600
      Width           =   1035
   End
   Begin VB.TextBox txtHrs 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   5
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   51
      TabStop         =   0   'False
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   3240
      Width           =   1035
   End
   Begin VB.TextBox txtOhd 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   4
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   50
      TabStop         =   0   'False
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   4680
      Width           =   1035
   End
   Begin VB.TextBox txtMat 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   4
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   49
      TabStop         =   0   'False
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   3960
      Width           =   1035
   End
   Begin VB.TextBox txtExp 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   4
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   48
      TabStop         =   0   'False
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   4320
      Width           =   1035
   End
   Begin VB.TextBox txtLab 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   4
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   3600
      Width           =   1035
   End
   Begin VB.TextBox txtHrs 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   4
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   46
      TabStop         =   0   'False
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   3240
      Width           =   1035
   End
   Begin VB.TextBox txtHrs 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   4920
      TabIndex        =   5
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   3240
      Width           =   1035
   End
   Begin VB.TextBox txtHrs 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   2
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   3240
      Width           =   1035
   End
   Begin VB.TextBox txtHrs 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   1
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   3240
      Width           =   1035
   End
   Begin VB.TextBox txtHrs 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   3240
      Width           =   1035
   End
   Begin VB.TextBox txtLab 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   3600
      Width           =   1035
   End
   Begin VB.TextBox txtLab 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   1
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   3600
      Width           =   1035
   End
   Begin VB.TextBox txtLab 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   2
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   3600
      Width           =   1035
   End
   Begin VB.TextBox txtLab 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   4920
      TabIndex        =   6
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   3600
      Width           =   1035
   End
   Begin VB.TextBox txtStd 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   5040
      Width           =   1035
   End
   Begin VB.TextBox txtBud 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   5040
      Width           =   1035
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7680
      Top             =   2400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6525
      FormDesignWidth =   8265
   End
   Begin VB.TextBox txtExp 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   4920
      TabIndex        =   8
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   4320
      Width           =   1035
   End
   Begin VB.TextBox txtMat 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   4920
      TabIndex        =   7
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   3960
      Width           =   1035
   End
   Begin VB.TextBox txtOhd 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   4920
      TabIndex        =   9
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   4680
      Width           =   1035
   End
   Begin VB.TextBox txtExp 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   2
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   4320
      Width           =   1035
   End
   Begin VB.TextBox txtMat 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   2
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   3960
      Width           =   1035
   End
   Begin VB.TextBox txtOhd 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   2
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   4680
      Width           =   1035
   End
   Begin VB.TextBox txtExp 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   1
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   4320
      Width           =   1035
   End
   Begin VB.TextBox txtMat 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   1
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   3960
      Width           =   1035
   End
   Begin VB.TextBox txtOhd 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   1
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   4680
      Width           =   1035
   End
   Begin VB.TextBox txtExp 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   3
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   4320
      Width           =   1035
   End
   Begin VB.TextBox txtMat 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   2
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   3960
      Width           =   1035
   End
   Begin VB.TextBox txtOhd 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   4
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   4680
      Width           =   1035
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   7320
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   65
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
      PictureUp       =   "diaSCe02a.frx":1626
      PictureDn       =   "diaSCe02a.frx":176C
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Update:"
      Height          =   285
      Index           =   24
      Left            =   120
      TabIndex        =   72
      Top             =   5520
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Default BOM Rev"
      Height          =   285
      Index           =   23
      Left            =   120
      TabIndex        =   71
      Top             =   2160
      Width           =   1395
   End
   Begin VB.Label lblBOM 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   70
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Standard Cost"
      Height          =   285
      Index           =   21
      Left            =   4200
      TabIndex        =   64
      Top             =   5400
      Width           =   1635
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Revised"
      Height          =   285
      Index           =   20
      Left            =   2520
      TabIndex        =   62
      Top             =   2160
      Width           =   1275
   End
   Begin VB.Label lblRev 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4080
      TabIndex        =   61
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "At All Levels?"
      Height          =   285
      Index           =   22
      Left            =   3120
      TabIndex        =   59
      Top             =   6120
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Proposed Cost?"
      Height          =   285
      Index           =   18
      Left            =   360
      TabIndex        =   57
      Top             =   5760
      Width           =   2145
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Standard Cost?"
      Height          =   285
      Index           =   15
      Left            =   360
      TabIndex        =   56
      Top             =   6120
      Width           =   2145
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   285
      Index           =   17
      Left            =   120
      TabIndex        =   45
      Top             =   5040
      Width           =   1035
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Previous"
      Height          =   285
      Index           =   14
      Left            =   7200
      TabIndex        =   44
      Top             =   2880
      Width           =   915
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Std Cost"
      Height          =   285
      Index           =   13
      Left            =   6000
      TabIndex        =   43
      Top             =   2880
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Hours"
      Height          =   285
      Index           =   16
      Left            =   120
      TabIndex        =   42
      Top             =   3240
      Width           =   1275
   End
   Begin VB.Label lblExt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   1560
      TabIndex        =   39
      Top             =   960
      Width           =   4455
   End
   Begin VB.Line Line2 
      X1              =   1560
      X2              =   1560
      Y1              =   2760
      Y2              =   5280
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Labor Cost"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   38
      Top             =   3600
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type"
      Height          =   285
      Index           =   12
      Left            =   120
      TabIndex        =   37
      Top             =   1800
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Make/Buy/Either"
      Height          =   285
      Index           =   11
      Left            =   2520
      TabIndex        =   36
      Top             =   1800
      Width           =   1275
   End
   Begin VB.Label lblLvl 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   35
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lblMbe 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4080
      TabIndex        =   34
      Top             =   1800
      Width           =   375
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8160
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Proposed"
      Height          =   285
      Index           =   10
      Left            =   4920
      TabIndex        =   33
      Top             =   2880
      Width           =   1035
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   285
      Index           =   9
      Left            =   3840
      TabIndex        =   32
      Top             =   2880
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lower Levels"
      Height          =   285
      Index           =   8
      Left            =   2760
      TabIndex        =   31
      Top             =   2880
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "This Level"
      Height          =   285
      Index           =   7
      Left            =   1680
      TabIndex        =   30
      Top             =   2880
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Proposed Cost Components"
      Height          =   525
      Index           =   2
      Left            =   120
      TabIndex        =   29
      Top             =   2700
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Expense Cost"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   28
      Top             =   4320
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Material Cost"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   27
      Top             =   3960
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Overhead Cost"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   26
      Top             =   4680
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   25
      Top             =   600
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   24
      Top             =   240
      Width           =   1395
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   23
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "diaSCe02a"
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

'**************************************************************************************
' diaCost22 - Cost Information
'
' Created: 11/30/01 (nth)
' Revisions:
'   11/21/01 (nth) Roll up lower level cost from BOM
'   12/26/01 (nth) Removed the Select button
'   03/15/02 (nth) Enable/Disable "this level" text boxes based on part type,BOM,and Routing
'   03/18/02 (nth) Add update "this level" cost logic to lost focus
'   03/25/02 (nth) Added check boxs and logic to toggle how lower level is calulated
'   03/26/02 (nth) Moved standard costing logic to module mod_ESI_StdCost
'   03/27/02 (nth) Revised interface to include MCS cst06 functionality
'   04/08/02 (nth) Made requested revions per JLH in regards too Total and Proposed cost
'   07/11/02 (nth) Changed how the Update Proposed and Update Std cost checkbox function per JLH
'   07/11/02 (nth) Added part lookup
'   07/11/02 (nth) Added BOM view
'   12/05/02 (nth) Added ClearGrid
'
'**************************************************************************************

Dim RdoPrt As ADODB.Recordset
Dim AdoQry1 As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter

Dim bOnLoad As Byte
Dim bGoodPart As Byte
Dim bCancel As Byte
Dim sMsg As String
Dim bOnTheFly As Byte

Const COSTMASK = "#,###,##0.0000"

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'**************************************************************************************

Private Sub chkPur_Click()
   If chkPur Then
      '    'chkStd.Enabled = True
      cmdUpd.enabled = True
   Else
      'chkStd.Enabled = False
      '    chkStd = vbUnchecked
      '    If chkStd = vbUnchecked And chkPur = vbUnchecked Then
      cmdUpd.enabled = False
      '    End If
   End If
End Sub

Private Sub chkStd_Click()
   If chkStd Then
      chkAll.enabled = True
      cmdUpd.enabled = True
   Else
      '    If chkStd = vbUnchecked And chkPur = vbUnchecked Then
      '       cmdUpd.Enabled = False
      '   End If
      chkAll.enabled = False
      chkAll = vbUnchecked
   End If
End Sub

Private Sub cmbPrt_GotFocus()
   SelectFormat Me
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             x As Single, y As Single)
   bCancel = True
End Sub

Private Sub cmdFnd_Click()
   optVew.Value = vbChecked
   ViewParts.Show
End Sub

Private Sub cmdFnd_MouseDown(Button As Integer, Shift As Integer, _
                             x As Single, y As Single)
   bCancel = True
End Sub

Private Sub cmdFnd_MouseUp(Button As Integer, Shift As Integer, _
                           x As Single, y As Single)
   bCancel = False
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, Me.Caption
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub cmdUpd_Click()
   If chkPur Then
      UpdateProposed
      chkPur = vbUnchecked
   End If
   
   If chkStd Then
      UpdateStd
      chkStd = vbUnchecked
      chkAll = vbUnchecked
   End If
End Sub

Private Sub cmdVew_Click()
   ViewBom.Show
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      MouseCursor 13
      bOnLoad = False
      cmbPrt = cUR.CurrentPart
      bGoodPart = GetPart()
      If bGoodPart Then
         FillCostGrid
         cmbPrt.SetFocus
      End If
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   ' All cost levels for part
   sSql = "SELECT PARTREF, PARTNUM, PADESC, PALEVEL, PAREVDATE," _
          & "PAEXTDESC, PAMAKEBUY, PALEVLABOR, PALEVEXP, PALEVMATL, PALEVOH," _
          & "PALEVHRS, PASTDCOST, PABOMLABOR, PABOMEXP, PABOMMATL, PABOMOH," _
          & "PABOMHRS, PABOMREV, PAPREVLABOR, PAPREVEXP, PAPREVMATL, PAPREVOH," _
          & "PAPREVHRS, PATOTHRS, PATOTEXP, PATOTLABOR, PATOTMATL, PATOTOH,PAROUTING,PARRQ,PAEOQ " _
          & "FROM PartTable WHERE PARTREF = ?"
   Set AdoQry1 = New ADODB.Command
   AdoQry1.CommandText = sSql
   
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adChar
   AdoParameter1.SIZE = 30
   AdoQry1.parameters.Append AdoParameter1
   
   
   ' Operations, workcenters, and shops associated with part
'   sSql = "SELECT OPREF,OPNO,OPSETUP,OPUNIT,WCNSTDRATE,WCNOHFIXED,SHPRATE,SHPOHTOTAL, " _
'          & "WCNOHPCT,WCNOHFIXED,OPSERVPART,PASTDCOST FROM RtopTable " _
'          & "INNER JOIN WcntTable ON RtopTable.OPCENTER = WcntTable.WCNREF " _
'          & "INNER JOIN ShopTable ON RtopTable.OPSHOP = ShopTable.SHPREF " _
'          & "LEFT JOIN PartTable On RtopTable.OPSERVPART = PartTable.PARTREF " _
'          & "WHERE (RtopTable.OPREF = ?)"
'   Set RdoQry2 = RdoCon.CreateQuery("", sSql)
   
   bOnLoad = True
   
   txtHrs(0).BackColor = Me.BackColor
   txtOhd(0).BackColor = Me.BackColor
   txtLab(0).BackColor = Me.BackColor
   txtExp(0).BackColor = Me.BackColor
   txtMat(0).BackColor = Me.BackColor
   
   txtHrs(1).BackColor = Me.BackColor
   txtOhd(1).BackColor = Me.BackColor
   txtLab(1).BackColor = Me.BackColor
   txtExp(1).BackColor = Me.BackColor
   txtMat(1).BackColor = Me.BackColor
   
   txtHrs(2).BackColor = Me.BackColor
   txtOhd(2).BackColor = Me.BackColor
   txtLab(2).BackColor = Me.BackColor
   txtExp(2).BackColor = Me.BackColor
   txtMat(2).BackColor = Me.BackColor
   
   txtHrs(4).BackColor = Me.BackColor
   txtOhd(4).BackColor = Me.BackColor
   txtLab(4).BackColor = Me.BackColor
   txtExp(4).BackColor = Me.BackColor
   txtMat(4).BackColor = Me.BackColor
   
   txtHrs(5).BackColor = Me.BackColor
   txtOhd(5).BackColor = Me.BackColor
   txtLab(5).BackColor = Me.BackColor
   txtExp(5).BackColor = Me.BackColor
   txtMat(5).BackColor = Me.BackColor
   
   txtBud.BackColor = Me.BackColor
   txtStd.BackColor = Me.BackColor
   txtCst.BackColor = Me.BackColor
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   If bGoodPart Then
      cUR.CurrentPart = Trim(cmbPrt)
      SaveCurrentSelections
   End If
   Set RdoPrt = Nothing
   Set AdoParameter1 = Nothing
   Set AdoQry1 = Nothing
   FormUnload
   
   Set diaSCe02a = Nothing
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub optVew_Click()
   If optVew.Value = vbUnchecked Then
      ' Part search is closing refresh form
      cmbPrt_LostFocus
   End If
End Sub

Private Sub txtExp_GotFocus(Index As Integer)
   SelectFormat Me
End Sub

Private Sub txtExp_LostFocus(Index As Integer)
   txtExp(Index) = CheckLen(txtExp(Index), 11)
   txtExp(Index) = Format(Abs(CSng(txtExp(Index))), COSTMASK)
   
   If Index = 0 Then
      If bGoodPart Then
         On Error Resume Next

         RdoPrt!PALEVEXP = Val(txtExp(Index))
         RdoPrt.Update
         If Err <> 0 Then
            ValidateEdit Me
         End If
      End If
   End If
   
   UpdateTotals
End Sub

Private Sub txthrs_GotFocus(Index As Integer)
   SelectFormat Me
End Sub

Private Sub txtHrs_LostFocus(Index As Integer)
   txtHrs(Index) = CheckLen(txtHrs(Index), 11)
   txtHrs(Index) = Format(txtHrs(Index), COSTMASK)
   
   If Index = 0 Then
      If bGoodPart Then
         On Error Resume Next

         RdoPrt!PALEVHRS = CSng(txtHrs(Index))
         RdoPrt.Update
         If Err <> 0 Then
            ValidateEdit Me
         End If
      End If
      UpdateTotals
   End If
End Sub

Private Sub txtLab_GotFocus(Index As Integer)
   SelectFormat Me
End Sub

Private Sub txtLab_LostFocus(Index As Integer)
   txtLab(Index) = CheckLen(txtLab(Index), 11)
   txtLab(Index) = Format(Abs(CSng(txtLab(Index))), COSTMASK)
   
   If Index = 0 Then
      If bGoodPart Then
         On Error Resume Next

         RdoPrt!PALEVLABOR = CSng(txtLab(Index))
         RdoPrt.Update
         If Err <> 0 Then
            ValidateEdit Me
         End If
      End If
   End If
   UpdateTotals
End Sub

Private Sub txtMat_GotFocus(Index As Integer)
   SelectFormat Me
End Sub

Private Sub txtMat_LostFocus(Index As Integer)
   
   txtMat(Index) = CheckLen(txtMat(Index), 11)
   txtMat(Index) = Format(Abs(CSng(txtMat(Index))), COSTMASK)
   
   If Index = 0 Then
      If bGoodPart Then
         On Error Resume Next

         RdoPrt!PALEVMATL = CSng(txtMat(Index))
         RdoPrt.Update
         If Err <> 0 Then
            ValidateEdit Me
         End If
      End If
   End If
   UpdateTotals
   
End Sub

Private Sub txtOhd_GotFocus(Index As Integer)
   SelectFormat Me
End Sub

Public Function GetPart() As Byte
   Dim SPartRef As String
   
   On Error GoTo DiaErr1
   MouseCursor 13
   
   SPartRef = Compress(cmbPrt)
   AdoQry1.parameters(0).Value = SPartRef
   bSqlRows = clsADOCon.GetQuerySet(RdoPrt, AdoQry1, ES_KEYSET, True)
   
   If bSqlRows Then
      ' Check for none costable parts
      If RdoPrt!PALEVEL = "5" Or RdoPrt!PALEVEL = "6" Then
         sMsg = "Cannot Cost Part Types 5 and 6"
         MsgBox sMsg, vbInformation, Caption
         Set RdoPrt = Nothing
         ClearGrid
         GetPart = False
         cmbPrt.SetFocus
         MouseCursor 0
         Exit Function
      End If
      
      lblDsc.ForeColor = Me.ForeColor
      With RdoPrt
         ' Fill part descriptions ect.
         cmbPrt = "" & Trim(!PARTNUM)
         lblDsc = Me.ForeColor
         lblDsc = "" & Trim(!PADESC)
         lblExt = "" & Trim(!PAEXTDESC)
         lblLvl = Format(!PALEVEL, "0")
         lblMbe = "" & Trim(!PAMAKEBUY)
         lblRev = Format(!PAREVDATE, "mm/dd/yy")
         lblBOM = Format(!PABOMREV, "#0")
      End With
      GetPart = True
      
   Else
      GetPart = False
      lblDsc.ForeColor = ES_RED
      lblDsc = "*** No Current Part ***"
      lblExt = ""
      lblLvl = ""
      lblMbe = ""
      lblRev = ""
      lblBOM = ""
      txtCst = ""
      txtBud = ""
      txtStd = ""
      ClearGrid
   End If
   
   MouseCursor 0
   Exit Function
   
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub UpdateTotals()
   txtLab(2) = Format(CCur(txtLab(0)) + CCur(txtLab(1)), COSTMASK)
   txtExp(2) = Format(CSng(txtExp(0)) + CSng(txtExp(1)), COSTMASK)
   txtMat(2) = Format(CSng(txtMat(0)) + CSng(txtMat(1)), COSTMASK)
   txtOhd(2) = Format(CSng(txtOhd(0)) + CSng(txtOhd(1)), COSTMASK)
   txtHrs(2) = Format(CSng(txtHrs(0)) + CSng(txtHrs(1)), COSTMASK)
   txtBud = Format(CSng(txtLab(2)) + CSng(txtExp(2)) + CSng(txtHrs(2)) _
            + CSng(txtMat(2)) + CSng(txtOhd(2)), COSTMASK)
End Sub

Private Sub txtOhd_LostFocus(Index As Integer)
   txtOhd(Index) = CheckLen(txtOhd(Index), 11)
   txtOhd(Index) = Format(Abs(CSng(txtOhd(Index))), COSTMASK)
   
   If Index = 0 Then
      If bGoodPart Then
         On Error Resume Next
         RdoPrt!PALEVOH = CSng(txtOhd(Index))
         RdoPrt.Update
         If Err <> 0 Then
            ValidateEdit Me
         End If
      End If
   End If
   
   UpdateTotals
End Sub

Private Sub UpdateProposed()
   ' Calc the new purposed cost
   txtLab(3) = Format(CSng(txtLab(0)) + CSng(txtLab(1)), COSTMASK)
   txtExp(3) = Format(CSng(txtExp(0)) + CSng(txtExp(1)), COSTMASK)
   txtMat(3) = Format(CSng(txtMat(0)) + CSng(txtMat(1)), COSTMASK)
   txtOhd(3) = Format(CSng(txtOhd(0)) + CSng(txtOhd(1)), COSTMASK)
   txtHrs(3) = Format(CSng(txtHrs(0)) + CSng(txtHrs(1)), COSTMASK)
   SysMsg "Purposed Cost Updated.", True
End Sub

Private Sub UpdateStd()
   Dim iResponse As Integer
   iResponse = MsgBox("Update Proposed To Standard Cost?", ES_YESQUESTION, Caption)
   If iResponse = vbNo Then
      Exit Sub
   End If
   
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   Err = 0
   With RdoPrt
      ' Old Std cost
      txtLab(5) = txtLab(4)
      txtExp(5) = txtExp(4)
      txtMat(5) = txtMat(4)
      txtOhd(5) = txtOhd(4)
      txtHrs(5) = txtHrs(4)
      
      !PAPREVHRS = CSng(txtHrs(4))
      !PAPREVEXP = CSng(txtExp(4))
      !PAPREVMATL = CSng(txtMat(4))
      !PAPREVOH = CSng(txtOhd(4))
      !PAPREVLABOR = CSng(txtLab(4))
      !PAREVDATE = Format(Now, "mm/dd/yy")
      
      ' New Std cost from purposed
      !PATOTHRS = CSng(txtHrs(3))
      !PATOTEXP = CSng(txtExp(3))
      !PATOTMATL = CSng(txtMat(3))
      !PATOTOH = CSng(txtOhd(3))
      !PATOTLABOR = CSng(txtLab(3))
      
      '!PASTDCOST = CSng(txtStd)
      !PASTDCOST = _
                   CSng(txtLab(2)) + _
                   CSng(txtExp(2)) + _
                   CSng(txtHrs(2)) + _
                   CSng(txtMat(2)) + _
                   CSng(txtOhd(2))
      
      If chkAll Then
         ' Update part lower and this level records
         !PALEVHRS = CSng(txtHrs(0))
         !PALEVEXP = CSng(txtExp(0))
         !PALEVMATL = CSng(txtMat(0))
         !PALEVOH = CSng(txtOhd(0))
         !PALEVLABOR = CSng(txtLab(0))
         
         !PABOMHRS = CSng(txtHrs(1))
         !PABOMEXP = CSng(txtExp(1))
         !PABOMMATL = CSng(txtMat(1))
         !PABOMOH = CSng(txtOhd(1))
         !PABOMLABOR = CSng(txtLab(1))
      End If
      .Update
   End With
   
   If Err > 0 Then
      ValidateEdit Me
   Else
      txtLab(4) = txtLab(3)
      txtExp(4) = txtExp(3)
      txtMat(4) = txtMat(3)
      txtOhd(4) = txtOhd(3)
      txtHrs(4) = txtHrs(3)
      txtStd = Format(CSng(txtLab(4)) + CSng(txtExp(4)) + CSng(txtHrs(4)) _
               + CSng(txtMat(4)) + CSng(txtOhd(4)), COSTMASK)
      
      lblRev = Format(Now, "mm/dd/yy")
      txtCst = Format(txtStd, COSTMASK)
      
      SysMsg "Standard Cost Updated.", True
   End If
End Sub


Public Sub FillCostGrid()
   Dim SPartRef As String
   Dim RdoBom As ADODB.Recordset
   Dim bHasRouting As Byte
   Dim bHasBom As Byte
   Dim ThisCost As PartCost
   
   MouseCursor 13
   
   bOnTheFly = True ' May be an option later on.
   
   With RdoPrt
      ' Turn on proposed
      txtHrs(3).Locked = False
      txtExp(3).Locked = False
      txtMat(3).Locked = False
      txtOhd(3).Locked = False
      txtLab(3).Locked = False
      
      txtHrs(3).BackColor = cmbPrt.BackColor
      txtExp(3).BackColor = cmbPrt.BackColor
      txtMat(3).BackColor = cmbPrt.BackColor
      txtOhd(3).BackColor = cmbPrt.BackColor
      txtLab(3).BackColor = cmbPrt.BackColor
      
      ' Check For BOM
      sSql = "SELECT COUNT(BMASSYPART) FROM BmplTable WHERE BMASSYPART = '" _
             & Trim(!PartRef) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom)
      If bSqlRows Then
         If RdoBom.Fields(0) > 0 Then
            bHasBom = 1
         End If
      End If
      Set RdoBom = Nothing
      
      ' Check For Routing
      If Trim(!PAROUTING) <> "" Then
         bHasRouting = 1
      Else
         bHasRouting = 0
      End If
      
      ' Current Standard Cost
      txtCst = Format(!PASTDCOST, COSTMASK)
      IniStdCost
      ThisCost = CostPart(Trim(!PartRef), 0)
      CostPartBOM SPartRef, "" & Trim(!PABOMREV), 1
      
      ' This Level
      'txtLab(0) = Format(!PALEVLABOR, COSTMASK)
      'txtExp(0) = Format(!PALEVEXP, COSTMASK)
      'txtMat(0) = Format(!PALEVMATL, COSTMASK)
      'txtOhd(0) = Format(!PALEVOH, COSTMASK)
      'txtHrs(0) = Format(!PALEVHRS, COSTMASK)
      
      txtLab(0) = Format(ThisCost.nLabor, COSTMASK)
      txtExp(0) = Format(ThisCost.nExpense, COSTMASK)
      txtMat(0) = Format(ThisCost.nMaterial, COSTMASK)
      txtOhd(0) = Format(ThisCost.nOverhead, COSTMASK)
      txtHrs(0) = Format(ThisCost.nHours, COSTMASK)
      
      ' Lower Levels
      'txtLab(1) = Format(!PABOMLABOR, COSTMASK)
      'txtExp(1) = Format(!PABOMEXP, COSTMASK)
      'txtMat(1) = Format(!PABOMMATL, COSTMASK)
      'txtOhd(1) = Format(!PABOMOH, COSTMASK)
      'txtHrs(1) = Format(!PABOMHRS, COSTMASK)
      
      txtLab(1) = Format(BomCost(0).nLabor, COSTMASK)
      txtExp(1) = Format(BomCost(0).nExpense, COSTMASK)
      txtMat(1) = Format(BomCost(0).nMaterial, COSTMASK)
      txtOhd(1) = Format(BomCost(0).nOverhead, COSTMASK)
      txtHrs(1) = Format(BomCost(0).nHours, COSTMASK)
      
      ' Total
      UpdateTotals
      
      ' Proposed
      txtLab(3) = Format(!PABOMLABOR + !PALEVLABOR, COSTMASK)
      txtExp(3) = Format(!PABOMEXP + !PALEVEXP, COSTMASK)
      txtMat(3) = Format(!PABOMMATL + !PALEVMATL, COSTMASK)
      txtOhd(3) = Format(!PABOMOH + !PALEVOH, COSTMASK)
      txtHrs(3) = Format(!PABOMHRS + !PALEVHRS, COSTMASK)
      
      ' Std
      txtLab(4) = Format(!PATOTLABOR, COSTMASK)
      txtExp(4) = Format(!PATOTEXP, COSTMASK)
      txtMat(4) = Format(!PATOTMATL, COSTMASK)
      txtOhd(4) = Format(!PATOTOH, COSTMASK)
      txtHrs(4) = Format(!PATOTHRS, COSTMASK)
      
      txtStd = Format(CSng(txtLab(4)) + CSng(txtExp(4)) + CSng(txtHrs(4)) _
               + CSng(txtMat(4)) + CSng(txtOhd(4)), COSTMASK)
      
      ' Previous
      txtLab(5) = Format(!PAPREVLABOR, COSTMASK)
      txtExp(5) = Format(!PAPREVEXP, COSTMASK)
      txtMat(5) = Format(!PAPREVMATL, COSTMASK)
      txtOhd(5) = Format(!PAPREVOH, COSTMASK)
      txtHrs(5) = Format(!PAPREVHRS, COSTMASK)
      
      Select Case lblLvl
         Case "4"
            txtExp(0).BackColor = Me.BackColor
            txtExp(0).Locked = True
            
            txtLab(0).BackColor = Me.BackColor
            txtLab(0).Locked = True
            txtLab(0).TabStop = False
            
            txtOhd(0).BackColor = Me.BackColor
            txtOhd(0).Locked = True
            txtOhd(0).TabStop = False
            
            txtHrs(0).BackColor = Me.BackColor
            txtHrs(0).Locked = True
            txtHrs(0).TabStop = False
            
            txtMat(0).BackColor = cmbPrt.BackColor
            txtMat(0).Locked = False
            txtMat(0).TabStop = True
            txtMat(0).SetFocus
            
         Case "7"
            txtLab(0).BackColor = Me.BackColor
            txtLab(0).Locked = True
            txtLab(0).TabStop = False
            
            txtOhd(0).BackColor = Me.BackColor
            txtOhd(0).Locked = True
            txtOhd(0).TabStop = False
            
            txtHrs(0).BackColor = Me.BackColor
            txtHrs(0).Locked = True
            txtHrs(0).TabStop = False
            
            txtMat(0).BackColor = Me.BackColor
            txtMat(0).Locked = True
            txtMat(0).TabStop = False
            
            txtExp(0).BackColor = cmbPrt.BackColor
            txtExp(0).Locked = False
            txtExp(0).TabStop = True
            txtExp(0).SetFocus
            
         Case "1", "2", "3"
            If bHasRouting = 0 And bHasBom = 0 Then
               txtExp(0).BackColor = cmbPrt.BackColor
               txtExp(0).Locked = False
               txtExp(0).TabStop = True
               
               txtLab(0).BackColor = cmbPrt.BackColor
               txtLab(0).Locked = False
               txtLab(0).TabStop = True
               
               txtOhd(0).BackColor = cmbPrt.BackColor
               txtOhd(0).Locked = False
               txtOhd(0).TabStop = True
               
               txtHrs(0).BackColor = cmbPrt.BackColor
               txtHrs(0).Locked = False
               txtHrs(0).TabStop = True
               
               txtMat(0).BackColor = cmbPrt.BackColor
               txtMat(0).Locked = False
               txtMat(0).TabStop = True
               
            ElseIf bHasRouting = 1 And bHasBom = 0 Then
               
               txtExp(0).BackColor = Me.BackColor
               txtExp(0).Locked = True
               txtExp(0).TabStop = False
               
               txtLab(0).BackColor = Me.BackColor
               txtLab(0).Locked = True
               txtLab(0).TabStop = False
               
               txtOhd(0).BackColor = Me.BackColor
               txtOhd(0).Locked = True
               txtOhd(0).TabStop = False
               
               txtHrs(0).BackColor = Me.BackColor
               txtHrs(0).Locked = True
               txtHrs(0).TabStop = False
               
               txtMat(0).BackColor = cmbPrt.BackColor
               txtMat(0).Locked = False
               txtMat(0).TabStop = True
               
            ElseIf bHasRouting = 0 And bHasBom = 1 Then
               
               txtExp(0).BackColor = cmbPrt.BackColor
               txtExp(0).Locked = False
               txtExp(0).TabStop = True
               
               txtLab(0).BackColor = cmbPrt.BackColor
               txtLab(0).Locked = False
               txtLab(0).TabStop = True
               
               txtOhd(0).BackColor = cmbPrt.BackColor
               txtOhd(0).Locked = False
               txtOhd(0).TabStop = True
               
               txtHrs(0).BackColor = cmbPrt.BackColor
               txtHrs(0).Locked = False
               txtHrs(0).TabStop = True
               
               txtMat(0).BackColor = Me.BackColor
               txtMat(0).Locked = True
               txtMat(0).TabStop = False
               
            Else
               txtExp(0).BackColor = Me.BackColor
               txtExp(0).Locked = True
               txtExp(0).TabStop = False
               
               txtLab(0).BackColor = Me.BackColor
               txtLab(0).Locked = True
               txtLab(0).TabStop = False
               
               txtOhd(0).BackColor = Me.BackColor
               txtOhd(0).Locked = True
               txtOhd(0).TabStop = False
               
               txtHrs(0).BackColor = Me.BackColor
               txtHrs(0).Locked = True
               txtHrs(0).TabStop = False
               
               txtMat(0).BackColor = Me.BackColor
               txtMat(0).Locked = True
               txtMat(0).TabStop = False
            End If
            
         Case Else
            txtExp(0).BackColor = Me.BackColor
            txtExp(0).Locked = True
            txtExp(0).TabStop = False
            
            txtLab(0).BackColor = Me.BackColor
            txtLab(0).Locked = True
            txtLab(0).TabStop = False
            
            txtOhd(0).BackColor = Me.BackColor
            txtOhd(0).Locked = True
            txtOhd(0).TabStop = False
            
            txtHrs(0).BackColor = Me.BackColor
            txtHrs(0).Locked = True
            txtHrs(0).TabStop = False
            
            txtMat(0).BackColor = Me.BackColor
            txtMat(0).Locked = True
            txtMat(0).TabStop = False
      End Select
   End With
   MouseCursor 0
   Exit Sub
DiaErr1:
   sProcName = "fillcostgrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cmbPrt_LostFocus()
   If Not bCancel Then
      cmbPrt = CheckLen(cmbPrt, 30)
      If Len(cmbPrt) Then
         bGoodPart = GetPart()
         If bGoodPart Then
            FillCostGrid
         End If
      End If
   End If
End Sub

Private Sub ClearGrid()
   Dim i As Byte
   For i = 0 To 5
      txtHrs(i).BackColor = Me.BackColor
      txtHrs(i).Text = "0.000"
      txtLab(i).BackColor = Me.BackColor
      txtLab(i).Text = "0.000"
      txtOhd(i).BackColor = Me.BackColor
      txtOhd(i).Text = "0.000"
      txtExp(i).BackColor = Me.BackColor
      txtExp(i).Text = "0.000"
      txtMat(i).BackColor = Me.BackColor
      txtMat(i).Text = "0.000"
      If i = 0 Or i = 3 Then
         txtHrs(i).Locked = True
         txtLab(i).Locked = True
         txtOhd(i).Locked = True
         txtExp(i).Locked = True
         txtMat(i).Locked = True
      End If
   Next
End Sub
