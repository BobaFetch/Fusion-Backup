VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form CyclCYf05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ABC Inventory Reconciliation"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8670
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CyclCYf05a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   84
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdLotRec 
      Caption         =   "R&econcile"
      Enabled         =   0   'False
      Height          =   285
      Index           =   12
      Left            =   7680
      TabIndex        =   79
      ToolTipText     =   "Adjust Part Quantity Or Reconcile"
      Top             =   6120
      Width           =   875
   End
   Begin VB.TextBox txtLotQty 
      Enabled         =   0   'False
      Height          =   285
      Index           =   12
      Left            =   6600
      TabIndex        =   78
      Tag             =   "2"
      ToolTipText     =   "Actual Count Quantity"
      Top             =   6075
      Width           =   975
   End
   Begin VB.CommandButton cmdLotRec 
      Caption         =   "R&econcile"
      Enabled         =   0   'False
      Height          =   285
      Index           =   11
      Left            =   7680
      TabIndex        =   75
      ToolTipText     =   "Adjust Part Quantity Or Reconcile"
      Top             =   5760
      Width           =   875
   End
   Begin VB.TextBox txtLotQty 
      Enabled         =   0   'False
      Height          =   285
      Index           =   11
      Left            =   6600
      TabIndex        =   74
      Tag             =   "2"
      ToolTipText     =   "Actual Count Quantity"
      Top             =   5745
      Width           =   975
   End
   Begin VB.TextBox txtLotQty 
      Enabled         =   0   'False
      Height          =   285
      Index           =   10
      Left            =   6600
      TabIndex        =   73
      Tag             =   "2"
      ToolTipText     =   "Actual Count Quantity"
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdLotRec 
      Caption         =   "R&econcile"
      Enabled         =   0   'False
      Height          =   285
      Index           =   10
      Left            =   7680
      TabIndex        =   72
      ToolTipText     =   "Adjust Part Quantity Or Reconcile"
      Top             =   5430
      Width           =   875
   End
   Begin VB.TextBox txtLotQty 
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   6600
      TabIndex        =   69
      Tag             =   "2"
      ToolTipText     =   "Actual Count Quantity"
      Top             =   5115
      Width           =   975
   End
   Begin VB.CommandButton cmdLotRec 
      Caption         =   "R&econcile"
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   7680
      TabIndex        =   68
      ToolTipText     =   "Adjust Part Quantity Or Reconcile"
      Top             =   5115
      Width           =   875
   End
   Begin VB.TextBox txtLotQty 
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   6600
      TabIndex        =   65
      Tag             =   "2"
      ToolTipText     =   "Actual Count Quantity"
      Top             =   4815
      Width           =   975
   End
   Begin VB.CommandButton cmdLotRec 
      Caption         =   "R&econcile"
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   7680
      TabIndex        =   64
      ToolTipText     =   "Adjust Part Quantity Or Reconcile"
      Top             =   4815
      Width           =   875
   End
   Begin VB.TextBox txtLotQty 
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   6600
      TabIndex        =   61
      Tag             =   "2"
      ToolTipText     =   "Actual Count Quantity"
      Top             =   4500
      Width           =   975
   End
   Begin VB.CommandButton cmdLotRec 
      Caption         =   "R&econcile"
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   7680
      TabIndex        =   60
      ToolTipText     =   "Adjust Part Quantity Or Reconcile"
      Top             =   4500
      Width           =   875
   End
   Begin VB.TextBox txtLotQty 
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   6600
      TabIndex        =   57
      Tag             =   "2"
      ToolTipText     =   "Actual Count Quantity"
      Top             =   4185
      Width           =   975
   End
   Begin VB.CommandButton cmdLotRec 
      Caption         =   "R&econcile"
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   7680
      TabIndex        =   56
      ToolTipText     =   "Adjust Part Quantity Or Reconcile"
      Top             =   4185
      Width           =   875
   End
   Begin VB.TextBox txtLotQty 
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   6600
      TabIndex        =   53
      Tag             =   "2"
      ToolTipText     =   "Actual Count Quantity"
      Top             =   3885
      Width           =   975
   End
   Begin VB.CommandButton cmdLotRec 
      Caption         =   "R&econcile"
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   7680
      TabIndex        =   52
      ToolTipText     =   "Adjust Part Quantity Or Reconcile"
      Top             =   3885
      Width           =   875
   End
   Begin VB.TextBox txtLotQty 
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   6600
      TabIndex        =   49
      Tag             =   "2"
      ToolTipText     =   "Actual Count Quantity"
      Top             =   3570
      Width           =   975
   End
   Begin VB.CommandButton cmdLotRec 
      Caption         =   "R&econcile"
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   7680
      TabIndex        =   48
      ToolTipText     =   "Adjust Part Quantity Or Reconcile"
      Top             =   3570
      Width           =   875
   End
   Begin VB.TextBox txtLotQty 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   6600
      TabIndex        =   45
      Tag             =   "2"
      ToolTipText     =   "Actual Count Quantity"
      Top             =   3255
      Width           =   975
   End
   Begin VB.CommandButton cmdLotRec 
      Caption         =   "R&econcile"
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   7680
      TabIndex        =   44
      ToolTipText     =   "Adjust Part Quantity Or Reconcile"
      Top             =   3255
      Width           =   875
   End
   Begin VB.TextBox txtLotQty 
      Height          =   285
      Index           =   2
      Left            =   6600
      TabIndex        =   41
      Tag             =   "2"
      ToolTipText     =   "Actual Count Quantity"
      Top             =   2955
      Width           =   975
   End
   Begin VB.CommandButton cmdLotRec 
      Caption         =   "R&econcile"
      Height          =   285
      Index           =   2
      Left            =   7680
      TabIndex        =   40
      ToolTipText     =   "Adjust Part Quantity Or Reconcile"
      Top             =   2955
      Width           =   875
   End
   Begin VB.TextBox txtLotQty 
      Height          =   285
      Index           =   1
      Left            =   6600
      TabIndex        =   9
      Tag             =   "2"
      ToolTipText     =   "Actual Count Quantity"
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdLotRec 
      Caption         =   "R&econcile"
      Height          =   285
      Index           =   1
      Left            =   7680
      TabIndex        =   37
      ToolTipText     =   "Adjust Part Quantity Or Reconcile"
      Top             =   2640
      Width           =   875
   End
   Begin VB.CheckBox optShow 
      Alignment       =   1  'Right Justify
      Caption         =   "Show Reconciled"
      Height          =   255
      Left            =   5760
      TabIndex        =   4
      ToolTipText     =   "Show All Items Including Previously Reconciled"
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CommandButton cmdLst 
      Caption         =   "<< &Last    "
      Enabled         =   0   'False
      Height          =   315
      Left            =   6750
      TabIndex        =   7
      ToolTipText     =   "Last Part Number"
      Top             =   6480
      Width           =   875
   End
   Begin VB.CommandButton cmdNxt 
      Caption         =   " &Next >>"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7680
      TabIndex        =   8
      ToolTipText     =   "Next Part Number"
      Top             =   6480
      Width           =   875
   End
   Begin VB.CommandButton cmdAdj 
      Caption         =   "R&econcile"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7680
      TabIndex        =   6
      ToolTipText     =   "Adjust Part Quantity Or Reconcile"
      Top             =   1800
      Width           =   875
   End
   Begin VB.TextBox txtAQty 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6600
      TabIndex        =   5
      Tag             =   "2"
      ToolTipText     =   "Actual Count Quantity (Page Up For Previous, Page Down For Next)"
      Top             =   1800
      Width           =   975
   End
   Begin VB.CheckBox optRec 
      Alignment       =   1  'Right Justify
      Caption         =   "RE"
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      TabIndex        =   26
      ToolTipText     =   "Reconciled"
      Top             =   2100
      Width           =   855
   End
   Begin VB.CheckBox optLots 
      Alignment       =   1  'Right Justify
      Caption         =   "LT"
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      TabIndex        =   17
      ToolTipText     =   "Lot Tracked"
      Top             =   1800
      Width           =   855
   End
   Begin VB.Frame z2 
      Height          =   60
      Left            =   240
      TabIndex        =   16
      Top             =   1440
      Width           =   8325
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "S&elect"
      Height          =   315
      Left            =   7680
      TabIndex        =   3
      ToolTipText     =   "Fill The Form With Qualifying Items"
      Top             =   720
      Width           =   875
   End
   Begin VB.ComboBox cmbCid 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "List Includes Cycle ID's Not Locked Or Completed"
      Top             =   360
      Width           =   2115
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   3375
   End
   Begin VB.ComboBox txtPlan 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   1
      Tag             =   "4"
      ToolTipText     =   "Planned Inventory Date"
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   7680
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   6240
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6915
      FormDesignWidth =   8670
   End
   Begin VB.Label lblQoh 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   0
      Left            =   5520
      TabIndex        =   85
      ToolTipText     =   "Current Quantity (Part QOH Now)"
      Top             =   2100
      Width           =   972
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reconcile All Lots In One Sitting               "
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
      Index           =   14
      Left            =   5520
      TabIndex        =   83
      ToolTipText     =   "Have Lots Prepared To Reconcile All At One Time"
      Top             =   2400
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "All Items Should Be Reconciled.  Including Those Without Change"
      Height          =   255
      Index           =   13
      Left            =   600
      TabIndex        =   82
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label lblLot 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   12
      Left            =   1200
      TabIndex        =   81
      Top             =   6075
      Width           =   3615
   End
   Begin VB.Label lblLotRem 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   12
      Left            =   5520
      TabIndex        =   80
      Top             =   6075
      Width           =   975
   End
   Begin VB.Label lblLot 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   11
      Left            =   1200
      TabIndex        =   77
      Top             =   5745
      Width           =   3615
   End
   Begin VB.Label lblLotRem 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   11
      Left            =   5520
      TabIndex        =   76
      Top             =   5745
      Width           =   975
   End
   Begin VB.Label lblLotRem 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   5520
      TabIndex        =   71
      Top             =   5430
      Width           =   975
   End
   Begin VB.Label lblLot 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   1200
      TabIndex        =   70
      Top             =   5430
      Width           =   3615
   End
   Begin VB.Label lblLotRem 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   5520
      TabIndex        =   67
      Top             =   5115
      Width           =   975
   End
   Begin VB.Label lblLot 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   1200
      TabIndex        =   66
      Top             =   5115
      Width           =   3615
   End
   Begin VB.Label lblLotRem 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   5520
      TabIndex        =   63
      Top             =   4815
      Width           =   975
   End
   Begin VB.Label lblLot 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   1200
      TabIndex        =   62
      Top             =   4815
      Width           =   3615
   End
   Begin VB.Label lblLotRem 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   5520
      TabIndex        =   59
      Top             =   4500
      Width           =   975
   End
   Begin VB.Label lblLot 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   1200
      TabIndex        =   58
      Top             =   4500
      Width           =   3615
   End
   Begin VB.Label lblLotRem 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   5520
      TabIndex        =   55
      Top             =   4185
      Width           =   975
   End
   Begin VB.Label lblLot 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   1200
      TabIndex        =   54
      Top             =   4185
      Width           =   3615
   End
   Begin VB.Label lblLotRem 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   5520
      TabIndex        =   51
      Top             =   3885
      Width           =   975
   End
   Begin VB.Label lblLot 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   1200
      TabIndex        =   50
      Top             =   3885
      Width           =   3615
   End
   Begin VB.Label lblLotRem 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   5520
      TabIndex        =   47
      Top             =   3570
      Width           =   975
   End
   Begin VB.Label lblLot 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   1200
      TabIndex        =   46
      Top             =   3570
      Width           =   3615
   End
   Begin VB.Label lblLotRem 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   5520
      TabIndex        =   43
      Top             =   3255
      Width           =   975
   End
   Begin VB.Label lblLot 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   1200
      TabIndex        =   42
      Top             =   3255
      Width           =   3615
   End
   Begin VB.Label lblLotRem 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   5520
      TabIndex        =   39
      Top             =   2955
      Width           =   975
   End
   Begin VB.Label lblLot 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   1200
      TabIndex        =   38
      Top             =   2955
      Width           =   3615
   End
   Begin VB.Label lblLotRem 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   5520
      TabIndex        =   36
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblLot 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   35
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lots Available For Inventory:"
      Enabled         =   0   'False
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
      Index           =   12
      Left            =   240
      TabIndex        =   34
      ToolTipText     =   "Part Number/Description"
      Top             =   2420
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Adjust         "
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
      Index           =   11
      Left            =   7680
      TabIndex        =   33
      ToolTipText     =   "Adjust Or Reconcil Quantity"
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Actual Qty    "
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
      Index           =   9
      Left            =   6600
      TabIndex        =   32
      ToolTipText     =   "Actual Count"
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   3840
      TabIndex        =   31
      ToolTipText     =   "Total Items Included"
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Item"
      Height          =   255
      Index           =   4
      Left            =   2640
      TabIndex        =   30
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Items"
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   29
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   1560
      TabIndex        =   28
      ToolTipText     =   "Total Items Included"
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblLoc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   27
      ToolTipText     =   "Inventory Location"
      Top             =   2100
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty On Hand"
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
      Left            =   5520
      TabIndex        =   25
      ToolTipText     =   "Recorded Quantity On Hand"
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Std Cost/Loc   "
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
      Left            =   1200
      TabIndex        =   24
      ToolTipText     =   "Standard Cost/Location"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblQoh 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   5520
      TabIndex        =   23
      ToolTipText     =   "Recorded Quantity On Hand (Date Of Count)"
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblCost 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   22
      ToolTipText     =   "Standard Cost"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number                                                          "
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
      Left            =   2400
      TabIndex        =   21
      ToolTipText     =   "Part Number/Description"
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   2400
      TabIndex        =   20
      ToolTipText     =   "Part Description"
      Top             =   2100
      Width           =   3015
   End
   Begin VB.Label lblPart 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   2400
      TabIndex        =   19
      ToolTipText     =   "Part Number"
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lots/Recon"
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
      Index           =   0
      Left            =   240
      TabIndex        =   18
      ToolTipText     =   "Lot Tracked Part/Already Reconciled"
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Retrieve List"
      Height          =   255
      Index           =   8
      Left            =   5760
      TabIndex        =   15
      Top             =   765
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cycle Count ID"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   14
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   13
      Top             =   765
      Width           =   1335
   End
   Begin VB.Label lblCabc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3840
      TabIndex        =   12
      ToolTipText     =   "ABC Code Selected"
      Top             =   360
      Width           =   405
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Planned Date"
      Height          =   255
      Index           =   7
      Left            =   4440
      TabIndex        =   11
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "CyclCYf05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'3/16/04 New
'1/12/05 Added CheckReconciled
'        Added additional NextDate trap
'        Proper filling when changing Cycle Counts
'        Corrected Inva where the actual count>current inventory
'        Ask to mark completed in FillList and cmdAdj
'1/17/04 Change Inva/Lots dates to CC planned date (Larry)
'1/17/04 Added Accounts per Larry. See notes
'12/19/06 Changed PAQOH to CIPAQOH (FillList). Added Current QOH lblQoh(0).
Option Explicit
Dim bOnLoad As Byte
Dim bGoodCount As Byte

Dim iTotalList As Integer
Dim iTotalLots As Integer
Dim iIndex As Integer
Dim lCOUNTER As Long

Dim sCreditAcct As String
Dim sDebitAcct As String

Dim vNextDate As Variant

Dim sParts(1000, 5) As String 'Location,PartRef, Number, Description, Lot Tracked
Dim cValue(1000, 5) As Currency 'Cost, Qoh (At Count), Actual Qty, Reconciled, PartTable.PAQOH
Dim vCycleLots(100, 4) As Variant 'Lotnumber, Remaining, UserId, AdjustQty

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'1/17/04
'Larry's note
'For Debit Account use "Inventory Over/Short" account.
'For Credit Account use inventory/expense material account.
'loaded in part number,

Private Sub GetAccounts(PartNumber As String)
   Dim rdoAct As ADODB.Recordset
   Dim bType As Byte
   Dim sPcode As String
   
   On Error GoTo DiaErr1
   sDebitAcct = ""
   sCreditAcct = ""
   sSql = "SELECT COADJACCT FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         If Not IsNull(!COADJACCT) Then _
                       sDebitAcct = "" & Trim(!COADJACCT)
         ClearResultSet rdoAct
         Set rdoAct = Nothing
      End With
   End If
   
   'Use current Part
   sSql = "Qry_GetExtPartAccounts '" & PartNumber & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         sPcode = "" & Trim(!PAPRODCODE)
         If sDebitAcct = "" Then sDebitAcct = "" & Trim(!PACGSMATACCT)
         sCreditAcct = "" & Trim(!PAINVEXPACCT)
         ClearResultSet rdoAct
         Set rdoAct = Nothing
      End With
   Else
      sCreditAcct = ""
      sDebitAcct = ""
      Exit Sub
   End If
   If sDebitAcct = "" Or sCreditAcct = "" Then
      'None in one or both there, try Product code
      sSql = "Qry_GetPCodeAccounts '" & sPcode & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
      If bSqlRows Then
         With rdoAct
            If sDebitAcct = "" Then sCreditAcct = "" & Trim(!PCCGSMATACCT)
            If sCreditAcct = "" Then sDebitAcct = "" & Trim(!PCINVMATACCT)
            ClearResultSet rdoAct
            Set rdoAct = Nothing
         End With
      End If
      If sDebitAcct = "" Or sCreditAcct = "" Then
         'Still none, we'll check the common
         sSql = "SELECT COCGSMATACCT" & Trim(str(bType)) & "," _
                & "COINVMATACCT" & Trim(str(bType)) & " " _
                & "FROM ComnTable WHERE COREF=1"
         bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
         If bSqlRows Then
            With rdoAct
               If sCreditAcct = "" Then sDebitAcct = "" & Trim(.Fields(0))
               If sDebitAcct = "" Then sCreditAcct = "" & Trim(.Fields(1))
               ClearResultSet rdoAct
               Set rdoAct = Nothing
            End With
         End If
      End If
   End If
   'After this excercise, we'll give up if none are found
   Set rdoAct = Nothing
   Exit Sub
   
DiaErr1:
   'Just bail for now. May not have anything set
   'CurrError.Number = Err
   'CurrError.Description = Err.Description
   'DoModuleErrors Me
   On Error GoTo 0
   
End Sub

Private Sub FillList()
   Dim RdoAbc As ADODB.Recordset
   Dim bResponse As Byte
   Dim iReconciled As Integer
   Dim sMsg As String
   On Error GoTo DiaErr1
   iIndex = 0
   iTotalList = 0
   lblCount = "0"
   lblItem = "0"
   txtAQty.Enabled = False
   cmdNxt.Enabled = False
   cmdLst.Enabled = False
   cmdAdj.Enabled = False
   
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALOCATION,PAQOH,PAABC," _
          & "PASTDCOST,CIREF,CIPARTREF,CILOTTRACK,CIPAQOH,CIACTUALQOH,CIRECONCILED " _
          & "FROM PartTable,CcitTable WHERE (CIREF='" & cmbCid & "' AND " _
          & "PARTREF=CIPARTREF "
   If optShow.Value = vbUnchecked Then sSql = sSql & "AND CIRECONCILED=0"
   sSql = sSql & ") ORDER BY PALOCATION,PARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAbc, ES_FORWARD)
   If bSqlRows Then
      With RdoAbc
         Do Until .EOF
            iTotalList = iTotalList + 1
            sParts(iTotalList, 0) = "" & Trim(!PALOCATION)
            sParts(iTotalList, 1) = "" & Trim(!PartRef)
            sParts(iTotalList, 2) = "" & Trim(!PartNum)
            sParts(iTotalList, 3) = "" & Trim(!PADESC)
            sParts(iTotalList, 4) = "" & str$(!CILOTTRACK)
            cValue(iTotalList, 0) = !PASTDCOST
            '12/19/06
            cValue(iTotalList, 1) = !CIPAQOH
            cValue(iTotalList, 2) = !CIACTUALQOH
            cValue(iTotalList, 3) = !CIRECONCILED
            cValue(iTotalList, 4) = !PAQOH
            If !CIRECONCILED = 1 Then iReconciled = iReconciled + 1
            .MoveNext
         Loop
         ClearResultSet RdoAbc
         Set RdoAbc = Nothing
      End With
   End If
   
   If iReconciled = iTotalList Then
      sMsg = "There Are No Items Left To Reconcile. " & vbCr _
             & "Mark As Reconciled (Complete) Now?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         MarkReconciled
         Exit Sub
      End If
   End If
   If iTotalList > 0 Then
      vNextDate = GetNextDate()
      cmdSel.Enabled = False
      cmdNxt.Enabled = True
      cmdLst.Enabled = True
      lblCount = iTotalList
      iIndex = 1
      GetNextItem
   End If
   Set RdoAbc = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "filllist"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetCycleCount() As Byte
   Dim RdoCid As ADODB.Recordset
   On Error GoTo DiaErr1
   CloseLotBoxes
   lblCost(1) = ""
   lblLoc(1) = ""
   lblPart(1) = ""
   lblDsc(1) = ""
   lblLotRem(1) = ""
   txtAQty = ""
   cmdAdj.Enabled = False
   sSql = "Qry_GetCycleCount '" & Trim(cmbCid) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCid, ES_FORWARD)
   If bSqlRows Then
      With RdoCid
         lblCabc = "" & Trim(!CCABCCODE)
         txtDsc = "" & Trim(!CCDESC)
         txtPlan = Format(!CCPLANDATE, "mm/dd/yy")
         GetCycleCount = 1
         cmdSel.Enabled = True
         ClearResultSet RdoCid
      End With
   Else
      GetCycleCount = 0
      cmdSel.Enabled = False
      Select Case MsgBox("That Count ID Wasn't Found, Is Locked, Or Is Not Saved.  Do you wish to cancel?", _
         vbQuestion + vbYesNo, Caption)
      Case vbYes
         Exit Function
      End Select
   End If
   Set RdoCid = Nothing
   Exit Function
   
DiaErr1:
   GetCycleCount = 0
   sProcName = "getcycleco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Sub cmbCid_Click()
   GetCycleCount
   
End Sub

Private Sub cmbCid_LostFocus()
   bGoodCount = GetCycleCount()
   
End Sub



Private Sub cmdAdj_Click()
   'These are not Lot Tracked Item
   
   Dim bResponse As Byte
   Dim cActivityQty As Currency
   Dim cAdjustmentQty As Currency
   Dim cQuantityOh As Currency
   Dim sMsg As String
   Dim sLotNumber As String
   Dim vAdate As Variant
   
   vAdate = Format(GetServerDateTime(), "mm/dd/yy hh:mm")
   bResponse = MsgBox("Reconcile Item " & Trim(lblItem) & "?", _
               ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      cmdAdj.Enabled = False
      txtAQty.Enabled = False
      On Error Resume Next
      lCOUNTER = GetLastActivity()
      sLotNumber = GetNextLotNumber()
      cValue(iIndex, 2) = Val(txtAQty)
      cQuantityOh = Val(lblQoh(1))
      If cQuantityOh = Val(txtAQty) Then
         'Equal don't adjust quantity
         clsADOCon.BeginTrans
         sSql = "UPDATE CcitTable SET CIACTUALQOH=" & cValue(iIndex, 2) _
                & ",CIRECONCILEDDATE='" & vAdate _
                & "',CIRECONCILED=1 WHERE (CIREF='" & cmbCid & "' AND " _
                & "CIPARTREF='" & sParts(iIndex, 1) & "')"
         clsADOCon.ExecuteSQL sSql
         
         sSql = "UPDATE PartTable SET PANEXTCYCLEDATE='" & vNextDate _
                & "',PAQOH=" & cValue(iIndex, 2) & ",PALOTQTYREMAINING=" _
                & cValue(iIndex, 2) & " WHERE PARTREF='" _
                & sParts(iIndex, 1) & "'"
         clsADOCon.ExecuteSQL sSql
         
         sSql = "UPDATE LohdTable SET LOTAVAILABLE=0,LOTREMAININGQTY=0 " _
                & "WHERE LOTPARTREF='" & sParts(iIndex, 1) & "'"
         clsADOCon.ExecuteSQL sSql
         
         cActivityQty = GetActivityQuantity(sParts(iIndex, 1))
         If cActivityQty <> cQuantityOh Then
            lCOUNTER = lCOUNTER + 1
            RepairInventory sParts(iIndex, 1), cQuantityOh - cActivityQty, lCOUNTER
         End If
         'This operation is complete
         GoTo Fin1
         

         
      Else
         'Not equal adjust quantity
         clsADOCon.ADOErrNum = 0
         clsADOCon.BeginTrans
         
         sSql = "UPDATE CcitTable SET CIACTUALQOH=" & cValue(iIndex, 2) _
                & ",CIRECONCILEDDATE='" & vAdate _
                & "',CIRECONCILED=1 WHERE (CIREF='" & cmbCid & "' AND " _
                & "CIPARTREF='" & sParts(iIndex, 1) & "')"
         clsADOCon.ExecuteSQL sSql
         
         sSql = "UPDATE PartTable SET PANEXTCYCLEDATE='" & vNextDate _
                & "',PAQOH=" & cValue(iIndex, 2) & ",PALOTQTYREMAINING=" _
                & cValue(iIndex, 2) & " WHERE PARTREF='" _
                & sParts(iIndex, 1) & "'"
         clsADOCon.ExecuteSQL sSql
         
         sSql = "UPDATE LohdTable SET LOTAVAILABLE=0,LOTREMAININGQTY=0 " _
                & "WHERE LOTPARTREF='" & sParts(iIndex, 1) & "'"
         clsADOCon.ExecuteSQL sSql
         
         cActivityQty = GetActivityQuantity(sParts(iIndex, 1))
         If cQuantityOh <> cActivityQty Then
            lCOUNTER = lCOUNTER + 1
            RepairInventory sParts(iIndex, 1), cQuantityOh - cActivityQty, lCOUNTER
            cActivityQty = cQuantityOh
         End If
         If cActivityQty <> cValue(iIndex, 2) Then
            'Inventory
            If cActivityQty > cValue(iIndex, 2) Then
               'Reduce Inventory Activity
               cAdjustmentQty = cActivityQty - cValue(iIndex, 2)
               lCOUNTER = InsertActivity(cAdjustmentQty, 1)
            Else
               'Increase Inventory Activity
               cAdjustmentQty = cValue(iIndex, 2) - cActivityQty
               lCOUNTER = InsertActivity(cAdjustmentQty, 0)
            End If
         End If
         If cQuantityOh < Val(txtAQty) Then
            'Actual Less than recorded
            sSql = "INSERT INTO LohdTable (LOTNUMBER,LOTUSERLOTID," _
                   & "LOTPARTREF,LOTPDATE,LOTORIGINALQTY,LOTREMAININGQTY," _
                   & "LOTUNITCOST,LOTDATECOSTED,LOTCOMMENTS,LOTLOCATION) " _
                   & "VALUES('" _
                   & sLotNumber & "','ABC Cycle-" & sLotNumber & "','" & sParts(iIndex, 1) _
                   & "','" & txtPlan & "'," & cValue(iIndex, 2) & "," _
                   & cValue(iIndex, 2) & "," & cValue(iIndex, 0) & ",'" _
                   & Format(ES_SYSDATE, "mm/dd/yy") & "','ABC Cycle Count','" _
                   & sParts(iIndex, 0) & "')"
            clsADOCon.ExecuteSQL sSql
            
            sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
                   & "LOITYPE,LOIPARTREF,LOIPDATE,LOIQUANTITY," _
                   & "LOIACTIVITY,LOICOMMENT) " _
                   & "VALUES('" _
                   & sLotNumber & "',1,30,'" & sParts(iIndex, 1) _
                   & "','" & txtPlan & "'," _
                   & cValue(iIndex, 2) & "," & lCOUNTER & ",'" _
                   & "ABC Cycle Count" & "')"
            clsADOCon.ExecuteSQL sSql
         Else
            'Actual is greater than recorded
            sSql = "INSERT INTO LohdTable (LOTNUMBER,LOTUSERLOTID," _
                   & "LOTPARTREF,LOTPDATE,LOTORIGINALQTY,LOTREMAININGQTY," _
                   & "LOTUNITCOST,LOTDATECOSTED,LOTCOMMENTS,LOTLOCATION) " _
                   & "VALUES('" _
                   & sLotNumber & "','ABC Cycle-" & sLotNumber & "','" & sParts(iIndex, 1) _
                   & "','" & txtPlan & "'," & cValue(iIndex, 2) & "," _
                   & cValue(iIndex, 2) & "," & cValue(iIndex, 0) & ",'" _
                   & Format(ES_SYSDATE, "mm/dd/yy") & "','ABC Cycle Count','" _
                   & sParts(iIndex, 0) & "')"
            clsADOCon.ExecuteSQL sSql
            
            sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
                   & "LOITYPE,LOIPARTREF,LOIPDATE,LOIQUANTITY," _
                   & "LOIACTIVITY,LOICOMMENT) " _
                   & "VALUES('" _
                   & sLotNumber & "',1,30,'" & sParts(iIndex, 1) _
                   & "','" & txtPlan & "'," _
                   & cValue(iIndex, 2) & "," & lCOUNTER & ",'" _
                   & "ABC Cycle Count" & "')"
            clsADOCon.ExecuteSQL sSql
         End If
Fin1:          MouseCursor 0
         If clsADOCon.ADOErrNum = 0 Then
            clsADOCon.CommitTrans
            sSql = "UPDATE LohdTable SET LOTCOUNTQTY=LOTORIGINALQTY " _
                   & "WHERE (LOTPARTREF='" & sParts(iIndex, 1) & "' AND " _
                   & "LOTAVAILABLE=1)"
            clsADOCon.ExecuteSQL sSql
            cValue(iIndex, 3) = 1
            optRec.Value = vbChecked
            SysMsg "The Item Was Reconciled", True
         Else
            clsADOCon.RollbackTrans
            clsADOCon.ADOErrNum = 0
            SysMsg "The Item Could Not Be Reconciled", True
         End If
      End If
   Else
      CancelTrans
   End If
   
   bResponse = CheckReconciled()
   If bResponse = 0 Then
      sMsg = "There Are No Items Left To Reconcile. " & vbCr _
             & "Mark As Reconciled (Complete) Now?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then MarkReconciled
   End If
   
End Sub

Private Sub cmdAdj_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "5455"
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdLotRec_Click(Index As Integer)
   Dim bResponse As Byte
   Dim lLOTRECORD As Long
   Dim cRLotQty As Currency
   Dim cAlotqty As Currency
   
   cRLotQty = Val(lblLotRem(Index))
   cAlotqty = Val(txtLotQty(Index))
   
   On Error Resume Next
   bResponse = MsgBox("Reconcile This Lot?", ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      CancelTrans
      Exit Sub
   End If
   lCOUNTER = GetLastActivity()
   If cAlotqty < 0 Then
      If cRLotQty + cAlotqty < 0 Then
         Beep
         MsgBox "The Quantity Entered Will Leave A Negative Balance.", _
            vbExclamation, Caption
         txtLotQty(Index) = lblLotRem(Index)
      End If
   End If
   Err.Clear
   If cRLotQty = cAlotqty Then
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      
      sSql = "UPDATE LohdTable SET LOTCOUNTQTY=" & cAlotqty & ",LOTCOUNTDATE='" _
             & txtPlan & "' WHERE LOTNUMBER='" _
             & vCycleLots(Index, 0) & "'"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "UPDATE CcltTable SET CLLOTADJUSTQTY=" & cAlotqty & " WHERE (CLREF='" _
             & cmbCid.Text & "' AND CLLOTNUMBER='" & vCycleLots(Index, 0) & "')"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         cmdLotRec(Index).Enabled = False
         txtLotQty(Index).Enabled = False
         SysMsg "Lot Reconciled.", True
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         
         MsgBox "Couldn't Reconcile the Lot.", _
            vbExclamation, Caption
      End If
   Else
      If cRLotQty > cAlotqty Then
         'Decrease Lot
         Err.Clear
         lLOTRECORD = GetNextLotRecord(str$(vCycleLots(Index, 0)))
         clsADOCon.BeginTrans
         clsADOCon.ADOErrNum = 0
         
         sSql = "UPDATE LohdTable SET LOTREMAININGQTY=" _
                & cAlotqty & ", LOTCOUNTQTY=" & cAlotqty & ",LOTCOUNTDATE='" _
                & txtPlan & "' WHERE LOTNUMBER='" _
                & vCycleLots(Index, 0) & "'"
         clsADOCon.ExecuteSQL sSql
         
         lCOUNTER = InsertActivity(cRLotQty - cAlotqty, 1)
         sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
                & "LOITYPE,LOIPARTREF,LOIPDATE,LOIQUANTITY," _
                & "LOIACTIVITY,LOICOMMENT) " _
                & "VALUES('" _
                & vCycleLots(Index, 0) & "'," & lLOTRECORD & ",30,'" _
                & sParts(iIndex, 1) _
                & "','" & txtPlan & "',-" _
                & cRLotQty - cAlotqty & "," & lCOUNTER & ",'" _
                & "ABC Cycle Count" & "')"
         clsADOCon.ExecuteSQL sSql
         
         If clsADOCon.ADOErrNum = 0 Then
            clsADOCon.CommitTrans
            cmdLotRec(Index).Enabled = False
            txtLotQty(Index).Enabled = False
            SysMsg "Lot Reconciled.", True
         Else
            clsADOCon.RollbackTrans
            MsgBox "Couldn't Reconcile the Lot.", _
               vbExclamation, Caption
         End If
      Else
         'Increase lot
         clsADOCon.ADOErrNum = 0
         clsADOCon.BeginTrans
         
         
         sSql = "UPDATE LohdTable SET LOTREMAININGQTY=" _
                & cAlotqty & ", LOTCOUNTQTY=" & cAlotqty & ",LOTCOUNTDATE='" _
                & txtPlan & "' WHERE LOTNUMBER='" _
                & vCycleLots(Index, 0) & "'"
         clsADOCon.ExecuteSQL sSql
         
         lCOUNTER = InsertActivity(cAlotqty - cRLotQty, 0)
         sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
                & "LOITYPE,LOIPARTREF,LOIPDATE,LOIQUANTITY," _
                & "LOIACTIVITY,LOICOMMENT) " _
                & "VALUES('" _
                & vCycleLots(Index, 0) & "'," & lLOTRECORD & " ,30,'" _
                & sParts(iIndex, 1) _
                & "','" & txtPlan & "',-" _
                & cAlotqty - cRLotQty & "," & lCOUNTER & ",'" _
                & "ABC Cycle Count" & "')"
         clsADOCon.ExecuteSQL sSql
         
         If clsADOCon.ADOErrNum = 0 Then
            clsADOCon.CommitTrans
            cmdLotRec(Index).Enabled = False
            txtLotQty(Index).Enabled = False
            SysMsg "Lot Reconciled.", True
         Else
            clsADOCon.RollbackTrans
            clsADOCon.ADOErrNum = 0
            MsgBox "Couldn't Reconcile the Lot.", _
               vbExclamation, Caption
         End If
      End If
   End If
   If clsADOCon.ADOErrNum = 0 Then TotalLotQuantity True
   Exit Sub
   
End Sub

Private Sub cmdLotRec_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub cmdLst_Click()
   iIndex = iIndex - 1
   If iIndex < 1 Then iIndex = 1
   GetNextItem
   
End Sub

Private Sub cmdNxt_Click()
   iIndex = iIndex + 1
   If iIndex > iTotalList Then iIndex = iTotalList
   GetNextItem
   
End Sub

Private Sub cmdSel_Click()
   FillList
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      vNextDate = GetNextDate()
   End If
   bOnLoad = 0
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
   Set CyclCYf05a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDsc.BackColor = Es_FormBackColor
   txtPlan.BackColor = Es_FormBackColor
   txtPlan.ToolTipText = "Planned Inventory Date"
   z1(13).ForeColor = ES_BLUE
   CloseLotBoxes
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   cmbCid.Clear
   sSql = "SELECT CCREF FROM CchdTable WHERE (CCCOUNTLOCKED=1 AND " _
          & "CCUPDATED=0)"
   LoadComboBox cmbCid, -1
   If cmbCid.ListCount > 0 Then
      If Trim(cmbCid) = "" Then cmbCid = cmbCid.List(0)
      'bGoodCount = GetCycleCount()
   Else
      MsgBox "There Are No Locked And Not Reconciled Counts Recorded.", _
         vbInformation, Caption
      Unload Me
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetNextItem()
   Dim bCycleLots As Byte
   Dim iRow As Integer
   On Error Resume Next
   CloseLotBoxes
   lblItem = iIndex
   lblLoc(1) = sParts(iIndex, 0)
   lblPart(1) = sParts(iIndex, 2)
   lblDsc(1) = sParts(iIndex, 3)
   lblCost(1) = Format(cValue(iIndex, 0), "###,###,##0.000")
   lblQoh(1) = Format(cValue(iIndex, 1), ES_QuantityDataFormat)
   lblQoh(0) = Format(cValue(iIndex, 4), ES_QuantityDataFormat)
   If cValue(iIndex, 3) > 0 Then
      txtAQty = Format(cValue(iIndex, 2), ES_QuantityDataFormat)
   Else
      txtAQty = Format(cValue(iIndex, 1), ES_QuantityDataFormat)
   End If
   optLots.Value = Val(sParts(iIndex, 4))
   bCycleLots = GetCycleLots(sParts(iIndex, 1))
   If cValue(iIndex, 3) = 1 Then
      optRec.Value = vbChecked
      txtAQty.Enabled = False
      cmdAdj.Enabled = False
   Else
      optRec.Value = vbUnchecked
      If bCycleLots = 0 Then cmdAdj.Enabled = True
   End If
   If bCycleLots = 1 Then
      txtAQty.Enabled = False
      cmdAdj.Enabled = False
      For iRow = 1 To iTotalLots
         If iRow > 12 Then Exit For
         If Val(vCycleLots(iRow, 3)) = 0 Then
            z1(12).Enabled = True
            z1(14).Visible = True
            lblLot(iRow) = vCycleLots(iRow, 2)
            lblLot(iRow).ToolTipText = vCycleLots(iRow, 0)
            lblLotRem(iRow) = vCycleLots(iRow, 1)
            txtLotQty(iRow) = vCycleLots(iRow, 1)
            cmdLotRec(iRow).Enabled = True
            txtLotQty(iRow).Enabled = True
            txtLotQty(iRow).BackColor = Es_TextBackColor
         Else
            lblLot(iRow) = vCycleLots(iRow, 2)
            lblLot(iRow).ToolTipText = vCycleLots(iRow, 0)
            lblLotRem(iRow) = vCycleLots(iRow, 3)
            txtLotQty(iRow) = vCycleLots(iRow, 3)
         End If
      Next
      TotalLotQuantity
      txtLotQty(1).SetFocus
   Else
      If cValue(iIndex, 3) = 0 Then
         cmdAdj.Enabled = True
         txtAQty.Enabled = True
         txtAQty.SetFocus
      End If
   End If
   
End Sub


Private Sub optShow_Click()
   If cmdSel.Enabled = False Then FillList
   
End Sub

Private Sub txtAQty_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtAQty_LostFocus()
   txtAQty = CheckLen(txtAQty, 10)
   txtAQty = Format(Abs(Val(txtAQty)), ES_QuantityDataFormat)
   cValue(iIndex, 2) = Val(txtAQty)
   
End Sub



Private Sub CloseLotBoxes()
   Dim bRow As Byte
   Dim bTabIdx As Byte
   z1(12).Enabled = False
   z1(14).Visible = False
   bTabIdx = 7
   If txtLotQty(1).Enabled Then
      For bRow = 1 To 11
         bTabIdx = bTabIdx + 2
         lblLot(bRow) = ""
         lblLot(bRow).ToolTipText = "No Lot Recorded"
         txtLotQty(bRow).TabIndex = bTabIdx
         txtLotQty(bRow) = ""
         lblLotRem(bRow) = ""
         txtLotQty(bRow).Enabled = False
         txtLotQty(bRow).BackColor = Es_FormBackColor
         cmdLotRec(bRow).TabIndex = bTabIdx + 1
         cmdLotRec(bRow).Enabled = False
      Next
      bTabIdx = bTabIdx + 2
      lblLot(bRow) = ""
      lblLot(bRow).ToolTipText = "No Lot Recorded"
      txtLotQty(bRow).TabIndex = bTabIdx
      txtLotQty(bRow) = ""
      lblLotRem(bRow) = ""
      txtLotQty(bRow).Enabled = False
      txtLotQty(bRow).BackColor = Es_FormBackColor
      cmdLotRec(bRow).TabIndex = bTabIdx + 1
      cmdLotRec(bRow).Enabled = False
   End If
   
End Sub

Private Function GetCycleLots(CyclePart As String) As Byte
   Dim RdoClots As ADODB.Recordset
   Erase vCycleLots
   iTotalLots = 0
   GetCycleLots = 0
   If optLots.Value = vbChecked Then
      sSql = "SELECT CLLOTNUMBER,CLLOTREMAININGQTY,LOTNUMBER,CLLOTADJUSTQTY,LOTUSERLOTID " _
             & "FROM CcltTable,LohdTable WHERE CLLOTNUMBER=LOTNUMBER AND " _
             & "(CLREF='" & cmbCid & "' AND " _
             & "CLPARTREF='" & CyclePart & "' AND CLLOTNUMBER<>'')"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoClots, ES_FORWARD)
      If bSqlRows Then
         With RdoClots
            Do Until .EOF
               iTotalLots = iTotalLots + 1
               vCycleLots(iTotalLots, 0) = "" & Trim(!CLLOTNUMBER)
               vCycleLots(iTotalLots, 1) = Format(!CLLOTREMAININGQTY, "########.000")
               vCycleLots(iTotalLots, 2) = "" & Trim(!LOTUSERLOTID)
               vCycleLots(iTotalLots, 3) = Format(!CLLOTADJUSTQTY, "########.000")
               .MoveNext
            Loop
            ClearResultSet RdoClots
         End With
         GetCycleLots = 1
      End If
   End If
   Set RdoClots = Nothing
   
   
End Function

Private Sub txtLotQty_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtLotQty_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub


Private Sub txtLotQty_LostFocus(Index As Integer)
   txtLotQty(Index) = CheckLen(txtLotQty(Index), 9)
   txtLotQty(Index) = Format(Val(txtLotQty(Index)), "########.000")
   
End Sub



Private Function GetNextDate() As Variant
   Dim RdoDate As ADODB.Recordset
   Dim iFrequency As Integer
   Dim dDate As Date
   
   On Error Resume Next
   dDate = Format(txtPlan, "mm/dd/yy")
   sSql = "SELECT COABCROW,COABCCODE,COABCFREQUENCY " _
          & "FROM CabcTable WHERE COABCCODE='" & lblCabc & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDate, ES_FORWARD)
   If bSqlRows Then iFrequency = RdoDate!COABCFREQUENCY
   GetNextDate = Format(dDate + iFrequency, "mm/dd/yy")
   Set RdoDate = Nothing
   
End Function


Private Function InsertActivity(Adjustment As Currency, Decrease As Byte) As Long
   Dim vAdate As Variant
   GetAccounts sParts(iIndex, 1)
   lCOUNTER = lCOUNTER + 1
   InsertActivity = lCOUNTER
   Adjustment = Abs(Adjustment)
   vAdate = Format(GetServerDateTime(), "mm/dd/yy hh:mm")
   If Decrease = 1 Then
      'Reduce Inventory Activity
      sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2,INPDATE," _
             & "INADATE,INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT,INNUMBER,INUSER) " _
             & "VALUES(30,'" & sParts(iIndex, 1) & "','ABC Cycle Count',''," _
             & "'" & txtPlan & "','" & txtPlan _
             & "',-" & Adjustment & ",-" & Adjustment & "," & cValue(iIndex, 0) _
             & ",'" & sCreditAcct & "','" & sDebitAcct & "'," & InsertActivity & ",'" & sInitials & "')"
   Else
      'Increase Inventory Activity
      sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2,INPDATE," _
             & "INADATE,INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT,INNUMBER,INUSER) " _
             & "VALUES(30,'" & sParts(iIndex, 1) & "','ABC Cycle Count',''," _
             & "'" & txtPlan & "','" & txtPlan _
             & "'," & Adjustment & "," & Adjustment & "," & cValue(iIndex, 0) _
             & ",'" & sCreditAcct & "','" & sDebitAcct & "'," & InsertActivity & ",'" & sInitials & "')"
   End If
   clsADOCon.ExecuteSQL sSql
   UpdateWipColumns lCOUNTER
   
End Function

'Total Lot Quantities and test, then close the lots

Private Sub TotalLotQuantity(Optional CloseLots As Boolean)
   Dim bByte As Byte
   Dim bRow As Byte
   Dim cActivityQty As Currency
   Dim cAdjustmentQty As Currency
   Dim cTotalLots As Currency
   
   For bRow = 1 To 11
      cTotalLots = cTotalLots + Val(txtLotQty(bRow))
   Next
   cTotalLots = cTotalLots + Val(txtLotQty(bRow))
   txtAQty = Format(cTotalLots, ES_QuantityDataFormat)
   If CloseLots Then
      For bRow = 1 To 11
         If cmdLotRec(bRow).Enabled = True Then bByte = 1
      Next
      If cmdLotRec(bRow).Enabled = True Then bByte = 1
      If bByte = 0 Then
         On Error Resume Next
         clsADOCon.ADOErrNum = 0
         clsADOCon.BeginTrans
         'Close and clear the parts
         sSql = "UPDATE PartTable SET PANEXTCYCLEDATE='" & vNextDate _
                & "',PAQOH=" & cTotalLots & ",PALOTQTYREMAINING=" _
                & cTotalLots & " WHERE PARTREF='" _
                & sParts(iIndex, 1) & "'"
         clsADOCon.ExecuteSQL sSql
         
         sSql = "UPDATE CcitTable SET CIACTUALQOH=" & cTotalLots _
                & ",CIRECONCILEDDATE='" & Format(ES_SYSDATE, "mm/dd/yy") _
                & "',CIRECONCILED=1 WHERE (CIREF='" & cmbCid & "' AND " _
                & "CIPARTREF='" & sParts(iIndex, 1) & "')"
         clsADOCon.ExecuteSQL sSql
         
         cActivityQty = GetActivityQuantity(sParts(iIndex, 1))
         If cActivityQty <> cTotalLots Then
            'Inventory
            If cActivityQty > cTotalLots Then
               'Reduce Inventory Activity
               cAdjustmentQty = cActivityQty - cTotalLots
               lCOUNTER = InsertActivity(cAdjustmentQty, 1)
            Else
               'Increase Inventory Activity
               cAdjustmentQty = cTotalLots - cAdjustmentQty
               lCOUNTER = InsertActivity(cAdjustmentQty, 0)
            End If
         End If
         If clsADOCon.ADOErrNum = 0 Then
            clsADOCon.CommitTrans
            MarkReconciled
         Else
            clsADOCon.RollbackTrans
            clsADOCon.ADOErrNum = 0
         End If
      End If
   End If
End Sub

Private Sub MarkReconciled()
   Dim rdoRec As ADODB.Recordset
   Dim bByte As Byte
   
   clsADOCon.ADOErrNum = 0
   On Error Resume Next
   sSql = "SELECT CIREF,CIPARTREF,CIRECONCILED FROM CcitTable WHERE " _
          & "(CIREF='" & cmbCid & "' AND CIRECONCILED=0)"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoRec, ES_FORWARD)
   Set rdoRec = Nothing
   If bSqlRows Then Exit Sub
   
   sSql = "UPDATE CchdTable SET CCUPDATEDDATE='" & Format(ES_SYSDATE, "mm/dd/yy") _
          & "',CCUPDATED=1 WHERE CCREF='" & cmbCid & "'"
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then MsgBox Trim(cmbCid) & " Has Been Reconciled.", _
            vbInformation, Caption
   FillCombo
   On Error Resume Next
   cmdNxt.SetFocus
   
End Sub

Private Function CheckReconciled() As Byte
   Dim rdoRec As ADODB.Recordset
   sSql = "SELECT CIREF,CIRECONCILED FROM CcitTable WHERE (CIREF='" _
          & cmbCid & "' AND CIRECONCILED=0) "
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoRec, ES_FORWARD)
   CheckReconciled = bSqlRows
   Set rdoRec = Nothing
   
End Function

Private Sub txtPlan_DropDown()
   ShowCalendar Me
   
End Sub
