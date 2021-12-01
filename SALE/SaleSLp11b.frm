VERSION 5.00
Begin VB.Form SaleSLp11b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer Sales Analysis"
   ClientHeight    =   9105
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   9180
   Icon            =   "SaleSLp11b.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   3720
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   9480
      Width           =   1065
   End
   Begin VB.CommandButton optPrn 
      DownPicture     =   "SaleSLp11b.frx":030A
      Height          =   415
      Left            =   7800
      Picture         =   "SaleSLp11b.frx":0494
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Print This Form"
      Top             =   120
      Width           =   415
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLp11b.frx":0A26
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.Label Legend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Legend"
      Height          =   255
      Index           =   0
      Left            =   7440
      TabIndex        =   92
      Top             =   1320
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Legend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Legend"
      Height          =   255
      Index           =   1
      Left            =   7440
      TabIndex        =   91
      Top             =   1680
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Legend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Legend"
      Height          =   255
      Index           =   2
      Left            =   7440
      TabIndex        =   90
      Top             =   2040
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Legend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Legend"
      Height          =   255
      Index           =   3
      Left            =   7440
      TabIndex        =   89
      Top             =   2400
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Legend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Legend"
      Height          =   255
      Index           =   4
      Left            =   7440
      TabIndex        =   88
      Top             =   2760
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Legend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Legend"
      Height          =   255
      Index           =   5
      Left            =   7440
      TabIndex        =   87
      Top             =   3120
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Legend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Legend"
      Height          =   255
      Index           =   6
      Left            =   7440
      TabIndex        =   86
      Top             =   3480
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Legend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Legend"
      Height          =   255
      Index           =   7
      Left            =   7440
      TabIndex        =   85
      Top             =   3840
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Legend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Legend"
      Height          =   255
      Index           =   8
      Left            =   7440
      TabIndex        =   84
      Top             =   4200
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Legend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Legend"
      Height          =   255
      Index           =   9
      Left            =   7440
      TabIndex        =   83
      Top             =   4560
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Legend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Legend"
      Height          =   255
      Index           =   10
      Left            =   7440
      TabIndex        =   82
      Top             =   4920
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Legend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Legend"
      Height          =   255
      Index           =   11
      Left            =   7440
      TabIndex        =   81
      Top             =   5280
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Legend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Legend"
      Height          =   255
      Index           =   12
      Left            =   7440
      TabIndex        =   80
      Top             =   5640
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Legend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Legend"
      Height          =   255
      Index           =   13
      Left            =   7440
      TabIndex        =   79
      Top             =   6000
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Legend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Legend"
      Height          =   255
      Index           =   14
      Left            =   7440
      TabIndex        =   78
      Top             =   6360
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Legend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Legend"
      Height          =   255
      Index           =   15
      Left            =   7440
      TabIndex        =   77
      Top             =   6720
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Legend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Legend"
      Height          =   255
      Index           =   16
      Left            =   7440
      TabIndex        =   76
      Top             =   7080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Legend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Legend"
      Height          =   255
      Index           =   17
      Left            =   7440
      TabIndex        =   75
      Top             =   7440
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Legend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Legend"
      Height          =   255
      Index           =   18
      Left            =   7440
      TabIndex        =   74
      Top             =   7800
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Legend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Legend"
      Height          =   255
      Index           =   19
      Left            =   7440
      TabIndex        =   73
      Top             =   8160
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label TotalSales 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   252
      Left            =   7320
      TabIndex        =   72
      ToolTipText     =   "Total All Selected Customers"
      Top             =   8520
      Width           =   852
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Selected Customers"
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   71
      ToolTipText     =   "From Work Center Calendars"
      Top             =   8520
      Width           =   2772
   End
   Begin VB.Label Sales 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   19
      Left            =   8280
      TabIndex        =   69
      ToolTipText     =   "Total Sales For This Customer"
      Top             =   8160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label SalesPercentage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   19
      Left            =   1680
      TabIndex        =   68
      Top             =   8160
      Visible         =   0   'False
      Width           =   5604
   End
   Begin VB.Label Customer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   19
      Left            =   240
      TabIndex        =   67
      Top             =   8160
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label Sales 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   18
      Left            =   8280
      TabIndex        =   66
      ToolTipText     =   "Total Sales For This Customer"
      Top             =   7800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label SalesPercentage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   18
      Left            =   1680
      TabIndex        =   65
      Top             =   7800
      Visible         =   0   'False
      Width           =   5604
   End
   Begin VB.Label Customer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   18
      Left            =   240
      TabIndex        =   64
      Top             =   7800
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label Sales 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   17
      Left            =   8280
      TabIndex        =   63
      ToolTipText     =   "Total Sales For This Customer"
      Top             =   7440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label SalesPercentage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   17
      Left            =   1680
      TabIndex        =   62
      Top             =   7440
      Visible         =   0   'False
      Width           =   5604
   End
   Begin VB.Label Customer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   17
      Left            =   240
      TabIndex        =   61
      Top             =   7440
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label Sales 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   16
      Left            =   8280
      TabIndex        =   60
      ToolTipText     =   "Total Sales For This Customer"
      Top             =   7080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label SalesPercentage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   16
      Left            =   1680
      TabIndex        =   59
      Top             =   7080
      Visible         =   0   'False
      Width           =   5604
   End
   Begin VB.Label Customer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   16
      Left            =   240
      TabIndex        =   58
      Top             =   7080
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label Sales 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   8280
      TabIndex        =   57
      ToolTipText     =   "Total Sales For This Customer"
      Top             =   6720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label SalesPercentage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   15
      Left            =   1680
      TabIndex        =   56
      Top             =   6720
      Visible         =   0   'False
      Width           =   5604
   End
   Begin VB.Label Customer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   15
      Left            =   240
      TabIndex        =   55
      Top             =   6720
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label Sales 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   8280
      TabIndex        =   54
      ToolTipText     =   "Total Sales For This Customer"
      Top             =   6360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label SalesPercentage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   14
      Left            =   1680
      TabIndex        =   53
      Top             =   6360
      Visible         =   0   'False
      Width           =   5604
   End
   Begin VB.Label Customer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   14
      Left            =   240
      TabIndex        =   52
      Top             =   6360
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label Sales 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   8280
      TabIndex        =   51
      ToolTipText     =   "Total Sales For This Customer"
      Top             =   6000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label SalesPercentage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   13
      Left            =   1680
      TabIndex        =   50
      Top             =   6000
      Visible         =   0   'False
      Width           =   5604
   End
   Begin VB.Label Customer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   13
      Left            =   240
      TabIndex        =   49
      Top             =   6000
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label Sales 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   8280
      TabIndex        =   48
      ToolTipText     =   "Total Sales For This Customer"
      Top             =   5640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label SalesPercentage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   12
      Left            =   1680
      TabIndex        =   47
      Top             =   5640
      Visible         =   0   'False
      Width           =   5604
   End
   Begin VB.Label Customer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   12
      Left            =   240
      TabIndex        =   46
      Top             =   5640
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label Sales 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   8280
      TabIndex        =   45
      ToolTipText     =   "Total Sales For This Customer"
      Top             =   5280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label SalesPercentage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   11
      Left            =   1680
      TabIndex        =   44
      Top             =   5280
      Visible         =   0   'False
      Width           =   5604
   End
   Begin VB.Label Customer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   11
      Left            =   240
      TabIndex        =   43
      Top             =   5280
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label Sales 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   8280
      TabIndex        =   42
      ToolTipText     =   "Total Sales For This Customer"
      Top             =   4920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label SalesPercentage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   10
      Left            =   1680
      TabIndex        =   41
      Top             =   4920
      Visible         =   0   'False
      Width           =   5604
   End
   Begin VB.Label Customer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   10
      Left            =   240
      TabIndex        =   40
      Top             =   4920
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label Sales 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   8280
      TabIndex        =   39
      ToolTipText     =   "Total Sales For This Customer"
      Top             =   4560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label SalesPercentage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   9
      Left            =   1680
      TabIndex        =   38
      Top             =   4560
      Visible         =   0   'False
      Width           =   5604
   End
   Begin VB.Label Customer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   9
      Left            =   240
      TabIndex        =   37
      Top             =   4560
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label Sales 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   8280
      TabIndex        =   36
      ToolTipText     =   "Total Sales For This Customer"
      Top             =   4200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label SalesPercentage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   8
      Left            =   1680
      TabIndex        =   35
      Top             =   4200
      Visible         =   0   'False
      Width           =   5604
   End
   Begin VB.Label Customer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   8
      Left            =   240
      TabIndex        =   34
      Top             =   4200
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label Sales 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   8280
      TabIndex        =   33
      ToolTipText     =   "Total Sales For This Customer"
      Top             =   3840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label SalesPercentage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   7
      Left            =   1680
      TabIndex        =   32
      Top             =   3840
      Visible         =   0   'False
      Width           =   5604
   End
   Begin VB.Label Customer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   7
      Left            =   240
      TabIndex        =   31
      Top             =   3840
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label Sales 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   8280
      TabIndex        =   30
      ToolTipText     =   "Total Sales For This Customer"
      Top             =   3480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label SalesPercentage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   6
      Left            =   1680
      TabIndex        =   29
      Top             =   3480
      Visible         =   0   'False
      Width           =   5604
   End
   Begin VB.Label Customer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   6
      Left            =   240
      TabIndex        =   28
      Top             =   3480
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label Sales 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   8280
      TabIndex        =   27
      ToolTipText     =   "Total Sales For This Customer"
      Top             =   3120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label SalesPercentage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   5
      Left            =   1680
      TabIndex        =   26
      Top             =   3120
      Visible         =   0   'False
      Width           =   5604
   End
   Begin VB.Label Customer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   5
      Left            =   240
      TabIndex        =   25
      Top             =   3120
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label Sales 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   8280
      TabIndex        =   24
      ToolTipText     =   "Total Sales For This Customer"
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label SalesPercentage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   4
      Left            =   1680
      TabIndex        =   23
      Top             =   2760
      Visible         =   0   'False
      Width           =   5604
   End
   Begin VB.Label Customer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   4
      Left            =   240
      TabIndex        =   22
      Top             =   2760
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label Sales 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   8280
      TabIndex        =   21
      ToolTipText     =   "Total Sales For This Customer"
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label SalesPercentage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   3
      Left            =   1680
      TabIndex        =   20
      Top             =   2400
      Visible         =   0   'False
      Width           =   5604
   End
   Begin VB.Label Customer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   3
      Left            =   240
      TabIndex        =   19
      Top             =   2400
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label Sales 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   8280
      TabIndex        =   18
      ToolTipText     =   "Total Sales For This Customer"
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label SalesPercentage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   2
      Left            =   1680
      TabIndex        =   17
      Top             =   2040
      Visible         =   0   'False
      Width           =   5604
   End
   Begin VB.Label Customer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   16
      Top             =   2040
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label Sales 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   8280
      TabIndex        =   15
      ToolTipText     =   "Total Sales For This Customer"
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label SalesPercentage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   1
      Left            =   1680
      TabIndex        =   14
      Top             =   1680
      Visible         =   0   'False
      Width           =   5604
   End
   Begin VB.Label Customer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label Sales 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   8280
      TabIndex        =   12
      ToolTipText     =   "Total Sales For This Customer"
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label SalesPercentage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   0
      Left            =   1680
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   5600
   End
   Begin VB.Label Customer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Top 20 Customers Maximum"
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   9
      ToolTipText     =   "From Work Center Calendars"
      Top             =   480
      Width           =   2652
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From "
      Height          =   252
      Index           =   2
      Left            =   2760
      TabIndex        =   8
      ToolTipText     =   "From Work Center Calendars"
      Top             =   840
      Width           =   612
   End
   Begin VB.Label lblStart 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   252
      Left            =   3360
      TabIndex        =   7
      ToolTipText     =   "Always Today"
      Top             =   840
      Width           =   732
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Though"
      Height          =   252
      Index           =   3
      Left            =   4200
      TabIndex        =   6
      ToolTipText     =   "From Work Center Calendars"
      Top             =   840
      Width           =   612
   End
   Begin VB.Label lblThrough 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   252
      Left            =   4920
      TabIndex        =   5
      ToolTipText     =   "Date Selected"
      Top             =   840
      Width           =   732
   End
   Begin VB.Label Customers 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   252
      Left            =   1320
      TabIndex        =   4
      ToolTipText     =   "Customer(s) Selected"
      Top             =   840
      Width           =   1212
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer(s)"
      Height          =   252
      Index           =   4
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "From Work Center Calendars"
      Top             =   840
      Width           =   852
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Sales Analysis (Bookings)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   7092
   End
End
Attribute VB_Name = "SaleSLp11b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'8/25/06 New
Option Explicit
Dim RdoChart As ADODB.Recordset

Const PercWidth = 5600
Dim bOnLoad As Byte
Dim iTotalCusts As String
Dim sTempTable As String

Dim sCustomers(1200, 3)

Private Sub cmdCan_Click()
   Form_Deactivate
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2130
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   If bOnLoad Then
      CreateTempTable
      bOnLoad = 0
   End If
   
End Sub

Private Sub Form_Deactivate()
   Unload Me
   
End Sub


Private Sub Form_Load()
   Dim iRow As Integer
   Me.BackColor = vbWhite
   Move 1000, 1000
   For iRow = 0 To 19
      SalesPercentage(iRow).Width = 10
      Sales(iRow).ToolTipText = "Total Sales/Bookings For This Customer"
   Next
   bOnLoad = 1
   
End Sub



Private Sub CreateTempTable()
   Dim sCust As String
   Dim sBeg As String
   Dim sEnd As String
   
   If Customers <> "ALL" Then sCust = Compress(Customers)
   sBeg = Format(lblStart, "mm-dd-yyyy")
   sEnd = Format(lblThrough, "mm-dd-yyyy")
   
   sTempTable = "A" & UCase$(Format(Now, "ddd")) & Right(Compress(GetNextLotNumber()), 8)
   sSql = "CREATE TABLE " & sTempTable _
          & " (Customer CHAR(10) NULL DEFAULT('')," _
          & "CustomerName CHAR(40) NULL DEFAULT(''), " _
          & "Sales INT NULL DEFAULT(0))"
   clsADOCon.ExecuteSQL sSql ' rdExecDirect
   
   sSql = "CREATE UNIQUE CLUSTERED INDEX cust_idx ON " & sTempTable & " " _
          & "(Customer) WITH FILLFACTOR=80"
   clsADOCon.ExecuteSQL sSql ' rdExecDirect
   iTotalCusts = -1
   sSql = "select DISTINCT SOCUST,CUNAME,ITSO FROM SohdTable,CustTable," _
          & "SoitTable WHERE (SONUMBER=ITSO AND ITBOOKDATE Between '" & sBeg & " 00:00' " _
          & "AND '" & sEnd & " 23:59' AND SOCUST=CUREF) AND SOCUST LIKE '" & sCust _
          & "%' ORDER BY SOCUST"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChart, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoChart
         Do Until .EOF
            Err.Clear
            clsADOCon.ADOErrNum = 0
            
            sSql = "INSERT INTO " & sTempTable & " " _
                   & "(Customer,CustomerName) " _
                   & "VALUES('" & !SOCUST & "','" & !CUNAME & "')"
            clsADOCon.ExecuteSQL sSql ' rdExecDirect
            If clsADOCon.ADOErrNum = 0 Then
               iTotalCusts = iTotalCusts + 1
               sCustomers(iTotalCusts, 0) = "" & Trim(!SOCUST)
            End If
            .MoveNext
         Loop
         ClearResultSet RdoChart
      End With
   Else
      MouseCursor 0
      MsgBox "No Data Was Found In The Time Frame.", vbInformation, Caption
      Unload Me
      Exit Sub
   End If
   CreateCustChart
   Exit Sub
DiaErr1:
   On Error Resume Next
   sSql = "DROP TABLE " & sTempTable & " "
   clsADOCon.ExecuteSQL sSql ' rdExecDirect
   MouseCursor 0
   MsgBox "Couldn't Create Temporary Table.", vbExclamation, Caption
   Unload Me
   
End Sub

Private Sub CreateCustChart()
   Dim iRow As Integer
   Dim cWidth As Currency
   Dim lTotalSales As Long
   Dim lCustSales(20) As Long
   
   'Query
   Dim sBeg As String
   Dim sEnd As String
   
   sBeg = Format(lblStart, "mm-dd-yyyy")
   sEnd = Format(lblThrough, "mm-dd-yyyy")
   
   On Error GoTo DiaErr1
   sProcName = "getchartsales"
   For iRow = 0 To iTotalCusts
      sSql = "SELECT SOCUST,SONUMBER,SUM(ITQTY*ITDOLLARS)AS CustSales FROM " _
             & "SohdTable,SoitTable WHERE (SONUMBER=ITSO AND ITCANCELED=0 AND " _
             & "ITBOOKDATE BETWEEN '" & sBeg & " 00:00' AND '" & sEnd & " 23:59' AND SOCUST='" _
             & sCustomers(iRow, 0) & "') GROUP BY SOCUST,SONUMBER "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoChart, ES_FORWARD)
      If bSqlRows Then
         With RdoChart
            Do Until .EOF
               sSql = "UPDATE " & sTempTable & " SET Sales=Sales + " _
                      & !CustSales & " WHERE Customer='" & sCustomers(iRow, 0) & "'"
               clsADOCon.ExecuteSQL sSql ' rdExecDirect
               .MoveNext
            Loop
            ClearResultSet RdoChart
         End With
      End If
   Next
   sSql = "DELETE FROM " & sTempTable & " Where (Sales<1 OR Sales IS NULL)"
   clsADOCon.ExecuteSQL sSql ' rdExecDirect
   
   Erase sCustomers
   iTotalCusts = -1
   sProcName = "limitcusts"
   sSql = "SELECT * FROM " & sTempTable & " ORDER BY Sales DESC"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChart, ES_FORWARD)
   If bSqlRows Then
      With RdoChart
         Do Until .EOF
            iTotalCusts = iTotalCusts + 1
            If iTotalCusts < 19 Then
               sCustomers(iTotalCusts, 0) = "" & Trim(!Customer)
               sCustomers(iTotalCusts, 1) = "" & Trim(!CustomerName)
               lCustSales(iTotalCusts) = !Sales
               lTotalSales = lTotalSales + !Sales
               Customer(iTotalCusts).Visible = True
               SalesPercentage(iTotalCusts).Visible = True
               Sales(iTotalCusts).Visible = True
               Legend(iTotalCusts).Visible = True
               Customer(iTotalCusts).Caption = sCustomers(iTotalCusts, 0)
               Customer(iTotalCusts).ToolTipText = sCustomers(iTotalCusts, 1)
               Sales(iTotalCusts).Caption = lCustSales(iTotalCusts)
            Else
               'Others
               sCustomers(19, 0) = "Others"
               sCustomers(19, 1) = "Other Customers"
               lCustSales(19) = lCustSales(19) + !Sales
               lTotalSales = lTotalSales + !Sales
               Customer(19).Visible = True
               SalesPercentage(19).Visible = True
               Sales(19).Visible = True
               Legend(19).Visible = True
               Customer(19).Caption = sCustomers(19, 0)
               Customer(19).ToolTipText = sCustomers(19, 1)
               Sales(19).Caption = lCustSales(19)
            End If
            If iTotalCusts > 40 Then Exit Do
            .MoveNext
         Loop
         ClearResultSet RdoChart
      End With
      If iTotalCusts > 19 Then iTotalCusts = 19
      z1(1).Top = Customer(iTotalCusts).Top + 440
      TotalSales.Top = Customer(iTotalCusts).Top + 440
      TotalSales.Caption = Format$(lTotalSales)
      If TotalSales.Top < 8520 Then Me.Height = TotalSales.Top + 1770
   End If
   'Build Chart
   sProcName = "buildchart"
   For iRow = 0 To iTotalCusts
      If iRow < 19 Then
         cWidth = lCustSales(iRow) / lTotalSales
         SalesPercentage(iRow).Width = PercWidth * cWidth
         'SalesPercentage(iRow).Caption = Format$(cWidth, "##0%")
         SalesPercentage(iRow).ToolTipText = Format$(cWidth, "##0%")
         Legend(iRow).Caption = Format$(cWidth, "##0%")
         Legend(iRow).Left = (SalesPercentage(iRow).Left + PercWidth * cWidth) + 25
      Else
         'Others last boxes
         cWidth = lCustSales(19) / lTotalSales
         SalesPercentage(19).Width = PercWidth * cWidth
         'SalesPercentage(19).Caption = Format$(cWidth, "##0%")
         SalesPercentage(19).ToolTipText = Format$(cWidth, "##0%")
         Legend(19).Caption = Format$(cWidth, "##0%")
         Legend(19).Left = (SalesPercentage(19).Left + PercWidth * cWidth) + 25
      End If
   Next
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   On Error Resume Next
   sSql = "DROP TABLE " & sTempTable & " "
   clsADOCon.ExecuteSQL sSql ' rdExecDirect
   
   sProcName = "CCC " & sProcName
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error Resume Next
   sSql = "DROP TABLE " & sTempTable & " "
   clsADOCon.ExecuteSQL sSql ' rdExecDirect
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set RdoChart = Nothing
   Set SaleSLp11b = Nothing
   
End Sub


Private Sub optPrn_Click()
   Dim bResponse As Byte
   bResponse = MsgBox("Print This Form?", ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      cmdHlp.Visible = False
      optPrn.Visible = False
      Sleep 500
      PrintForm
   End If
   optPrn.Visible = True
   cmdHlp.Visible = True
   optPrn.Value = False
   
End Sub
