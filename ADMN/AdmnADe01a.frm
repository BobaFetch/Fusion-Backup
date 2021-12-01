VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form AdmnADe01a 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Settings"
   ClientHeight    =   10875
   ClientLeft      =   1440
   ClientTop       =   1485
   ClientWidth     =   8475
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "AdmnADe01a.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "AdmnADe01a.frx":000C
   ScaleHeight     =   10875
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   Begin VB.Frame tabFrame 
      Caption         =   "Lots"
      Height          =   5052
      Index           =   6
      Left            =   9060
      TabIndex        =   210
      Top             =   1140
      Width           =   7572
      Begin VB.CheckBox chkAbbreviatedLotNumbers 
         Height          =   255
         Left            =   3180
         TabIndex        =   328
         ToolTipText     =   "On Puchout logout."
         Top             =   3480
         Width           =   375
      End
      Begin VB.CheckBox chkSheetInventory 
         Caption         =   "Check2"
         Height          =   195
         Left            =   3180
         TabIndex        =   316
         Top             =   3120
         Width           =   255
      End
      Begin VB.Frame z3 
         BorderStyle     =   0  'None
         Height          =   612
         Index           =   3
         Left            =   2640
         TabIndex        =   211
         Top             =   960
         Width           =   2292
         Begin VB.OptionButton optFifo 
            Caption         =   "FIFO"
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optLifo 
            Caption         =   "LIFO"
            Height          =   255
            Left            =   1080
            TabIndex        =   66
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox txtAbcHigh 
         Height          =   285
         Left            =   3180
         TabIndex        =   69
         Tag             =   "1"
         ToolTipText     =   "Standard High (Max) Value Of Codes"
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txtAbcLow 
         Height          =   285
         Left            =   3180
         TabIndex        =   68
         Tag             =   "1"
         ToolTipText     =   "Standard Low (Min) Value Of Codes"
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox TxtAbcCount 
         Height          =   285
         Left            =   3180
         TabIndex        =   67
         Tag             =   "1"
         ToolTipText     =   "How Often Are Counters Counting (In Days) 1-360"
         Top             =   2040
         Width           =   735
      End
      Begin VB.CheckBox optLots 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   2760
         TabIndex        =   64
         ToolTipText     =   "Verify Accounts And Journals Through The System"
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Use Abbreviated User Lot Numbers"
         Height          =   255
         Left            =   240
         TabIndex        =   329
         ToolTipText     =   "On Punchout Auto Logout from Current Job"
         Top             =   3480
         Width           =   2775
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Use Sheet Inventory Features"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   128
         Left            =   240
         TabIndex        =   317
         Top             =   3120
         Width           =   2280
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Standard High Limit"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   62
         Left            =   240
         TabIndex        =   218
         ToolTipText     =   "Standard High (Max) Value Of Codes"
         Top             =   2760
         Width           =   1860
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Standard Low Limit"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   61
         Left            =   240
         TabIndex        =   217
         ToolTipText     =   "Standard Low (Min) Value Of Codes"
         Top             =   2400
         Width           =   2340
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Persons Counting (In Days)"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   60
         Left            =   240
         TabIndex        =   216
         ToolTipText     =   "How Often Are Counters Counting (In Days) 1-360"
         Top             =   2040
         Width           =   2340
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cycle Counting:"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   59
         Left            =   240
         TabIndex        =   215
         Top             =   1680
         Width           =   2580
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Lot Tracking:"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   58
         Left            =   240
         TabIndex        =   214
         Top             =   360
         Width           =   2580
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory Is Relieved Using"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   49
         Left            =   240
         TabIndex        =   213
         Top             =   1200
         Width           =   2580
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Lot Tracking Is Turned On"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   52
         Left            =   240
         TabIndex        =   212
         Top             =   720
         Width           =   2580
      End
   End
   Begin VB.Frame tabFrame 
      Caption         =   "Inventory"
      Height          =   5775
      Index           =   1
      Left            =   9780
      TabIndex        =   106
      Top             =   2820
      Width           =   7572
      Begin VB.ComboBox txtMos 
         Height          =   315
         Left            =   1920
         TabIndex        =   28
         Tag             =   "3"
         Top             =   4680
         Width           =   1935
      End
      Begin VB.ComboBox txtCogs 
         Height          =   315
         Left            =   1920
         TabIndex        =   18
         Tag             =   "3"
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdWip 
         Caption         =   "&Inventory"
         Height          =   315
         Left            =   6120
         TabIndex        =   107
         TabStop         =   0   'False
         ToolTipText     =   "Enter/Revise Inventory, CGS Accounts And WIP Accounts"
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox txtFgi 
         Height          =   315
         Left            =   1920
         TabIndex        =   27
         Tag             =   "3"
         Top             =   4320
         Width           =   1935
      End
      Begin VB.ComboBox txtArev8 
         Height          =   315
         Left            =   1920
         TabIndex        =   26
         Tag             =   "3"
         Top             =   3960
         Width           =   1935
      End
      Begin VB.ComboBox txtArev7 
         Height          =   315
         Left            =   1920
         TabIndex        =   25
         Tag             =   "3"
         Top             =   3600
         Width           =   1935
      End
      Begin VB.ComboBox txtArev6 
         Height          =   315
         Left            =   1920
         TabIndex        =   24
         Tag             =   "3"
         Top             =   3240
         Width           =   1935
      End
      Begin VB.ComboBox txtArev5 
         Height          =   315
         Left            =   1920
         TabIndex        =   23
         Tag             =   "3"
         Top             =   2880
         Width           =   1935
      End
      Begin VB.ComboBox txtArev4 
         Height          =   315
         Left            =   1920
         TabIndex        =   22
         Tag             =   "3"
         Top             =   2520
         Width           =   1935
      End
      Begin VB.ComboBox txtArev3 
         Height          =   315
         Left            =   1920
         TabIndex        =   21
         Tag             =   "3"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.ComboBox txtArev2 
         Height          =   315
         Left            =   1920
         TabIndex        =   20
         Tag             =   "3"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.ComboBox txtArev1 
         Height          =   315
         Left            =   1920
         TabIndex        =   19
         Tag             =   "3"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label lblMos 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   3960
         TabIndex        =   133
         Top             =   4680
         Width           =   2592
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Material Over/Short"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   75
         Left            =   240
         TabIndex        =   132
         Top             =   4680
         Width           =   3000
      End
      Begin VB.Label lblCogs 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   3960
         TabIndex        =   131
         Top             =   840
         Width           =   2592
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Default COG Sold"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   41
         Left            =   240
         TabIndex        =   130
         Top             =   840
         Width           =   2100
      End
      Begin VB.Label lblFgi 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   3960
         TabIndex        =   129
         Top             =   4320
         Width           =   2592
      End
      Begin VB.Label lblArev8 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   3960
         TabIndex        =   128
         Top             =   3960
         Width           =   2592
      End
      Begin VB.Label lblArev7 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   3960
         TabIndex        =   127
         Top             =   3600
         Width           =   2592
      End
      Begin VB.Label lblArev6 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   3960
         TabIndex        =   126
         Top             =   3240
         Width           =   2592
      End
      Begin VB.Label lblArev5 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   3960
         TabIndex        =   125
         Top             =   2880
         Width           =   2592
      End
      Begin VB.Label lblArev4 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   3960
         TabIndex        =   124
         Top             =   2520
         Width           =   2592
      End
      Begin VB.Label lblArev3 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   3960
         TabIndex        =   123
         Top             =   2160
         Width           =   2592
      End
      Begin VB.Label lblArev2 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   3960
         TabIndex        =   122
         Top             =   1800
         Width           =   2592
      End
      Begin VB.Label lblArev1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   3960
         TabIndex        =   121
         Top             =   1440
         Width           =   2592
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory/Expense And Cost Of Goods Accounts:"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   40
         Left            =   2520
         TabIndex        =   120
         Top             =   360
         Width           =   4020
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Finished Goods "
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   38
         Left            =   240
         TabIndex        =   119
         Top             =   4320
         Width           =   3000
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "8. Project"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   19
         Left            =   240
         TabIndex        =   118
         Top             =   3960
         Width           =   3000
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "7. OS Service"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   18
         Left            =   240
         TabIndex        =   117
         Top             =   3600
         Width           =   3000
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "6. Service"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   17
         Left            =   240
         TabIndex        =   116
         Top             =   3240
         Width           =   3000
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "5. Expendable"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   16
         Left            =   240
         TabIndex        =   115
         Top             =   2880
         Width           =   3000
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "4. Raw Matl"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   15
         Left            =   240
         TabIndex        =   114
         Top             =   2520
         Width           =   3000
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "3. Base Assy"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   14
         Left            =   240
         TabIndex        =   113
         Top             =   2160
         Width           =   3000
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "2. Inter Assy"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   13
         Left            =   240
         TabIndex        =   112
         Top             =   1800
         Width           =   3000
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1. Top Assy"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   12
         Left            =   240
         TabIndex        =   111
         Top             =   1440
         Width           =   3000
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Revenue Accounts"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   9
         Left            =   1920
         TabIndex        =   110
         Top             =   1200
         Width           =   1860
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Part Type"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   8
         Left            =   240
         TabIndex        =   109
         Top             =   1200
         Width           =   1260
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Default Inventory Accounts:"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   7
         Left            =   240
         TabIndex        =   108
         Top             =   360
         Width           =   2100
      End
   End
   Begin VB.Frame tabFrame 
      Caption         =   "Engineering"
      Height          =   5052
      Index           =   10
      Left            =   9960
      TabIndex        =   264
      Top             =   3120
      Width           =   7572
      Begin VB.TextBox txtEngineeringRate 
         Height          =   285
         Left            =   2280
         TabIndex        =   318
         Tag             =   "1"
         Top             =   2160
         Width           =   840
      End
      Begin VB.CheckBox chkBOMSetQty 
         Caption         =   "____"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   2280
         TabIndex        =   306
         Top             =   1800
         Width           =   255
      End
      Begin VB.CheckBox chkRout 
         Caption         =   "____"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   2280
         TabIndex        =   267
         Top             =   720
         Width           =   255
      End
      Begin VB.CheckBox chkDocLst 
         Caption         =   "____"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   2280
         TabIndex        =   266
         Top             =   1440
         Width           =   255
      End
      Begin VB.CheckBox chkBOM 
         Caption         =   "____"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   2280
         TabIndex        =   265
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Engineering Labor Rate"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   101
         Left            =   480
         TabIndex        =   319
         Top             =   2160
         Width           =   1725
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BOM Setup Qty"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   125
         Left            =   480
         TabIndex        =   307
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rounting Changes"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   110
         Left            =   480
         TabIndex        =   271
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Document List"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   116
         Left            =   480
         TabIndex        =   270
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BOM changes"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   117
         Left            =   480
         TabIndex        =   269
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Configuration Management Security :"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   118
         Left            =   240
         TabIndex        =   268
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame tabFrame 
      Caption         =   "Labor"
      Height          =   5052
      Index           =   9
      Left            =   9600
      TabIndex        =   233
      Top             =   2460
      Width           =   7572
      Begin VB.CheckBox chkNoTCOvh 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   2040
         TabIndex        =   280
         ToolTipText     =   "Allows Prepackaging, But Not Shipped Items On Pack Slips"
         Top             =   2040
         Width           =   500
      End
      Begin VB.CheckBox optOHcost 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   2040
         TabIndex        =   262
         ToolTipText     =   "Allows Prepackaging, But Not Shipped Items On Pack Slips"
         Top             =   1560
         Width           =   500
      End
      Begin VB.ComboBox txtDefTime 
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   79
         Tag             =   "3"
         ToolTipText     =   "Used If No Account In The Work Center Or Employee At Time Entry"
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox txtDefLabor 
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   80
         Tag             =   "3"
         ToolTipText     =   "Distribution Without Payroll"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Do not share TC Overhead rate"
         ForeColor       =   &H00000000&
         Height          =   585
         Index           =   114
         Left            =   240
         TabIndex        =   281
         ToolTipText     =   "Allows Packaging, But Not Shipped Items On Pack Slips"
         Top             =   2040
         Width           =   1740
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Labor Defaults"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   90
         Left            =   240
         TabIndex        =   234
         ToolTipText     =   "Used If No Account In The Work Center Or Employee At Time Entry"
         Top             =   360
         Width           =   2580
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ignore posting Overhead cost to GL"
         ForeColor       =   &H00000000&
         Height          =   585
         Index           =   107
         Left            =   240
         TabIndex        =   263
         ToolTipText     =   "Allows Packaging, But Not Shipped Items On Pack Slips"
         Top             =   1560
         Width           =   1740
      End
      Begin VB.Label lblDefTime 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4080
         TabIndex        =   238
         Top             =   720
         Width           =   2592
      End
      Begin VB.Label lblDefLabor 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4080
         TabIndex        =   237
         Top             =   1080
         Width           =   2592
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Labor Account "
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   91
         Left            =   240
         TabIndex        =   236
         ToolTipText     =   "Distribution Without Payroll"
         Top             =   1080
         Width           =   2220
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Time Entry Account"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   89
         Left            =   240
         TabIndex        =   235
         ToolTipText     =   "Used If No Account In The Work Center Or Employee At Time Entry"
         Top             =   720
         Width           =   2580
      End
   End
   Begin VB.Frame tabFrame 
      Caption         =   "Sales"
      Height          =   4935
      Index           =   8
      Left            =   9480
      TabIndex        =   225
      Top             =   2160
      Width           =   7572
      Begin VB.CheckBox chkIgnQty 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   3000
         TabIndex        =   286
         Top             =   2880
         Width           =   255
      End
      Begin VB.TextBox txtWrnCusCL 
         Height          =   285
         Left            =   2280
         TabIndex        =   255
         Tag             =   "3"
         Text            =   "50000"
         ToolTipText     =   "Warn Customer Credit Limit"
         Top             =   2400
         Width           =   735
      End
      Begin VB.Frame z3 
         BorderStyle     =   0  'None
         Height          =   612
         Index           =   5
         Left            =   2520
         TabIndex        =   227
         Top             =   1200
         Width           =   3012
         Begin VB.OptionButton optSaleFor 
            Caption         =   "0.000"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   83
            Top             =   240
            Value           =   -1  'True
            Width           =   950
         End
         Begin VB.OptionButton optSaleFor 
            Caption         =   "0.0000"
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   84
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame z3 
         BorderStyle     =   0  'None
         Height          =   612
         Index           =   4
         Left            =   2520
         TabIndex        =   226
         Top             =   480
         Width           =   3012
         Begin VB.OptionButton optCom 
            Caption         =   "Cash Receipts"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   82
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton optCom 
            Caption         =   "Invoice"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   81
            Top             =   240
            Width           =   1044
         End
      End
      Begin VB.TextBox txtLastSon 
         Height          =   285
         Left            =   2280
         TabIndex        =   85
         Tag             =   "3"
         Text            =   "S00000"
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Option to Ignore Qty when repeat SO"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   119
         Left            =   240
         TabIndex        =   287
         ToolTipText     =   "Include Data collection module"
         Top             =   2880
         Width           =   2820
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Warn Customer Credit Limit"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   103
         Left            =   240
         TabIndex        =   254
         ToolTipText     =   "Allows Changes To The PO Line Item Price After Invoice (Does Not Change Actual)"
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Six Characters (Please Include Default Prefix)"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   88
         Left            =   3120
         TabIndex        =   232
         ToolTipText     =   "Allows Changes To The PO Line Item Price After Invoice (Does Not Change Actual)"
         Top             =   1920
         Width           =   4095
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Sales Order"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   80
         Left            =   240
         TabIndex        =   231
         ToolTipText     =   "Allows Changes To The PO Line Item Price After Invoice (Does Not Change Actual)"
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Selling Pricing Formats"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   79
         Left            =   240
         TabIndex        =   230
         ToolTipText     =   "Allows Changes To The PO Line Item Price After Invoice (Does Not Change Actual)"
         Top             =   1440
         Width           =   3492
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Commissions Based On"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   78
         Left            =   240
         TabIndex        =   229
         Top             =   720
         Width           =   2580
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Selling Defaults:"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   56
         Left            =   240
         TabIndex        =   228
         Top             =   360
         Width           =   2580
      End
   End
   Begin VB.Frame tabFrame 
      Caption         =   "Packing"
      Height          =   5052
      Index           =   7
      Left            =   9360
      TabIndex        =   219
      Top             =   1860
      Width           =   7572
      Begin VB.CheckBox chkEarlyLWarning 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   3120
         TabIndex        =   312
         ToolTipText     =   "Show pop-up when Shipped Early/Late"
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CheckBox chkOverShip 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   2280
         TabIndex        =   296
         ToolTipText     =   "Print Custom Shipping Labels"
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CheckBox cbCustomShipLabel 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   2280
         TabIndex        =   259
         ToolTipText     =   "Print Custom Shipping Labels"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CheckBox optInvPS 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   2280
         TabIndex        =   76
         ToolTipText     =   "Allows Prepackaging, But Not Shipped Items On Pack Slips"
         Top             =   2280
         Width           =   500
      End
      Begin VB.CheckBox optSoit 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   2280
         TabIndex        =   75
         ToolTipText     =   "Allows Prepackaging, But Not Shipped Items On Pack Slips"
         Top             =   1920
         Width           =   500
      End
      Begin VB.TextBox txtPackPrefix 
         Height          =   285
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   73
         Text            =   "PS"
         Top             =   1260
         Width           =   315
      End
      Begin VB.TextBox txtPsl 
         Height          =   285
         Left            =   2280
         MaxLength       =   8
         TabIndex        =   74
         Tag             =   "1"
         Text            =   "88888888"
         ToolTipText     =   "Set To The Last Packing Slip - Automatically Cycles"
         Top             =   1590
         Width           =   855
      End
      Begin VB.CommandButton cmdLock 
         Caption         =   "&Lock"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5880
         TabIndex        =   78
         ToolTipText     =   "Lock Transfer Invoice As Transfer Invoice"
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdInvoice 
         Caption         =   "&Find"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5040
         TabIndex        =   77
         ToolTipText     =   "Find Transfer Invoice"
         Top             =   960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtTransfer 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   72
         TabStop         =   0   'False
         Text            =   "000000"
         Top             =   960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox optTransfer 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   2280
         TabIndex        =   71
         ToolTipText     =   "Allows Inventory Transfers Within The Company"
         Top             =   960
         Width           =   500
      End
      Begin VB.CheckBox optPre 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   2280
         TabIndex        =   70
         ToolTipText     =   "Allows Prepackaging, But Not Shipped Items On Pack Slips"
         Top             =   600
         Width           =   500
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Early/Late Shipping Pop-up Warning "
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   126
         Left            =   240
         TabIndex        =   313
         ToolTipText     =   "Allows Packaging, But Not Shipped Items On Pack Slips"
         Top             =   3360
         Width           =   2700
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Allow Over Shipping Qty"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   123
         Left            =   240
         TabIndex        =   297
         ToolTipText     =   "Allows Packaging, But Not Shipped Items On Pack Slips"
         Top             =   3000
         Width           =   1980
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Custom Shipping Labels"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   106
         Left            =   240
         TabIndex        =   258
         ToolTipText     =   "Allows Packaging, But Not Shipped Items On Pack Slips"
         Top             =   2640
         Width           =   1980
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Set Invoice Num as PS"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   105
         Left            =   240
         TabIndex        =   257
         ToolTipText     =   "Allows Packaging, But Not Shipped Items On Pack Slips"
         Top             =   2280
         Width           =   1980
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Allow to add new SO Item"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   104
         Left            =   240
         TabIndex        =   256
         ToolTipText     =   "Allows Packaging, But Not Shipped Items On Pack Slips"
         Top             =   1920
         Width           =   1980
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "May be left blank to allow 8-digit PS numbers"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   95
         Left            =   3180
         TabIndex        =   243
         ToolTipText     =   "Set To The Last Packing Slip - Automatically Cycles"
         Top             =   1320
         Width           =   3540
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Packing Slip Prefix"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   93
         Left            =   240
         TabIndex        =   242
         ToolTipText     =   "Set To The Last Packing Slip - Automatically Cycles"
         Top             =   1290
         Width           =   1500
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Packing Slips And Shipping:"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   55
         Left            =   240
         TabIndex        =   224
         ToolTipText     =   "Allows Packaging, But Not Shipped Items On Pack Slips"
         Top             =   240
         Width           =   2580
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Packing Slip"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   87
         Left            =   240
         TabIndex        =   223
         ToolTipText     =   "Set To The Last Packing Slip - Automatically Cycles"
         Top             =   1620
         Width           =   1500
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Transfer Invoice"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   83
         Left            =   2880
         TabIndex        =   222
         ToolTipText     =   "Allows Inventory Transfers Within The Company"
         Top             =   960
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Allow Transfers"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   82
         Left            =   240
         TabIndex        =   221
         ToolTipText     =   "Allows Inventory Transfers Within The Company"
         Top             =   960
         Width           =   2220
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Allow Prepackaging "
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   53
         Left            =   240
         TabIndex        =   220
         ToolTipText     =   "Allows Packaging, But Not Shipped Items On Pack Slips"
         Top             =   600
         Width           =   1740
      End
   End
   Begin VB.Frame tabFrame 
      Caption         =   "Payables"
      Height          =   5052
      Index           =   3
      Left            =   9180
      TabIndex        =   158
      Top             =   1500
      Width           =   7572
      Begin VB.TextBox txtLinesPerStub 
         Height          =   285
         Left            =   3120
         MaxLength       =   2
         TabIndex        =   167
         Tag             =   "1"
         Text            =   "15"
         Top             =   3600
         Width           =   315
      End
      Begin VB.CheckBox optApFrDs 
         Caption         =   "____"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   3120
         TabIndex        =   166
         Top             =   3300
         Width           =   735
      End
      Begin VB.ComboBox txtCdds 
         Height          =   315
         Left            =   3120
         Sorted          =   -1  'True
         TabIndex        =   165
         Tag             =   "3"
         Top             =   2880
         Width           =   1935
      End
      Begin VB.ComboBox txtCdcc 
         Height          =   315
         Left            =   3120
         Sorted          =   -1  'True
         TabIndex        =   164
         Tag             =   "3"
         Top             =   2520
         Width           =   1935
      End
      Begin VB.ComboBox txtCdxc 
         Height          =   315
         Left            =   3120
         Sorted          =   -1  'True
         TabIndex        =   163
         Tag             =   "3"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.ComboBox txtPant 
         Height          =   315
         Left            =   3120
         Sorted          =   -1  'True
         TabIndex        =   162
         Tag             =   "3"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.ComboBox txtPatf 
         Height          =   315
         Left            =   3120
         Sorted          =   -1  'True
         TabIndex        =   161
         Tag             =   "3"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.ComboBox txtPatx 
         Height          =   315
         Left            =   3120
         Sorted          =   -1  'True
         TabIndex        =   160
         Tag             =   "3"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.ComboBox txtPaap 
         Height          =   315
         Left            =   3120
         Sorted          =   -1  'True
         TabIndex        =   159
         Tag             =   "3"
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum invoices per check stub"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   96
         Left            =   240
         TabIndex        =   244
         Top             =   3660
         Width           =   2715
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Calculate Discounts On AP Freight"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   47
         Left            =   240
         TabIndex        =   183
         Top             =   3300
         Width           =   2655
      End
      Begin VB.Label lblCdds 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   5160
         TabIndex        =   182
         Top             =   2880
         Width           =   2052
      End
      Begin VB.Label lblCdcc 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   5160
         TabIndex        =   181
         Top             =   2520
         Width           =   2052
      End
      Begin VB.Label lblCdxc 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   5160
         TabIndex        =   180
         Top             =   2160
         Width           =   2052
      End
      Begin VB.Label lblPant 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   5160
         TabIndex        =   179
         Top             =   1800
         Width           =   2052
      End
      Begin VB.Label lblPatf 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   5160
         TabIndex        =   178
         Top             =   1440
         Width           =   2052
      End
      Begin VB.Label lblPatx 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   5160
         TabIndex        =   177
         Top             =   1080
         Width           =   2052
      End
      Begin VB.Label lblPaap 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   5160
         TabIndex        =   176
         Top             =   720
         Width           =   2052
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Payables Discount Account"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   37
         Left            =   240
         TabIndex        =   175
         Top             =   2880
         Width           =   2700
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Computer Checks Account"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   36
         Left            =   240
         TabIndex        =   174
         Top             =   2520
         Width           =   2700
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "External Checks Account"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   35
         Left            =   240
         TabIndex        =   173
         Top             =   2160
         Width           =   2700
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Non-Taxable Freight Account"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   34
         Left            =   240
         TabIndex        =   172
         Top             =   1800
         Width           =   2700
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Taxable Freight Account"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   33
         Left            =   240
         TabIndex        =   171
         Top             =   1440
         Width           =   2700
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Taxes Payable Account"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   32
         Left            =   240
         TabIndex        =   170
         Top             =   1080
         Width           =   2700
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Payables Account"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   31
         Left            =   240
         TabIndex        =   169
         Top             =   720
         Width           =   2700
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Default Payables Journal Accounts:"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   30
         Left            =   240
         TabIndex        =   168
         Top             =   360
         Width           =   3492
      End
   End
   Begin VB.Frame tabFrame 
      Caption         =   "Receivables"
      Height          =   5535
      Index           =   2
      Left            =   8880
      TabIndex        =   134
      Top             =   840
      Width           =   7572
      Begin VB.TextBox txtBIVSOffSetAcc 
         Height          =   285
         Left            =   3240
         MaxLength       =   12
         TabIndex        =   276
         Tag             =   "3"
         Top             =   5040
         Width           =   1935
      End
      Begin VB.TextBox txtLastInvoiceNumber 
         Height          =   285
         Left            =   3240
         MaxLength       =   6
         TabIndex        =   40
         Tag             =   "3"
         Text            =   "000000"
         Top             =   4680
         Width           =   675
      End
      Begin VB.ComboBox txtTrns 
         Height          =   315
         Left            =   3240
         Sorted          =   -1  'True
         TabIndex        =   39
         Tag             =   "3"
         Top             =   4320
         Width           =   1935
      End
      Begin VB.ComboBox txtFtx 
         Height          =   315
         Left            =   3240
         Sorted          =   -1  'True
         TabIndex        =   38
         Tag             =   "3"
         Top             =   3960
         Width           =   1935
      End
      Begin VB.ComboBox txtSjnt 
         Height          =   315
         Left            =   3240
         Sorted          =   -1  'True
         TabIndex        =   37
         Tag             =   "3"
         Top             =   3600
         Width           =   1935
      End
      Begin VB.ComboBox txtSjtf 
         Height          =   315
         Left            =   3240
         Sorted          =   -1  'True
         TabIndex        =   36
         Tag             =   "3"
         Top             =   3240
         Width           =   1935
      End
      Begin VB.ComboBox txtSjtp 
         Height          =   315
         Left            =   3240
         Sorted          =   -1  'True
         TabIndex        =   35
         Tag             =   "3"
         Top             =   2880
         Width           =   1935
      End
      Begin VB.ComboBox txtSjar 
         Height          =   315
         Left            =   3240
         Sorted          =   -1  'True
         TabIndex        =   34
         Tag             =   "3"
         Top             =   2520
         Width           =   1935
      End
      Begin VB.ComboBox txtCrrv 
         Height          =   315
         Left            =   3240
         Sorted          =   -1  'True
         TabIndex        =   33
         Tag             =   "3"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.ComboBox txtCrcm 
         Height          =   315
         Left            =   3240
         Sorted          =   -1  'True
         TabIndex        =   32
         Tag             =   "3"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.ComboBox txtCrex 
         Height          =   315
         Left            =   3240
         Sorted          =   -1  'True
         TabIndex        =   31
         Tag             =   "3"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.ComboBox txtCrds 
         Height          =   315
         Left            =   3240
         Sorted          =   -1  'True
         TabIndex        =   30
         Tag             =   "3"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.ComboBox txtCrcs 
         Height          =   315
         Left            =   3240
         Sorted          =   -1  'True
         TabIndex        =   29
         Tag             =   "3"
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BIVS Offset Account"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   113
         Left            =   240
         TabIndex        =   277
         Top             =   5040
         Width           =   1935
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Invoice Number"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   94
         Left            =   240
         TabIndex        =   241
         Top             =   4680
         Width           =   1935
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Wire Transfer Account"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   11
         Left            =   240
         TabIndex        =   157
         Top             =   4320
         Width           =   3000
      End
      Begin VB.Label lblTrns 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   5280
         TabIndex        =   156
         Top             =   4320
         Width           =   2052
      End
      Begin VB.Label lblCrds 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   5280
         TabIndex        =   155
         Top             =   1080
         Width           =   2052
      End
      Begin VB.Label lblCrex 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   5280
         TabIndex        =   154
         Top             =   1440
         Width           =   2052
      End
      Begin VB.Label lblCrcm 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   5280
         TabIndex        =   153
         Top             =   1800
         Width           =   2052
      End
      Begin VB.Label lblCrrv 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   5280
         TabIndex        =   152
         Top             =   2160
         Width           =   2052
      End
      Begin VB.Label lblSjar 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   5280
         TabIndex        =   151
         Top             =   2520
         Width           =   2052
      End
      Begin VB.Label lblSjtp 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   5280
         TabIndex        =   150
         Top             =   2880
         Width           =   2052
      End
      Begin VB.Label lblSjtf 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   5280
         TabIndex        =   149
         Top             =   3240
         Width           =   2052
      End
      Begin VB.Label lblSjnt 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   5280
         TabIndex        =   148
         Top             =   3600
         Width           =   2052
      End
      Begin VB.Label lblFtx 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   5280
         TabIndex        =   147
         Top             =   3960
         Width           =   2052
      End
      Begin VB.Label lblCrcs 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   5280
         TabIndex        =   146
         Top             =   720
         Width           =   2052
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Federal Sales Tax"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   39
         Left            =   240
         TabIndex        =   145
         Top             =   3960
         Width           =   3000
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Journal Sales Tax Payable"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   29
         Left            =   240
         TabIndex        =   144
         Top             =   2880
         Width           =   3000
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Journal Non-Taxable Freight"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   28
         Left            =   240
         TabIndex        =   143
         Top             =   3600
         Width           =   3000
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Journal Taxable Freight Account"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   27
         Left            =   240
         TabIndex        =   142
         Top             =   3240
         Width           =   3000
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Journal AR Account"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   26
         Left            =   240
         TabIndex        =   141
         Top             =   2520
         Width           =   3000
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Receipts Revenue Account"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   25
         Left            =   240
         TabIndex        =   140
         Top             =   2160
         Width           =   3000
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Receipts Commission Account"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   24
         Left            =   240
         TabIndex        =   139
         Top             =   1800
         Width           =   3000
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Receipts Expense Account"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   23
         Left            =   240
         TabIndex        =   138
         Top             =   1440
         Width           =   3000
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Receipts Discount Account"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   22
         Left            =   240
         TabIndex        =   137
         Top             =   1080
         Width           =   3000
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Receipts Cash Account"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   21
         Left            =   240
         TabIndex        =   136
         Top             =   720
         Width           =   3000
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Default Sales/Cash Journal Accounts:"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   20
         Left            =   240
         TabIndex        =   135
         Top             =   360
         Width           =   3504
      End
   End
   Begin VB.Frame tabFrame 
      Caption         =   "Company"
      Height          =   7575
      Index           =   0
      Left            =   8760
      TabIndex        =   90
      Top             =   600
      Width           =   7572
      Begin VB.CheckBox chkPartSrch 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   5520
         TabIndex        =   298
         Top             =   6600
         Width           =   255
      End
      Begin VB.CheckBox cbDisAutoScan 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   5520
         TabIndex        =   294
         Top             =   6240
         Width           =   255
      End
      Begin VB.CheckBox cbDocumentLok 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   3000
         TabIndex        =   290
         Top             =   6240
         Width           =   255
      End
      Begin VB.CheckBox chkSumAcct 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   3000
         TabIndex        =   288
         Top             =   6600
         Width           =   375
      End
      Begin VB.CheckBox chkDataCol 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   3000
         TabIndex        =   282
         Top             =   5880
         Width           =   975
      End
      Begin VB.CheckBox cbHideObsoleteParts 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   1680
         TabIndex        =   17
         ToolTipText     =   "Do not load/show obsolete parts on Part lookups and drop down boxes"
         Top             =   5520
         Width           =   735
      End
      Begin VB.CheckBox cbHideInactiveParts 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   1680
         TabIndex        =   16
         ToolTipText     =   "Do not load/show inactive parts on Part lookups and drop down boxes"
         Top             =   5160
         Width           =   855
      End
      Begin VB.TextBox txtResaleNo 
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Top             =   2520
         Width           =   2580
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   1680
         TabIndex        =   251
         Top             =   4680
         Width           =   3135
         Begin VB.OptionButton optCosting 
            Caption         =   "Actual"
            Height          =   195
            Index           =   1
            Left            =   1320
            TabIndex        =   15
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton optCosting 
            Caption         =   "Standard"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   14
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   1680
         TabIndex        =   249
         Top             =   4080
         Width           =   3375
         Begin VB.OptionButton OptWend 
            Caption         =   "Sunday"
            Height          =   195
            Index           =   1
            Left            =   1320
            TabIndex        =   13
            ToolTipText     =   "Weekending Days"
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton OptWend 
            Caption         =   "Saturday"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   12
            ToolTipText     =   "Weekending Days"
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame z3 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   0
         Left            =   1680
         TabIndex        =   91
         Top             =   4200
         Width           =   3015
      End
      Begin VB.TextBox txtDivEnd 
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Tag             =   "1"
         ToolTipText     =   "1 - 12 (Saved Only If Division Accounts Is Checked)"
         Top             =   3840
         Width           =   375
      End
      Begin VB.TextBox txtDivStart 
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Tag             =   "1"
         ToolTipText     =   "1 - 12 (Saved Only If Division Accounts Is Checked)"
         Top             =   3480
         Width           =   375
      End
      Begin VB.CheckBox optGlDivisions 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         ToolTipText     =   "Verify Accounts And Journals Through The System"
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox txtIntf 
         Height          =   285
         Left            =   4440
         TabIndex        =   4
         Tag             =   "1"
         ToolTipText     =   "International Prefix (Country Code)"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtIntp 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Tag             =   "1"
         ToolTipText     =   "International Prefix (Country Code)"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtNme 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Tag             =   "2"
         Top             =   360
         Width           =   4470
      End
      Begin VB.TextBox txtAdr 
         Height          =   1005
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   1
         Tag             =   "9"
         Top             =   720
         Width           =   4470
      End
      Begin VB.TextBox txtTid 
         Height          =   315
         Left            =   1680
         TabIndex        =   6
         Tag             =   "3"
         Top             =   2160
         Width           =   2580
      End
      Begin VB.CheckBox optAct 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         ToolTipText     =   "Verify Accounts And Journals Through The System"
         Top             =   3000
         Width           =   735
      End
      Begin MSMask.MaskEdBox txtPhn 
         Height          =   288
         Left            =   2052
         TabIndex        =   3
         Tag             =   "1"
         Top             =   1800
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "###-###-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFax 
         Height          =   288
         Left            =   4812
         TabIndex        =   5
         Tag             =   "1"
         Top             =   1800
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "###-###-####"
         PromptChar      =   "_"
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Part Search Button"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   124
         Left            =   3960
         TabIndex        =   299
         ToolTipText     =   "Document Imaging System Integration"
         Top             =   6600
         Width           =   1380
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Disable AutoScan"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   122
         Left            =   3960
         TabIndex        =   295
         ToolTipText     =   "Document Imaging System Integration"
         Top             =   6240
         Width           =   1380
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Document Imaging System Integration"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   121
         Left            =   120
         TabIndex        =   291
         ToolTipText     =   "Document Imaging System Integration"
         Top             =   6240
         Width           =   2820
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Top Summary Account"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   120
         Left            =   120
         TabIndex        =   289
         ToolTipText     =   "Include Data collection module"
         Top             =   6600
         Width           =   2340
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Include Data Collection Module"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   115
         Left            =   120
         TabIndex        =   283
         ToolTipText     =   "Include Data collection module"
         Top             =   5880
         Width           =   2340
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hide Obsolete Parts"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   112
         Left            =   120
         TabIndex        =   275
         Top             =   5520
         Width           =   1500
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hide Inactive Parts"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   111
         Left            =   120
         TabIndex        =   274
         Top             =   5160
         Width           =   1500
      End
      Begin VB.Label Label3 
         Caption         =   "Resale Number"
         Height          =   255
         Left            =   120
         TabIndex        =   253
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Part Costing Default"
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   100
         Left            =   120
         TabIndex        =   250
         Top             =   4800
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "More >>>"
         Height          =   252
         Left            =   6600
         TabIndex        =   239
         Top             =   480
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(Ending Position Of Divison (etc) In the Account Number)"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   68
         Left            =   2484
         TabIndex        =   105
         ToolTipText     =   "Verify Accounts And Journals Through The System"
         Top             =   3876
         Width           =   4488
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(Starting Position Of Divison (etc) In the Account Number)"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   67
         Left            =   2484
         TabIndex        =   104
         ToolTipText     =   "Verify Accounts And Journals Through The System"
         Top             =   3516
         Width           =   4488
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ending At Char"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   66
         Left            =   120
         TabIndex        =   103
         ToolTipText     =   "Verify Accounts And Journals Through The System"
         Top             =   3840
         Width           =   1848
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Starting At Char"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   65
         Left            =   120
         TabIndex        =   102
         ToolTipText     =   "Verify Accounts And Journals Through The System"
         Top             =   3480
         Width           =   1848
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Division Accounts"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   64
         Left            =   120
         TabIndex        =   101
         ToolTipText     =   "Verify Accounts And Journals Through The System"
         Top             =   3240
         Width           =   1848
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(Use Department/Branch/Divisions In Account Structure)"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   63
         Left            =   2484
         TabIndex        =   100
         ToolTipText     =   "Verify Accounts And Journals Through The System"
         Top             =   3240
         Width           =   4488
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(Requires Accounts And Journals)"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   10
         Left            =   2484
         TabIndex        =   99
         Top             =   3000
         Width           =   2928
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Week Ends"
         ForeColor       =   &H00000000&
         Height          =   336
         Index           =   5
         Left            =   120
         TabIndex        =   98
         Top             =   4320
         Width           =   1380
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   0
         Left            =   120
         TabIndex        =   97
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   1
         Left            =   120
         TabIndex        =   96
         Top             =   720
         Width           =   1260
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Number"
         ForeColor       =   &H00000000&
         Height          =   336
         Index           =   2
         Left            =   120
         TabIndex        =   95
         Top             =   1800
         Width           =   1380
      End
      Begin VB.Label z1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
         ForeColor       =   &H00000000&
         Height          =   336
         Index           =   3
         Left            =   3840
         TabIndex        =   94
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Id Number"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   4
         Left            =   120
         TabIndex        =   93
         Top             =   2160
         Width           =   1380
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Verify Journals"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   6
         Left            =   120
         TabIndex        =   92
         ToolTipText     =   "Verify Open Journals Throughout The System"
         Top             =   3000
         Width           =   1848
      End
   End
   Begin VB.Frame tabFrame 
      Caption         =   "Shop"
      Height          =   8715
      Index           =   4
      Left            =   600
      TabIndex        =   184
      Top             =   1200
      Width           =   7572
      Begin VB.CheckBox chkRequireApprovedRoutings 
         Height          =   255
         Left            =   6960
         TabIndex        =   326
         ToolTipText     =   "On Puchout logout."
         Top             =   6960
         Width           =   375
      End
      Begin VB.CheckBox chkShowWcNamesAtPomLogin 
         Height          =   255
         Left            =   6960
         TabIndex        =   324
         ToolTipText     =   "Display Last MO Run Number First"
         Top             =   8040
         Width           =   375
      End
      Begin VB.CheckBox chkDenyLoginIfPriorOpOpen 
         Height          =   255
         Left            =   6960
         TabIndex        =   320
         ToolTipText     =   "Display Last MO Run Number First"
         Top             =   7680
         Width           =   375
      End
      Begin VB.CheckBox ChkLogoutOnPuchout 
         Height          =   255
         Left            =   3240
         TabIndex        =   311
         ToolTipText     =   "On Puchout logout."
         Top             =   8040
         Width           =   375
      End
      Begin VB.CheckBox chkOnlyOpenPO 
         Height          =   255
         Left            =   3240
         TabIndex        =   308
         ToolTipText     =   "Display Last MO Run Number First"
         Top             =   7680
         Width           =   375
      End
      Begin VB.CheckBox chkLotNum 
         Height          =   255
         Left            =   6960
         TabIndex        =   304
         ToolTipText     =   "Display Last MO Run Number First"
         Top             =   7320
         Width           =   375
      End
      Begin VB.CheckBox chkUMOComt 
         Height          =   255
         Left            =   3240
         TabIndex        =   302
         ToolTipText     =   "Display Last MO Run Number First"
         Top             =   7320
         Width           =   375
      End
      Begin VB.CheckBox chkLoginSC 
         Height          =   255
         Left            =   3240
         TabIndex        =   300
         ToolTipText     =   "Display Last MO Run Number First"
         Top             =   6960
         Width           =   375
      End
      Begin VB.CheckBox chkOverMOComp 
         Height          =   255
         Left            =   3240
         TabIndex        =   292
         ToolTipText     =   "Display Last MO Run Number First"
         Top             =   6480
         Width           =   375
      End
      Begin VB.CheckBox chkMelLog 
         Height          =   255
         Left            =   3240
         TabIndex        =   284
         ToolTipText     =   "Display Last MO Run Number First"
         Top             =   6120
         Width           =   375
      End
      Begin VB.CheckBox chkTCSetup 
         Height          =   255
         Left            =   3240
         TabIndex        =   278
         ToolTipText     =   "Display Last MO Run Number First"
         Top             =   5640
         Width           =   375
      End
      Begin VB.CheckBox optShowLastRunFirst 
         Height          =   255
         Left            =   3240
         TabIndex        =   55
         ToolTipText     =   "Display Last MO Run Number First"
         Top             =   5280
         Width           =   375
      End
      Begin VB.CheckBox optDontAllowMOIfNotPC 
         Height          =   255
         Left            =   3240
         TabIndex        =   48
         Top             =   3000
         Width           =   255
      End
      Begin VB.CheckBox optTCSerOp 
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   3240
         TabIndex        =   52
         ToolTipText     =   "Use Pom For Time"
         Top             =   4200
         Width           =   255
      End
      Begin VB.Frame z3 
         BorderStyle     =   0  'None
         Height          =   492
         Index           =   1
         Left            =   3240
         TabIndex        =   185
         Top             =   1356
         Width           =   2652
         Begin VB.OptionButton optTfor 
            Caption         =   "0.0000"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   44
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton optTfor 
            Caption         =   "0.000"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   43
            Top             =   120
            Width           =   1092
         End
      End
      Begin VB.TextBox txtQNM 
         Height          =   285
         Left            =   3240
         TabIndex        =   54
         Tag             =   "1"
         ToolTipText     =   "Scheduling Conversion - 1 Day = n Hours (Valid 8 To 24)"
         Top             =   4920
         Width           =   372
      End
      Begin VB.CheckBox optCreate 
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   3240
         TabIndex        =   53
         ToolTipText     =   "MO's Created With New Sales Order Items"
         Top             =   4560
         Width           =   375
      End
      Begin VB.CheckBox optPom 
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   3240
         TabIndex        =   51
         ToolTipText     =   "Use Pom For Time"
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox txtSplit 
         Height          =   285
         Left            =   5280
         TabIndex        =   50
         Tag             =   "1"
         ToolTipText     =   "Start Splits For Each Manufacturing Order At Run"
         Top             =   3480
         Width           =   735
      End
      Begin VB.CheckBox optSplits 
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   3240
         TabIndex        =   49
         ToolTipText     =   "MO's May Be Split (See Help)"
         Top             =   3480
         Width           =   375
      End
      Begin VB.CheckBox OptVerifyInvoice 
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   3240
         TabIndex        =   47
         ToolTipText     =   "Test Allocated PO Items For Invoices"
         Top             =   2520
         Width           =   735
      End
      Begin VB.CheckBox optQtyChg 
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   3240
         TabIndex        =   46
         ToolTipText     =   "If Checked, Allow MO Quantities To Be Changed After PL (Provides A Warning)"
         Top             =   2160
         Width           =   735
      End
      Begin VB.CheckBox optMo4 
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   3240
         TabIndex        =   45
         ToolTipText     =   "If Checked, Allow Modifications To Type 4 Parts"
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtDefLead 
         Height          =   285
         Left            =   3240
         TabIndex        =   42
         Tag             =   "1"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtDefFlow 
         Height          =   285
         Left            =   3240
         TabIndex        =   41
         Tag             =   "1"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Require Approved Routings"
         Height          =   255
         Left            =   3840
         TabIndex        =   327
         ToolTipText     =   "On Punchout Auto Logout from Current Job"
         Top             =   6960
         Width           =   2775
      End
      Begin VB.Label Label13 
         Caption         =   "Show WC Names at POM Logon"
         Height          =   255
         Left            =   3840
         TabIndex        =   325
         ToolTipText     =   "Display Last MO Run Number First"
         Top             =   8040
         Width           =   2475
      End
      Begin VB.Label Label12 
         Caption         =   "Deny Login if previous op is open"
         Height          =   255
         Left            =   3840
         TabIndex        =   321
         ToolTipText     =   "Display Last MO Run Number First"
         Top             =   7680
         Width           =   2475
      End
      Begin VB.Label Label11 
         Caption         =   "Auto Logout on Punchout"
         Height          =   255
         Left            =   240
         TabIndex        =   310
         ToolTipText     =   "On Punchout Auto Logout from Current Job"
         Top             =   8040
         Width           =   2775
      End
      Begin VB.Label Label10 
         Caption         =   "Show Only Open OP Number"
         Height          =   255
         Left            =   240
         TabIndex        =   309
         ToolTipText     =   "Display Last MO Run Number First"
         Top             =   7680
         Width           =   2775
      End
      Begin VB.Label Label9 
         Caption         =   "Update Lot Number"
         Height          =   255
         Left            =   3840
         TabIndex        =   305
         ToolTipText     =   "Display Last MO Run Number First"
         Top             =   7320
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Update MO Comment"
         Height          =   255
         Left            =   240
         TabIndex        =   303
         ToolTipText     =   "Display Last MO Run Number First"
         Top             =   7320
         Width           =   2775
      End
      Begin VB.Label Label7 
         Caption         =   "Allow Login to SC Status"
         Height          =   255
         Left            =   240
         TabIndex        =   301
         ToolTipText     =   "Display Last MO Run Number First"
         Top             =   6960
         Width           =   2775
      End
      Begin VB.Label Label6 
         Caption         =   "Allow MO Completion Qty greater than Org Qty"
         Height          =   495
         Left            =   240
         TabIndex        =   293
         ToolTipText     =   "Display Last MO Run Number First"
         Top             =   6480
         Width           =   2775
      End
      Begin VB.Label Label5 
         Caption         =   "Enable Melter's Log from POM"
         Height          =   255
         Left            =   240
         TabIndex        =   285
         ToolTipText     =   "Display Last MO Run Number First"
         Top             =   6120
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "Allow option to Time Charge Setup From POM"
         Height          =   495
         Left            =   240
         TabIndex        =   279
         ToolTipText     =   "Display Last MO Run Number First"
         Top             =   5640
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Display Last MO Run Number First"
         Height          =   255
         Left            =   240
         TabIndex        =   252
         ToolTipText     =   "Display Last MO Run Number First"
         Top             =   5280
         Width           =   2895
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Don't Allow MO Completion if not at Pick Complete Status"
         ForeColor       =   &H00000000&
         Height          =   465
         Index           =   99
         Left            =   240
         TabIndex        =   248
         ToolTipText     =   "Test Allocated PO Items For Invoices"
         Top             =   2880
         Width           =   3495
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Allow Time Charges for Service Ops"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   97
         Left            =   240
         TabIndex        =   245
         ToolTipText     =   "MO's May Be Split (See Help)"
         Top             =   4200
         Width           =   3135
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hours (Scheduling)"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   86
         Left            =   3960
         TabIndex        =   200
         ToolTipText     =   "Scheduling Conversion - 1 Day = n Hours (Valid 8 To 24)"
         Top             =   5040
         Width           =   1935
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Queue And Move Conversion "
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   85
         Left            =   240
         TabIndex        =   199
         ToolTipText     =   "Scheduling Conversion - 1 Day = n Hours (Valid 8 To 24)"
         Top             =   4920
         Width           =   2415
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Allow MO Creation From Sales Orders"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   57
         Left            =   240
         TabIndex        =   198
         ToolTipText     =   "MO's May Be Split (See Help)"
         Top             =   4560
         Width           =   3495
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Allow Creation of POM Time Charges"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   84
         Left            =   240
         TabIndex        =   197
         ToolTipText     =   "MO's May Be Split (See Help)"
         Top             =   3840
         Width           =   2775
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Beginning Run"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   77
         Left            =   4080
         TabIndex        =   196
         ToolTipText     =   "Start Splits For Each Manufacturing Order At Run"
         Top             =   3540
         Width           =   1215
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Allow Split Manufacturing Orders"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   76
         Left            =   240
         TabIndex        =   195
         ToolTipText     =   "MO's May Be Split (See Help)"
         Top             =   3480
         Width           =   3495
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Verify Invoicing (PO) Before Closing MO"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   74
         Left            =   240
         TabIndex        =   194
         ToolTipText     =   "Test Allocated PO Items For Invoices"
         Top             =   2520
         Width           =   3495
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Allow Quantity Changes For PL ,PP, PC"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   72
         Left            =   240
         TabIndex        =   193
         ToolTipText     =   "If Checked, Allow MO Quantities To Be Changed After PL (Provides A Warning)"
         Top             =   2160
         Width           =   3495
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Allow Manufacturing Orders For Type 4"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   54
         Left            =   240
         TabIndex        =   192
         ToolTipText     =   "If Checked, Allow Modifications To Type 4 Parts"
         Top             =   1800
         Width           =   3495
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Time Format For Routing/MO Units"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   50
         Left            =   240
         TabIndex        =   191
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(Days)"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   46
         Left            =   4200
         TabIndex        =   190
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(Days)"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   45
         Left            =   4200
         TabIndex        =   189
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Default Manufacturing Flow Time"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   44
         Left            =   240
         TabIndex        =   188
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Default Purchasing Lead Time"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   43
         Left            =   240
         TabIndex        =   187
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Default Shop Floor Settings:"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   42
         Left            =   240
         TabIndex        =   186
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame tabFrame 
      Caption         =   "Purchasing"
      Height          =   6015
      Index           =   5
      Left            =   8580
      TabIndex        =   201
      Top             =   300
      Width           =   7572
      Begin VB.CheckBox chkNextAvailablePO 
         Caption         =   "____"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   3720
         TabIndex        =   322
         Top             =   5400
         Width           =   735
      End
      Begin VB.CheckBox chkWarnServOp 
         Caption         =   "____"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   3720
         TabIndex        =   314
         Top             =   5040
         Width           =   735
      End
      Begin VB.CheckBox chkMoreQty 
         Caption         =   "____"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   3720
         TabIndex        =   273
         ToolTipText     =   "Allows Changes To The PO Line Item Price After Invoice (Does Not Change Actual)"
         Top             =   4440
         Width           =   255
      End
      Begin VB.CheckBox chkShortQty 
         Caption         =   "____"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   3720
         TabIndex        =   260
         ToolTipText     =   "Allows Changes To The PO Line Item Price After Invoice (Does Not Change Actual)"
         Top             =   3840
         Width           =   255
      End
      Begin VB.CheckBox optVendApproval 
         Caption         =   "____"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   3720
         TabIndex        =   247
         Top             =   720
         Width           =   735
      End
      Begin VB.CheckBox chkDefaultPrice 
         Caption         =   "____"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   3720
         TabIndex        =   60
         ToolTipText     =   "Allows Changes To The PO Line Item Price After Invoice (Does Not Change Actual)"
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtLastPon 
         Height          =   285
         Left            =   3720
         TabIndex        =   63
         Tag             =   "1"
         Text            =   "000000"
         Top             =   3360
         Width           =   735
      End
      Begin VB.CheckBox optPoPrice 
         Caption         =   "____"
         ForeColor       =   &H8000000F&
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   59
         ToolTipText     =   "Allows Changes To The PO Line Item Price After Invoice (Does Not Change Actual)"
         Top             =   2160
         Width           =   735
      End
      Begin VB.CheckBox optPoPrice 
         Caption         =   "____"
         ForeColor       =   &H8000000F&
         Height          =   255
         Index           =   0
         Left            =   3720
         TabIndex        =   58
         ToolTipText     =   "Allows Changes To The PO Line Item Price After Receipt (Does Not Change Actual)"
         Top             =   1800
         Width           =   735
      End
      Begin VB.CheckBox optService 
         Caption         =   "____"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   3720
         TabIndex        =   57
         Top             =   1440
         Width           =   735
      End
      Begin VB.CheckBox optPurAcct 
         Caption         =   "____"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   3720
         TabIndex        =   56
         Top             =   1080
         Width           =   735
      End
      Begin VB.Frame z3 
         BorderStyle     =   0  'None
         Height          =   492
         Index           =   2
         Left            =   3600
         TabIndex        =   202
         Top             =   2760
         Width           =   2412
         Begin VB.OptionButton optPfor 
            Caption         =   "0.0000"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   62
            Top             =   140
            Width           =   855
         End
         Begin VB.OptionButton optPfor 
            Caption         =   "0.000"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   61
            Top             =   140
            Width           =   855
         End
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Select next available PO number"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   102
         Left            =   240
         TabIndex        =   323
         Top             =   5400
         Width           =   2355
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Warn Open Service Op"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   127
         Left            =   240
         TabIndex        =   315
         Top             =   5040
         Width           =   3495
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Accept more Qty delivered on a PO Line item. At On Dock Inspection."
         ForeColor       =   &H00000000&
         Height          =   465
         Index           =   109
         Left            =   240
         TabIndex        =   272
         ToolTipText     =   "Allows Changes To The PO Line Item Price After Receipt (Does Not Change Actual)"
         Top             =   4440
         Width           =   3255
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Accept short Qty delivered on a PO Line item. At On Dock Inspection."
         ForeColor       =   &H00000000&
         Height          =   465
         Index           =   108
         Left            =   240
         TabIndex        =   261
         ToolTipText     =   "Allows Changes To The PO Line Item Price After Receipt (Does Not Change Actual)"
         Top             =   3840
         Width           =   3255
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor Approval Required"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   98
         Left            =   240
         TabIndex        =   246
         Top             =   720
         Width           =   3492
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Default PO items to last invoiced price"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   92
         Left            =   240
         TabIndex        =   240
         ToolTipText     =   "Allows Changes To The PO Line Item Price After Invoice (Does Not Change Actual)"
         Top             =   2520
         Width           =   3135
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Purchase Order Number"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   81
         Left            =   240
         TabIndex        =   209
         ToolTipText     =   "Allows Changes To The PO Line Item Price After Invoice (Does Not Change Actual)"
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Order Data Entry Formats"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   73
         Left            =   240
         TabIndex        =   208
         ToolTipText     =   "Allows Changes To The PO Line Item Price After Invoice (Does Not Change Actual)"
         Top             =   2880
         Width           =   3495
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Allow Post Invoice Item Pricing"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   71
         Left            =   240
         TabIndex        =   207
         ToolTipText     =   "Allows Changes To The PO Line Item Price After Invoice (Does Not Change Actual)"
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Allow Post Receipt Item Pricing"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   70
         Left            =   240
         TabIndex        =   206
         ToolTipText     =   "Allows Changes To The PO Line Item Price After Receipt (Does Not Change Actual)"
         Top             =   1800
         Width           =   3492
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Force Routing Use For Part Type 7 (Services)"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   69
         Left            =   240
         TabIndex        =   205
         Top             =   1440
         Width           =   3492
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Accounts In Purchase Order Items"
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   51
         Left            =   240
         TabIndex        =   204
         Top             =   1080
         Width           =   3492
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Purchasing Defaults:"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   48
         Left            =   240
         TabIndex        =   203
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "AdmnADe01a.frx":034E
      Style           =   1  'Graphical
      TabIndex        =   88
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   444
      Index           =   1
      Left            =   6720
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   0
      Width           =   975
   End
   Begin VB.CheckBox optWip 
      Caption         =   "Wip Shows"
      Height          =   255
      Left            =   1680
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   6960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   6840
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   10875
      FormDesignWidth =   8475
   End
   Begin MSComctlLib.TabStrip tab1 
      Height          =   10275
      Left            =   0
      TabIndex        =   89
      Top             =   600
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   18124
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      TabFixedWidth   =   2117
      TabFixedHeight  =   473
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   11
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Company"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Inventory"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Receivables"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Payables"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "S&hop Floor"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "P&urchasing"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "L&ots/Cycle"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "P&acking Slips"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Sales"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Labor"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Engineering"
            Object.ToolTipText     =   "Engineering"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "AdmnADe01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***

Dim bAllowSplits As Byte
Dim bOnLoad As Byte
Dim bGoodJrns As Byte
Dim bLotTracking As Byte
Dim sAcctDesc As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub ClearCombos()
   Dim iControl As Integer
   For iControl = 0 To Controls.count - 1
      If TypeOf Controls(iControl) Is ComboBox Then _
                         Controls(iControl).SelLength = 0
   Next
   
End Sub

Private Sub cbDocumentLok_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET CODOCLOK= " _
          & cbDocumentLok.Value & " WHERE COREF=1"
   clsADOCon.ExecuteSql sSql
End Sub
Private Sub cbDisAutoScan_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET CODISAUTOSCAN= " _
          & cbDisAutoScan.Value & " WHERE COREF=1"
   clsADOCon.ExecuteSql sSql
End Sub






Private Sub chkAbbreviatedLotNumbers_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET CoUseAbbreviatedLotNumbers = " _
          & chkAbbreviatedLotNumbers.Value & " WHERE COREF = 1"
   clsADOCon.ExecuteSql sSql
End Sub

Private Sub chkBOMSetQty_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET COBOMSETQTY = " _
          & chkBOMSetQty.Value & " WHERE COREF=1"
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   
End Sub


Private Sub chkDenyLoginIfPriorOpOpen_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET DenyLoginIfPriorOpOpen = " _
          & chkDenyLoginIfPriorOpOpen.Value & " WHERE COREF = 1"
   clsADOCon.ExecuteSql sSql
End Sub

Private Sub chkEarlyLWarning_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET COPSPOPUPWARNING= " _
          & chkEarlyLWarning.Value & " WHERE COREF=1"
   clsADOCon.ExecuteSql sSql
End Sub

Private Sub chkNextAvailablePO_Click()
   On Error Resume Next
   sSql = "UPDATE Preferences SET GetNextAvailablePoNumber = " _
          & chkNextAvailablePO.Value
   clsADOCon.ExecuteSql sSql ', rdExecDirect
End Sub

Private Sub chkOverShip_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET COPSOVERSHIP= " _
          & chkOverShip.Value & " WHERE COREF=1"
   clsADOCon.ExecuteSql sSql
End Sub

Private Sub cbCustomShipLabel_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET COCUSTOMSHIPLABEL= " _
          & cbCustomShipLabel.Value & " WHERE COREF=1"
   clsADOCon.ExecuteSql sSql
End Sub


Private Sub cbHideInactiveParts_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET COHIDEINACTIVEPART= " _
          & cbHideInactiveParts.Value & " WHERE COREF=1"
   clsADOCon.ExecuteSql sSql ', rdExecDirect
End Sub

Private Sub cbHideObsoleteParts_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET COHIDEOBSOLETEPART= " _
          & cbHideObsoleteParts.Value & " WHERE COREF=1"
   clsADOCon.ExecuteSql sSql ', rdExecDirect
End Sub

Private Sub chkRequireApprovedRoutings_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET CoRequireApprovedRoutings = " _
          & chkRequireApprovedRoutings.Value & " WHERE COREF = 1"
   clsADOCon.ExecuteSql sSql
End Sub

Private Sub chkSheetInventory_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET COUSESHEETINVENTORY= " _
          & Me.chkSheetInventory.Value & " WHERE COREF=1"
   clsADOCon.ExecuteSql sSql 'rdExecDirect
End Sub

Private Sub chkShortQty_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET COALLOWDELSHORTQTY= " _
          & chkShortQty.Value & " WHERE COREF=1"
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   
End Sub

Private Sub chkDataCol_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET CODATACOLMOD= " _
          & chkDataCol.Value & " WHERE COREF=1"
   clsADOCon.ExecuteSql sSql 'rdExecDirect
End Sub

Private Sub chkMoreQty_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET COALLOWDELOVERQTY= " _
          & chkMoreQty.Value & " WHERE COREF=1"
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   
End Sub

Private Sub chkShowWcNamesAtPomLogin_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET CoUseNamesInPOM = " _
          & chkShowWcNamesAtPomLogin.Value & " WHERE COREF = 1"
   clsADOCon.ExecuteSql sSql
End Sub

Private Sub chkWarnServOp_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET COWARNSERVICEOPOPEN= " _
          & chkWarnServOp.Value & " WHERE COREF=1"
   clsADOCon.ExecuteSql sSql 'rdExecDirect

End Sub

Private Sub cmdCan_Click(Index As Integer)
   sFacility = "" & txtNme
   UpdateCompany
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 1101
      MouseCursor False
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdInvoice_Click()
   Dim RdoFind As ADODB.Recordset
   sSql = "SELECT MAX(INVNO) AS LastInvoice FROM CihdTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoFind, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoFind!LastInvoice) Then
         txtTransfer = Format((RdoFind!LastInvoice + 1), "000000")
      Else
         txtTransfer = "000001"
      End If
   End If
   cmdLock.Enabled = True
   Set RdoFind = Nothing
   
End Sub

Private Sub cmdLock_Click()
   Dim bResponse As Byte
   Dim lTransfer As Long
   Dim sMsg As String
   On Error Resume Next
   sMsg = "This Function Will Establish " & txtTransfer & " As The" & vbCr _
          & "Reference Invoice For Inter Company Transfers." & vbCr _
          & "Once Set, It Cannot Be Changed. Continue?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbNo Then
      CancelTrans
   Else
      cmdInvoice.Enabled = False
      lTransfer = Val(txtTransfer)
      sSql = "INSERT INTO CihdTable (INVNO,INVPRE,INVTYPE," _
             & "INVCOMMENTS,INVCANCELED) VALUES(" & lTransfer & ",'T','IT'," _
             & "'Inter Company Transfers',1)"
      clsADOCon.ExecuteSql sSql
      If Err > 0 Then
         cmdInvoice.Enabled = True
         MsgBox "That Invoice Was Used While We Were Away. Try Again.", _
            vbInformation, Caption
      Else
         sSql = "update Preferences SET AllowTransfers=1," _
                & "TransferInvoice = " & lTransfer & " "
         clsADOCon.ExecuteSql sSql
         SysMsg "Transfer Options Set.", True
         cmdLock.Enabled = False
      End If
      cmdLock.Enabled = False
   End If
   
End Sub


Private Sub cmdWip_Click()
   optWip.Value = vbChecked
   AdmnADe01b.Show
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      Dim I As Integer
      For I = 1 To 10
         tabFrame(I).ZOrder (0)     'move to front
      Next
      ES_TimeFormat = GetTimeFormat()
      optTfor(1).Value = True
      If ES_TimeFormat = "##0.000" Then _
                         optTfor(0).Value = True
      bGoodJrns = CheckJournals()
      FillCombos
      GetShopDefaults
      
      GetCompany
      
      bOnLoad = 0
   End If
   If optWip.Value = vbChecked Then Unload AdmnADe01b
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   Dim b As Byte
   MouseCursor 13
   FormLoad Me
   
   For b = 0 To 10
      With tabFrame(b)
         .Left = 40
         .Visible = False
         .BorderStyle = 0
         .Top = 1200
         .Caption = ""
      End With
   Next
   tabFrame(0).Visible = True
   cmdCan(1).Top = 40
   FormatControls
   bOnLoad = 1
   
End Sub

Private Sub GetCompany()
   Dim RdoGet As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet, ES_KEYSET)
   If bSqlRows Then
      With RdoGet
         'On Error Resume Next
         txtAdr = "" & Trim(!COADR)
         txtFax = "" & Trim(!COFAX)
         txtPhn = "" & Trim(!COPHONE)
         txtTid = "" & Trim(!COTAXNO)
         txtResaleNo = "" & Trim(!RESALENUMBER)
         txtCogs = "" & GetAccount(Trim(!COCGSMATACCT1))
         lblCogs = sAcctDesc
         txtArev1 = "" & GetAccount(Trim(!COREVACCT1))
         lblArev1 = sAcctDesc
         txtArev2 = "" & GetAccount(Trim(!COREVACCT2))
         lblArev2 = sAcctDesc
         txtArev3 = "" & GetAccount(Trim(!COREVACCT3))
         lblArev3 = sAcctDesc
         txtArev4 = "" & GetAccount(Trim(!COREVACCT4))
         lblArev4 = sAcctDesc
         txtArev5 = "" & GetAccount(Trim(!COREVACCT5))
         lblArev5 = sAcctDesc
         txtArev6 = "" & GetAccount(Trim(!COREVACCT6))
         lblArev6 = sAcctDesc
         txtArev7 = "" & GetAccount(Trim(!COREVACCT7))
         lblArev7 = sAcctDesc
         txtArev8 = "" & GetAccount(Trim(!COREVACCT8))
         lblArev8 = sAcctDesc
         
         cbHideInactiveParts.Value = Val("" & Trim(!COHIDEINACTIVEPART))
         cbHideObsoleteParts.Value = Val("" & Trim(!COHIDEOBSOLETEPART))
         
         txtFgi = "" & GetAccount(Trim(!COSJINVACCT))
         lblFgi = sAcctDesc
         
         'Receivables
         txtCrcs = "" & GetAccount(Trim(!COCRCASHACCT))
         lblCrcs = "" & sAcctDesc
         txtCrds = "" & GetAccount(Trim(!COCRDISCACCT))
         lblCrds = "" & sAcctDesc
         txtCrex = "" & GetAccount(Trim(!COCREXPACCT))
         lblCrex = "" & sAcctDesc
         txtCrcm = "" & GetAccount(Trim(!COCRCOMMACCT))
         lblCrcm = "" & sAcctDesc
         txtCrrv = "" & GetAccount(Trim(!COCRREVACCT))
         lblCrrv = "" & sAcctDesc
         txtSjar = "" & GetAccount(Trim(!COSJARACCT))
         lblSjar = "" & sAcctDesc
         txtSjtp = "" & GetAccount(Trim(!COSJTAXACCT))
         lblSjtp = "" & sAcctDesc
         txtSjtf = "" & GetAccount(Trim(!COSJTFRTACCT))
         lblSjtf = "" & sAcctDesc
         txtSjnt = "" & GetAccount(Trim(!COSJNFRTACCT))
         lblSjnt = "" & sAcctDesc
         txtFtx = "" & GetAccount(Trim(!COFEDTAXACCT))
         lblFtx = "" & sAcctDesc
         txtTrns = "" & GetAccount(Trim(!COTRANSFEEACCT))
         lblTrns = "" & sAcctDesc
         
         'Payables
         txtPaap = "" & GetAccount(Trim(!COAPACCT))
         lblPaap = "" & sAcctDesc
         txtPatx = "" & GetAccount(Trim(!COPJTAXACCT))
         lblPatx = "" & sAcctDesc
         txtPatf = "" & GetAccount(Trim(!COPJTFRTACCT))
         lblPatf = "" & sAcctDesc
         txtPant = "" & GetAccount(Trim(!COPJNFRTACCT))
         lblPant = "" & sAcctDesc
         txtCdxc = "" & GetAccount(Trim(!COXCCASHACCT))
         lblCdxc = "" & sAcctDesc
         txtCdcc = "" & GetAccount(Trim(!COCCCASHACCT))
         lblCdcc = "" & sAcctDesc
         txtCdds = "" & GetAccount(Trim(!COAPDISCACCT))
         lblCdds = "" & sAcctDesc
         If bGoodJrns Then
            optAct.Value = 0 + !COGLVERIFY
         End If
         If !WEEKENDS = "Sat" Then
            OptWend(0).Value = True
         Else
            OptWend(1).Value = True
         End If
         
         optCosting(Val("" & !DEFCOSTINGMETHOD)).Value = True
        
         If Not IsNull(!COAPDISC) Then optApFrDs.Value = !COAPDISC
         optLots.Value = !COLOTSACTIVE
         bLotTracking = !COLOTSACTIVE
         If !COLOTSFIFO = 1 Then optFifo.Value = True Else optLifo.Value = True
         optLots.Value = !COLOTSACTIVE
         
         'Commissions
         If Not IsNull(!COCOMMISSION) Then
            If !COCOMMISSION = 0 Then
               optCom(0).Value = True
            End If
         End If
         
         optPre.Value = !COALLOWPSPREPICKS
         optSoit.Value = IIf(IsNull(!COALLOWPSNEWSOITEM), 0, !COALLOWPSNEWSOITEM)
         chkShortQty.Value = IIf(IsNull(!COALLOWDELSHORTQTY), 0, !COALLOWDELSHORTQTY)
         chkMoreQty.Value = IIf(IsNull(!COALLOWDELOVERQTY), 0, !COALLOWDELOVERQTY)
         optInvPS.Value = IIf(IsNull(!COALLOWINVNUMPS), 0, !COALLOWINVNUMPS)
         cbCustomShipLabel.Value = IIf(IsNull(!COCUSTOMSHIPLABEL), 0, !COCUSTOMSHIPLABEL)
            
         chkOverShip.Value = IIf(IsNull(!COPSOVERSHIP), 0, !COPSOVERSHIP)
         chkBOMSetQty.Value = IIf(IsNull(!COBOMSETQTY), 0, !COBOMSETQTY)
        
         '10/8/03 International phones
         txtIntp = "" & Trim(!COINTPHONE)
         txtIntf = "" & Trim(!COINTFAX)
         optMo4 = !COALLOWTYPEFOURMO
         '11/19/03
         txtDefTime = "" & Trim(!CODEFTIMEACCT)
         txtDefLabor = "" & Trim(!CODEFLABORACCT)
         
         '12/9/03 ABC Codes
         TxtAbcCount = Format(!COABCCOUNTERS, "##0")
         txtAbcLow = Format(!COABCLOWLIMITCOST, "##0.00")
         txtAbcHigh = Format(!COABCHIGHLIMITCOST, "#####0.00")
         txtNme = "" & Trim(!CONAME)
         '2/16/04 Division characters for some uncomprehensible reason
         optGlDivisions.Value = !COGLDIVISIONS
         txtDivStart = Format(!COGLDIVSTARTPOS, "0")
         txtDivEnd = Format(!COGLDIVENDPOS, "0")
         '11/11/04
         OptVerifyInvoice = Format(!COVERIFYINVOICES, "0")
         '07/07/2010
         optDontAllowMOIfNotPC.Value = Val("" & !CODONTALLOWMONOTPC)
         '1/22/05
         txtMos = "" & GetAccount(Trim(!COADJACCT))
         lblMos = "" & sAcctDesc
         '2/8/05
         optSplits.Value = !COALLOWSPLITS
         txtSplit = !COSPLITSTARTRUN
         bAllowSplits = !COALLOWSPLITS
         '7/15/05
         txtLastPon = Format(!COLASTPURCHASEORDER, "000000")
         txtLastSon = "" & Trim(!COLASTSALESORDER)
         '8/17/05
         optPom.Value = !COPOMTIME
         optTCSerOp.Value = IIf(IsNull(!COPOTIMESERVOP), 0, !COPOTIMESERVOP)
  
         If (optPom.Value = 0) Then
            optTCSerOp.Enabled = False
         Else
            optTCSerOp.Enabled = True
         End If
         
         ' include OH rate post to GL
         optOHcost.Value = IIf(IsNull(!COLABOROHTOGL), 0, !COLABOROHTOGL)

         
'         '7/6/06
'         If Not IsNull(!CURPSNUMBER) Or Trim(!CURPSNUMBER) <> "" Then
'            txtPsl = "" & Trim(!CURPSNUMBER)
'         Else
'            txtPsl = "000000"
'         End If

         'revised 5/16 to handle new optional prefix logic
         'done to accomodate 8-digit ps numbers for Pegasus
         txtPackPrefix = Trim(!COPSPREFIX)
         FormatPSNumber !COLASTPSNUMBER
         
         Me.chkDefaultPrice.Value = IIf(!COPODEFAULTTOLASTPRICE, 1, 0)
         Me.txtLastInvoiceNumber = !COLASTINVOICENUMBER
         Me.txtLinesPerStub = !COLinesPerCheckStub
         Me.txtWrnCusCL = !CUWARNCREDITLMT
         
         chkRout = IIf(IsNull(!COROUTSEC), 0, !COROUTSEC)
         chkBOM = IIf(IsNull(!COBOMSEC), 0, !COBOMSEC)
         chkDocLst = IIf(IsNull(!CODOCLSTSEC), 0, !CODOCLSTSEC)
         chkDataCol.Value = IIf(IsNull(!CODATACOLMOD), 0, !CODATACOLMOD)
         
         cbDocumentLok.Value = IIf(IsNull(!CODOCLOK), 0, !CODOCLOK)
         cbDisAutoScan.Value = IIf(IsNull(!CODISAUTOSCAN), 0, !CODISAUTOSCAN)
         
         
         txtBIVSOffSetAcc = "" & IIf(IsNull(!COBIVSOFFSETACCT), 0, Trim(!COBIVSOFFSETACCT))
         chkTCSetup = IIf(IsNull(!COSETUPTCPOM), 0, !COSETUPTCPOM)
         chkMelLog = IIf(IsNull(!ENABLEMELTERSLOG), 0, !ENABLEMELTERSLOG)
         
         chkIgnQty = IIf(IsNull(!COIGNQYTRPTSO), 0, !COIGNQYTRPTSO)
         
         chkNoTCOvh = IIf(IsNull(!CONOTCSHAREOVH), 0, !CONOTCSHAREOVH)
         
         chkSumAcct = IIf(IsNull(!COTOPSUMACCT), 0, !COTOPSUMACCT)
         
         chkOverMOComp = IIf(IsNull(!ALLOWOVERQTYCOMP), 0, !ALLOWOVERQTYCOMP)
         
         chkPartSrch = IIf(IsNull(!PARTSEARCHOP), 0, !PARTSEARCHOP)
         chkLoginSC = IIf(IsNull(!COALLOWSCPOM), 0, !COALLOWSCPOM)
         chkUMOComt = IIf(IsNull(!COALLOWMOCOMT), 0, !COALLOWMOCOMT)
         chkLotNum = IIf(IsNull(!COLOTATPOM), 0, !COLOTATPOM)
         chkOnlyOpenPO = IIf(IsNull(!COONLYOPENPO), 0, !COONLYOPENPO)
         ChkLogoutOnPuchout = IIf(IsNull(!COLOGOUTONPOUT), 0, !COLOGOUTONPOUT)
         chkEarlyLWarning = IIf(IsNull(!COPSPOPUPWARNING), 0, !COPSPOPUPWARNING)
         chkDenyLoginIfPriorOpOpen = IIf(IsNull(!DenyLoginIfPriorOpOpen), 0, !DenyLoginIfPriorOpOpen)
         
         chkWarnServOp = IIf(IsNull(!COWARNSERVICEOPOPEN), 0, !COWARNSERVICEOPOPEN)
         chkSheetInventory = IIf(IsNull(!COUSESHEETINVENTORY), 0, !COUSESHEETINVENTORY)  'AWJ sheet inv 10/25/2016 TEL
         chkShowWcNamesAtPomLogin = IIf(IsNull(!CoUseNamesInPOM), 0, !CoUseNamesInPOM)
         chkRequireApprovedRoutings = IIf(IsNull(!CoRequireApprovedRoutings), 0, !CoRequireApprovedRoutings)
         chkAbbreviatedLotNumbers = IIf(IsNull(!CoUseAbbreviatedLotNumbers), 0, !CoUseAbbreviatedLotNumbers)
         
         ClearResultSet RdoGet
      End With
   Else
      sSql = "INSERT INTO ComnTable (CONAME) VALUES('')"
      clsADOCon.ExecuteSql sSql
   End If
   sSql = "SELECT REQVENDORAPPROVAL, PurchaseAccount,ForceServices,AllowPostReceiptPricing,AllowPostInvoicePricing," & vbCrLf _
      & "PurchasedDataFormat,SellingPriceFormat, AutoSelectLastRun, EngineeringLaborRate,GetNextAvailablePoNumber" & vbCrLf _
      & "FROM Preferences WHERE PreRecord=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet, ES_FORWARD)
   If bSqlRows Then
      With RdoGet
         If Not IsNull(!REQVENDORAPPROVAL) Then optVendApproval.Value = !REQVENDORAPPROVAL    'BBS Added for New Vendor Approval Option
         If Not IsNull(!PurchaseAccount) Then optPurAcct.Value = !PurchaseAccount
         If Not IsNull(!ForceServices) Then optService.Value = !ForceServices
         If Not IsNull(!AllowPostReceiptPricing) Then optPoPrice(0).Value = !AllowPostReceiptPricing
         If Not IsNull(!AllowPostinvoicePricing) Then optPoPrice(1).Value = !AllowPostinvoicePricing
         If Not IsNull(!AutoSelectLastRun) Then optShowLastRunFirst.Value = !AutoSelectLastRun
         If Not IsNull(!GetNextAvailablePoNumber) Then chkNextAvailablePO.Value = !GetNextAvailablePoNumber
        
         'If Trim(Right(!PurchasedDataFormat, 4)) = ".000" Then
         
         ' 3 and 4 decimal error. 4/19/2009
         If Right(Trim(!PurchasedDataFormat), 4) = ".000" Then
            optPfor(0).Value = True
            optPfor(1).Value = False
         Else
            optPfor(0).Value = False
            optPfor(1).Value = True
         End If
         
         If Right(Trim(!SellingPriceFormat), 4) = ".000" Then
            optSaleFor(0).Value = True
            optSaleFor(1).Value = False
         Else
            optSaleFor(0).Value = False
            optSaleFor(1).Value = True
         End If
         
         txtEngineeringRate = Format(!EngineeringLaborRate, "##0.00")
         
         ClearResultSet RdoGet
      End With
   End If
   '8/1/05
   sSql = "SELECT TransferInvoice,AllowTransfers,AllowMOCreationFromSO FROM Preferences WHERE PreRecord=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet, ES_FORWARD)
   If bSqlRows Then
      With RdoGet
         txtTransfer = Format(!TransferInvoice, "000000")
         optTransfer.Value = !AllowTransfers
         optCreate.Value = !AllowMOCreationFromSO
         ClearResultSet RdoGet
      End With
   End If
   sSql = "SELECT QueueMoveConversion From Preferences WHERE PreRecord=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet, ES_FORWARD)
   If bSqlRows Then
      With RdoGet
         txtQNM = Format(!QueueMoveConversion, "###0")
         ClearResultSet RdoGet
      End With
   End If
   
   If Val(txtTransfer) > 0 Then cmdInvoice.Enabled = False
   On Error Resume Next
   txtNme.SetFocus
   Set RdoGet = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getcompan"

   MsgBox ("Catch: Get COmpany:" & CStr(Err.Number))
   MsgBox ("Catch: Get COmpany:" & CStr(Err.Description))
   
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'9/1/04 to ease shipping

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   '
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   If Len(Trim(txtNme)) Then SaveSetting "Esi2000", "System", "CompanyName", txtNme
   MouseCursor 0
   Set AdmnADe01a = Nothing
   
End Sub



Private Sub UpdateCompany()
   Dim bCom As Byte
   Dim bFifo As Byte
   Dim bGLStart As Byte
   Dim bGLEnd As Byte
   MouseCursor 13
   If bAllowSplits = 0 Then
      If optSplits.Value = vbChecked Then
         MsgBox "Warning:" & vbCr _
            & "You Have Chosen To Allow Split Manufacturing " & vbCr _
            & "Orders. Be Certain That You Have Read Help  " & vbCr _
            & "And Understand The Process Before Splits " & vbCr _
            & "Are Made.", vbExclamation, Caption
      End If
   End If
   If optFifo.Value = True Then bFifo = 1 Else bFifo = 0
   If optCom(0).Value = True Then bCom = 0 Else bCom = 1
   If optGlDivisions.Value = vbChecked Then
      bGLStart = Val(txtDivStart)
      bGLEnd = Val(txtDivEnd)
   Else
      bGLStart = 0
      bGLEnd = 0
   End If
   On Error GoTo DiaErr1
   txtNme = "" & RTrim(txtNme)
   txtAdr = "" & RTrim(txtAdr)
   txtFax = "" & RTrim(txtFax)
   txtPhn = "" & RTrim(txtPhn)
   txtTid = "" & RTrim(txtTid)
   txtResaleNo = "" & RTrim(txtResaleNo)
   sSql = "UPDATE ComnTable SET CONAME='" & txtNme & "'," & vbCrLf _
          & "COFAX='" & txtFax & "'," & vbCrLf _
          & "COPHONE='" & txtPhn & "'," & vbCrLf _
          & "COTAXNO='" & txtTid & "'," & vbCrLf _
          & "RESALENUMBER='" & txtResaleNo & "'," & vbCrLf _
          & "COGLVERIFY=" & optAct.Value & "," & vbCrLf _
          & "COREVACCT1='" & Compress(txtArev1) & "'," & vbCrLf _
          & "COREVACCT2='" & Compress(txtArev2) & "'," & vbCrLf _
          & "COREVACCT3='" & Compress(txtArev3) & "'," & vbCrLf _
          & "COREVACCT4='" & Compress(txtArev4) & "'," & vbCrLf _
          & "COREVACCT5='" & Compress(txtArev5) & "'," & vbCrLf _
          & "COREVACCT6='" & Compress(txtArev6) & "'," & vbCrLf _
          & "COREVACCT7='" & Compress(txtArev7) & "'," & vbCrLf _
          & "COREVACCT8='" & Compress(txtArev8) & "' "
   clsADOCon.ExecuteSql sSql
   
   sSql = "UPDATE ComnTable SET " & vbCrLf _
          & "COCGSMATACCT1='" & Compress(txtCogs) & "'," & vbCrLf _
          & "COSJARACCT='" & Compress(txtSjar) & "'," & vbCrLf _
          & "COSJTAXACCT='" & Compress(txtSjtp) & "'," & vbCrLf _
          & "COSJTFRTACCT='" & Compress(txtSjtf) & "'," & vbCrLf _
          & "COSJNFRTACCT='" & Compress(txtSjnt) & "'," & vbCrLf _
          & "COCRCASHACCT='" & Compress(txtCrcs) & "'," & vbCrLf _
          & "COCRDISCACCT='" & Compress(txtCrds) & "'," & vbCrLf _
          & "COCREXPACCT='" & Compress(txtCrex) & "'," & vbCrLf _
          & "COCRCOMMACCT='" & Compress(txtCrcm) & "'," & vbCrLf _
          & "COCRREVACCT='" & Compress(txtCrrv) & "'," & vbCrLf _
          & "COAPACCT='" & Compress(txtPaap) & "'," & vbCrLf _
          & "COPJTAXACCT='" & Compress(txtPatx) & "'," & vbCrLf _
          & "COPJTFRTACCT='" & Compress(txtPatf) & "'," & vbCrLf _
          & "COPJNFRTACCT='" & Compress(txtPant) & "'," & vbCrLf _
          & "COXCCASHACCT='" & Compress(txtCdxc) & "'," & vbCrLf _
          & "COCCCASHACCT='" & Compress(txtCdcc) & "'," & vbCrLf _
          & "COAPDISCACCT='" & Compress(txtCdds) & "'," & vbCrLf _
          & "COSJINVACCT='" & Compress(txtFgi) & "'," & vbCrLf _
          & "COFEDTAXACCT='" & Compress(txtFtx) & "'," & vbCrLf _
          & "COTRANSFEEACCT='" & Compress(txtTrns) & "'," & vbCrLf _
          & "COAPDISC =" & optApFrDs.Value & "," & vbCrLf _
          & "COLOTSACTIVE=" & optLots.Value & "," & vbCrLf _
          & "COLOTSFIFO=" & bFifo & ","
   sSql = sSql & vbCrLf _
          & "COCOMMISSION=" & bCom & "," & vbCrLf _
          & "COALLOWPSPREPICKS=" & optPre.Value & "," & vbCrLf _
          & "COINTPHONE='" & txtIntp & "'," & vbCrLf _
          & "COINTFAX='" & txtIntf & "'," & vbCrLf _
          & "COALLOWTYPEFOURMO=" & optMo4.Value & "," & vbCrLf _
          & "CODEFTIMEACCT='" & Compress(txtDefTime) & "'," & vbCrLf _
          & "CODEFLABORACCT='" & Compress(txtDefLabor) & "'," & vbCrLf _
          & "COABCCOUNTERS=" & Val(TxtAbcCount) & "," & vbCrLf _
          & "COABCLOWLIMITCOST=" & Val(txtAbcLow) & "," & vbCrLf _
          & "COABCHIGHLIMITCOST=" & Val(txtAbcHigh) & "," & vbCrLf _
          & "COGLDIVISIONS=" & optGlDivisions.Value & "," & vbCrLf _
          & "COGLDIVSTARTPOS=" & bGLStart & "," & vbCrLf _
          & "COGLDIVENDPOS=" & bGLEnd & "," & vbCrLf _
          & "COADJACCT='" & Compress(txtMos) & "'," & vbCrLf _
          & "COVERIFYINVOICES=" & OptVerifyInvoice.Value & "," & vbCrLf _
          & "COALLOWSPLITS=" & optSplits.Value & "," & vbCrLf _
          & "COSPLITSTARTRUN=" & Val(txtSplit) & "," & vbCrLf _
          & "COLASTPURCHASEORDER=" & Val(txtLastPon) & "," & vbCrLf _
          & "COLASTSALESORDER='" & txtLastSon & "'," & vbCrLf _
          & "COPOMTIME=" & optPom.Value & "," & vbCrLf _
          & "COPOTIMESERVOP=" & optTCSerOp.Value & "," & vbCrLf
          
      'note: CURPSNUMBER is no longer used 5/16/08
      ' "CURPSNUMBER='" & txtPsl & "'," & vbCrLf
      sSql = sSql _
         & "COPSPREFIX='" & Trim(txtPackPrefix) & "'," & vbCrLf _
         & "COLASTPSNUMBER=" & txtPsl & "," & vbCrLf _
         & "COPODEFAULTTOLASTPRICE=" & chkDefaultPrice & "," & vbCrLf _
         & "COLASTINVOICENUMBER=" & Me.txtLastInvoiceNumber & "," & vbCrLf _
         & "COLinesPerCheckStub=" & Me.txtLinesPerStub & "," & vbCrLf _
         & "CODONTALLOWMONOTPC=" & optDontAllowMOIfNotPC.Value & "," & vbCrLf _
         & "CUWARNCREDITLMT=" & Val(txtWrnCusCL) & "," & vbCrLf _
         & "COBIVSOFFSETACCT='" & Val(txtBIVSOffSetAcc) & "'" & vbCrLf _
         & "WHERE COREF=1"
         
'         Clipboard.Clear
'         Clipboard.SetText sSql
         
   clsADOCon.ExecuteSql sSql
   Sleep 500
   Unload Me
   Exit Sub
   
DiaErr1:
   sProcName = "updatecom"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub




Private Sub optAct_Click()
   Dim b As Byte
   If Not bOnLoad Then
      If bGoodJrns Then
         b = optAct.Value
      Else
         b = 0
         If optAct.Value = vbChecked Then
            MsgBox "Requires Accounts And GL Journals.", _
               vbExclamation, Caption
            optAct.Value = vbUnchecked
         End If
      End If
      On Error Resume Next
      sSql = "UPDATE ComnTable SET COGLVERIFY=" & b & " "
      clsADOCon.ExecuteSql sSql
   End If
   
End Sub


Private Sub optOHcost_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET COLABOROHTOGL= " _
          & optOHcost.Value & " WHERE COREF=1"
   clsADOCon.ExecuteSql sSql ', rdExecDirect

End Sub

Private Sub optAct_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optApFrDs_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


'Commissions By Invoice (Index 0) Or CR (Index 1)

Private Sub optCom_Click(Index As Integer)
   Dim b As Byte
   On Error Resume Next
   If optCom(0).Value = True Then b = 0 Else b = 1
   sSql = "UPDATE ComnTable SET COCOMMISSION=" & b & " " _
          & "WHERE COREF=1"
   clsADOCon.ExecuteSql sSql
   
End Sub

Private Sub optCosting_Click(Index As Integer)
   sSql = "UPDATE ComnTable SET DEFCOSTINGMETHOD = " & LTrim(str(Index)) & " WHERE COREF=1"
   clsADOCon.ExecuteSql sSql
End Sub

Private Sub optCreate_Click()
   On Error Resume Next
   sSql = "UPDATE Preferences SET AllowMOCreationFromSO=" _
          & optCreate.Value & " WHERE PreRecord=1"
   clsADOCon.ExecuteSql sSql
End Sub



Private Sub optGlDivisions_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub



Private Sub optLots_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   If bLotTracking = 1 And optLots.Value = vbUnchecked Then
      sMsg = "Lot Tracking Was Turned On (Enabled). " & vbCr _
             & "Do You Really Want To Disable It?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbNo Then
         optLots.Value = vbChecked
         Exit Sub
      End If
   End If
   On Error Resume Next
   sSql = "UPDATE ComnTable SET COLOTSACTIVE=" & optLots.Value & " "
   clsADOCon.ExecuteSql sSql
   bLotTracking = optLots.Value
   
End Sub

Private Sub optLots_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optMo4_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPfor_Click(Index As Integer)
   If bOnLoad = 0 Then
      If optPfor(0).Value = True Then
         ES_PurchasedDataFormat = "#####0.000"  '4/19/2009 'ES_QuantityDataFormat
      Else
         ES_PurchasedDataFormat = "######0.0000"
      End If
      sSql = "UPDATE Preferences SET PurchasedDataFormat='" & ES_PurchasedDataFormat _
             & "' WHERE PreRecord=1"
      clsADOCon.ExecuteSql sSql
   End If
   
End Sub

Private Sub optPom_Click()
    If (optPom.Value = 0) Then
       optTCSerOp.Enabled = False
       optTCSerOp.Value = 0
    Else
       optTCSerOp.Enabled = True
    End If
End Sub

Private Sub optPoPrice_Click(Index As Integer)
   On Error Resume Next
   If Index = 0 Then
      sSql = "UPDATE Preferences SET AllowPostReceiptPricing=" & optPoPrice(0).Value _
             & " WHERE PreRecord=1"
      clsADOCon.ExecuteSql sSql
   Else
      sSql = "UPDATE Preferences SET AllowPostInvoicePricing=" & optPoPrice(1).Value _
             & " WHERE PreRecord=1"
      clsADOCon.ExecuteSql sSql
   End If
End Sub

Private Sub optPoPrice_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub chkRout_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET COROUTSEC = " _
          & chkRout.Value & " WHERE COREF=1"
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   
End Sub

Private Sub chkBOM_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET COBOMSEC = " _
          & chkBOM.Value & " WHERE COREF=1"
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   
End Sub

Private Sub chkDocLst_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET CODOCLSTSEC = " _
          & chkDocLst.Value & " WHERE COREF=1"
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   
End Sub



Private Sub optPre_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET COALLOWPSPREPICKS= " _
          & optPre.Value & " WHERE COREF=1"
   clsADOCon.ExecuteSql sSql
   
End Sub

Private Sub chkTCSetup_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET COSETUPTCPOM= " _
          & chkTCSetup.Value & " WHERE COREF=1"
   clsADOCon.ExecuteSql sSql
   
End Sub

Private Sub chkMelLog_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET ENABLEMELTERSLOG = " _
          & chkMelLog.Value & " WHERE COREF = 1"
   clsADOCon.ExecuteSql sSql ', rdExecDirect
   
End Sub


Private Sub chkOverMOComp_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET ALLOWOVERQTYCOMP = " _
          & chkOverMOComp.Value & " WHERE COREF = 1"
   clsADOCon.ExecuteSql sSql ', rdExecDirect
   
End Sub

Private Sub chkPartSrch_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET PARTSEARCHOP = " _
          & chkPartSrch.Value & " WHERE COREF = 1"
   clsADOCon.ExecuteSql sSql ', rdExecDirect
   
End Sub


Private Sub chkLoginSC_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET COALLOWSCPOM = " _
          & chkLoginSC.Value & " WHERE COREF = 1"
   clsADOCon.ExecuteSql sSql ', rdExecDirect
   
End Sub
Private Sub chkUMOComt_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET COALLOWMOCOMT = " _
          & chkUMOComt.Value & " WHERE COREF = 1"
   clsADOCon.ExecuteSql sSql ', rdExecDirect
   
End Sub

Private Sub chkLotNum_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET COLOTATPOM = " _
          & chkLotNum.Value & " WHERE COREF = 1"
   clsADOCon.ExecuteSql sSql
   
End Sub

Private Sub ChkLogoutOnPuchout_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET COLOGOUTONPOUT = " _
          & ChkLogoutOnPuchout.Value & " WHERE COREF = 1"
   clsADOCon.ExecuteSql sSql
   
End Sub

Private Sub chkOnlyOpenPO_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET COONLYOPENPO = " _
          & chkOnlyOpenPO.Value & " WHERE COREF = 1"
   clsADOCon.ExecuteSql sSql
   
End Sub

Private Sub chkIgnQty_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET COIGNQYTRPTSO = " _
          & chkIgnQty.Value & " WHERE COREF = 1"
   clsADOCon.ExecuteSql sSql ', rdExecDirect
   
End Sub

Private Sub chkNoTCOvh_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET CONOTCSHAREOVH = " _
          & chkNoTCOvh.Value & " WHERE COREF=1"
   clsADOCon.ExecuteSql sSql
   
End Sub

Private Sub chkSumAcct_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET COTOPSUMACCT = " _
          & chkSumAcct.Value & " WHERE COREF=1"
   clsADOCon.ExecuteSql sSql
   
End Sub

Private Sub optSoit_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET COALLOWPSNEWSOITEM= " _
          & optSoit.Value & " WHERE COREF=1"
   clsADOCon.ExecuteSql sSql
   
End Sub

Private Sub optInvPS_Click()
   On Error Resume Next
   sSql = "UPDATE ComnTable SET COALLOWINVNUMPS= " _
          & optInvPS.Value & " WHERE COREF=1"
   clsADOCon.ExecuteSql sSql
   
End Sub

Private Sub optPre_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPurAcct_Click()
   On Error Resume Next
   sSql = "UPDATE Preferences SET PurchaseAccount=" & optPurAcct.Value _
          & " WHERE PreRecord=1"
   clsADOCon.ExecuteSql sSql
   
End Sub

Private Sub optPurAcct_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optQtyChg_Click()
   sSql = "UPDATE Preferences SET AllowMOQuantityChanges=" & optQtyChg.Value _
          & " WHERE PreRecord=1"
   clsADOCon.ExecuteSql sSql
   
End Sub

Private Sub optQtyChg_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optSaleFor_Click(Index As Integer)
   If bOnLoad = 0 Then
      If optSaleFor(0).Value = True Then
         ES_SellingPriceFormat = "#####0.000" '4/19/2009 'ES_QuantityDataFormat
      Else
         ES_SellingPriceFormat = "######0.0000"
      End If
      sSql = "UPDATE Preferences SET SellingPriceFormat='" & ES_SellingPriceFormat _
             & "' WHERE PreRecord=1"
      clsADOCon.ExecuteSql sSql
   End If
   
End Sub

Private Sub optService_Click()
   On Error Resume Next
   sSql = "UPDATE Preferences SET ForceServices=" & optService.Value _
          & " WHERE PreRecord=1"
   clsADOCon.ExecuteSql sSql
   
End Sub

Private Sub optShowLastRunFirst_Click()
   On Error Resume Next
   sSql = "UPDATE Preferences SET AutoSelectLastRun=" _
          & optShowLastRunFirst.Value & " WHERE PreRecord=1"
   clsADOCon.ExecuteSql sSql

End Sub

Private Sub optSplits_Click()
    If optSplits.Value = vbChecked Then
        txtSplit.Enabled = True
    Else
        txtSplit.Enabled = False
        txtSplit.Text = ""
    End If
End Sub

Private Sub optSplits_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optTfor_Click(Index As Integer)
   If bOnLoad = 0 Then
      If optTfor(0).Value = True Then
         ES_TimeFormat = "##0.000"
      Else
         ES_TimeFormat = "##0.0000"
      End If
      sSql = "UPDATE Preferences SET TimeFormat='" & ES_TimeFormat _
             & "' WHERE PreRecord=1"
      clsADOCon.ExecuteSql sSql
   End If
   
End Sub

Private Sub optTransfer_Click()
   If optTransfer.Value = vbChecked Then
      If Val(txtTransfer) = 0 Then cmdInvoice.Enabled = True _
             Else cmdInvoice.Enabled = False
   Else
      cmdInvoice.Enabled = False
   End If
   sSql = "update Preferences SET AllowTransfers=" & optTransfer.Value & " "
   clsADOCon.ExecuteSql sSql
   
End Sub

Private Sub optVendApproval_Click()
On Error Resume Next
   sSql = "UPDATE Preferences SET REQVENDORAPPROVAL=" & optVendApproval.Value _
          & " WHERE PreRecord=1"
   clsADOCon.ExecuteSql sSql
End Sub

Private Sub optVendApproval_KeyPress(KeyAscii As Integer)
    KeyLock KeyAscii
End Sub

Private Sub OptVerifyInvoice_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub OptWend_Click(Index As Integer)
   Dim sWeekDay As String
   If Index = 0 Then sWeekDay = "Sat" Else sWeekDay = "Sun"
   sSql = "UPDATE ComnTable SET WEEKENDS='" & sWeekDay & "' "
   clsADOCon.ExecuteSql sSql
   
End Sub

Private Sub OptWend_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub tab1_Click()
   Dim b As Byte
   On Error Resume Next
   ClearCombos
   For b = 0 To 10
      tabFrame(b).Visible = False
   Next
   tabFrame(tab1.SelectedItem.Index - 1).Visible = True
   Select Case tab1.SelectedItem.Index
      Case 1
         txtNme.SetFocus
      Case 2
         txtCogs.SetFocus
      Case 3
         txtCrcs.SetFocus
      Case 4
         txtPaap.SetFocus
      Case 5
         txtDefFlow.SetFocus
      Case 6
         optVendApproval.SetFocus   ' BBS Changed for New Vendor Approval option
      Case 7
         optLots.SetFocus
      Case 8
         optPre.SetFocus
      'Case 9
         'txtDummy.SetFocus
      Case 9
         txtDefTime.SetFocus
      Case 10
         chkRout.SetFocus
      Case Else
         txtNme.SetFocus
         
   End Select
   
End Sub



Private Sub TxtAbcCount_LostFocus()
   TxtAbcCount = CheckLen(TxtAbcCount, 3)
   If Val(TxtAbcCount) > 360 Then
      'Beep
      TxtAbcCount = 360
   End If
   TxtAbcCount = Format(Abs(Val(TxtAbcCount)), "##0")
   
End Sub


Private Sub txtAbcHigh_LostFocus()
   txtAbcHigh = CheckLen(txtAbcHigh, 10)
   If Val(txtAbcLow) > Val(txtAbcHigh) Then
      'Beep
      txtAbcLow = 0
   End If
   txtAbcHigh = Format(Abs(Val(txtAbcHigh)), "#####0.00")
   
End Sub


Private Sub txtAbcLow_LostFocus()
   txtAbcLow = CheckLen(txtAbcLow, 9)
   If Val(txtAbcLow) > Val(txtAbcHigh) Then
      'Beep
      txtAbcLow = 0
   End If
   txtAbcLow = Format(Abs(Val(txtAbcLow)), "##0.00")
   
End Sub


Private Sub txtAdr_LostFocus()
   txtAdr = CheckLen(txtAdr, 255)
   sSql = "UPDATE ComnTable SET COADR='" & txtAdr & "' " _
          & "WHERE COREF=1"
   clsADOCon.ExecuteSql sSql
   
End Sub

Private Sub txtArev1_Click()
   txtArev1 = GetAccount(txtArev1)
   lblArev1 = sAcctDesc
   
End Sub

Private Sub txtArev1_LostFocus()
   txtArev1 = CheckLen(txtArev1, 12)
   txtArev1 = GetAccount(txtArev1)
   lblArev1 = sAcctDesc
   
End Sub


Private Sub txtArev2_Click()
   txtArev2 = GetAccount(txtArev2)
   lblArev2 = sAcctDesc
   
End Sub

Private Sub txtArev2_LostFocus()
   txtArev2 = CheckLen(txtArev2, 12)
   txtArev2 = GetAccount(txtArev2)
   lblArev2 = sAcctDesc
   
End Sub


Private Sub txtArev3_Click()
   txtArev3 = GetAccount(txtArev3)
   lblArev3 = sAcctDesc
End Sub

Private Sub txtArev3_LostFocus()
   txtArev3 = CheckLen(txtArev3, 12)
   txtArev3 = GetAccount(txtArev3)
   lblArev3 = sAcctDesc
   
End Sub


Private Sub txtArev4_Click()
   txtArev4 = GetAccount(txtArev4)
   lblArev4 = sAcctDesc
   
End Sub

Private Sub txtArev4_LostFocus()
   txtArev4 = CheckLen(txtArev4, 12)
   txtArev4 = GetAccount(txtArev4)
   lblArev4 = sAcctDesc
   
End Sub


Private Sub txtArev5_Click()
   txtArev5 = GetAccount(txtArev5)
   lblArev5 = sAcctDesc
   
End Sub

Private Sub txtArev5_LostFocus()
   txtArev5 = CheckLen(txtArev5, 12)
   txtArev5 = GetAccount(txtArev5)
   lblArev5 = sAcctDesc
   
End Sub


Private Sub txtArev6_Click()
   txtArev6 = GetAccount(txtArev6)
   lblArev6 = sAcctDesc
   
End Sub

Private Sub txtArev6_LostFocus()
   txtArev6 = CheckLen(txtArev6, 12)
   txtArev6 = GetAccount(txtArev6)
   lblArev6 = sAcctDesc
   
End Sub


Private Sub txtArev7_Click()
   txtArev7 = GetAccount(txtArev7)
   lblArev7 = sAcctDesc
   
End Sub

Private Sub txtArev7_LostFocus()
   txtArev7 = CheckLen(txtArev7, 12)
   txtArev7 = GetAccount(txtArev7)
   lblArev7 = sAcctDesc
   
End Sub


Private Sub txtArev8_Click()
   txtArev8 = GetAccount(txtArev8)
   lblArev8 = sAcctDesc
   
End Sub

Private Sub txtArev8_LostFocus()
   txtArev8 = CheckLen(txtArev8, 12)
   txtArev8 = GetAccount(txtArev8)
   lblArev8 = sAcctDesc
   
End Sub


Private Sub txtCdcc_Click()
   txtCdcc = GetAccount(txtCdcc)
   lblCdcc = "" & sAcctDesc
   
End Sub

Private Sub txtCdcc_LostFocus()
   txtCdcc = CheckLen(txtCdcc, 12)
   txtCdcc = GetAccount(txtCdcc)
   lblCdcc = "" & sAcctDesc
   
End Sub


Private Sub txtCdds_Click()
   txtCdds = GetAccount(txtCdds)
   lblCdds = "" & sAcctDesc
   
End Sub

Private Sub txtCdds_LostFocus()
   txtCdds = CheckLen(txtCdds, 12)
   txtCdds = GetAccount(txtCdds)
   lblCdds = "" & sAcctDesc
   
End Sub


Private Sub txtCdxc_Click()
   txtCdxc = GetAccount(txtCdxc)
   lblCdxc = "" & sAcctDesc
   
End Sub

Private Sub txtCdxc_LostFocus()
   txtCdxc = CheckLen(txtCdxc, 12)
   txtCdxc = GetAccount(txtCdxc)
   lblCdxc = "" & sAcctDesc
   
End Sub


Private Sub txtCogs_Click()
   txtCogs = GetAccount(txtCogs)
   lblCogs = sAcctDesc
   
End Sub


Private Sub txtCogs_LostFocus()
   txtCogs = CheckLen(txtCogs, 12)
   txtCogs = GetAccount(txtCogs)
   lblCogs = sAcctDesc
   
End Sub


Private Sub txtCrcm_Click()
   txtCrcm = GetAccount(txtCrex)
   lblCrcm = sAcctDesc
   
End Sub

Private Sub txtCrcm_LostFocus()
   txtCrcm = CheckLen(txtCrcm, 12)
   txtCrcm = GetAccount(txtCrcm)
   lblCrcm = sAcctDesc
   
End Sub


Private Sub txtCrcs_Click()
   txtCrcs = GetAccount(txtCrcs)
   lblCrcs = sAcctDesc
   
End Sub

Private Sub txtCrcs_LostFocus()
   txtCrcs = CheckLen(txtCrcs, 12)
   txtCrcs = GetAccount(txtCrcs)
   lblCrcs = sAcctDesc
   
End Sub


Private Sub txtCrds_Click()
   txtCrds = GetAccount(txtCrds)
   lblCrds = sAcctDesc
   
End Sub

Private Sub txtCrds_LostFocus()
   txtCrds = CheckLen(txtCrds, 12)
   txtCrds = GetAccount(txtCrds)
   lblCrds = sAcctDesc
   
End Sub


Private Sub txtCrex_Click()
   txtCrex = GetAccount(txtCrex)
   lblCrex = sAcctDesc
   
End Sub

Private Sub txtCrex_LostFocus()
   txtCrex = CheckLen(txtCrex, 12)
   txtCrex = GetAccount(txtCrex)
   lblCrex = sAcctDesc
   
End Sub


Private Sub txtCrrv_Click()
   txtCrrv = GetAccount(txtCrrv)
   lblCrrv = sAcctDesc
   
End Sub

Private Sub txtCrrv_LostFocus()
   txtCrrv = CheckLen(txtCrrv, 12)
   txtCrrv = GetAccount(txtCrrv)
   lblCrrv = sAcctDesc
   
End Sub


Private Sub txtDefFlow_LostFocus()
   On Error Resume Next
   txtDefFlow = CheckLen(txtDefFlow, 3)
   txtDefFlow = Format(Abs(Val(txtDefFlow)), "##0")
   sSql = "UPDATE Preferences SET DEFFLOWTIME=" & Val(Trim(txtDefFlow)) & " " _
          & "WHERE PreRecord=1"
   clsADOCon.ExecuteSql sSql
   
   
End Sub


Private Sub txtDefLabor_Click()
   txtDefLabor = GetAccount(txtDefLabor)
   lblDefLabor = sAcctDesc
   
End Sub


Private Sub txtDefLabor_LostFocus()
   txtDefLabor = CheckLen(txtDefLabor, 12)
   txtDefLabor = GetAccount(txtDefLabor)
   lblDefLabor = sAcctDesc
   
End Sub


Private Sub txtDefLead_LostFocus()
   On Error Resume Next
   txtDefLead = CheckLen(txtDefLead, 3)
   txtDefLead = Format(Abs(Val(txtDefLead)), "##0")
   sSql = "UPDATE Preferences SET DEFLEADTIME=" & Val(Trim(txtDefLead)) & " " _
          & "WHERE PreRecord=1"
   clsADOCon.ExecuteSql sSql
   
End Sub


Private Sub txtDefTime_Click()
   txtDefTime = GetAccount(txtDefTime)
   lblDefTime = sAcctDesc
   
End Sub


Private Sub txtDefTime_LostFocus()
   txtDefTime = CheckLen(txtDefTime, 12)
   txtDefTime = GetAccount(txtDefTime)
   lblDefTime = sAcctDesc
   
End Sub


Private Sub txtDivEnd_LostFocus()
   txtDivEnd = CheckLen(txtDivEnd, 2)
   txtDivEnd = Format(Abs(Val(txtDivEnd)), "#0")
   If Val(txtDivStart) > Val(txtDivEnd) Then
      'Beep
      txtDivStart = "0"
      txtDivEnd = "0"
   Else
      If Val(txtDivEnd) > 12 Then txtDivEnd = "12"
   End If
   
End Sub


Private Sub txtDivStart_LostFocus()
   txtDivStart = CheckLen(txtDivStart, 2)
   txtDivStart = Format(Abs(Val(txtDivStart)), "#0")
   If Val(txtDivStart) > 12 Then txtDivStart = "0"
   
End Sub


Private Sub txtDummy_Change()
   'Just a dummy
   
End Sub

Private Sub txtEngineeringRate_LostFocus()
   On Error Resume Next
   CheckCurrencyTextBox txtEngineeringRate, False
   'txtDefFlow = Format(Abs(Val(txtDefFlow)), "##0.00")
   sSql = "UPDATE Preferences SET EngineeringLaborRate = " & CCur(txtEngineeringRate) & " " _
          & "WHERE PreRecord=1"
   clsADOCon.ExecuteSql sSql
End Sub

Private Sub txtFax_LostFocus()
   txtFax = CheckLen(txtFax, 14)
   
End Sub

Private Sub txtFgi_Click()
   txtFgi = GetAccount(txtFgi)
   lblFgi = sAcctDesc
   
   
End Sub

Private Sub txtFtx_Click()
   txtFtx = GetAccount(txtFtx)
   lblFtx = sAcctDesc
   
End Sub

Private Sub txtFtx_LostFocus()
   txtFtx = CheckLen(txtFtx, 12)
   txtFtx = GetAccount(txtFtx)
   lblFtx = sAcctDesc
   
End Sub

Private Sub txtIntf_Change()
   If Len(txtIntf) > 3 Then SendKeys "{tab}"
   
End Sub

Private Sub txtIntf_LostFocus()
   txtIntf = CheckLen(txtIntf, 4)
   txtIntf = Format(Abs(Val(txtIntf)), "###")
   
End Sub


Private Sub txtIntp_Change()
   If Len(txtIntp) > 3 Then SendKeys "{tab}"
   
End Sub

Private Sub txtIntp_LostFocus()
   txtIntp = CheckLen(txtIntp, 4)
   txtIntp = Format(Abs(Val(txtIntp)), "###")
   
End Sub

Private Sub txtLastPon_LostFocus()
   txtLastPon = CheckLen(txtLastPon, 6)
   txtLastPon = Format(Abs(Val(txtLastPon)), "000000")
   
End Sub

Private Sub txtLastSon_LostFocus()
   txtLastSon = CheckLen(txtLastSon, 7)
   If Len(txtLastSon) < 7 Then _
          MsgBox "(7) Characters Please. Like S001234.", _
          vbInformation, Caption
   If Asc(Left(txtLastSon, 1)) < 65 Or Asc(Left(txtLastSon, 1)) > 89 Then
      MsgBox "Requires A Leading Character (A-Z) Like S001234.", _
         vbInformation, Caption
   End If
End Sub

Private Sub txtLinesPerStub_LostFocus()
   If Not IsNumeric(txtLinesPerStub) Then
      MsgBox "Lines per check stub must be 1 - 20"
      txtLinesPerStub.SetFocus
   End If
   
   If CInt(txtLinesPerStub) < 1 Or CInt(txtLinesPerStub > 20) Then
      MsgBox "Lines per check stub must be 1 - 20"
      txtLinesPerStub.SetFocus
   End If
End Sub

Private Sub txtMos_Click()
   txtMos = GetAccount(txtMos)
   lblMos = sAcctDesc
   
End Sub


Private Sub txtMos_LostFocus()
   txtMos = CheckLen(txtMos, 12)
   txtMos = GetAccount(txtMos)
   lblMos = sAcctDesc
   
End Sub


Private Sub txtNme_LostFocus()
   txtNme = CheckLen(txtNme, 50)
   txtNme = StrCase(txtNme)
   
   
End Sub

Private Sub txtPaap_Click()
   txtPaap = GetAccount(txtPaap)
   lblPaap = "" & sAcctDesc
   
End Sub

Private Sub txtPaap_LostFocus()
   txtPaap = CheckLen(txtPaap, 12)
   txtPaap = GetAccount(txtPaap)
   lblPaap = "" & sAcctDesc
   
End Sub


Private Sub txtPackPrefix_LostFocus()
   'reformat ps number to appropriate number of digits (8 - length of prefix)
   txtPackPrefix = Trim(txtPackPrefix)
   FormatPSNumber CLng("0" & txtPsl)
End Sub

Private Sub txtPant_Click()
   txtPant = GetAccount(txtPant)
   lblPant = "" & sAcctDesc
   
End Sub

Private Sub txtPant_LostFocus()
   txtPant = CheckLen(txtPant, 12)
   txtPant = GetAccount(txtPant)
   lblPant = "" & sAcctDesc
   
End Sub


Private Sub txtPatf_Click()
   txtPatf = GetAccount(txtPatf)
   lblPatf = "" & sAcctDesc
   
End Sub

Private Sub txtPatf_LostFocus()
   txtPatf = CheckLen(txtPatf, 12)
   txtPatf = GetAccount(txtPatf)
   lblPatf = "" & sAcctDesc
   
End Sub


Private Sub txtPatx_Click()
   txtPatx = GetAccount(txtPatx)
   lblPatx = "" & sAcctDesc
   
End Sub

Private Sub txtPatx_LostFocus()
   txtPatx = CheckLen(txtPatx, 12)
   txtPatx = GetAccount(txtPatx)
   lblPatx = "" & sAcctDesc
   
End Sub


Private Sub txtPhn_LostFocus()
   txtPhn = CheckLen(txtPhn, 14)
   
End Sub

Private Sub txtPsl_LostFocus()
'   txtPsl = CheckLen(txtPsl, 6)
'   txtPsl = Format(Abs(Val(txtPsl)), "000000")
   
   If Not IsNumeric(txtPsl) Then
      MsgBox "Packing slip number must be numeric"
      txtPsl.SetFocus
      Exit Sub
   End If
   FormatPSNumber CLng("0" & txtPsl)
End Sub


Private Sub txtQNM_LostFocus()
   If Val(txtQNM) > 24 Then txtQNM = "24"
   If Val(txtQNM) < 8 Then txtQNM = "8"
   txtQNM = Format(Abs(Val(txtQNM)), "##0")
   sSql = "UPDATE Preferences SET QueueMoveConversion=" & txtQNM & " WHERE PreRecord=1"
   clsADOCon.ExecuteSql sSql
   
End Sub



Private Sub txtResaleNo_LostFocus()
  txtResaleNo = CheckLen(txtResaleNo, 20)
End Sub

Private Sub txtSjar_Click()
   txtSjar = GetAccount(txtSjar)
   lblSjar = sAcctDesc
   
End Sub

Private Sub txtSjar_LostFocus()
   txtSjar = CheckLen(txtSjar, 12)
   txtSjar = GetAccount(txtSjar)
   lblSjar = sAcctDesc
   
End Sub


Private Sub txtSjnt_Click()
   txtSjnt = GetAccount(txtSjnt)
   lblSjnt = sAcctDesc
   
End Sub

Private Sub txtSjnt_LostFocus()
   txtSjnt = CheckLen(txtSjnt, 12)
   txtSjnt = GetAccount(txtSjnt)
   lblSjnt = sAcctDesc
   
End Sub


Private Sub txtSjtf_Click()
   txtSjtf = GetAccount(txtSjtf)
   lblSjtf = sAcctDesc
   
End Sub

Private Sub txtSjtf_LostFocus()
   txtSjtf = CheckLen(txtSjtf, 12)
   txtSjtf = GetAccount(txtSjtf)
   lblSjtf = sAcctDesc
   
End Sub


Private Sub txtSjtp_Click()
   txtSjtp = GetAccount(txtSjtp)
   lblSjtp = sAcctDesc
   
End Sub

Private Sub txtSjtp_LostFocus()
   txtSjtp = CheckLen(txtSjtp, 12)
   txtSjtp = GetAccount(txtSjtp)
   lblSjtp = sAcctDesc
   
End Sub


Private Sub txtSplit_LostFocus()
   txtSplit = CheckLen(txtSplit, 4)
   txtSplit = Format(Abs(Val(txtSplit)), "####0")
   If Val(txtSplit) = 0 Then txtSplit = 900
   
End Sub


Private Sub txtTid_LostFocus()
   txtTid = CheckLen(txtTid, 20)
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   'txtDummy.BackColor = Es_FormBackColor
   txtTransfer.BackColor = Es_TextDisabled
   
End Sub


Private Sub FillCombos()
   Dim RdoGlm As ADODB.Recordset
   Dim iList As Integer
   
   On Error GoTo DiaErr1
   sSql = "Qry_FillLowAccounts"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGlm, ES_FORWARD)
   If bSqlRows Then
      With RdoGlm
         Do Until .EOF
            iList = iList + 1
            AddComboStr txtArev1.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtArev2.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtArev3.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtArev4.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtArev5.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtArev6.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtArev7.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtArev8.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtFgi.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtCogs.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtMos.hwnd, "" & Trim(!GLACCTNO)
            
            'Receivables
            AddComboStr txtCrcs.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtCrds.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtCrex.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtCrcm.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtCrrv.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtSjar.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtSjtp.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtSjtf.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtSjnt.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtFtx.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtTrns.hwnd, "" & Trim(!GLACCTNO)
            
            'Payables
            AddComboStr txtPaap.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtPatx.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtPatf.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtPant.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtCdxc.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtCdcc.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtCdds.hwnd, "" & Trim(!GLACCTNO)
            
            'Labor
            AddComboStr txtDefTime.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtDefLabor.hwnd, "" & Trim(!GLACCTNO)
            .MoveNext
         Loop
         ClearResultSet RdoGlm
      End With
   End If
   Set RdoGlm = Nothing
   On Error Resume Next
   txtNme.SetFocus
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombos"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetAccount(sAccount) As String
   Dim RdoGlm As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_GetAccount '" & Compress(sAccount) & "'"
   If sAccount <> "" Then
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoGlm, ES_FORWARD)
      If bSqlRows Then
         With RdoGlm
            GetAccount = "" & Trim(!GLACCTNO)
            sAcctDesc = "" & Trim(!GLDESCR)
         End With
      Else
         'Beep
         GetAccount = ""
         sAcctDesc = ""
      End If
   Else
      GetAccount = ""
      sAcctDesc = ""
   End If
   Set RdoGlm = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getaccoun"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Function CheckJournals() As Byte
   Dim rdoJrn As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT MJTYPE FROM JrhdTable "
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn, ES_FORWARD)
   If bSqlRows Then
      CheckJournals = 1
      ClearResultSet rdoJrn
   Else
      CheckJournals = 0
   End If
   Set rdoJrn = Nothing
   Exit Function
   
DiaErr1:
   CheckJournals = 0
   
End Function


Private Sub GetShopDefaults()
   Dim RdoShp As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DEFFLOWTIME,DEFLEADTIME,AllowMOQuantityChanges " _
          & "FROM Preferences WHERE PreRecord=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoShp, ES_FORWARD)
   If bSqlRows Then
      With RdoShp
         txtDefFlow = Format(!DEFFLOWTIME, "##0")
         txtDefLead = Format(!DEFLEADTIME, "##0")
         optQtyChg.Value = !AllowMOQuantityChanges
         ClearResultSet RdoShp
      End With
   End If
   Set RdoShp = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getshopdef"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtTrns_Click()
   txtTrns = GetAccount(txtTrns)
   lblTrns = sAcctDesc
   
End Sub


Private Sub txtTrns_LostFocus()
   txtTrns = CheckLen(txtTrns, 12)
   txtTrns = GetAccount(txtTrns)
   lblTrns = sAcctDesc
   
End Sub

Private Sub FormatPSNumber(psno As Long)
   Dim ps As String
   ps = Format(psno, "00000000")
   txtPsl = ""
   txtPsl.MaxLength = 8 - Len(Trim(txtPackPrefix))
   txtPsl = Mid(ps, 1 + Len(Trim(txtPackPrefix)), txtPsl.MaxLength)
End Sub

