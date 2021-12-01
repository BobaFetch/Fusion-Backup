VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form CyclCYe01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create ABC Classes"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CyclCYe01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   72
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optShow 
      Alignment       =   1  'Right Justify
      Caption         =   "Show Only Used Rows"
      Height          =   255
      Left            =   240
      TabIndex        =   71
      ToolTipText     =   "Check This And Click Reorder (Requires Some Used Rows)"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CheckBox optHelp 
      Caption         =   "Help"
      Height          =   255
      Left            =   1440
      TabIndex        =   70
      Top             =   5280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame z2 
      Height          =   70
      Index           =   1
      Left            =   120
      TabIndex        =   67
      Top             =   1080
      Width           =   5415
   End
   Begin VB.Frame z2 
      Height          =   70
      Index           =   0
      Left            =   120
      TabIndex        =   66
      Top             =   5160
      Width           =   5415
   End
   Begin VB.CommandButton cmdLst 
      Caption         =   "<< &Last    "
      Height          =   315
      Left            =   3800
      TabIndex        =   65
      ToolTipText     =   "Last Page"
      Top             =   5280
      Width           =   875
   End
   Begin VB.CommandButton cmdNxt 
      Caption         =   " &Next >>"
      Height          =   315
      Left            =   4680
      TabIndex        =   64
      ToolTipText     =   "Next Page"
      Top             =   5280
      Width           =   875
   End
   Begin VB.TextBox txtFrq 
      Height          =   285
      Index           =   10
      Left            =   960
      TabIndex        =   62
      ToolTipText     =   "Frequency Of Count (Days)"
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox txtLCost 
      Height          =   285
      Index           =   10
      Left            =   1920
      TabIndex        =   61
      ToolTipText     =   "Bottom End (Min) Standard Cost Of The Item"
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox txtHCost 
      Height          =   285
      Index           =   10
      Left            =   3240
      TabIndex        =   60
      ToolTipText     =   "Top End (Max) Standard Cost Of The Item"
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CheckBox optUsed 
      Alignment       =   1  'Right Justify
      Caption         =   "____"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   10
      Left            =   4680
      TabIndex        =   59
      ToolTipText     =   "Mark This Code To Be Used"
      Top             =   4800
      Width           =   735
   End
   Begin VB.CheckBox optInit 
      Caption         =   "Initialized"
      Height          =   255
      Left            =   120
      TabIndex        =   58
      Top             =   5280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtFrq 
      Height          =   285
      Index           =   9
      Left            =   960
      TabIndex        =   56
      ToolTipText     =   "Frequency Of Count (Days)"
      Top             =   4480
      Width           =   735
   End
   Begin VB.TextBox txtLCost 
      Height          =   285
      Index           =   9
      Left            =   1920
      TabIndex        =   55
      ToolTipText     =   "Bottom End (Min) Standard Cost Of The Item"
      Top             =   4480
      Width           =   1095
   End
   Begin VB.TextBox txtHCost 
      Height          =   285
      Index           =   9
      Left            =   3240
      TabIndex        =   54
      ToolTipText     =   "Top End (Max) Standard Cost Of The Item"
      Top             =   4480
      Width           =   1095
   End
   Begin VB.CheckBox optUsed 
      Alignment       =   1  'Right Justify
      Caption         =   "____"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   9
      Left            =   4680
      TabIndex        =   53
      ToolTipText     =   "Mark This Code To Be Used"
      Top             =   4480
      Width           =   735
   End
   Begin VB.TextBox txtFrq 
      Height          =   285
      Index           =   8
      Left            =   960
      TabIndex        =   51
      ToolTipText     =   "Frequency Of Count (Days)"
      Top             =   4160
      Width           =   735
   End
   Begin VB.TextBox txtLCost 
      Height          =   285
      Index           =   8
      Left            =   1920
      TabIndex        =   50
      ToolTipText     =   "Bottom End (Min) Standard Cost Of The Item"
      Top             =   4160
      Width           =   1095
   End
   Begin VB.TextBox txtHCost 
      Height          =   285
      Index           =   8
      Left            =   3240
      TabIndex        =   49
      ToolTipText     =   "Top End (Max) Standard Cost Of The Item"
      Top             =   4160
      Width           =   1095
   End
   Begin VB.CheckBox optUsed 
      Alignment       =   1  'Right Justify
      Caption         =   "____"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   4680
      TabIndex        =   48
      ToolTipText     =   "Mark This Code To Be Used"
      Top             =   4160
      Width           =   735
   End
   Begin VB.TextBox txtFrq 
      Height          =   285
      Index           =   7
      Left            =   960
      TabIndex        =   46
      ToolTipText     =   "Frequency Of Count (Days)"
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox txtLCost 
      Height          =   285
      Index           =   7
      Left            =   1920
      TabIndex        =   45
      ToolTipText     =   "Bottom End (Min) Standard Cost Of The Item"
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtHCost 
      Height          =   285
      Index           =   7
      Left            =   3240
      TabIndex        =   44
      ToolTipText     =   "Top End (Max) Standard Cost Of The Item"
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CheckBox optUsed 
      Alignment       =   1  'Right Justify
      Caption         =   "____"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   4680
      TabIndex        =   43
      ToolTipText     =   "Mark This Code To Be Used"
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox txtFrq 
      Height          =   285
      Index           =   6
      Left            =   960
      TabIndex        =   41
      ToolTipText     =   "Frequency Of Count (Days)"
      Top             =   3520
      Width           =   735
   End
   Begin VB.TextBox txtLCost 
      Height          =   285
      Index           =   6
      Left            =   1920
      TabIndex        =   40
      ToolTipText     =   "Bottom End (Min) Standard Cost Of The Item"
      Top             =   3520
      Width           =   1095
   End
   Begin VB.TextBox txtHCost 
      Height          =   285
      Index           =   6
      Left            =   3240
      TabIndex        =   39
      ToolTipText     =   "Top End (Max) Standard Cost Of The Item"
      Top             =   3520
      Width           =   1095
   End
   Begin VB.CheckBox optUsed 
      Alignment       =   1  'Right Justify
      Caption         =   "____"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   4680
      TabIndex        =   38
      ToolTipText     =   "Mark This Code To Be Used"
      Top             =   3520
      Width           =   735
   End
   Begin VB.TextBox txtFrq 
      Height          =   285
      Index           =   5
      Left            =   960
      TabIndex        =   36
      ToolTipText     =   "Frequency Of Count (Days)"
      Top             =   3200
      Width           =   735
   End
   Begin VB.TextBox txtLCost 
      Height          =   285
      Index           =   5
      Left            =   1920
      TabIndex        =   35
      ToolTipText     =   "Bottom End (Min) Standard Cost Of The Item"
      Top             =   3200
      Width           =   1095
   End
   Begin VB.TextBox txtHCost 
      Height          =   285
      Index           =   5
      Left            =   3240
      TabIndex        =   34
      ToolTipText     =   "Top End (Max) Standard Cost Of The Item"
      Top             =   3200
      Width           =   1095
   End
   Begin VB.CheckBox optUsed 
      Alignment       =   1  'Right Justify
      Caption         =   "____"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   4680
      TabIndex        =   33
      ToolTipText     =   "Mark This Code To Be Used"
      Top             =   3200
      Width           =   735
   End
   Begin VB.TextBox txtFrq 
      Height          =   285
      Index           =   4
      Left            =   960
      TabIndex        =   31
      ToolTipText     =   "Frequency Of Count (Days)"
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txtLCost 
      Height          =   285
      Index           =   4
      Left            =   1920
      TabIndex        =   30
      ToolTipText     =   "Bottom End (Min) Standard Cost Of The Item"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtHCost 
      Height          =   285
      Index           =   4
      Left            =   3240
      TabIndex        =   29
      ToolTipText     =   "Top End (Max) Standard Cost Of The Item"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CheckBox optUsed 
      Alignment       =   1  'Right Justify
      Caption         =   "____"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   4680
      TabIndex        =   28
      ToolTipText     =   "Mark This Code To Be Used"
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txtFrq 
      Height          =   285
      Index           =   3
      Left            =   960
      TabIndex        =   26
      ToolTipText     =   "Frequency Of Count (Days)"
      Top             =   2560
      Width           =   735
   End
   Begin VB.TextBox txtLCost 
      Height          =   285
      Index           =   3
      Left            =   1920
      TabIndex        =   25
      ToolTipText     =   "Bottom End (Min) Standard Cost Of The Item"
      Top             =   2560
      Width           =   1095
   End
   Begin VB.TextBox txtHCost 
      Height          =   285
      Index           =   3
      Left            =   3240
      TabIndex        =   24
      ToolTipText     =   "Top End (Max) Standard Cost Of The Item"
      Top             =   2560
      Width           =   1095
   End
   Begin VB.CheckBox optUsed 
      Alignment       =   1  'Right Justify
      Caption         =   "____"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   4680
      TabIndex        =   23
      ToolTipText     =   "Mark This Code To Be Used"
      Top             =   2560
      Width           =   735
   End
   Begin VB.TextBox txtFrq 
      Height          =   285
      Index           =   2
      Left            =   960
      TabIndex        =   21
      ToolTipText     =   "Frequency Of Count (Days)"
      Top             =   2240
      Width           =   735
   End
   Begin VB.TextBox txtLCost 
      Height          =   285
      Index           =   2
      Left            =   1920
      TabIndex        =   20
      ToolTipText     =   "Bottom End (Min) Standard Cost Of The Item"
      Top             =   2240
      Width           =   1095
   End
   Begin VB.TextBox txtHCost 
      Height          =   285
      Index           =   2
      Left            =   3240
      TabIndex        =   19
      ToolTipText     =   "Top End (Max) Standard Cost Of The Item"
      Top             =   2240
      Width           =   1095
   End
   Begin VB.CheckBox optUsed 
      Alignment       =   1  'Right Justify
      Caption         =   "____"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   18
      ToolTipText     =   "Mark This Code To Be Used"
      Top             =   2240
      Width           =   735
   End
   Begin VB.TextBox txtFrq 
      Height          =   285
      Index           =   1
      Left            =   960
      TabIndex        =   16
      ToolTipText     =   "Frequency Of Count (Days)"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox txtLCost 
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   15
      ToolTipText     =   "Bottom End (Min) Standard Cost Of The Item"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtHCost 
      Height          =   285
      Index           =   1
      Left            =   3240
      TabIndex        =   14
      ToolTipText     =   "Top End (Max) Standard Cost Of The Item"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CheckBox optUsed 
      Alignment       =   1  'Right Justify
      Caption         =   "____"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   13
      ToolTipText     =   "Mark This Code To Be Used"
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmdReorder 
      Caption         =   "R&eorder"
      Height          =   315
      Left            =   3720
      TabIndex        =   12
      ToolTipText     =   "Reorders Showing Used Classes First.  Saves Changes With Little Testing."
      Top             =   1200
      Width           =   875
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   4680
      TabIndex        =   11
      ToolTipText     =   "Update The Rows With The Current Information."
      Top             =   1200
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4680
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   5640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5685
      FormDesignWidth =   5640
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Page"
      Height          =   255
      Index           =   7
      Left            =   2640
      TabIndex        =   69
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label lblPage 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   255
      Left            =   3120
      TabIndex        =   68
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label lblCode 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   240
      TabIndex        =   63
      ToolTipText     =   "ABC Code"
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label lblCode 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   240
      TabIndex        =   57
      ToolTipText     =   "ABC Code"
      Top             =   4480
      Width           =   495
   End
   Begin VB.Label lblCode 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   52
      ToolTipText     =   "ABC Code"
      Top             =   4160
      Width           =   495
   End
   Begin VB.Label lblCode 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   47
      ToolTipText     =   "ABC Code"
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label lblCode 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   42
      ToolTipText     =   "ABC Code"
      Top             =   3520
      Width           =   495
   End
   Begin VB.Label lblCode 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   37
      ToolTipText     =   "ABC Code"
      Top             =   3200
      Width           =   495
   End
   Begin VB.Label lblCode 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   32
      ToolTipText     =   "ABC Code"
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lblCode 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   27
      ToolTipText     =   "ABC Code"
      Top             =   2560
      Width           =   495
   End
   Begin VB.Label lblCode 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   22
      ToolTipText     =   "ABC Code"
      Top             =   2240
      Width           =   495
   End
   Begin VB.Label lblCode 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   17
      ToolTipText     =   "ABC Code"
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      ToolTipText     =   "Establishing ABC Class Codes And Values Initializes ABC Functions"
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label lblHigh 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4560
      TabIndex        =   9
      ToolTipText     =   "From The Company Setup"
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblLow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      ToolTipText     =   "From The Company Setup"
      Top             =   720
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Your High Value Is:"
      Height          =   255
      Index           =   6
      Left            =   3000
      TabIndex        =   7
      ToolTipText     =   "From The Company Setup"
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Low Value Is:"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "From The Company Setup"
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Used        "
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
      Left            =   4680
      TabIndex        =   5
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "High Limit Cost"
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
      Left            =   3240
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Low Limit Cost"
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
      Left            =   1920
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Frequency"
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
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Class  "
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
      TabIndex        =   1
      Top             =   1680
      Width           =   615
   End
End
Attribute VB_Name = "CyclCYe01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'New 11/10/03
'6/30/05 Added TestAbc
Option Explicit
Dim bFilling As Byte
Dim bOnLoad As Byte
Dim bCurrIdx As Byte
Dim bSaved As Byte

Dim iTotalRows As Integer
Dim cLowCost As Currency
Dim cHighCost As Currency

Dim cSettings(79, 5) As Currency
'   0 = Row
'   1 = Frequency
'   2 = Low Cost
'   3 = High Cost
'   4 = Used
Dim sCode(79) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmdCan_Click()
   Dim bByte As Byte
   Dim bResponse As Byte
   Dim sMsg As String
   If bSaved = 1 Then
      bResponse = CheckCosts()
      If optInit.Value = vbUnchecked Then
         sMsg = "Your Class Codes Have Not Been Initialized. You May" & vbCr _
                & "Loose Your Work Do You Wish To Quit Anyway?"
         bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
         If bResponse = vbYes Then Unload Me _
                        Else bByte = 0
      Else
         bByte = 0
      End If
      bResponse = CheckFrequencies()
      If bResponse > 0 Then
         If bResponse = 1 Then
            sMsg = bResponse & " Frequency Is Not Set. " & vbCr _
                   & "Do You Wish To Quit Anyway?"
            bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
         Else
            sMsg = bResponse & " Frequencies Are Not Set. " & vbCr _
                   & "Do You Wish To Quit Anyway?"
            bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
         End If
         If bResponse = vbYes Then bByte = 1 Else bByte = 0
      Else
         bByte = 1
      End If
   Else
      sMsg = "You've Made Changes And Not Saved The Work  " & vbCr _
             & "And May Loose It. Continue To Exit Anyway?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then bByte = 1
   End If
   If bByte = 1 Then Unload Me
   
End Sub



Private Sub cmdHlp_Click()
   If cmdHlp Then
      optHelp.Value = vbChecked
      OpenHelpContext 1604
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdLst_Click()
   If bCurrIdx > 0 Then bCurrIdx = bCurrIdx - 10
   GetNextGroup
   
End Sub

Private Sub cmdNxt_Click()
   bCurrIdx = bCurrIdx + 10
   If bCurrIdx > 70 Then bCurrIdx = 70
   GetNextGroup
   
End Sub


Private Sub cmdReorder_Click()
   Dim bResponse As Byte
   bResponse = MsgBox("Reorder The Group?", _
               ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      cmdReorder.Enabled = False
      For bResponse = 1 To iTotalRows
         On Error Resume Next
         sSql = "UPDATE CabcTable SET COABCUSED=" _
                & cSettings(bResponse, 4) & "," _
                & "COABCFREQUENCY=" & cSettings(bResponse, 1) & "," _
                & "COABCLOWCOST=" & cSettings(bResponse, 2) & "," _
                & "COABCHIGHCOST=" & cSettings(bResponse, 3) & " " _
                & "WHERE COABCROW=" & cSettings(bResponse, 0) & " "
         clsADOCon.ExecuteSQL sSql
      Next
      MouseCursor 0
      cmdReorder.Enabled = True
      GetCodes
   Else
      CancelTrans
   End If
   
End Sub

Private Sub cmdUpd_Click()
   Dim bResponse As Byte
   bResponse = CheckColumns()
   If bResponse > 0 Then
      MsgBox "The Selected Columns Aren't Correct. " & vbCr _
         & "You Must Check Your Work.", vbInformation, Caption
      optInit.Value = vbUnchecked
      sSql = "UPDATE Preferences SET CycleCountInitialized=0"
      clsADOCon.ExecuteSQL sSql
      GetSetup
      Exit Sub
   End If
   bResponse = CheckCosts()
   If bResponse = 0 Then
      MsgBox "There Are No Costs Attributed To The Group.", _
         vbInformation, Caption
   Else
      If optInit.Value = vbChecked Then
         bResponse = MsgBox("Would You Like To Update The ABC Classes?", _
                     ES_YESQUESTION, Caption)
      Else
         bResponse = MsgBox("Would You Like To Update The ABC Classes" & vbCr _
                     & "And Initialize Them?", ES_YESQUESTION, Caption)
      End If
      If bResponse = vbYes Then
         On Error Resume Next
         MouseCursor 13
         cmdUpd.Enabled = False
         On Error Resume Next
         clsADOCon.ADOErrNum = 0
         For bResponse = 1 To iTotalRows
            sSql = "UPDATE CabcTable SET COABCUSED=" _
                   & Format(cSettings(bResponse, 4), "0") & "," _
                   & "COABCFREQUENCY=" & cSettings(bResponse, 1) & "," _
                   & "COABCLOWCOST=" & cSettings(bResponse, 2) & "," _
                   & "COABCHIGHCOST=" & cSettings(bResponse, 3) & " " _
                   & "WHERE COABCROW=" & Format(cSettings(bResponse, 0), "0") & " "
            clsADOCon.ExecuteSQL sSql
         Next
         cmdUpd.Enabled = True
         If clsADOCon.ADOErrNum = 0 Then
            optInit.Value = vbChecked
            sSql = "UPDATE Preferences SET CycleCountInitialized=1"
            clsADOCon.ExecuteSQL sSql
            sSql = "UPDATE CabcTable SET " _
                   & "COABCFREQUENCY=0," _
                   & "COABCLOWCOST=0," _
                   & "COABCHIGHCOST=0" _
                   & "WHERE COABCUSED=0"
            clsADOCon.ExecuteSQL sSql
            Sleep 500
            GetSetup
            MouseCursor 0
            MsgBox "The Group Has Been Updated.", vbInformation, _
               Caption
            bSaved = 1
            GetCodes
         Else
            MouseCursor 0
            MsgBox "Couldn't Update The Group.", vbInformation, _
               Caption
         End If
      Else
         CancelTrans
      End If
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      TestAbc
      bSaved = 1
      GetSetup
      GetCodes
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   GetOptions
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim bByte As Byte
   SaveOptions
   On Error Resume Next
   If optInit.Value = vbUnchecked Then
      For bByte = 1 To iTotalRows
         sSql = "UPDATE CabcTable SET COABCUSED=" _
                & cSettings(bByte, 4) & "," _
                & "COABCFREQUENCY=0," _
                & "COABCLOWCOST=0," _
                & "COABCHIGHCOST=0" _
                & "WHERE COABCROW=" & cSettings(bByte, 0) & " "
         clsADOCon.ExecuteSQL sSql
      Next
      sSql = "UPDATE Preferences SET CycleCountInitialized=0"
      clsADOCon.ExecuteSQL sSql
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set CyclCYe01a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   Dim iIndex As Integer
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   lblStatus.ForeColor = ES_BLUE
   
   bFilling = 1
   For b = 1 To 10
      lblCode(b).Visible = False
      lblCode(b).ToolTipText = "ABC Class"
      txtFrq(b).Visible = False
      txtLCost(b).Visible = False
      txtHCost(b).Visible = False
      optUsed(b).Visible = False
      txtFrq(b) = "0"
      txtFrq(b).TabIndex = iIndex
      iIndex = iIndex + 1
      txtLCost(b) = "0.00"
      txtLCost(b).TabIndex = iIndex
      iIndex = iIndex + 1
      txtHCost(b) = "0.00"
      txtHCost(b).TabIndex = iIndex
      iIndex = iIndex + 1
      optUsed(b).TabIndex = iIndex
      iIndex = iIndex + 1
   Next
   bFilling = 0
   
End Sub

Private Sub GetSetup()
   Dim bByte As Byte
   Dim RdoSet As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "Qry_GetABCPreference"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSet, ES_FORWARD)
   If bSqlRows Then
      With RdoSet
         If Not IsNull(.Fields(0)) Then
            bByte = .Fields(0)
         Else
            bByte = 0
         End If
         ClearResultSet RdoSet
      End With
   End If
   If bByte = 0 Then
      lblStatus = "The ABC Class Setup Has Not Been Initialized"
      optInit.Value = vbUnchecked
   Else
      lblStatus = "The ABC Class Setup Has Been Initialized"
      optInit.Value = vbChecked
   End If
   sSql = "SELECT COABCLOWLIMITCOST,COABCHIGHLIMITCOST FROM " _
          & "ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSet, ES_FORWARD)
   If bSqlRows Then
      With RdoSet
         If Not IsNull(.Fields(0)) Then
            cLowCost = .Fields(0)
            cHighCost = .Fields(1)
            lblLow = Format(.Fields(0), "#0.00")
            lblHigh = Format(.Fields(1), "###,##0.00")
         Else
            lblLow = "0.00"
            lblHigh = lblLow
         End If
         ClearResultSet RdoSet
      End With
   End If
   Set RdoSet = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getsetup"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub optInit_Click()
   'never visible - used for checking
   
End Sub


Private Sub optUsed_Click(Index As Integer)
   If bFilling = 0 Then bSaved = 0
   If bFilling = 0 Then cSettings(Index + bCurrIdx, 4) = optUsed(Index).Value
   
End Sub

Private Sub optUsed_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optUsed_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtFrq_Change(Index As Integer)
   If bFilling = 0 Then bSaved = 0
   
End Sub

Private Sub txtFrq_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtFrq_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtFrq_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub


Private Sub txtFrq_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtFrq_LostFocus(Index As Integer)
   txtFrq(Index) = CheckLen(txtFrq(Index), 3)
   txtFrq(Index) = Format(Abs(Val(txtFrq(Index))), "##0")
   If Val(txtFrq(Index)) > 360 Then
      'SysSysbeep
      txtFrq(Index) = "360"
   End If
   If bFilling = 0 Then cSettings(Index + bCurrIdx, 1) = Val(txtFrq(Index))
   
End Sub

Private Sub txtHCost_Change(Index As Integer)
   If bFilling = 0 Then bSaved = 0
   
End Sub

Private Sub txtHCost_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtHCost_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtHCost_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub


Private Sub txtHCost_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtHCost_LostFocus(Index As Integer)
   txtHCost(Index) = CheckLen(txtHCost(Index), 10)
   If Val(txtHCost(Index)) > cHighCost Then
      'SysSysbeep
      txtHCost(Index) = cHighCost
   End If
   txtHCost(Index) = Format(Abs(Val(txtHCost(Index))), "#####0.00")
   If bFilling = 0 Then cSettings(Index + bCurrIdx, 3) = Val(txtHCost(Index))
   If Val(txtHCost(Index)) > 0 Then
      If Val(txtLCost(Index)) < cLowCost Then
         txtLCost(Index) = Format(cLowCost, "#####0.00")
         If bFilling = 0 Then cSettings(Index + bCurrIdx, 2) = Val(txtLCost(Index))
      End If
   End If
   
End Sub

Private Sub txtLCost_Change(Index As Integer)
   If bFilling = 0 Then bSaved = 0
   
End Sub

Private Sub txtLCost_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtLCost_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtLCost_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub


Private Sub txtLCost_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtLCost_LostFocus(Index As Integer)
   txtLCost(Index) = CheckLen(txtLCost(Index), 10)
   If Val(txtLCost(Index)) > 0 Then
      If Val(txtLCost(Index)) < cLowCost Then
         'SysSysbeep
         txtLCost(Index) = cLowCost
      End If
   End If
   txtLCost(Index) = Format(Abs(Val(txtLCost(Index))), "#####0.00")
   If bFilling = 0 Then cSettings(Index + bCurrIdx, 2) = Val(txtLCost(Index))
   
End Sub



Private Sub GetCodes()
   Dim RdoCde As ADODB.Recordset
   Dim iRow As Integer
   Erase cSettings
   Erase sCode
   iRow = 0
   'Fill With Selected
   On Error GoTo DiaErr1
   bFilling = 1
   sSql = "Qry_FillABCUsed"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCde, ES_FORWARD)
   If bSqlRows Then
      With RdoCde
         Do Until .EOF
            iRow = iRow + 1
            sCode(iRow) = "" & Trim(!COABCCODE)
            cSettings(iRow, 0) = !COABCROW
            cSettings(iRow, 1) = !COABCFREQUENCY
            cSettings(iRow, 2) = !COABCLOWCOST
            cSettings(iRow, 3) = !COABCHIGHCOST
            cSettings(iRow, 4) = !COABCUSED
            cSettings(iRow, 4) = Format(!COABCUSED, "0")
            .MoveNext
         Loop
         ClearResultSet RdoCde
      End With
   End If
   If iRow = 0 Then
      optShow.Enabled = False
      optShow.Value = vbUnchecked
   Else
      optShow.Enabled = True
   End If
   'Fill unselected
   If optShow.Value = vbUnchecked Then
      sSql = "Qry_FillABCNotUsed"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCde, ES_FORWARD)
      If bSqlRows Then
         With RdoCde
            Do Until .EOF
               iRow = iRow + 1
               sCode(iRow) = "" & Trim(!COABCCODE)
               cSettings(iRow, 0) = !COABCROW
               cSettings(iRow, 1) = !COABCFREQUENCY
               cSettings(iRow, 2) = !COABCLOWCOST
               cSettings(iRow, 3) = !COABCHIGHCOST
               cSettings(iRow, 4) = Format(!COABCUSED, "0")
               .MoveNext
            Loop
            ClearResultSet RdoCde
         End With
      End If
   End If
   iTotalRows = iRow
   If iTotalRows < 11 Then
      cmdNxt.Enabled = False
      cmdLst.Enabled = False
   End If
   For iRow = 1 To 9
      lblCode(iRow).Visible = False
      txtFrq(iRow).Visible = False
      txtLCost(iRow).Visible = False
      txtHCost(iRow).Visible = False
      optUsed(iRow).Visible = False
   Next
   lblCode(iRow).Visible = False
   txtFrq(iRow).Visible = False
   txtLCost(iRow).Visible = False
   txtHCost(iRow).Visible = False
   optUsed(iRow).Visible = False
   
   For iRow = 1 To 10
      If iRow > iTotalRows Then Exit For
      lblCode(iRow).Visible = True
      txtFrq(iRow).Visible = True
      txtLCost(iRow).Visible = True
      txtHCost(iRow).Visible = True
      optUsed(iRow).Visible = True
      lblCode(iRow) = sCode(iRow)
      txtFrq(iRow) = Format(cSettings(iRow, 1), "##0")
      txtLCost(iRow) = Format(cSettings(iRow, 2), "#####0.00")
      txtHCost(iRow) = Format(cSettings(iRow, 3), "#####0.00")
      optUsed(iRow).Value = Format(cSettings(iRow, 4), "0")
   Next
   If iTotalRows > 10 Then
      cmdLst.Enabled = True
      cmdNxt.Enabled = True
   End If
   bFilling = 0
   bCurrIdx = 0
   lblPage = 1
   On Error Resume Next
   If txtFrq(1).Enabled Then txtFrq(1).SetFocus
   Set RdoCde = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getcodes"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetNextGroup()
   Dim iRow As Integer
   On Error Resume Next
   bFilling = 1
   lblPage = (bCurrIdx / 10) + 1
   For iRow = 1 To 9
      lblCode(iRow).Visible = False
      txtFrq(iRow).Visible = False
      txtLCost(iRow).Visible = False
      txtHCost(iRow).Visible = False
      optUsed(iRow).Visible = False
   Next
   lblCode(iRow).Visible = False
   txtFrq(iRow).Visible = False
   txtLCost(iRow).Visible = False
   txtHCost(iRow).Visible = False
   optUsed(iRow).Visible = False
   
   For iRow = 1 To 10
      If iRow + bCurrIdx > iTotalRows Then Exit For
      If bCurrIdx = 70 And iRow > 7 Then Exit For
      lblCode(iRow).Visible = True
      txtFrq(iRow).Visible = True
      txtLCost(iRow).Visible = True
      txtHCost(iRow).Visible = True
      optUsed(iRow).Visible = True
      lblCode(iRow) = sCode(iRow + bCurrIdx)
      txtFrq(iRow) = Format(cSettings(iRow + bCurrIdx, 1), "##0")
      txtLCost(iRow) = Format(cSettings(iRow + bCurrIdx, 2), "#####0.00")
      txtHCost(iRow) = Format(cSettings(iRow + bCurrIdx, 3), "#####0.00")
      optUsed(iRow).Value = Format(cSettings(iRow + bCurrIdx, 4), "0")
   Next
   bFilling = 0
   If txtFrq(1).Enabled Then txtFrq(1).SetFocus
   
End Sub

Private Function CheckCosts() As Byte
   Dim bRow As Byte
   Dim cCost As Currency
   For bRow = 1 To 77
      cCost = cCost + (cSettings(bRow, 2) + cSettings(bRow, 3))
   Next
   cCost = cCost + (cSettings(bRow, 2) + cSettings(bRow, 3))
   If cCost > 0 Then
      CheckCosts = 1
   Else
      CheckCosts = 0
      optInit.Value = vbUnchecked
   End If
   
End Function


Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiInvc", "ABCSetup", optShow.Value
   
End Sub

Private Sub GetOptions()
   On Error Resume Next
   optShow.Value = Val(GetSetting("Esi2000", "EsiInvc", "ABCSetup", optShow.Value))
   
End Sub

Private Function CheckFrequencies() As Byte
   Dim bRow As Byte
   Dim iFreq As Integer
   Dim cCost As Currency
   For bRow = 1 To iTotalRows
      If cSettings(bRow, 4) > 0 Then
         cCost = cSettings(bRow, 2) + cSettings(bRow, 3)
         If cCost > 0 Then
            If cSettings(bRow, 1) = 0 Then iFreq = iFreq + 1
         End If
      End If
   Next
   CheckFrequencies = iFreq
   
   
   
End Function

Private Function CheckColumns() As Byte
   Dim bByte As Byte
   Dim bRow As Byte
   Dim cCost As Currency
   For bRow = 1 To iTotalRows
      If cSettings(bRow, 4) > 0 Then
         If bRow > 1 Then
            If (cSettings(bRow, 2) >= cSettings(bRow - 1, 3)) _
                Then bByte = 1
            End If
            If (cSettings(bRow, 2) >= cSettings(bRow, 3)) Then _
                bByte = 1
         End If
      Next
      CheckColumns = bByte
      
   End Function
   
   '6/30/05 In case ABC Codes aren't installed on initial Setup

   Public Sub TestAbc()
      Dim RdoAbc As ADODB.Recordset
      Dim iCode As Integer
      Dim iRow As Integer
      
      On Error GoTo DiaErr1
      sSql = "Qry_TestABCCodes"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoAbc, ES_FORWARD)
      If Not bSqlRows Then
         For iRow = 65 To 90
            iCode = iCode + 1
            sSql = "INSERT INTO dbo.CabcTable (COABCROW," _
                   & "COABCCODE) VALUES(" & iCode & ",'" _
                   & Chr$(iRow) & "+')"
            clsADOCon.ExecuteSQL sSql
            
            iCode = iCode + 1
            sSql = "INSERT INTO dbo.CabcTable (COABCROW," _
                   & "COABCCODE) VALUES(" & iCode & ",'" _
                   & Chr$(iRow) & "')"
            clsADOCon.ExecuteSQL sSql
            
            iCode = iCode + 1
            sSql = "INSERT INTO dbo.CabcTable (COABCROW," _
                   & "COABCCODE) VALUES(" & iCode & ",'" _
                   & Chr$(iRow) & "-')"
            clsADOCon.ExecuteSQL sSql
         Next
      Else
         ClearResultSet RdoAbc
      End If
      Set RdoAbc = Nothing
      Exit Sub
      
DiaErr1:
      sProcName = "testabc"
      CurrError.Number = Err.Number
      CurrError.Description = Err.Description
      DoModuleErrors Me
      
   End Sub
