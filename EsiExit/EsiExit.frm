VERSION 5.00
Begin VB.Form diaExit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Normal Workstation Shut Down"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   Icon            =   "EsiExit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   3360
      Width           =   135
   End
   Begin VB.CommandButton cmdHelp 
      Appearance      =   0  'Flat
      DownPicture     =   "EsiExit.frx":014A
      Height          =   280
      Left            =   4600
      Picture         =   "EsiExit.frx":06DC
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   300
   End
   Begin VB.Timer Timer3 
      Interval        =   32700
      Left            =   0
      Top             =   1920
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   0
      Top             =   2280
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   0
      Top             =   2640
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      ToolTipText     =   "Close And Allow Reset Of Sections"
      Top             =   0
      Width           =   855
   End
   Begin VB.Label lblShutDown 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   840
      TabIndex        =   18
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label OffTime 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   3000
      TabIndex        =   17
      ToolTipText     =   "Time Closed"
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Section 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   840
      TabIndex        =   16
      ToolTipText     =   "Section Closed"
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label OffTime 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   3000
      TabIndex        =   15
      ToolTipText     =   "Time Closed"
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label OffTime 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   3000
      TabIndex        =   14
      ToolTipText     =   "Time Closed"
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label OffTime 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   3000
      TabIndex        =   13
      ToolTipText     =   "Time Closed"
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label OffTime 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   12
      ToolTipText     =   "Time Closed"
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label OffTime 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   11
      ToolTipText     =   "Time Closed"
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label OffTime 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   10
      ToolTipText     =   "Time Closed"
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Section 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   840
      TabIndex        =   9
      ToolTipText     =   "Section Closed"
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Section 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   840
      TabIndex        =   8
      ToolTipText     =   "Section Closed"
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Section 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   840
      TabIndex        =   7
      ToolTipText     =   "Section Closed"
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Section 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   6
      ToolTipText     =   "Section Closed"
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Section 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   5
      ToolTipText     =   "Section Closed"
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Section 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   4
      ToolTipText     =   "Section Closed"
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "has performed a normal shut down of one or more Sections for security reasons.  The action is a result of user idle time."
      ForeColor       =   &H00800000&
      Height          =   795
      Left            =   840
      TabIndex        =   3
      Top             =   120
      Width           =   3015
   End
   Begin VB.Image img1 
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   0
      Picture         =   "EsiExit.frx":0C6E
      Top             =   0
      Width           =   540
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      ToolTipText     =   "Current Time"
      Top             =   2400
      Width           =   855
   End
End
Attribute VB_Name = "diaExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
   Unload Me
   
End Sub



Private Sub cmdHelp_Click()
   If cmdHelp Then
      OpenWebHelp "hs926"
      cmdHelp = False
      On Error Resume Next
      Text1.SetFocus
   End If
   
End Sub

Private Sub Form_Initialize()
   Dim wFlags As Long
   wFlags = 1
   SetWindowPos hWnd, hWnd_TopMost, 0, 0, 0, 0, wFlags
   
End Sub

Private Sub Form_Load()
   Dim iSect As Integer
   Dim sThisApp As String
   Label1.ForeColor = ES_BLUE
   lblShutDown.ForeColor = ES_RED
   lblShutDown.BorderStyle = 0
   lblTime = Format(Time, "h:mm AM/PM")
   Label1.Caption = "ES/" & Format(Now, "yyyy") & " ERP " & Label1.Caption
   sAppTitle = GetSetting("Esi2000", "System", "CloseSection", sAppTitle)
   'sAppTitle = "foobar"   'Testing
   'CloseRemainder
   If Trim(sAppTitle) <> "" Then
      SaveSetting "Esi2000", "System", "CloseSection", ""
      Section(1) = sAppTitle
      CloseCurrentApp
   Else
   End
End If

End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
End

End Sub


Private Sub Timer1_Timer()
   lblTime = Format(Time, "h:mm AM/PM")
   
End Sub


Private Sub Timer2_Timer()
   Static bByte As Byte
   bByte = bByte + 1
   If bUserAction = 1 Then bByte = 1
   bUserAction = 0
   'If bByte > 60 Then CloseRemainder
   
End Sub




Private Sub Timer3_Timer()
   sAppTitle = Trim(GetSetting("Esi2000", "System", "CloseSection", sAppTitle))
   SaveSetting "Esi2000", "System", "CloseSection", ""
   If sAppTitle <> "" Then CloseCurrentApp
   
End Sub
