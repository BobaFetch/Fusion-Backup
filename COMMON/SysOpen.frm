VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SysOpen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fusion  ERP"
   ClientHeight    =   990
   ClientLeft      =   1740
   ClientTop       =   4800
   ClientWidth     =   3540
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "SysOpen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   990
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   252
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   2772
      _ExtentX        =   4895
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   8125
      Left            =   3000
      Top             =   840
   End
   Begin VB.Label pnl 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   2772
   End
   Begin VB.Image img1 
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   0
      Picture         =   "SysOpen.frx":030A
      Top             =   0
      Width           =   540
   End
End
Attribute VB_Name = "SysOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'6/21/06 Changed to SysOpen
Option Explicit
'The only function of this form is to return proper focus to
'ActiveBar after loading.

Private Sub Form_Deactivate()
   Timer1.Enabled = False
   Unload Me
   
End Sub

'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions

Private Sub Form_Load()
   MouseCursor 13
   BackColor = ES_SystemBackcolor
   pnl.ForeColor = ES_BLUE
   Caption = "Fusion ERP"
   If Screen.Width < 10000 Then
      Move (MDISect.Width) - (MDISect.Width - 2000), MDISect.Height - 1800
   Else
      Move (MDISect.Width) - (MDISect.Width - 2500), MDISect.Height - 1900
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   MouseCursor 0
   Set SysOpen = Nothing
   
End Sub

Private Sub Timer1_Timer()
   'Set for 2min in case it won't go away
   Static b As Byte
   If MDISect.WindowState = 1 Then Hide
   b = b + 1
   If b = 8 Then Unload Me
   
End Sub
