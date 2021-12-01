VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form SysAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Key Software"
   ClientHeight    =   3525
   ClientLeft      =   6000
   ClientTop       =   1005
   ClientWidth     =   3435
   ClipControls    =   0   'False
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
   ForeColor       =   &H00C0C0C0&
   HelpContextID   =   5
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3525
   ScaleWidth      =   3435
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDmy 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1215
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3000
      Top             =   3240
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3525
      FormDesignWidth =   3435
   End
   Begin VB.Label lblOs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " "
      ForeColor       =   &H00800000&
      Height          =   252
      Left            =   1700
      TabIndex        =   11
      Top             =   1460
      Width           =   1600
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Operating System"
      ForeColor       =   &H00800000&
      Height          =   252
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   1460
      Width           =   1692
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Index           =   1
      Left            =   240
      Picture         =   "SysAbout.frx":0000
      Top             =   0
      Width           =   2985
   End
   Begin VB.Label lblVer 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   2160
      Width           =   1605
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Section Version"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "www.keysoftwarellc.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   4
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "ESI On The Web"
      Top             =   2520
      Width           =   3012
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Index           =   4
      Left            =   240
      Picture         =   "SysAbout.frx":37E2
      Top             =   4440
      Width           =   3630
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Index           =   3
      Left            =   240
      Picture         =   "SysAbout.frx":152A0
      Top             =   3600
      Width           =   3630
   End
   Begin VB.Label lblApp 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   1940
      Width           =   1845
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
      ForeColor       =   &H00800000&
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1940
      Width           =   1692
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Index           =   0
      Left            =   240
      Picture         =   "SysAbout.frx":26D5E
      Top             =   5760
      Width           =   3630
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MB"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   3
      Top             =   1710
      Width           =   400
   End
   Begin VB.Label lblMem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " "
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1960
      TabIndex        =   2
      Top             =   1710
      Width           =   1035
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Physical Memory"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1710
      Width           =   1815
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C0C0&
      X1              =   0
      X2              =   3255
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "SysAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'6/30/04 Corrected lblApp (do not attempt to show the Database).
'        See Form_Load.
'12/5/05 Added GetWindowsVersion. Added link to website

Option Explicit




Private Sub cmdCan_Click()
   Form_Deactivate
   
End Sub


Private Sub Form_Deactivate()
   Unload Me
   
End Sub

Private Sub Form_Load()
   Dim bByte As Byte
   Dim FreeSpc As Long
   Dim sBuild As String
   Dim sYear As String
   Dim sAppCap As String
   Dim SysMemory As MEMORYSTATUS
   top = 10
   If iBarOnTop Then
      Left = MdiSect.Width - (Width + 600)
   Else
      Left = MdiSect.Width - (Width + MdiSect.SideBar.Width + 430)
   End If
   Image1(1).Left = (Width - Image1(1).Width) / 2
   cmdCan.Left = (Width - cmdCan.Width) / 2
   z1(4).Left = (Width - z1(4).Width) / 2
   GetWindowsVersion
   GlobalMemoryStatus SysMemory
   FreeSpc = ((SysMemory.dwTotalPhys) / 1024 * (1 / 1024)) + 0.4
   lblMem = Format(FreeSpc&, "#,###,##0")
   sAppCap = MdiSect.Caption
   bByte = InStr(sAppCap, "-")
   sAppCap = Left$(sAppCap, bByte - 2)
   lblApp = sAppCap
   lblVer = App.Major & "." & App.Minor & ".0" & "." & App.Revision
   sYear = Format$(Now, "yyyy")
   '\bitmaps about2007, etc
   Select Case sYear
      Case "2007"
         Image1(1).Picture = Image1(0).Picture
      Case "2008"
         Image1(1).Picture = Image1(3).Picture
      Case "2006"
         Image1(1).Picture = Image1(4).Picture
      Case Else
         Image1(1).Picture = Image1(1).Picture
   End Select
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set SysAbout = Nothing
   
End Sub


Private Sub Image1_Click(Index As Integer)
   '\bitmaps about2007, etc.
End Sub

Private Sub z1_Click(Index As Integer)
   If Index = 4 Then
      z1(Index).ForeColor = RGB(255, 92, 92)
      OpenWebPage "http://" & Trim(z1(4))
   End If
   
End Sub



Private Sub GetWindowsVersion()
   Dim lRet As Long
   Dim typOS As OSVERSIONINFO
   Dim sServicePack As String
   typOS.dwOSVersionInfoSize = Len(typOS)
   lRet = GetVersionEx(typOS)
   If typOS.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
      lblOs = "Windows 95/98"
   Else
      sServicePack = Mid$(typOS.szCSDVersion, 14, 1)
      If typOS.dwMinorVersion = 0 Then
         lblOs = "Windows 2000"
         If Val(sServicePack) > 0 Then _
                lblOs = lblOs & " SP" & sServicePack
      Else
         lblOs = "Windows XP"
         If Val(sServicePack) > 0 Then _
                lblOs = lblOs & " SP" & sServicePack
      End If
   End If
   
End Sub
