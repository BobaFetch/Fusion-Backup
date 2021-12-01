VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form AdmnUsecr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Advanced Security Installation"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Fra 
      Height          =   5292
      Index           =   2
      Left            =   100
      TabIndex        =   21
      Top             =   0
      Width           =   4932
      Begin VB.CommandButton cmdPrv 
         Caption         =   "< Back"
         Height          =   375
         Index           =   2
         Left            =   2140
         TabIndex        =   36
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton cmdNxt 
         Caption         =   "Finish >"
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   3000
         TabIndex        =   35
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton cmdCan 
         Caption         =   "Cancel"
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox txtKey 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   31
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtKey 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   30
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtKey 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   2880
         MaxLength       =   4
         TabIndex        =   29
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   1800
         TabIndex        =   33
         Top             =   1320
         Width           =   348
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   2640
         TabIndex        =   32
         Top             =   1320
         Width           =   348
      End
      Begin VB.Label z1 
         Caption         =   "You should have been given a Key. Enter that number here:"
         Height          =   612
         Index           =   16
         Left            =   360
         TabIndex        =   28
         Top             =   600
         Width           =   3012
      End
   End
   Begin VB.Frame Fra 
      Height          =   5292
      Index           =   1
      Left            =   100
      TabIndex        =   20
      Top             =   0
      Width           =   4932
      Begin VB.CommandButton cmdCan 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Index           =   2
         Left            =   3960
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton cmdNxt 
         Caption         =   "Next >"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   3000
         TabIndex        =   26
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton cmdPrv 
         Caption         =   "< Back"
         Height          =   375
         Index           =   1
         Left            =   2160
         TabIndex        =   25
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox txtLic 
         Height          =   3375
         Left            =   480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   360
         Width           =   4095
      End
      Begin VB.OptionButton optLic 
         Caption         =   "Agree"
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   23
         Top             =   4200
         Width           =   1215
      End
      Begin VB.OptionButton optLic 
         Caption         =   "Don't Agree"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   22
         Top             =   4200
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame Fra 
      Height          =   5292
      Index           =   0
      Left            =   100
      TabIndex        =   0
      Top             =   0
      Width           =   4905
      Begin VB.CommandButton cmdPrv 
         Caption         =   "<<< Back"
         Height          =   375
         Index           =   0
         Left            =   3960
         TabIndex        =   5
         Top             =   4200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdNxt 
         Caption         =   "Next >"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   3000
         TabIndex        =   4
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton cmdCan 
         Caption         =   "Cancel"
         Height          =   375
         Index           =   0
         Left            =   3960
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   4680
         Width           =   855
      End
      Begin VB.CheckBox optRead 
         Caption         =   "I Have Read All Of This And The Help "
         ForeColor       =   &H00800000&
         Height          =   400
         Left            =   480
         TabIndex        =   2
         Top             =   4656
         Width           =   2400
      End
      Begin VB.CommandButton cmdHlp 
         Appearance      =   0  'Flat
         Height          =   250
         Left            =   120
         Picture         =   "AdmnUsecr.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Subject Help"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   250
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         Caption         =   "This procedure is designed to upgrade the "
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   2
         Left            =   480
         TabIndex        =   19
         Top             =   600
         Width           =   3492
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         Caption         =   "early systems  (before 2.6.0) to the latest "
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   3
         Left            =   480
         TabIndex        =   18
         Top             =   840
         Width           =   3492
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         Caption         =   "security features.  It will install security on"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   4
         Left            =   480
         TabIndex        =   17
         Top             =   1080
         Width           =   3600
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         Caption         =   "later versions. Please read the [?] Help now."
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   5
         Left            =   480
         TabIndex        =   16
         Top             =   1320
         Width           =   3492
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         Caption         =   "You will need the following before you begin:"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   6
         Left            =   480
         TabIndex        =   15
         Top             =   1680
         Width           =   3600
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         Caption         =   "User ID - 30 char max, not case sensitive"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   7
         Left            =   480
         TabIndex        =   14
         Top             =   1920
         Width           =   3600
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         Caption         =   "User Name - 40 char max (the persons full name)"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   8
         Left            =   480
         TabIndex        =   13
         Top             =   2160
         Width           =   3600
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         Caption         =   "Password - 15 char max and case sensitive"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   9
         Left            =   480
         TabIndex        =   12
         Top             =   2400
         Width           =   3600
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         Caption         =   "Nickname - 20 char max"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   10
         Left            =   480
         TabIndex        =   11
         Top             =   2640
         Width           =   3600
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         Caption         =   "Initials - Uppercase 3 chars"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   11
         Left            =   480
         TabIndex        =   10
         Top             =   2880
         Width           =   3600
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         Caption         =   "Class - Administrator or Users"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   12
         Left            =   480
         TabIndex        =   9
         Top             =   3120
         Width           =   3600
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         Caption         =   "Note: Administrator auto all features and Users auto none"
         ForeColor       =   &H80000008&
         Height          =   396
         Index           =   13
         Left            =   720
         TabIndex        =   8
         Top             =   3360
         Width           =   3612
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         Caption         =   "Managed by Section, Group and Tab (Edit etc)"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   14
         Left            =   480
         TabIndex        =   7
         Top             =   3816
         Width           =   3600
      End
      Begin VB.Label z1 
         Appearance      =   0  'Flat
         Caption         =   "Once Installed, current user must be installed to use ES/2000ERP"
         ForeColor       =   &H00000080&
         Height          =   492
         Index           =   15
         Left            =   480
         TabIndex        =   6
         Top             =   4080
         Width           =   3492
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5040
      Top             =   5040
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5385
      FormDesignWidth =   5145
   End
End
Attribute VB_Name = "AdmnUsecr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'Key 550 041 330
Dim bOnLoad As Byte
Dim bHelpShown As Byte
Dim bIsntSet As Boolean

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub FormatTextBox()
   Dim sMsg As String
   sMsg = sSysCaption & vbCrLf
   sMsg = sMsg & "This product and the associated programs, reports and " _
          & "all portions of the product not furnished by first or third " _
          & "parties, are the sole property of Enterprise Systems, Inc. " _
          & "Illegal use of this product will be prosecuted to the full " _
          & "extent of national and international trade laws."
   
   sMsg = sMsg & vbCrLf & vbCrLf _
          & "In addition, this product is covered by individual " _
          & "corporate license agreements clearly stating the features " _
          & "that may accessed by those license agreements. Currently, " _
          & "those agreements may not be enforced by software, but by " _
          & "good faith with customers.  Those license agreements can " _
          & "and will be enforced."
   txtLic = sMsg
   
End Sub

Private Sub cmdCan_Click(Index As Integer)
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   Dim l&
   If cmdHlp Then
      MouseCursor 13
      ' l& = WinHelp(hwnd, sReportPath & "security.hlp", HELP_KEY, "Advanced Security Setup")
      MouseCursor 0
      bHelpShown = 1
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdNxt_Click(Index As Integer)
   Dim bResponse As Byte
   Dim sMsg As String
   bHelpShown = 1
   If bHelpShown = 0 Then
      sMsg = "You Have Not Read The Subject Help. " & vbCr _
             & "Please Click On The [?], Read And " & vbCr _
             & "Possibly Print The Help Content."
      MsgBox sMsg, vbExclamation, Caption
      Exit Sub
   End If
   If Index = 0 Then
      sMsg = "I have really read all of that and attest " & vbCr _
             & "that I am in no way confused and am" & vbCr _
             & "really to move on to the next step."
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbNo Then Exit Sub
   Else
      If Index = 2 Then
         sMsg = "I understand that the person that I am " & vbCr _
                & "about to establish as a user is Me.  I " & vbCr _
                & "will grant myself full Admin Permissions."
         bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
         If bResponse = vbNo Then Exit Sub
      End If
   End If
   On Error Resume Next
   Select Case Index
      Case 0
         Fra(0).Visible = False
         Fra(1).Visible = True
         Fra(2).Visible = False
      Case 1
         Fra(0).Visible = False
         Fra(1).Visible = False
         Fra(2).Visible = True
         txtKey(1).SetFocus
      Case Else
         CheckCdKey
   End Select
   
End Sub


Private Sub cmdPrv_Click(Index As Integer)
   Select Case Index
      Case 1
         Fra(0).Visible = True
         Fra(1).Visible = False
         Fra(2).Visible = False
      Case 2
         Fra(0).Visible = False
         Fra(2).Visible = False
         Fra(1).Visible = True
      Case Else
         
   End Select
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      bIsntSet = True
      FormatTextBox
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   Width = 4905
   FormatControls
   bOnLoad = 1
   Show
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim bResponse As Byte
   Dim sMsg As String
   If bIsntSet Then
      sMsg = "Security Is Not Correctly Installed." & vbCr _
             & "Do You Wish To Quit Now?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbNo Then
         Cancel = True
      Else
         MsgBox "Installation Aborted And May Be Installed Later.", _
            vbInformation, Caption
      End If
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set AdmnUsecr = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   For b = 0 To 2
      Fra(b).BackColor = Es_FormBackColor
   Next
   Fra(b).BackColor = Es_FormBackColor
   optRead.ForeColor = ES_BLUE
   
End Sub

Private Sub optLic_Click(Index As Integer)
   If Index = (0) Then cmdNxt(1).Enabled = True Else cmdNxt(1).Enabled = False
   
End Sub


Private Sub optRead_Click()
   If optRead.value = vbChecked Then
      cmdNxt(0).Enabled = True
   Else
      cmdNxt(0).Enabled = False
   End If
   
End Sub

Private Sub txtKey_Change(Index As Integer)
   If Index < 3 Then
      If Len(Trim(txtKey(Index))) = 3 Then txtKey(Index + 1).SetFocus
   End If
   
End Sub


Private Sub txtKey_LostFocus(Index As Integer)
   If Len(Trim(txtKey(1))) = 3 And Len(Trim(txtKey(2))) = 3 And _
          Len(Trim(txtKey(3))) = 3 Then cmdNxt(2).Enabled = True
   
   
End Sub



'Key is 550-041-330 without dashes

Private Sub CheckCdKey()
   Dim sCdKey As String
   sCdKey = Trim(txtKey(1)) & Trim(txtKey(2)) & Trim(txtKey(3))
   If sCdKey = "550041330" Then
      bIsntSet = False
      AdmnUnewu.Show
      Unload Me
   Else
      MsgBox "Key Number Entered In Error." & vbCr _
         & "Please Re-enter The Correct Number.", _
         vbExclamation
   End If
   
End Sub
