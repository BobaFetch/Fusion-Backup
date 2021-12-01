VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form InvcINe04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Part Locations"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   3015
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "InvcINe04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   72
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optLoc 
      Alignment       =   1  'Right Justify
      Caption         =   "Show Only Parts Without A Location"
      Height          =   252
      Left            =   3000
      TabIndex        =   2
      ToolTipText     =   "Show Only Part Numbers That Currently Have Not Been Assigned An Inventory Location"
      Top             =   360
      Width           =   2892
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Index           =   10
      Left            =   7200
      TabIndex        =   16
      Tag             =   "3"
      Top             =   5160
      Width           =   615
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Index           =   9
      Left            =   7200
      TabIndex        =   15
      Tag             =   "3"
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton cmdNxt 
      Caption         =   " &Next >>"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6960
      TabIndex        =   18
      ToolTipText     =   "Page 2"
      Top             =   5640
      Width           =   875
   End
   Begin VB.CommandButton cmdLst 
      Caption         =   "<< &Last    "
      Enabled         =   0   'False
      Height          =   315
      Left            =   6060
      TabIndex        =   17
      ToolTipText     =   "Page 1"
      Top             =   5640
      Width           =   875
   End
   Begin VB.Frame z2 
      Height          =   40
      Index           =   1
      Left            =   120
      TabIndex        =   54
      Top             =   5520
      Width           =   7695
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Index           =   8
      Left            =   7200
      TabIndex        =   14
      Tag             =   "3"
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Index           =   7
      Left            =   7200
      TabIndex        =   13
      Tag             =   "3"
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Index           =   6
      Left            =   7200
      TabIndex        =   12
      Tag             =   "3"
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Index           =   5
      Left            =   7200
      TabIndex        =   11
      Tag             =   "3"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Index           =   4
      Left            =   7200
      TabIndex        =   10
      Tag             =   "3"
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Index           =   3
      Left            =   7200
      TabIndex        =   9
      Tag             =   "3"
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Index           =   2
      Left            =   7200
      TabIndex        =   8
      Tag             =   "3"
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Index           =   1
      Left            =   7200
      TabIndex        =   7
      Tag             =   "3"
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6960
      TabIndex        =   6
      ToolTipText     =   "Fill The Grid (300 Parts Maximum)"
      Top             =   1200
      Width           =   875
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6060
      TabIndex        =   5
      ToolTipText     =   "Cancel Work Not Updated And Return To Selection"
      Top             =   1200
      Width           =   875
   End
   Begin VB.Frame z2 
      Height          =   70
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   1080
      Width           =   7695
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "InvcINe04a.frx":07AE
      Height          =   315
      Left            =   4560
      Picture         =   "InvcINe04a.frx":0AF0
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   720
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.CheckBox optVew 
      Caption         =   "View"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   6600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      ToolTipText     =   "Leading Characters Or Blank For All (Selects Up To 300 Parts >= The Leading Characters)"
      Top             =   720
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "S&elect"
      Height          =   315
      Index           =   0
      Left            =   5040
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Selects A Maximum Of 300 Items"
      Top             =   720
      Width           =   855
   End
   Begin VB.ComboBox cmbLvl 
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "InvcINe04a.frx":0E32
      Left            =   1440
      List            =   "InvcINe04a.frx":0E34
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Select Part Type From List"
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6960
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   6120
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6150
      FormDesignWidth =   7905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lvl"
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
      Left            =   6120
      TabIndex        =   71
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lblLvl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   6120
      TabIndex        =   70
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label lblLvl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   6120
      TabIndex        =   69
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label lblLvl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   6120
      TabIndex        =   68
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lblLvl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   6120
      TabIndex        =   67
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label lblLvl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   6120
      TabIndex        =   66
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label lblLvl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   6120
      TabIndex        =   65
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label lblLvl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   6120
      TabIndex        =   64
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label lblLvl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   6120
      TabIndex        =   63
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label lblLvl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   6120
      TabIndex        =   62
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label lblLvl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   6120
      TabIndex        =   61
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   120
      TabIndex        =   60
      Top             =   5160
      Width           =   2895
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   3120
      TabIndex        =   59
      Top             =   5160
      Width           =   2895
   End
   Begin VB.Label lblLoc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   6480
      TabIndex        =   58
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   57
      Top             =   4800
      Width           =   2895
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   3120
      TabIndex        =   56
      Top             =   4800
      Width           =   2895
   End
   Begin VB.Label lblLoc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   6480
      TabIndex        =   55
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   53
      Top             =   4440
      Width           =   2895
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   3120
      TabIndex        =   52
      Top             =   4440
      Width           =   2895
   End
   Begin VB.Label lblLoc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   6480
      TabIndex        =   51
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   50
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   3120
      TabIndex        =   49
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Label lblLoc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   6480
      TabIndex        =   48
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   47
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   3120
      TabIndex        =   46
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Label lblLoc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   6480
      TabIndex        =   45
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   44
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   3120
      TabIndex        =   43
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label lblLoc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   6480
      TabIndex        =   42
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   41
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   3120
      TabIndex        =   40
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label lblLoc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   6480
      TabIndex        =   39
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   38
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   3120
      TabIndex        =   37
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label lblLoc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   6480
      TabIndex        =   36
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   35
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   3120
      TabIndex        =   34
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label lblLoc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   6480
      TabIndex        =   33
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   32
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   3120
      TabIndex        =   31
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label lblLoc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   6480
      TabIndex        =   30
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current     New         "
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
      Left            =   6480
      TabIndex        =   29
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number                                                "
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
      Left            =   120
      TabIndex        =   28
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description                                                "
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
      Index           =   5
      Left            =   3120
      TabIndex        =   27
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Found"
      Height          =   255
      Index           =   8
      Left            =   6600
      TabIndex        =   22
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblNum 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7320
      TabIndex        =   21
      Top             =   720
      Width           =   510
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   20
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "InvcINe04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'12/2/03 New
'12/5/05 Clarified ToolTipText >= the entry. Fixed cmdNxt.
'12/8/   Added optLoc.  Show only Parts with no loc
Option Explicit
Dim bOnLoad As Byte

Dim iTotalParts As Integer
Dim iCurrIdx As Integer

Dim sPartLocs(301, 6) As String
'0 = PARTREF
'1 = PARTNUM
'2 = PADESC
'3 = PALOCATION
'4 = New Location (Same as PALOCATION unless changed)
'5 = Level
Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdEnd_Click()
   Dim bResponse As Byte
   bResponse = MsgBox("Are You Sure That You Want To Cancel?", _
               ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      ManageBoxes
      cmbLvl.Enabled = True
      txtPrt.Enabled = True
      cmdFnd.Enabled = True
      optLoc.Enabled = True
      cmdGo(0).Enabled = True
      cmdUpd.Enabled = False
      cmdNxt.Enabled = False
      cmdLst.Enabled = False
      cmdEnd.Enabled = False
   End If
   
End Sub

Private Sub cmdFnd_Click()
   ViewParts.lblControl = "TXTPRT"
   ViewParts.txtPrt = txtPrt
   optVew.Value = vbChecked
   ViewParts.Show
   
End Sub


Private Sub cmdGo_Click(Index As Integer)
   SelectParts
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 1602
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdLst_Click()
   iCurrIdx = iCurrIdx - 10
   If iCurrIdx < 0 Then iCurrIdx = 0
   If iCurrIdx = 0 Then cmdLst.Enabled = False
   If iCurrIdx > 0 Then cmdNxt.Enabled = True
   GetNextGroup
   
End Sub

Private Sub cmdNxt_Click()
   iCurrIdx = iCurrIdx + 10
   'If iCurrIdx > iTotalParts / 10 Then iCurrIdx = iCurrIdx - 11
   cmdLst.Enabled = True
   If iCurrIdx > (iTotalParts - 10) Then cmdNxt.Enabled = False
   GetNextGroup
   
End Sub


Private Sub cmdUpd_Click()
   Dim bResponse As Byte
   bResponse = MsgBox("Update To The Current Locations?", _
               ES_YESQUESTION, Caption)
   If bResponse = vbYes Then UpdateParts Else CancelTrans
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then bOnLoad = 0
   MouseCursor 0
   FillCombo
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   cmbPrt.AddItem "1 - ALL"
   cmbLvl.AddItem "ALL"
   cmbLvl.AddItem "1 - Top"
   cmbLvl.AddItem "2 - Mid"
   cmbLvl.AddItem "3 - Base"
   cmbLvl.AddItem "4 - Raw"
   cmbLvl.AddItem "5 - Expendables"
   cmbLvl.AddItem "8 - Project"
   cmbLvl = cmbLvl.List(0)
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set InvcINe04a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtPrt = "ALL"
   lblNum = 0
   cmdEnd.ToolTipText = "Cancel Work Not Updated And Return To Selection"
   cmdLst.ToolTipText = "Last Page (Page Up)"
   cmdNxt.ToolTipText = "Next Page (Page Down)"
   ManageBoxes
   
End Sub

Private Sub FillCombo()
   sSql = "Qry_FillSortedParts"
   LoadComboBox cmbPrt
 '  FillVendors
   If cmbPrt.ListCount > 0 Then
      cmbPrt = cmbPrt.List(0)
      'bGoodPart = GetAliasedPart()
 '     GetAlias True
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "FillCombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Sub SelectParts()
   Dim RdoSel As ADODB.Recordset
   Dim bLen As Byte
   Dim iRows As Integer
   
   Dim sParts As String
   Erase sPartLocs
   iTotalParts = 0
   lblNum = 0
   ManageBoxes
   
   On Error GoTo DiaErr1
   If cmbPrt <> "ALL" Then sParts = Compress(cmbPrt)
   bLen = Len(sParts)
   If bLen = 0 Then bLen = 1
   If cmbLvl = "ALL" Then
      sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PALOCATION " _
             & "FROM PartTable WHERE (LEFT(PARTREF," & str$(bLen) & ")> ='" & sParts _
             & "' AND PALEVEL<>6 AND PALEVEL<>7 AND PATOOL=0"
   Else
      sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PALOCATION " _
             & "FROM PartTable WHERE (LEFT(PARTREF," & str$(bLen) & ")> ='" & sParts _
             & "' AND PALEVEL=" & Val(Left(cmbLvl, 1)) & " AND PATOOL=0"
   End If
   If optLoc.Value = vbChecked Then sSql = sSql & " AND PALOCATION=''"
   sSql = sSql & ") ORDER BY PARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSel, ES_FORWARD)
   If bSqlRows Then
      With RdoSel
         Do Until .EOF
            iRows = iRows + 1
            If iRows > 300 Then
               iRows = iRows - 1
               Exit Do
            End If
            sPartLocs(iRows, 0) = "" & Trim(!PartRef)
            sPartLocs(iRows, 1) = "" & Trim(!PartNum)
            sPartLocs(iRows, 2) = "" & Trim(!PADESC)
            sPartLocs(iRows, 3) = "" & Trim(!PALOCATION)
            sPartLocs(iRows, 4) = "" & Trim(!PALOCATION)
            sPartLocs(iRows, 5) = "" & Trim(str(!PALEVEL))
            If iRows < 11 Then
               lblPrt(iRows) = "" & Trim(!PartNum)
               lblDsc(iRows) = "" & Trim(!PADESC)
               lblLoc(iRows) = "" & Trim(!PALOCATION)
               lblLvl(iRows) = "" & Trim(str(!PALEVEL))
               txtLoc(iRows) = "" & Trim(!PALOCATION)
               txtLoc(iRows).Enabled = True
               txtLoc(iRows).BackColor = Es_TextBackColor
            End If
            .MoveNext
         Loop
         ClearResultSet RdoSel
      End With
      On Error Resume Next
      iCurrIdx = 0
      iTotalParts = iRows
      txtLoc(1).SetFocus
      cmdUpd.Enabled = True
      If iTotalParts > 10 Then cmdNxt.Enabled = True
      cmdEnd.Enabled = True
      cmbLvl.Enabled = False
      txtPrt.Enabled = False
      cmdFnd.Enabled = False
      optLoc.Enabled = False
      cmdGo(0).Enabled = False
      lblNum = iTotalParts
      If txtLoc(1).Enabled Then txtLoc(1).SetFocus
   Else
      MsgBox "No Matching Parts Found.", _
         vbInformation, Caption
   End If
   Set RdoSel = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "selectparts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtLoc_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtLoc_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
End Sub

Private Sub txtLoc_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyCase KeyAscii
   
End Sub


Private Sub txtLoc_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtLoc_LostFocus(Index As Integer)
   txtLoc(Index) = Compress(txtLoc(Index))
   txtLoc(Index) = CheckLen(txtLoc(Index), 4)
   sPartLocs(Index + iCurrIdx, 4) = txtLoc(Index)
   
End Sub


Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "TXTPRT"
      ViewParts.txtPrt = txtPrt
      optVew.Value = vbChecked
      ViewParts.Show
   End If
   
End Sub

Private Sub txtPrt_LostFocus()
   txtPrt = CheckLen(txtPrt, 30)
   
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If (Not ValidPartNumber(cmbPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If
   
End Sub


Private Sub ManageBoxes()
   On Error Resume Next
   Dim iRow As Integer
   For iRow = 1 To 10
      lblPrt(iRow) = ""
      lblDsc(iRow) = ""
      lblLvl(iRow) = ""
      lblLoc(iRow) = ""
      txtLoc(iRow) = ""
      txtLoc(iRow).BackColor = Es_FormBackColor
      txtLoc(iRow).Enabled = False
   Next
   
End Sub

Private Sub GetNextGroup()
   Dim iRow As Integer
   Dim iEnd As Integer
   ManageBoxes
   On Error Resume Next
   For iRow = 1 To 10
      If iRow + iCurrIdx > iTotalParts Then Exit For
      lblPrt(iRow) = sPartLocs(iRow + iCurrIdx, 1)
      lblDsc(iRow) = sPartLocs(iRow + iCurrIdx, 2)
      lblLvl(iRow) = sPartLocs(iRow + iCurrIdx, 5)
      lblLoc(iRow) = sPartLocs(iRow + iCurrIdx, 3)
      txtLoc(iRow) = sPartLocs(iRow + iCurrIdx, 4)
      txtLoc(iRow).Enabled = True
      txtLoc(iRow).BackColor = Es_TextBackColor
   Next
   If txtLoc(1).Enabled Then txtLoc(1).SetFocus
   If iRow + iCurrIdx < iTotalParts Then cmdNxt.Enabled = True
   
End Sub

Private Sub UpdateParts()
   Dim iRows As Integer
   
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   For iRows = 1 To iTotalParts
      If sPartLocs(iRows, 3) <> sPartLocs(iRows, 4) Then
         sSql = "UPDATE PartTable SET PALOCATION='" _
                & sPartLocs(iRows, 4) & "' WHERE PARTREF='" _
                & sPartLocs(iRows, 0) & "' "
         clsADOCon.ExecuteSQL sSql
      End If
   Next
   If clsADOCon.ADOErrNum = 0 Then
      SysMsg "Locations Where Updated.", True
   Else
      MsgBox "Couldn't Update Selections.", _
         vbInformation, Caption
   End If
   ManageBoxes
   cmbLvl.Enabled = True
   txtPrt.Enabled = True
   cmdFnd.Enabled = True
   optLoc.Enabled = True
   cmdGo(0).Enabled = True
   cmdUpd.Enabled = False
   cmdNxt.Enabled = False
   cmdLst.Enabled = False
   cmdEnd.Enabled = False
   
End Sub
