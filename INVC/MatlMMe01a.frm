VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form MatlMMe01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Assign ABC Classes"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   1603
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   7935
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
      Picture         =   "MatlMMe01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   72
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbAbc 
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   10
      ItemData        =   "MatlMMe01a.frx":07AE
      Left            =   7245
      List            =   "MatlMMe01a.frx":07BE
      TabIndex        =   15
      Tag             =   "3"
      ToolTipText     =   "Select From List (A,B,C)  Or Leave Blank"
      Top             =   5160
      Width           =   615
   End
   Begin VB.ComboBox cmbAbc 
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   9
      ItemData        =   "MatlMMe01a.frx":07D0
      Left            =   7245
      List            =   "MatlMMe01a.frx":07E0
      TabIndex        =   14
      Tag             =   "3"
      ToolTipText     =   "Select From List (A,B,C)  Or Leave Blank"
      Top             =   4800
      Width           =   615
   End
   Begin VB.ComboBox cmbAbc 
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   8
      ItemData        =   "MatlMMe01a.frx":07F2
      Left            =   7245
      List            =   "MatlMMe01a.frx":0802
      TabIndex        =   13
      Tag             =   "3"
      ToolTipText     =   "Select From List (A,B,C)  Or Leave Blank"
      Top             =   4440
      Width           =   615
   End
   Begin VB.ComboBox cmbAbc 
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   7
      ItemData        =   "MatlMMe01a.frx":0814
      Left            =   7245
      List            =   "MatlMMe01a.frx":0824
      TabIndex        =   12
      Tag             =   "3"
      ToolTipText     =   "Select From List (A,B,C)  Or Leave Blank"
      Top             =   4080
      Width           =   615
   End
   Begin VB.ComboBox cmbAbc 
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   6
      ItemData        =   "MatlMMe01a.frx":0836
      Left            =   7245
      List            =   "MatlMMe01a.frx":0846
      TabIndex        =   11
      Tag             =   "3"
      ToolTipText     =   "Select From List (A,B,C)  Or Leave Blank"
      Top             =   3720
      Width           =   615
   End
   Begin VB.ComboBox cmbAbc 
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   5
      ItemData        =   "MatlMMe01a.frx":0858
      Left            =   7245
      List            =   "MatlMMe01a.frx":0868
      TabIndex        =   10
      Tag             =   "3"
      ToolTipText     =   "Select From List (A,B,C)  Or Leave Blank"
      Top             =   3360
      Width           =   615
   End
   Begin VB.ComboBox cmbAbc 
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   4
      ItemData        =   "MatlMMe01a.frx":087A
      Left            =   7245
      List            =   "MatlMMe01a.frx":088A
      TabIndex        =   9
      Tag             =   "3"
      ToolTipText     =   "Select From List (A,B,C)  Or Leave Blank"
      Top             =   3000
      Width           =   615
   End
   Begin VB.ComboBox cmbAbc 
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   3
      ItemData        =   "MatlMMe01a.frx":089C
      Left            =   7245
      List            =   "MatlMMe01a.frx":08AC
      TabIndex        =   8
      Tag             =   "3"
      ToolTipText     =   "Select From List (A,B,C)  Or Leave Blank"
      Top             =   2640
      Width           =   615
   End
   Begin VB.ComboBox cmbAbc 
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   2
      ItemData        =   "MatlMMe01a.frx":08BE
      Left            =   7245
      List            =   "MatlMMe01a.frx":08CE
      TabIndex        =   7
      Tag             =   "3"
      ToolTipText     =   "Select From List (A,B,C)  Or Leave Blank"
      Top             =   2280
      Width           =   615
   End
   Begin VB.ComboBox cmbAbc 
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   1
      ItemData        =   "MatlMMe01a.frx":08E0
      Left            =   7245
      List            =   "MatlMMe01a.frx":08F0
      TabIndex        =   6
      Tag             =   "3"
      ToolTipText     =   "Select From List (A,B,C)  Or Leave Blank"
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton cmdNxt 
      Caption         =   " &Next >>"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6960
      TabIndex        =   17
      Top             =   5640
      Width           =   875
   End
   Begin VB.CommandButton cmdLst 
      Caption         =   "<< &Last    "
      Enabled         =   0   'False
      Height          =   315
      Left            =   6060
      TabIndex        =   16
      Top             =   5640
      Width           =   875
   End
   Begin VB.Frame z2 
      Height          =   40
      Index           =   1
      Left            =   120
      TabIndex        =   53
      Top             =   5520
      Width           =   7695
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6960
      TabIndex        =   5
      ToolTipText     =   "Update Classes And Apply Changes"
      Top             =   1200
      Width           =   875
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6060
      TabIndex        =   4
      ToolTipText     =   "Cancel Current Changes"
      Top             =   1200
      Width           =   875
   End
   Begin VB.Frame z2 
      Height          =   70
      Index           =   0
      Left            =   120
      TabIndex        =   25
      Top             =   1080
      Width           =   7695
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "MatlMMe01a.frx":0902
      Height          =   315
      Left            =   4560
      Picture         =   "MatlMMe01a.frx":0C44
      Style           =   1  'Graphical
      TabIndex        =   24
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
      TabIndex        =   23
      Top             =   6600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      ToolTipText     =   "Leading Characters Or Blank For All"
      Top             =   720
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "S&elect"
      Height          =   315
      Index           =   0
      Left            =   5040
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Selects A Maximum Of 300 Items"
      Top             =   720
      Width           =   855
   End
   Begin VB.ComboBox cmbLvl 
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "MatlMMe01a.frx":0F86
      Left            =   1440
      List            =   "MatlMMe01a.frx":0F88
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Select From List"
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6960
      TabIndex        =   18
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
      FormDesignWidth =   7935
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status"
      Height          =   255
      Left            =   120
      TabIndex        =   71
      ToolTipText     =   "You Must Run And Setup Inventory ABC "
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Label lblCost 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   5760
      TabIndex        =   70
      Top             =   5160
      Width           =   1040
   End
   Begin VB.Label lblCost 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   5760
      TabIndex        =   69
      Top             =   4800
      Width           =   1040
   End
   Begin VB.Label lblCost 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   5760
      TabIndex        =   68
      Top             =   4440
      Width           =   1040
   End
   Begin VB.Label lblCost 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   5760
      TabIndex        =   67
      Top             =   4080
      Width           =   1040
   End
   Begin VB.Label lblCost 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   5760
      TabIndex        =   66
      Top             =   3720
      Width           =   1040
   End
   Begin VB.Label lblCost 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   5760
      TabIndex        =   65
      Top             =   3360
      Width           =   1040
   End
   Begin VB.Label lblCost 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   5760
      TabIndex        =   64
      Top             =   3000
      Width           =   1040
   End
   Begin VB.Label lblCost 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   5760
      TabIndex        =   63
      Top             =   2640
      Width           =   1040
   End
   Begin VB.Label lblCost 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   5760
      TabIndex        =   62
      Top             =   2280
      Width           =   1040
   End
   Begin VB.Label lblCost 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   5760
      TabIndex        =   61
      Top             =   1920
      Width           =   1040
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Std Cost          "
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
      Left            =   5760
      TabIndex        =   60
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   120
      TabIndex        =   59
      Top             =   5160
      Width           =   2800
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   3000
      TabIndex        =   58
      Top             =   5160
      Width           =   2715
   End
   Begin VB.Label lblLoc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   6870
      TabIndex        =   57
      Top             =   5160
      Width           =   315
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   56
      Top             =   4800
      Width           =   2800
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   3000
      TabIndex        =   55
      Top             =   4800
      Width           =   2715
   End
   Begin VB.Label lblLoc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   6870
      TabIndex        =   54
      Top             =   4800
      Width           =   315
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   52
      Top             =   4440
      Width           =   2800
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   3000
      TabIndex        =   51
      Top             =   4440
      Width           =   2715
   End
   Begin VB.Label lblLoc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   6870
      TabIndex        =   50
      Top             =   4440
      Width           =   315
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   49
      Top             =   4080
      Width           =   2800
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   3000
      TabIndex        =   48
      Top             =   4080
      Width           =   2715
   End
   Begin VB.Label lblLoc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   6870
      TabIndex        =   47
      Top             =   4080
      Width           =   315
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   46
      Top             =   3720
      Width           =   2800
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   3000
      TabIndex        =   45
      Top             =   3720
      Width           =   2715
   End
   Begin VB.Label lblLoc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   6870
      TabIndex        =   44
      Top             =   3720
      Width           =   315
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   43
      Top             =   3360
      Width           =   2800
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   3000
      TabIndex        =   42
      Top             =   3360
      Width           =   2715
   End
   Begin VB.Label lblLoc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   6870
      TabIndex        =   41
      Top             =   3360
      Width           =   315
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   40
      Top             =   3000
      Width           =   2800
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   3000
      TabIndex        =   39
      Top             =   3000
      Width           =   2715
   End
   Begin VB.Label lblLoc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   6870
      TabIndex        =   38
      Top             =   3000
      Width           =   315
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   37
      Top             =   2640
      Width           =   2800
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   3000
      TabIndex        =   36
      Top             =   2640
      Width           =   2715
   End
   Begin VB.Label lblLoc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   6870
      TabIndex        =   35
      Top             =   2640
      Width           =   315
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   34
      Top             =   2280
      Width           =   2800
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   3000
      TabIndex        =   33
      Top             =   2280
      Width           =   2715
   End
   Begin VB.Label lblLoc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   6870
      TabIndex        =   32
      Top             =   2280
      Width           =   315
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   31
      Top             =   1920
      Width           =   2800
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   30
      Top             =   1920
      Width           =   2715
   End
   Begin VB.Label lblLoc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   6870
      TabIndex        =   29
      Top             =   1920
      Width           =   315
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Curr   New     "
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
      Left            =   6840
      TabIndex        =   28
      Top             =   1680
      Width           =   1095
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
      TabIndex        =   27
      Top             =   1680
      Width           =   2775
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
      Left            =   3000
      TabIndex        =   26
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Found"
      Height          =   255
      Index           =   8
      Left            =   6600
      TabIndex        =   21
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblNum 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7320
      TabIndex        =   20
      Top             =   720
      Width           =   510
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   19
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "MatlMMe01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'New 12/3/03
'New ABC's 12/12/03
'9/1/04 omit tools
'12/21/04 Corrected query Select Parts (PASTDCOST)
Option Explicit
Dim bOnLoad As Byte

Dim iTotalParts As Integer
Dim iCurrIdx As Integer

Dim sPartABC(301, 7) As String
'0 = PARTREF
'1 = PARTNUM
'2 = PADESC
'3 = PAABC
'4 = New ABC (Same as PAABC unless changed)
'5 = Standard Cost
'6 = ToolTipText
Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

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
      lblStatus.ForeColor = ES_RED
      lblStatus.ToolTipText = "You Should Run And Properly Install Inventory ABC Class Setup"
      lblStatus = "Caution: The ABC Class Setup Has Not Been Initialized"
   Else
      lblStatus = "The ABC Class Setup Has Been Initialized"
   End If
   Set RdoSet = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getsetup"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbAbc_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub cmbAbc_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyCase KeyAscii
   
End Sub

Private Sub cmbAbc_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub cmbAbc_LostFocus(Index As Integer)
   Dim b As Byte
   Dim iRow As Byte
   cmbAbc(Index) = CheckLen(cmbAbc(Index), 2)
   If cmbAbc(Index) = "" Then cmbAbc(Index) = "  "
   For iRow = 0 To cmbAbc(Index).ListCount - 1
      If cmbAbc(Index) = cmbAbc(Index).List(iRow) Then b = 1
   Next
   If b = 0 Then
      'SysSysbeep
      cmbAbc(Index) = lblLoc(Index)
   End If
   sPartABC(Index + iCurrIdx, 4) = cmbAbc(Index)
   
End Sub


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
      cmbPrt.Enabled = True
      cmdFnd.Enabled = True
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
      OpenHelpContext 1603
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
   bResponse = MsgBox("Update To The Current ABC Classes?", _
               ES_YESQUESTION, Caption)
   If bResponse = vbYes Then UpdateParts Else CancelTrans
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      GetSetup
      FillCombos
      FillCombo
      cmbPrt = "ALL"
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
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
   Set MatlMMe01a = Nothing
   
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



Private Sub SelectParts()
   Dim RdoSel As ADODB.Recordset
   Dim iRows As Integer
   
   Dim sParts As String
   Erase sPartABC
   iTotalParts = 0
   lblNum = 0
   ManageBoxes
   
   On Error GoTo DiaErr1
   If cmbPrt <> "ALL" Then sParts = Compress(cmbPrt)
   If cmbLvl = "ALL" Then
      sSql = "SELECT PARTREF,PARTNUM,PADESC,PASTDCOST,PAABC," _
             & "PATOOL FROM PartTable WHERE (PARTREF LIKE '" & sParts _
             & "%' AND PALEVEL<>6 AND PALEVEL<>7 AND PATOOL=0) " _
             & "ORDER BY PARTREF"
   Else
      sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PASTDCOST,PAABC,PATOOL " _
             & "FROM PartTable WHERE (PARTREF LIKE '" & sParts _
             & "%' AND PALEVEL=" & Val(Left(cmbLvl, 1)) & " AND PATOOL=0) " _
             & "ORDER BY PARTREF"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSel, ES_FORWARD)
   If bSqlRows Then
      With RdoSel
         Do Until .EOF
            iRows = iRows + 1
            If iRows > 300 Then
               iRows = iRows - 1
               Exit Do
            End If
            sPartABC(iRows, 0) = "" & Trim(!PartRef)
            sPartABC(iRows, 1) = "" & Trim(!PartNum)
            sPartABC(iRows, 2) = "" & Trim(!PADESC)
            sPartABC(iRows, 3) = "" & Trim(!PAABC)
            sPartABC(iRows, 4) = "" & Trim(!PAABC)
            sPartABC(iRows, 5) = "" & Format$(!PASTDCOST, ES_QuantityDataFormat)
            sPartABC(iRows, 6) = "Standard Cost " & Format$(!PASTDCOST, "#,###,##0.000")
            If iRows < 11 Then
               lblPrt(iRows) = "" & Trim(!PartNum)
               lblDsc(iRows) = "" & Trim(!PADESC)
               lblLoc(iRows) = "" & Trim(!PAABC)
               lblCost(iRows) = "" & Format$(!PASTDCOST, ES_QuantityDataFormat)
               lblCost(iRows).ToolTipText = sPartABC(iRows, 6)
               cmbAbc(iRows) = "" & Trim(!PAABC)
               cmbAbc(iRows).Enabled = True
               cmbAbc(iRows).BackColor = Es_TextBackColor
            End If
            .MoveNext
         Loop
         ClearResultSet RdoSel
      End With
      iCurrIdx = 0
      iTotalParts = iRows
      cmbAbc(1).SetFocus
      cmdUpd.Enabled = True
      If iTotalParts > 10 Then cmdNxt.Enabled = True
      cmdEnd.Enabled = True
      cmbLvl.Enabled = False
      cmbPrt.Enabled = False
      cmdFnd.Enabled = False
      cmdGo(0).Enabled = False
      lblNum = iTotalParts
      If cmbAbc(1).Enabled Then cmbAbc(1).SetFocus
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
   
End Sub

Private Sub ManageBoxes()
   On Error Resume Next
   Dim iRow As Integer
   For iRow = 1 To 10
      lblPrt(iRow) = ""
      lblDsc(iRow) = ""
      lblCost(iRow) = ""
      lblLoc(iRow) = ""
      cmbAbc(iRow) = ""
      cmbAbc(iRow).BackColor = Es_FormBackColor
      cmbAbc(iRow).Enabled = False
   Next
   
End Sub

Private Sub GetNextGroup()
   Dim iRow As Integer
   Dim iEnd As Integer
   ManageBoxes
   On Error Resume Next
   For iRow = 1 To 10
      If iRow + iCurrIdx > iTotalParts Then Exit For
      lblPrt(iRow) = sPartABC(iRow + iCurrIdx, 1)
      lblDsc(iRow) = sPartABC(iRow + iCurrIdx, 2)
      lblCost(iRow) = sPartABC(iRow + iCurrIdx, 5)
      lblCost(iRow).ToolTipText = sPartABC(iRow + iCurrIdx, 6)
      lblLoc(iRow) = sPartABC(iRow + iCurrIdx, 3)
      cmbAbc(iRow) = sPartABC(iRow + iCurrIdx, 4)
      cmbAbc(iRow).Enabled = True
      cmbAbc(iRow).BackColor = Es_TextBackColor
   Next
   If cmbAbc(1).Enabled Then cmbAbc(1).SetFocus
   
End Sub

Private Sub UpdateParts()
   Dim iRows As Integer
   
   clsADOCon.ADOErrNum = 0
   On Error Resume Next
   For iRows = 1 To iTotalParts
      If sPartABC(iRows, 3) <> sPartABC(iRows, 4) Then
         sSql = "UPDATE PartTable SET PAABC='" _
                & sPartABC(iRows, 4) & "' WHERE PARTREF='" _
                & sPartABC(iRows, 0) & "' "
         clsADOCon.ExecuteSQL sSql
      End If
   Next
   If clsADOCon.ADOErrNum = 0 Then
      SysMsg "ABC Classs Where Updated.", True
   Else
      MsgBox "Couldn't Update Selections.", _
         vbInformation, Caption
   End If
   ManageBoxes
   cmbLvl.Enabled = True
   cmbPrt.Enabled = True
   cmdFnd.Enabled = True
   cmdGo(0).Enabled = True
   cmdUpd.Enabled = False
   cmdNxt.Enabled = False
   cmdLst.Enabled = False
   cmdEnd.Enabled = False
   
End Sub

Private Sub FillCombo()
   sSql = "Qry_FillSortedParts"
   LoadComboBox cmbPrt
   'FillVendors
   If cmbPrt.ListCount > 0 Then
      cmbPrt = cmbPrt.List(0)
      'bGoodPart = GetAliasedPart()
      'GetAlias True
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "FillCombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub
Private Sub FillCombos()
   Dim RdoCmb As ADODB.Recordset
   Dim b As Byte
   Dim iList As Integer
   For iList = 1 To 10
      cmbAbc(iList).Clear
      cmbAbc(iList).AddItem "  "
   Next
   sSql = "Qry_FillABCCombo"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            For iList = 1 To 10
               AddComboStr cmbAbc(iList).hWnd, Trim(!COABCCODE)
            Next
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   Else
      For iList = 1 To 10
         For b = 65 To 89
            cmbAbc(iList).AddItem Chr$(b) & "+"
            cmbAbc(iList).AddItem Chr$(b)
            cmbAbc(iList).AddItem Chr$(b) & "-"
         Next
         cmbAbc(iList).AddItem Chr$(b) & "+"
         cmbAbc(iList).AddItem Chr$(b)
         cmbAbc(iList).AddItem Chr$(b) & "-"
      Next
      
   End If
   Set RdoCmb = Nothing
End Sub
