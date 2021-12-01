VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form AdmnADf02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Current Logons"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "AdmnADf02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.Timer Timer1 
      Interval        =   64000
      Left            =   120
      Top             =   4920
   End
   Begin VB.CommandButton cmdRef 
      Caption         =   "R&efresh"
      Height          =   315
      Left            =   3840
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Refresh List Of Logons"
      Top             =   480
      Width           =   875
   End
   Begin VB.ListBox lstUsr 
      ForeColor       =   &H00400000&
      Height          =   2790
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Current network logged on users and Time "
      Top             =   1200
      Width           =   4380
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   3840
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4440
      Top             =   4920
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4590
      FormDesignWidth =   4860
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   2520
      TabIndex        =   9
      Top             =   4200
      Width           =   2052
      _ExtentX        =   3625
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Logged On                "
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
      Left            =   3000
      TabIndex        =   7
      ToolTipText     =   "Date and Time Logged On to ES/2000 ERP"
      Top             =   960
      Width           =   1500
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name          "
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
      TabIndex        =   6
      ToolTipText     =   "Network User Logon Name"
      Top             =   960
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Workstation        ___   "
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
      TabIndex        =   5
      ToolTipText     =   "Net Workstation"
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblRef 
      Caption         =   "Refreshing List."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Logons To ES/2000 ERP"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3615
   End
End
Attribute VB_Name = "AdmnADf02a"
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

Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmdCan_Click()
   Timer1.Enabled = False
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 1151
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub



Private Sub cmdRef_Click()
   Timer1.Enabled = False
   FillList
   Timer1.Enabled = True
   
End Sub




Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillList
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub



Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST, ES_RESIZE
   FormatControls
   
   sProgDir = GetSetting("Esi2000", "System", "FilePath", sProgDir)
   bOnLoad = 1
   Show
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdmnADf02a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   z1(1).Caption = " Current Logons To " & sSysCaption
   
End Sub

Private Sub FillList()
   Dim RdoUsr As ADODB.Recordset
   
   On Error GoTo DiaErr1
   lblRef.Visible = True
   lblRef.Refresh
   lstUsr.Clear
   sSql = "use msdb"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "SELECT * FROM SystemUserLog WHERE Log_Off='Active'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoUsr, ES_FORWARD)
   If bSqlRows Then
      With RdoUsr
         Do Until .EOF
            lstUsr.AddItem "" & !Log_WorkStation _
               & !Log_User _
               & !Log_On
            .MoveNext
         Loop
      End With
   End If
   sSql = "use " & sDataBase
   clsADOCon.ExecuteSQL sSql
   
   Sleep 500
   lblRef.Visible = False
   Exit Sub
   
DiaErr1:
   sSql = "use " & sDataBase
   clsADOCon.ExecuteSQL sSql
   MsgBox "User List Wasn't Available.", _
      vbInformation, Caption
   
End Sub


Private Sub Timer1_Timer()
   FillList
   
End Sub







Private Sub SetProgress()
   Dim iList As Integer
   MouseCursor 13
   
   On Error Resume Next
   prg1.Visible = True
   For iList = 1 To 20
      prg1.Value = prg1.Value + 5
      DoEvents
      Sleep 2000
   Next
   MouseCursor 0
   prg1.Visible = False
   prg1.Value = 0
   Refresh
'   sSql = "UPDATE Alerts SET ALERTMSG='' " _
'          & "WHERE ALERTREF=1"
'   RdoCon.Execute sSql, rdExecDirect
   cmdCan.Enabled = True
   
End Sub
