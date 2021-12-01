VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form AdmnADf01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Reports"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFile 
      DisabledPicture =   "AdmnADf01a.frx":0000
      Enabled         =   0   'False
      Height          =   280
      Left            =   6240
      MaskColor       =   &H8000000F&
      Picture         =   "AdmnADf01a.frx":04C2
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   280
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "AdmnADf01a.frx":0984
      Style           =   1  'Graphical
      TabIndex        =   56
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtRpt 
      Height          =   285
      Left            =   2280
      TabIndex        =   51
      Text            =   "*.rpt"
      ToolTipText     =   "Enter A Leading Char Search Or *.rpt To See All"
      Top             =   10
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   3720
      TabIndex        =   52
      ToolTipText     =   "Click To Insert A File In The Current TextBox"
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "C&ancel"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5100
      TabIndex        =   53
      ToolTipText     =   "Cancel Without Saving"
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "S&elect"
      Height          =   285
      Left            =   5880
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txtNew 
      Enabled         =   0   'False
      Height          =   285
      Index           =   10
      Left            =   3600
      TabIndex        =   22
      Tag             =   "2"
      ToolTipText     =   "Enter The Custom Report Name (No Extension)"
      Top             =   5400
      Width           =   2415
   End
   Begin VB.TextBox txtOld 
      Enabled         =   0   'False
      Height          =   285
      Index           =   10
      Left            =   1080
      TabIndex        =   21
      Tag             =   "2"
      ToolTipText     =   "Enter The Standard Report Name (No Extension)"
      Top             =   5400
      Width           =   2415
   End
   Begin VB.TextBox txtNew 
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   3600
      TabIndex        =   20
      Tag             =   "2"
      ToolTipText     =   "Enter The Custom Report Name (No Extension)"
      Top             =   5040
      Width           =   2415
   End
   Begin VB.TextBox txtOld 
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   1080
      TabIndex        =   19
      Tag             =   "2"
      ToolTipText     =   "Enter The Standard Report Name (No Extension)"
      Top             =   5040
      Width           =   2415
   End
   Begin VB.TextBox txtNew 
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   3600
      TabIndex        =   18
      Tag             =   "2"
      ToolTipText     =   "Enter The Custom Report Name (No Extension)"
      Top             =   4680
      Width           =   2415
   End
   Begin VB.TextBox txtOld 
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   1080
      TabIndex        =   17
      Tag             =   "2"
      ToolTipText     =   "Enter The Standard Report Name (No Extension)"
      Top             =   4680
      Width           =   2415
   End
   Begin VB.TextBox txtNew 
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   3600
      TabIndex        =   16
      Tag             =   "2"
      ToolTipText     =   "Enter The Custom Report Name (No Extension)"
      Top             =   4320
      Width           =   2415
   End
   Begin VB.TextBox txtOld 
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   1080
      TabIndex        =   15
      Tag             =   "2"
      ToolTipText     =   "Enter The Standard Report Name (No Extension)"
      Top             =   4320
      Width           =   2415
   End
   Begin VB.TextBox txtNew 
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   3600
      TabIndex        =   14
      Tag             =   "2"
      ToolTipText     =   "Enter The Custom Report Name (No Extension)"
      Top             =   3960
      Width           =   2415
   End
   Begin VB.TextBox txtOld 
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   1080
      TabIndex        =   13
      Tag             =   "2"
      ToolTipText     =   "Enter The Standard Report Name (No Extension)"
      Top             =   3960
      Width           =   2415
   End
   Begin VB.TextBox txtNew 
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   3600
      TabIndex        =   12
      Tag             =   "2"
      ToolTipText     =   "Enter The Custom Report Name (No Extension)"
      Top             =   3600
      Width           =   2415
   End
   Begin VB.TextBox txtOld 
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   1080
      TabIndex        =   11
      Tag             =   "2"
      ToolTipText     =   "Enter The Standard Report Name (No Extension)"
      Top             =   3600
      Width           =   2415
   End
   Begin VB.TextBox txtNew 
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   3600
      TabIndex        =   10
      Tag             =   "2"
      ToolTipText     =   "Enter The Custom Report Name (No Extension)"
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox txtOld 
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   1080
      TabIndex        =   9
      Tag             =   "2"
      ToolTipText     =   "Enter The Standard Report Name (No Extension)"
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox txtNew 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   3600
      TabIndex        =   8
      Tag             =   "2"
      ToolTipText     =   "Enter The Custom Report Name (No Extension)"
      Top             =   2880
      Width           =   2415
   End
   Begin VB.TextBox txtOld 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   1080
      TabIndex        =   7
      Tag             =   "2"
      ToolTipText     =   "Enter The Standard Report Name (No Extension)"
      Top             =   2880
      Width           =   2415
   End
   Begin VB.TextBox txtNew 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   3600
      TabIndex        =   6
      Tag             =   "2"
      ToolTipText     =   "Enter The Custom Report Name (No Extension)"
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox txtOld 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   1080
      TabIndex        =   5
      Tag             =   "2"
      ToolTipText     =   "Enter The Standard Report Name (No Extension)"
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox txtNew 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   3600
      TabIndex        =   4
      Tag             =   "2"
      ToolTipText     =   "Enter The Custom Report Name (No Extension)"
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox txtOld 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   3
      Tag             =   "2"
      ToolTipText     =   "Enter The Standard Report Name (No Extension)"
      Top             =   2160
      Width           =   2415
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5880
      TabIndex        =   30
      ToolTipText     =   "Apply Changes"
      Top             =   1560
      Width           =   735
   End
   Begin VB.Frame z2 
      Height          =   40
      Left            =   240
      TabIndex        =   29
      Top             =   1440
      Width           =   6495
   End
   Begin VB.ComboBox cmbGroup 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Tag             =   "8"
      ToolTipText     =   "Select Group From List"
      Top             =   960
      Width           =   3255
   End
   Begin VB.ComboBox cmbSection 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Select Section From List"
      Top             =   480
      Width           =   3255
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5760
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   5880
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5880
      FormDesignWidth =   6840
   End
   Begin VB.Label lblReports 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1200
      TabIndex        =   55
      Top             =   1560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Report Path:"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   54
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblNum 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
      Height          =   285
      Index           =   10
      Left            =   240
      TabIndex        =   50
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label lblIdx 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   240
      TabIndex        =   49
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblNum 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9"
      Height          =   285
      Index           =   9
      Left            =   240
      TabIndex        =   48
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label lblIdx 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   240
      TabIndex        =   47
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblNum 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8"
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   46
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label lblIdx 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   45
      Top             =   4680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblNum 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   44
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lblIdx 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   43
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblNum 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   42
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label lblIdx 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   41
      Top             =   3960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblNum 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   40
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblIdx 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   39
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblNum 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   38
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label lblIdx 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   37
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblNum 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   36
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblIdx 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   35
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblNum 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   34
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lblIdx 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   33
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblNum 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   32
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblIdx 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   31
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Custom Report Name                    "
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
      Left            =   3600
      TabIndex        =   28
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Standard Report Name                 "
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
      Left            =   1080
      TabIndex        =   27
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Number "
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
      Left            =   240
      TabIndex        =   26
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Group"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   25
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   24
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "AdmnADf01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'1/17/04 Change Financial Administration to Financial Accounting
'3/16/05 Added missing Groups to fina
'12/20/06 Added InstallCustomReports
'12/21/06 Added columns from AddMissing
Option Explicit
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter


Dim bOnLoad As Byte
Dim bGoodList As Byte
Dim bIndex As Byte

Dim sCurrentBox As String
Dim sOldSection As String
Dim sReport(10, 4) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'12/20/06 Fix Broken Custom Reports
'12/21/06 Removed Credit Managment and added 31-36 (from Administration)

Private Sub InstallCustomReports()
   Dim RdoTest As ADODB.Recordset
   Dim bByte As Byte
   Dim a As Integer
   Dim iRow As Integer
   Dim iList As Integer
   
   Dim sSection As String
   Dim sGroup As String
   On Error Resume Next
   Err = 0
   clsADOCon.ADOErrNum = 0
   
   sSql = "SELECT REPORT_INDEX FROM dbo.CustomReports WHERE REPORT_INDEX = 1"
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum > 0 Then
      Err.Clear
      clsADOCon.ADOErrNum = 0
      
      sSql = "CREATE TABLE dbo.CustomReports (" _
             & "REPORT_INDEX SMALLINT NOT NULL DEFAULT(0)," _
             & "REPORT_SECTION CHAR(30) NULL DEFAULT('')," _
             & "REPORT_GROUP CHAR(30) NULL DEFAULT('')," _
             & "REPORT_REF CHAR(12) NULL DEFAULT('')," _
             & "REPORT_NAME CHAR(12) NULL DEFAULT('')," _
             & "REPORT_CUSTOMREPORT CHAR(12) NULL DEFAULT(''))"
      clsADOCon.ExecuteSQL sSql
      
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.ADOErrNum = 0
         
         sSql = "ALTER TABLE CustomReports ADD Constraint PK_CustomReports_REPORTREF PRIMARY KEY CLUSTERED (REPORT_INDEX) " _
                & "WITH FILLFACTOR=80 "
         clsADOCon.ExecuteSQL sSql
         
         sSql = "CREATE INDEX ReportSection ON CustomReports(REPORT_SECTION) WITH FILLFACTOR = 80"
         clsADOCon.ExecuteSQL sSql
         
         sSql = "CREATE INDEX ReportRef ON CustomReports(REPORT_REF) WITH FILLFACTOR = 80"
         clsADOCon.ExecuteSQL sSql
      End If
      
      Err.Clear
      clsADOCon.ADOErrNum = 0
      sSql = "SELECT REPORT_INDEX FROM CustomReports WHERE REPORT_INDEX=1"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoTest, ES_FORWARD)
      If Not bSqlRows Then
         Err.Clear
         'None, so we'll make some
         For iRow = 1 To 44
            Select Case iRow
               Case 1
                  sSection = "Administration"
                  sGroup = "System"
               Case 2
                  sSection = "Administration"
                  sGroup = "Sales"
               Case 3
                  sSection = "Administration"
                  sGroup = "Production Control"
               Case 4
                  sSection = "Administration"
                  sGroup = "Time Charges"
               Case 5
                  sSection = "Administration"
                  sGroup = "Inventory Management"
               Case 6
                  sSection = "Sales"
                  sGroup = "Order processing"
               Case 7
                  sSection = "Sales"
                  sGroup = "Packing Slips"
               Case 8
                  sSection = "Sales"
                  sGroup = "Bookings/Backlog"
               Case 9
                  sSection = "Engineering"
                  sGroup = "Routings"
               Case 10
                  sSection = "Engineering"
                  sGroup = "Bills Of Material"
               Case 11
                  sSection = "Engineering"
                  sGroup = "Document Control"
               Case 12
                  sSection = "Engineering"
                  sGroup = "Tooling"
               Case 13
                  sSection = "Engineering"
                  sGroup = "Estimating"
               Case 14
                  sSection = "Production Control"
                  sGroup = "Shop Floor Control"
               Case 15
                  sSection = "Production Control"
                  sGroup = "Capacity Planning"
               Case 16
                  sSection = "Production Control"
                  sGroup = "Purchasing"
               Case 17
                  sSection = "Production Control"
                  sGroup = "Time Charges"
               Case 18
                  sSection = "Production Control"
                  sGroup = "Material Requirements"
               Case 19
                  sSection = "Production Control"
                  sGroup = "Material Requisitions"
               Case 20
                  sSection = "Inventory Control"
                  sGroup = "Inventory"
               Case 21
                  sSection = "Inventory Control"
                  sGroup = "Material/Picks"
               Case 22
                  sSection = "Inventory Control"
                  sGroup = "Receiving"
               Case 23
                  sSection = "Inventory Control"
                  sGroup = "Inventory Management"
               Case 24
                  sSection = "Inventory Control"
                  sGroup = "Lot Tracking"
               Case 25
                  sSection = "Quality Assurance"
                  sGroup = "Inspection Reports"
               Case 26
                  sSection = "Quality Assurance"
                  sGroup = "First Article Inspection"
               Case 27
                  sSection = "Quality Assurance"
                  sGroup = "Statistical Process Control"
               Case 28
                  sSection = "Quality Assurance"
                  sGroup = "On Dock Inspection"
               Case 29
                  sSection = "Financial Accounting"
                  sGroup = "Accounts Receivable"
               Case 30
                  sSection = "Financial Accounting"
                  sGroup = "Accounts Payable"
               Case 31
                  sSection = "Financial Accounting"
                  sGroup = "Journals"
               Case 32
                  sSection = "Financial Accounting"
                  sGroup = "Job Costing"
               Case 33
                  sSection = "Financial Accounting"
                  sGroup = "Lot Costing"
               Case 34
                  sSection = "Financial Accounting"
                  sGroup = "General Ledger"
               Case 35
                  sSection = "Financial Accounting"
                  sGroup = "Closing"
               Case 36
                  sSection = "Financial Accounting"
                  sGroup = "Product Costing"
               Case Else
                  bByte = 1
            End Select
            
            If bByte = 0 Then
               clsADOCon.BeginTrans
               For a = 1 To 10
                  iList = iList + 1
                  sSql = "INSERT INTO CustomReports (" _
                         & "REPORT_INDEX,REPORT_SECTION,REPORT_GROUP) " _
                         & "VALUES(" & iList & ",'" _
                         & sSection & "','" & sGroup & "')"
                  clsADOCon.ExecuteSQL sSql
               Next
               clsADOCon.CommitTrans
            Else
               Exit For
            End If
         Next
      End If
   End If
   Set RdoTest = Nothing
   SaveSetting "ES2000", "system", "customreports", 1
   
End Sub

Private Sub cmbSection_Click()
   FillGroups
   
End Sub


Private Sub cmbSection_LostFocus()
   FillGroups
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub




Private Sub cmdEnd_Click()
   ManageBoxes False
   
End Sub

Private Sub cmdFile_Click()
   If File1.Visible = True Then
      txtRpt.Visible = False
      File1.Visible = False
      cmbSection.Width = 3255
      cmbGroup.Width = 3255
   Else
      txtRpt.Visible = True
      File1.Visible = True
      cmbSection.Width = 2600
      cmbGroup.Width = 2600
   End If
   
End Sub



Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 1150
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdSel_Click()
   bGoodList = GetReports()
   If bGoodList Then ManageBoxes True
   
   
End Sub


Private Sub cmdUpd_Click()
   UpdateReports
   
End Sub

Private Sub File1_Click()
   Dim sFile As String
   sFile = Trim(File1)
   If LCase$(Right$(sFile, 4)) = ".rpt" Then _
             sFile = Left$(sFile, Len(sFile) - 4)
   If bIndex = 0 Then bIndex = 1
   If sCurrentBox = "" Then sCurrentBox = "txtOld"
   If sCurrentBox = "txtOld" Then
      txtOld(bIndex) = sFile
   Else
      txtNew(bIndex) = sFile
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      InstallCustomReports
      AddMissing
      UpdateHeaders
      FillSections
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   sSql = "SELECT DISTINCT REPORT_INDEX,REPORT_GROUP FROM CustomReports WHERE " _
          & "REPORT_SECTION= ? ORDER BY REPORT_INDEX"
          
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Size = 30
   AdoParameter.Type = adChar
   
   AdoQry.Parameters.Append AdoParameter
   
          
          
   'Set RdoQry = RdoCon.CreateQuery("", sSql)
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdoQry = Nothing
   Set AdoParameter = Nothing
   Set AdmnADf01a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   On Error GoTo DiaErr1
   txtRpt = "*.rpt"
   lblReports = sReportPath
   txtRpt.Visible = False
   File1.Visible = False
   File1.Path = sReportPath
   File1.Pattern = txtRpt
   cmdEnd.ToolTipText = "Cancel Work Not Updated And Return To Selection"
   Exit Sub
DiaErr1:
   cmdFile.Visible = False
   
End Sub


Private Sub FillSections()
   Dim RdoCmb As ADODB.Recordset
   Dim sLastSection As String
   On Error GoTo DiaErr1
   'Distinct doesn 't work here
   sSql = "SELECT REPORT_SECTION FROM CustomReports " _
          & "ORDER BY REPORT_INDEX"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            If sLastSection <> Trim(!REPORT_SECTION) Then
               AddComboStr cmbSection.hwnd, "" & Trim(!REPORT_SECTION)
               sLastSection = "" & Trim(!REPORT_SECTION)
            End If
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   End If
   If cmbSection.ListCount > 0 Then
      cmbSection = cmbSection.List(0)
      FillGroups
   Else
      MsgBox "Custom Reports Has Not Been Installed.", _
         vbInformation, Caption
      Unload Me
   End If
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   If Left$(Err.Description, 5) = "37000" Then
      MsgBox "Custom Reports Has Not Been Installed On This" & vbCr _
         & "Database. Close All Sections And Re-Logon.", _
         vbInformation, Caption
      Unload Me
      Exit Sub
   End If
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume DiaErr2:
DiaErr2:
   sProcName = "fillsections"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillGroups()
   Dim RdoGrp As ADODB.Recordset
   Dim sLastGroup As String
   
   If sOldSection = cmbSection Then Exit Sub
   sOldSection = cmbSection
   cmbGroup.Clear

   AdoQry.Parameters(0).Value = Trim(cmbSection)
   bSqlRows = clsADOCon.GetQuerySet(RdoGrp, AdoQry, ES_FORWARD)
   If bSqlRows Then
      With RdoGrp
         Do Until .EOF
            If sLastGroup <> Trim(!REPORT_GROUP) Then
               AddComboStr cmbGroup.hwnd, "" & Trim(!REPORT_GROUP)
               sLastGroup = "" & Trim(!REPORT_GROUP)
            End If
            .MoveNext
         Loop
         ClearResultSet RdoGrp
      End With
   End If
   If cmbGroup.ListCount > 0 Then cmbGroup = cmbGroup.List(0)
   Set RdoGrp = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillgroups"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'True to Open Bottom

Private Sub ManageBoxes(Action As Boolean)
   Dim b As Byte
   cmdFile.Value = False
   For b = 1 To 10
      txtOld(b).Enabled = Action
      txtNew(b).Enabled = Action
   Next
   If Action = False Then
      cmbSection.Enabled = True
      cmbGroup.Enabled = True
      cmdSel.Enabled = True
      cmbSection.Width = 3255
      cmbGroup.Width = 3255
      txtRpt.Visible = False
      File1.Visible = False
      cmdUpd.Enabled = False
      cmdEnd.Enabled = False
      cmdFile.Enabled = False
      For b = 1 To 10
         txtOld(b) = ""
         txtNew(b) = ""
         lblIdx(b) = ""
      Next
   Else
      cmbSection.Enabled = False
      cmbGroup.Enabled = False
      cmdSel.Enabled = False
      cmdUpd.Enabled = True
      cmdEnd.Enabled = True
      cmdFile.Enabled = True
   End If
   
End Sub

Private Function GetReports() As Byte
   Dim RdoRpt As ADODB.Recordset
   Dim b As Byte
   
   Erase sReport
   sSql = "SELECT * FROM CustomReports WHERE " _
          & "REPORT_SECTION='" & cmbSection & "' AND " _
          & "REPORT_GROUP='" & cmbGroup & "' " _
          & "ORDER BY REPORT_INDEX"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpt, ES_FORWARD)
   If bSqlRows Then
      With RdoRpt
         Do Until .EOF
            b = b + 1
            sReport(b, 0) = str$(!REPORT_INDEX)
            sReport(b, 1) = "" & Trim(!REPORT_REF)
            sReport(b, 2) = "" & Trim(!REPORT_NAME)
            sReport(b, 3) = "" & Trim(!REPORT_CUSTOMREPORT)
            lblIdx(b) = str$(!REPORT_INDEX)
            txtOld(b) = "" & Trim(!REPORT_NAME)
            txtNew(b) = "" & Trim(!REPORT_CUSTOMREPORT)
            If b = 10 Then Exit Do
            .MoveNext
         Loop
         ClearResultSet RdoRpt
      End With
      GetReports = 1
   Else
      GetReports = 0
   End If
   sCurrentBox = "txtold"
   bIndex = 1
   Set RdoRpt = Nothing
   
End Function

Private Sub txtNew_Click(Index As Integer)
   bIndex = Index
   sCurrentBox = "txtnew"
   
End Sub

Private Sub txtNew_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtNew_KeyPress(Index As Integer, KeyAscii As Integer)
   'KeyCheck KeyAscii
   
End Sub


Private Sub txtNew_LostFocus(Index As Integer)
   txtNew(Index) = Trim(txtNew(Index))
   If LCase$(Right$(txtNew(Index), 4)) = ".rpt" Then
      txtNew(Index) = Left$(txtNew(Index), Len(txtNew(Index)) - 4)
   End If
   txtNew(Index) = CheckLen(txtNew(Index), 30)
   bIndex = Index
   sCurrentBox = "txtNew"
   
End Sub


Private Sub txtOld_Click(Index As Integer)
   bIndex = Index
   sCurrentBox = "txtOld"
   
End Sub

Private Sub txtOld_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtOld_LostFocus(Index As Integer)
   bIndex = Index
   sCurrentBox = "txtOld"
   txtOld(Index) = Trim(txtOld(Index))
   If LCase$(Right$(txtOld(Index), 4)) = ".rpt" Then
      txtOld(Index) = Left$(txtOld(Index), Len(txtOld(Index)) - 4)
   End If
   txtOld(Index) = CheckLen(txtOld(Index), 30)
   
End Sub



'Reorder these for storage

Private Sub UpdateReports()
   Dim RdoUpd As ADODB.Recordset
   Dim a As Integer
   Dim b As Byte
   Dim sTemp(10, 4) As String
   For b = 1 To 10
      sTemp(b, 0) = lblIdx(b)
      If Trim(txtOld(b)) <> "" Then
         a = a + 1
         sTemp(a, 1) = LCase$(txtOld(b))
         sTemp(a, 2) = txtOld(b)
         sTemp(a, 3) = txtNew(b)
      End If
   Next
   
   For b = 1 To 10
      lblIdx(b) = sTemp(b, 0)
      txtOld(b) = sTemp(b, 2)
      txtNew(b) = sTemp(b, 3)
   Next
   
   'Update them
   For b = 1 To 10
      sSql = "SELECT * FROM CustomReports WHERE " _
             & "REPORT_INDEX=" & sTemp(b, 0) & " "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoUpd, ES_DYNAMIC)
      If bSqlRows Then
         With RdoUpd
            '.Edit
            !REPORT_REF = Compress(sTemp(b, 1))
            !REPORT_NAME = sTemp(b, 2)
            !REPORT_CUSTOMREPORT = sTemp(b, 3)
            .Update
         End With
      End If
   Next
   ManageBoxes False
   Set RdoUpd = Nothing
   
End Sub

Private Sub txtRpt_KeyPress(KeyAscii As Integer)
   KeyCheck KeyAscii
   
End Sub


Private Sub txtRpt_LostFocus()
   txtRpt = LCase(txtRpt)
   If Trim(Right(txtRpt, 4)) <> ".rpt" Then txtRpt = txtRpt & ".rpt"
   File1.Pattern = txtRpt
   
End Sub



Private Sub UpdateHeaders()
   sSql = "UPDATE CustomReports SET REPORT_SECTION='Financial Accounting' " _
          & "WHERE REPORT_SECTION='Financial Administration'"
   clsADOCon.ExecuteSQL sSql
   
End Sub

Private Sub AddMissing()
   Dim RdoMissed As ADODB.Recordset
   Dim iList As Integer
   Dim iRow As Integer
   
   'Check and bail if it's new
   sSql = "SELECT REPORT_GROUP FROM CustomReports WHERE REPORT_GROUP='General Ledger'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoMissed, ES_FORWARD)
   If Not bSqlRows Then
      'RdoMissed.Close
      iRow = 340
      sSql = "DELETE FROM CustomReports WHERE REPORT_GROUP='Credit Management'"
      clsADOCon.ExecuteSQL sSql
      For iList = 1 To 10
         iRow = iRow + 1
         sSql = "INSERT INTO CustomReports (REPORT_INDEX,REPORT_SECTION," _
                & "REPORT_GROUP) VALUES(" _
                & iRow & ",'Financial Accounting','General Ledger')"
         clsADOCon.ExecuteSQL sSql
      Next
      
      For iList = 1 To 10
         iRow = iRow + 1
         sSql = "INSERT INTO CustomReports (REPORT_INDEX,REPORT_SECTION," _
                & "REPORT_GROUP) VALUES(" _
                & iRow & ",'Financial Accounting','Closing')"
         clsADOCon.ExecuteSQL sSql
      Next
      
      For iList = 1 To 10
         iRow = iRow + 1
         sSql = "INSERT INTO CustomReports (REPORT_INDEX,REPORT_SECTION," _
                & "REPORT_GROUP) VALUES(" _
                & iRow & ",'Financial Accounting','Product Costing')"
         clsADOCon.ExecuteSQL sSql
      Next
   End If
   
   'add Time Management -
   sSql = "SELECT REPORT_GROUP FROM CustomReports WHERE REPORT_GROUP='Database Maintenance'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoMissed, ES_FORWARD)
   If Not bSqlRows Then
      RdoMissed.Close
      iRow = 370
      'sSql = "DELETE FROM CustomReports WHERE REPORT_GROUP='Credit Management'"
      'RdoCon.Execute sSql, rdExecDirect
      For iList = 1 To 10
         iRow = iRow + 1
         sSql = "INSERT INTO CustomReports (REPORT_INDEX,REPORT_SECTION," _
                & "REPORT_GROUP) VALUES(" _
                & iRow & ",'Time Management','Time Charges')"
         clsADOCon.ExecuteSQL sSql
      Next
      
      For iList = 1 To 10
         iRow = iRow + 1
         sSql = "INSERT INTO CustomReports (REPORT_INDEX,REPORT_SECTION," _
                & "REPORT_GROUP) VALUES(" _
                & iRow & ",'Time Management','Time and Attendance')"
         clsADOCon.ExecuteSQL sSql
      Next
      
      For iList = 1 To 10
         iRow = iRow + 1
         sSql = "INSERT INTO CustomReports (REPORT_INDEX,REPORT_SECTION," _
                & "REPORT_GROUP) VALUES(" _
                & iRow & ",'Administration','Database Maintenance')"
         clsADOCon.ExecuteSQL sSql
      Next
   End If
   Set RdoMissed = Nothing

End Sub


