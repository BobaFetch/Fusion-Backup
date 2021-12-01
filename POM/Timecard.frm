VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmTimecard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Daily Employee Time Charges"
   ClientHeight    =   5400
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7365
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Timecard.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5400
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdWeekly 
      Caption         =   "Week"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   840
      Width           =   1800
   End
   Begin VB.CommandButton cmdDaily 
      BackColor       =   &H0000FF00&
      Caption         =   "Day"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   1740
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   840
      Width           =   1800
   End
   Begin VB.CommandButton optPrn 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   2820
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Print The Report"
      Top             =   4020
      UseMaskColor    =   -1  'True
      Width           =   1800
   End
   Begin VB.CommandButton optDis 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Display The Report"
      Top             =   4020
      UseMaskColor    =   -1  'True
      Width           =   1800
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   4920
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4020
      Width           =   1800
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   420
      Top             =   1620
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5400
      FormDesignWidth =   7365
   End
   Begin Crystal.CrystalReport Crw 
      Left            =   8580
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowLeft      =   110
      WindowTop       =   35
      WindowWidth     =   460
      WindowHeight    =   410
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   7
      DiscardSavedData=   -1  'True
      WindowState     =   1
      PrintFileLinesPerPage=   60
      WindowShowProgressCtls=   0   'False
   End
   Begin VB.Label lblOfEnding 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "of"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2160
      TabIndex        =   8
      Top             =   2220
      Width           =   2955
   End
   Begin VB.Image imgRight 
      BorderStyle     =   1  'Fixed Single
      Height          =   900
      Left            =   5820
      Picture         =   "Timecard.frx":000C
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   900
   End
   Begin VB.Image imgLeft 
      BorderStyle     =   1  'Fixed Single
      Height          =   900
      Left            =   720
      Picture         =   "Timecard.frx":D776
      Stretch         =   -1  'True
      Top             =   2820
      Width           =   900
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tue 09/09/2008"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1680
      TabIndex        =   5
      Top             =   3000
      Width           =   4095
   End
   Begin VB.Label lblEmployeeNo 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88888"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   240
      TabIndex        =   2
      Top             =   180
      Width           =   1395
   End
   Begin VB.Label lblEmployeeName 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Joseph Briggs-Stratton"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1740
      TabIndex        =   1
      Top             =   180
      Width           =   5415
   End
End
Attribute VB_Name = "frmTimecard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit

Private Const COLOR_NOTSELECTED = &H8000000F
Private Const COLOR_SELECTED = &HFF00&

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdDaily_Click()
   cmdDaily.BackColor = COLOR_SELECTED
   cmdWeekly.BackColor = COLOR_NOTSELECTED
   lblOfEnding = "of"
End Sub

Private Sub cmdWeekly_Click()
   cmdDaily.BackColor = COLOR_NOTSELECTED
   cmdWeekly.BackColor = COLOR_SELECTED
   lblOfEnding = "ending"
End Sub

'Private Sub cmdHlp_Click(Value As Integer)
'   If cmdHlp Then
'      MouseCursor ccHourglass
'      'OpenWebHelp "hs907"
'      MouseCursor ccArrow
'      cmdHlp = False
'   End If
'
'End Sub

Private Sub Form_Load()
   CenterForm Me
   SetDateLabel Now
   cmdDaily.BackColor = COLOR_SELECTED
   cmdWeekly.BackColor = COLOR_NOTSELECTED
   lblOfEnding = "of"
End Sub

Private Sub SetDateLabel(dt As Date)
   lblDate = WeekdayName(DatePart("w", dt), True) & " " & Format(dt, "mm/dd/yyyy")
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub PrintReport()
   If Me.cmdDaily.BackColor = COLOR_SELECTED Then
      PrintReportDaily
   Else
      PrintReportWeekly
   End If
End Sub

Private Sub PrintReportDaily()
   Dim sDate As String
   Dim sCustomReport As String
   Dim sReportPath As String
   sDate = Format(Mid(lblDate, 5), "yyyy,mm,dd")
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   'MouseCursor ccHourglass
   On Error GoTo DiaErr1
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("admhu03")
   If RunningInIDE Then
      sReportPath = GetSetting("Esi2000", "System", "ReportPath", sReportPath)
   End If
   If sReportPath = "" Then sReportPath = App.Path & "\"

   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & CStr(lblEmployeeNo) & "'")
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{TchdTable.TMEMP}=" & Val(lblEmployeeNo) & " " _
          & "AND {TchdTable.TMDAY}=Date(" & sDate & ") "
   cCRViewer.SetReportSelectionFormula sSql
   'cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName, 1, "", False, True
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
   MouseCursor ccArrow
   
   Exit Sub
   
  
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   'DoModuleErrors Me
   
End Sub

Private Sub PrintReportWeekly()
   Dim sDate As String
   Dim sCustomReport As String
   Dim sReportPath As String
   sDate = Mid(Me.lblDate, 5)
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("admhu05")
   If RunningInIDE Then
      sReportPath = GetSetting("Esi2000", "System", "ReportPath", sReportPath)
   End If
   If sReportPath = "" Then sReportPath = App.Path & "\"
   
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Employee " & CStr(lblEmployeeNo) & " for week ending " & Me.lblDate & "'")
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{TchdTable.TMDAY} in" & CrystalDate(DateAdd("d", -6, sDate)) & " to " & CrystalDate(sDate)
   sSql = sSql & "AND {TchdTable.TMEMP}=" & lblEmployeeNo & " "

   cCRViewer.SetReportSelectionFormula sSql
   'cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName, 1, "", False, True
 
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   MouseCursor ccArrow
   
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub imgLeft_Click()
   'increment date up to current date
   Dim sDate As Date
   sDate = DateAdd("d", -1, Mid(lblDate, 5))
   SetDateLabel sDate
End Sub

Private Sub imgRight_Click()
   'increment date up to one week in the future
   Dim sDate As Date
   sDate = DateAdd("d", 1, Mid(lblDate, 5))
   If DateDiff("d", Now, sDate) <= 7 Then
      SetDateLabel sDate
   End If
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub


'Private Sub SetCrystalAction(frm As Form)
'   'Allow Crystal to catch up. Especially to refresh DAO.
'   'A bug in Crystal that does not propertly refresh data until the
'   'second hit.
'    If crw.DataFiles(1) <> "" Then Sleep 1000
'
'    'set report size
'
'    'set report size
'    crw.WindowState = crptMaximized
'
'    'other settings
'    crw.WindowBorderStyle = crptSizable
'    crw.WindowControlBox = True
'    crw.WindowMaxButton = True
'    crw.WindowMinButton = True
'    crw.WindowShowCancelBtn = True
'    crw.WindowShowCloseBtn = True
'    crw.WindowShowExportBtn = True
'    crw.WindowShowGroupTree = False
'    crw.WindowShowNavigationCtls = True
'    crw.WindowShowPrintBtn = True
'    crw.WindowShowPrintSetupBtn = True
'    crw.WindowShowRefreshBtn = True
'    crw.WindowShowZoomCtl = True
'    crw.WindowShowSearchBtn = False
'
'    'set connection
'    crw.Connect = "uid=" & gstrSaAdmin & ";pwd=" & gstrSaPassword & ";driver={SQL Server};" _
'                & "server=" & gstrServer & ";database=" & gstrDatabase & ";"
'
'    crw.ReportTitle = frm.Caption
'    crw.WindowTitle = frm.Caption
'    If frm.optPrn.Value = True Then
'    Else
'        crw.Destination = crptToWindow
'    End If
'    On Error Resume Next
'    Err.Clear
'    crw.Action = 1
'    If Err Then
'        MsgBox Err.description
'    End If
'
'End Sub
