VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form zGr1Admn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Administration"
   ClientHeight    =   4620
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   5010
   ControlBox      =   0   'False
   Icon            =   "Gr1Admn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Gr1Admn.frx":08CA
   ScaleHeight     =   4620
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton optPrn 
      Height          =   330
      Left            =   1680
      Picture         =   "Gr1Admn.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Print The Report"
      Top             =   4440
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton optDis 
      Height          =   330
      Left            =   1080
      Picture         =   "Gr1Admn.frx":131E
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Display The Report"
      Top             =   4440
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2760
      Top             =   4320
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   19
      ImageHeight     =   23
      MaskColor       =   -2147483633
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr1Admn.frx":149C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr1Admn.frx":1A52
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr1Admn.frx":1E5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr1Admn.frx":1FBA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstSelect 
      Height          =   2790
      Left            =   430
      TabIndex        =   3
      Top             =   600
      Width           =   4065
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   3780
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "    Close (Escape)    "
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   0
      Picture         =   "Gr1Admn.frx":20E8
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin MSComctlLib.TabStrip Tab1 
      Height          =   3972
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4932
      _ExtentX        =   8705
      _ExtentY        =   7011
      MultiRow        =   -1  'True
      TabFixedHeight  =   473
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Edit             "
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&View           "
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Functions"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   480
      Top             =   4092
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4620
      FormDesignWidth =   5010
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2280
      Top             =   4320
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483628
      ImageWidth      =   22
      ImageHeight     =   20
      MaskColor       =   -2147483636
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr1Admn.frx":221B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr1Admn.frx":26CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr1Admn.frx":2B6B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblCustomer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "This Feature Is Not Available"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   840
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   2892
   End
End
Attribute VB_Name = "zGr1Admn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Dim b As Byte
Dim ActiveTab As String

'Private Sub ShowUserLog() ' non-functional.  removed 4/11/19
'   Dim cCRViewer As EsCrystalRptViewer
'   Dim sCustomReport As String
'   Dim aRptPara As New Collection
'   Dim aRptParaType As New Collection
'   Dim aFormulaValue As New Collection
'   Dim aFormulaName As New Collection
'   MouseCursor 13
'   On Error GoTo DiaErr1
'   aFormulaName.Add "CompanyName"
'   aFormulaName.Add "RequestBy"
'   aFormulaValue.Add CStr("'" & sFacility & "'")
'   aFormulaValue.Add CStr("'Requested By: " & sInitials & "'")
'   sSql = ""
'   Set cCRViewer = New EsCrystalRptViewer
'   cCRViewer.Init
'   sCustomReport = GetCustomReport("UserLog")
'   cCRViewer.SetReportFileName sCustomReport, sReportPath
'   cCRViewer.SetReportTitle = sCustomReport
'   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
'   cCRViewer.ShowGroupTree False
'   cCRViewer.SetReportSelectionFormula sSql
'   cCRViewer.CRViewerSize Me
'   cCRViewer.SetDbTableConnection
'   cCRViewer.OpenCrystalReportObject Me, aFormulaName
'
'   cCRViewer.ClearFieldCollection aRptPara
'   cCRViewer.ClearFieldCollection aFormulaName
'   cCRViewer.ClearFieldCollection aFormulaValue
'   MouseCursor 0
'   Exit Sub
'
'DiaErr1:
'   sProcName = "printreport"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub

'Private Sub ShowUserLog1()
'   MouseCursor 13
'   On Error GoTo DiaErr1
'   'SetMdiReportsize MDISect
'   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MDISect.Crw.Formulas(1) = "RequestBy='Requested By: " & sInitials & "'"
'   sSql = ""
'   sCustomReport = GetCustomReport("UserLog")
'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
'   'SetCrystalAction Me
'   MouseCursor 0
'   Exit Sub
'
'DiaErr1:
'   sProcName = "printreport"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub
'
'Private Sub ShowEventLog() ' non-functional.  removed 4/11/19
'   Dim cCRViewer As EsCrystalRptViewer
'   Dim sCustomReport As String
'   Dim aRptPara As New Collection
'   Dim aRptParaType As New Collection
'   Dim aFormulaValue As New Collection
'   Dim aFormulaName As New Collection
'
'   MouseCursor 13
'   On Error GoTo DiaErr1
'   aFormulaName.Add "CompanyName"
'   aFormulaName.Add "RequestBy"
'   aFormulaName.Add "Includes"
'   aFormulaValue.Add CStr("'" & sFacility & "'")
'   aFormulaValue.Add CStr("'Requested By: " & sInitials & "'")
'   aFormulaValue.Add CStr("'" & sSysCaption & " Event Log'")
'   sSql = ""
'   Set cCRViewer = New EsCrystalRptViewer
'   cCRViewer.Init
'   sCustomReport = GetCustomReport("EventLog")
'   cCRViewer.SetReportFileName sCustomReport, sReportPath
'   cCRViewer.SetReportTitle = sCustomReport
'   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
'   cCRViewer.ShowGroupTree False
'   cCRViewer.SetReportSelectionFormula sSql
'   cCRViewer.CRViewerSize Me
'   cCRViewer.SetDbTableConnection
'   cCRViewer.OpenCrystalReportObject Me, aFormulaName
'
'   cCRViewer.ClearFieldCollection aRptPara
'   cCRViewer.ClearFieldCollection aFormulaName
'   cCRViewer.ClearFieldCollection aFormulaValue
'   MouseCursor 0
'   Exit Sub
'
'DiaErr1:
'   sProcName = "printreport"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'
'End Sub

'Private Sub ShowEventLog1()
'   MouseCursor 13
'   On Error GoTo DiaErr1
'   'SetMdiReportsize MDISect
'   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MDISect.Crw.Formulas(1) = "RequestBy='Requested By: " & sInitials & "'"
'   MDISect.Crw.Formulas(2) = "Includes='" & sSysCaption & " Event Log'"
'   sSql = ""
'   sCustomReport = GetCustomReport("EventLog")
'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
'   'SetCrystalAction Me
'   MouseCursor 0
'   Exit Sub
'
'DiaErr1:
'   sProcName = "printreport"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'
'End Sub
Private Sub TruncateIdleLog()
   Dim bResponse As Byte
   Dim iFreeFile As Integer
   Dim sDate As String * 16
   Dim sSection As String * 8
   Dim sForm As String * 12
   Dim sErrNum As String * 10
   Dim sErrSev As String * 2
   Dim sProc As String * 10
   Dim sUserName As String * 20
   Dim sMsg As String
   On Error Resume Next
   sMsg = "The Idle Time Log Should Be Truncated Occassionally. The " & vbCr _
          & "Log Events Should Printed And Be Given To The System Admin " & vbCr _
          & "First. Continue To Truncate And Dump The Current Log?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, "System Function")
   If bResponse = vbYes Then
      If Dir(sFilePath & "EsiClose.log") <> "" Then _
             Kill sFilePath & "EsiClose.log"
      MsgBox "The Idle Time Log Was Truncated.", vbInformation, "System Function"
      iFreeFile = FreeFile
      Open sFilePath & "EsiClose.log" For Append Shared As #FreeFile
      Print #iFreeFile, "ES/2005 ERP closed the following Workstations due to excessive idle time: "
      Print #iFreeFile, vbCr
      Print #iFreeFile, "Workstation        ", "Windows Log On"
      Print #iFreeFile, String$(110, "-")
      Close #iFreeFile
   Else
      CancelTrans
   End If
   
End Sub


Private Sub TruncateWarningLog()
   Dim bResponse As Byte
   Dim iFreeFile As Integer
   Dim sDate As String * 16
   Dim sSection As String * 8
   Dim sForm As String * 12
   Dim sErrNum As String * 10
   Dim sErrSev As String * 2
   Dim sProc As String * 10
   Dim sUserName As String * 20
   Dim sMsg As String
   
   On Error Resume Next
   sMsg = "The System Event Log Should Be Truncated Occassionally. The " & vbCr _
          & "Log Events Should Printed And Be Given To The System Trainer " & vbCr _
          & "First. Continue To Truncate And Dump The Current Log?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, "System Function")
   If bResponse = vbYes Then
      If Dir(sFilePath & "EsiError.log") <> "" Then _
             Kill sFilePath & "EsiError.log"
      clsADOCon.ExecuteSql "use msdb"
      sSql = "truncate table SystemEvents"
      clsADOCon.ExecuteSql sSql
      clsADOCon.ExecuteSql "use " & sDataBase
      
      MsgBox "The Event Log Was Truncated.", vbInformation, "System Function"
      iFreeFile = FreeFile
      '        Open sFilePath & "EsiError.log" For Append Shared As #iFreeFile
      '            sDate = "Date"
      '            sSection = "Section"
      '            sForm = "Form"
      '            sUserName = "User"
      '            sErrNum = "Err"
      '            sProc = "Procedure"
      '            Print #iFreeFile, "System Event Log"
      '            Print #iFreeFile, "Note: Information messages and User input notices are not recorded"
      '            Print #iFreeFile, sDate; "    "; sSection; "  "; sForm; "  "; sUserName; _
      '              "          "; sErrNum; "   "; sErrSev; " "; sProc; " "
      '            Print #iFreeFile, String(156, "-")
      '            Close #iFreeFile
   Else
      CancelTrans
   End If
End Sub


Private Sub WhereAmI()
   Dim sMsg As String
   sMsg = "The Files Are In: " _
          & App.Path & vbCr _
          & "My Files Setting is: " _
          & sFilePath & vbCr _
          & "My Report Setting Is: " _
          & sReportPath
   MsgBox sMsg, vbInformation, sSysCaption
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 923
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   cUR.CurrentGroup = "Admn"
   
End Sub

Private Sub Form_Deactivate()
   On Error Resume Next
   Hide
   
End Sub

Private Sub Form_Load()
   ActiveTab = "E"
   'If bSecSet = 1 Then User.Group1 = 1
   'If UCase$(cur.CurrentUser) = "ADMINISTRATOR" Then User.Group1 = 1
   'If UCase$(cur.CurrentUser) = "LARRYH" Then User.Group1 = 1
   FillEdit
   FormLoad Me, ES_DONTLIST, ES_RESIZE
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set zGr1Admn = Nothing
   
End Sub




Private Sub lstSelect_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then Exit Sub
   MouseCursor 13
   b = 1
   If ActiveTab = "E" Then
      Select Case lstSelect.ListIndex
         Case 0
            AdmnADe01a.Show
         Case 1
            EstiESe01b.Show
         Case 2
            AdmnADe03a.Show
         Case 3
            AdmnADe04a.Show
         Case 4
            'If bSecSet = 1 Then
            AdmnUuser2.Show
            'Else
            '    AdmnUuser.Show
            'End If
         Case 5
            'AdmnUnewu.Show
            AdmnUsecr.Show
         Case Else
            b = 0
      End Select
   ElseIf ActiveTab = "V" Then
      Select Case lstSelect.ListIndex
         Case 0
            AdmnADp06a.Show
         Case 1
            AdmnADp07a.Show
'         Case 2     ' non-functional.  removed 4/11/19
'            ShowEventLog
'            b = 0
'         Case 3     ' non-functional.  removed 4/11/19
'            ShowUserLog
'            b = 0
'         Case 2
'            OpenWebPage sFilePath & "EsiClose.log"
'            b = 0
         Case 2
            WhereAmI
            b = 0
         Case 3
            AdmnADp08a.Show
         Case Else
            b = 0
      End Select
   Else 'Functions
      Select Case lstSelect.ListIndex
         Case 0
            AdmnADf01a.Show
         Case 1
            AdmnADf06a.Show
         Case 2
            AdmnADf04a.Show
         Case 3
            AdmnADf02a.Show
         Case 4
            AdmnADf03a.Show
         Case 5
            TruncateWarningLog
            b = 0
         Case 6
            TruncateIdleLog
            b = 0
         Case 7
            SysMessage.Show
         Case 8
            PomMessage.Show
         Case 9
            AdmnADf05a.Show
         Case Else
            b = 0
      End Select
   End If
   If b = 1 Then Hide Else MouseCursor 0
   
End Sub


Private Sub lstSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   
   lstSelect_KeyPress 13
End Sub


Private Sub tab1_Click()
   bActiveTab(1) = Tab1.SelectedItem.Index
   Select Case bActiveTab(1)
      Case 1
         FillEdit
      Case 2
         FillView
      Case 3
         FillFunctions
   End Select
   
End Sub

Private Sub FillEdit()
   ActiveTab = "E"
   lstSelect.Clear
   If Secure.UserAdmnG1E = 1 Then
      'If User.Group1 Then
      lstSelect.AddItem "System Settings "
      lstSelect.AddItem "Estimating Parameters "
      lstSelect.AddItem "Standard Comments "
      lstSelect.AddItem "Status Codes"
      'the next must be the last on the list
      'If User.Adduser = 1 Or ...
      If SecPw.UserAdmn = 1 Or RunningInIDE Then
         lstSelect.AddItem "User Manager "
      End If
      If bSecSet = 0 And bCannotOpen = 0 Then lstSelect.AddItem "Advanced Security Setup "
      lstSelect.Enabled = True
      'Else
      '    lstSelect.Enabled = False
      '    lstSelect.AddItem "No Group Permissions"
      'End If
   Else
      lstSelect.Enabled = False
      lstSelect.AddItem "No User Permissions"
   End If
   
End Sub

Private Sub FillView()
   ActiveTab = "V"
   lstSelect.Clear
   If Secure.UserAdmnG1V = 1 Then
      'If User.Group1 Then
      lstSelect.AddItem "Standard Comments"
      lstSelect.AddItem "Custom Report Settings"
'      lstSelect.AddItem "View Event Log" ' non-functional.  removed 4/11/19
'      lstSelect.AddItem "View User Log"  ' non-functional.  removed 4/11/19
'     lstSelect.AddItem "View Idle Time Closing Log"   ' non-functional.  removed 5/16/19
      lstSelect.AddItem "Show The Application Path (Where Am I?)"
      lstSelect.AddItem "User List"
      lstSelect.Enabled = True
      'Else
      '    lstSelect.Enabled = False
      '    lstSelect.AddItem "No Group Permissions"
      'End If
   Else
      lstSelect.Enabled = False
      lstSelect.AddItem "No User Permissions"
   End If
   
End Sub

Public Sub FillFunctions()
   ActiveTab = "F"
   lstSelect.Clear
   If Secure.UserAdmnG1F = 1 Then
      'If User.Group1 Then
      lstSelect.AddItem "Custom Reports"
      lstSelect.AddItem "Custom Fields"
      lstSelect.AddItem "Company Logo"
      lstSelect.AddItem "Current Logons"
      lstSelect.AddItem "Delete A Standard Comment"
      lstSelect.AddItem "Truncate Event Log"
      lstSelect.AddItem "Truncate Idle Time Closing Log"
      lstSelect.AddItem "Broadcast Message"
      lstSelect.AddItem "Set Message for POM Users"
      lstSelect.AddItem "Delete Status Code"
      lstSelect.Enabled = True
      'Else
      '    lstSelect.Enabled = False
      '    lstSelect.AddItem "No Group Permissions"
      'End If
   Else
      lstSelect.Enabled = False
      lstSelect.AddItem "No User Permissions"
   End If
   
End Sub
