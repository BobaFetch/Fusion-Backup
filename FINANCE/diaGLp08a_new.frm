VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaGLp08a_new 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Balance Sheet - New (Report)"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optReg 
      Caption         =   "    "
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   27
      Top             =   2760
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.OptionButton optSum 
      Caption         =   "    "
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   26
      Top             =   3120
      Width           =   615
   End
   Begin VB.ComboBox cmbDiv 
      Height          =   315
      Left            =   2280
      Sorted          =   -1  'True
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Enter/Revise A Division (2 char)"
      Top             =   2280
      Width           =   660
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   2520
      Top             =   3720
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3825
      FormDesignWidth =   7125
   End
   Begin VB.CheckBox optCon 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   4920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optExcWO 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   4560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox optIna 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   1920
      Width           =   735
   End
   Begin VB.CheckBox optPY 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtLvl 
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Tag             =   "1"
      Top             =   1560
      Width           =   285
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Tag             =   "4"
      Top             =   720
      Width           =   1095
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton CmdCan 
      Caption         =   "Close"
      Height          =   375
      Left            =   5880
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   5760
      TabIndex        =   12
      Top             =   240
      Width           =   1335
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Print The Report"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Display The Report"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   10
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaGLp08a_new.frx":0000
      PictureDn       =   "diaGLp08a_new.frx":0146
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   25
      ToolTipText     =   "Show System Printers"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaGLp08a_new.frx":028C
      PictureDn       =   "diaGLp08a_new.frx":03D2
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Summary Account"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   29
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Regular Account"
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   28
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   10
      Left            =   3720
      TabIndex        =   24
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   23
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Divisionalized Reports Only)"
      Height          =   285
      Index           =   8
      Left            =   3720
      TabIndex        =   22
      Top             =   3480
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Consolidated"
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   21
      Top             =   4920
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Exclude Accounts W/O Divisions"
      Height          =   405
      Index           =   6
      Left            =   120
      TabIndex        =   20
      Top             =   4440
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Inactive Accounts"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   19
      Top             =   1920
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Previous Year Difference"
      Height          =   435
      Index           =   4
      Left            =   120
      TabIndex        =   18
      Top             =   3960
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through Detail Level"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(9 For All)"
      Height          =   285
      Index           =   1
      Left            =   3720
      TabIndex        =   16
      Top             =   1560
      Width           =   945
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Period Beginning"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Period Ending"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "diaGLp08a_new"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'*************************************************************************************
' diaGLp08a_new - Balance Sheet
'
' Notes:
'
' Created: 03/22/01 (nth)
' Revisions:
'   08/07/03 (nth) Fix errors per WCK.
'   10/15/03 (nth) more revisions per WCK / now ties to MCS / ESI balance sheet
'   02/23/04 (JCW) Divisionalized Report
'   08/16/04 (nth) Added getoptions and saveoptions
'
'*************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim iStart As Integer
Dim iEnd As Integer

Dim vAccounts(10, 4) As Variant

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*************************************************************************************

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "No Subject Help"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub Form_Activate()
   Dim iSumAcct As Integer
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      
      iSumAcct = GetTopSumAcctFlag
      If (iSumAcct = 1) Then
         optSum.enabled = True
      Else
         optSum.enabled = False
      End If
      
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   'txtBeg = Format(Now, "mm/01/yy")
   'txtEnd = GetMonthEnd(txtBeg)
   txtLvl = 9
   'optPrn.Picture = Resources.imgPrn.Picture
   'optDis.Picture = Resources.imgDis.Picture
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub cmbDiv_LostFocus()
   On Error Resume Next
   cmbDiv = CheckLen(cmbDiv, 2)
   If Trim(cmbDiv) <> "" And Not bValidElement(cmbDiv) Then
      cmbDiv = ""
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveOptions
   FormUnload
   On Error Resume Next
   Set diaGLp08a_new = Nothing
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub FillCombo()
   If bDivisionAccounts(iStart, iEnd) Then
      sProcName = "filldivisions"
      FillDivisions Me
   Else
      cmbDiv.enabled = False
   End If
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub PrintReport()
   Dim sCustomReport As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   On Error GoTo DiaErr1
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   
   If (optReg = True) Then
      sCustomReport = GetCustomReport("fingl08_new.rpt")
      'cCRViewer.SetReportTitle = "fingl08_new.rpt"   ' has to happen after setreportfilename
   Else
      sCustomReport = GetCustomReport("fingl08Top.rpt")
      cCRViewer.SetReportTitle = "fingl08Top.rpt"
   End If
   
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.crxReport.ReportTitle = "fingl08_new.rpt"
   cCRViewer.ShowGroupTree False
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Title1"
   aFormulaName.Add "Title2"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "nDetailLevel"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Level " & txtLvl _
                        & " Balance Sheet For The Period Ending " & txtEnd & "'")
   aFormulaValue.Add CStr("'Period Beginning:  " _
                        & txtBeg & " And Ending:  " & txtEnd & "'")
   aFormulaValue.Add CStr("'Requested By: " & sInitials & "'")
                        
   aFormulaValue.Add CInt(Val(txtLvl))
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   If (optReg = True) Then
      sSql = "({RptAcctBalanceSheet;1.SUMCURBAL} <> 0) OR ({RptAcctBalanceSheet;1.SUMPREVBAL} <> 0)"
   Else
      sSql = "({RptAcctTopBalanceSheet;1.SUMCURBAL} <> 0) OR ({RptAcctTopBalanceSheet;1.SUMPREVBAL} <> 0)"
   End If
   
   cCRViewer.SetReportSelectionFormula (sSql)
   
   cCRViewer.CRViewerSize Me
   ' Set report parameter
   cCRViewer.SetDbTableConnection True
   ' report parameter
   aRptPara.Add CStr(txtBeg)
   aRptPara.Add CStr(txtEnd)
   aRptPara.Add CStr(optIna)
   
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("String")
   ' Set report parameter
   cCRViewer.SetReportDBParameters aRptPara, aRptParaType
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
   MouseCursor 0
   
   Exit Sub
DiaErr1:
   sProcName = "PrintReport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub ShowPrinters_Click(Value As Integer)
   'SysPrinters.Show
   'ShowPrinters.Value = True
End Sub

Private Sub txtBeg_LostFocus()
   txtBeg = CheckDate(txtBeg)
End Sub

Private Sub txtend_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEnd_LostFocus()
   txtEnd = CheckDate(txtEnd)
End Sub

Private Sub txtLvl_LostFocus()
   If Val(txtLvl) < 1 Or Val(txtLvl) > 9 Then
      txtLvl = 9
   End If
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendar Me
End Sub

Private Function bDivisionAccounts(iStart As Integer, iEnd As Integer) As Boolean
   Dim RdoDiv As ADODB.Recordset
   On Error GoTo DiaErr1
   
   sSql = "SELECT COGLDIVISIONS, COGLDIVSTARTPOS, COGLDIVENDPOS FROM ComnTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDiv)
   If bSqlRows Then
      With RdoDiv
         If Val("" & !COGLDIVISIONS) <> 0 Then
            If Val(!COGLDIVSTARTPOS) <> 0 And Val(!COGLDIVENDPOS) <> 0 Then
               iStart = Val(!COGLDIVSTARTPOS)
               iEnd = Val(!COGLDIVENDPOS)
               bDivisionAccounts = True
            End If
         End If
      End With
   End If
   Set RdoDiv = Nothing
   Exit Function
DiaErr1:
   sProcName = "bDivisionAccounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Function bValidElement(cmbCombo As ComboBox) As Boolean
   Dim i As Integer
   On Error GoTo DiaErr1
   If cmbCombo.ListCount > 0 Then
      For i = 0 To cmbCombo.ListCount - 1
         If Val(cmbCombo.List(i)) = Val(cmbCombo.Text) Then
            bValidElement = True
            cmbCombo.ListIndex = i
         End If
      Next
   End If
   Exit Function
   
DiaErr1:
   sProcName = "bValidElement"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim(txtBeg.Text) & Trim(txtEnd.Text)
   SaveSetting "Esi2000", "EsiFina", Me.Name, sOptions
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   Dim dToday As Integer
   
   dToday = CInt(Mid(Format(Now, "mm/dd/yy"), 4, 2))
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)

   
   If Len(Trim(sOptions)) > 0 Then
     
     If dToday < 21 Then
      txtBeg = Mid(sOptions, 1, 8)
      txtEnd = Mid(sOptions, 9, 8)
     Else
      txtBeg = Format(Now, "mm/01/yy")
      txtEnd = GetMonthEnd(txtBeg)
     End If
   
   End If
   

   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = TTDEFAULT
   End If
End Sub
