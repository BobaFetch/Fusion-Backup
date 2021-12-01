VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaGLp17a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rolling Income Statement (Report)"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optYTD 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cmbDiv 
      Height          =   315
      Left            =   2520
      Sorted          =   -1  'True
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Enter/Revise A Division (2 char)"
      Top             =   1320
      Width           =   660
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6240
      Top             =   2760
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2820
      FormDesignWidth =   7080
   End
   Begin VB.CheckBox optIna 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   2040
      Width           =   855
   End
   Begin VB.CheckBox optDiv 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   5040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optCon 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtLvl 
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Tag             =   "1"
      Top             =   1680
      Width           =   285
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   3480
      TabIndex        =   2
      Tag             =   "4"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Tag             =   "4"
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   5880
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   5760
      TabIndex        =   14
      Top             =   240
      Width           =   1335
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Display The Report"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Print The Report"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox txtYearBeg 
      Height          =   315
      Left            =   3240
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "4"
      ToolTipText     =   "Enter New Team Member  (15 Char) Or Select From List"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   23
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
      PictureUp       =   "diaGLp17a.frx":0000
      PictureDn       =   "diaGLp17a.frx":0146
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   24
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
      PictureUp       =   "diaGLp17a.frx":028C
      PictureDn       =   "diaGLp17a.frx":03D2
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Year To Date"
      Height          =   285
      Index           =   11
      Left            =   240
      TabIndex        =   27
      Top             =   2280
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   10
      Left            =   3960
      TabIndex        =   26
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   25
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inactive Accounts"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   22
      Top             =   2040
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Exclude Accounts W/O Divisions"
      Height          =   405
      Index           =   6
      Left            =   240
      TabIndex        =   21
      Top             =   5040
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Consolidated"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   20
      Top             =   5520
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Divisionalized Reports Only)"
      Height          =   285
      Index           =   8
      Left            =   3720
      TabIndex        =   19
      Top             =   5040
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(9 For All)"
      Height          =   285
      Index           =   4
      Left            =   3960
      TabIndex        =   18
      Top             =   1680
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through Detail Level"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   17
      Top             =   1680
      Width           =   1545
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Period Ending Date"
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   13
      Top             =   2760
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rolling Beginning Date"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Year Beginning Date"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   11
      Top             =   2880
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "diaGLp17a"
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
'
' diaGLp17a - Rolling Income Statement
'
' Notes: Used the income statement with percentages form and report as a base.
'
' Created:  9/30/01 (nth)
' Revisions:
' 08/01/03 (nth) Added at fourth jet table (FSS) to anchor report structure.
' 08/07/03 (nth) Fixed misc errors per WCK income statement now matchs MCS.
' 02/23/04 (JCW) Divisionalized reports, Misc. Bug fixes
' 01/19/05 (nth) Added option boxs to show or hide YTD and Previous Year Columns.
'
'*************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim vAccounts(10, 4) As Variant
Dim iStart As Integer
Dim iEnd As Integer

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
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   txtBeg = Format(ES_SYSDATE, "mm/01/yy")
   txtEnd = GetMonthEnd(txtBeg)
   txtYearBeg = "01/01/" & Format(txtBeg, "yy")
   'optPrn.Picture = Resources.imgPrn.Picture
   'optDis.Picture = Resources.imgDis.Picture
   GetOptions
   If Trim(txtLvl) = "" Then
      txtLvl = "9"
   End If
   optCon.Value = 1 ' temporary
   bOnLoad = True
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   If bDivisionAccounts(iStart, iEnd) Then
      FillDivisions Me
   Else
      cmbDiv.enabled = False
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "filldivisions"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveOptions
   FormUnload
   On Error Resume Next
   Set diaGLp17a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub



Private Sub PrintReport()
   Dim sCustomReport As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   Dim strCrdate As String
   
   On Error GoTo DiaErr1
   
   sCustomReport = GetCustomReport("fingl17.rpt")
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   
   cCRViewer.SetReportTitle = "fingl17.rpt"
   cCRViewer.ShowGroupTree False
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Title1"
   aFormulaName.Add "Title2"
   aFormulaName.Add "nDetailLevel"
   aFormulaName.Add "StartDate"
   aFormulaName.Add optYTD.Name
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Level " & txtLvl _
                        & " Rolling Income Statement For Year Beginning " & txtYearBeg & "'")
   aFormulaValue.Add CStr("'Period Beginning:  " _
                        & txtBeg & " And Ending:  " & txtEnd & "'")
   aFormulaValue.Add CInt(Val(txtLvl))
   
   strCrdate = year(CDate(txtBeg)) & "," & Month(CDate(txtBeg)) & "," & day(CDate(txtBeg))
   aFormulaValue.Add CStr("#" & strCrdate & "#")
   
   aFormulaValue.Add CInt(optYTD)
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   sSql = ""
   
   cCRViewer.SetReportSelectionFormula (sSql)
   
   cCRViewer.CRViewerSize Me
   ' Set report parameter
   cCRViewer.SetDbTableConnection True
   'cCRViewer.SetTableConnection aRptPara
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
   If (BuildReport) Then
      PrintReport
   Else
      MsgBox "Cannot create the report.", vbInformation, Caption
   End If
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendar Me
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
   If Trim(txtLvl) = "" Or Val(txtLvl) > 9 Or Val(txtLvl) < 1 Then txtLvl = 9
End Sub

Private Sub txtYearBeg_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtYearBeg_LostFocus()
   txtYearBeg = CheckDate(txtYearBeg)
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
   Exit Function
DiaErr1:
   sProcName = "bDivisionAccounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim(txtLvl) _
              & Trim(optIna) _
              & Trim(optDiv) _
              & Trim(optCon) _
              & Trim(optYTD)
   SaveSetting "Esi2000", "EsiFina", Me.Name, sOptions
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
      txtLvl = Left(sOptions, 1)
      optIna = Mid(sOptions, 2, 1)
      optDiv = Mid(sOptions, 3, 1)
      optCon = Mid(sOptions, 4, 1)
      optYTD = Mid(sOptions, 5, 1)
   Else
      txtLvl = "9"
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = TTDEFAULT
   End If
End Sub

Private Function BuildReport() As Boolean

   Dim strBegDate As String
   Dim strEndDate As String
   Dim strBegYear As String
   
   On Error GoTo DiaErr1
   
   txtEnd = GetMonthEnd(txtBeg)
   txtYearBeg = "01/01/" & Format(txtBeg, "yy")
   
   strBegDate = txtBeg
   strEndDate = txtEnd
   strBegYear = txtYearBeg
   
   sSql = "RptRollingIncStat '" & strBegDate & "', '" & strEndDate & "', '" & strBegYear & "','1'"
   clsADOCon.ExecuteSql sSql

   If (Err.Number <> 0) Then
      BuildReport = False
   Else
      BuildReport = True
   End If
   
   Exit Function
DiaErr1:
   BuildReport = False
   sProcName = "bDivisionAccounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function
