VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaGLp06a_new 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Income Statement - New (Report)"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optPre 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   3240
      Width           =   855
   End
   Begin VB.CheckBox optYTD 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   3000
      Width           =   855
   End
   Begin VB.ComboBox cmbDiv 
      Height          =   315
      Left            =   2520
      Sorted          =   -1  'True
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Enter/Revise A Division (2 char)"
      Top             =   2040
      Width           =   660
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5940
      Top             =   3000
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3825
      FormDesignWidth =   7080
   End
   Begin VB.CheckBox optIna 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   2760
      Width           =   855
   End
   Begin VB.CheckBox optDiv 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   5040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optCon 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtLvl 
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Tag             =   "1"
      Top             =   2400
      Width           =   285
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   2520
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   5880
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   0
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   5760
      TabIndex        =   15
      Top             =   240
      Width           =   1335
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Display The Report"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Print The Report"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox txtYearBeg 
      Height          =   315
      Left            =   2520
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "4"
      ToolTipText     =   "Enter New Team Member  (15 Char) Or Select From List"
      Top             =   720
      Width           =   1095
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   24
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
      PictureUp       =   "diaGLp06a_new.frx":0000
      PictureDn       =   "diaGLp06a_new.frx":0146
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
      PictureUp       =   "diaGLp06a_new.frx":028C
      PictureDn       =   "diaGLp06a_new.frx":03D2
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Previous Year"
      Height          =   285
      Index           =   12
      Left            =   240
      TabIndex        =   29
      Top             =   3240
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Year To Date"
      Height          =   285
      Index           =   11
      Left            =   240
      TabIndex        =   28
      Top             =   3000
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   10
      Left            =   3960
      TabIndex        =   27
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   26
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inactive Accounts"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   23
      Top             =   2760
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Exclude Accounts W/O Divisions"
      Height          =   405
      Index           =   6
      Left            =   240
      TabIndex        =   22
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
      TabIndex        =   21
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
      TabIndex        =   20
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
      TabIndex        =   19
      Top             =   2400
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through Detail Level"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   18
      Top             =   2400
      Width           =   1545
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   17
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Period Ending Date"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   14
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Period Beginning Date"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   13
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Year Beginning Date"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   720
      Width           =   2415
   End
End
Attribute VB_Name = "diaGLp06a_new"
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
' diaGLp06a_new - Income Statement
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
   'Dim iSumAcct As Integer
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      
'      iSumAcct = GetTopSumAcctFlag
'      If (iSumAcct = 1) Then
'         optSum.enabled = True
'      Else
'         optSum.enabled = False
'      End If
      
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   ReopenJet
   sCurrForm = Caption
   txtYearBeg = "01/01/" & Format(ES_SYSDATE, "yy")
   txtBeg = Format(ES_SYSDATE, "mm/01/yy")
   txtEnd = GetMonthEnd(txtBeg)
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
   Set diaGLp06a_new = Nothing
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
   
   On Error GoTo DiaErr1
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   
'  summary for NORPRO merging of two databases has been removed
'   If (optReg = True) Then
      sCustomReport = GetCustomReport("fingl06_new.rpt")
      cCRViewer.SetReportTitle = "fingl06_new.rpt"
'   Else
'      sCustomReport = GetCustomReport("fingl06Top.rpt")
'      cCRViewer.SetReportTitle = "fingl06Top.rpt"
'   End If
   
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.ShowGroupTree False
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Title1"
   aFormulaName.Add "Title2"
   aFormulaName.Add "nDetailLevel"
   aFormulaName.Add optYTD.Name
   aFormulaName.Add optPre.Name
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Level " & txtLvl _
                        & " Income Statement For Year Beginning " & txtYearBeg & "'")
   aFormulaValue.Add CStr("'Period Beginning:  " _
                        & txtBeg & " And Ending:  " & txtEnd & "'")
   aFormulaValue.Add CInt(Val(txtLvl))
   aFormulaValue.Add CInt(optYTD)
   aFormulaValue.Add CInt(optPre)
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
'   If (optReg = True) Then
      sSql = "{RptIncomeStatement;1.SUMCURBAL}  <> 0"
      
      If (CInt(optYTD) = 1) Then
         sSql = sSql & " OR {RptIncomeStatement;1.SUMYTD} <> 0 "
      End If
      
      If (CInt(optPre) = 1) Then
         sSql = sSql & " OR {RptIncomeStatement;1.SUMPREVBAL} <> 0 "
      End If
'   Else
'      sSql = "{RptTopIncomeStatement;1.SUMCURBAL}  <> 0"
'
'      If (CInt(optYTD) = 1) Then
'         sSql = sSql & " OR {RptTopIncomeStatement;1.SUMYTD} <> 0 "
'      End If
'
'      If (CInt(optPre) = 1) Then
'         sSql = sSql & " OR {RptTopIncomeStatement;1.SUMPREVBAL} <> 0 "
'      End If
'   End If
   
   cCRViewer.SetReportSelectionFormula (sSql)
   
   cCRViewer.CRViewerSize Me
   ' Set report parameter
   cCRViewer.SetDbTableConnection True
   ' report parameter
   aRptPara.Add CStr(txtBeg)
   aRptPara.Add CStr(txtEnd)
   aRptPara.Add CStr(txtYearBeg)
   aRptPara.Add CStr(optIna)
   
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("String")
   ' Set report parameter
   cCRViewer.SetReportDBParameters aRptPara, aRptParaType    'must happen AFTER SetDbTableConnection call!
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
   Set RdoDiv = Nothing
   Exit Function
DiaErr1:
   sProcName = "bDivisionAccounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub SaveOptions()
   Dim sOptions As String

   sOptions = Trim(txtBeg.Text) _
              & Trim(txtEnd.Text) _
              & Trim(txtLvl) _
              & Trim(optIna) _
              & Trim(optDiv) _
              & Trim(optCon) _
              & Trim(optYTD) _
              & Trim(optPre)
   SaveSetting "Esi2000", "EsiFina", Me.Name, sOptions
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   Dim dToday As Integer

   On Error Resume Next
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

      txtLvl = Mid(sOptions, 17, 1)
      optIna = Mid(sOptions, 18, 1)
      optDiv = Mid(sOptions, 18, 1)
      optCon = Mid(sOptions, 20, 1)
      optYTD = Mid(sOptions, 21, 1)
      optPre = Mid(sOptions, 22, 1)
   Else: txtLvl = "9"
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = TTDEFAULT
   End If
End Sub
