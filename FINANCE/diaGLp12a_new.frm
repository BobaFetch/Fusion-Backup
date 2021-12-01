VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaGLp12a_New 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pro Forma Income Statement - New"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   6750
   Begin VB.ComboBox cmbDiv 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Enter/Revise A Division (2 char)"
      Top             =   2040
      Width           =   660
   End
   Begin VB.ComboBox cmbPer2 
      Height          =   315
      ItemData        =   "diaGLp12a_new.frx":0000
      Left            =   1440
      List            =   "diaGLp12a_new.frx":0002
      TabIndex        =   2
      Tag             =   "1"
      Top             =   1560
      Width           =   615
   End
   Begin VB.ComboBox cmbPer1 
      Height          =   315
      ItemData        =   "diaGLp12a_new.frx":0004
      Left            =   1440
      List            =   "diaGLp12a_new.frx":0006
      TabIndex        =   1
      Tag             =   "1"
      Top             =   1080
      Width           =   615
   End
   Begin VB.ComboBox cmbYer 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Tag             =   "1"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   5520
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   5400
      TabIndex        =   11
      Top             =   240
      Width           =   1335
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Print The Report"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Display The Report"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.TextBox txtLvl 
      Height          =   285
      Left            =   2400
      TabIndex        =   4
      Tag             =   "1"
      Top             =   2640
      Width           =   285
   End
   Begin VB.CheckBox optCon 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optDiv 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optIna 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   3000
      Width           =   735
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6120
      Top             =   2880
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3510
      FormDesignWidth =   6750
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   12
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
      PictureUp       =   "diaGLp12a_new.frx":0008
      PictureDn       =   "diaGLp12a_new.frx":014E
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   13
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
      PictureUp       =   "diaGLp12a_new.frx":0294
      PictureDn       =   "diaGLp12a_new.frx":03DA
   End
   Begin VB.Label z1 
      Caption         =   "Division"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   31
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label z1 
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   30
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label lblYerStart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4080
      TabIndex        =   29
      Top             =   600
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Year Beginning"
      Height          =   255
      Index           =   13
      Left            =   2760
      TabIndex        =   28
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   27
      Top             =   1560
      Width           =   1515
   End
   Begin VB.Label lblEnd 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4080
      TabIndex        =   26
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblStart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4080
      TabIndex        =   25
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Period Ending"
      Height          =   255
      Index           =   11
      Left            =   2760
      TabIndex        =   24
      Top             =   1560
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Period Starting"
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   23
      Top             =   1080
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From Period"
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   22
      Top             =   1080
      Width           =   1515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      Height          =   255
      Index           =   16
      Left            =   240
      TabIndex        =   21
      Top             =   600
      Width           =   915
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   20
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through Detail Level"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   19
      Top             =   2640
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(9 For All)"
      Height          =   285
      Index           =   4
      Left            =   3240
      TabIndex        =   18
      Top             =   2640
      Width           =   705
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Divisionalized Reports Only)"
      Height          =   285
      Index           =   8
      Left            =   3240
      TabIndex        =   17
      Top             =   3600
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Consolidated"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   16
      Top             =   4080
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Exclude Accounts W/O Divisions"
      Height          =   405
      Index           =   6
      Left            =   240
      TabIndex        =   15
      Top             =   3600
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Inactive Accounts"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   14
      Top             =   3000
      Width           =   2025
   End
End
Attribute VB_Name = "diaGLp12a_New"
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
' diaGLp12a_New - Proforma Income
'
' Notes: Used the income statement as a base.
'
' Created:  2/06/04 (nth)
' Revisions:
'   2/23/04 (JCW) Fixed misc. Bugs
'
'*************************************************************************************

Option Explicit

Dim rdoPer As ADODB.Recordset
Dim bOnLoad As Byte
Dim vAccounts(10, 4) As Variant
Dim iStart As Integer
Dim iEnd As Integer

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd
'*************************************************************************************


Private Sub cmbDiv_LostFocus()
   On Error Resume Next
   cmbDiv = CheckLen(cmbDiv, 2)
   If Trim(cmbDiv) <> "" And Not bValidElement(cmbDiv) Then
      cmbDiv = ""
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
   GetOptions
   sCurrForm = Caption
   If Trim(txtLvl) = "" Then txtLvl = "9"
   'optPrn.Picture = Resources.imgPrn.Picture
   'optDis.Picture = Resources.imgDis.Picture
   bOnLoad = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   
   On Error Resume Next
   Set rdoPer = Nothing
   Set diaGLp12a_New = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub cmbPer1_LostFocus()
   On Error Resume Next
   If Not bValidElement(cmbPer1) Then
      cmbPer1 = ""
      lblStart = ""
   End If
End Sub

Private Sub cmbPer1_Click()
   GetPeriodDate
End Sub

Private Sub cmbPer2_LostFocus()
   On Error Resume Next
   If Not bValidElement(cmbPer2) Then
      cmbPer2 = ""
      lblEnd = ""
   End If
End Sub

Private Sub cmbPer2_Click()
   GetPeriodDate
End Sub

Private Sub optPrn_Click()
   Dim sMessage As String
   On Error Resume Next
   
   If Trim(cmbYer) <> "" Then
      If Trim(cmbPer1) <> "" And Trim(cmbPer2) <> "" Then
         PrintReport
         Exit Sub
      Else
         sMessage = "Enter Valid Period Values."
      End If
   Else
      sMessage = "Enter A Valid Fiscal Year."
   End If
   MsgBox sMessage, vbInformation, Caption
End Sub

Private Sub optDis_Click()
   Dim sMessage As String
   On Error Resume Next
   
   If Trim(cmbYer) <> "" Then
      If Trim(cmbPer1) <> "" And Trim(cmbPer2) <> "" Then
         PrintReport
         Exit Sub
      Else
         sMessage = "Enter Valid Period Values."
      End If
   Else
      sMessage = "Enter A Valid Fiscal Year."
   End If
   MsgBox sMessage, vbInformation, Caption
End Sub

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



Private Sub txtLvl_LostFocus()
   If Trim(txtLvl) = "" Or Val(txtLvl) > 9 Or Val(txtLvl) < 1 Then txtLvl = 9
End Sub

Private Sub cmbYer_LostFocus()
   On Error Resume Next
   If Val(cmbYer) < 32000 Then
      cmbYer = CInt(Val(cmbYer))
      If bValidElement(cmbYer) Then
         LoadPeriods
      Else
         ClearPer False
         cmbYer = ""
      End If
   Else
      cmbYer = ""
      ClearPer False
   End If
End Sub

Private Sub cmbYer_Click()
   LoadPeriods
End Sub


Private Sub FillCombo() 'Gets Years; Fills Combo
   Dim rdoYrs As ADODB.Recordset
   Dim i As Integer
   On Error GoTo DiaErr1
   
   sSql = "SELECT FYYEAR FROM GlfyTable WHERE FYPERIODS > 0"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoYrs)
   If bSqlRows Then
      With rdoYrs
         cmbYer.Clear
         While Not .EOF
            AddComboStr cmbYer.hWnd, "" & !FYYEAR
            .MoveNext
         Wend
      End With
   End If
   
   'If bGoodYear(Val(Format(Now, "yyyy"))) Then
   '    cmbYer = Format(Now, "yyyy")
   'End If
   
   
   'Replace With Valid Element (Below)
   For i = 0 To cmbYer.ListCount - 1
      If cmbYer.List(i) = Format(Now, "yyyy") Then
         cmbYer.ListIndex = i
      End If
   Next
   
   If Trim(cmbYer) = "" And cmbYer.ListCount > 0 Then cmbYer.ListIndex = 0
   
   If bDivisionAccounts(iStart, iEnd) Then
      FillDivisions Me
   Else
      cmbDiv.enabled = False
   End If
   
   
   Set rdoYrs = Nothing
   Exit Sub
DiaErr1:
   sProcName = "FillCombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub ClearPer(bBool As Boolean)
   On Error GoTo DiaErr1
   cmbPer1.Clear
   cmbPer2.Clear
   cmbPer1.enabled = bBool
   cmbPer2.enabled = bBool
   cmbPer1 = ""
   cmbPer2 = ""
   cmbPer1.SelLength = 0
   cmbPer2.SelLength = 0
   lblStart = ""
   lblEnd = ""
   Exit Sub
   
DiaErr1:
   sProcName = "clearper"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Sub GetPeriodDate()
   Dim iStart As Integer
   Dim iEnd As Integer
   
   On Error GoTo DiaErr1
   iStart = Val(cmbPer1) - 1
   If Val(cmbPer2) > Val(cmbPer1) Then
      iEnd = 13 + Val(cmbPer2) - 1
   Else
      iEnd = 13 + Val(cmbPer1) - 1
   End If
   
   With rdoPer
      'lblStart = Format(.rdoColumns(Val(cmbPer1) - 1), "mm/dd/yy")
      'lblEnd = .rdoColumns(13 + Val(cmbPer2) - 1)
      lblStart = Format(.Fields(iStart), "mm/dd/yy")
      lblEnd = Format(.Fields(iEnd), "mm/dd/yy")
   End With
   
   Exit Sub
   
DiaErr1:
   sProcName = "GetPeriodDate"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Sub LoadPeriods()
   Dim i As Integer
   
   On Error GoTo DiaErr1
   ClearPer True
   sSql = "SELECT FYPERSTART1, FYPERSTART2, FYPERSTART3, FYPERSTART4, FYPERSTART5," _
          & " FYPERSTART6, FYPERSTART7, FYPERSTART8, FYPERSTART9,  FYPERSTART10, " _
          & " FYPERSTART11, FYPERSTART12, FYPERSTART13, FYPEREND1, FYPEREND2, " _
          & " FYPEREND3, FYPEREND4, FYPEREND5, FYPEREND6, FYPEREND7, FYPEREND8," _
          & " FYPEREND9 , FYPEREND10, FYPEREND11, FYPEREND12, FYPEREND13, FYPERIODS, FYSTART " _
          & " From GlfyTable Where (FYYEAR = " & cmbYer & ")"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoPer)
   
   If bSqlRows And Val(rdoPer!FYPERIODS) <> 0 Then
      With rdoPer
         lblYerStart = Format(!fyStart, "mm/dd/yy")
         For i = 1 To Val(!FYPERIODS)
            AddComboNum cmbPer1.hWnd, CLng(i)
            AddComboNum cmbPer2.hWnd, CLng(i)
            If CDate(.Fields(12 + i)) > CDate(Now) And CDate(Now) > CDate(.Fields(i - 1)) Then
               cmbPer1.ListIndex = i - 1
               cmbPer2.ListIndex = i - 1
            End If
         Next
      End With
      If Trim(cmbPer1) = "" Then
         cmbPer1.ListIndex = 0
         cmbPer2.ListIndex = Val("" & rdoPer!FYPERIODS) - 1
      End If
   Else
      ClearPer False
   End If
   
   Exit Sub
   
DiaErr1:
   sProcName = "LoadPeriods"
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
   Dim sDiv As String
   
   On Error GoTo DiaErr1
   
   
   sCustomReport = GetCustomReport("fingl12_new.rpt")
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   
   cCRViewer.SetReportTitle = "fingl12_new1.rpt"
   cCRViewer.ShowGroupTree False
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Title1"
   aFormulaName.Add "Title2"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "nDetailLevel"
   'aFormulaName.Add "Division"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Level " & txtLvl _
                        & " Income Statement For Year Beginning " & lblYerStart & "'")
   aFormulaValue.Add CStr("'Period Beginning:  " _
                        & lblStart & " And Ending:  " & lblEnd & "'")
   aFormulaValue.Add CStr("'Requested By: " & sInitials & "'")
                        
   aFormulaValue.Add CInt(Val(txtLvl))
   
   If Trim(cmbDiv) = "" Then
      sDiv = "ALL"
   Else
      sDiv = cmbDiv
   End If
   'aFormulaValue.Add CStr(sDiv)
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   'sSql = "({RptProFormaIncStatement;1.SUMBUDCURBAL} <> 0) OR ({RptProFormaIncStatement;1.SUMBUDYTD} <> 0)"
   sSql = "not isNull({RptProFormaIncStatement;1.GLACCTREF})"

   cCRViewer.SetReportSelectionFormula (sSql)
   
   cCRViewer.CRViewerSize Me
   ' Set report parameter
   cCRViewer.SetDbTableConnection True
   ' report parameter
   aRptPara.Add CStr(lblStart)
   aRptPara.Add CStr(lblEnd)
   aRptPara.Add CStr(cmbYer)
   aRptPara.Add CStr(cmbPer1)
   aRptPara.Add CStr(cmbPer2)
   aRptPara.Add CStr(optIna)
   
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("Int")
   aRptParaType.Add CStr("Int")
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



Public Sub SaveOptions()
   Dim sOptions As String
   Dim i As Integer
   sOptions = optIna.Value
   SaveSetting "Esi2000", "EsiFina", Me.Name, Trim(sOptions)
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Public Sub GetOptions()
   Dim i As Integer
   Dim sOptions As String
   
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
      optIna.Value = Val(sOptions)
   Else
      optIna.Value = 0
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = "Default Printer"
   End If
   
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
