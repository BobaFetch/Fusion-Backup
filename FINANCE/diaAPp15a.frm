VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaAPp15a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Computer Check Summary (Report)"
   ClientHeight    =   2775
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6510
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2775
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   3840
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Vendors With Checks"
      Top             =   840
      Width           =   1680
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5280
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5280
      TabIndex        =   4
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaAPp15a.frx":0000
      PictureDn       =   "diaAPp15a.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3600
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2775
      FormDesignWidth =   6510
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
      PictureUp       =   "diaAPp15a.frx":028C
      PictureDn       =   "diaAPp15a.frx":03D2
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   14
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   2
      Left            =   3480
      TabIndex        =   12
      Top             =   840
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Posting Dates From"
      Height          =   405
      Index           =   4
      Left            =   240
      TabIndex        =   11
      Top             =   1500
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   3
      Left            =   3000
      TabIndex        =   10
      Top             =   1680
      Width           =   915
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   1425
   End
End
Attribute VB_Name = "diaAPp15a"
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
' diaAPp15a - Computer Check Summary
'
' Notes:
'
' Created: 08/07/01 (nth)
' Revisions:
'   10/13/03 (nth) Added ablility to see void checks and report rerouting.
'   08/16/04 (nth) Added printer to getoptions and saveoptions.
'
'*************************************************************************************

Option Explicit
Dim bOnLoad As Byte

Dim sBegDate As String
Dim sEnddate As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Public Sub TestCheckDates()
   Dim RdoTst As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT MIN(CHKPOSTDATE) FROM ChksTable WHERE " _
          & "CHKPRINTED=1 AND CHKVOID=0"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTst, ES_FORWARD)
   If bSqlRows Then
      With RdoTst
         If Not IsNull(.Fields(0)) Then
            sBegDate = Format(.Fields(0), "mm/dd/yy")
         Else
            sBegDate = Format(Now - 240, "mm/dd/yy")
         End If
         .Cancel
      End With
   Else
   End If
   Set RdoTst = Nothing
   sSql = "SELECT MAX(CHKPOSTDATE) FROM ChksTable WHERE " _
          & "CHKPRINTED=1 AND CHKVOID=0"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTst, ES_FORWARD)
   If bSqlRows Then
      With RdoTst
         If Not IsNull(.Fields(0)) Then
            sEnddate = Format(.Fields(0), "mm/dd/yy")
         Else
            sEnddate = Format(Now - 240, "mm/dd/yy")
         End If
         .Cancel
      End With
   Else
      sEnddate = txtEnd
   End If
   Set RdoTst = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "testcheckda"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cmbVnd_Click()
   GetCheckVendor
End Sub

Private Sub cmbVnd_LostFocus()
   cmbVnd = CheckLen(cmbVnd, 10)
   If Len(cmbVnd) = 0 Then cmbVnd = "ALL"
   GetCheckVendor
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Reports"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub FillCombo()
   On Error GoTo DiaErr1
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   AddComboStr cmbVnd.hWnd, "ALL"
   sSql = "SELECT DISTINCT CHKVENDOR,VEREF,VENICKNAME FROM " _
          & "ChksTable,VndrTable WHERE CHKVENDOR=VEREF " _
          & "ORDER BY CHKVENDOR"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            AddComboStr cmbVnd.hWnd, "" & Trim(!VENICKNAME)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoCmb = Nothing
   cmbVnd = "ALL"
   If cmbVnd.ListCount > 0 Then GetCheckVendor
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub Form_Activate()
   On Error Resume Next
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
   txtBeg = Format(Now, "mm/01/yy")
   txtEnd = Format(Now, "mm/dd/yy")
   'optPrn.Picture = Resources.imgPrn.Picture
   'optDis.Picture = Resources.imgDis.Picture
   GetOptions
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveOptions
   FormUnload
   Set diaAPp15a = Nothing
End Sub

Private Sub PrintReport()
   Dim sVendor As String
   Dim sBeg As String
   Dim sEnd As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   On Error GoTo DiaErr1
   
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Includes" & CStr(cmbVnd & "... " _
                        & "From " & txtBeg & " Through " & txtEnd) & "'")
    aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
    
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    sCustomReport = GetCustomReport("finch02.rpt")
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   
   If Trim(cmbVnd) = "ALL" Or cmbVnd = "" Then
      sVendor = ""
      sSql = ""
   Else
      sVendor = Compress(cmbVnd)
      sSql = "{ChksTable.CHKVENDOR} = '" & cmbVnd & "' AND"
   End If
   
   sSql = sSql & "{ChksTable.CHKPOSTDATE} >=#" & txtBeg _
          & "# AND {ChksTable.CHKPOSTDATE} <=#" _
          & txtEnd & "# AND {ChksTable.CHKPRINTED} = 1"
   
    cCRViewer.CRViewerSize Me
    ' Set report parameter
    cCRViewer.SetDbTableConnection
    ' print the copies
    cCRViewer.SetReportSelectionFormula sSql
    
   sSql = "{ChksTable.CHKVOIDDATE} >=#" & txtBeg _
          & "# AND {ChksTable.CHKVOIDDATE} <=#" _
          & txtEnd & "# AND {ChksTable.CHKVOID} = 1"
          
   
   ' set the sub sql variable pass the sub report name
   cCRViewer.SetSubRptSelFormula "sr_VoidedChecks", sSql
    
    
    cCRViewer.OpenCrystalReportObject Me, aFormulaName
    cCRViewer.ShowGroupTree False
    
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

Private Sub PrintReport1()
   Dim sVendor As String
   Dim sBeg As String
   Dim sEnd As String
   Dim sCustomReport As String
   
   MouseCursor 13
   On Error GoTo DiaErr1
   
   'SetMdiReportsize MdiSect
   
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "Includes='Includes " & cmbVnd & "... " _
                        & "From " & txtBeg & " Through " & txtEnd & "'"
   MdiSect.crw.Formulas(2) = "RequestBy='Requested By: " & sInitials & "'"
   
   sCustomReport = GetCustomReport("finch02.rpt")
   
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   
   If Trim(cmbVnd) = "ALL" Or cmbVnd = "" Then
      sVendor = ""
      sSql = ""
   Else
      sVendor = Compress(cmbVnd)
      sSql = "{ChksTable.CHKVENDOR} = '" & cmbVnd & "' AND"
   End If
   
   sSql = sSql & "{ChksTable.CHKPOSTDATE} >=#" & txtBeg _
          & "# AND {ChksTable.CHKPOSTDATE} <=#" _
          & txtEnd & "# AND {ChksTable.CHKPRINTED} = 1"
   MdiSect.crw.SelectionFormula = sSql
   
   'SetCrystalAction Me
   
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "PrintReport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = "Default Printer"
   End If
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub GetCheckVendor()
   Dim RdoChk As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT VEREF,VEBNAME FROM VndrTable WHERE " _
          & "VEREF='" & Compress(cmbVnd) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk, ES_FORWARD)
   If bSqlRows Then
      With RdoChk
         lblNme = "" & Trim(!VEBNAME)
         .Cancel
      End With
   Else
      lblNme = "Multiple Vendors Selected"
   End If
   Set RdoChk = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getcheckven"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) > 0 Then
      txtBeg = CheckDate(txtBeg)
   Else
      txtBeg = "ALL"
   End If
End Sub

Private Sub txtend_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) > 0 Then
      txtEnd = CheckDate(txtEnd)
   Else
      txtEnd = "ALL"
   End If
End Sub
