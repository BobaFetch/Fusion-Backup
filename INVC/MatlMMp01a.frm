VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form MatlMMp01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Part Activity With Cost"
   ClientHeight    =   3060
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3060
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   1080
      Width           =   3075
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "MatlMMp01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Leading Char Search  (*  In Front Is A Legal Wild Card)"
      Top             =   1080
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.CommandButton cmdFnd 
      Height          =   315
      Left            =   4560
      Picture         =   "MatlMMp01a.frx":07AE
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   1080
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.Frame z2 
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   2280
      Width           =   3015
      Begin VB.OptionButton optPln 
         Caption         =   "User Date"
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         ToolTipText     =   "Planned (User Entered Date)"
         Top             =   200
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optAct 
         Caption         =   "Actual Date"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Actual (System Date)"
         Top             =   200
         Width           =   1335
      End
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   3720
      TabIndex        =   3
      Tag             =   "4"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   9
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "MatlMMp01a.frx":0AF0
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "MatlMMp01a.frx":0C6E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6840
      Top             =   2280
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3060
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For ALL)"
      Height          =   285
      Index           =   6
      Left            =   5160
      TabIndex        =   21
      Top             =   1920
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dates By"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   20
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   4
      Left            =   2760
      TabIndex        =   19
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dates From"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   18
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
      Height          =   252
      Index           =   2
      Left            =   5040
      TabIndex        =   17
      Top             =   1080
      Width           =   852
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty On Hand"
      Height          =   252
      Index           =   17
      Left            =   5040
      TabIndex        =   16
      Top             =   1440
      Width           =   1092
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   15
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   14
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   13
      Top             =   1440
      Width           =   3075
   End
   Begin VB.Label lblLvl 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   6120
      TabIndex        =   12
      Top             =   1080
      Width           =   612
   End
   Begin VB.Label lblQoh 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   6120
      TabIndex        =   11
      Top             =   1440
      Width           =   972
   End
End
Attribute VB_Name = "MatlMMp01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'2/28/05 Changed date handling
'5/10/05 Removed Combo and added lookup
'12/15/05 Changed Reports to total INAQTY*INAMT
Option Explicit

Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter

Dim bOnLoad As Byte
Dim bGoodPart As Byte
Dim bShowPart As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If Len(cmbPrt) And bShowPart = 0 Then bGoodPart = GetPart() _
          Else lblDsc = "*** No Part Selected ***"

End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   txtPrt = ""
   
End Sub

Private Sub cmdFnd_Click()
   ViewParts.lblControl = "TXTPRT"
   ViewParts.txtPrt = txtPrt
   ViewParts.Show
   
End Sub

Private Sub cmdFnd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bShowPart = 1
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 907
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then FixDates
   bOnLoad = 0
   FillCombo
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sSql = "SELECT TOP 1 PARTREF,PARTNUM,PADESC,PALEVEL," _
          & "PAQOH FROM PartTable WHERE PARTREF= ? "
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Type = adChar
   AdoParameter.Size = 30
   
   AdoQry.Parameters.Append AdoParameter
   
   'Set RdoQry = RdoCon.CreateQuery("", sSql)
   'RdoQry.MaxRows = 1
   bOnLoad = 1
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   On Error Resume Next
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   Set MatlMMp01a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sPartNumber As String
   Dim sBegDate As String
   Dim sEndDate As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   On Error Resume Next
   If Not IsDate(txtBeg) Then
      sBegDate = "1995,01,01"
   Else
      sBegDate = Format(txtBeg, "yyyy,mm,dd")
   End If
   If Not IsDate(txtEnd) Then
      sEndDate = "2024,12,31"
   Else
      sEndDate = Format(txtEnd, "yyyy,mm,dd")
   End If
   
   sPartNumber = Compress(cmbPrt)
   On Error GoTo DiaErr1
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "FromDate"
    aFormulaName.Add "ToDate"
    aFormulaName.Add "RequestBy"
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'" & CStr(txtBeg) & "'")
    aFormulaValue.Add CStr("'" & CStr(txtEnd) & "'")
    aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   sSql = ""
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   If optAct.Value = True Then
      sCustomReport = GetCustomReport("admin01")
      cCRViewer.SetReportFileName sCustomReport, sReportPath
      cCRViewer.SetReportTitle = sCustomReport
     sSql = "{PartTable.PARTREF}='" & sPartNumber & "' " _
             & "and {InvaTable.INADATE} in Date(" & sBegDate & ") " _
             & "to Date(" & sEndDate & ")"
   Else
      sCustomReport = GetCustomReport("admin01p")
      cCRViewer.SetReportFileName sCustomReport, sReportPath
      cCRViewer.SetReportTitle = sCustomReport
      sSql = "{PartTable.PARTREF}='" & sPartNumber & "' " _
             & "and {InvaTable.INPDATE} in Date(" & sBegDate & ") " _
             & "to Date(" & sEndDate & ") "
   End If
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.ShowGroupTree False
   cCRViewer.SetReportSelectionFormula (sSql)
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.OpenCrystalReportObject Me, aFormulaName

   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtBeg = "01/01/" & Right(txtEnd, 4)
   
End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 3) = "***" Then lblDsc.ForeColor = ES_RED _
           Else lblDsc.ForeColor = vbBlack
   
End Sub

Private Sub optDis_Click()
   If lblDsc.ForeColor = ES_RED Then MsgBox "Please Enter Or Select A Valid Part Number.", _
                         vbInformation, Caption Else PrintReport
   
End Sub

Private Sub optPrn_Click()
   If lblDsc.ForeColor = ES_RED Then MsgBox "Please Enter Or Select A Valid Part Number.", _
                         vbInformation, Caption Else PrintReport
   
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDateEx(txtBeg)
   
End Sub

Private Sub txtEnd_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If txtEnd <> "ALL" Then txtEnd = CheckDateEx(txtEnd)
   
End Sub

Private Sub FillCombo()
   sSql = "Qry_FillSortedParts"
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then
      cmbPrt = cmbPrt.List(0)
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "FillCombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Function GetPart()
   Dim RdoGet As ADODB.Recordset
   On Error GoTo DiaErr1
   'cmbPrt = Compress(cmbPrt)
   'RdoQry(0) = Compress(txtPrt)
    AdoQry.Parameters(0).Value = Compress(cmbPrt)
   bSqlRows = clsADOCon.GetQuerySet(RdoGet, AdoQry, ES_KEYSET)
   If bSqlRows Then
      With RdoGet
         GetPart = True
         txtPrt = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         lblLvl = "" & Format(!PALEVEL, "0")
         lblQoh = "" & Format(!PAQOH, ES_QuantityDataFormat)
         ClearResultSet RdoGet
      End With
   Else
      GetPart = False
      lblDsc = "*** Invalid Part Number ***"
      lblLvl = ""
      lblQoh = "0.000"
   End If
   Set RdoGet = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub FixDates()
   On Error Resume Next
   sSql = "UPDATE InvaTable SET INPDATE=INADATE WHERE INPDATE IS NULL"
   clsADOCon.ExecuteSQL sSql
End Sub

Private Sub txtPrt_GotFocus()
   bShowPart = 0
   If Len(Trim(txtPrt)) Then bGoodPart = GetPart()
   
End Sub

Private Sub cmbPrt_GotFocus()
   bShowPart = 0
   If Len(Trim(cmbPrt)) Then bGoodPart = GetPart()
   
End Sub


Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "TXTPRT"
      ViewParts.txtPrt = txtPrt
      ViewParts.Show
   End If
   
End Sub

Private Sub txtPrt_LostFocus()
   txtPrt = CheckLen(txtPrt, 30)
   If Len(txtPrt) And bShowPart = 0 Then bGoodPart = GetPart() _
          Else lblDsc = "*** No Part Selected ***"
   
End Sub

Private Sub cmbPrt_Change()
'   cmbPrt = CheckLen(cmbPrt, 30)
'   If Len(cmbPrt) And bShowPart = 0 Then bGoodPart = GetPart() _
'          Else lblDsc = "*** No Part Selected ***"
   
End Sub

Private Sub cmbPrt_Click()
'   cmbPrt = CheckLen(cmbPrt, 30)
'   If Len(cmbPrt) And bShowPart = 0 Then bGoodPart = GetPart() _
'          Else lblDsc = "*** No Part Selected ***"
   
End Sub

