VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form BookBKp02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bookings By Part Number "
   ClientHeight    =   4035
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4035
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "BookBKp02a.frx":0000
      Height          =   315
      Left            =   5160
      Picture         =   "BookBKp02a.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   840
      Width           =   350
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "BookBKp02a.frx":0684
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Tag             =   "3"
      Top             =   860
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Contains Customers With Sales Orders"
      Top             =   1560
      Width           =   1555
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      MaskColor       =   &H8000000F&
      TabIndex        =   8
      Top             =   3480
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      MaskColor       =   &H8000000F&
      TabIndex        =   7
      Top             =   3240
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.ComboBox cmbDiv 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   5
      Tag             =   "8"
      ToolTipText     =   "Select Division From List"
      Top             =   2640
      Width           =   860
   End
   Begin VB.ComboBox cmbReg 
      Height          =   288
      Left            =   4200
      Sorted          =   -1  'True
      TabIndex        =   6
      Tag             =   "8"
      ToolTipText     =   "Select Region From List"
      Top             =   2640
      Width           =   780
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4215
      TabIndex        =   4
      Tag             =   "4"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Tag             =   "4"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   12
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
         Picture         =   "BookBKp02a.frx":0E32
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "BookBKp02a.frx":0FB0
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   3720
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4035
      FormDesignWidth =   7260
   End
   Begin VB.Label lblCUName 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   2160
      TabIndex        =   28
      Top             =   1920
      Width           =   3132
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer(s)"
      Height          =   285
      Index           =   12
      Left            =   240
      TabIndex        =   26
      Top             =   1560
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   11
      Left            =   5640
      TabIndex        =   25
      Tag             =   " "
      Top             =   1560
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   192
      Index           =   8
      Left            =   240
      TabIndex        =   24
      Top             =   3000
      Width           =   1908
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   192
      Index           =   7
      Left            =   240
      TabIndex        =   23
      Top             =   3240
      Width           =   2028
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   192
      Index           =   6
      Left            =   240
      TabIndex        =   22
      Top             =   3480
      Width           =   2028
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   5
      Left            =   5640
      TabIndex        =   21
      Tag             =   " "
      Top             =   860
      Width           =   1905
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2160
      TabIndex        =   20
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   19
      Top             =   860
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   10
      Left            =   5640
      TabIndex        =   18
      Tag             =   " "
      Top             =   2280
      Width           =   1788
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   9
      Left            =   5640
      TabIndex        =   17
      Tag             =   " "
      Top             =   2640
      Width           =   1548
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
      Height          =   288
      Index           =   4
      Left            =   3360
      TabIndex        =   16
      Tag             =   " "
      Top             =   2640
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   3
      Left            =   3480
      TabIndex        =   15
      Top             =   2280
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Booking Start Date"
      Height          =   288
      Index           =   2
      Left            =   240
      TabIndex        =   14
      Top             =   2280
      Width           =   1548
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      Height          =   288
      Index           =   1
      Left            =   240
      TabIndex        =   13
      Tag             =   " "
      Top             =   2652
      Width           =   1428
   End
End
Attribute VB_Name = "BookBKp02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'2/25/05 Changed dates and Options
'4/27/05 Removed Part Combo
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCst_Click()
   GetThisCustomer
   
End Sub

Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If Trim(cmbCst) = "" Then cmbCst = "ALL"
   GetThisCustomer
   
End Sub

Private Sub cmbDiv_LostFocus()
   cmbDiv = CheckLen(cmbDiv, 4)
   If Len(cmbDiv) = 0 Then cmbDiv = "ALL"
   
End Sub

Private Sub cmbPrt_Click()
   GetPart
   
End Sub

Private Sub cmdFnd_Click()
   If txtPrt.Visible Then
      cmbPrt = txtPrt
      ViewParts.lblControl = "TXTPRT"
   Else
      ViewParts.lblControl = "CMBPRT"
   End If
   ViewParts.txtPrt = cmbPrt
   ViewParts.Show
   
End Sub


Private Sub txtPrt_Change()
   cmbPrt = txtPrt
End Sub

Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "TXTPRT"
      ViewParts.txtPrt = txtPrt
      ViewParts.Show
   End If
   
End Sub

Private Sub txtPrt_LostFocus()
   cmbPrt = txtPrt
   If Trim(txtPrt) = "" Or Trim(txtPrt) = "ALL" Then txtPrt = "ALL"
   GetPart
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If Trim(cmbPrt) = "" Or Trim(cmbPrt) = "ALL" Then cmbPrt = "ALL"
   GetPart
   
End Sub

Private Sub cmbReg_LostFocus()
   cmbReg = CheckLen(cmbReg, 3)
   If Len(cmbReg) = 0 Then cmbReg = "ALL"
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 907
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_GetCustomerSalesOrder"
   LoadComboBox cmbCst
   cmbCst = "ALL"
   GetThisCustomer
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmdVew_Click(Index As Integer)
   ViewParts.lblControl = "CMBPRT"
   ViewParts.txtPrt = cmbPrt
   ViewParts.Show
   
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillDivisions
      FillRegions
      If cmbDiv = "" Then cmbDiv = "ALL"
      If cmbReg = "" Then cmbReg = "ALL"
      FillCombo
      
      Dim bPartSearch As Boolean
      
      bPartSearch = GetPartSearchOption
      SetPartSearchOption (bPartSearch)
      
      If (Not bPartSearch) Then FillPartCombo cmbPrt
      
      FormatControls
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
'   FormatControls
   GetOptions
   bOnLoad = 1
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set BookBKp02a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sPartNumber As String
   Dim sCust As String
   Dim sDiv As String
   Dim sReg As String
   Dim sBegDate As String
   Dim sEndDate As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   If Trim(cmbPrt) = "" Then cmbPrt = "ALL"
   If Trim(cmbPrt) <> "ALL" Then sPartNumber = Compress(cmbPrt) _
           Else sPartNumber = ""
   
   If Trim(cmbCst) = "" Then cmbCst = "ALL"
   If Trim(cmbCst) <> "ALL" Then sCust = Compress(cmbCst) _
           Else sCust = ""
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
   If Trim(cmbDiv) = "ALL" Then sDiv = "" Else sDiv = cmbDiv
   If Trim(cmbReg) = "ALL" Then sReg = "" Else sReg = cmbReg
   
   MouseCursor 13
   On Error GoTo DiaErr1
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowDescription"
   aFormulaName.Add "ShowExDescription"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Part(s) " & CStr(cmbPrt & ", Division " _
                        & cmbDiv & ", Region " & cmbReg & " From " & txtBeg _
                        & " To " & txtEnd) & "...'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add optDsc.Value
   aFormulaValue.Add optExt.Value
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("slebk02")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

   sSql = "{SoitTable.ITPART} Like '" & sPartNumber & "*' " _
          & "AND {SohdTable.SODIVISION} LIKE '" & sDiv & "*' " _
          & "AND {SohdTable.SOREGION} LIKE '" & sReg & "*' " _
          & "AND {SohdTable.SOCUST} LIKE '" & sCust & "*' " _
          & "AND {SoitTable.ITBOOKDATE} In Date(" & sBegDate & ") " _
          & "To Date(" & sEndDate & ")" _
          & " and {SoitTable.ITCANCELED} = 0.00"
   
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
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
   txtBeg = "" 'Format(ES_SYSDATE, "mm/01/yyyy")
   txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   cmbDiv.AddItem "ALL"
   cmbReg.AddItem "ALL"
   cmbPrt = "ALL"
   GetPart
   
   
End Sub

Private Sub SaveOptions()
   Dim sDiv As String * 4
   Dim sReg As String * 3
   Dim sOptions As String
   'Save by Menu Option
   sDiv = cmbDiv
   sReg = cmbReg
   sOptions = sDiv & sReg _
              & Trim(str(optDsc.Value)) & Trim(str(optExt.Value))
   SaveSetting "Esi2000", "EsiSale", "bk02a", sOptions
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiSale", "bk02a", Trim(sOptions))
   If Len(sOptions) > 0 Then
      cmbDiv = Mid(sOptions, 1, 4)
      cmbReg = Mid(sOptions, 5, 3)
      optDsc.Value = Val(Mid(sOptions, 8, 1))
      optExt.Value = Val(Mid(sOptions, 9, 1))
   End If
   txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtBeg = "" 'Left(txtEnd, 3) & "01" & Right(txtEnd, 5)
   
End Sub

Private Sub optDis_Click()
   PrintReport
   
End Sub

Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optExt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub


Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDateEx(txtBeg)
   
End Sub

Private Sub GetPart()
   Dim RdoPrt As ADODB.Recordset
   On Error GoTo DiaErr1
   
   sSql = "Qry_GetPartNumberBasics '" & Compress(cmbPrt) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
   If bSqlRows Then
      With RdoPrt
         cmbPrt = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         ClearResultSet RdoPrt
      End With
   Else
      lblDsc = "Range Of Parts Selected."
   End If
   Set RdoPrt = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub txtend_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If Trim(txtEnd) <> "ALL" Then txtEnd = CheckDateEx(txtEnd)
   
End Sub

Function SetPartSearchOption(bPartSearch As Boolean)
   
   If (bPartSearch = True) Then
      cmbPrt.Visible = False
      txtPrt.Visible = True
      cmdFnd.Visible = True
   Else
      cmbPrt.Visible = True
      txtPrt.Visible = False
      cmdFnd.Visible = False
   End If
End Function

