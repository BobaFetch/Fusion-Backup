VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form SaleSLp04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer Sales Order List"
   ClientHeight    =   3255
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7095
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3255
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLp04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4080
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6000
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6000
      TabIndex        =   13
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "SaleSLp04a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   490
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "SaleSLp04a.frx":0938
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   3
      Top             =   2295
      Width           =   735
   End
   Begin VB.CheckBox optCmt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   4
      Top             =   2580
      Width           =   735
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select From List Or Blank For All"
      Top             =   960
      Width           =   1555
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6480
      Top             =   2640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3255
      FormDesignWidth =   7095
   End
   Begin VB.Label lblNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   17
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   7
      Left            =   5400
      TabIndex        =   16
      Top             =   1680
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   6
      Left            =   5400
      TabIndex        =   15
      Top             =   960
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   12
      Top             =   2040
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   2295
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Comments"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   2580
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   2
      Left            =   3360
      TabIndex        =   9
      Top             =   1680
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Booking Start Date"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Nickname"
      Height          =   288
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   1788
   End
End
Attribute VB_Name = "SaleSLp04a"
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
Dim bOnLoad As Byte
Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   Dim sBeg As String * 8
   Dim sEnd As String * 8
   sBeg = txtBeg
   sEnd = txtEnd
   'Save by Menu Option
   sOptions = RTrim(optExt.Value) _
              & RTrim(optCmt.Value) _
              & sBeg & sEnd
   SaveSetting "Esi2000", "EsiSale", "sl04", Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   'Get By Menu Option
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiSale", "sl04", sOptions)
   If Len(sOptions) > 0 Then
      optExt.Value = Val(Left(sOptions, 1))
      optCmt.Value = Val(Mid(sOptions, 2, 1))
      '   txtBeg = Mid(sOptions, 3, 8)
      '  txtEnd = Mid(sOptions, 11, 8)
      
   End If
   txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtBeg = "01/01/" & Right(txtEnd, 4)
   
End Sub

Private Sub cmbCst_Click()
   FindCustomer Me, cmbCst, False
   
End Sub

Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If cmbCst = "" Then
      cmbCst = "ALL"
      lblNme = "All Customers"
   End If
   
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

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = 0
      cmbCst.AddItem "ALL"
      FillCustomers
      If cUR.CurrentCustomer <> "" Then
         cmbCst = cUR.CurrentCustomer
         FindCustomer Me, cmbCst, False
      Else
         cmbCst = "ALL"
      End If
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtBeg = Format(ES_SYSDATE, "mm/01/yyyy")
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
   On Error Resume Next
   If cmbCst <> "ALL" Then
      cUR.CurrentCustomer = cmbCst
      SaveCurrentSelections
   End If
   FormUnload
   Set SaleSLp04a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sCust As String
   Dim sBegDte As String
   Dim sEndDte As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   If cmbCst = "" Then cmbCst = "ALL"
   If cmbCst <> "ALL" Then sCust = Compress(cmbCst)
   
   If Not IsDate(txtBeg) Then
      sBegDte = "1995,01,01"
   Else
      sBegDte = Format(txtBeg, "yyyy,mm,dd")
   End If
   If Not IsDate(txtEnd) Then
      sEndDte = "2024,12,31"
   Else
      sEndDte = Format(txtEnd, "yyyy,mm,dd")
   End If
   
   On Error GoTo DiaErr1
  aFormulaName.Add "CompanyName"
  aFormulaName.Add "Includes"
  aFormulaName.Add "RequestBy"
  aFormulaName.Add "ShowExDescription"
  aFormulaName.Add "ShowComments"
  
  aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
  aFormulaValue.Add CStr("'Includes Customer(s) " & CStr(cmbCst & "... Booked " _
                        & "From " & txtBeg & " To " & txtEnd) & "'")
  aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
  aFormulaValue.Add optExt.Value
  aFormulaValue.Add optCmt.Value
  
  Set cCRViewer = New EsCrystalRptViewer
  cCRViewer.Init
  sCustomReport = GetCustomReport("sleco04")
  cCRViewer.SetReportFileName sCustomReport, sReportPath
  cCRViewer.SetReportTitle = sCustomReport
  cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{CustTable.CUREF} LIKE '" & sCust & "*' " _
          & "AND {SoitTable.ITBOOKDATE} in Date(" & sBegDte _
          & ") to Date(" & sEndDte & ")" _
          & " AND {SoitTable.ITCANCELED} = 0"
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

Private Sub optCmt_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub

Private Sub optCmt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optDis_Click()
   MouseCursor 13
   PrintReport
   
End Sub

Private Sub optExt_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub

Private Sub optExt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optPrn_Click()
   MouseCursor 13
   PrintReport
   
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDateEx(txtBeg)
   
End Sub

Private Sub txtend_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If Trim(txtEnd) <> "ALL" Then txtEnd = CheckDateEx(txtEnd)
   
End Sub
