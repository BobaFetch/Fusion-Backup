VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaARp03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer Invoice Register"
   ClientHeight    =   4305
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6570
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4305
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optCD 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1680
      TabIndex        =   21
      Top             =   2760
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   1680
      TabIndex        =   17
      Top             =   3240
      Width           =   3375
      Begin VB.OptionButton optSortBy 
         Caption         =   "Date Only"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optSortBy 
         Caption         =   "Customer/Date"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CheckBox optCA 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   2280
      Width           =   735
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   3720
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Contains Customers With Invoices"
      Top             =   840
      Width           =   1555
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5400
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5400
      TabIndex        =   5
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaARp03a.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "diaARp03a.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   4
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
      PictureUp       =   "diaARp03a.frx":0308
      PictureDn       =   "diaARp03a.frx":044E
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6000
      Top             =   3360
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4305
      FormDesignWidth =   6570
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
      PictureUp       =   "diaARp03a.frx":0594
      PictureDn       =   "diaARp03a.frx":06DA
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include DM/CM"
      Height          =   405
      Index           =   6
      Left            =   240
      TabIndex        =   22
      Top             =   2760
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sort Report by"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   20
      Top             =   3480
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Advance Payments"
      Height          =   405
      Index           =   2
      Left            =   240
      TabIndex        =   16
      Top             =   2280
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   0
      Left            =   3480
      TabIndex        =   15
      Top             =   840
      Width           =   1545
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   14
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   3
      Left            =   3000
      TabIndex        =   12
      Top             =   1680
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   11
      Top             =   1680
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer "
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   840
      Width           =   1425
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   9
      Top             =   1200
      Width           =   2775
   End
End
Attribute VB_Name = "diaARp03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'************************************************************************************
' Form: diaPar03 - Customer Invoice Register
'
' Notes: Prints or displays a customer invoice register.
'
' Created: (cjs)
' Modified:
'   06/25/01 (nth) Fix the customer combo so it now loads customers
'   05/09/03 (nth) Made report like MCS and modfied crystal selection formula
'
'************************************************************************************

Option Explicit
Dim bOnLoad As Byte
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'************************************************************************************

Private Sub cmbCst_Click()
   FindCustomer Me, cmbCst
End Sub

Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If Len(cmbCst) Then
      FindCustomer Me, cmbCst
   Else
      cmbCst = "ALL"
   End If
   If cmbCst = "ALL" Then lblNme = "All Customers Selected."
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
   Dim rdoCst As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT CUREF,CUNICKNAME,INVCUST FROM " _
          & "CustTable,CihdTable WHERE CUREF=INVCUST "
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCst)
   If bSqlRows Then
      With rdoCst
         Do Until .EOF
            AddComboStr cmbCst.hWnd, "" & Trim(!CUNICKNAME)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set rdoCst = Nothing
   If cmbCst.ListCount > 0 Then
      cmbCst = cUR.CurrentCustomer
      FindCustomer Me, cmbCst
   End If
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
   GetOptions
   txtEnd = Format(GetServerDateTime, "mm/dd/yy")
   txtBeg = Format(txtEnd, "mm/01/yy")
   bOnLoad = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Len(cmbCst) Then
      cUR.CurrentCustomer = cmbCst
      SaveCurrentSelections
   End If
   FormUnload
   Set diaARp03a = Nothing
End Sub

Private Sub PrintReport()
    Dim sCust As String
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
   
   sBeg = Format(txtBeg, "mm/dd/yyyy")
   sEnd = Format(txtEnd, "mm/dd/yyyy")
   
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "ShowDCM"
    
    aFormulaName.Add "ShowGroup1Subtotal"
    aFormulaName.Add "Group1Field"
    aFormulaName.Add "Group2Field"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'" & CStr(cmbCst _
                        & " From " & txtBeg & " To " & txtEnd) & "... '")
    aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
    
   aFormulaValue.Add optCD

    If Me.optSortBy(1).Value = True Then aFormulaValue.Add 0 Else aFormulaValue.Add 1
    If Me.optSortBy(1).Value = True Then
        aFormulaValue.Add CStr("{CihdTable.INVDATE}")
        aFormulaValue.Add CStr("{CihdTable.INVNO}")
    Else
        aFormulaValue.Add CStr("{CustTable.CUNICKNAME}")
        aFormulaValue.Add CStr("{CihdTable.INVNO}")
    End If
        
    
   
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    sCustomReport = GetCustomReport("finar03.rpt")
        
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = sCustomReport
   
   
   sSql = ""
   
   sSql = cCRViewer.GetReportSelectionFormula
   
   If (sSql <> "") Then
      sSql = sSql & " AND "
   End If
   
   
   If cmbCst <> "ALL" Then
      sCust = Compress(cmbCst)
      sSql = sSql & "{CihdTable.INVCUST} = '" & sCust & "' AND "
   End If
   
   sSql = sSql & "{JrhdTable.MJTYPE} = 'SJ' AND {CihdTable.INVDATE} >= #" _
          & sBeg & "#" & " AND {CihdTable.INVDATE} <= #" & sEnd _
          & "# AND {CihdTable.INVCANCELED}=0"
   
   If optCA = vbUnchecked Then
      sSql = sSql & " AND {CihdTable.INVTYPE} <> 'CA'"
      aFormulaName.Add "Title1"
      aFormulaValue.Add CStr("'*** Excludes Cash Advances ***'")
   End If
   
   If optCD = vbUnchecked Then
      sSql = sSql & " AND ({CihdTable.INVTYPE} <> 'CM') AND ({CihdTable.INVTYPE} <> 'DM')"
      aFormulaName.Add "Title2"
      aFormulaValue.Add CStr("'*** Excludes Debit and Credit Memo***'")
   End If
   
   'not need
   'sSql = sSql & " and {CihdTable.INVCANCELED} = 0.00"
    
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   cCRViewer.CRViewerSize Me
   ' Set report parameter
   cCRViewer.SetDbTableConnection
   ' print the copies
   cCRViewer.SetReportSelectionFormula sSql
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

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub


Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
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

Public Sub SaveOptions()
   Dim sbuf As String
   sbuf = optCA.Value
   sbuf = sbuf & optCD.Value
   
   If optSortBy(1).Value = True Then sbuf = sbuf & "1" Else sbuf = sbuf & "0"
   SaveSetting "Esi2000", "EsiFina", Me.Name, sbuf
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Public Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(sOptions) Then
      optCA.Value = Mid(sOptions, 1, 1)
      optCD.Value = Mid(sOptions, 1, 1)
      If Len(sOptions) > 1 Then optSortBy(Val(Mid(sOptions, 2, 1))).Value = 1
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
End Sub
