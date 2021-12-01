VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form CapaCPp10a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Work Center Load Analysis"
   ClientHeight    =   3060
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3060
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CapaCPp10a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1680
      Width           =   1250
   End
   Begin VB.ComboBox cboShop 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Select From List"
      Top             =   960
      Width           =   1815
   End
   Begin VB.ComboBox cboWorkCenter 
      Height          =   288
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select Work Center, Leading Characters, Or Blank"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   3
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "CapaCPp10a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "CapaCPp10a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   3000
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
      Caption         =   "Cut Off Date"
      Height          =   288
      Index           =   4
      Left            =   240
      TabIndex        =   11
      Top             =   1680
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Select From List)"
      Height          =   288
      Index           =   3
      Left            =   3960
      TabIndex        =   10
      Top             =   960
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Centers"
      Height          =   288
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shop"
      Height          =   288
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   2
      Left            =   3960
      TabIndex        =   7
      Top             =   1320
      Width           =   1428
   End
End
Attribute VB_Name = "CapaCPp10a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'8/22/06 New
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FillCombo()
   '   Dim RdoCmb As ADODB.recordset
   '   On Error GoTo DiaErr1
   '    sSql = "SELECT DISTINCT CUREF,CUNICKNAME,SOCUST FROM " _
   '        & "CustTable,SohdTable WHERE CUREF=SOCUST"
   '    bsqlrows = clsadocon.getdataset(ssql, RdoCmb, ES_FORWARD)
   '        If bSqlRows Then
   '            With RdoCmb
   '                Do Until .EOF
   '                    AddComboStr cmbCst.hWnd, "" & Trim(!CUNICKNAME)
   '                    .MoveNext
   '                Loop
   '                .Cancel
   '            End With
   '        Else
   '            lblNme = "*** No Customers With SO's Found ***"
   '        End If
   '    Set RdoCmb = Nothing
   '    cmbCst = "ALL"
   '    GetCustomer
   '   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub cboShop_Click()
   FillWorkCenters
   
End Sub


Private Sub cboShop_LostFocus()
   If cboShop = "" Then
      If cboShop.ListCount > 0 Then cboShop = cboShop.List(0)
   End If
   FillWorkCenters
   
End Sub


Private Sub cboWorkCenter_LostFocus()
   If cboWorkCenter = "" Then cboWorkCenter = "ALL"
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4230
      MouseCursor 0
      cmdHlp = False
   End If
End Sub



Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillShops
'      cboWorkCenter.AddItem "ALL"
'      cboWorkCenter = "ALL"
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub FillShops()
   Dim wc As New ClassWorkCenter
   wc.PopulateShopCombo cboShop, cboWorkCenter
End Sub

Private Sub FillWorkCenters()
   Dim wc As New ClassWorkCenter
   wc.PoulateWorkCenterCombo cboShop, cboWorkCenter
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
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
   Set CapaCPp10a = Nothing
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   'b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtEnd = Format(Now + 30, "mm/dd/yyyy")
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   On Error Resume Next
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   
End Sub


Private Sub optDis_Click()
   'CapaCPp10b.lblThrough = Format(txtEnd, "mm/dd/yy")
   'CapaCPp10b.Shop = cboShop
   'CapaCPp10b.Center = cboWorkCenter
   'CapaCPp10b.Show
   PrintReport
End Sub


Private Sub optPrn_Click()
   'CapaCPp10b.lblThrough = txtEnd
   'CapaCPp10b.Shop = cboShop
   'CapaCPp10b.Center = cboWorkCenter
   'CapaCPp10b.Show
   PrintReport
End Sub


Private Sub txtend_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtEnd_LostFocus()
   txtEnd = CheckDateEx(txtEnd)
   If Format(txtEnd, "mm/dd/yy") < Format(Now, "mm/dd/yy") Then
      Beep
      txtEnd = Format(Now + 30, "mm/dd/yyyy")
   End If
   
End Sub


Private Sub PrintReport()
   Dim sShop As String
   Dim sCenter As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   Dim sDateStart As String
   sDateStart = Format(Now, "yyyy,mm,dd")
   Dim sDateEnd As String
   sDateEnd = Format(txtEnd, "yyyy,mm,dd")
   
   MouseCursor 13
   
   On Error GoTo DiaErr1
   If cboShop <> "ALL" Then sShop = Compress(cboShop)
   If cboWorkCenter <> "ALL" Then sCenter = Compress(cboWorkCenter)
   
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "Includes"
    aFormulaName.Add "CutOffDate"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
    aFormulaValue.Add CStr("'Shop: " & CStr(cboShop) & " Work Center(s):" & CStr(cboWorkCenter) & "  CutOff Date: " & CStr(txtEnd) & "'")
    aFormulaValue.Add CStr("'" & txtEnd & "'")
    
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("prdca10")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   sSql = "{WcntTable.WCNREF} LIKE '" & sCenter & "*' AND " _
          & "{WcntTable.WCNSHOP} = '" & sShop & "' AND {WcntTable.WCNSERVICE} = 0 "
'          & " AND {RnOpTable.OPCOMPLETE}=0 " _
 '         & " AND {RnopTable.OPSCHEDDATE} <= " & CrystalDate(sDateEnd)
      '    & " AND {WcclTable.WCCDATE} >= " & CrystalDate(sDateStart) & " AND {WcclTable.WCCDATE} <= " & CrystalDate(sDateEnd)
          
          


'   sSql = sSql & " AND {RunsTable.RUNSTATUS} <> 'CA' "
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.SetReportSelectionFormula sSql
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


