VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form CapaCPp08a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Capacity And Load"
   ClientHeight    =   3525
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3525
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CapaCPp08a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optGrp 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1920
      TabIndex        =   4
      Top             =   2520
      Width           =   735
   End
   Begin VB.ComboBox cmbShp 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select From List, Leading Characters Or Blank For All"
      Top             =   960
      Width           =   1815
   End
   Begin VB.ComboBox cmbWcn 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select From List, Leading Characters Or Blank For All"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1680
      Width           =   1250
   End
   Begin VB.CheckBox optDet 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1920
      TabIndex        =   3
      Top             =   2280
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "CapaCPp08a.frx":07AE
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
         Picture         =   "CapaCPp08a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6840
      Top             =   2880
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3525
      FormDesignWidth =   7260
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   1920
      TabIndex        =   20
      Top             =   2880
      Visible         =   0   'False
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Detail"
      Height          =   285
      Index           =   9
      Left            =   240
      TabIndex        =   18
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Chart"
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   17
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   360
      Picture         =   "CapaCPp08a.frx":0AB6
      ToolTipText     =   "Chart Results"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   7
      Left            =   4080
      TabIndex        =   16
      Top             =   960
      Width           =   2265
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   6
      Left            =   4080
      TabIndex        =   15
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Past Due As Of"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   14
      ToolTipText     =   "Select Or Enter A New Work Center"
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shop(s)"
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   13
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center(s)"
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   12
      ToolTipText     =   "Select Or Enter A New Work Center"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Tag             =   " "
      ToolTipText     =   "The Day That Your Week Starts In The Company Setup"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Week Starts On"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Tag             =   " "
      ToolTipText     =   "The Day That Your Week Starts In The Company Setup"
      Top             =   480
      Width           =   2145
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   1785
   End
End
Attribute VB_Name = "CapaCPp08a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'New 11/13/03
Option Explicit
Dim bOnLoad As Byte
Dim iTotalCenters As Integer
Dim sPartNumber As String
Dim sWkStarts As String

Dim cHours(50) As Currency
Dim sCenters(500, 7) As String
'0 = Shop Ref
'1 = Shop Name
'2 = Shop Desc
'3 = Work Center Ref
'4 = Work Center Name
'5 = Work Center Desc

Dim dCapaDates(36, 2) As Date
' 0 = Start Date
' 1 = End Date
Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT SHPREF,SHPNUM " _
          & "FROM ShopTable,WcntTable " _
          & "Where (WCNSHOP = SHPREF) AND WCNSERVICE = 0 "
   LoadComboBox cmbShp
   If cmbShp.ListCount > 0 Then cmbShp = cmbShp.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbShp_Click()
   FillCenters
   
End Sub

Private Sub cmbShp_LostFocus()
   cmbShp = CheckLen(cmbShp, 10)
   If cmbShp = "" Then cmbShp = "ALL"
   FillCenters
   
End Sub

Private Sub cmbWcn_LostFocus()
   cmbWcn = CheckLen(cmbWcn, 10)
   If cmbWcn = "" Then cmbWcn = "ALL"
   
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
   MouseCursor 13
   If bOnLoad Then
      CreateTable1
      CreateTable2
      CreateTable3
      z1(0) = z1(0) & " " & GetStartDay()
      FillCombo
      bOnLoad = 0
   End If
   MDISect.lblBotPanel = Caption
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   GetOptions
   bOnLoad = 1
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   On Error Resume Next
   sSql = "UPDATE EsReportCapa15d SET RPTLOCKED=0"
   clsADOCon.ExecuteSql sSql
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set CapaCPp08a = Nothing
   
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
   
   MouseCursor 13
   
   On Error GoTo DiaErr1
   If cmbShp <> "ALL" Then sShop = Compress(cmbShp)
   If cmbWcn <> "ALL" Then sCenter = Compress(cmbWcn)
   
    
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "ShowComments"
    aFormulaName.Add "ShowGroup"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Shop(s)" & cmbShp & ", Work Centers(s) " _
                        & cmbWcn & " Late Starting " & txtBeg & "...'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
    aFormulaValue.Add optDet.Value
    aFormulaValue.Add optGrp.Value
    
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("prdca15")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   
'  ' MDISect.Crw.ReportFileName = sReportPath & sCustomReport
          
'   If optDet.value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(0) = "GROUPHDR.0.0;F;;;"
'      MDISect.Crw.SectionFormat(1) = "DETAIL.0.0;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(0) = "GROUPHDR.0.0;T;;;"
'      MDISect.Crw.SectionFormat(1) = "DETAIL.0.0;T;;;"
'   End If
'   If optGrp.value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(2) = "REPORTFTR.0.1;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(2) = "REPORTFTR.0.1;T;;;"
'   End If
   
   sSql = "{EsReportCapa15a.RPTSHOPREF} LIKE '" & sShop & "*' " _
          & "AND {EsReportCapa15a.RPTWCNREF} LIKE '" & sCenter & "*'" _
          & "AND {EsReportCapa15b.RPTRUNSTATUS} <> 'CA'"
         
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.SetDbTableConnection
   
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName

   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
'   MDISect.Crw.SelectionFormula = sSql
'   SetCrystalAction Me
   
   On Error Resume Next
   sSql = "UPDATE EsReportCapa15d SET RPTLOCKED=0"
   clsADOCon.ExecuteSql sSql
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sSql = "UPDATE EsReportCapa15d SET RPTLOCKED=0"
   clsADOCon.ExecuteSql sSql
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtBeg = Format(ES_SYSDATE, "mm/dd/yyyy")
   cmbWcn = "ALL"
   
End Sub

Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiProd", "ca15", optDet.Value
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "ca15", sOptions)
   If sOptions = "" Then optDet.Value = vbChecked _
                 Else optDet.Value = Val(sOptions)
   
End Sub

Private Sub Image1_Click()
   If optGrp.Value = vbChecked Then
      optGrp.Value = vbUnchecked
   Else
      optGrp.Value = vbChecked
   End If
   
End Sub

Private Sub optDet_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   'PrintReport
   BuildReport
   
End Sub

Private Sub optGrp_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optPrn_Click()
   'PrintReport
   
End Sub

'sCapaDates
' 0 = Start Date
' 1 = End Date
'
'sCenters
'0 = Shop Ref
'1 = Shop Name
'2 = Shop Desc
'3 = Work Center Ref
'4 = Work Center Name
'5 = Work Center Desc

Private Sub GetDates()
   Dim RdoCap As ADODB.Recordset
   Dim b As Byte
   Dim iCenters As Integer
   Dim dDate As Date
   
   Dim sBegMonth As String
   Dim sStartDate As String
   Dim sGetDate As String
   
   'Hourly calculations
   Dim cResources As Currency
   
   On Error GoTo DiaErr1
   z1(2).Visible = True
   z1(2).Caption = "Building Capacity"
   z1(2).Refresh
   prg1.Visible = True
   dDate = Format(txtBeg, "mm/dd/yy 00:00")
   
   'Get the starting Dates - Match the Calendar entries
   sBegMonth = Format(txtBeg, "mmm") & "-" & "20" & Right(txtBeg, 2)
   sStartDate = Format(txtBeg, "ddd")
   
   prg1.Value = 10
   'Past Due
   dCapaDates(0, 0) = "09/01/94 00:00"
   dCapaDates(0, 1) = dDate - 1
   'Starting
   dCapaDates(1, 0) = dDate
   b = Format(dDate, "w")
   If sWkStarts = "Sun" Then b = 8 - b Else _
                  b = 9 - b
   dDate = dDate + (b - 1)
   dCapaDates(1, 1) = dDate
   For b = 2 To 11
      dCapaDates(b, 0) = dCapaDates(b - 1, 1) + 1
      dCapaDates(b, 1) = dCapaDates(b - 1, 1) + 7
   Next
   dCapaDates(b, 0) = dCapaDates(b - 1, 1) + 1
   dCapaDates(b, 1) = dCapaDates(b - 1, 1) + 7
   sGetDate = Format(dCapaDates(b, 1), "mmm-yyyy")
   If Trim(sCenters(iTotalCenters, 0)) = "" Then
      iTotalCenters = iTotalCenters - 1
   End If
   prg1.Value = 20
   For iCenters = 0 To iTotalCenters
      Erase cHours
      For b = 0 To 11
         'cHours(b) = 0
         sSql = "SELECT * FROM WcclTable WHERE " _
                & "(WCCSHOP='" & Trim(sCenters(iCenters, 0)) & "' " _
                & "AND WCCCENTER='" & Trim(sCenters(iCenters, 3)) & "') " _
                & "AND WCCDATE " _
                & "BETWEEN '" & dCapaDates(b, 0) & " 00:00' " _
                & "AND '" & dCapaDates(b, 1) & " 23:59' ORDER BY WCCDATE"
         bSqlRows = clsADOCon.GetDataSet(sSql, RdoCap, ES_FORWARD)
         If bSqlRows Then
            With RdoCap
               Do Until .EOF
                  cResources = !WCCSHR1
                  If cResources = 0 Then cResources = 1
                  cHours(b) = cHours(b) + (!WCCSHH1 * cResources)
                  
                  cResources = !WCCSHR2
                  If cResources = 0 Then cResources = 1
                  cHours(b) = cHours(b) + (!WCCSHH2 * cResources)
                  
                  'cResources = !WCCSHR3
                  'If cResources = 0 Then cResources = 1
                  'cHours(b) = cHours(b) + (!WCCSHH3 * cResources)
                  
                  cResources = !WCCSHR3
                  If cResources = 0 Then cResources = 1
                  cHours(b) = cHours(b) + (!WCCSHH3 * cResources)
                  
                  cResources = !WCCSHR4
                  If cResources = 0 Then cResources = 1
                  cHours(b) = cHours(b) + (!WCCSHH4 * cResources)
                  .MoveNext
               Loop
               ClearResultSet RdoCap
            End With
         End If
      Next
      cHours(0) = 0
      sSql = "INSERT INTO EsReportCapa15a (" _
             & "RPTSHOPREF,RPTSHOPNUM,RPTSHOPDESC," _
             & "RPTWCNREF,RPTWCNNUM,RPTWCNDESC,"
      sSql = sSql & "RPTBEGDATE1,RPTENDDATE1,RPTHOURS1,"
      sSql = sSql & "RPTBEGDATE2,RPTENDDATE2,RPTHOURS2,"
      sSql = sSql & "RPTBEGDATE3,RPTENDDATE3,RPTHOURS3,"
      sSql = sSql & "RPTBEGDATE4,RPTENDDATE4,RPTHOURS4,"
      sSql = sSql & "RPTBEGDATE5,RPTENDDATE5,RPTHOURS5,"
      sSql = sSql & "RPTBEGDATE6,RPTENDDATE6,RPTHOURS6,"
      sSql = sSql & "RPTBEGDATE7,RPTENDDATE7,RPTHOURS7,"
      sSql = sSql & "RPTBEGDATE8,RPTENDDATE8,RPTHOURS8,"
      sSql = sSql & "RPTBEGDATE9,RPTENDDATE9,RPTHOURS9,"
      sSql = sSql & "RPTBEGDATE10,RPTENDDATE10,RPTHOURS10,"
      sSql = sSql & "RPTBEGDATE11,RPTENDDATE11,RPTHOURS11) "
      
      sSql = sSql & "VALUES('" _
             & Trim(sCenters(iCenters, 0)) & "','" _
             & Trim(sCenters(iCenters, 1)) & "','" _
             & Trim(sCenters(iCenters, 2)) & "','" _
             & Trim(sCenters(iCenters, 3)) & "','" _
             & Trim(sCenters(iCenters, 4)) & "','" _
             & Trim(sCenters(iCenters, 5)) & "','" _
             & Format$(dCapaDates(0, 0), "mm/dd/yy") & "','" _
             & Format$(dCapaDates(0, 1), "mm/dd/yy") & "'," _
             & cHours(0) & ",'" _
             & Format$(dCapaDates(1, 0), "mm/dd/yy") & "','" _
             & Format$(dCapaDates(1, 1), "mm/dd/yy") & "'," _
             & cHours(1) & ",'" _
             & Format$(dCapaDates(2, 0), "mm/dd/yy") & "','" _
             & Format$(dCapaDates(2, 1), "mm/dd/yy") & "'," _
             & cHours(2) & ",'" _
             & Format$(dCapaDates(3, 0), "mm/dd/yy") & "','" _
             & Format$(dCapaDates(3, 1), "mm/dd/yy") & "'," _
             & cHours(3) & ",'" _
             & Format$(dCapaDates(4, 0), "mm/dd/yy") & "','" _
             & Format$(dCapaDates(4, 1), "mm/dd/yy") & "'," _
             & cHours(4) & ",'"
      sSql = sSql & Format$(dCapaDates(5, 0), "mm/dd/yy") & "','" _
             & Format$(dCapaDates(5, 1), "mm/dd/yy") & "'," _
             & cHours(5) & ",'" _
             & Format$(dCapaDates(6, 0), "mm/dd/yy") & "','" _
             & Format$(dCapaDates(6, 1), "mm/dd/yy") & "'," _
             & cHours(6) & ",'" _
             & Format$(dCapaDates(7, 0), "mm/dd/yy") & "','" _
             & Format$(dCapaDates(7, 1), "mm/dd/yy") & "'," _
             & cHours(7) & ",'" _
             & Format$(dCapaDates(8, 0), "mm/dd/yy") & "','" _
             & Format$(dCapaDates(8, 1), "mm/dd/yy") & "'," _
             & cHours(8) & ",'" _
             & Format$(dCapaDates(9, 0), "mm/dd/yy") & "','" _
             & Format$(dCapaDates(9, 1), "mm/dd/yy") & "'," _
             & cHours(9) & ",'" _
             & Format$(dCapaDates(10, 0), "mm/dd/yy") & "','" _
             & Format$(dCapaDates(10, 1), "mm/dd/yy") & "'," _
             & cHours(10) & ")"
      clsADOCon.ExecuteSql sSql
   Next
   prg1.Value = 100
   Sleep 500
   prg1.Value = 0
   Set RdoCap = Nothing
   Exit Sub
   
DiaErr1:
   On Error Resume Next
   sSql = "UPDATE EsReportCapa15d SET RPTLOCKED=0"
   clsADOCon.ExecuteSql sSql
   
   sProcName = "getdates"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtBeg_LostFocus()
   txtBeg = CheckDateEx(txtBeg)
   
End Sub

'Local Errors

Function GetStartDay() As String
   On Error Resume Next
   Dim RdoStr As ADODB.Recordset
   sSql = "SELECT WEEKENDS FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoStr, ES_FORWARD)
   If bSqlRows Then
      If IsNull(RdoStr.Fields(0)) Then _
                GetStartDay = "Sat" Else GetStartDay = RdoStr.Fields(0)
      ClearResultSet RdoStr
   End If
   If GetStartDay = "Sat" Then GetStartDay = "Sunday" Else _
                    GetStartDay = "Monday"
   sWkStarts = Left$(GetStartDay, 3)
   Set RdoStr = Nothing
   
End Function

'0 = Shop Ref
'1 = Shop Name
'2 = Shop Desc
'3 = Work Center Ref
'4 = Work Center Name
'5 = Work Center Desc

Private Function GetCenters() As Byte
   Dim RdoCnt As ADODB.Recordset
   Dim b As Integer
   Dim iCenters As Integer
   Dim sShop As String
   Dim sWnc As String
   
   Erase sCenters
   On Error GoTo DiaErr1
   If cmbShp <> "ALL" Then sShop = Compress(cmbShp)
   If cmbWcn <> "ALL" Then sWnc = Compress(cmbWcn)
   
   sSql = "SELECT SHPREF,SHPNUM,SHPDESC,WCNREF,WCNNUM," _
          & "WCNDESC FROM ShopTable,WcntTable " _
          & "WHERE (WCNSHOP = SHPREF AND WCNSHOP LIKE '" & sShop & "%') " _
          & " AND WCNREF LIKE '" & sWnc & "%' " _
          & "AND WCNSERVICE = 0 " _
          & "ORDER BY WCNSHOP,WCNREF"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCnt, ES_FORWARD)
   If bSqlRows Then
      iCenters = -1
      With RdoCnt
         Do Until .EOF
            iCenters = iCenters + 1
            For b = 0 To 4
               sCenters(iCenters, b) = "" & .Fields(b)
            Next
            sCenters(iCenters, b) = "" & Trim(.Fields(b))
            .MoveNext
         Loop
         ClearResultSet RdoCnt
      End With
   End If
   iTotalCenters = iCenters
   If iTotalCenters > -1 Then
      sSql = "INSERT INTO EsReportCapa15d (RPTALLREF) " _
             & "VALUES('" & sCenters(iTotalCenters, 0) & "')"
      clsADOCon.ExecuteSql sSql
      GetCenters = 1
   Else
      MouseCursor 0
      MsgBox "No Selections To Report.", _
         vbInformation, Caption
      GetCenters = 0
   End If
   Set RdoCnt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getcenters"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

'DSS Table for Capacity

Private Sub CreateTable1()
   On Error Resume Next
   sSql = "SELECT RPTSHOPREF FROM EsReportCapa15a"
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum = 40002 Then
      clsADOCon.ADOErrNum = 0
      sSql = "Create Table dbo.EsReportCapa15a (" _
             & "RPTSHOPREF CHAR(10) NULL DEFAULT('')," _
             & "RPTSHOPNUM CHAR(10) NULL DEFAULT('')," _
             & "RPTSHOPDESC CHAR(30) NULL DEFAULT('')," _
             & "RPTWCNREF CHAR(10) NULL DEFAULT('')," _
             & "RPTWCNNUM CHAR(10) NULL DEFAULT('')," _
             & "RPTWCNDESC CHAR(30) NULL DEFAULT('')," _
             & "RPTBEGDATE1 CHAR(8) NULL DEFAULT('')," _
             & "RPTENDDATE1 CHAR(8) NULL DEFAULT('')," _
             & "RPTHOURS1 REAL NULL DEFAULT(0)," _
             & "RPTUSEDHOURS1 REAL NULL DEFAULT(0)," _
             & "RPTBEGDATE2 CHAR(8) NULL DEFAULT('')," _
             & "RPTENDDATE2 CHAR(8) NULL DEFAULT('')," _
             & "RPTHOURS2 REAL NULL DEFAULT(0)," _
             & "RPTUSEDHOURS2 REAL NULL DEFAULT(0),"
      sSql = sSql _
             & "RPTBEGDATE3 CHAR(8) NULL DEFAULT('')," _
             & "RPTENDDATE3 CHAR(8) NULL DEFAULT('')," _
             & "RPTHOURS3 REAL NULL DEFAULT(0)," _
             & "RPTUSEDHOURS3 REAL NULL DEFAULT(0)," _
             & "RPTBEGDATE4 CHAR(8) NULL DEFAULT('')," _
             & "RPTENDDATE4 CHAR(8) NULL DEFAULT('')," _
             & "RPTHOURS4 REAL NULL DEFAULT(0)," _
             & "RPTUSEDHOURS4 REAL NULL DEFAULT(0)," _
             & "RPTBEGDATE5 CHAR(8) NULL DEFAULT('')," _
             & "RPTENDDATE5 CHAR(8) NULL DEFAULT('')," _
             & "RPTHOURS5 REAL NULL DEFAULT(0)," _
             & "RPTUSEDHOURS5 REAL NULL DEFAULT(0)," _
             & "RPTBEGDATE6 CHAR(8) NULL DEFAULT('')," _
             & "RPTENDDATE6 CHAR(8) NULL DEFAULT('')," _
             & "RPTHOURS6 REAL NULL DEFAULT(0)," _
             & "RPTUSEDHOURS6 REAL NULL DEFAULT(0)," _
             & "RPTBEGDATE7 CHAR(8) NULL DEFAULT('')," _
             & "RPTENDDATE7 CHAR(8) NULL DEFAULT('')," _
             & "RPTHOURS7 REAL NULL DEFAULT(0)," _
             & "RPTUSEDHOURS7 REAL NULL DEFAULT(0),"
      sSql = sSql _
             & "RPTBEGDATE8 CHAR(8) NULL DEFAULT('')," _
             & "RPTENDDATE8 CHAR(8) NULL DEFAULT('')," _
             & "RPTHOURS8 REAL NULL DEFAULT(0)," _
             & "RPTUSEDHOURS8 REAL NULL DEFAULT(0)," _
             & "RPTBEGDATE9 CHAR(8) NULL DEFAULT('')," _
             & "RPTENDDATE9 CHAR(8) NULL DEFAULT('')," _
             & "RPTHOURS9 REAL NULL DEFAULT(0)," _
             & "RPTUSEDHOURS9 REAL NULL DEFAULT(0)," _
             & "RPTBEGDATE10 CHAR(8) NULL DEFAULT('')," _
             & "RPTENDDATE10 CHAR(8) NULL DEFAULT('')," _
             & "RPTHOURS10 REAL NULL DEFAULT(0)," _
             & "RPTUSEDHOURS10 REAL NULL DEFAULT(0)," _
             & "RPTBEGDATE11 CHAR(8) NULL DEFAULT('')," _
             & "RPTENDDATE11 CHAR(8) NULL DEFAULT('')," _
             & "RPTHOURS11 REAL NULL DEFAULT(0)," _
             & "RPTUSEDHOURS11 REAL NULL DEFAULT(0))"
      clsADOCon.ExecuteSql sSql
      If clsADOCon.ADOErrNum = 0 Then
         sSql = "CREATE CLUSTERED INDEX ReportRef ON dbo.EsReportCapa15a(RPTSHOPREF,RPTWCNREF) WITH  FILLFACTOR = 80"
         clsADOCon.ExecuteSql sSql
         
         sSql = "CREATE INDEX ShopRef ON dbo.EsReportCapa15a(RPTSHOPREF) WITH  FILLFACTOR = 80"
         clsADOCon.ExecuteSql sSql
         
         sSql = "CREATE INDEX WcnRef ON dbo.EsReportCapa15a(RPTWCNREF) WITH  FILLFACTOR = 80"
         clsADOCon.ExecuteSql sSql
      End If
      clsADOCon.ADOErrNum = 0
      
   End If
   
End Sub

Private Sub BuildReport()
   MouseCursor 13
   Dim RdoLck As ADODB.Recordset
   Dim b As Byte
   
   On Error Resume Next
   'In Use?
   sSql = "SELECT RPTLOCKED FROM EsReportCapa15d"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLck, ES_FORWARD)
   If bSqlRows Then b = RdoLck!RPTLOCKED
   Set RdoLck = Nothing
   If b = 1 Then
      MouseCursor 0
      MsgBox "The Report Tables Are In Use.  Please Come Back In A Few Minutes.", _
         vbInformation, Caption
   Else
      sSql = "TRUNCATE TABLE EsReportCapa15a"
      clsADOCon.ExecuteSql sSql
      sSql = "TRUNCATE TABLE EsReportCapa15b"
      clsADOCon.ExecuteSql sSql
      sSql = "TRUNCATE TABLE EsReportCapa15c"
      clsADOCon.ExecuteSql sSql
      sSql = "TRUNCATE TABLE EsReportCapa15d"
      clsADOCon.ExecuteSql sSql
      On Error GoTo DiaErr1
      sProcName = "getcenters"
      b = GetCenters()
      If b = 1 Then
         sProcName = "getdates"
         GetDates
         sProcName = "getusedhrs"
         GetUsedHours
         sProcName = "printreport"
         PrintReport
         On Error Resume Next
         sSql = "UPDATE EsReportCapa15d SET RPTLOCKED=0"
         clsADOCon.ExecuteSql sSql
      End If
   End If
   Exit Sub
   
DiaErr1:
   On Error Resume Next
   sSql = "UPDATE EsReportCapa15d SET RPTLOCKED=0"
   clsADOCon.ExecuteSql sSql
   
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetUsedHours()
   Dim RdoHrs As ADODB.Recordset
   
   Dim b As Byte
   Dim C As Byte
   
   Dim lRow As Long
   Dim cQuanity As Currency
   Dim cOpHrs As Currency
   
   Dim sColumn(2) As String
   Dim sCenter As String
   Dim sShop As String
   Dim sDesc As String
   sShop = ""
   sCenter = ""
   
   
   z1(2).Caption = "Building Used Hours"
   z1(2).Refresh
   prg1.Value = 10
   On Error GoTo DiaErr1
   
   If cmbShp <> "ALL" Then sShop = Compress(cmbShp)
   If cmbWcn <> "ALL" Then sCenter = Compress(cmbWcn)
   
   'For b = 0 To 11     'RESULTS IN RPTENDDATE12, not a valid column, below
   For b = 0 To 10
      prg1.Value = prg1.Value + 6
      
      If cmbWcn = "ALL" Then sCenter = ""

      sSql = "SELECT DISTINCT OPREF,OPRUN,OPNO,OPSHOP,OPCENTER," _
             & "OPSUHRS,OPUNITHRS,OPSUDATE,OPSCHEDDATE,RUNREF,RUNNO," _
             & "RUNSTATUS,RUNREMAININGQTY,WCNREF,WCNSERVICE FROM RnopTable," _
             & "RunsTable,WcntTable WHERE (OPREF=RUNREF AND " _
             & "OPRUN=RUNNO AND OPCENTER=WCNREF AND WCNSERVICE=0 AND " _
             & "OPCOMPLETE=0) AND OPSCHEDDATE BETWEEN '" & dCapaDates(b, 0) & " 00:00' " _
             & " AND '" & dCapaDates(b, 1) & " 23:09' " _
             & " AND OPSHOP LIKE '" & sShop & "%' AND OPCENTER LIKE '" & sCenter & "%' " _
             & " ORDER BY OPSCHEDDATE,OPREF,OPNO "
      'Debug.Print sSql
      
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoHrs, ES_FORWARD)
      If bSqlRows Then
         With RdoHrs
            Do Until .EOF
If lRow >= 7137 Then
Debug.Print lRow
End If
               cOpHrs = !OPSUHRS + (!OPUNITHRS * !RUNREMAININGQTY)
               sShop = "" & Trim(!OPSHOP)
               sCenter = "" & Trim(!OPCENTER)
               sDesc = GetDescription(!OPREF)
               'If not passed due then split the times
               If b > 0 Then
                  If Format(!OPSUDATE, "mm/dd/yy") < Format(dCapaDates(b, 0), "mm/dd/yy") Then
                     lRow = lRow + 1
                     cOpHrs = cOpHrs / 2
                     C = b
                     sColumn(0) = "RPTUSEDHOURS" & Trim(str(C))
                     sColumn(1) = "RPTENDDATE" & Trim(str(C))
                     sSql = "UPDATE EsReportCapa15a SET " & sColumn(0) & "=" _
                            & sColumn(0) & "+" _
                            & cOpHrs & " WHERE (" & sColumn(1) & "='" _
                            & Format$(dCapaDates(C - 1, 1), "mm/dd/yy") & "' AND " _
                            & "RPTSHOPREF='" & sShop & "' AND " _
                            & "RPTWCNREF='" & sCenter & "')"
                     clsADOCon.ExecuteSql sSql
                     
                     sSql = "INSERT INTO EsReportCapa15b (" _
                            & "RPTRECORD,RPTPARTREF,RPTPARTNUM,RPTPARTDESC," _
                            & "RPTSHOPREF,RPTWCNREF,RPTRUNNO,RPTOPNO,RPTRUNSTATUS," _
                            & "RPTREMAININGQTY," _
                            & "RPTSCHEDCOMPL,RPTENDDATE" & Trim(str(C)) _
                            & ",RPTHOURS" & Trim(str(C)) & ") VALUES(" _
                            & lRow & ",'" & Trim(!OPREF) & "','" & Trim(sPartNumber) _
                            & "','" & sDesc & "','" _
                            & sShop & "','" & sCenter & "'," & !OPRUN & ",'" _
                            & Format$(!opNo, "000") & "','" & !RUNSTATUS & "'," _
                            & Format(!RUNREMAININGQTY, "#0.000") & ",'" _
                            & Format$(!OPSCHEDDATE, "mm/dd/yy") & "','" _
                            & Format$(dCapaDates(C, 1), "mm/dd/yy") & "'," _
                            & cOpHrs & ")"
                     clsADOCon.ExecuteSql sSql
                  End If
               End If
               lRow = lRow + 1
               sColumn(0) = "RPTUSEDHOURS" & Trim(str(b + 1))
               sColumn(1) = "RPTENDDATE" & Trim(str(b + 1))
               sSql = "UPDATE EsReportCapa15a SET " & sColumn(0) & "=" _
                      & sColumn(0) & "+" _
                      & cOpHrs & " WHERE (" & sColumn(1) & "='" _
                      & Format$(dCapaDates(b, 1), "mm/dd/yy") & "' AND " _
                      & "RPTSHOPREF='" & sShop & "' AND " _
                      & "RPTWCNREF='" & sCenter & "')"
               clsADOCon.ExecuteSql sSql
               
               sSql = "INSERT INTO EsReportCapa15b (" _
                      & "RPTRECORD,RPTPARTREF,RPTPARTNUM,RPTPARTDESC," _
                      & "RPTSHOPREF,RPTWCNREF,RPTRUNNO,RPTOPNO,RPTRUNSTATUS," _
                      & "RPTREMAININGQTY," _
                      & "RPTSCHEDCOMPL,RPTENDDATE" & Trim(str(b + 1)) _
                      & ",RPTHOURS" & Trim(str(b + 1)) & ") VALUES(" _
                      & lRow & ",'" & Trim(!OPREF) & "','" & Trim(sPartNumber) _
                      & "','" & sDesc & "','" _
                      & sShop & "','" & sCenter & "'," & !OPRUN & ",'" _
                      & Format$(!opNo, "000") & "','" & !RUNSTATUS & "'," _
                      & Format(!RUNREMAININGQTY, "#0.000") & ",'" _
                      & Format$(!OPSCHEDDATE, "mm/dd/yy") & "','" _
                      & Format$(dCapaDates(b, 1), "mm/dd/yy") & "'," _
                      & cOpHrs & ")"

               clsADOCon.ExecuteSql sSql
               .MoveNext
            Loop
            ClearResultSet RdoHrs
         End With
      End If
   Next
   'House Cleaning
   prg1.Value = 100
   Sleep 500
   
   z1(2).Caption = "Finishing Report"
   z1(2).Refresh
   prg1.Value = 10
   sSql = "INSERT INTO EsReportCapa15c (RPTSHOP)" _
          & "SELECT distinct RPTSHOPREF FROM EsReportCapa15b"
   clsADOCon.ExecuteSql sSql
   For b = 2 To 11
      prg1.Value = prg1.Value + 3
      sSql = "UPDATE EsReportCapa15c SET RPTHOURS" & Trim(str(b)) & "=" _
             & "(SELECT SUM(RPTHOURS" & Trim(str(b)) & ") FROM EsReportCapa15a " _
             & "WHERE EsReportCapa15a.RPTSHOPREF=EsReportCapa15c.RPTSHOP)"
      clsADOCon.ExecuteSql sSql
   Next
   For b = 2 To 11
      prg1.Value = prg1.Value + 3
      sSql = "UPDATE EsReportCapa15d SET RPTALLHOURS" & Trim(str(b)) & "=" _
             & "(SELECT SUM(RPTHOURS" & Trim(str(b)) & ") FROM EsReportCapa15a) "
      clsADOCon.ExecuteSql sSql
   Next
   
   prg1.Value = 90
   sSql = "UPDATE EsReportCapa15b SET RPTPASTDUE='*' WHERE " _
          & "RPTHOURS1>0"
   clsADOCon.ExecuteSql sSql
   
   Set RdoHrs = Nothing
   prg1.Value = 100
   Sleep 500
   prg1.Visible = False
   z1(2).Visible = False
   Exit Sub
   
DiaErr1:
   'On Error Resume Next
   sSql = "UPDATE EsReportCapa15d SET RPTLOCKED=0"
   
   clsADOCon.ExecuteSql sSql
   sProcName = "GetUsedHours"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub CreateTable2()
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   sSql = "SELECT RPTSHOPREF FROM EsReportCapa15b"
   clsADOCon.ExecuteSql sSql
   
   'need to change RPTOPNO from 3 TO 5 characters
   If clsADOCon.ADOErrNum = 0 Then
      sSql = "insert EsReportCapa15b (RPTOPNO) values ( '1000' )"
      clsADOCon.ExecuteSql sSql
   End If
   If clsADOCon.ADOErrNum = 0 Then
      sSql = "truncate table EsReportCapa15b"
      clsADOCon.ExecuteSql sSql
   Else
   
   'If Err = 40002 Then
   'If Err <> 0 Then
      sSql = "drop table EsReportCapa15b"
      clsADOCon.ExecuteSql sSql
      clsADOCon.ADOErrNum = 0
      
      sSql = "Create Table dbo.EsReportCapa15b (" _
             & "RPTRECORD INT NULL DEFAULT(0)," _
             & "RPTPARTREF CHAR(30) NULL DEFAULT('')," _
             & "RPTPARTNUM CHAR(30) NULL DEFAULT('')," _
             & "RPTPARTDESC CHAR(30) NULL DEFAULT('')," _
             & "RPTSHOPREF CHAR(10) NULL DEFAULT('')," _
             & "RPTWCNREF CHAR(10) NULL DEFAULT('')," _
             & "RPTRUNNO INT NULL DEFAULT(0)," _
             & "RPTOPNO CHAR(5) NULL DEFAULT('')," _
             & "RPTRUNSTATUS CHAR(2) NULL DEFAULT('')," _
             & "RPTSCHEDCOMPL CHAR(8) NULL DEFAULT('')," _
             & "RPTREMAININGQTY REAL NULL DEFAULT(0)," _
             & "RPTENDDATE1 CHAR(8) NULL DEFAULT('')," _
             & "RPTHOURS1 REAL NULL DEFAULT(0)," _
             & "RPTENDDATE2 CHAR(8) NULL DEFAULT('')," _
             & "RPTHOURS2 REAL NULL DEFAULT(0)," _
             & "RPTENDDATE3 CHAR(8) NULL DEFAULT('')," _
             & "RPTHOURS3 REAL NULL DEFAULT(0)," _
             & "RPTENDDATE4 CHAR(8) NULL DEFAULT('')," _
             & "RPTHOURS4 REAL NULL DEFAULT(0)," _
             & "RPTENDDATE5 CHAR(8) NULL DEFAULT('')," _
             & "RPTHOURS5 REAL NULL DEFAULT(0),"
      sSql = sSql _
             & "RPTENDDATE6 CHAR(8) NULL DEFAULT('')," _
             & "RPTHOURS6 REAL NULL DEFAULT(0)," _
             & "RPTENDDATE7 CHAR(8) NULL DEFAULT('')," _
             & "RPTHOURS7 REAL NULL DEFAULT(0)," _
             & "RPTENDDATE8 CHAR(8) NULL DEFAULT('')," _
             & "RPTHOURS8 REAL NULL DEFAULT(0)," _
             & "RPTENDDATE9 CHAR(8) NULL DEFAULT('')," _
             & "RPTHOURS9 REAL NULL DEFAULT(0)," _
             & "RPTENDDATE10 CHAR(8) NULL DEFAULT('')," _
             & "RPTHOURS10 REAL NULL DEFAULT(0)," _
             & "RPTENDDATE11 CHAR(8) NULL DEFAULT('')," _
             & "RPTHOURS11 REAL NULL DEFAULT(0)," _
             & "RPTPASTDUE CHAR(1) NULL DEFAULT(''))"
      clsADOCon.ExecuteSql sSql
      If clsADOCon.ADOErrNum = 0 Then
         sSql = "CREATE CLUSTERED INDEX ReportRef ON dbo.EsReportCapa15b(RPTRECORD) WITH  FILLFACTOR = 80"
         clsADOCon.ExecuteSql sSql
         
         sSql = "CREATE INDEX ShopRef ON dbo.EsReportCapa15b(RPTSHOPREF) WITH  FILLFACTOR = 80"
         clsADOCon.ExecuteSql sSql
         
         sSql = "CREATE INDEX WcnRef ON dbo.EsReportCapa15b(RPTWCNREF) WITH  FILLFACTOR = 80"
         clsADOCon.ExecuteSql sSql
      End If
      clsADOCon.ADOErrNum = 0
      
   End If
   
End Sub

'Local errors

Private Function GetDescription(MONUMBER As String) As String
   Dim RdoDsc As ADODB.Recordset
   sSql = "SELECT PARTREF,PARTNUM,PADESC FROM PartTable WHERE PARTREF='" _
          & MONUMBER & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDsc, ES_FORWARD)
   If bSqlRows Then
      With RdoDsc
         GetDescription = "" & Trim(!PADESC)
         sPartNumber = "" & Trim(!PartNum)
         ClearResultSet RdoDsc
      End With
   End If
   Set RdoDsc = Nothing
   
End Function

Private Sub CreateTable3()
   On Error Resume Next
   sSql = "SELECT RPTSHOP FROM EsReportCapa15c"
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum = 40002 Then
      clsADOCon.ADOErrNum = 0
      sSql = "Create Table dbo.EsReportCapa15c (" _
             & "RPTSHOP CHAR(10) NULL DEFAULT('')," _
             & "RPTHOURS1 REAL NULL DEFAULT(0)," _
             & "RPTHOURS2 REAL NULL DEFAULT(0)," _
             & "RPTHOURS3 REAL NULL DEFAULT(0)," _
             & "RPTHOURS4 REAL NULL DEFAULT(0)," _
             & "RPTHOURS5 REAL NULL DEFAULT(0)," _
             & "RPTHOURS6 REAL NULL DEFAULT(0)," _
             & "RPTHOURS7 REAL NULL DEFAULT(0)," _
             & "RPTHOURS8 REAL NULL DEFAULT(0)," _
             & "RPTHOURS9 REAL NULL DEFAULT(0)," _
             & "RPTHOURS10 REAL NULL DEFAULT(0)," _
             & "RPTHOURS11 REAL NULL DEFAULT(0))"
      clsADOCon.ExecuteSql sSql
      If clsADOCon.ADOErrNum = 0 Then
         sSql = "CREATE UNIQUE CLUSTERED INDEX ReportRef ON dbo.EsReportCapa15c(RPTSHOP) " _
                & "WITH  FILLFACTOR = 80"
         clsADOCon.ExecuteSql sSql
      End If
      clsADOCon.ADOErrNum = 0
   End If
   
   'Third Table
   clsADOCon.ADOErrNum = 0
   sSql = "SELECT RPTSHOP FROM EsReportCapa15d"
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum = 40002 Then
      clsADOCon.ADOErrNum = 0
      sSql = "Create Table dbo.EsReportCapa15d (" _
             & "RPTALLREF CHAR(10) NULL DEFAULT('')," _
             & "RPTALLHOURS1 REAL NULL DEFAULT(0)," _
             & "RPTALLHOURS2 REAL NULL DEFAULT(0)," _
             & "RPTALLHOURS3 REAL NULL DEFAULT(0)," _
             & "RPTALLHOURS4 REAL NULL DEFAULT(0)," _
             & "RPTALLHOURS5 REAL NULL DEFAULT(0)," _
             & "RPTALLHOURS6 REAL NULL DEFAULT(0)," _
             & "RPTALLHOURS7 REAL NULL DEFAULT(0)," _
             & "RPTALLHOURS8 REAL NULL DEFAULT(0)," _
             & "RPTALLHOURS9 REAL NULL DEFAULT(0)," _
             & "RPTALLHOURS10 REAL NULL DEFAULT(0)," _
             & "RPTALLHOURS11 REAL NULL DEFAULT(0)," _
             & "RPTLOCKED TINYINT NULL DEFAULT(1))"
      clsADOCon.ExecuteSql sSql
      If clsADOCon.ADOErrNum = 0 Then
         sSql = "CREATE UNIQUE CLUSTERED INDEX ReportRef ON dbo.EsReportCapa15d(RPTALLREF) " _
                & "WITH  FILLFACTOR = 80"
         clsADOCon.ExecuteSql sSql
      End If
   End If
   
End Sub

Private Sub FillCenters()
   'Dim RdoWcn As ADODB.Recordset
   cmbWcn.Clear
   On Error GoTo DiaErr1
   sSql = "SELECT WCNREF,WCNNUM FROM WcntTable WHERE (WCNSHOP='" _
          & Compress(cmbShp) & "' AND WCNSERVICE = 0) " _
          & "ORDER BY WCNREF"
   LoadComboBox cmbWcn
   cmbWcn = "ALL"
   Exit Sub
   
DiaErr1:
   sProcName = "fillcenters"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
