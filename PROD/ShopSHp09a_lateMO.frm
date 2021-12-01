VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ShopSHp09a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Late MO's By Operation"
   ClientHeight    =   3180
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7230
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3180
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optSoNum 
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1920
      TabIndex        =   6
      ToolTipText     =   "Include Split Number"
      Top             =   2640
      Value           =   1  'Checked
      Width           =   375
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "ShopSHp09a.frx":0000
      Height          =   315
      Left            =   5040
      Picture         =   "ShopSHp09a.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   1320
      Width           =   350
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHp09a.frx":0684
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtDay 
      Height          =   285
      Left            =   3720
      TabIndex        =   4
      Tag             =   "1"
      Top             =   1680
      Width           =   495
   End
   Begin VB.CheckBox optDsc 
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1920
      TabIndex        =   5
      Top             =   2280
      Value           =   1  'Checked
      Width           =   375
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   1920
      TabIndex        =   3
      Tag             =   "4"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Leading Character Search"
      Top             =   1320
      Width           =   3075
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
      TabIndex        =   7
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "ShopSHp09a.frx":0E32
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "ShopSHp09a.frx":0FB0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbWcn 
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select From List"
      Top             =   960
      Width           =   1815
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6840
      Top             =   2520
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3180
      FormDesignWidth =   7230
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include SO Number"
      Height          =   285
      Index           =   9
      Left            =   240
      TabIndex        =   21
      Tag             =   " "
      Top             =   2640
      Width           =   1545
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Or"
      Height          =   285
      Index           =   8
      Left            =   3000
      TabIndex        =   18
      Tag             =   " "
      Top             =   1680
      Width           =   705
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Days Or More Late (0 For ALL)"
      Height          =   285
      Index           =   7
      Left            =   4440
      TabIndex        =   17
      Tag             =   " "
      Top             =   1695
      Width           =   2610
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   4
      Left            =   5760
      TabIndex        =   16
      Top             =   1320
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   2
      Left            =   5760
      TabIndex        =   15
      Top             =   960
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   14
      Tag             =   " "
      Top             =   2280
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   13
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Late As Of"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   12
      Tag             =   " "
      Top             =   1700
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Tag             =   " "
      Top             =   960
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1425
   End
End
Attribute VB_Name = "ShopSHp09a"
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




Private Sub cmbWcn_LostFocus()
   cmbWcn = CheckLen(cmbWcn, 12)
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdFnd_Click()
   ViewParts.lblControl = "TXTPRT"
   ViewParts.txtPrt = txtPrt
   ViewParts.Show
   
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
   sSql = "Qry_FillWorkCentersAll"
   LoadComboBox cmbWcn
   cmbWcn = "ALL"
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
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
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set ShopSHp09a = Nothing
   
End Sub




Private Sub PrintReport()
   Dim sCenter As String
   Dim sBegDate As String
   Dim sEnddate As String
   Dim sIncDate As String
   Dim sPartNumber As String
   
   MouseCursor 13
   sIncDate = GetReportDate()
   sBegDate = Format(GetReportDate(), "yyyy,mm,dd")
   sEnddate = Format(txtDte, "yyyy,mm,dd")
   
   If txtPrt = "ALL" Then sPartNumber = "" Else sPartNumber = Compress(txtPrt)
   If cmbWcn = "ALL" Then sCenter = "" Else sCenter = Compress(cmbWcn)
   
   On Error GoTo DiaErr1
   SetMdiReportsize MDISect
   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MDISect.Crw.Formulas(1) = "Includes='Parts " & txtPrt & " And " _
                        & "Centers " & cmbWcn & "...'"
   MDISect.Crw.Formulas(2) = "RequestBy = 'Requested By: " & sInitials & "'"
   MDISect.Crw.Formulas(3) = "IncDate='From " & sIncDate & " To " & txtDte & "'"
   
   sCustomReport = GetCustomReport("prdsh10")
   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   sSql = "{RunsTable.RUNREF} Like '" & sPartNumber & "*' AND " _
          & "{WcntTable.WCNREF} Like '" & sCenter & "*' AND " _
          & "{RnopTable.OPSCHEDDATE} >= Date(" & sEnddate & ")"
   
   If optDsc.value = vbUnchecked Then
      MDISect.Crw.Formulas(4) = "Desc=''"
   Else
      MDISect.Crw.Formulas(4) = "Desc='1'"
   End If

   If optSoNum.value = vbUnchecked Then
      MDISect.Crw.Formulas(5) = "showSONum=''"
   Else
      MDISect.Crw.Formulas(5) = "showSONum='1'"
   End If
   
   
   MDISect.Crw.SelectionFormula = sSql
   SetCrystalAction Me
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
   txtPrt = "ALL"
   txtDte = Format(ES_SYSDATE, "mm/dd/yy")
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   'Save by Menu Option
   sOptions = Trim(str(optDsc.value)) _
              & Format$(txtDay, "000")
   SaveSetting "Esi2000", "EsiProd", "sh10", Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "sh10", sOptions)
   If Len(sOptions) Then
      optDsc.value = Val(Mid(sOptions, 1, 1))
      txtDay = Val(Mid(sOptions, 2, 3))
   Else
      txtDay = 0
   End If
   
   
End Sub

Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub


Private Sub txtDay_LostFocus()
   txtDay = CheckLen(txtDay, 3)
   txtDay = Format(Abs(Val(txtDay)), "##0")
   GetEndDate
   
End Sub


Private Sub txtDte_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtDte_LostFocus()
   txtDte = CheckDate(txtDte)
   
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
   If txtPrt = "" Then txtPrt = "ALL"
   
End Sub



'Get the dates for the report

Private Function GetReportDate() As String
   Dim iList As Integer
   Dim dDate As Date
   
   On Error GoTo DiaErr1
   'More than 4 years then they are screwed anyway
   iList = Val(txtDay)
   If iList = 0 Then iList = 1460
   dDate = Format(txtDte, "mm/dd/yy")
   dDate = Format(dDate - iList, "mm/dd/yy")
   GetReportDate = dDate
   Exit Function
   
DiaErr1:
   sProcName = "getreportda"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
   
End Function

Private Sub GetEndDate()
   Dim l As Long
   l = DateValue(txtDte)
   txtDte = Format(l + Val(txtDay), "mm/dd/yy")
   
End Sub
