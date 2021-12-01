VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaAPp11a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Received And Not Invoiced (Report)"
   ClientHeight    =   4635
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6735
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4635
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CheckBox ChkSub 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   3720
      Width           =   855
   End
   Begin VB.CheckBox ChkTyp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   7
      Left            =   4200
      TabIndex        =   9
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox ChkTyp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   6
      Left            =   3840
      TabIndex        =   8
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox ChkTyp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   5
      Left            =   3480
      TabIndex        =   7
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox ChkTyp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   4
      Left            =   3120
      TabIndex        =   6
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox ChkTyp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   5
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox ChkTyp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   4
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox ChkTyp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox chkExt 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Top             =   3480
      Width           =   855
   End
   Begin VB.CheckBox chkDesc 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   3240
      Width           =   855
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select From List Or Leave Blank"
      Top             =   840
      Width           =   1555
   End
   Begin VB.CheckBox chkPO 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2040
      TabIndex        =   13
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5520
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5520
      TabIndex        =   17
      Top             =   360
      Width           =   1215
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaAPp11a.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "diaAPp11a.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   16
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
      PictureUp       =   "diaAPp11a.frx":0308
      PictureDn       =   "diaAPp11a.frx":044E
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6000
      Top             =   1080
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4635
      FormDesignWidth =   6735
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
      PictureUp       =   "diaAPp11a.frx":0594
      PictureDn       =   "diaAPp11a.frx":06DA
   End
   Begin VB.CheckBox ChkTyp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   2
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "As Of What Date?"
      Height          =   285
      Index           =   11
      Left            =   120
      TabIndex        =   41
      Top             =   1700
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type Sub Report"
      Height          =   285
      Index           =   10
      Left            =   240
      TabIndex        =   40
      Top             =   3720
      Width           =   1665
   End
   Begin VB.Label zTyp 
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      Height          =   255
      Index           =   7
      Left            =   4200
      TabIndex        =   39
      Top             =   2280
      Width           =   180
   End
   Begin VB.Label zTyp 
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      Height          =   255
      Index           =   6
      Left            =   3840
      TabIndex        =   38
      Top             =   2280
      Width           =   180
   End
   Begin VB.Label zTyp 
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      Height          =   255
      Index           =   5
      Left            =   3480
      TabIndex        =   37
      Top             =   2280
      Width           =   180
   End
   Begin VB.Label zTyp 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      Height          =   255
      Index           =   4
      Left            =   3120
      TabIndex        =   36
      Top             =   2280
      Width           =   180
   End
   Begin VB.Label zTyp 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   35
      Top             =   2280
      Width           =   180
   End
   Begin VB.Label zTyp 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   34
      Top             =   2280
      Width           =   180
   End
   Begin VB.Label zTyp 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   33
      Top             =   2280
      Width           =   180
   End
   Begin VB.Label zTyp 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   32
      Top             =   2280
      Width           =   180
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Part Types:"
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   31
      Top             =   2520
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "By PO Number"
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   30
      Top             =   4200
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   29
      Top             =   3000
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Ext Description"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   28
      Top             =   3480
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Description"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   27
      Top             =   3240
      Width           =   1425
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   26
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Default Sorted By Date Received)"
      Height          =   285
      Index           =   3
      Left            =   3000
      TabIndex        =   24
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   0
      Left            =   3480
      TabIndex        =   23
      Top             =   840
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nickname"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   22
      Top             =   840
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor Name"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   21
      Top             =   1200
      Width           =   1425
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   20
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sort:"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   19
      Top             =   3960
      Width           =   1335
   End
End
Attribute VB_Name = "diaAPp11a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

' See the UpdateTables prodecure for database revisions

Option Explicit

'************************************************************************************
'
' diaAPp11a - Received but not invoiced.
'
' Notes:
'
' Created: (cjs)
' Revisions:
' 12/23/02 (nth) Added Part Desc and Ext Desc to report per JLH
' 10/23/03 (jcw) Redo Everything Including Database Transaction/Strip Out Access Code/Add Part Type
'
'***********************************************************************************

Dim bOnLoad As Byte
Dim bGoodVendor As Boolean
Dim iTotalChk As Integer

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'***********************************************************************************

Private Sub ChkTyp_GotFocus(Index As Integer)
   zTyp(Index).BorderStyle = 1
End Sub

Private Sub ChkTyp_LostFocus(Index As Integer)
   zTyp(Index).BorderStyle = 0
End Sub

Private Sub GetTotalCheck()
   Dim i As Integer
   iTotalChk = 0
   For i = 0 To 7
      If ChkTyp(i).Value = 1 Then
         iTotalChk = iTotalChk + 1
      End If
   Next
End Sub

Private Sub cmbVnd_Click()
   If cmbVnd <> "ALL" Then
      bGoodVendor = FindVendor(Me)
   Else
      lblNme = "All Vendors.."
   End If
End Sub


Private Sub cmbVnd_LostFocus()
   cmbVnd = CheckLen(cmbVnd, 10)
   If Len(cmbVnd) = 0 Then cmbVnd = "ALL"
   If cmbVnd <> "ALL" Then
      bGoodVendor = FindVendor(Me)
      If Trim(cmbVnd) = "" Then
         cmbVnd = "ALL"
      End If
   Else
      lblNme = "All Vendors.."
   End If
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
   Dim RdoVed As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT POVENDOR,VEREF,VENICKNAME " _
          & "FROM PohdTable,VndrTable WHERE POVENDOR=VEREF " _
          & "ORDER BY POVENDOR"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVed)
   If bSqlRows Then
      With RdoVed
         cmbVnd = "ALL"
         'cmbVnd.AddItem "ALL"
         AddComboStr cmbVnd.hWnd, "ALL"
         Do Until .EOF
            'cmbVnd.AddItem "" & Trim(!VENICKNAME)
            AddComboStr cmbVnd.hWnd, "" & Trim(!VENICKNAME)
            .MoveNext
         Loop
      End With
   End If
   lblNme = "All Vendors.."
   Set RdoVed = Nothing
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
      txtDte = Format(Now, "mm/dd/yy")
      bOnLoad = False
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   GetOptions
   bOnLoad = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Set diaAPp11a = Nothing
   FormUnload
End Sub

Private Sub PrintReport()
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   Dim aSortList As New Collection
   Dim sWindows As String
   Dim strBegNextDay As String
   Dim sEnd As String
   Dim i As Integer
   Dim sSqlTemp As String
   MouseCursor 13
   On Error GoTo DiaErr1
   
'   SetMdiReportsize MdiSect
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    sCustomReport = GetCustomReport("finap11b.rpt")
    
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Includes Vendor(s) " & CStr(cmbVnd) & "'")
    aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   
   
   strBegNextDay = Format(DateAdd("d", 1, CDate(txtDte)), "mm/dd/yy")
   
   sSql = "{PoitTable.PIADATE}" _
          & " <= DateTime('" & Trim(txtDte) & "')" _
          & " and ({VihdTable.VIDATE} >= DateTime('" & Trim(strBegNextDay) & "')  or  {PoitTable.PITYPE} = 15 )"
   
   If Trim(cmbVnd) <> "ALL" Then
      sSql = sSql & " and {PoitTable.PIVENDOR} = '" & Compress(cmbVnd) & "' "
      
      aFormulaName.Add "Vendor"
      aFormulaName.Add "VendorRef"
      aFormulaValue.Add CStr("'Vendor:" & CStr(Trim(cmbVnd)) & "'")
      aFormulaValue.Add CStr("'" & CStr(Compress(cmbVnd)) & "'")
   Else
      aFormulaName.Add "Vendor"
      aFormulaName.Add "VendorRef"
      aFormulaValue.Add CStr("'Vendor: All'")
      aFormulaValue.Add CStr("'All'")
   End If
   
   GetTotalCheck
   
   If iTotalChk <> 8 And iTotalChk <> 0 Then
      sSql = sSql & " and {PartTable.PALEVEL} IN["
      For i = 0 To 7
         If ChkTyp(i).Value = 1 Then
            sSql = sSql & zTyp(i) & ","
         End If
      Next
      sSql = Left(sSql, Len(sSql) - 1)
      sSql = sSql & "]"
   End If
   
   If chkDesc = vbChecked Then
      aFormulaName.Add "Desc"
      aFormulaValue.Add "'0'"
   Else
      aFormulaName.Add "Desc"
      aFormulaValue.Add "'1'"
   End If
   
   If chkExt = vbChecked Then
      aFormulaName.Add "ExtDesc"
      aFormulaValue.Add "'0'"
   Else
      aFormulaName.Add "ExtDesc"
      aFormulaValue.Add "'1'"
   End If
   
   If ChkSub = vbChecked Then
      aFormulaName.Add "Sub"
      aFormulaValue.Add "'1'"
   Else
      aFormulaName.Add "Sub"
      aFormulaValue.Add "'0'"
   End If
   
   If iTotalChk <> 8 And iTotalChk <> 0 Then
      For i = 1 To 8
         If ChkTyp(i - 1).Value = vbChecked Then
             aFormulaName.Add "pType" & i
             aFormulaValue.Add CStr("'" & CStr(i) & "'")
         Else
             aFormulaName.Add "pType" & i
             aFormulaValue.Add CStr("'0'")
         End If
      Next
   Else
      For i = 1 To 8
          aFormulaName.Add "pType" & i
          aFormulaValue.Add CStr("'" & CStr(i) & "'")
      Next
   End If
   
    aFormulaName.Add "AsOf"
    aFormulaValue.Add CStr("'" & CStr(Trim(txtDte)) & "'")
   
   If chkPO.Value = vbChecked Then
      aSortList.Add "PINUMBER"
      aSortList.Add "PIITEM"
      aSortList.Add "PIREV"
   Else
      aSortList.Add "PIADATE"
      aSortList.Add "PINUMBER"
      aSortList.Add "PIITEM"
      aSortList.Add "PIREV"
   End If
   
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.SetSortFields aSortList
    cCRViewer.SetReportSelectionFormula sSql
    cCRViewer.CRViewerSize Me
    cCRViewer.SetDbTableConnection
    cCRViewer.OpenCrystalReportObject Me, aFormulaName
    cCRViewer.ShowGroupTree False
    
    cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aFormulaName
    cCRViewer.ClearFieldCollection aFormulaValue
    
    
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub PrintReport1()
   Dim sWindows As String
   Dim sBeg As String
   Dim sEnd As String
   Dim i As Integer
   Dim sSqlTemp As String
   MouseCursor 13
   On Error GoTo DiaErr1
   
   
   'SetMdiReportsize MdiSect
   
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "Includes='Includes Vendor(s) " & cmbVnd & "...'"
   MdiSect.crw.Formulas(2) = "RequestBy='Requested By: " & sInitials & "'"
   
   sSql = "{PoitTable.PIADATE}" _
          & " <= DateTime('" & Trim(txtDte) & "')" _
          & " and ({VihdTable.VIDATE} > DateTime('" & Trim(txtDte) & "')  or  {PoitTable.PITYPE} = 15 )"
   
   If Trim(cmbVnd) <> "ALL" Then
      sSql = sSql & " and {PoitTable.PIVENDOR} = '" & Compress(cmbVnd) & "' "
      MdiSect.crw.Formulas(3) = "Vendor='Vendor:" & Trim(cmbVnd) & "'"
      MdiSect.crw.Formulas(4) = "VendorRef='" & Compress(cmbVnd) & "'"
   Else
      MdiSect.crw.Formulas(3) = "Vendor='Vendor: All'"
      MdiSect.crw.Formulas(4) = "VendorRef='All'"
   End If
   
   GetTotalCheck
   
   If iTotalChk <> 8 And iTotalChk <> 0 Then
      sSql = sSql & " and {PartTable.PALEVEL} IN["
      For i = 0 To 7
         If ChkTyp(i).Value = 1 Then
            sSql = sSql & zTyp(i) & ","
         End If
      Next
      sSql = Left(sSql, Len(sSql) - 1)
      sSql = sSql & "]"
   End If
   
   If chkDesc = vbChecked Then
      MdiSect.crw.Formulas(5) = "Desc = '0'"
   Else
      MdiSect.crw.Formulas(5) = "Desc = '1'"
   End If
   
   If chkExt = vbChecked Then
      MdiSect.crw.Formulas(6) = "ExtDesc = '0'"
   Else
      MdiSect.crw.Formulas(6) = "ExtDesc = '1'"
   End If
   
   If ChkSub = vbChecked Then
      MdiSect.crw.Formulas(7) = "Sub='1'"
   Else
      MdiSect.crw.Formulas(7) = "Sub='0'"
   End If
   
   If iTotalChk <> 8 And iTotalChk <> 0 Then
      For i = 1 To 8
         If ChkTyp(i - 1).Value = vbChecked Then
            MdiSect.crw.Formulas(7 + i) = "pType" & i & " ='" & i & "'"
         Else
            MdiSect.crw.Formulas(7 + i) = "pType" & i & " ='0'"
         End If
      Next
   Else
      For i = 1 To 8
         MdiSect.crw.Formulas(7 + i) = "pType" & i & " ='" & i & "'"
      Next
   End If
   
   MdiSect.crw.Formulas(16) = "AsOf='" & Trim(txtDte) & "'"
   
   If chkPO.Value = vbChecked Then
      MdiSect.crw.SortFields(0) = "+{PoitTable.PINUMBER}"
      MdiSect.crw.SortFields(1) = "+{PoitTable.PIITEM}"
      MdiSect.crw.SortFields(2) = "+{PoitTable.PIREV}"
   Else
      MdiSect.crw.SortFields(0) = "+{PoitTable.PIADATE}"
      MdiSect.crw.SortFields(1) = "+{PoitTable.PINUMBER}"
      MdiSect.crw.SortFields(2) = "+{PoitTable.PIITEM}"
      MdiSect.crw.SortFields(3) = "+{PoitTable.PIREV}"
   End If
   
   MdiSect.crw.SelectionFormula = sSql
   MdiSect.crw.ReportFileName = sReportPath & "finap11b.rpt"
   'SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Public Sub SaveOptions()
   Dim sOptions As String
   Dim i As Integer
   'Save by Menu Option
   sOptions = Trim(str(chkDesc.Value)) _
              & Trim(str(chkExt.Value)) _
              & Trim(str(ChkSub.Value)) _
              & Trim(str(chkPO.Value))
   For i = 0 To 7
      sOptions = sOptions & Trim(str(ChkTyp(i).Value))
   Next
   SaveSetting "Esi2000", "EsiFina", Me.Name, Trim(sOptions)
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Public Sub GetOptions()
   Dim sOptions As String
   Dim i As Integer
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
      chkDesc.Value = Val(Mid(sOptions, 1, 1))
      chkExt.Value = Val(Mid(sOptions, 2, 1))
      ChkSub.Value = Val(Mid(sOptions, 3, 1))
      chkPO.Value = Val(Mid(sOptions, 4, 1))
      For i = 0 To 7
         ChkTyp(i) = Val(Mid(sOptions, 5 + i, 1))
      Next
   Else
      chkDesc.Value = vbUnchecked
      chkExt.Value = vbUnchecked
      ChkSub.Value = vbChecked
      chkPO.Value = vbUnchecked
      For i = 0 To 7
         ChkTyp(i) = vbChecked
      Next
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
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

Private Sub txtDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtDte_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtDte_LostFocus()
   txtDte = CheckDate(txtDte)
End Sub
