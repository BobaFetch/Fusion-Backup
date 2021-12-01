VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ShopSHp09a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Late MO's By Operation"
   ClientHeight    =   3180
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7050
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3180
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   1320
      Width           =   3075
   End
   Begin VB.CheckBox optSO 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   252
      Left            =   1680
      TabIndex        =   21
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "ShopSHp09a.frx":0000
      Height          =   315
      Left            =   4800
      Picture         =   "ShopSHp09a.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   1320
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHp09a.frx":0684
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtDay 
      Height          =   285
      Left            =   3480
      TabIndex        =   6
      Tag             =   "1"
      Top             =   1680
      Width           =   495
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1680
      TabIndex        =   7
      Top             =   2280
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   1680
      TabIndex        =   5
      Tag             =   "4"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Leading Character Search"
      Top             =   1320
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5880
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5880
      TabIndex        =   8
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "ShopSHp09a.frx":0E32
         Style           =   1  'Graphical
         TabIndex        =   9
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
         TabIndex        =   10
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbWcn 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select From List"
      Top             =   960
      Width           =   1815
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6240
      Top             =   2520
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3180
      FormDesignWidth =   7050
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include SO"
      Height          =   285
      Index           =   9
      Left            =   480
      TabIndex        =   22
      Top             =   2640
      Width           =   915
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Or"
      Height          =   285
      Index           =   8
      Left            =   2760
      TabIndex        =   19
      Tag             =   " "
      Top             =   1680
      Width           =   705
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Days Or More Late (0 For ALL)"
      Height          =   285
      Index           =   7
      Left            =   4200
      TabIndex        =   18
      Tag             =   " "
      Top             =   1695
      Width           =   2610
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   4
      Left            =   5520
      TabIndex        =   17
      Top             =   1320
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   2
      Left            =   5520
      TabIndex        =   16
      Top             =   960
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   15
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
      TabIndex        =   14
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Late As Of"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   13
      Tag             =   " "
      Top             =   1695
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   12
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
      TabIndex        =   4
      Top             =   1320
      Width           =   1185
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


Public Sub FillPartCombo(cmbPrt As ComboBox)
   Dim RdoPart As ADODB.Recordset
   On Error GoTo DiaErr1
   
   sSql = "Qry_FillParts"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPart, ES_FORWARD)
   
   If bSqlRows Then
      With RdoPart
         While Not .EOF
            AddComboStr cmbPrt.hwnd, "" & Trim(!PartNum)
            .MoveNext
         Wend
         .Cancel
      End With
   End If
   Set RdoPart = Nothing
   cmbPrt.ListIndex = 0
   Exit Sub
   
DiaErr1:
   sProcName = "FillPartCombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
End Sub


Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      FillPartCombo cmbPrt
      cmbPrt = "ALL"
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
    Dim sEndDate As String
    Dim sIncDate As String
    Dim sPartNumber As String
    
    MouseCursor 13
    sIncDate = GetReportDate()
    sBegDate = Format(GetReportDate(), "yyyy,mm,dd")
    sEndDate = Format(txtDte, "yyyy,mm,dd")
    
    If cmbPrt = "ALL" Then sPartNumber = "" Else sPartNumber = Compress(cmbPrt)
    If cmbWcn = "ALL" Then sCenter = "" Else sCenter = Compress(cmbWcn)
    
    On Error GoTo DiaErr1
   
    Dim sCustomReport As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
    Dim strIncludes As String
    Dim strRequestBy As String
    Dim strIncDate As String
    Dim strFullPath As String
   
    sCustomReport = GetCustomReport("prdsh10")
    strFullPath = sReportPath & sCustomReport
    If (Not CheckPath(strFullPath)) Then
        MsgBox ("Crystal report File was not found - " & strFullPath)
        Exit Sub
    End If

    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath

    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.ShowGroupTree False
    
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "IncDate"
    aFormulaName.Add "ShowPartDesc"
    aFormulaName.Add "showSO"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")

    strIncludes = "'Parts " & txtDte & " And Centers " & cmbWcn & "...'"
    aFormulaValue.Add CStr(strIncludes)
    strRequestBy = "'Requested By: " & sInitials & "'"
    aFormulaValue.Add CStr(strRequestBy)
    strIncDate = "'From " & sIncDate & " To " & txtDte & "'"
    aFormulaValue.Add CStr(strIncDate)
    aFormulaValue.Add CStr(optDsc)
    aFormulaValue.Add CStr(optSO)
    
        sSql = "{RnopTable.OPCOMPLETE} = 0.00 AND " _
            & "{RunsTable.RUNSTATUS} in ['SC', 'PC', 'PP', 'PL', 'RL'] AND " _
            & "{RunsTable.RUNREF} Like '" & sPartNumber & "*' AND " _
           & "{WcntTable.WCNREF} Like '" & sCenter & "*' AND " _
           & "{RnopTable.OPSCHEDDATE} <= Date(" & sEndDate & ")"
    
    ' Set Formula values
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
    cCRViewer.SetReportDistinctRecords True
    ' set the report Selection
    cCRViewer.SetReportSelectionFormula (sSql)
    cCRViewer.CRViewerSize Me
    
    ' Set report parameter
    cCRViewer.SetDbTableConnection


    cCRViewer.OpenCrystalReportObject Me, aFormulaName

    cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aRptParaType
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
   txtPrt = "ALL"
   txtDte = Format(ES_SYSDATE, "mm/dd/yyyy")
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   'Save by Menu Option
   sOptions = Trim(str(optDsc.Value)) _
              & Format$(txtDay, "000")
   SaveSetting "Esi2000", "EsiProd", "sh10", Trim(sOptions)
   SaveSetting "Esi2000", "EsiProd", "sh10", Trim(optSO)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   Dim sSO As String
   
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "sh10", sOptions)
   If Len(sOptions) Then
      optDsc.Value = Val(Mid(sOptions, 1, 1))
      txtDay = Val(Mid(sOptions, 2, 3))
   Else
      txtDay = 0
   End If
   
   sSO = GetSetting("Esi2000", "EsiProd", "sh10", sSO)
   optSO.Value = Val(sSO)

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
   ShowCalendarEx Me
   
End Sub


Private Sub txtDte_LostFocus()
   txtDte = CheckDateEx(txtDte)
   
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
   dDate = Format(txtDte, "mm/dd/yyyy")
   dDate = Format(dDate - iList, "mm/dd/yyyy")
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
   txtDte = Format(l + Val(txtDay), "mm/dd/yyyy")
   
End Sub

