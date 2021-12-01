VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaGLp09a 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Budgets"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   7200
   Begin VB.ComboBox cmbDiv 
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Enter/Revise A Division (2 char)"
      Top             =   1920
      Width           =   660
   End
   Begin VB.ComboBox cmbAct 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CheckBox optPer 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Index           =   12
      Left            =   4680
      TabIndex        =   15
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox optPer 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Index           =   11
      Left            =   4320
      TabIndex        =   14
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox optPer 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Index           =   10
      Left            =   3960
      TabIndex        =   13
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox optPer 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Index           =   9
      Left            =   3600
      TabIndex        =   12
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox optPer 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Index           =   8
      Left            =   3240
      TabIndex        =   11
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox optPer 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Index           =   7
      Left            =   2880
      TabIndex        =   10
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox optPer 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Index           =   6
      Left            =   2520
      TabIndex        =   9
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox optPer 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Index           =   5
      Left            =   2160
      TabIndex        =   8
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox optPer 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Index           =   4
      Left            =   1800
      TabIndex        =   7
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox optPer 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Index           =   3
      Left            =   1440
      TabIndex        =   6
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox optPer 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Index           =   2
      Left            =   1080
      TabIndex        =   5
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox optPer 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Index           =   1
      Left            =   720
      TabIndex        =   4
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox optPer 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   6000
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Width           =   1095
   End
   Begin VB.CheckBox optIna 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   3600
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   5760
      TabIndex        =   20
      Top             =   240
      Width           =   1335
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Display The Report"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Print The Report"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbYer 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Tag             =   "1"
      Top             =   600
      Width           =   1095
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   4200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4350
      FormDesignWidth =   7200
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   21
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaGLp09a.frx":0000
      PictureDn       =   "diaGLp09a.frx":0146
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   22
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
      PictureUp       =   "diaGLp09a.frx":028C
      PictureDn       =   "diaGLp09a.frx":03D2
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   4
      Left            =   3720
      TabIndex        =   44
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   43
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   42
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   41
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Account"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   40
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Periods:"
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   39
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Label lblPer 
      Alignment       =   2  'Center
      Caption         =   "13"
      Height          =   255
      Index           =   12
      Left            =   4650
      TabIndex        =   38
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label lblPer 
      Alignment       =   2  'Center
      Caption         =   "12"
      Height          =   255
      Index           =   11
      Left            =   4290
      TabIndex        =   37
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label lblPer 
      Alignment       =   2  'Center
      Caption         =   "11"
      Height          =   255
      Index           =   10
      Left            =   3900
      TabIndex        =   36
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label lblPer 
      Alignment       =   2  'Center
      Caption         =   "10"
      Height          =   255
      Index           =   9
      Left            =   3570
      TabIndex        =   35
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label lblPer 
      Alignment       =   2  'Center
      Caption         =   "9"
      Height          =   255
      Index           =   8
      Left            =   3210
      TabIndex        =   34
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label lblPer 
      Alignment       =   2  'Center
      Caption         =   "8"
      Height          =   255
      Index           =   7
      Left            =   2850
      TabIndex        =   33
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label lblPer 
      Alignment       =   2  'Center
      Caption         =   "7"
      Height          =   255
      Index           =   6
      Left            =   2490
      TabIndex        =   32
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label lblPer 
      Alignment       =   2  'Center
      Caption         =   "6"
      Height          =   255
      Index           =   5
      Left            =   2130
      TabIndex        =   31
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label lblPer 
      Alignment       =   2  'Center
      Caption         =   "5"
      Height          =   255
      Index           =   4
      Left            =   1770
      TabIndex        =   30
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label lblPer 
      Alignment       =   2  'Center
      Caption         =   "4"
      Height          =   255
      Index           =   3
      Left            =   1410
      TabIndex        =   29
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label lblPer 
      Alignment       =   2  'Center
      Caption         =   "3"
      Height          =   255
      Index           =   2
      Left            =   1050
      TabIndex        =   28
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label lblPer 
      Alignment       =   2  'Center
      Caption         =   "2"
      Height          =   255
      Index           =   1
      Left            =   690
      TabIndex        =   27
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label lblPer 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   330
      TabIndex        =   26
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Inactive Accounts"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   25
      Top             =   3600
      Width           =   2025
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   24
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fiscal Year"
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   23
      Top             =   600
      Width           =   915
   End
End
Attribute VB_Name = "diaGLp09a"
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
'
' diaGLp09a - View Budgets
'
' Created:  2/06/04 (nth)
' Revisions:
'   02/23/04 (JCW) Fixed Misc. Bugs
'   08/16/04 (nth) Added printer to getoptions and saveoptions
'
'*************************************************************************************

Option Explicit

Dim sOptions As String
Dim bOnLoad As Byte
Dim iStart As Integer
Dim iEnd As Integer


Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd
'*************************************************************************************


Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, Caption
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub Form_Activate()
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
   'optPrn.Picture = Resources.imgPrn.Picture
   'optDis.Picture = Resources.imgDis.Picture
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
   FormUnload
   
   Set diaGLp09a = Nothing
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub optPer_GotFocus(Index As Integer)
   lblPer(Index).BorderStyle = 1
End Sub

Private Sub optPer_LostFocus(Index As Integer)
   lblPer(Index).BorderStyle = 0
   sOptions = reParseOptions
End Sub

Private Sub cmbDiv_LostFocus()
   On Error Resume Next
   If Trim(cmbDiv) <> "" And Not bValidElement(cmbDiv) Then
      cmbDiv = ""
   End If
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub cmbYer_LostFocus()
   On Error Resume Next
   
   If Val(cmbYer) < 32000 Then
      cmbYer = CInt(Val(cmbYer))
      If bValidElement(cmbYer) Then
         LoadPeriods
      Else
         ClearPer False
         cmbYer = ""
      End If
   Else
      cmbYer = ""
      ClearPer False
   End If
   
End Sub

Private Sub cmbAct_Click()
   FindAccount Me
End Sub

Private Sub cmbAct_LostFocus()
   On Error Resume Next
   cmbAct = CheckLen(cmbAct, 12)
   FindAccount Me
End Sub

Private Sub cmbYer_Click()
   LoadPeriods
End Sub

Public Sub FillCombo()
   Dim rdoCombo As ADODB.Recordset
   Dim i As Integer
   On Error GoTo DiaErr1
   
   If bDivisionAccounts(iStart, iEnd) Then
      FillDivisions Me
   Else
      cmbDiv.enabled = False
   End If
   
   sSql = "Qry_FillAccountCombo"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCombo)
   If bSqlRows Then
      With rdoCombo
         Do While Not .EOF
            AddComboStr cmbAct.hWnd, "" & !GLACCTNO
            rdoCombo.MoveNext
         Loop
      End With
   End If
   
   If cmbAct.ListCount > 0 Then
      cmbAct.ListIndex = 0
   End If
   
   
   Set rdoCombo = Nothing
   
   sSql = "SELECT FYYEAR from GlfyTable "
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCombo)
   If bSqlRows Then
      With rdoCombo
         Do While Not .EOF
            AddComboStr cmbYer.hWnd, "" & !FYYEAR
            rdoCombo.MoveNext
         Loop
      End With
   End If
   
   cmbYer = Format(Now, "yyyy")
   If Not bValidElement(cmbYer) Then
      If cmbYer.ListCount > 0 Then
         cmbYer.ListIndex = 0
      Else
         cmbYer = ""
      End If
   End If
   
   Set rdoCombo = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub SaveOptions()
   Dim sOptions2 As String
   Dim i As Integer
   sOptions2 = reParseOptions
   SaveSetting "Esi2000", "EsiFina", Me.Name, Trim(sOptions2)
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim i As Integer
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
      For i = 0 To 12
         optPer(i).Value = Val(Mid(sOptions, i + 1, 1))
      Next
      optIna.Value = Val(Mid(sOptions, 14, 1))
   Else
      For i = 0 To 12
         optPer(i).Value = vbChecked
      Next
      optIna.Value = 0
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = TTDEFAULT
   End If
End Sub

Private Sub ClearPer(bBool As Boolean)
   Dim i As Integer
   On Error Resume Next
   
   For i = 0 To 12
      optPer(i).enabled = bBool
      optPer(i).Value = bBool
      
   Next
End Sub


Private Sub LoadPeriods()
   Dim i As Integer
   Dim rdoPer As ADODB.Recordset
   On Error GoTo DiaErr1
   
   ClearPer False
   
   sSql = "SELECT FYPERIODS " _
          & " From GlfyTable Where (FYYEAR = " & cmbYer & ")"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoPer)
   
   If bSqlRows And Val(rdoPer!FYPERIODS) <> 0 Then
      With rdoPer
         For i = 0 To Val(!FYPERIODS) - 1
            optPer(i).enabled = True
            optPer(i).Value = Val(Mid(sOptions, i + 1, 1))
         Next
      End With
   Else
      ClearPer False
   End If
   
   Set rdoPer = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "LoadPeriods"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub PrintReport()
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    'Dim aRptPara As New Collection
    'Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection

    Dim sInclude As String
    Dim sAccount As String
    Dim i As Integer
    'Dim X As Integer
    
    On Error GoTo DiaErr1
    
    sCustomReport = GetCustomReport("fingl09.rpt")
    'SetMdiReportsize MdiSect
    'MdiSect.crw.ReportFileName = sReportPath & sCustomReport
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = sCustomReport
    
    If Trim(cmbYer) = "" Then
       MsgBox "Enter a Valid Fiscal Year.", vbInformation, Caption
       Exit Sub
    Else
       If Left(lblDsc, 10) = "*** Accoun" Then
          MsgBox "Enter a Valid Account.", vbInformation, Caption
          Exit Sub
       Else
          sSql = "{BdgtTable.BUDFY} = " & cmbYer
          If Trim(cmbAct) <> "" Then
             sAccount = Trim(cmbAct)
             sSql = sSql & " and  {GlacTable.GLACCTREF} = '" & Compress(cmbAct) & "'"
          Else
             sAccount = "ALL"
          End If
       End If
    End If
    
'    If Trim(cmbYer) = "" Then
'       MsgBox "Enter a Valid Fiscal Year.", vbInformation, Caption
'       Exit Sub
'    Else
'       If Left(lblDsc, 10) = "*** Accoun" Then
'          MsgBox "Enter a Valid Account.", vbInformation, Caption
'          Exit Sub
'       Else
'          sSql = "{BdgtTable.BUDFY} = " & cmbYer
'          If Trim(cmbAct) <> "" Then
'             sAccount = Trim(cmbAct)
'             sSql = sSql & " and  {GlacTable.GLACCTREF} = '" & Compress(cmbAct) & "'"
'          Else
'             sAccount = "ALL"
'          End If
'       End If
'    End If
    
'    X = 1
'    For i = 0 To 12
'       If optPer(i).Value = 1 Then
'          MdiSect.crw.Formulas(X - 1) = "Per" & X _
'                               & " = cdbl(right(cstr({BdgtTable.BUDPER" & lblPer(i) _
'                               & "}),len(cstr({BdgtTable.BUDPER" & lblPer(i) & "}))-1))"
'          MdiSect.crw.Formulas(X + 12) = "lbl" & X & " = ' Period " & lblPer(i) & "'"
'          sInclude = sInclude & lblPer(i) & ", "
'          X = X + 1
'       End If
'    Next
'
'    Do While X <= 13
'       MdiSect.crw.Formulas(X - 1) = "Per" & X & "=''"
'       X = X + 1
'    Loop
'
    
    'X = 1
    For i = 0 To 12
       If optPer(i).Value = 1 Then
'          MdiSect.crw.Formulas(i) = "Per" & CStr(i + 1) _
'                               & " = cdbl(right(cstr({BdgtTable.BUDPER" & lblPer(i) _
'                               & "}),len(cstr({BdgtTable.BUDPER" & lblPer(i) & "}))-1))"
            aFormulaName.Add "Per" & CStr(i + 1)
            aFormulaValue.Add "cdbl(right(cstr({BdgtTable.BUDPER" & lblPer(i) _
                               & "}),len(cstr({BdgtTable.BUDPER" & lblPer(i) & "}))-1))"
       End If
    Next
    
    For i = 0 To 12
       If optPer(i).Value = 1 Then
          'MdiSect.crw.Formulas(i + 13) = "lbl" & CStr(i + 1) & " = ' Period " & lblPer(i) & "'"
          aFormulaName.Add "lbl" & CStr(i + 1)
          aFormulaValue.Add "' Period " & lblPer(i) & "'"
          sInclude = sInclude & lblPer(i) & ", "
       End If
    Next
    
'    Do While X <= 13
'       MdiSect.crw.Formulas(X - 1) = "Per" & X & "=''"
'       X = X + 1
'    Loop
    
    If Trim(sInclude) <> "" Then
       sInclude = Left(sInclude, Len(sInclude) - 2)
    End If
    
'    MdiSect.crw.Formulas(26) = "Title2 = 'Include Periods: " & sInclude & "'"
'    MdiSect.crw.Formulas(27) = "CompanyName='" & sFacility & "'"
'    MdiSect.crw.Formulas(28) = "RequestedBy='Requested By: " & sInitials & "'"
'    MdiSect.crw.Formulas(29) = "Title1='Budgets For " & cmbYer & "'"
'    MdiSect.crw.Formulas(30) = "Account= 'Account: " & sAccount & "'"
    
    'MdiSect.crw.Formulas(26) = "Title2 = 'Include Periods: " & sInclude & "'"
    aFormulaName.Add "Title2"
    aFormulaValue.Add "'Include Periods: " & sInclude & "'"
    
    'MdiSect.crw.Formulas(27) = "CompanyName='" & sFacility & "'"
    aFormulaName.Add "CompanyName"
    aFormulaValue.Add "'" & sFacility & "'"
    
    'MdiSect.crw.Formulas(28) = "RequestedBy='Requested By: " & sInitials & "'"
    aFormulaName.Add "RequestedBy"
    aFormulaValue.Add "'Requested By: " & sInitials & "'"
    
    'MdiSect.crw.Formulas(29) = "Title1='Budgets For " & cmbYer & "'"
    aFormulaName.Add "Title1"
    aFormulaValue.Add "'Budgets For " & cmbYer & "'"
    
    'MdiSect.crw.Formulas(30) = "Account= 'Account: " & sAccount & "'"
    aFormulaName.Add "Account"
    aFormulaValue.Add "'Account: " & sAccount & "'"
    
    sInclude = "Include Inactive Accounts? "
    If optIna.Value = vbUnchecked Then
       sSql = sSql & " and {GlacTable.GLINACTIVE} = 0 "
       sInclude = sInclude & "N"
    Else
       sInclude = sInclude & "Y"
    End If
    'MdiSect.crw.Formulas(31) = "Title3='" & sInclude & "'"
    aFormulaName.Add "Title3"
    aFormulaValue.Add "'" & sInclude & "'"
    
    If Trim(cmbDiv) <> "" Then
       sSql = sSql & " and (RIGHT(LEFT({GlacTable.GLACCTNO} & '            ', " _
              & iEnd & "), " & iEnd & " - (" & iStart & " - 1)) = '" & cmbDiv & "')"
       sInclude = Trim(cmbDiv)
    Else
       sInclude = "ALL"
    End If
    'MdiSect.crw.Formulas(32) = "Division='Division: " & sInclude & "'"
    aFormulaName.Add "Division"
    aFormulaValue.Add "'Division: " & sInclude & "'"
    
    'MdiSect.crw.SelectionFormula = sSql
    'SetCrystalAction Me
    
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    cCRViewer.CRViewerSize Me
     ' Set report parameter
    cCRViewer.SetDbTableConnection
     ' print the copies
    cCRViewer.SetReportSelectionFormula sSql
    cCRViewer.OpenCrystalReportObject Me, aFormulaName
    cCRViewer.ShowGroupTree False
     
    'cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aFormulaName
    cCRViewer.ClearFieldCollection aFormulaValue
    
    optDis.enabled = True
    optPrn.enabled = True
    
    MouseCursor 0
    Exit Sub
   
DiaErr1:
    sProcName = "PrintReport"
    CurrError.Number = Err.Number
    CurrError.Description = Err.Description
    DoModuleErrors Me
End Sub


Private Sub PrintReport_Old()
   Dim sCustomReport As String
   Dim sInclude As String
   Dim sAccount As String
   Dim i As Integer
   Dim X As Integer
   
   On Error GoTo DiaErr1
   
   sCustomReport = GetCustomReport("fingl09.rpt")
   'SetMdiReportsize MdiSect
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   
   
   If Trim(cmbYer) = "" Then
      MsgBox "Enter a Valid Fiscal Year.", vbInformation, Caption
      Exit Sub
   Else
      If Left(lblDsc, 10) = "*** Accoun" Then
         MsgBox "Enter a Valid Account.", vbInformation, Caption
         Exit Sub
      Else
         sSql = "{BdgtTable.BUDFY} = " & cmbYer
         If Trim(cmbAct) <> "" Then
            sAccount = Trim(cmbAct)
            sSql = sSql & " and  {GlacTable.GLACCTREF} = '" & Compress(cmbAct) & "'"
         Else
            sAccount = "ALL"
         End If
      End If
   End If
   
   X = 1
   For i = 0 To 12
      If optPer(i).Value = 1 Then
         MdiSect.crw.Formulas(X - 1) = "Per" & X _
                              & " = cdbl(right(cstr({BdgtTable.BUDPER" & lblPer(i) _
                              & "}),len(cstr({BdgtTable.BUDPER" & lblPer(i) & "}))-1))"
         MdiSect.crw.Formulas(X + 12) = "lbl" & X & " = ' Period " & lblPer(i) & "'"
         sInclude = sInclude & lblPer(i) & ", "
         X = X + 1
      End If
   Next
   
   Do While X <= 13
      MdiSect.crw.Formulas(X - 1) = "Per" & X & "=''"
      X = X + 1
   Loop
   
   If Trim(sInclude) <> "" Then
      sInclude = Left(sInclude, Len(sInclude) - 2)
   End If
   
   MdiSect.crw.Formulas(26) = "Title2 = 'Include Periods: " & sInclude & "'"
   MdiSect.crw.Formulas(27) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(28) = "RequestedBy='Requested By: " & sInitials & "'"
   MdiSect.crw.Formulas(29) = "Title1='Budgets For " & cmbYer & "'"
   MdiSect.crw.Formulas(30) = "Account= 'Account: " & sAccount & "'"
   
   sInclude = "Include Inactive Accounts? "
   If optIna.Value = vbUnchecked Then
      sSql = sSql & " and {GlacTable.GLINACTIVE} = 0 "
      sInclude = sInclude & "N"
   Else
      sInclude = sInclude & "Y"
   End If
   MdiSect.crw.Formulas(31) = "Title3='" & sInclude & "'"
   
   If Trim(cmbDiv) <> "" Then
      sSql = sSql & " and (RIGHT(LEFT({GlacTable.GLACCTNO} & '            ', " _
             & iEnd & "), " & iEnd & " - (" & iStart & " - 1)) = '" & cmbDiv & "')"
      sInclude = Trim(cmbDiv)
   Else
      sInclude = "ALL"
   End If
   MdiSect.crw.Formulas(32) = "Division='Division: " & sInclude & "'"
   
   
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


Private Function bValidElement(cmbCombo As ComboBox) As Boolean
   Dim i As Integer
   On Error GoTo DiaErr1
   
   If cmbCombo.ListCount > 0 Then
      For i = 0 To cmbCombo.ListCount - 1
         If Val(cmbCombo.List(i)) = Val(cmbCombo.Text) Then
            bValidElement = True
            cmbCombo.ListIndex = i
         End If
      Next
   End If
   Exit Function
   
DiaErr1:
   sProcName = "ValidElement"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Function reParseOptions() As String
   Dim i As Integer
   Dim sDummy As String
   On Error GoTo DiaErr1
   
   For i = 0 To 12
      If optPer(i).enabled = True Then
         sDummy = sDummy & CStr(Val(optPer(i)))
      Else
         sDummy = sDummy & CStr(Val(Mid(sOptions, i + 1, 1)))
      End If
   Next
   sDummy = sDummy & CStr(optIna.Value)
   reParseOptions = sDummy
   
   Exit Function
   
DiaErr1:
   sProcName = "reParseOption"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Function bDivisionAccounts(iStart As Integer, iEnd As Integer) As Boolean
   Dim RdoDiv As ADODB.Recordset
   On Error GoTo DiaErr1
   
   sSql = "SELECT COGLDIVISIONS, COGLDIVSTARTPOS, COGLDIVENDPOS FROM ComnTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDiv)
   If bSqlRows Then
      With RdoDiv
         If Val("" & !COGLDIVISIONS) <> 0 Then
            If Val(!COGLDIVSTARTPOS) <> 0 And Val(!COGLDIVENDPOS) <> 0 Then
               iStart = Val(!COGLDIVSTARTPOS)
               iEnd = Val(!COGLDIVENDPOS)
               bDivisionAccounts = True
            End If
         End If
      End With
   End If
   Set RdoDiv = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "bDivisionAccounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function
