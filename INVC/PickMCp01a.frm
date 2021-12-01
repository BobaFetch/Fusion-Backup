VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form PickMCp01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Individual Pick List"
   ClientHeight    =   4320
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   7485
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   HelpContextID   =   5201
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PickMCp01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   37
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optLots 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   3240
      Width           =   735
   End
   Begin VB.CheckBox optSpc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   3480
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbRev 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   6240
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Revision-Select From List"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CheckBox optFrom 
      Caption         =   "from"
      Height          =   255
      Left            =   480
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6360
      TabIndex        =   28
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "PickMCp01a.frx":07AE
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
         Picture         =   "PickMCp01a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CheckBox optSht 
      Caption         =   "____"
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   5280
      TabIndex        =   9
      Top             =   4440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optCmt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   3000
      Width           =   735
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   2760
      Width           =   735
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   2520
      Width           =   735
   End
   Begin VB.CheckBox optCan 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   4440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbRun 
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   6240
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Select Run Number"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Part Number"
      Top             =   1080
      Width           =   3545
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6360
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6840
      Top             =   2640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4320
      FormDesignWidth =   7485
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   2400
      TabIndex        =   38
      Top             =   3840
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblRun 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   5040
      TabIndex        =   36
      Top             =   600
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Label lblMon 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   2040
      TabIndex        =   35
      Top             =   600
      Visible         =   0   'False
      Width           =   2892
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Lots"
      Height          =   255
      Index           =   13
      Left            =   240
      TabIndex        =   34
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Building Pick List"
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   33
      Top             =   3840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6840
      TabIndex        =   32
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Use AWJ Special"
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   31
      Top             =   3480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PL Rev"
      Height          =   255
      Index           =   10
      Left            =   5400
      TabIndex        =   30
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label lblPck 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4080
      TabIndex        =   27
      Top             =   1800
      Width           =   950
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Pick"
      Height          =   255
      Index           =   9
      Left            =   3240
      TabIndex        =   26
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shortage List"
      Height          =   252
      Index           =   8
      Left            =   3360
      TabIndex        =   25
      Top             =   4440
      Visible         =   0   'False
      Width           =   2052
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pick Comments"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   24
      Top             =   3000
      Width           =   2050
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   23
      Top             =   2760
      Width           =   2050
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   22
      Top             =   2520
      Width           =   2050
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Canceled Items"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   21
      Top             =   4440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   20
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblDte 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   19
      Top             =   1800
      Width           =   950
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Printed"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   18
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status  "
      Height          =   255
      Index           =   15
      Left            =   5760
      TabIndex        =   17
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   16
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label lblStat 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6840
      TabIndex        =   15
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   1
      Left            =   5760
      TabIndex        =   14
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   1080
      Width           =   1335
   End
End
Attribute VB_Name = "PickMCp01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'11/13/02 Added PKRECORD for new index
'2/9/05 Removed "Show Canceled Items"
'2/16/05 Added exclusion for Part Type 8
'4/18/05 Corrected Response byte to allow printing
'4/5/06 Corrected bad jump from Revise an MO
'2/21/07 Added Phantom (INVC 7.2.0/PROD 7.2.6)

Option Explicit
Dim rdoQry As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter

Dim bCanceled As Byte
Dim bGoodMo As Byte
Dim bGoodRuns As Byte
Dim bOnLoad As Byte
Dim bFromRev As Byte
Dim bPrinted As Byte

Dim dPkStart As Date

Dim iOldRun As Integer
Dim iPkRecord As Integer
Dim iRunNo As Integer

Dim cRunqty As Currency

Dim sBomRev As String
Dim sPartNumber As String
Dim sPrntPrt As String
Dim sOldPart As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd

Private Sub FormatControls()
   Dim bByte As Byte
   ReDim txtKeyPress(2) As New EsiKeyBd
   On Error Resume Next
   Set txtKeyPress(0).esCmbKeyCase = cmbPrt
   Set txtKeyPress(1).esCmbKeyValue = cmbRun
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   'Save by Menu Option
   sOptions = Trim(str(optCan.Value)) _
              & Trim(str(optDsc.Value)) _
              & Trim(str(optExt.Value)) _
              & Trim(str(OptCmt.Value)) _
              & Trim(str(optSht.Value))
   SaveSetting "Esi2000", "EsiProd", "ma01", Trim(sOptions)
   SaveSetting "Esi2000", "EsiProd", "ma01a", Trim(optLots.Value)
   
End Sub

Private Sub GetOptions()
   Dim bByte As Byte
   Dim sOptions As String
   'Get By Menu Option
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "ma01", sOptions)
   If Len(sOptions) > 0 Then
      optCan.Value = Val(Left(sOptions, 1))
      optDsc.Value = Val(Mid(sOptions, 2, 1))
      optExt.Value = Val(Mid(sOptions, 3, 1))
      OptCmt.Value = Val(Mid(sOptions, 4, 1))
      optSht.Value = Val(Mid(sOptions, 5, 1))
   End If
   bByte = CheckLotTracking()
   If bByte = 0 Then
      z1(13).Enabled = False
      optLots.Value = vbUnchecked
      optLots.Enabled = False
   Else
      optLots.Value = Val(GetSetting("Esi2000", "EsiProd", "ma01a", Trim(optLots.Value)))
   End If
   
End Sub

Private Sub cmbPrt_Click()
   prg1.Visible = False
   z1(12).Visible = False
   GetType
   bGoodRuns = GetRuns()
   
End Sub


Private Sub cmbPrt_GotFocus()
   SelectFormat Me
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If bCanceled Then Exit Sub
   
   If (Not ValidPartNumber(cmbPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If
   
   If cmbPrt <> "" Then
      GetType
      bGoodRuns = GetRuns()
      GetRevisions
   Else
      cmbRev.Clear
   End If
   optPrn.Enabled = True
   optDis.Enabled = True
   
End Sub


Private Sub cmbRev_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   cmbRev = UCase(CheckLen(cmbRev, 4))
   For iList = 0 To cmbRev.ListCount - 1
      If Trim(cmbRev) = Trim(cmbRev.List(iList)) Then b = 1
   Next
   If b = 0 And cmbRev.ListCount > 0 Then
      Beep
      cmbRev = cmbRev.List(0)
   End If
   
End Sub


Private Sub cmbRun_Click()
   optPrn.Enabled = True
   optDis.Enabled = True
   bGoodMo = GetThisRun()
   
End Sub

Private Sub cmbRun_GotFocus()
    cmbRun.SelStart = 0
    cmbRun.SelLength = Len(cmbRun.Text)
End Sub

Private Sub cmbRun_LostFocus()
   cmbRun = CheckLen(cmbRun, 5)
   cmbRun = Format(Abs(Val(cmbRun)), "####0")
   bGoodMo = GetThisRun
   
End Sub


Private Sub cmdCan_Click()
   If bFromRev = 1 Then
      On Error Resume Next
      ShopSHe02a.lblFrom = cmbPrt
      ShopSHe02a.optPick.Value = vbChecked
      ShopSHe02a.cmbPrt = cmbPrt
      ShopSHe02a.cmbRun = cmbRun
      ShopSHe02a.txtQty.SetFocus
      ShopSHe02a.Show
      optFrom.Value = vbUnchecked
   Else
      Unload Me
   End If
   
   
End Sub



Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCanceled = 1
   If bFromRev = 0 Then cmbPrt = ""
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 5201
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub



Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = 0
      FillCombo
      optCan.Value = vbUnchecked
      If optFrom.Value = vbUnchecked Then
         'If cUR.CurrentPart <> "" Then cmbPrt = cUR.CurrentPart
         If sPassedMo <> "" Then cmbPrt = sPassedMo
      Else
         iOldRun = Val(cmbRun)
         sOldPart = cmbPrt
         bGoodMo = GetThisRun()
         If bGoodMo Then GetRevisions
         bFromRev = 1
      End If
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sSql = "SELECT RUNREF,RUNSTATUS,RUNNO,RUNPLDATE FROM " _
          & "RunsTable WHERE RUNREF = ? " _
          & "AND RUNSTATUS<>'CA'"
   Set rdoQry = New ADODB.Command
   rdoQry.CommandText = sSql
   
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adChar
   AdoParameter1.SIZE = 30
   
   rdoQry.Parameters.Append AdoParameter1
   
   sPrntPrt = ""
   sCustomReport = GetCustomReport("prdma01")
   If sCustomReport = "awima01.rpt" Then
      optSpc.Visible = True
      z1(11).Visible = True
   Else
      optSpc.Visible = False
      z1(11).Visible = False
   End If
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
   SaveCurrentSelections
   bFromRev = 0
   Set rdoQry = Nothing
   If optFrom.Value = vbChecked Then
      ShopSHe02a.lblStat = lblStat
      ShopSHe02a.Show
   Else
      FormUnload
   End If
   Set PickMCp01a = Nothing
   
End Sub



Private Sub FillCombo()
   On Error GoTo DiaErr1
   DoEvents
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,PADESC,RUNREF " _
          & "FROM PartTable,RunsTable WHERE (PARTREF=RUNREF " _
          & "AND RUNSTATUS<>'CA') ORDER BY PARTREF "
   LoadComboBox cmbPrt
   If Val(lblRun) > 0 Then
      cmbPrt = lblMon
      cmbRun = lblRun
   End If
   If Trim(cmbPrt) <> "" Then
      bGoodRuns = GetRuns()
      GetType
      sOldPart = cmbPrt
   Else
      MsgBox "No Qualifying Manufacturing Orders Available.", _
         vbInformation, Caption
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Function GetRuns() As Byte
   Dim RdoRns As ADODB.Recordset
   
   If sOldPart <> cmbPrt Then
      cmbRun.Clear
      sOldPart = cmbPrt
   Else
      GetRuns = 1
      Exit Function
   End If
   sOldPart = cmbPrt
   sPartNumber = Compress(cmbPrt)
   On Error GoTo DiaErr1
   'RdoQry(0) = sPartNumber
   rdoQry.Parameters(0).Value = sPartNumber
   bSqlRows = clsADOCon.GetQuerySet(RdoRns, rdoQry)
   If bSqlRows Then
      With RdoRns
         Do Until .EOF
            AddComboStr cmbRun.hwnd, Format$(!Runno, "####0")
            .MoveNext
         Loop
         ClearResultSet RdoRns
      End With
      GetRuns = 1
   Else
      sPartNumber = ""
      GetRuns = 0
   End If
   If optFrom.Value = vbUnchecked Then
      If cmbRun.ListCount > 0 Then cmbRun = cmbRun.List(0)
   Else
      If Val(lblRun) > 0 Then cmbRun = lblRun
   End If
   bGoodMo = GetThisRun()
   Set RdoRns = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function GetThisRun() As Byte
   Dim RdoPln As ADODB.Recordset
   
   sPartNumber = Compress(cmbPrt)
   If Val(cmbRun) = 0 Then Exit Function
   On Error GoTo DiaErr1
   If Val(cmbRun) = 0 And optFrom.Value = vbChecked Then cmbRun = iOldRun
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,PADESC,PABOMREV,RUNREF," _
          & "RUNSTATUS,RUNPLDATE,RUNPKSTART,RUNQTY FROM PartTable,RunsTable " _
          & "WHERE PARTREF=RUNREF " _
          & "AND PARTREF='" & sPartNumber & "' AND RUNNO=" & Val(cmbRun) & " " _
          & "AND RUNSTATUS<>'CA'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPln, ES_STATIC)
   If bSqlRows Then
      With RdoPln
         cmbPrt = "" & Trim(!PartNum)
         sPassedMo = cmbPrt
         lblDsc = "" & Trim(!PADESC)
         lblStat = "" & !RUNSTATUS
         lblDte = "" & Format(!RUNPLDATE, "mm/dd/yyyy")
         lblPck = "" & Format(!RUNPKSTART, "mm/dd/yyyy")
         sBomRev = "" & Trim(!PABOMREV)
         cmbRev = sBomRev
         cRunqty = Format(!RUNQTY, ES_QuantityDataFormat)
         sBomRev = Compress(sBomRev)
         If Left(lblStat, 1) = "S" Or Left(lblStat, 1) = "R" Then
            cmbRev.Enabled = True
         Else
            cmbRev.Enabled = False
         End If
         ClearResultSet RdoPln
      End With
      GetThisRun = True
   Else
      GetThisRun = False
      cRunqty = 0
      sPartNumber = ""
      ClearBoxes
   End If
   Set RdoPln = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getthisru"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub ClearBoxes()
   lblStat = ""
   lblDte = ""
   lblDsc = ""
   lblPck = ""
   cmbRev.Clear
   
End Sub




Private Sub lblMon_Change()
   cmbPrt = lblMon
   
End Sub

Private Sub lblRun_Change()
   cmbRun = lblRun
   
End Sub

Private Sub optCan_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub

Private Sub optCan_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optCmt_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optCmt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   Dim bByte As Byte
   bByte = CheckList()
   '    If bByte = 1 Then
   '        optPrn.Enabled = False
   '        optDis.Enabled = False
   '    End If
   
End Sub

Private Sub optDsc_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optExt_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optExt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optFrom_Click()
   'Never visible from ShopSHe02a
   
End Sub

Private Sub optLots_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   Dim bByte As Byte
   bByte = CheckList(1)
   'If bByte = 1 Then optPrn.Enabled = False
'   If bByte = 1 Then
'      optPrn.Enabled = False
'      optDis.Enabled = False
'   End If
   
End Sub

Private Sub optSht_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optSht_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub PrintReport()

   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection

   MouseCursor 13
   On Error GoTo Pma01Pr
   'SetMdiReportsize MDISect
   prg1.Visible = False
   z1(12).Visible = False
   If sPrntPrt = "" Then sPrntPrt = sPartNumber
'   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MDISect.Crw.Formulas(1) = "ShowDescription=" & optDsc.value
'   MDISect.Crw.Formulas(2) = "ShowExtendedDescription=" & optExt.value
'   MDISect.Crw.Formulas(3) = "ShowPickComments=" & optCmt.value
'   MDISect.Crw.Formulas(4) = "ShowLots=" & optLots.value

   sCustomReport = GetCustomReport("prdma01")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport

    aFormulaName.Add "CompanyName"
    aFormulaName.Add "ShowDescription"
    aFormulaName.Add "ShowExtendedDescription"
    aFormulaName.Add "ShowPickComments"
    aFormulaName.Add "ShowLots"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add optDsc.Value
    aFormulaValue.Add optExt.Value
    aFormulaValue.Add OptCmt.Value
    aFormulaValue.Add optLots.Value
    
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
  ' MDISect.Crw.ReportFileName = sReportPath & sCustomReport
'   If optDsc.Value = vbUnchecked Then
'      MdiSect.Crw.SectionFormat(0) = "GROUPHDR.1.0;F;;;"
'   Else
'      MdiSect.Crw.SectionFormat(0) = "GROUPHDR.1.0;T;;;"
'   End If
'   If optExt.Value = vbUnchecked Then
'      MdiSect.Crw.SectionFormat(1) = "GROUPHDR.1.1;F;;;"
'   Else
'      MdiSect.Crw.SectionFormat(1) = "GROUPHDR.1.1;T;;;"
'   End If
'   If OptCmt.Value = vbUnchecked Then
'      MdiSect.Crw.SectionFormat(2) = "GROUPFTR.1.1;F;;;"
'   Else
'      MdiSect.Crw.SectionFormat(2) = "GROUPFTR.1.1;T;;;"
'   End If
'   If sCustomReport <> "awima01" Then
'      If optLots.Value = vbUnchecked Then
'         MdiSect.Crw.SectionFormat(3) = "DETAIL.0.0;F;;;"
'      Else
'         MdiSect.Crw.SectionFormat(3) = "DETAIL.0.0;T;;;"
'      End If
'   End If
   sSql = "{RunsTable.RUNREF} = '" & sPrntPrt & "' " _
          & "AND {RunsTable.RUNNO}=" & Trim(str(iRunNo)) & " "
   ' If optCan = vbUnchecked Then sSql = sSql & "AND {MopkTable.PKTYPE}<>12"
   'MDISect.Crw.SelectionFormula = sSql
   'SetCrystalAction Me
   
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.SetDbTableConnection
   
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName

   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   bPrinted = 1
   sPrntPrt = ""
   
   MouseCursor 0
   Exit Sub
   
Pma01Pr:
   bPrinted = 0
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   sPrntPrt = ""
   DoModuleErrors Me
   
End Sub



'
'Private Sub AddPickList()
'   '================================
'   'Phantoms installed 2/21/07
'   '================================
'   Dim RdoLst As rdoResultset
'   Dim bGoodHeader As Byte
'   Dim bGoodPl As Byte
'   Dim bResponse As Byte
'
'   Dim iRow As Integer
'   Dim iTotalItems As Integer
'   Dim n As Integer
'
'   Dim dDate As Date
'
'   Dim cQuantity As Currency
'   Dim cConversion As Currency
'   Dim cSetup As Currency
'   Dim sMsg As String
''   Dim sAssyPart As String
''   Dim sAssyRev As String
''   Dim bPhantom(300) As Byte
''   Dim vPickList(300, 6) As Variant
'   '0 = BMASSYPART
'   '1 = BMPARTREF
'   '2 = Quantity
'   '3 = BMCOMT (Comment)
'   '4 = BMUNITS
'   '5 = BMREV
'
'   On Error GoTo DiaErr2
'
'   sPrntPrt = sPartNumber
'   sBomRev = Trim(cmbRev)
'   iPkRecord = 0
'   dDate = Format(ES_SYSDATE, "mm/dd/yy")
'   If Len(Trim(lblPck)) > 0 Then
'      dPkStart = lblPck
'   Else
'      dPkStart = dDate
'   End If
'
'   sMsg = "This Run Is Scheduled (SC) Or (RL)." & vbCr _
'          & "Do You Want To Create The Pick List " & vbCr _
'          & "And Release The MO?"
'   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
'   If bResponse = vbYes Then
'      MouseCursor 13
'      n = 10
'
'      'determine whether any part list for this part
'      sSql = "SELECT DISTINCT BMASSYPART FROM BmplTable " _
'             & "WHERE BMASSYPART='" & sPartNumber & "' "
'      bSqlRows = GetDataSet(RdoLst, ES_FORWARD)
'      If bSqlRows Then
'         prg1.Visible = True
'         z1(12).Visible = True
'         z1(12).Refresh
'         prg1.Value = n
'         bGoodPl = True
'         'On Error Resume Next
'         ClearResultSet RdoLst
'      Else
'         MouseCursor 0
'         bGoodPl = False
'         MsgBox "This Part Does Not Have A Parts List.", vbInformation, Caption
'         cmbPrt.SetFocus
'         Exit Sub
'      End If
''      If Not bGoodPl Then
''         'On Error Resume Next
''         cmbPrt.SetFocus
''         Exit Sub
''      End If
'      'found one, let's go on
'      sSql = "Select BMHREF FROM BmhdTable WHERE BMHREF='" & sPartNumber & "' " _
'             & "AND BMHREV='" & Trim(cmbRev) & "'"
'      bSqlRows = GetDataSet(RdoLst, ES_KEYSET)
'      iRow = 0
'      If bSqlRows Then
'         With RdoLst
'            Do Until .EOF
'               iRow = iRow + 1
'               .MoveNext
'            Loop
'            ClearResultSet RdoLst
'         End With
'      End If
'
'      'See how many...one?...use it..else find one
'      If iRow > 0 Then
'         If iRow = 1 Then
'            'There is one, but no revisions
'            sBomRev = ""
'            bGoodHeader = True
'         Else
'            'get the first rev that matches
'            sSql = "SELECT BMHREF,BMHREV,BMHOBSOLETE,BMHRELEASED,BMHEFFECTIVE " _
'                   & "FROM BmhdTable WHERE BMHREF='" & sPartNumber & "' AND BHMREV='" & sBomRev & "' " _
'                   & "AND BMHOBSOLETE >='" & dDate & "' AND BMHRELEASED=1"
'            bSqlRows = GetDataSet(RdoLst, ES_FORWARD)
'            If bSqlRows Then
'               bGoodHeader = True
'            Else
'               bGoodHeader = False
'            End If
'            'On Error Resume Next
'            ClearResultSet RdoLst
'         End If
'      Else
'         bGoodHeader = False
'      End If
'      If Not bGoodHeader Then
'         'oops the header is gone, date invalid or not released?
'         MouseCursor 0
'         MsgBox "The Parts List Is Not Valid, Released, " & vbCr _
'            & "Or Outdated For This Part.", vbInformation, Caption
'         'On Error Resume Next
'         cmbPrt.SetFocus
'         Exit Sub
'      End If
'
'      'that's enough, let's build it
'      RdoCon.BeginTrans
'      sSql = "SELECT * FROM BmplTable" & vbCrLf _
'         & "WHERE BMASSYPART='" & sPartNumber & "'" & vbCrLf _
'         & "AND BMREV='" & sBomRev & "'" & vbCrLf _
'         & "ORDER BY BMSEQUENCE"
'      bSqlRows = GetDataSet(RdoLst, ES_STATIC)
'      iRow = -1
'      If bSqlRows Then
'         With RdoLst
'            Do Until .EOF
'               iRow = iRow + 1
'               n = n + 5
'               If n > 70 Then n = 70
'               prg1.Value = n
'               If Not IsNull(!BMSETUP) Then
'                  cSetup = !BMSETUP
'               Else
'                  cSetup = 0
'               End If
'               cQuantity = Format((cRunqty * (!BMQTYREQD + !BMADDER) + cSetup), "######0.000")
'               'cConversion = !BMCONVERSION
'               'If cConversion = 0 Then cConversion = 1
'
'               If !BMCONVERSION <> 0 Then
'                  cQuantity = cQuantity / !BMCONVERSION
'               End If
''               If Not IsNull(!BMPHANTOM) Then
''                  bPhantom(iRow) = !BMPHANTOM
''               Else
''                  bPhantom(iRow) = 0
''               End If
'
'               'if phantom item, then explode it
'               If !BMPHANTOM = 1 Then
'                  InsertPhantom Trim(!BMPARTREF), Trim(!BMPARTREV), cQuantity
'
'               'add non-phantom item to pick list
'               Else
''                  vPickList(iRow, 0) = "" & Trim(!BMASSYPART)
''                  vPickList(iRow, 1) = "" & Trim(!BMPARTREF)
''                  vPickList(iRow, 2) = Format(cQuantity / cConversion, ES_QuantityDataFormat)
''                  vPickList(iRow, 3) = "" & Trim(!BMCOMT)
''                  vPickList(iRow, 4) = "" & Trim(!BMUNITS)
''                  vPickList(iRow, 5) = "" & Trim(!BMREV)
'
'                  iPkRecord = iPkRecord + 1
''                  sSql = "INSERT INTO MopkTable (PKPARTREF,PKMOPART,PKMORUN," _
''                         & "PKTYPE,PKPDATE,PKPQTY,PKBOMQTY,PKRECORD,PKUNITS," _
''                         & "PKCOMT) VALUES('" & vPickList(iRow, 1) & "','" _
''                         & vPickList(iRow, 0) & "'," & cmbRun & ",9,'" & dPkStart _
''                         & "'," & cQuantity & "," & cQuantity & "," & iPkRecord _
''                         & ",'" & vPickList(iRow, 4) & "','" & vPickList(iRow, 3) & "') "
'                  sSql = "INSERT INTO MopkTable (PKPARTREF,PKMOPART,PKMORUN," _
'                         & "PKTYPE,PKPDATE,PKPQTY,PKBOMQTY,PKRECORD,PKUNITS," _
'                         & "PKCOMT) VALUES('" & Trim(!BMPARTREF) & "','" _
'                         & Compress(cmbPrt) & "'," & cmbRun & ",9,'" & dPkStart _
'                         & "'," & cQuantity & "," & cQuantity & "," & iPkRecord & "," _
'                         & "'" & Trim(!BMUNITS) & "','" & Trim(!BMCOMT) & "') "
'                  RdoCon.Execute sSql, rdExecDirect
'
'               End If
'
'               .MoveNext
'            Loop
'            ClearResultSet RdoLst
'         End With
''         iTotalItems = iRow
''         'On Error Resume Next
''         'got some. now build it
''         RdoCon.BeginTrans
''         For iRow = 0 To iTotalItems
''            n = n + 5
''            If n > 90 Then n = 90
''            prg1.Value = n
''            cQuantity = Format(vPickList(iRow, 2), ES_QuantityDataFormat)
''            If bPhantom(iRow) = 0 Then
''               iPkRecord = iPkRecord + 1
''               sSql = "INSERT INTO MopkTable (PKPARTREF,PKMOPART,PKMORUN," _
''                      & "PKTYPE,PKPDATE,PKPQTY,PKBOMQTY,PKRECORD,PKUNITS," _
''                      & "PKCOMT) VALUES('" & vPickList(iRow, 1) & "','" _
''                      & vPickList(iRow, 0) & "'," & cmbRun & ",9,'" & dPkStart _
''                      & "'," & cQuantity & "," & cQuantity & "," & iPkRecord _
''                      & ",'" & vPickList(iRow, 4) & "','" & vPickList(iRow, 3) & "') "
''               RdoCon.Execute sSql, rdExecDirect
''            Else
''               'Don't want to use ByVal for the variants
''               sAssyPart = vPickList(iRow, 1)
''               sAssyRev = vPickList(iRow, 5)
''               InsertPhantom sAssyPart, sAssyRev, cQuantity
''            End If
''         Next
'         sSql = "UPDATE RunsTable SET RUNSTATUS='PL'," _
'                & "RUNPLDATE='" & dDate & "' " _
'                & "WHERE RUNREF='" & sPartNumber & "' " _
'                & "AND RUNNO=" & cmbRun & " "
'         RdoCon.Execute sSql, rdExecDirect
'         MouseCursor 0
'         prg1.Value = 100
'         MouseCursor 0
''         If Err = 0 Then
'            RdoCon.CommitTrans
'            optPrn = True
'            MsgBox "Pick List Added Successfully.", vbInformation, Caption
'            lblStat = "PL"
''         Else
''            prg1.Visible = False
''            z1(12).Visible = False
''            prg1.Value = 0
''            RdoCon.RollbackTrans
''            MsgBox "Could Not Successfully Print The Pick List.", _
''               vbInformation, Caption
''         End If
'      Else
'         MouseCursor 0
'         MsgBox "Couldn't Find Items For Revision " & sBomRev & ".", vbInformation, Caption
'         cmbPrt.SetFocus
'      End If
'   Else
'      optPrn.Enabled = True
'      optDis.Enabled = True
'      CancelTrans
'   End If
'   prg1.Visible = False
'   z1(12).Visible = False
'   Set RdoLst = Nothing
'   Exit Sub
'
''DiaErr1:
''   CurrError.Description = Err.Description
''   Resume DiaErr2
'DiaErr2:
'   'On Error Resume Next
'   RdoCon.RollbackTrans
'   MouseCursor 0
'   prg1.Visible = False
'   z1(12).Visible = False
'   'MsgBox CurrError.Description & vbCr & "Can't Add Pick List.", vbExclamation, Caption
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub





''''''''''''''''''''''


Private Sub AddPickList(Optional bPrn As Byte)
   Dim RdoLst As ADODB.Recordset
   
   Dim bGoodHeader As Byte
   Dim bGoodPl As Byte
   Dim bResponse As Byte
   Dim bOrphanedParts As Byte
   
   Dim iRow As Integer
   Dim iTotalItems As Integer
   Dim n As Integer
   
   Dim dDate As Date
   
   Dim cQuantity As Currency
   Dim cConversion As Currency
   Dim cSetup As Currency
   Dim sMsg As String
   '0 = BMASSYPART
   '1 = BMPARTREF
   '2 = Quantity
   '3 = BMCOMT (Comment)
   '4 = BMUNITS
   '5 = BMREV
   
   On Error GoTo DiaErr2
   
   sPrntPrt = sPartNumber
   sBomRev = Trim(cmbRev)
   iPkRecord = 0
   dDate = Format(ES_SYSDATE, "mm/dd/yyyy")
   If Len(Trim(lblPck)) > 0 Then
      dPkStart = lblPck
   Else
      dPkStart = dDate
   End If
   
   sMsg = "This Run Is Scheduled (SC) Or (RL)." & vbCr _
          & "Do You Want To Create The Pick List " & vbCr _
          & "And Release The MO?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      n = 10
      
      'determine whether any part list for this part and rev
      sSql = "SELECT BMASSYPART FROM BmplTable " & vbCrLf _
             & "WHERE BMASSYPART = '" & sPartNumber & "'" & vbCrLf _
             & "AND BMREV = '" & sBomRev & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_FORWARD)
      If bSqlRows Then
         prg1.Visible = True
         z1(12).Visible = True
         z1(12).Refresh
         prg1.Value = n
         bGoodPl = True
         ClearResultSet RdoLst
         Set RdoLst = Nothing
      Else
         MouseCursor 0
         bGoodPl = False
         MsgBox "This part does not have a parts list rev " & sBomRev, vbInformation, Caption
         cmbPrt.SetFocus
         Exit Sub
      End If
      
      
'      sSql = "Select BMHREF FROM BmhdTable WHERE BMHREF='" & sPartNumber & "' " _
'             & "AND BMHREV='" & sBomRev & "'"
'      bSqlRows = GetDataSet(RdoLst, ES_KEYSET)
'      iRow = 0
'      If bSqlRows Then
'         With RdoLst
'            Do Until .EOF
'               iRow = iRow + 1
'               .MoveNext
'            Loop
'            ClearResultSet RdoLst
'         End With
'      End If
'
'      'See how many...one?...use it..else find one
'      If iRow > 0 Then
'         If iRow = 1 Then
'            'There is one, but no revisions
'            sBomRev = ""
'            bGoodHeader = True
'         Else
'            'get the first rev that matches


            sSql = "SELECT BMHREF,BMHREV,BMHOBSOLETE,BMHRELEASED,BMHEFFECTIVE " & vbCrLf _
                   & "FROM BmhdTable" & vbCrLf _
                   & "WHERE BMHREF='" & sPartNumber & "' AND BMHREV='" & sBomRev & "' " & vbCrLf _
                   & "AND (BMHOBSOLETE IS NULL OR BMHOBSOLETE >='" & dDate & "') AND BMHRELEASED=1"
            bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_FORWARD)
            If bSqlRows Then
               bGoodHeader = True
            Else
               bGoodHeader = False
            End If
'            'On Error Resume Next
'            ClearResultSet RdoLst
'         End If
'      Else
'         bGoodHeader = False
'      End If
      If Not bGoodHeader Then
         'oops the header is gone, date invalid or not released?
         MouseCursor 0
         MsgBox "The Parts List Is Not Valid, Released, " & vbCr _
            & "Or Outdated For This Part.", vbInformation, Caption
         cmbPrt.SetFocus
         
         
         Exit Sub
      End If
      
      'that's enough, let's build it
      Dim strComt As String
      
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      
      sSql = "SELECT PARTREF,PAUNITS, * FROM BmplTable" & vbCrLf _
         & "LEFT OUTER JOIN PartTable ON PARTREF=BMPARTREF " & vbCrLf _
         & "WHERE BMASSYPART='" & sPartNumber & "'" & vbCrLf _
         & "AND BMREV='" & sBomRev & "'" & vbCrLf _
         & "ORDER BY BMSEQUENCE"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_DYNAMIC)
      iRow = -1
      bOrphanedParts = 0
      If bSqlRows Then
         With RdoLst
            Do Until .EOF
               iRow = iRow + 1
               n = n + 5
               If n > 70 Then n = 70
               prg1.Value = n
               If Not IsNull(!BMSETUP) Then
                  cSetup = !BMSETUP
               Else
                  cSetup = 0
               End If
               If (SetupQtyEnabled = True) Then
                  cQuantity = Format(((cRunqty + cSetup) * (!BMQTYREQD + !BMADDER)), "######0.000")
               Else
                  cQuantity = Format((cRunqty * (!BMQTYREQD + !BMADDER) + cSetup), "######0.000")
               End If
               'cQuantity = Format(((cRunqty + cSetup) * (!BMQTYREQD + !BMADDER)), "######0.000")
               If !BMCONVERSION <> 0 Then
                  cQuantity = cQuantity / !BMCONVERSION
               End If

               'if phantom item, then explode it
               If !BMPHANTOM = 1 Then
                  InsertPhantom Trim(!BMPARTREF), Trim(!BMPARTREV), cQuantity
               
               'else add non-phantom item to pick list
               Else
                  iPkRecord = iPkRecord + 1
                  
                  If (Not IsNull(!BMCOMT)) Then
                    strComt = ReplaceSingleQuote(Trim(!BMCOMT))
                  Else
                    strComt = ""
                  End If
                  
                  sSql = "INSERT INTO MopkTable (PKPARTREF,PKMOPART,PKMORUN," _
                         & "PKTYPE,PKPDATE,PKPQTY,PKBOMQTY,PKRECORD,PKUNITS," _
                         & "PKCOMT) VALUES('" & Trim(!BMPARTREF) & "','" _
                         & Compress(cmbPrt) & "'," & cmbRun & ",9,'" & dPkStart _
                         & "'," & cQuantity & "," & cQuantity & "," & iPkRecord & "," _
                         & "'" & Trim(!PAUNITS) & "','" & strComt & "') "
                  If Len(Trim(!PartRef)) = 0 Then
                      bOrphanedParts = 1
                  Else
                      clsADOCon.ExecuteSQL sSql
                  End If
                  
               End If
               
               .MoveNext
            Loop
            ClearResultSet RdoLst
            Set RdoLst = Nothing
         End With
         sSql = "UPDATE RunsTable SET RUNSTATUS='PL'," _
                & "RUNPLDATE='" & dDate & "' " _
                & "WHERE RUNREF='" & sPartNumber & "' " _
                & "AND RUNNO=" & cmbRun & " "
         clsADOCon.ExecuteSQL sSql
         MouseCursor 0
         prg1.Value = 100
         MouseCursor 0
         clsADOCon.CommitTrans
         bPrinted = bPrn
         If bPrinted Then
           optPrn = True
           bPrn = 0
         Else
           optDis = True
         End If
         If bOrphanedParts Then
           MsgBox "Pick List Added Successfully. However, your BOM Parts List has Orphaned Parts." & vbCrLf & "Please Contact Fusion Support"
         'Else
           'MsgBox "Pick List Added Successfully.", vbInformation, Caption
         End If
         lblStat = "PL"
      Else
         MouseCursor 0
         MsgBox "Couldn't Find Items For Revision " & sBomRev & ".", vbInformation, Caption
         cmbPrt.SetFocus
      End If
   Else
      optPrn.Enabled = True
      optDis.Enabled = True
      CancelTrans
   End If
   prg1.Visible = False
   z1(12).Visible = False
   Set RdoLst = Nothing
   Exit Sub
   
DiaErr2:
   clsADOCon.RollbackTrans
   MouseCursor 0
   prg1.Visible = False
   z1(12).Visible = False
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub




Private Function CheckList(Optional bPrn As Byte) As Byte
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   sPartNumber = Compress(cmbPrt)
   bGoodMo = GetThisRun
   iRunNo = cmbRun
   sProcName = "checklist"
   If bGoodMo Then
      CheckList = 1
      If Trim(lblStat) = "SC" Or Trim(lblStat) = "RL" Then
         If Val(lblType) = 8 Then
            PrintReport
         Else
            sProcName = "addpickli"
            AddPickList (bPrn)
         End If
      Else
         PrintReport
      End If
   Else
      CheckList = 0
      sMsg = "Run Isn't Listed." & vbCr _
             & "Status Must Be SC,RL,PL,PP,CO,Cl."
      MsgBox sMsg, vbInformation, Caption
   End If
   Exit Function
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
   
End Function

Private Sub GetRevisions()
   cmbRev.Clear
   On Error GoTo DiaErr1
   sSql = "SELECT BMHREV FROM BmhdTable WHERE BMHREF='" _
          & Compress(cmbPrt) & "' ORDER BY BMHREV"
   LoadComboBox cmbRev, -1
   Exit Sub
   
DiaErr1:
   sProcName = "getrevisions"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub optSpc_KeyDown(KeyCode As Integer, Shift As Integer)
   KeyCheck KeyCode
   
End Sub


Private Sub optSpc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub



Private Sub GetType()
   On Error Resume Next
   Dim RdoTyp As ADODB.Recordset
   
   sSql = "SELECT PARTREF,PALEVEL FROM PartTable WHERE " _
          & "PARTREF='" & Compress(cmbPrt) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTyp, ES_FORWARD)
   If bSqlRows Then lblType = RdoTyp!PALEVEL Else lblType = ""
   ClearResultSet RdoTyp
   Set RdoTyp = Nothing
   
End Sub

'2/21/07

Public Sub InsertPhantom(AssyPart As String, AssyRev As String, AssyQuantity As Currency)
   Dim RdoPhn As ADODB.Recordset
   Dim iList As Integer
   Dim iTotalPhantom As Integer
   Dim cPQuantity As Currency
   Dim cPConversion As Currency
   Dim cPSetup As Currency
   
'   Dim vPhantomList(300, 5) As Variant
   
'   sSql = "SELECT * FROM BmplTable WHERE BMASSYPART='" & AssyPart & "' " _
'          & "AND BMREV='" & AssyRev & "' "
'   bSqlRows = GetDataSet(RdoPhn, ES_STATIC)
'   iList = -1
'   If bSqlRows Then
'      With RdoPhn
'         Do Until .EOF
'            iList = iList + 1
'            If Not IsNull(!BMSETUP) Then
'               cPSetup = !BMSETUP
'            Else
'               cPSetup = 0
'            End If
'            cPQuantity = Format((AssyQuantity * (!BMQTYREQD + !BMADDER) + cPSetup), "######0.000")
'            cPConversion = !BMCONVERSION
'            If cPConversion = 0 Then cPConversion = 1
'            vPhantomList(iList, 0) = "" & Trim(!BMASSYPART)
'            vPhantomList(iList, 1) = "" & Trim(!BMPARTREF)
'            vPhantomList(iList, 2) = Format(cPQuantity / cPConversion, ES_QuantityDataFormat)
'            vPhantomList(iList, 3) = "" & Trim(!BMCOMT)
'            vPhantomList(iList, 4) = "" & Trim(!BMUNITS)
'            .MoveNext
'         Loop
'         ClearResultSet RdoPhn
'      End With
'      iTotalPhantom = iList
'      For iList = 0 To iTotalPhantom
'         cPQuantity = Format(vPhantomList(iList, 2), ES_QuantityDataFormat)
'         iPkRecord = iPkRecord + 1
'         sSql = "INSERT INTO MopkTable (PKPARTREF,PKMOPART,PKMORUN," _
'                & "PKTYPE,PKPDATE,PKPQTY,PKBOMQTY,PKRECORD,PKUNITS," _
'                & "PKCOMT) VALUES('" & vPhantomList(iList, 1) & "','" _
'                & vPhantomList(iList, 0) & "'," & cmbRun & ",9,'" & dPkStart _
'                & "'," & cPQuantity & "," & cPQuantity & "," & iPkRecord _
'                & ",'" & vPhantomList(iList, 4) & "','" & vPhantomList(iList, 3) & "') "
'         RdoCon.Execute sSql, rdExecDirect
'      Next
'   End If
'
'   sSql = "SELECT * FROM BmplTable WHERE BMASSYPART='" & AssyPart & "' " _
'          & "AND BMREV='" & AssyRev & "' "
   sSql = "SELECT * FROM BmplTable" & vbCrLf _
      & "WHERE BMASSYPART='" & AssyPart & "'" & vbCrLf _
      & "AND BMREV='" & AssyRev & "'" & vbCrLf _
      & "ORDER BY BMSEQUENCE"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPhn, ES_STATIC)
   iList = -1
   If bSqlRows Then
      With RdoPhn
         Do Until .EOF
            iList = iList + 1
            If Not IsNull(!BMSETUP) Then
               cPSetup = !BMSETUP
            Else
               cPSetup = 0
            End If
            
            If (SetupQtyEnabled = True) Then
               cPQuantity = Format(((AssyQuantity + cPSetup) * (!BMQTYREQD + !BMADDER)), "######0.000")
            Else
               cPQuantity = Format(((AssyQuantity * (!BMQTYREQD + !BMADDER)) + cPSetup), "######0.000")
            End If
            'cPConversion = !BMCONVERSION
            'If cPConversion = 0 Then cPConversion = 1
            
            If !BMCONVERSION <> 0 Then
               cPQuantity = cPQuantity / !BMCONVERSION
            End If
            
            If !BMPHANTOM = 1 Then
               InsertPhantom Trim(!BMPARTREF), Trim(!BMPARTREV), cPQuantity
            Else
'               vPhantomList(iList, 0) = "" & Trim(!BMASSYPART)
'               vPhantomList(iList, 1) = "" & Trim(!BMPARTREF)
'   '            vPhantomList(iList, 2) = Format(cPQuantity / cPConversion, ES_QuantityDataFormat)
'               vPhantomList(iList, 2) = Format(cPQuantity, ES_QuantityDataFormat)
'               vPhantomList(iList, 3) = "" & Trim(!BMCOMT)
'               vPhantomList(iList, 4) = "" & Trim(!BMUNITS)
            
               iPkRecord = iPkRecord + 1
'               sSql = "INSERT INTO MopkTable (PKPARTREF,PKMOPART,PKMORUN," _
'                  & "PKTYPE,PKPDATE,PKPQTY,PKBOMQTY,PKRECORD,PKUNITS," _
'                  & "PKCOMT) VALUES('" & vPhantomList(iList, 1) & "','" _
'                  & vPhantomList(iList, 0) & "'," & cmbRun & ",9,'" & dPkStart _
'                  & "'," & cPQuantity & "," & cPQuantity & "," & iPkRecord _
'                  & ",'" & vPhantomList(iList, 4) & "','" & vPhantomList(iList, 3) & "') "
               
               
'               vPhantomList(iList, 0) = "" & Trim(!BMASSYPART)
'               vPhantomList(iList, 1) = "" & Trim(!BMPARTREF)
'   '            vPhantomList(iList, 2) = Format(cPQuantity / cPConversion, ES_QuantityDataFormat)
'               vPhantomList(iList, 2) = Format(cPQuantity, ES_QuantityDataFormat)
'               vPhantomList(iList, 3) = "" & Trim(!BMCOMT)
'               vPhantomList(iList, 4) = "" & Trim(!BMUNITS)
               sSql = "INSERT INTO MopkTable (PKPARTREF,PKMOPART,PKMORUN," _
                  & "PKTYPE,PKPDATE,PKPQTY,PKBOMQTY,PKRECORD,PKUNITS," _
                  & "PKCOMT) VALUES('" & Trim(!BMPARTREF) & "','" _
                  & Compress(cmbPrt) & "'," & cmbRun & ",9,'" & dPkStart _
                  & "'," & cPQuantity & "," & cPQuantity & "," & iPkRecord & "," _
                  & "'" & Trim(!BMUNITS) & "','" & Trim(!BMCOMT) & "') "
               clsADOCon.ExecuteSQL sSql
            
            End If
            .MoveNext
         Loop
         ClearResultSet RdoPhn
      End With
   End If
'      iTotalPhantom = iList
'      For iList = 0 To iTotalPhantom
'         cPQuantity = Format(vPhantomList(iList, 2), ES_QuantityDataFormat)
'         iPkRecord = iPkRecord + 1
'         sSql = "INSERT INTO MopkTable (PKPARTREF,PKMOPART,PKMORUN," _
'                & "PKTYPE,PKPDATE,PKPQTY,PKBOMQTY,PKRECORD,PKUNITS," _
'                & "PKCOMT) VALUES('" & vPhantomList(iList, 1) & "','" _
'                & vPhantomList(iList, 0) & "'," & cmbRun & ",9,'" & dPkStart _
'                & "'," & cPQuantity & "," & cPQuantity & "," & iPkRecord _
'                & ",'" & vPhantomList(iList, 4) & "','" & vPhantomList(iList, 3) & "') "
'         RdoCon.Execute sSql, rdExecDirect
'      Next
'   End If
   Set RdoPhn = Nothing
End Sub
