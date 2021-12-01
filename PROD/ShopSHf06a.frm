VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form ShopSHf06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Split A Manufacturing Order"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   4156
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHf06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optSplit 
      Enabled         =   0   'False
      Height          =   255
      Left            =   1440
      TabIndex        =   16
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox txtQty 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6840
      TabIndex        =   4
      ToolTipText     =   "Quantity To Split To New Manufacturing Order"
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdSel 
      Cancel          =   -1  'True
      Caption         =   "S&elect"
      Height          =   315
      Left            =   6840
      TabIndex        =   2
      ToolTipText     =   "Select This Manufacturing Order Run"
      Top             =   960
      Width           =   875
   End
   Begin VB.CommandButton cmdSplit 
      Caption         =   "S&plit"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6840
      TabIndex        =   5
      ToolTipText     =   "Split This Manufacturing Order"
      Top             =   2160
      Width           =   875
   End
   Begin VB.TextBox txtRun 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5640
      TabIndex        =   3
      ToolTipText     =   "New Run Number (Any Unused)"
      Top             =   2160
      Width           =   855
   End
   Begin VB.ComboBox cmbRun 
      Height          =   315
      Left            =   5640
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Contains Qualified Runs"
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Part Number (Contains Qualifying Manufacturing Orders)"
      Top             =   960
      Width           =   3545
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   6840
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
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
      FormDesignHeight=   3480
      FormDesignWidth =   7875
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   1440
      TabIndex        =   25
      Top             =   2760
      Visible         =   0   'False
      Width           =   3252
      _ExtentX        =   5741
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Split MO's"
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   23
      ToolTipText     =   "Company Setup Option"
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblPrg 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   2880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Index           =   8
      Left            =   5040
      TabIndex        =   21
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7200
      TabIndex        =   20
      ToolTipText     =   "Manufacturing Order Status"
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Split Percentage"
      Height          =   255
      Index           =   7
      Left            =   5040
      TabIndex        =   19
      ToolTipText     =   "Quantity To Split To New Manufacturing Order"
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label lblPerc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6840
      TabIndex        =   18
      ToolTipText     =   "Percentage Of The Remaining Quantity"
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Splits Are Allowed (Company Setup)"
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   6
      Left            =   1440
      TabIndex        =   17
      ToolTipText     =   "Company Setup Option"
      Top             =   480
      Width           =   4575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity To Split"
      Height          =   285
      Index           =   5
      Left            =   5040
      TabIndex        =   15
      ToolTipText     =   "Quantity To Split To New Manufacturing Order"
      Top             =   2565
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining Quantity"
      Height          =   285
      Index           =   4
      Left            =   5040
      TabIndex        =   14
      ToolTipText     =   "Remaining Quantity"
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lblRemQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6840
      TabIndex        =   13
      ToolTipText     =   "Remaining Quantity"
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   3
      Left            =   5040
      TabIndex        =   12
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Split To"
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label lblNew 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      ToolTipText     =   "New Manufacturing Order"
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current MO"
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "ShopSHf06a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'2/8/05 New
'2/15/05 Submitted for testing
'2/22/05 Notified Larry that the function was available
Option Explicit
Dim bOnLoad As Byte
Dim bSplitFailed As Byte
Dim bOutSideFailed As Byte
Dim bTablesFailed As Byte
Dim iDefaultRun As Integer
Dim cSplitPerc As Currency
Dim sOldPart As String
Dim newMORun As String
Dim loitColumns As String

'TempTables
Dim sTimeTemp As String
Dim sPickTemp As String
Dim sInvaTemp As String
Dim sLotsTemp As String
Dim sRunsTemp As String
Dim sRnopTemp As String
Dim sOutSideSrvTemp  As String
Dim sPoitSrvTemp As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbPrt_Change()
   lblPrg.Visible = False
   
End Sub

Private Sub cmbPrt_Click()
   cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
   lblNew = cmbPrt
   FillRuns
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
   lblNew = cmbPrt
   FillRuns
   
End Sub


Private Sub cmbRun_Click()
   txtRun = ""
   txtQty = ""
   lblPerc = ""
   lblRemQty = ""
   lblStatus = ""
   cmdSplit.Enabled = False
   txtRun.Enabled = False
   txtQty.Enabled = False
   
End Sub

Private Sub cmbRun_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   For iList = 0 To cmbRun.ListCount - 1
      If cmbRun.List(iList) = cmbRun Then bByte = 1
   Next
   If bByte = 0 Then
      Beep
      cmbRun = cmbRun.List(0)
   End If
   
End Sub


Private Sub cmdCan_Click()
   MouseCursor 13
   Unload Me
   
End Sub



Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4156
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdSel_Click()
   If cmbRun.ListCount = 0 Then
      MsgBox "Requires A Valid Run.", _
         vbInformation, Caption
      Exit Sub
   End If
   lblRemQty = Format(GetRemaining(), ES_QuantityDataFormat)
   txtRun = GetSplitHistory()
   If Val(lblRemQty) < 2 Then
      MsgBox "The Remaining Quantity May Not Be Split.", _
         vbInformation, Caption
      txtRun = ""
      lblPerc = ""
      txtQty = "0.000"
      Exit Sub
   End If
   If optSplit.Value = vbChecked Then
      cmdSplit.Enabled = True
      txtRun.Enabled = True
      txtQty.Enabled = True
   End If
   
End Sub

Private Sub cmdSplit_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   bResponse = CheckExistingRun()
   
   If bResponse = 1 Then
      MsgBox "The Selected Manufacturing Order Run Is In Use.", _
         vbInformation, Caption
      Exit Sub
   End If
   
   If Val(txtQty) = 0 Then
      MsgBox "Requires a valid Quantity.", _
         vbInformation, Caption
      Exit Sub
   End If
   
   If Val(txtQty) >= Val(lblRemQty) Then
      MsgBox "The Quantity Split Must Be Less Than Remaining.", _
         vbInformation, Caption
      Exit Sub
   End If
   
   sMsg = "Warning:" & vbCr _
          & "If There Are Purchase Order Allocations Assigned" & vbCr _
          & "To The Original Manufacturing Order, Then Only " & vbCr _
          & "Received Items Will Be Resolved." & vbCr _
          & "Do You Wish To Continue?"
   bResponse = MsgBox(sMsg, vbYesNo + vbExclamation, Caption)
   
   If bResponse = vbNo Then
      txtRun = ""
      txtQty = ""
      lblPerc = ""
      cmdSplit.Enabled = False
      txtRun.Enabled = False
      txtQty.Enabled = False
      CancelTrans
      Exit Sub
   End If
   sMsg = "The Existing Manufacturing Order Run Will Be" & vbCr _
          & "Split To Create A New Manufacturing Order." & vbCr _
          & "Existing Items Will Be Split Based On A " & vbCr _
          & "Percentage Of " & lblPerc & " Except Items With A" & vbCr _
          & "Unit Of Measure Of 'EA' (Uses The Count). " & vbCr _
          & "Do You Wish To Continue?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbNo Then
      txtRun = ""
      txtQty = ""
      lblPerc = ""
      cmdSplit.Enabled = False
      txtRun.Enabled = False
      txtQty.Enabled = False
      CancelTrans
      Exit Sub
   End If
   newMORun = txtRun
   SplitThisMo
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      MouseCursor 13
      lblPrg.ForeColor = ES_BLUE
      lblPrg.Caption = "Parameters "
      lblPrg.Visible = True
      prg1.Visible = True
      prg1.Value = 10
      lblPrg.Refresh
      zCreateTimeTable
      zCreatePickTable
      prg1.Value = 20
      zCreateInvaTable
      zCreateLotsTable
      lblPrg.Caption = "Temp Tables "
      lblPrg.Refresh
      prg1.Value = 40
      zCreateRunsTable
      zCreateRnopTable
      zCreateOutSideTable
      zCreatePOitTable
      prg1.Value = 70
      optSplit = GetSplitStatus()
      FillCombo
      prg1.Value = 100
      Sleep 1000
      prg1.Visible = False
   End If
   bOnLoad = 0
   MouseCursor 0
   If bTablesFailed = 1 Then
      MsgBox "Warning:" & vbCr _
         & "Could Not Create One Or More Necessary " & vbCr _
         & "Temporary Tables To Create Splits. Have " & vbCr _
         & "Your System Administrator Check Your  " & vbCr _
         & "Database Permissions.", vbExclamation, Caption
      Unload Me
   End If
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   'Redundant because RdoCon will drop them when closed
   '(reduces clutter)
   On Error Resume Next
   If sRunsTemp <> "" Then clsADOCon.ExecuteSql "DROP TABLE " & sRunsTemp
   If sRnopTemp <> "" Then clsADOCon.ExecuteSql "DROP TABLE " & sRnopTemp
   If sTimeTemp <> "" Then clsADOCon.ExecuteSql "DROP TABLE " & sTimeTemp
   If sPickTemp <> "" Then clsADOCon.ExecuteSql "DROP TABLE " & sPickTemp
   If sInvaTemp <> "" Then clsADOCon.ExecuteSql "DROP TABLE " & sInvaTemp
   If sLotsTemp <> "" Then clsADOCon.ExecuteSql "DROP TABLE " & sLotsTemp
   If sOutSideSrvTemp <> "" Then clsADOCon.ExecuteSql "DROP TABLE " & sOutSideSrvTemp
   
   Set ShopSHf06a = Nothing
   MouseCursor 0
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   cmbPrt.Clear
   sSql = "SELECT DISTINCT RUNREF,PARTREF,PARTNUM FROM " _
          & "RunsTable,PartTable WHERE (RUNREF=PARTREF AND " _
          & "RUNSTATUS NOT LIKE 'C%' AND RUNREMAININGQTY>1)"
   LoadComboBox cmbPrt, 1
   If cmbPrt.ListCount > 0 Then
      cmbPrt = cmbPrt.List(0)
      GetCurrentPart cmbPrt, lblDsc
      lblNew = cmbPrt
   Else
      MsgBox "No Qualifying Runs To Split.", _
         vbInformation, Caption
      cmdSel.Enabled = False
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillRuns()
   txtRun = ""
   txtQty = ""
   lblPerc = ""
   cmdSplit.Enabled = False
   txtRun.Enabled = False
   txtQty.Enabled = False
   If sOldPart = cmbPrt Then Exit Sub
   sOldPart = cmbPrt
   cmbRun.Clear
   cmdSel.Enabled = False
   sSql = "SELECT RUNNO FROM RunsTable WHERE (RUNREF='" _
          & Compress(cmbPrt) & "' AND RUNSTATUS NOT LIKE 'C%' " _
          & "AND RUNREMAININGQTY>1)"
   LoadNumComboBox cmbRun, 0
   If cmbRun.ListCount > 0 Then
      cmbRun = cmbRun.List(0)
      cmdSel.Enabled = True
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 6) = "*** Pa" Then
      lblDsc.ForeColor = ES_RED
      cmdSel.Enabled = False
   Else
      lblDsc.ForeColor = vbBlack
      cmdSel.Enabled = True
   End If
   
End Sub


Private Function GetSplitStatus()
   Dim RdoStatus As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT COALLOWSPLITS,COSPLITSTARTRUN " _
          & "FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoStatus, ES_FORWARD)
   If bSqlRows Then
      With RdoStatus
         GetSplitStatus = !COALLOWSPLITS
         iDefaultRun = !COSPLITSTARTRUN
         ClearResultSet RdoStatus
      End With
   Else
      GetSplitStatus = 0
   End If
   On Error Resume Next
   sSql = "UPDATE InvaTable SET INUNITS=PAUNITS " _
          & "FROM InvaTable,PartTable WHERE (INPART=PARTREF AND " _
          & "INUNITS='')"
   clsADOCon.ExecuteSql sSql
   
   sSql = "UPDATE LoitTable SET LOIUNITS=PAUNITS " _
          & "FROM LoitTable,PartTable WHERE (LOIPARTREF=PARTREF AND " _
          & "LOIUNITS='')"
   clsADOCon.ExecuteSql sSql
   
   Set RdoStatus = Nothing
   Exit Function
   
DiaErr1:
   GetSplitStatus = 0
   
End Function

Private Function GetRemaining()
   On Error Resume Next
   Dim RdoRem As ADODB.Recordset
   sSql = "SELECT RUNREMAININGQTY,RUNSTATUS FROM RunsTable WHERE " _
          & "RUNREF='" & Compress(cmbPrt) & "' AND RUNNO=" _
          & Val(cmbRun) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRem, ES_FORWARD)
   If bSqlRows Then
      With RdoRem
         GetRemaining = !RUNREMAININGQTY
         lblStatus = "" & Trim(!RUNSTATUS)
         ClearResultSet RdoRem
      End With
   End If
   Set RdoRem = Nothing
   
End Function

Private Function GetSplitHistory()
   On Error GoTo DiaErr1
   Dim RdoOld As ADODB.Recordset
   sSql = "SELECT MAX(RUNLASTSPLITRUNNO)AS LastSplit FROM " _
          & "RunsTable WHERE RUNREF='" & Compress(cmbPrt) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoOld, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoOld!LastSplit) Then _
                    GetSplitHistory = RdoOld!LastSplit + 1
      ClearResultSet RdoOld
   End If
   If GetSplitHistory < 2 Then GetSplitHistory = iDefaultRun
   Set RdoOld = Nothing
   Exit Function
   
DiaErr1:
   MsgBox "Splits Are Not Properly Initialized.", _
      vbInformation, Caption
   
End Function

Private Function GetNextPoItem()
   On Error GoTo DiaErr1
   Dim RdoOld As ADODB.Recordset
   sSql = "SELECT MAX(PIITEM) + 1 AS NextPOIT FROM poitTable " _
          & " WHERE PINUMBER IN ( SELECT DISTINCT PINUMBER FROM poitTable WHERE pirunpart ='" & Compress(cmbPrt) & "' AND pirunno =" _
          & Val(cmbRun) & ")"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoOld, ES_FORWARD)
   
   If bSqlRows Then
      If Not IsNull(RdoOld!NextPOIT) Then _
                    GetNextPoItem = RdoOld!NextPOIT
      ClearResultSet RdoOld
   Else
    GetNextPoItem = -1
   End If
   Set RdoOld = Nothing
   Exit Function
   
DiaErr1:
   MsgBox "Get Next Vendor Invoice.", _
      vbInformation, Caption
   
End Function


Private Sub txtQty_LostFocus()
   txtQty = Format(Abs(Val(txtQty)), ES_QuantityDataFormat)
   If Val(txtQty) > 0 Then
      lblPerc = Format(Val(txtQty) / Val(lblRemQty), "#0.0%")
      cSplitPerc = Format(Val(txtQty) / Val(lblRemQty), "#0.000")
   Else
      cSplitPerc = 0
      lblPerc = ""
      lblStatus = ""
   End If
   
End Sub



Private Function CheckExistingRun() As Byte
   Dim RdoGet As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT RUNREF,RUNNO FROM RunsTable WHERE " _
          & "(RUNREF='" & Compress(cmbPrt) & "' AND " _
          & "RUNNO=" & Val(txtRun) & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet, ES_FORWARD)
   If bSqlRows Then CheckExistingRun = 1 _
                                       Else CheckExistingRun = 0
   Set RdoGet = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "checkexistingr"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub SplitThisMo()
   Dim bByte As Byte
   
   On Error Resume Next
   'Dump any exiting trash
   clsADOCon.ExecuteSql "TRUNCATE TABLE " & sRunsTemp
   clsADOCon.ExecuteSql "TRUNCATE TABLE " & sRnopTemp
   clsADOCon.ExecuteSql "TRUNCATE TABLE " & sTimeTemp
   clsADOCon.ExecuteSql "TRUNCATE TABLE " & sPickTemp
   clsADOCon.ExecuteSql "TRUNCATE TABLE " & sInvaTemp
   clsADOCon.ExecuteSql "TRUNCATE TABLE " & sLotsTemp
   
   clsADOCon.ExecuteSql "TRUNCATE TABLE " & sOutSideSrvTemp
   
   cmdSel.Enabled = False
   cmdSplit.Enabled = False
   
   lblPrg.ForeColor = vbBlack
   lblPrg.Visible = True
   prg1.Visible = True
   lblPrg = "Creating MO:"
   prg1.Value = 10
   lblPrg.Refresh
   Sleep 700
   
   sProcName = "createmo"
   ' MM I am not sure whiy this was commented. 9/24/2009
   bByte = CreateManufacturingOrder()
   prg1.Value = 100
   lblPrg.Refresh
   Sleep 700
   ' MM bByte = 1
   On Error GoTo DiaErr1
   If bByte = 1 Then
      bSplitFailed = 0
      MouseCursor 13
      sProcName = "splittime"
      SplitTime
      
      sProcName = "splitpicks"
      SplitPicks
      
      sProcName = "splitactivity"
      SplitActivity
      
      sProcName = "splitlots"
      SplitLots
      
      sProcName = "CopyDocuments"
      CopyMODocs
      
      sProcName = "Split Outside service"
      SplitOutsideService
      
      sProcName = "Split PO Outside service"
      SplitPOService
      
      MouseCursor 0
      MsgBox "The Manufacturing Order Was Successfully Split.", _
         vbInformation, Caption
      lblStatus = ""
      lblRemQty = ""
      FillCombo
   Else
      MsgBox "Failed To Create The New Manufacturing Order.", _
         vbInformation, Caption
   End If
   lblPrg.Visible = False
   prg1.Visible = False
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub SplitTime()
   Dim cOldPerc As Currency
   cOldPerc = 1 - cSplitPerc
   On Error Resume Next
   Err.Clear
   lblPrg = "Time Charges:"
   prg1.Value = 10
   lblPrg.Refresh
   Sleep 700
   
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   sSql = "INSERT " & sTimeTemp & " SELECT * FROM TcitTable WHERE " _
          & "TCPARTREF='" & Compress(cmbPrt) & "' AND TCRUNNO=" _
          & Val(cmbRun) & " "
   clsADOCon.ExecuteSql sSql
   prg1.Value = 20
   Sleep 700
   
   sSql = "UPDATE " & sTimeTemp & " SET TCRUNNO=" & Val(txtRun) & "," _
          & "TCHOURS=TCHOURS * " & cSplitPerc & "," _
          & "TCYIELD=TCYIELD * " & cSplitPerc & "," _
          & "TCACCEPT=TCACCEPT * " & cSplitPerc & "," _
          & "TCREJECT=TCREJECT * " & cSplitPerc & "," _
          & "TCSCRAP=TCSCRAP * " & cSplitPerc & " "
   clsADOCon.ExecuteSql sSql
   prg1.Value = 40
   Sleep 700
   
   sSql = "UPDATE TcitTable SET " _
          & "TCHOURS=TCHOURS * " & cOldPerc & "," _
          & "TCYIELD=TCYIELD * " & cOldPerc & "," _
          & "TCACCEPT=TCACCEPT * " & cOldPerc & "," _
          & "TCREJECT=TCREJECT * " & cOldPerc & "," _
          & "TCSCRAP=TCSCRAP * " & cOldPerc & " " _
          & "WHERE TCPARTREF='" & Compress(cmbPrt) & "' " _
          & "AND TCRUNNO=" & Val(cmbRun) & " "
   clsADOCon.ExecuteSql sSql
   prg1.Value = 60
   Sleep 700
   
   sSql = "INSERT TcitTable SELECT * FROM " & sTimeTemp & " WHERE " _
          & "TCPARTREF='" & Compress(lblNew) & "' "
   clsADOCon.ExecuteSql sSql
   prg1.Value = 100
   Sleep 700
   
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
   Else
      bSplitFailed = 1
      clsADOCon.RollbackTrans
   End If
End Sub

Private Sub SplitOutsideService()
   Dim cOldPerc As Currency
   cOldPerc = 1 - cSplitPerc
   ES_SYSDATE = GetServerDateTime()
   
   lblPrg = "OutSide Service:"
   prg1.Value = 10
   lblPrg.Refresh
   Sleep 700
   
   On Error Resume Next
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   Dim viitem As Integer
   viitem = GetNextVendorInv
   If (viitem <> -1) Then
      sSql = "INSERT " & sOutSideSrvTemp & " SELECT * FROM viitTable WHERE " _
             & " vitmo='" & Compress(cmbPrt) & "' AND vitmorun =" _
             & Val(cmbRun) & " "
      clsADOCon.ExecuteSql sSql
      prg1.Value = 20
      Sleep 700
      
      sSql = "UPDATE " & sOutSideSrvTemp & " SET VITITEM =" & Val(viitem) & "," _
             & " vitmorun = " & Val(txtRun) & ", " _
             & "VITQTY = VITQTY * " & cSplitPerc & " "
             
      clsADOCon.ExecuteSql sSql
      prg1.Value = 40
      Sleep 700
      
      sSql = "UPDATE viitTable SET " _
             & "VITQTY = VITQTY * " & cOldPerc & "" _
             & "WHERE vitmo ='" & Compress(cmbPrt) & "' " _
             & "AND vitmorun =" & Val(cmbRun) & " "
      clsADOCon.ExecuteSql sSql
      prg1.Value = 60
      Sleep 700
      
      sSql = "INSERT viitTable SELECT * FROM " & sOutSideSrvTemp & " WHERE " _
             & "vitmo ='" & Compress(cmbPrt) & "' " _
             & "AND vitmorun =" & Val(txtRun) & " "
      clsADOCon.ExecuteSql sSql
      prg1.Value = 100
      Sleep 700
      
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
      Else
         bOutSideFailed = 1
         clsADOCon.RollbackTrans
      End If
      
      sSql = "DELETE FROM " & sOutSideSrvTemp
      clsADOCon.ExecuteSql sSql
   End If

End Sub

Private Function GetNextVendorInv()
   On Error GoTo DiaErr1
   Dim RdoOld As ADODB.Recordset
   sSql = "SELECT MAX(VITITEM) + 1 AS NextVIT FROM viitTable " _
          & " WHERE VITNO IN ( SELECT DISTINCT VITNO FROM viitTable WHERE vitmo='" & Compress(cmbPrt) & "' AND vitmorun =" _
          & Val(cmbRun) & ")"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoOld, ES_FORWARD)
   
   If bSqlRows Then
      If Not IsNull(RdoOld!NextVIT) Then _
                    GetNextVendorInv = RdoOld!NextVIT
      ClearResultSet RdoOld
   Else
    GetNextVendorInv = -1
   End If
   Set RdoOld = Nothing
   Exit Function
   
DiaErr1:
   MsgBox "Get Next Vendor Invoice.", _
      vbInformation, Caption
   
End Function


Private Sub SplitPOService()
   Dim cOldPerc As Currency
   cOldPerc = 1 - cSplitPerc
   ES_SYSDATE = GetServerDateTime()
   
   lblPrg = "PO Service:"
   prg1.Value = 10
   lblPrg.Refresh
   Sleep 700
   
   On Error Resume Next
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   Dim POITEM As Integer
   POITEM = GetNextPoItem
   If (POITEM <> -1) Then
      sSql = "INSERT " & sPoitSrvTemp & " SELECT * FROM poitTable WHERE " _
             & " pirunpart ='" & Compress(cmbPrt) & "' AND pirunno =" _
             & Val(cmbRun) & " "
      clsADOCon.ExecuteSql sSql
      prg1.Value = 20
      Sleep 700
      
      'MsgBox (sSql)
      'MsgBox (clsADOCon.ADOErrNum)
      
      
      sSql = "UPDATE " & sPoitSrvTemp & " SET PIITEM =" & Val(POITEM) & "," _
             & " pirunno = " & Val(newMORun) & ", " _
             & " PIPQTY = PIPQTY * " & cSplitPerc & ", " _
             & " PIAQTY = PIAQTY * " & cSplitPerc & " " _
             
             
      clsADOCon.ExecuteSql sSql
      prg1.Value = 40
      Sleep 700
      
      'MsgBox (sSql)
      'MsgBox (clsADOCon.ADOErrNum)
      
      sSql = "UPDATE PoitTable SET " _
             & " PIPQTY = PIPQTY * " & cOldPerc & ", " _
             & " PIAQTY = PIAQTY * " & cOldPerc & " " _
             & " WHERE pirunpart ='" & Compress(cmbPrt) & "' " _
             & " AND pirunno =" & Val(cmbRun) & " "
             
      clsADOCon.ExecuteSql sSql
      prg1.Value = 60
      Sleep 700
      
      'MsgBox (sSql)
      'MsgBox (clsADOCon.ADOErrNum)
      
      sSql = "INSERT PoitTable SELECT * FROM " & sPoitSrvTemp & " WHERE " _
             & " pirunpart ='" & Compress(cmbPrt) & "' " _
             & " AND pirunno =" & Val(newMORun) & " "
      clsADOCon.ExecuteSql sSql
      prg1.Value = 100
      Sleep 700
      
      'MsgBox (sSql)
      'MsgBox (clsADOCon.ADOErrNum)
      If clsADOCon.ADOErrNum = 0 Then
         
         clsADOCon.CommitTrans
      Else
         bOutSideFailed = 1
         clsADOCon.RollbackTrans
      End If
      
      sSql = "DELETE FROM " & sPoitSrvTemp
      clsADOCon.ExecuteSql sSql
      
   End If

End Sub



Private Sub SplitPicks()
   Dim cOldPerc As Currency
   cOldPerc = 1 - cSplitPerc
   ES_SYSDATE = GetServerDateTime()
   
   lblPrg = "MO Picks:"
   prg1.Value = 10
   lblPrg.Refresh
   Sleep 700
   
   On Error Resume Next
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   sSql = "INSERT " & sPickTemp & " SELECT * FROM MopkTable WHERE " _
          & "PKMOPART='" & Compress(cmbPrt) & "' AND PKMORUN=" _
          & Val(cmbRun) & " "
   clsADOCon.ExecuteSql sSql
   prg1.Value = 20
   Sleep 700
   
   '// MM - When we split an MO we would always want to split by percentage
   '// Commented these update - 3/27/2010
   '// Ticket# 48479
'   sSql = "UPDATE " & sPickTemp & " SET " _
'          & "PKMORUN=" & Val(txtRun) & "," _
'          & "PKAQTY=" & Val(txtQty) & "," _
'          & "PKPQTY=" & Val(txtQty) & "," _
'          & "PKADATE='" & Format(ES_SYSDATE, "mm/dd/yy hh:mm") & "'," _
'          & "PKPDATE='" & Format(ES_SYSDATE, "mm/dd/yy hh:mm") & "' "
''          & "WHERE PKUNITS='EA'"
' MM   clsADOCon.ExecuteSQL sSql
   
   sSql = "UPDATE " & sPickTemp & " SET " _
          & "PKMORUN=" & Val(txtRun) & "," _
          & "PKAQTY=PKAQTY * " & cSplitPerc & "," _
          & "PKPQTY=PKPQTY * " & cSplitPerc & "," _
          & "PKADATE='" & Format(ES_SYSDATE, "mm/dd/yy") & "'," _
          & "PKPDATE='" & Format(ES_SYSDATE, "mm/dd/yy") & "' "
          '& "WHERE PKUNITS<>'EA'"
   clsADOCon.ExecuteSql sSql
   prg1.Value = 40
   Sleep 700
   
'   sSql = "UPDATE MopkTable SET " _
'          & "PKAQTY=PKAQTY-" & Val(txtQty) & "," _
'          & "PKPQTY=PKPQTY-" & Val(txtQty) & " " _
'          & "WHERE (PKMOPART='" & Compress(cmbPrt) & "' AND " _
'          & "PKMORUN=" & Val(cmbRun) & " AND PKUNITS='EA')"
'   clsAdoCon.ExecuteSQL sSql
   prg1.Value = 60
   Sleep 700
   
   sSql = "UPDATE MopkTable SET " _
          & "PKAQTY=PKAQTY * " & cOldPerc & "," _
          & "PKPQTY=PKPQTY * " & cOldPerc & " " _
          & "WHERE (PKMOPART='" & Compress(cmbPrt) & "' AND " _
          & "PKMORUN=" & Val(cmbRun) & ")"
          'AND PKUNITS<>'EA')"
   clsADOCon.ExecuteSql sSql
   prg1.Value = 80
   Sleep 700
   
   sSql = "INSERT MopkTable SELECT * FROM " & sPickTemp & " WHERE " _
          & "PKMOPART='" & Compress(cmbPrt) & "' "
   clsADOCon.ExecuteSql sSql
   prg1.Value = 100
   Sleep 700
   
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
   Else
      bSplitFailed = 1
      clsADOCon.RollbackTrans
   End If
   
End Sub

Private Sub SplitActivity()
   Dim RdoUpd As ADODB.Recordset
   Dim iLength As Integer
   Dim lCOUNTER As Long
   Dim cOldPerc As Currency
   
   Dim sMoRun As String * 9
   Dim sMoPart As String * 31
   
   lblPrg = "Activity Rows:"
   prg1.Value = 10
   lblPrg.Refresh
   Sleep 700
   
   cOldPerc = 1 - cSplitPerc
   iLength = Len(Trim(str(cmbRun)))
   iLength = 4 - iLength
   sMoPart = cmbPrt
   sMoRun = "RUN" & Space$(iLength) & txtRun
   
   On Error Resume Next
   Err.Clear
   sSql = "INSERT " & sInvaTemp & " (INTYPE,INPART,INREF1,INREF2,INPDATE,INADATE,INPQTY,INAQTY,INAMT," & _
               "INTOTMATL,INTOTLABOR,INTOTEXP,INTOTOH,INTOTHRS,INCREDITACCT,INDEBITACCT," & _
               "INGLJOURNAL,INGLPOSTED,INGLDATE,INMOPART,INMORUN,INSONUMBER,INSOITEM," & _
               "INSOREV,INPONUMBER,INPORELEASE,INPOITEM,INPOREV,INPSNUMBER,INPSITEM,INWIPLABACCT," & _
               "INWIPMATACCT,INWIPOHDACCT,INWIPEXPACCT,INNUMBER,INLOTNUMBER,INUSER,INUNITS," & _
               "INDRLABACCT,INDRMATACCT,INDREXPACCT,INDROHDACCT,INCRLABACCT,INCRMATACCT,INCREXPACCT," & _
               "INCROHDACCT,INLOTTRACK, INUSEACTUALCOST, INCOSTEDBY, INMAINTCOSTED) " & _
            "SELECT INTYPE,INPART,INREF1,INREF2,INPDATE,INADATE,INPQTY,INAQTY,INAMT," & _
               "INTOTMATL,INTOTLABOR,INTOTEXP,INTOTOH,INTOTHRS,INCREDITACCT,INDEBITACCT," & _
               "INGLJOURNAL,INGLPOSTED,INGLDATE,INMOPART,INMORUN,INSONUMBER,INSOITEM," & _
               "INSOREV,INPONUMBER,INPORELEASE,INPOITEM,INPOREV,INPSNUMBER,INPSITEM,INWIPLABACCT," & _
               "INWIPMATACCT,INWIPOHDACCT,INWIPEXPACCT,INNUMBER,INLOTNUMBER,INUSER,INUNITS," & _
               "INDRLABACCT,INDRMATACCT,INDREXPACCT,INDROHDACCT,INCRLABACCT,INCRMATACCT,INCREXPACCT," & _
               "INCROHDACCT,INLOTTRACK, INUSEACTUALCOST, INCOSTEDBY, INMAINTCOSTED FROM InvaTable WHERE " & _
            "INMOPART='" & Compress(cmbPrt) & "' AND INMORUN=" & Val(cmbRun) & ""
   
   clsADOCon.ExecuteSql sSql
   prg1.Value = 20
   Sleep 700
   
   'Update the rows - Tested with 100 dummy rows
   'Leave the INNUMBER as is for joins (relationships)
   '    lCounter = GetLastActivity()
   '    sSql = "SELECT INPART,INNUMBER FROM " & sInvaTemp & " "
   '    bsqlrows = clsadocon.getdataset(ssql, RdoUpd, ES_STATIC)
   '        If bSqlRows Then
   '            With RdoUpd
   '                Do Until .EOF
   '                    lCounter = lCounter + 1
   '                    sSql = "UPDATE " & sInvaTemp & " SET INNUMBER=" _
   '                        & lCounter & " WHERE (INPART='" & Trim(!INPART) & "' AND " _
   '                        & "INNUMBER=" & !INNUMBER & ") "
   '                    clsAdoCon.ExecuteSQL sSql
   '                    .MoveNext
   '                Loop
   '                .Cancel
   '            End With
   '        End If
   prg1.Value = 30
   Sleep 700
   
   sSql = "UPDATE " & sInvaTemp & " SET INREF2='" & sMoPart & sMoRun & "'," _
          & "INMORUN=" & Val(txtRun) & " "
   clsADOCon.ExecuteSql sSql
   prg1.Value = 40
   Sleep 700
   
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   'UPDATE InvaTable
   ' 10/31/2010 - These are pick items so the INAQTY & INPQTY should be negative.
   '// MM - When we split an MO we would always want to split by percentage
   '// Commented these update - 3/27/2011
   '// Ticket# 48479
'   sSql = "UPDATE InvaTable SET INAQTY=INAQTY+" & Val(txtQty) _
'          & ",INPQTY=INPQTY+" & Val(txtQty) _
'          & " FROM InvaTable, PartTable WHERE (INPART=PARTREF " _
'          & "AND INMOPART='" & Compress(cmbPrt) & "' AND INMORUN=" _
'          & Val(cmbRun) & " AND PAUNITS='EA')"
'   clsAdoCon.ExecuteSQL sSql
   
   '    sSql = "DELETE FROM " & sInvaTemp & " WHERE INAQTY< " & Val(txtQty) _
   '        & " AND INUNITS='EA'"
   '    clsAdoCon.ExecuteSQL sSql
   prg1.Value = 50
   Sleep 700
   
   sSql = "UPDATE InvaTable SET INAQTY=INAQTY * " & cOldPerc _
          & ",INPQTY=INPQTY * " & cOldPerc _
          & ",INTOTMATL = INTOTMATL * " & cOldPerc _
          & ",INTOTLABOR = INTOTLABOR * " & cOldPerc _
          & ",INTOTEXP = INTOTEXP * " & cOldPerc _
          & ",INTOTOH = INTOTOH * " & cOldPerc _
          & " FROM InvaTable, PartTable WHERE (INPART=PARTREF " _
          & "AND INMOPART='" & Compress(cmbPrt) & "' AND INMORUN=" & Val(cmbRun) & ")"
          '& " AND PAUNITS<>'EA')"
   clsADOCon.ExecuteSql sSql
   prg1.Value = 60
   Sleep 700
   
'   sSql = "UPDATE " & sInvaTemp & " SET INAQTY=-" & Val(txtQty) _
'          & " WHERE INUNITS='EA'"
'   clsAdoCon.ExecuteSQL sSql
   prg1.Value = 70
   Sleep 700
   
   sSql = "UPDATE " & sInvaTemp & " SET INAQTY=INAQTY * " & cSplitPerc _
          & ",INTOTMATL = INTOTMATL * " & cSplitPerc _
          & ",INTOTLABOR = INTOTLABOR * " & cSplitPerc _
          & ",INTOTEXP = INTOTEXP * " & cSplitPerc _
          & ",INTOTOH = INTOTOH * " & cSplitPerc
          '& " WHERE INUNITS<>'EA'"
   clsADOCon.ExecuteSql sSql
   prg1.Value = 80
   Sleep 700
   
   sSql = "UPDATE " & sInvaTemp & " SET INPQTY=INAQTY," _
          & "INADATE='" & Format(ES_SYSDATE, "mm/dd/yy") & "'," _
          & "INPDATE='" & Format(ES_SYSDATE, "mm/dd/yy") & "' "
   clsADOCon.ExecuteSql sSql
   
   
   sSql = "INSERT InvaTable (INTYPE,INPART,INREF1,INREF2,INPDATE,INADATE,INPQTY,INAQTY,INAMT," & _
            "INTOTMATL,INTOTLABOR,INTOTEXP,INTOTOH,INTOTHRS,INCREDITACCT,INDEBITACCT," & _
            "INGLJOURNAL,INGLPOSTED,INGLDATE,INMOPART,INMORUN,INSONUMBER,INSOITEM," & _
            "INSOREV,INPONUMBER,INPORELEASE,INPOITEM,INPOREV,INPSNUMBER,INPSITEM,INWIPLABACCT," & _
            "INWIPMATACCT,INWIPOHDACCT,INWIPEXPACCT,INNUMBER,INLOTNUMBER,INUSER,INUNITS," & _
            "INDRLABACCT,INDRMATACCT,INDREXPACCT,INDROHDACCT,INCRLABACCT,INCRMATACCT,INCREXPACCT," & _
            "INCROHDACCT,INLOTTRACK, INUSEACTUALCOST, INCOSTEDBY, INMAINTCOSTED)" & _
                  " SELECT INTYPE,INPART,INREF1,INREF2,INPDATE,INADATE,INPQTY,INAQTY,INAMT," & _
                  "INTOTMATL,INTOTLABOR,INTOTEXP,INTOTOH,INTOTHRS,INCREDITACCT,INDEBITACCT," & _
                  "INGLJOURNAL,INGLPOSTED,INGLDATE,INMOPART,INMORUN,INSONUMBER,INSOITEM," & _
                  "INSOREV,INPONUMBER,INPORELEASE,INPOITEM,INPOREV,INPSNUMBER,INPSITEM,INWIPLABACCT," & _
                  "INWIPMATACCT,INWIPOHDACCT,INWIPEXPACCT,INNUMBER,INLOTNUMBER,INUSER,INUNITS," & _
                  "INDRLABACCT,INDRMATACCT,INDREXPACCT,INDROHDACCT,INCRLABACCT,INCRMATACCT,INCREXPACCT," & _
                  "INCROHDACCT , INLOTTRACK, INUSEACTUALCOST, INCOSTEDBY, INMAINTCOSTED" & _
         " FROM " & sInvaTemp & " WHERE INMOPART='" & Compress(cmbPrt) & "' "
   
   clsADOCon.ExecuteSql sSql
   prg1.Value = 100
   Sleep 700
   
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
   Else
      bSplitFailed = 1
      clsADOCon.RollbackTrans
   End If
   Set RdoUpd = Nothing
   
End Sub

Private Sub SplitLots()
   Dim RdoUpd As ADODB.Recordset
   Dim lCOUNTER As Long
   Dim cOldPerc As Currency
   
   cOldPerc = 1 - cSplitPerc
   
   lblPrg = "Available Lots:"
   prg1.Value = 10
   lblPrg.Refresh
   
   On Error Resume Next
   
   Err.Clear
'   sSql = "INSERT " & sLotsTemp & " SELECT * FROM LoitTable WHERE " _
'          & "LOIMOPARTREF='" & Compress(cmbPrt) & "' AND LOIMORUNNO=" _
'          & Val(cmbRun) & " "
   sSql = "INSERT " & sLotsTemp & " (" & loitColumns & ") SELECT " & loitColumns & " FROM LoitTable WHERE " _
          & "LOIMOPARTREF='" & Compress(cmbPrt) & "' AND LOIMORUNNO=" _
          & Val(cmbRun) & " "
   clsADOCon.ExecuteSql sSql
   prg1.Value = 20
   Sleep 700
   
   sSql = "UPDATE " & sLotsTemp & " SET LOIMORUNNO=" & Val(txtRun) & " "
   clsADOCon.ExecuteSql sSql
   prg1.Value = 40
   Sleep 700
   
   Dim strPrvLoiNum As String
   strPrvLoiNum = ""
   sSql = "SELECT LOINUMBER,LOIPARTREF,LOIRECORD FROM " & sLotsTemp & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoUpd, ES_STATIC)
   If bSqlRows Then
      With RdoUpd
         
         Do Until .EOF
            If (!LOINUMBER = strPrvLoiNum) Then
               lCOUNTER = lCOUNTER + 1
            Else
               lCOUNTER = GetNextLotRecord(!LOINUMBER)
               strPrvLoiNum = !LOINUMBER
            End If
            sSql = "UPDATE " & sLotsTemp & " SET LOIRECORD=" _
                   & lCOUNTER & " WHERE (LOINUMBER='" & Trim(!LOINUMBER) _
                   & "' AND LOIRECORD=" & !LOIRECORD & ") "
            clsADOCon.ExecuteSql sSql
            .MoveNext
         Loop
         ClearResultSet RdoUpd
      End With
   End If
   prg1.Value = 40
   Sleep 700
   
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   'UPDATE LotsTable
   ' 10/2/2010 The Adding the qty beciasue LOIQUANTITY is negative.
   
   ' 10/31/2010 - These are pick items so the INAQTY & INPQTY should be negative.
   '// MM - When we split an MO we would always want to split by percentage
   '// Commented these update - 3/27/2011
   '// Ticket# 48479
'   sSql = "UPDATE LoitTable SET LOIQUANTITY=LOIQUANTITY+" & Val(txtQty) _
'          & " WHERE (LOIMOPARTREF='" & Compress(cmbPrt) & "'  AND LOIMORUNNO=" _
'          & Val(cmbRun) & " AND LOIUNITS='EA')"
'   clsAdoCon.ExecuteSQL sSql
'          & ",LOTTOTMATL = LOTTOTMATL * " & cOldPerc _
'          & ",LOTTOTLABOR = LOTTOTLABOR * " & cOldPerc _
'          & ",LOTTOTEXP = LOTTOTEXP * " & cOldPerc _
'          & ",LOTTOTOH = LOTTOTOH * " & cOldPerc _

   '    sSql = "DELETE FROM " & sLotsTemp & " WHERE LOIQUANTITY>=" _
   '        & Val(txtQty) & " AND LOIUNITS='EA')"
   '    clsAdoCon.ExecuteSQL sSql
   
   prg1.Value = 50
   Sleep 700
   
   sSql = "UPDATE LoitTable SET LOIQUANTITY=LOIQUANTITY * " & cOldPerc _
          & " WHERE (LOIMOPARTREF='" & Compress(cmbPrt) & "' AND LOIMORUNNO=" & Val(cmbRun) & ")"
          '& " AND LOIUNITS<>'EA')"
   clsADOCon.ExecuteSql sSql
   prg1.Value = 60
   Sleep 700
   ' 10/2/2010 - These are pick items so the LOIQUANTITY should be negative.
'   sSql = "UPDATE " & sLotsTemp & " SET LOIQUANTITY=-" & Val(txtQty) _
'          & " WHERE LOIUNITS='EA'"
'   clsAdoCon.ExecuteSQL sSql
   prg1.Value = 70
   Sleep 700
   
   ' 10/2/2010 - These are pick items so the LOIQUANTITY should be negative.
   sSql = "UPDATE " & sLotsTemp & " SET LOIQUANTITY=LOIQUANTITY * " & cSplitPerc
'          & ",LOTTOTMATL = LOTTOTMATL * " & cSplitPerc _
'          & ",LOTTOTLABOR = LOTTOTLABOR * " & cSplitPerc _
'          & ",LOTTOTEXP = LOTTOTEXP * " & cSplitPerc _
'          & ",LOTTOTOH = LOTTOTOH * " & cSplitPerc

          '& " WHERE LOIUNITS<>'EA'"
   clsADOCon.ExecuteSql sSql
   prg1.Value = 80
   Sleep 700
   
   sSql = "UPDATE " & sLotsTemp & " SET LOIADATE='" & Format(ES_SYSDATE, "mm/dd/yy") _
          & "',LOIPDATE='" & Format(ES_SYSDATE, "mm/dd/yy") & "'," _
          & "LOIMORUNNO=" & Val(txtRun) & " "
   clsADOCon.ExecuteSql sSql
   prg1.Value = 90
   Sleep 700
   
   sSql = "INSERT LoitTable SELECT * FROM " & sLotsTemp & " WHERE " _
          & "LOIMOPARTREF='" & Compress(cmbPrt) & "' "
   clsADOCon.ExecuteSql sSql
   
   prg1.Value = 100
   Sleep 700
   
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
   Else
      bSplitFailed = 1
      clsADOCon.RollbackTrans
   End If
   Set RdoUpd = Nothing
   
End Sub

Private Sub CopyMODocs()
   
   lblPrg = "Copy MO Docs:"
   prg1.Value = 10
   lblPrg.Refresh
   
   On Error Resume Next
   Err.Clear
   
   sSql = "INSERT INTO RndlTable (RUNDLSNUM,RUNDLSRUNREF," _
         & " RUNDLSRUNNO,RUNDLSREV,RUNDLSDOCREF,RUNDLSDOCREV," _
         & " RUNDLSDOCREFLONG,RUNDLSDOCREFDESC,RUNDLSDOCREFSHEET," _
         & " RUNDLSDOCREFCLASS,RUNDLSDOCREFADCN," _
         & " RUNDLSDOCREFECO)" _
   & " SELECT RUNDLSNUM,'" & Compress(cmbPrt) & "'," & Val(txtRun) & "," _
          & " RUNDLSREV,RUNDLSDOCREF,RUNDLSDOCREV," _
          & " RUNDLSDOCREFLONG,RUNDLSDOCREFDESC,RUNDLSDOCREFSHEET," _
          & " RUNDLSDOCREFCLASS,RUNDLSDOCREFADCN," _
          & " RUNDLSDOCREFECO FROM   rndltable WHERE RUNDLSRUNREF = '" & Compress(cmbPrt) & "'" _
          & " AND RUNDLSRUNNO=" & Val(cmbRun)
       
    Debug.Print sSql
    
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   prg1.Value = 100
   Sleep 700
   
End Sub


Private Sub zCreateTimeTable()
   Dim RdoCols As ADODB.Recordset
   
   Dim b As Byte
   Dim iRows As Integer
   Dim sTableN As String
   Dim sCol1 As Variant
   Dim sCol2 As Variant
   Dim sCol3 As Variant
   Dim sCol4 As Variant
   Dim sTable(100) As Variant
   MouseCursor 13
   On Error GoTo DiaErr1
   Err.Clear
   sSql = "sp_columns 'TcitTable'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCols, ES_FORWARD)
   If bSqlRows Then
      With RdoCols
         sTableN = "##TIME" & Right(Compress(GetNextLotNumber()), 8)
         sTimeTemp = Trim$(sTableN)
         Do Until .EOF
            iRows = iRows + 1
            sCol1 = Trim(.Fields(3))
            sCol2 = Trim(.Fields(5))
            sCol3 = Trim(.Fields(7))
            If sCol1 = "" Then Exit Do
            sCol4 = ""
            If iRows > 1 Then sCol1 = "," & sCol1
            If sCol2 = "char" Or sCol2 = "varchar" Then
               sCol3 = "(" & sCol3 & ") Null "
               sCol4 = "default('')"
            Else
               sCol3 = " Null "
               If sCol2 <> "smalldatetime" Then sCol4 = "default(0)"
            End If
            sTable(iRows) = sCol1 & " " & sCol2 & sCol3 & sCol4
            .MoveNext
         Loop
         sTable(iRows) = sTable(iRows) & ")"
         ClearResultSet RdoCols
      End With
   End If
   If iRows > 0 Then
      sTableN = "create table " & sTimeTemp & " ("
      For b = 1 To iRows
         sTableN = sTableN & sTable(b)
      Next
      clsADOCon.ExecuteSql sTableN
      sSql = "create clustered index TimeRef on " & sTimeTemp & " " _
             & "(TCCARD) WITH FILLFACTOR=80"
      clsADOCon.ExecuteSql sSql
   End If
   Set RdoCols = Nothing
   Exit Sub
   
DiaErr1:
   bTablesFailed = 1
   sTimeTemp = ""
   
End Sub


Public Sub zCreatePOitTable()
   Dim RdoCols As ADODB.Recordset
   
   Dim b As Byte
   Dim iRows As Integer
   Dim sTableN As String
   
   MouseCursor 13
   On Error GoTo DiaErr1
   Err.Clear
   
   sTableN = "##POIT" & Right(Compress(GetNextLotNumber()), 8)
   sPoitSrvTemp = Trim$(sTableN)
   Dim strConstraintName As String
   strConstraintName = "POIT" & Right(Compress(GetNextLotNumber()), 8)
   
   sSql = "CREATE TABLE [dbo].[" & sPoitSrvTemp & "](" & vbCrLf
   sSql = sSql & "   [PINUMBER] [int] NOT NULL," & vbCrLf
   sSql = sSql & "   [PIRELEASE] [smallint] NULL," & vbCrLf
   sSql = sSql & "   [PIITEM] [smallint] NOT NULL," & vbCrLf
   sSql = sSql & "   [PIREV] [char](2) NOT NULL," & vbCrLf
   sSql = sSql & "   [PITYPE] [smallint] NULL," & vbCrLf
   sSql = sSql & "   [PIPART] [char](30) NULL," & vbCrLf
   sSql = sSql & "   [PIPDATE] [smalldatetime] NULL," & vbCrLf
   sSql = sSql & "   [PIADATE] [smalldatetime] NULL," & vbCrLf
   sSql = sSql & "   [PIPQTY] [decimal](12, 4) NULL," & vbCrLf
   sSql = sSql & "   [PIAQTY] [decimal](12, 4) NULL," & vbCrLf
   sSql = sSql & "   [PIAMT] [decimal](12, 4) NULL," & vbCrLf
   sSql = sSql & "   [PIESTUNIT] [decimal](12, 4) NULL," & vbCrLf
   sSql = sSql & "   [PIADDERS] [decimal](12, 4) NULL," & vbCrLf
   sSql = sSql & "   [PILOT] [decimal](12, 4) NULL," & vbCrLf
   sSql = sSql & "   [PIRUNPART] [char](30) NULL," & vbCrLf
   sSql = sSql & "   [PIRUNNO] [int] NULL," & vbCrLf
   sSql = sSql & "   [PIRUNOPNO] [smallint] NULL," & vbCrLf
   sSql = sSql & "   [PISN] [tinyint] NULL," & vbCrLf
   sSql = sSql & "   [PISNNO] [char](16) NULL," & vbCrLf
   sSql = sSql & "   [PIFRTADDERS] [decimal](7, 2) NULL," & vbCrLf
   sSql = sSql & "   [PIWIP] [char](4) NULL," & vbCrLf
   sSql = sSql & "   [PIONDOC] [tinyint] NULL," & vbCrLf
   sSql = sSql & "   [PIREJECTED] [decimal](12, 4) NULL," & vbCrLf
   sSql = sSql & "   [PIWASTE] [decimal](12, 3) NULL," & vbCrLf
   sSql = sSql & "   [PIINSBY] [char](4) NULL," & vbCrLf
   sSql = sSql & "   [PIINSDATE] [smalldatetime] NULL," & vbCrLf
   sSql = sSql & "   [PIUSER] [char](4) NULL," & vbCrLf
   sSql = sSql & "   [PIENTERED] [smalldatetime] NULL," & vbCrLf
   sSql = sSql & "   [PIODATE] [smalldatetime] NULL," & vbCrLf
   sSql = sSql & "   [PIRECEIVED] [smalldatetime] NULL," & vbCrLf
   sSql = sSql & "   [PIORIGSCHEDQTY] [decimal](12, 4) NULL," & vbCrLf
   sSql = sSql & "   [PICOMT] [varchar](2048) NULL," & vbCrLf
   sSql = sSql & "   [PILOTNUMBER] [char](15) NULL," & vbCrLf
   sSql = sSql & "   [PIONDOCK] [tinyint] NULL," & vbCrLf
   sSql = sSql & "   [PIONDOCKINSPECTED] [tinyint] NULL," & vbCrLf
   sSql = sSql & "   [PIONDOCKINSPDATE] [smalldatetime] NULL," & vbCrLf
   sSql = sSql & "   [PIONDOCKREJTAG] [char](12) NULL," & vbCrLf
   sSql = sSql & "   [PIONDOCKQTYACC] [decimal](12, 4) NULL," & vbCrLf
   sSql = sSql & "   [PIONDOCKQTYREJ] [decimal](12, 4) NULL," & vbCrLf
   sSql = sSql & "   [PIONDOCKINSPECTOR] [char](30) NULL," & vbCrLf
   sSql = sSql & "   [PIONDOCKCOMMENT] [varchar](255) NULL," & vbCrLf
   sSql = sSql & "   [PIODDELIVERED] [tinyint] NULL," & vbCrLf
   sSql = sSql & "   [PIODDELDATE] [smalldatetime] NULL," & vbCrLf
   sSql = sSql & "   [PIODDELPSNUMBER] [char](20) NULL," & vbCrLf
   sSql = sSql & "   [PIODDELQTY] [decimal](12, 4) NULL," & vbCrLf
   sSql = sSql & "   [PIACCOUNT] [char](12) NULL," & vbCrLf
   sSql = sSql & "   [PIVENDOR] [char](10) NULL," & vbCrLf
   sSql = sSql & "   [PIPICKRECORD] [smallint] NULL," & vbCrLf
   sSql = sSql & "   [PIPRESPLITFROM] [char](6) NULL," & vbCrLf
   sSql = sSql & "   [PIONDOCKQTYWASTE] [decimal](12, 4) NULL," & vbCrLf
   sSql = sSql & "   [PIPORIGDATE] [smalldatetime] NULL," & vbCrLf
   sSql = sSql & " CONSTRAINT [" & strConstraintName & "] PRIMARY KEY CLUSTERED " & vbCrLf
   sSql = sSql & "(" & vbCrLf
   sSql = sSql & "   [PINUMBER] ASC," & vbCrLf
   sSql = sSql & "   [PIITEM] ASC," & vbCrLf
   sSql = sSql & "   [PIREV] ASC" & vbCrLf
   sSql = sSql & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 80) ON [PRIMARY]" & vbCrLf
   sSql = sSql & ") ON [PRIMARY]"
   
   clsADOCon.ExecuteSql sSql
   
   Err.Clear
   Set RdoCols = Nothing
   Exit Sub
   
DiaErr1:
   bTablesFailed = 1
   sInvaTemp = ""
   
End Sub


Public Sub zCreateOutSideTable()
   Dim RdoCols As ADODB.Recordset
   
   Dim b As Byte
   Dim iRows As Integer
   Dim sTableN As String
   Dim sCol1 As Variant
   Dim sCol2 As Variant
   Dim sCol3 As Variant
   Dim sCol4 As Variant
   Dim sCol5 As Variant
   Dim sTable(100) As Variant
   MouseCursor 13
   On Error GoTo DiaErr1
   Err.Clear
   
   sTableN = "##VIIT" & Right(Compress(GetNextLotNumber()), 8)
   sOutSideSrvTemp = Trim$(sTableN)
   
   sSql = "CREATE TABLE [dbo].[" & sOutSideSrvTemp & "](" & vbCrLf
   sSql = sSql & "   [VITNO] [char](20) NOT NULL," & vbCrLf
   sSql = sSql & "   [VITVENDOR] [char](10) NOT NULL," & vbCrLf
   sSql = sSql & "   [VITITEM] [smallint] NOT NULL," & vbCrLf
   sSql = sSql & "   [VITPO] [int] NULL," & vbCrLf
   sSql = sSql & "   [VITPORELEASE] [smallint] NULL," & vbCrLf
   sSql = sSql & "   [VITPOITEM] [smallint] NULL," & vbCrLf
   sSql = sSql & "   [VITPOITEMREV] [char](2) NULL," & vbCrLf
   sSql = sSql & "   [VITQTY] [decimal](15, 4) NULL," & vbCrLf
   sSql = sSql & "   [VITCOST] [decimal](15, 4) NULL," & vbCrLf
   sSql = sSql & "   [VITMO] [char](30) NULL," & vbCrLf
   sSql = sSql & "   [VITMORUN] [int] NULL," & vbCrLf
   sSql = sSql & "   [VITACCOUNT] [char](12) NULL," & vbCrLf
   sSql = sSql & "   [VITNOTE] [char](40) NULL," & vbCrLf
   sSql = sSql & "   [VITCHECKNO] [char](12) NULL," & vbCrLf
   sSql = sSql & "   [VITCHECKDT] [smalldatetime] NULL," & vbCrLf
   sSql = sSql & "   [VITCASHACCOUNT] [char](12) NULL," & vbCrLf
   sSql = sSql & "   [VITDISCOUNT] [decimal](12, 4) NULL," & vbCrLf
   sSql = sSql & "   [VITPAID] [tinyint] NULL," & vbCrLf
   sSql = sSql & "   [VITTOTPAID] [decimal](13, 2) NULL," & vbCrLf
   sSql = sSql & "   [VITADDERS] [decimal](12, 4) NULL," & vbCrLf
   sSql = sSql & " CONSTRAINT [PK_ViidTable_VENDORINVOICE] PRIMARY KEY CLUSTERED " & vbCrLf
   sSql = sSql & "(" & vbCrLf
   sSql = sSql & "   [VITNO] ASC," & vbCrLf
   sSql = sSql & "   [VITVENDOR] ASC," & vbCrLf
   sSql = sSql & "   [VITITEM] ASC" & vbCrLf
   sSql = sSql & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 90) ON [PRIMARY]" & vbCrLf
   sSql = sSql & ") ON [PRIMARY]"
   
   clsADOCon.ExecuteSql sSql
   
   Err.Clear
   Set RdoCols = Nothing
   Exit Sub
   
DiaErr1:
   bTablesFailed = 1
   sInvaTemp = ""
   
End Sub

Private Sub zCreatePickTable()
   Dim RdoCols As ADODB.Recordset
   
   Dim b As Byte
   Dim iRows As Integer
   Dim sTableN As String
   Dim sCol1 As Variant
   Dim sCol2 As Variant
   Dim sCol3 As Variant
   Dim sCol4 As Variant
   Dim sCol5 As Variant
   Dim sTable(100) As Variant
   MouseCursor 13
   On Error GoTo DiaErr1
   Err.Clear
   sSql = "sp_columns 'MopkTable'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCols, ES_FORWARD)
   If bSqlRows Then
      With RdoCols
         sTableN = "##PICK" & Right(Compress(GetNextLotNumber()), 8)
         sPickTemp = Trim$(sTableN)
         Do Until .EOF
            iRows = iRows + 1
            sCol1 = Trim(.Fields(3))
            sCol2 = Trim(.Fields(5))
            sCol3 = Trim(.Fields(7))
            sCol5 = Trim(.Fields(8))
            If sCol1 = "" Then Exit Do
            sCol4 = ""
            If iRows > 1 Then sCol1 = "," & sCol1
            If sCol2 = "char" Or sCol2 = "varchar" Then
               sCol3 = "(" & sCol3 & ") Null "
               sCol4 = "default('')"
            ElseIf (sCol2 = "decimal") Then
               sCol2 = sCol2 & "(" & sCol3 & "," & sCol5 & ")"
               sCol3 = " Null "
               sCol4 = "default(0)"
            Else
               sCol3 = " Null "
               If sCol2 <> "smalldatetime" Then sCol4 = "default(0)"
            End If
            sTable(iRows) = sCol1 & " " & sCol2 & sCol3 & sCol4
            .MoveNext
         Loop
         sTable(iRows) = sTable(iRows) & ")"
         ClearResultSet RdoCols
      End With
   End If
   If iRows > 0 Then
      sTableN = "create table " & sPickTemp & " ("
      For b = 1 To iRows
         sTableN = sTableN & sTable(b)
      Next
      clsADOCon.ExecuteSql sTableN
      sSql = "create unique clustered index PickRef on " & sPickTemp & " " _
             & "(PKMOPART,PKMORUN,PKRECORD) WITH FILLFACTOR=80"
      clsADOCon.ExecuteSql sSql
   End If
   Set RdoCols = Nothing
   Exit Sub
   
DiaErr1:
   bTablesFailed = 1
   sPickTemp = ""
   
End Sub

Public Sub zCreateInvaTable()
   Dim RdoCols As ADODB.Recordset
   
   Dim b As Byte
   Dim iRows As Integer
   Dim sTableN As String
   Dim sCol1 As Variant
   Dim sCol2 As Variant
   Dim sCol3 As Variant
   Dim sCol4 As Variant
   Dim sCol5 As Variant
   Dim sTable(100) As Variant
   MouseCursor 13
   On Error GoTo DiaErr1
   Err.Clear
   sSql = "sp_columns 'InvaTable'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCols, ES_FORWARD)
   If bSqlRows Then
      With RdoCols
         sTableN = "##INVA" & Right(Compress(GetNextLotNumber()), 8)
         sInvaTemp = Trim$(sTableN)
         Do Until .EOF
            iRows = iRows + 1
            sCol1 = Trim(.Fields(3))
            sCol2 = Trim(.Fields(5))
            sCol3 = Trim(.Fields(7))
            sCol5 = Trim(.Fields(8))
            If sCol1 = "" Then Exit Do
            sCol4 = ""
            If iRows > 1 Then sCol1 = "," & sCol1
            If sCol2 = "char" Or sCol2 = "varchar" Then
               sCol3 = "(" & sCol3 & ") Null "
               sCol4 = "default('')"
            ElseIf (sCol2 = "decimal") Then
               sCol2 = sCol2 & "(" & sCol3 & "," & sCol5 & ")"
               sCol3 = " Null "
               sCol4 = "default(0)"
            Else
               sCol3 = " Null "
               If sCol2 <> "smalldatetime" Then sCol4 = "default(0)"
            End If
            
            If sCol1 <> ",INNO" Then
               sTable(iRows) = sCol1 & " " & sCol2 & sCol3 & sCol4
            Else
               sTable(iRows) = sCol1 & " " & sCol2
            End If
            .MoveNext
         Loop
         ClearResultSet RdoCols
         sTable(iRows) = sTable(iRows) & ")"
      End With
   End If
   If iRows > 0 Then
      sTableN = "create table " & sInvaTemp & " (" & vbCrLf
      For b = 1 To iRows
         sTableN = sTableN & sTable(b) & vbCrLf
      Next
      clsADOCon.ExecuteSql sTableN
      
      sSql = "create index InvNum on " & sInvaTemp & " " _
             & "(INNUMBER) WITH FILLFACTOR=80"
      clsADOCon.ExecuteSql sSql
      
      sSql = "create index InvPart on " & sInvaTemp & " " _
             & "(INPART) WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSql sSql
      
      sSql = "CREATE INDEX InvType ON " & sInvaTemp & " " _
             & "(INPART, INTYPE) WITH  FILLFACTOR = 80 "
      clsADOCon.ExecuteSql sSql
   
      ' Disable the Identity column on INNO
      sSql = "SET IDENTITY_INSERT " & sInvaTemp & " OFF"
      clsADOCon.ExecuteSql sSql

   End If
   Err.Clear
   Set RdoCols = Nothing
   Exit Sub
   
DiaErr1:
   bTablesFailed = 1
   sInvaTemp = ""
   
End Sub

Private Sub zCreateLotsTable()
   Dim RdoCols As ADODB.Recordset
   
   Dim b As Byte
   Dim iRows As Integer
   Dim sTableN As String
   Dim sCol1 As Variant
   Dim sCol2 As Variant
   Dim sCol3 As Variant
   Dim sCol4 As Variant
   Dim sCol5 As Variant
   Dim sTable(100) As Variant
   MouseCursor 13
   On Error GoTo DiaErr1
   Err.Clear
   loitColumns = ""
   sSql = "sp_columns 'LoitTable'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCols, ES_FORWARD)
   If bSqlRows Then
      With RdoCols
         sTableN = "##LOTS" & Right(Compress(GetNextLotNumber()), 8)
         'sTableN = "_LOTS" & Right(Compress(GetNextLotNumber()), 8)
         sLotsTemp = Trim$(sTableN)
         Do Until .EOF
            sCol1 = Trim(.Fields(3))
            If sCol1 <> "LOIAREA" Then
               iRows = iRows + 1
               sCol2 = Trim(.Fields(5))
               sCol3 = Trim(.Fields(7))
               sCol5 = Trim(.Fields(8))
               If sCol1 = "" Then Exit Do
               sCol4 = ""
               If iRows > 1 Then
                  sCol1 = "," & sCol1
               End If
               loitColumns = loitColumns & sCol1
               If sCol2 = "char" Or sCol2 = "varchar" Then
                  sCol3 = "(" & sCol3 & ") Null "
                  sCol4 = "default('')"
               ElseIf (sCol2 = "decimal") Then
                  sCol2 = sCol2 & "(" & sCol3 & "," & sCol5 & ")"
                  sCol3 = " Null "
                  sCol4 = "default(0)"
               Else
                  sCol3 = " Null "
                  If sCol2 <> "smalldatetime" Then sCol4 = "default(0)"
               End If
               sTable(iRows) = sCol1 & " " & sCol2 & sCol3 & sCol4
            End If
            .MoveNext
         Loop
         ClearResultSet RdoCols
         sTable(iRows) = sTable(iRows) & ")"
      End With
   End If
   If iRows > 0 Then
      sTableN = "create table " & sLotsTemp & " ("
      For b = 1 To iRows
         sTableN = sTableN & sTable(b)
      Next
      clsADOCon.ExecuteSql sTableN
      sSql = "create clustered index LotsRef on " & sLotsTemp & " " _
             & "(LOINUMBER,LOIRECORD) WITH FILLFACTOR=80"
      clsADOCon.ExecuteSql sSql
   End If
   Set RdoCols = Nothing
   Exit Sub
   
DiaErr1:
   bTablesFailed = 1
   sLotsTemp = ""
   
End Sub

Public Sub zCreateRunsTable()
   Dim RdoCols As ADODB.Recordset
   
   Dim b As Byte
   Dim iRows As Integer
   Dim sTableN As String
   Dim sCol1 As Variant
   Dim sCol2 As Variant
   Dim sCol3 As Variant
   Dim sCol4 As Variant
   Dim sCol5 As Variant
   
   Dim sTable(100) As Variant
   MouseCursor 13
   On Error GoTo DiaErr1
   Err.Clear
   sSql = "sp_columns 'RunsTable'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCols, ES_FORWARD)
   If bSqlRows Then
      With RdoCols
         sTableN = "##RUNS" & Right(Compress(GetNextLotNumber()), 8)
         sRunsTemp = Trim$(sTableN)
         Do Until .EOF
            iRows = iRows + 1
            sCol1 = Trim(.Fields(3))
            sCol2 = Trim(.Fields(5))
            sCol3 = Trim(.Fields(7))
            sCol5 = Trim(.Fields(8))
            If sCol1 = "" Then Exit Do
            sCol4 = ""
            If iRows > 1 Then sCol1 = "," & sCol1
            If sCol2 = "char" Or sCol2 = "varchar" Then
               sCol3 = "(" & sCol3 & ") Null "
               sCol4 = "default('')"
            ElseIf (sCol2 = "decimal") Then
               sCol2 = sCol2 & "(" & sCol3 & "," & sCol5 & ")"
               sCol3 = " Null "
               sCol4 = "default(0)"
            Else
               sCol3 = " Null "
               If sCol2 <> "smalldatetime" Then sCol4 = "default(0)"
            End If
            sTable(iRows) = sCol1 & " " & sCol2 & sCol3 & sCol4
            .MoveNext
         Loop
         ClearResultSet RdoCols
         sTable(iRows) = sTable(iRows) & ")"
      End With
   End If
   If iRows > 0 Then
      sTableN = "create table " & sRunsTemp & " ("
      For b = 1 To iRows
         sTableN = sTableN & sTable(b)
      Next
      clsADOCon.ExecuteSql sTableN
      sSql = "create unique clustered index RunRef on " & sRunsTemp & " " _
             & "(RUNREF,RUNNO) WITH FILLFACTOR=80"
      clsADOCon.ExecuteSql sSql
   End If
   Set RdoCols = Nothing
   Exit Sub
   
DiaErr1:
   bTablesFailed = 1
   sRunsTemp = ""
   
End Sub

Private Sub zCreateRnopTable()
   Dim RdoCols As ADODB.Recordset
   
   Dim b As Byte
   Dim iRows As Integer
   Dim sTableN As String
   Dim sCol1 As Variant
   Dim sCol2 As Variant
   Dim sCol3 As Variant
   Dim sCol4 As Variant
   Dim sTable(100) As Variant
   MouseCursor 13
   On Error GoTo DiaErr1
   Err.Clear
   sSql = "sp_columns 'RnopTable'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCols, ES_FORWARD)
   If bSqlRows Then
      With RdoCols
         sTableN = "##RNOP" & Right(Compress(GetNextLotNumber()), 8)
         sRnopTemp = Trim$(sTableN)
         Do Until .EOF
            iRows = iRows + 1
            sCol1 = Trim(.Fields(3))
            sCol2 = Trim(.Fields(5))
            sCol3 = Trim(.Fields(7))
            If sCol1 = "" Then Exit Do
            sCol4 = ""
            If iRows > 1 Then sCol1 = "," & sCol1
            If sCol2 = "char" Or sCol2 = "varchar" Then
               sCol3 = "(" & sCol3 & ") Null "
               sCol4 = "default('')"
            Else
               sCol3 = " Null "
               If sCol2 <> "smalldatetime" Then sCol4 = "default(0)"
            End If
            sTable(iRows) = sCol1 & " " & sCol2 & sCol3 & sCol4
            .MoveNext
         Loop
         ClearResultSet RdoCols
         sTable(iRows) = sTable(iRows) & ")"
      End With
   End If
   If iRows > 0 Then
      sTableN = "create table " & sRnopTemp & " ("
      For b = 1 To iRows
         sTableN = sTableN & sTable(b)
      Next
      clsADOCon.ExecuteSql sTableN
      sSql = "create unique clustered index OpRef on " & sRnopTemp & " " _
             & "(OPREF,OPRUN,OPNO) WITH FILLFACTOR=80"
      clsADOCon.ExecuteSql sSql
   End If
   Set RdoCols = Nothing
   Exit Sub
   
DiaErr1:
   bTablesFailed = 1
   sRnopTemp = ""
   
End Sub

'tested 2/11/05

Private Function CreateManufacturingOrder() As Byte
   Dim iNewOrderQty As Integer
   
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   On Error Resume Next
   iNewOrderQty = Val(txtQty)
   sSql = "INSERT " & sRunsTemp & " SELECT * FROM RunsTable WHERE " _
          & "RUNREF='" & Compress(cmbPrt) & "' AND RUNNO=" _
          & Val(cmbRun) & " "
   clsADOCon.ExecuteSql sSql
   
   sSql = "UPDATE " & sRunsTemp & " SET RUNREF='" & Compress(lblNew) _
          & "',RUNNO=" & Val(txtRun) & " "
   clsADOCon.ExecuteSql sSql
   
   sSql = "INSERT RunsTable SELECT * FROM " & sRunsTemp & " WHERE " _
          & "RUNREF='" & Compress(lblNew) & "' AND RUNNO=" _
          & Val(txtRun) & " "
   clsADOCon.ExecuteSql sSql
   
   'Update rows
   sSql = "UPDATE RunsTable SET RUNQTY=RUNQTY-" & iNewOrderQty & "," _
          & "RUNREMAININGQTY=RUNREMAININGQTY-" & iNewOrderQty _
          & ",RUNLASTSPLITREF='" & Compress(lblNew) & "'," _
          & "RUNLASTSPLITRUNNO=" & Val(txtRun) & " " _
          & "WHERE RUNREF='" & Compress(cmbPrt) & "' AND RUNNO=" _
          & Val(cmbRun) & " "
   clsADOCon.ExecuteSql sSql
   
   sSql = "UPDATE RunsTable SET RUNQTY=" & iNewOrderQty & "," _
          & "RUNREMAININGQTY=" & iNewOrderQty _
          & ",RUNSPLITFROMREF='" & Compress(cmbPrt) & "'," _
          & "RUNSPLITFROMRUNNO=" & Val(cmbRun) & " " _
          & "WHERE RUNREF='" & Compress(lblNew) & "' AND RUNNO=" _
          & Val(txtRun) & " "
   clsADOCon.ExecuteSql sSql
   
   'Ops
   sSql = "INSERT " & sRnopTemp & " SELECT * FROM RnopTable WHERE " _
          & "OPREF='" & Compress(cmbPrt) & "' AND OPRUN=" _
          & Val(cmbRun) & " "
   clsADOCon.ExecuteSql sSql
   
   sSql = "UPDATE " & sRnopTemp & " SET OPREF='" & Compress(lblNew) _
          & "',OPRUN=" & Val(txtRun) & " "
   clsADOCon.ExecuteSql sSql
   
   sSql = "INSERT RnopTable SELECT * FROM " & sRnopTemp & " WHERE " _
          & "OPREF='" & Compress(lblNew) & "' AND OPRUN=" _
          & Val(txtRun) & " "
   clsADOCon.ExecuteSql sSql
   
   'Update Ops
   sSql = "UPDATE RnopTable SET OPYIELD=OPYIELD-" & iNewOrderQty _
          & " WHERE (OPYIELD>" & iNewOrderQty & " AND OPREF='" _
          & Compress(cmbPrt) & "' AND OPRUN=" & Val(cmbRun) & ")"
   clsADOCon.ExecuteSql sSql
   
   sSql = "UPDATE RnopTable SET OPYIELD=" & iNewOrderQty _
          & " WHERE (OPYIELD>0 AND OPREF='" & Compress(lblNew) & "' AND " _
          & "OPRUN=" & Val(txtRun) & ")"
   clsADOCon.ExecuteSql sSql
   
   'Insert Tracking
   sSql = "INSERT INTO RnspTable (SPLIT_TORUNREF,SPLIT_TORUNRUNNO," _
          & "SPLIT_FROMRUNREF,SPLIT_FROMRUNRUNNO,SPLIT_SPLQTY," _
          & "SPLIT_SPLORIGQTY,SPLIT_SPUSER) VALUES('" _
          & Compress(lblNew) & "'," & Val(txtRun) & ",'" _
          & Compress(cmbPrt) & "'," & Val(cmbRun) & "," _
          & iNewOrderQty & "," & Val(lblRemQty) & ",'" _
          & sInitials & "')"
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum <> 0 Then
      CreateManufacturingOrder = 0
      clsADOCon.RollbackTrans
   Else
      CreateManufacturingOrder = 1
      clsADOCon.CommitTrans
   End If
   
End Function
