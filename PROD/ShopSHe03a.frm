VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form ShopSHe03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Operation Completions/Assignments"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHe03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optOps 
      Caption         =   " (Otherwise All Operations Will Be Shown)"
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   2350
      Value           =   1  'Checked
      Width           =   3495
   End
   Begin VB.CheckBox optLoaded 
      Height          =   255
      Left            =   5880
      TabIndex        =   10
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox optCmt 
      Caption         =   "           "
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   2040
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.ComboBox cmbRun 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   5640
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Select Run Number"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      ToolTipText     =   "Select Part Number"
      Top             =   1080
      Width           =   3545
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6360
      TabIndex        =   5
      Top             =   0
      Width           =   915
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Operations"
      Height          =   315
      Left            =   6360
      TabIndex        =   4
      ToolTipText     =   "Retrieve Manufacturing Order Operations"
      Top             =   600
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   2280
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3000
      FormDesignWidth =   7320
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   12
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Only Incomplete Ops"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   2350
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   3
      Left            =   5160
      TabIndex        =   9
      Top             =   1110
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operation Comments?"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   1110
      Width           =   1335
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   1440
      Width           =   3255
   End
End
Attribute VB_Name = "ShopSHe03a"
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
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter
Dim bCancel As Boolean
Dim bGoodPart As Byte
Dim bOnLoad As Byte
Dim bGoodRuns As Byte
Dim bGoodRun As Byte

Dim sPartNumber As String

Private txtKeyPress(2) As New EsiKeyBd
Private txtGotFocus(2) As New EsiKeyBd

Private Sub FormatControls()
   On Error Resume Next
   Set txtGotFocus(0).esCmbGotfocus = cmbPrt
   Set txtGotFocus(1).esCmbGotfocus = cmbRun
   
   Set txtKeyPress(0).esCmbKeyCase = cmbPrt
   Set txtKeyPress(1).esCmbKeyValue = cmbRun
   
End Sub

Private Function GetRuns() As Byte
   Dim RdoRns As ADODB.Recordset
   On Error GoTo DiaErr1
   cmbRun.Clear
   sPartNumber = Compress(cmbPrt)
   AdoQry.Parameters(0).Value = sPartNumber
   bSqlRows = clsADOCon.GetQuerySet(RdoRns, AdoQry)
   If bSqlRows Then
      With RdoRns
         cmbRun = Format(!Runno, "####0")
         Do Until .EOF
            AddComboStr cmbRun.hwnd, Format$(!Runno, "####0")
            .MoveNext
         Loop
         ClearResultSet RdoRns
      End With
      GetRuns = True
      If GetPreferenceValue("AutoSelectLastRun") = "1" Then cmbRun = cmbRun.List(cmbRun.ListCount - 1)
   Else
      sPartNumber = ""
      GetRuns = False
   End If
   Set RdoRns = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub cmbPrt_Click()
   bGoodPart = GetPart()
   
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If Not bCancel Then bGoodPart = GetPart()
   
End Sub




Private Sub cmbRun_LostFocus()
   bGoodRun = GetThisRun
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = True
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4103
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub cmdOk_Click()
   bGoodRun = GetThisRun
   If bGoodRuns And bGoodRun Then
      MouseCursor 13
      optLoaded.Value = vbChecked
      On Error Resume Next
      sPartNumber = Compress(cmbPrt)
      ShopSHe03b.lblPrt = sPartNumber
      If Err > 0 Then
         MouseCursor 0
         Exit Sub
      End If
      ShopSHe03b.lblRun = str(cmbRun)
      ShopSHe03b.Caption = ShopSHe03b.Caption & " " & cmbPrt & " Run " & str(cmbRun)
      ShopSHe03b.Show
      
   Else
      MsgBox "Part, Run Combination Wasn't Found.", vbExclamation, Caption
      On Error Resume Next
      optLoaded.Value = False
      cmbPrt.SetFocus
   End If
   
End Sub

Private Sub Form_Activate()
   MouseCursor 0
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillRuns Me, "NOT LIKE 'C%'"
      bGoodPart = GetPart()
      bOnLoad = 0
   End If
   If optLoaded = vbChecked Then
      Unload ShopSHe03b
      optLoaded = vbUnchecked
   End If
   
End Sub

Private Sub Form_Load()

'   Dim S As Single
'   Dim l As Long
'   Dim m As Long
'   'Dim T As String
'   'Dim timeCardNo As String
'   Dim loginDate As String
'   m = DateValue(Format(GetServerDateTime, "yyyy,mm,dd"))
'   S = TimeValue(Format(GetServerDateTime, "hh:mm:ss"))
'   l = S * 1000000
'   Dim GetTimeCardID As String
'   GetTimeCardID = Format(m, "00000") & Format(l, "000000")
'
   
   FormLoad Me
   FormatControls
   optCmt.Value = GetSetting("Esi2000", "EsiProd", "scoop", optCmt.Value)
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PARUN,RUNREF,RUNSTATUS," _
          & "RUNNO FROM PartTable,RunsTable WHERE PARTREF= ? " _
          & "AND PARTREF=RUNREF AND RUNSTATUS NOT LIKE 'C%'"
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Type = adChar
   AdoParameter.SIZE = 30
   
   AdoQry.Parameters.Append AdoParameter
   bOnLoad = 1
   
End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveSetting "Esi2000", "EsiProd", "scoop", optCmt.Value
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   If optLoaded.Value = 1 Then Unload ShopSHe03b Else FormUnload
   Set ShopSHe03a = Nothing
   
End Sub





Private Function GetPart() As Byte
   Dim RdoPrt As ADODB.Recordset
   
   optLoaded.Value = False
   sPartNumber = Compress(cmbPrt)
   On Error Resume Next
   If Len(sPartNumber) > 0 Then
      sSql = "SELECT PARTREF,PARTNUM,PADESC FROM PartTable WHERE PARTREF='" & sPartNumber & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
      If bSqlRows Then
         With RdoPrt
            cmbPrt = "" & Trim(!PartNum)
            lblDsc = "" & !PADESC
            GetPart = True
            ClearResultSet RdoPrt
         End With
      Else
         MsgBox "Part Wasn't Found.", vbExclamation, Caption
         cmbPrt = ""
         lblDsc = ""
         GetPart = False
      End If
      On Error Resume Next
      RdoPrt.Close
   Else
      sPartNumber = ""
      cmbPrt = ""
   End If
   
   Set RdoPrt = Nothing
   
   If GetPart Then bGoodRuns = GetRuns()
   Exit Function
   
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function






Private Function GetThisRun()
   Dim RdoRun As ADODB.Recordset
   sSql = "SELECT RUNREF,RUNNO FROM RunsTable WHERE " _
          & "RUNREF='" & Compress(cmbPrt) & "' AND " _
          & "RUNNO=" & Val(cmbRun) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRun, ES_FORWARD)
   If bSqlRows Then
      GetThisRun = 1
   Else
      GetThisRun = 0
      MsgBox "You Must Select Or Enter A Valid Run.", _
         vbInformation, Caption
   End If
   
   Set RdoRun = Nothing
End Function
