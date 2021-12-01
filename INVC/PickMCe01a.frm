VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form PickMCe01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pick Items"
   ClientHeight    =   2475
   ClientLeft      =   2475
   ClientTop       =   645
   ClientWidth     =   6345
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PickMCe01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtTyp 
      Height          =   285
      Left            =   5400
      MaxLength       =   1
      TabIndex        =   13
      Tag             =   "1"
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.OptionButton optInd 
      Caption         =   "&Individual  "
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   1920
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton optExp 
      Caption         =   "&Exceptions  "
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CheckBox optItm 
      Caption         =   "picks"
      Height          =   255
      Left            =   5160
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdPck 
      Caption         =   "&Items"
      Height          =   315
      Left            =   5400
      TabIndex        =   2
      ToolTipText     =   "Show Pick List Items"
      Top             =   480
      Width           =   915
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      ToolTipText     =   "Select Part Number"
      Top             =   720
      Width           =   3545
   End
   Begin VB.ComboBox cmbRun 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      ToolTipText     =   "Select Run Number"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5400
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   2280
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2475
      FormDesignWidth =   6345
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type"
      Height          =   285
      Index           =   3
      Left            =   4440
      TabIndex        =   14
      Top             =   1440
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run Number"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblStat 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3720
      TabIndex        =   6
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status (PL, PP)"
      Height          =   255
      Index           =   15
      Left            =   2520
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "PickMCe01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'12/7/04 changed Lots - See procedure in PickMCe01c
Option Explicit
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter
Dim bGoodRuns As Byte
Dim bGoodMo As Byte
Public bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'when calling from other forms, place MO info here
Public PassedInMoPartNo As String
Public PassedInMoRunNo As Integer

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub cmbPrt_Click()
   bGoodRuns = GetRuns()
   
End Sub


Private Sub cmbPrt_GotFocus()
   cmbPrt_Click
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   
   If (Not ValidPartNumber(cmbPrt.Text) And cmbPrt.Text <> "") Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If
   
   If cmbPrt <> "" Then bGoodRuns = GetRuns()
   
End Sub


Private Sub cmbRun_Click()
   bGoodMo = GetPart()
   
End Sub

Private Sub cmbRun_GotFocus()
   cmbRun_Click
   
End Sub


Private Sub cmbRun_LostFocus()
   cmbRun = CheckLen(cmbRun, 5)
   cmbRun = Format(Abs(Val(cmbRun)), "####0")
   bGoodMo = GetPart()
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmbPrt = ""
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      If Left(Caption, 3) <> "Rev" Then
         OpenHelpContext "5203"
      Else
         OpenHelpContext "5201"
      End If
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub cmdPck_Click()
   If Trim(cmbRun) = "" Then
      MsgBox "run number not specified"
      Exit Sub
   End If

   If Not bGoodMo Then
      MsgBox "Run (PL or PP) For This Part Wasn't Found.", vbInformation, Caption
      Exit Sub
   Else
      optItm.Value = vbChecked
      If Left(Caption, 3) = "Rev" Then
         
         PickMCe01b.cbfrom1a = vbChecked
         PickMCe01b.lblMon = cmbPrt
         PickMCe01b.lblRun = cmbRun
         
         PickMCe01b.txtTyp = txtTyp
         PickMCe01b.lblStat = lblStat
         PickMCe01b.Show
      Else
         PickMCe01c.Show
      End If
   End If
   
End Sub

Private Sub Form_Activate()
   Dim b As Byte
   Dim iList As Integer
   
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      'Can't find the caption
      iList = SetRecent(Me)
      If sPassedMo <> "" Then cmbPrt = sPassedMo
      sJournalID = GetOpenJournal("IJ", Format$(ES_SYSDATE, "mm/dd/yy"))
      If Left(sJournalID, 4) = "None" Then
         sJournalID = ""
         b = 1
      Else
         If sJournalID = "" Then b = 0 Else b = 1
      End If
      If b = 0 Then
         MouseCursor 0
         MsgBox "There Is No Open Inventory Journal For This Period.", _
            vbExclamation, Caption
         Sleep 500
         Unload Me
         Exit Sub
      End If
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   sSql = "SELECT RUNREF,RUNSTATUS,RUNNO FROM " _
          & "RunsTable WHERE RUNREF = ? " _
          & "AND (RUNSTATUS='PL' OR RUNSTATUS='PP')"
   'Set RdoQry = RdoCon.CreateQuery("", sSql)
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Size = 30
   AdoParameter.Type = adChar
   
   AdoQry.Parameters.Append AdoParameter
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If optItm = vbChecked Then
      If Left(Caption, 3) = "Rev" Then
         Unload PickMCe01b
      Else
         Unload PickMCe01c
      End If
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   SaveCurrentSelections
   FormUnload
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   Set PickMCe01a = Nothing
   
End Sub



Public Sub FillCombo()
   On Error GoTo DiaErr1
   cmbPrt.Clear
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,PADESC,RUNREF " _
          & " FROM PartTable,RunsTable WHERE PARTREF=RUNREF AND PAINACTIVE = 0 AND PAOBSOLETE = 0 " _
          & " AND (RUNSTATUS='PL' OR RUNSTATUS='PP') ORDER BY PARTREF"
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then
      cmbPrt = cmbPrt.List(0)
      GetPart
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
   
   cmbRun.Clear
   On Error GoTo DiaErr1
   AdoQry.Parameters(0).Value = Compress(cmbPrt)
   bSqlRows = clsADOCon.GetQuerySet(RdoRns, AdoQry)
   If bSqlRows Then
      With RdoRns
         cmbRun = Format(!Runno, "####0")
         lblStat = "" & !RUNSTATUS
         Do Until .EOF
            AddComboStr cmbRun.hWnd, Format$(!Runno, "####0")
            .MoveNext
         Loop
         ClearResultSet RdoRns
      End With
      GetRuns = True
   Else
      GetRuns = False
   End If
   If GetRuns Then bGoodMo = GetPart()
   Set RdoRns = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function GetPart() As Byte
   Dim RdoPrt As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,PADESC,PALEVEL,RUNREF,RUNSTATUS " _
          & "FROM PartTable,RunsTable WHERE PARTREF=RUNREF " _
          & "AND PARTREF='" & Compress(cmbPrt) & "' AND RUNNO=" & str(Val(cmbRun)) & " " _
          & "AND (RUNSTATUS='PL' OR RUNSTATUS='PP')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt)
   If bSqlRows Then
      With RdoPrt
         cmbPrt = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         lblStat = "" & Trim(!RUNSTATUS)
         txtTyp = Format(!PALEVEL, "0")
         cUR.CurrentPart = cmbPrt
         ClearResultSet RdoPrt
         GetPart = True
      End With
   Else
      GetPart = False
   End If
   Set RdoPrt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub optExp_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optInd_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optItm_Click()
   'never visible. marks PickMCe01c as loaded
   
End Sub
