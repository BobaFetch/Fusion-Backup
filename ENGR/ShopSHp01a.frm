VERSION 5.00
Begin VB.Form ShopSHp01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Engineering Time Charges"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Add"
      Height          =   360
      Left            =   2280
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1065
   End
   Begin VB.TextBox Text2 
      Height          =   3135
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   18
      Text            =   "ShopSHp01a.frx":0000
      Top             =   2880
      Width           =   7455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   15
      Tag             =   "2"
      ToolTipText     =   "Tool Description (40 chars)"
      Top             =   2160
      Width           =   585
   End
   Begin VB.TextBox txtCGSPOP 
      Height          =   285
      Left            =   1440
      TabIndex        =   11
      Tag             =   "2"
      ToolTipText     =   "Tool Description (40 chars)"
      Top             =   240
      Width           =   1065
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1440
      TabIndex        =   9
      Tag             =   "4"
      Top             =   960
      Width           =   1335
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Contains Part Numbers With Manufacturing Orders"
      Top             =   1440
      Width           =   3545
   End
   Begin VB.ComboBox cmbRun 
      Height          =   315
      Left            =   5400
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Contains Part Numbers With Manufacturing Orders"
      Top             =   1440
      Width           =   1140
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   360
      Left            =   6000
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Recent Charges"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   17
      Top             =   2520
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Hours"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   16
      Top             =   2160
      Width           =   795
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   14
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   13
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PIN"
      Height          =   285
      Index           =   28
      Left            =   240
      TabIndex        =   12
      Top             =   240
      Width           =   795
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   7
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type/Status"
      Height          =   255
      Index           =   15
      Left            =   5400
      TabIndex        =   5
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lblTyp 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6480
      TabIndex        =   4
      Top             =   1800
      Width           =   300
   End
   Begin VB.Label lblSta 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7080
      TabIndex        =   3
      Top             =   1800
      Width           =   495
   End
End
Attribute VB_Name = "ShopSHp01a"
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
' 3/25/04 Removed Jet tables and reorged prdshcvr.rpt
' 3/28/05 Revamped the Cover Sheet and formatting.
Option Explicit
Dim rdoQry As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter

'Dim DbDoc   As Recordset 'Jet
'Dim DbPls   As Recordset 'Jet
Dim bPrinting As Boolean

Dim bGoodPart As Byte
Dim bGoodMo As Byte
Dim bOnLoad As Byte
Dim bTablesCreated As Byte
Dim bUserTypedRun As Byte

Dim sBomRev As String
Dim sRunPkstart As String
Dim sPartNumber As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub GetOptions()
   Dim iList As Integer
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "sh01", "000000000")
'   chkComments = Mid(sOptions, 1, 1)
'   chkTime = Mid(sOptions, 2, 1)
'   chkSvcs = Mid(sOptions, 3, 1)
'   chkSoAlloc = Mid(sOptions, 4, 1)
'   chkDoc = Mid(sOptions, 5, 1)
'   chkBOM = Mid(sOptions, 6, 1)
'   chkPickList = Mid(sOptions, 7, 1)
'   chkBudget = Mid(sOptions, 8, 1)
'   chkShowToolList = Mid(sOptions, 9, 1)
'   chkDocList = Mid(sOptions, 10, 1)
'   If Len(sOptions) > 10 Then chkShowIntCmt = Mid(sOptions, 11, 1) Else chkShowIntCmt.Value = 0
'
'   'chkSoAlloc.Value = GetSetting("Esi2000", "EsiProd", "sh01all", chkSoAlloc.Value)
'   lblPrinter = GetSetting("Esi2000", "EsiProd", "sh01Printer", lblPrinter)
'   If lblPrinter = "" Then lblPrinter = "Default Printer"
   
End Sub


Private Sub SaveOptions()
   Dim iList As Integer
   Dim sOptions As String
   
'   sOptions = CStr(chkComments) & CStr(chkTime) & CStr(chkSvcs) & CStr(chkSoAlloc) _
'      & CStr(chkDoc) & CStr(chkBOM) & CStr(chkPickList) & CStr(chkBudget) & CStr(chkShowToolList) & CStr(chkDocList) & CStr(chkShowIntCmt) & "000000000"
   SaveSetting "Esi2000", "EsiProd", "sh01", Trim(sOptions)
'   SaveSetting "Esi2000", "EsiProd", "sh01all", Trim(chkSoAlloc.Value)
'   SaveSetting "Esi2000", "EsiProd", "sh01Printer", lblPrinter
   
End Sub



Private Sub cmbRun_Click()
   GetThisRun
End Sub


Private Sub cmbRun_KeyDown(KeyCode As Integer, Shift As Integer)
    bUserTypedRun = 1
End Sub

Private Sub cmbRun_LostFocus()
   cmbRun = CheckLen(cmbRun, 5)
   If Val(cmbRun) > 32767 Then cmbRun = "32767"
   cmbRun = Format(Abs(Val(cmbRun)), "####0")
   GetThisRun
   
End Sub


Private Sub Form_Activate()
   If bOnLoad Then
      'FillAllRuns cmbPrt
      bGoodPart = GetRuns()
      bOnLoad = 0
      bPrinting = False
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   bUserTypedRun = 0
   
   GetOptions
   bTablesCreated = 0
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PABOMREV,PARUN,RUNREF,RUNSTATUS," _
          & "RUNNO FROM PartTable,RunsTable WHERE PARTREF= ? " _
          & "AND PARTREF=RUNREF "
   Set rdoQry = New ADODB.Command
   rdoQry.CommandText = sSql
   
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adChar
   AdoParameter1.Size = 30
   rdoQry.Parameters.Append AdoParameter1
   
   bOnLoad = 1
   
End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Set AdoParameter1 = Nothing
   Set rdoQry = Nothing
   FormUnload
   Set ShopSHp01a = Nothing
   
End Sub




Private Function GetRuns() As Byte
   Dim RdoRns As ADODB.Recordset
   Dim iOriginalRun As Integer
   Dim bOriginalRunFound As Byte
   
   bOriginalRunFound = 0
   
   On Error GoTo DiaErr1
   iOriginalRun = Val(cmbRun)
   MouseCursor 13
   cmbRun.Clear
   sPartNumber = Compress(cmbPrt)
   rdoQry.Parameters(0).Value = sPartNumber
'   rdoQry(0) = sPartNumber
   bSqlRows = clsADOCon.GetQuerySet(RdoRns, rdoQry)
   If bSqlRows Then
      With RdoRns
         cmbRun = Format(!Runno, "####0")
         lblDsc = "" & Trim(!PADESC)
         lblTyp = Format(!PALEVEL, "#")
         Do Until .EOF
            AddComboStr cmbRun.hwnd, Format$(!Runno, "####0")
            If iOriginalRun = !Runno Then bOriginalRunFound = 1
            .MoveNext
         Loop
         ClearResultSet RdoRns
      End With
      cmbRun = cmbRun.List(cmbRun.ListCount - 1)
      GetRuns = True
      GetThisRun
   Else
      sPartNumber = ""
      GetRuns = False
   End If
   MouseCursor 0
   Set RdoRns = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub lblQty_Click()
   'run qty
   
End Sub

Private Sub chkbudget_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub GetThisRun()
   Dim RdoRun As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "SELECT RUNREF,RUNNO,RUNSTATUS,RUNPKSTART,RUNQTY FROM RunsTable WHERE " _
          & "RUNREF='" & Compress(cmbPrt) & "' AND " _
          & "RUNNO=" & cmbRun & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRun, ES_FORWARD)
   If bSqlRows Then
      With RdoRun
         lblSta = "" & Trim(!RUNSTATUS)
         If Not IsNull(!RUNPKSTART) Then
            sRunPkstart = Format(!RUNPKSTART, "mm/dd/yy")
         Else
            sRunPkstart = Format(ES_SYSDATE, "mm/dd/yy")
         End If
         ClearResultSet RdoRun
      End With
   End If
   Set RdoRun = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getthisrun"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub txtBeg_DropDown()
   ShowCalendar Me
End Sub

