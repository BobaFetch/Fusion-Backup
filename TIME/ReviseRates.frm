VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form ReviseRates 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revise Time Charge Pay Rates"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPayRate 
      Height          =   285
      Left            =   1860
      TabIndex        =   3
      Tag             =   "2"
      Top             =   1800
      Width           =   915
   End
   Begin VB.ComboBox cboEnd 
      Height          =   315
      Left            =   1860
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1380
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   315
      Left            =   5280
      TabIndex        =   10
      Top             =   600
      Width           =   875
   End
   Begin VB.ComboBox cboEmp 
      Height          =   315
      Left            =   1860
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Select From List Or Enter Number"
      Top             =   180
      Width           =   1095
   End
   Begin VB.ComboBox cboStart 
      Height          =   315
      Left            =   1860
      TabIndex        =   1
      Tag             =   "4"
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   4
      Top             =   90
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   5
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
      PictureUp       =   "ReviseRates.frx":0000
      PictureDn       =   "ReviseRates.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5880
      Top             =   1680
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2250
      FormDesignWidth =   6345
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revised Pay Rate"
      Height          =   255
      Index           =   4
      Left            =   420
      TabIndex        =   12
      Top             =   1800
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      Height          =   255
      Index           =   3
      Left            =   420
      TabIndex        =   11
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
      Height          =   255
      Index           =   2
      Left            =   420
      TabIndex        =   9
      Top             =   180
      Width           =   1575
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1860
      TabIndex        =   8
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Index           =   0
      Left            =   420
      TabIndex        =   7
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   255
      Index           =   1
      Left            =   420
      TabIndex        =   6
      Top             =   1020
      Width           =   975
   End
End
Attribute VB_Name = "ReviseRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit

Dim bOnLoad As Byte
Dim bGoodCard As Byte
Dim sCardNumber As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cboEmp_Click()
   GetEmployee
   
End Sub


Private Sub cboEmp_KeyUp(KeyCode As Integer, Shift As Integer)
   cboEmp = CheckLen(cboEmp, 6)
   If Len(cboEmp) Then
      cboEmp = Format(cboEmp, "000000")
      GetEmployee
   End If
   
End Sub


Private Sub cboEnd_DropDown()
   ShowCalendar Me
End Sub

Private Sub cboEnd_LostFocus()
   cboEnd = CheckDate(cboEnd)
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "hs1550"
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdUpdate_Click()
   
   UpdateTimeCharges
End Sub

Private Sub UpdateTimeCharges()
   'validate employee number
   If Not IsNumeric(cboEmp.Text) Then
      MsgBox "Invalid employee number"
      Exit Sub
   End If
   
   'validate pay rate
   Dim cUR As String
   cUR = CheckCurrency(txtPayRate.Text, False)
   If cUR <> "*" Then
      txtPayRate = cUR
   Else
      txtPayRate.SetFocus
      Exit Sub
   End If
   
   'validate date range
   On Error Resume Next
   Dim dtStart As Variant
   Dim dtEnd As Variant
   Dim days As Long
   dtStart = CheckDate(cboStart.Text)
   dtEnd = CheckDate(cboEnd.Text)
   If Err Then
      MsgBox "Invalid start or end date"
      Exit Sub
   End If
   If IsDate(dtStart) And IsDate(dtEnd) Then
      days = DateDiff("d", dtStart, dtEnd) + 1
      If days <= 0 Then
         MsgBox "Invalid date range: from " & cboStart.Text & " to " & cboEnd.Text
         Exit Sub
      End If
   End If
   
   
   On Error GoTo DiaErr1
   
   ' find time charges in the desired date range where
   ' there is no open journal.  if there are any, we can't proceed
   sSql = "select distinct card.TMDAY from TchdTable card" & vbCrLf _
          & "join TcitTable chg on chg.TCCARD = card.TMCARD" & vbCrLf _
          & "join JrhdTable jnl on jnl.MJSTART <= card.TMDAY and jnl.MJEND >= card.TMDAY" & vbCrLf _
          & "and jnl.MJTYPE = 'TJ' and jnl.MJCLOSED is not null" & vbCrLf _
          & "where chg.TCEMP = " & Format(cboEmp.Text, "0") & vbCrLf _
          & "and card.TMDAY >= '" & cboStart.Text & "'" & vbCrLf _
          & "and card.TMDAY <= '" & cboEnd.Text & "'"
   Debug.Print sSql
   Dim rdo As ADODB.Recordset
   
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD)
   If bSqlRows Then
      Dim sMsg
      Dim count As Integer
      count = 0
      sMsg = "Time Journal is closed for charges on the following dates: "
      With rdo
         Do Until .EOF
            count = count + 1
            If count > 1 Then
               sMsg = sMsg & ", "
            End If
            sMsg = sMsg & Format(.Fields(0), "mm/dd/yy")
            .MoveNext
         Loop
      End With
      MsgBox sMsg
      Set rdo = Nothing
      Exit Sub
   End If
   Set rdo = Nothing
   
   'now update any time charges in this date range for this employee
   'Update TcitTable
   'Set TCRATE = 25 * code.TYPEADDER
   'from TchdTable card
   'join TcitTable chg on chg.TCCARD = card.TMCARD
   'join JrhdTable jnl on jnl.MJSTART <= card.TMDAY and jnl.MJEND >= card.TMDAY
   'and jnl.MJTYPE = 'TJ' and jnl.MJCLOSED is null
   'join tmcdTable code on chg.TCCODE = code.TYPECODE
   'Where chg.TCEMP = 1
   'and card.TMDAY >= '4/9/05'
   'and card.TMDAY <= '4/9/05'
   sSql = "Update TcitTable" & vbCrLf _
          & "set TCRATE = " & txtPayRate.Text & " * code.TYPEADDER" & vbCrLf _
          & "from TchdTable card" & vbCrLf _
          & "join TcitTable chg on chg.TCCARD = card.TMCARD" & vbCrLf _
          & "join JrhdTable jnl on jnl.MJSTART <= card.TMDAY and jnl.MJEND >= card.TMDAY" & vbCrLf _
          & "and jnl.MJTYPE = 'TJ' and jnl.MJCLOSED is null" & vbCrLf _
          & "join tmcdTable code on chg.TCCODE = code.TYPECODE" & vbCrLf _
          & "where chg.TCEMP = " & Format(cboEmp.Text, "0") & vbCrLf _
          & "and card.TMDAY >= '" & cboStart.Text & "'" & vbCrLf _
          & "and card.TMDAY <= '" & cboEnd.Text & "'"
   Debug.Print sSql
   clsADOCon.ExecuteSQL sSql
   If Err Then
      MsgBox Err.Description
      Exit Sub
   End If
   SysMsg clsADOCon.RowsAffected & " time charges were updated", True
   Exit Sub
   
DiaErr1:
   sProcName = "UpdateTimeCharges"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Activate()
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
   
   'fill employee list
   sSql = "select PREMNUMBER from EmplTable where PREMSTATUS <> 'D' " _
          & "order by PREMNUMBER"
   LoadNumComboBox cboEmp, "000000"
   If bSqlRows Then cboEmp = cboEmp.List(0)
   
   cboStart = Format(Now, "mm/dd/yy")
   cboEnd = Format(Now, "mm/dd/yy")
   bOnLoad = 1
   Show
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaHdlch = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_FillEmployees"
   LoadNumComboBox cboEmp, "000000"
   If cboEmp.ListCount > 0 Then
      If Trim(sCurrEmployee) = "" Then
         cboEmp = cboEmp.List(0)
      Else
         cboEmp = sCurrEmployee
      End If
      GetEmployee
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetEmployee()
   Dim RdoEmp As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_EmployeeName " & Val(cboEmp)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoEmp)
   If bSqlRows Then
      With RdoEmp
         cboEmp = Format(!PREMNUMBER, "000000")
         lblName = "" & Trim(!PREMLSTNAME) & ", " _
                   & Trim(!PREMFSTNAME) & " " _
                   & Trim(!PREMMINIT)
         .Cancel
         sCurrEmployee = cboEmp
      End With
   Else
      MsgBox "Employee Wasn't Found.", vbExclamation, Caption
      lblName = "No Current Employee"
   End If
   Set RdoEmp = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getemploy"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cboStart_DropDown()
   ShowCalendar Me
End Sub


Private Sub cboStart_LostFocus()
   cboStart = CheckDate(cboStart)
End Sub

Private Sub txtPayRate_LostFocus()
   Dim cUR As String
   cUR = CheckCurrency(txtPayRate.Text, False)
   If cUR <> "*" Then
      txtPayRate = cUR
   Else
      txtPayRate.SetFocus
   End If
End Sub
