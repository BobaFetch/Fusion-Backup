VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaTCCpy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy Daily Time Charge"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Destination Time Card"
      Height          =   1935
      Left            =   240
      TabIndex        =   11
      Top             =   3360
      Width           =   6015
      Begin VB.ComboBox cmbDesEmp 
         Height          =   315
         Left            =   1560
         TabIndex        =   14
         Tag             =   "1"
         ToolTipText     =   "Select From List Or Enter Number"
         Top             =   720
         Width           =   1095
      End
      Begin VB.ComboBox txtDesDte 
         Height          =   315
         Left            =   4560
         TabIndex        =   12
         Tag             =   "4"
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4560
         TabIndex        =   18
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblDesNme 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1560
         TabIndex        =   16
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Number"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Card Date"
         Height          =   255
         Index           =   3
         Left            =   3480
         TabIndex        =   13
         Top             =   735
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Source Time Card"
      Height          =   1935
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   6015
      Begin VB.ComboBox txtScrDte 
         Height          =   315
         Left            =   4560
         TabIndex        =   5
         Tag             =   "4"
         Top             =   720
         Width           =   1095
      End
      Begin VB.ComboBox cmbSrcEmp 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Tag             =   "1"
         ToolTipText     =   "Select From List Or Enter Number"
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Card Date"
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   10
         Top             =   735
         Width           =   855
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblSsn 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4560
         TabIndex        =   8
         Top             =   1200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblNme 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Number"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy TC"
      Height          =   315
      Left            =   5280
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   1
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
      PictureUp       =   "diaTCcpy.frx":0000
      PictureDn       =   "diaTCcpy.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4320
      Top             =   480
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5505
      FormDesignWidth =   6420
   End
End
Attribute VB_Name = "diaTCCpy"
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
Dim sCurrSrcEmp As String
Dim sCurrDesEmp As String


Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbSrcEmp_Click()
   GetSrcEmployee
End Sub


Private Sub cmbSrcEmp_KeyUp(KeyCode As Integer, Shift As Integer)
   
   cmbSrcEmp = CheckLen(cmbSrcEmp, 6)
   If Len(cmbSrcEmp) Then
      cmbSrcEmp = Format(cmbSrcEmp, "000000")
      GetSrcEmployee
   End If
   
End Sub

Private Sub cmbDesEmp_Click()
   GetDesEmployee
   
End Sub


Private Sub cmbDesEmp_KeyUp(KeyCode As Integer, Shift As Integer)
   
   cmbDesEmp = CheckLen(cmbDesEmp, 6)
   If Len(cmbDesEmp) Then
      cmbDesEmp = Format(cmbDesEmp, "000000")
      GetDesEmployee
   End If

End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCopy_Click()
   Dim bResponse As Byte
   Dim strMsg As String
   Dim strCardNum As String
   Dim bGoodCard As Boolean
   Dim strSrcDate As String
   Dim strDesDate As String
   Dim strNewCardNum As String
   Dim bFound As Boolean
   
   On Error GoTo DiaErr1
   
   strSrcDate = CStr(Format(txtScrDte, "mm/dd/yyyy"))
   strDesDate = CStr(Format(txtDesDte, "mm/dd/yyyy"))
   
   If (CDate(strSrcDate) >= CDate(strDesDate)) Then
      strMsg = "The Source date should be less than Destination date " & vbCrLf
      bResponse = MsgBox(strMsg, vbCritical)
      Exit Sub
   End If
   
   Dim tc As New ClassTimeCharge
   Dim strJournalID As String
   bFound = tc.GetOpenTimeJournalForThisDate(strDesDate, strJournalID)
   If Not bFound Then
      Exit Sub
   End If
   
   bGoodCard = GetCard(strCardNum)
   If bGoodCard Then
      
      strMsg = "Do You Want To Copy The Time Card From " & strSrcDate & vbCrLf _
             & "date  to " & strDesDate & " For Employee " & lblNme & "?"
      bResponse = MsgBox(strMsg, ES_NOQUESTION, Caption)
      If bResponse = vbYes Then
         
            Dim iDays As Integer
            Dim bret As Boolean
            Dim iEmpNo As Long
            Dim iDesEmpNo As Long
            iEmpNo = Val(cmbSrcEmp)
            iDesEmpNo = Val(cmbDesEmp)
            
            strNewCardNum = GetNewNumber()
            bret = CheckDateExists(iDesEmpNo, strDesDate)
            
            If (Not bret) Then
               Dim strWndDate As String
               Dim strTCStartDt As Date
               Dim strTCEndDt As Date
               
               iDays = DateDiff("d", strSrcDate, strDesDate)
                     
               strWndDate = GetWeekEnd(strDesDate)
               strTCStartDt = DateAdd("d", 0, strSrcDate)
               strTCEndDt = DateAdd("d", iDays, strSrcDate)
               
               clsADOCon.BeginTrans
               sSql = "INSERT INTO TchdTable (TMCARD, TMEMP, TMDATE, TMDAY, TMWEEK, " & _
                        "TMSTART, TMSTOP, TMREGHRS, TMOVTHRS, TMDBLHRS)" & _
                     " SELECT '" & CStr(strNewCardNum) & "','" & CStr(cmbDesEmp) & "','" & strDesDate & "', " & _
                     " '" & strDesDate & "','" & strWndDate & "'," & " TMSTART, TMSTOP, " & _
                        "TMREGHRS, TMOVTHRS, TMDBLHRS FROM TchdTable " & _
                     " WHERE TMEMP=" & Val(cmbSrcEmp) & " AND TMDAY='" & strSrcDate & "'" & _
                           " AND TMCARD = '" & strCardNum & "'"
            
               clsADOCon.ExecuteSQL sSql ', rdExecDirect
               
               ' Insert inti tcitTable
               sSql = "INSERT INTO TcitTable (TCCARD, TCEMP, TCSTART, TCSTOP, TCHOURS, " & _
                     "TCTIME, TCCODE, TCRATE, TCOHRATE, TCRATENO, TCACCT, TCSHOP, TCWC, TCPAYTYPE," & _
                      "TCSURUN, TCYIELD, TCPARTREF, TCRUNNO, TCOPNO, TCHDRATE, TCACCEPT, " & _
                      "TCREJECT, TCSCRAP, TCARTAG, TCBILLTC, TCMULTIJOB, TCSORT," & _
                      "TCOHFIXED , TCACCOUNT, TCGLJOURNAL, TCGLREF, TCPRORATE, TCSOURCE, " & _
                      "TCSTARTTIME, TCSTOPTIME, TCENTERED, TCCOMMENTS) " & _
                  "SELECT '" & CStr(strNewCardNum) & "','" & Val(cmbDesEmp) & "', TCSTART, TCSTOP, TCHOURS, " & _
                     "TCTIME, TCCODE, TCRATE, TCOHRATE, TCRATENO, TCACCT, TCSHOP, TCWC, TCPAYTYPE," & _
                      "TCSURUN, TCYIELD, TCPARTREF, TCRUNNO, TCOPNO, TCHDRATE, TCACCEPT, " & _
                      "TCREJECT, TCSCRAP, TCARTAG, TCBILLTC, TCMULTIJOB, TCSORT," & _
                      "TCOHFIXED , TCACCOUNT, TCGLJOURNAL, TCGLREF, TCPRORATE, TCSOURCE, " & _
                      "DATEADD(d, " & CStr(iDays) & ", TCSTARTTIME), " & _
                      "DATEADD(d, " & CStr(iDays) & ", TCSTOPTIME), TCENTERED, TCCOMMENTS " & _
                     " FROM TcitTable WHERE TCEMP=" & Val(cmbSrcEmp) & " AND TCCARD = '" & strCardNum & "'"
               
                  clsADOCon.ExecuteSQL sSql ' rdExecDirect
               
                  clsADOCon.CommitTrans
            End If
         
         If clsADOCon.RowsAffected > 0 Then
            MsgBox "New Time Card Created..", vbInformation, Caption
         Else
            MsgBox "Couldn't Create New Time Card.", vbExclamation, Caption
         End If
      Else
         CancelTrans
      End If
   End If

   Exit Sub
DiaErr1:
   clsADOCon.RollbackTrans
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "hs1550"
      cmdHlp = False
      MouseCursor 0
   End If
   
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
   
   txtScrDte = Format(Now - 1, "mm/dd/yy")
   If sCurrDate = "" Then
      If Format(txtScrDte, "w") = 1 Then
         txtScrDte = Format(Now - 2, "mm/dd/yy")
      End If
   Else
      txtScrDte = sCurrDate
   End If
   
   txtDesDte = Format(Now - 1, "mm/dd/yy")
   If sCurrDate = "" Then
      If Format(txtScrDte, "w") = 1 Then
         txtDesDte = Format(Now - 2, "mm/dd/yy")
      End If
   Else
      txtDesDte = sCurrDate
   End If
   
   
   bOnLoad = 1
   Show
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   sCurrDate = txtScrDte
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaTCCpy = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   Dim strEmpNo As String
   Dim strLabel As String
   
   On Error GoTo DiaErr1
   sSql = "Qry_FillEmployees"
   LoadNumComboBox cmbSrcEmp, "000000"
   If cmbSrcEmp.ListCount > 0 Then
      If Trim(sCurrSrcEmp) = "" Then
         cmbSrcEmp = cmbSrcEmp.List(0)
      Else
         cmbSrcEmp = sCurrSrcEmp
      End If
      GetSrcEmployee
   End If
   
   sSql = "Qry_FillEmployees"
   LoadNumComboBox cmbDesEmp, "000000"
   If cmbDesEmp.ListCount > 0 Then
      If Trim(sCurrDesEmp) = "" Then
         cmbDesEmp = cmbDesEmp.List(0)
      Else
         cmbDesEmp = sCurrDesEmp
      End If
      GetDesEmployee
   End If
   
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetSrcEmployee()
   Dim RdoEmp As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_EmployeeName " & Val(cmbSrcEmp)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoEmp)
   If bSqlRows Then
      With RdoEmp
         cmbSrcEmp = Format(!PREMNUMBER, "000000")
         lblNme = "" & Trim(!PREMLSTNAME) & ", " _
                  & Trim(!PREMFSTNAME) & " " _
                  & Trim(!PREMMINIT)
         lblSsn = "" & Trim(!PREMSOCSEC)
         .Cancel
         sCurrSrcEmp = cmbSrcEmp
      End With
   Else
      MsgBox "Employee Wasn't Found.", vbExclamation, Caption
      lblNme = "No Current Employee"
      'lblSsn = ""
   End If
   Set RdoEmp = Nothing
   Exit Sub

DiaErr1:
   sProcName = "getemploy"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub

Private Sub GetDesEmployee()
   Dim RdoEmp As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_EmployeeName " & Val(cmbDesEmp)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoEmp)
   If bSqlRows Then
      With RdoEmp
         cmbDesEmp = Format(!PREMNUMBER, "000000")
         lblDesNme = "" & Trim(!PREMLSTNAME) & ", " _
                  & Trim(!PREMFSTNAME) & " " _
                  & Trim(!PREMMINIT)
         'lblSsn = "" & Trim(!PREMSOCSEC)
         .Cancel
         sCurrDesEmp = cmbDesEmp
      End With
   Else
      MsgBox "Employee Wasn't Found.", vbExclamation, Caption
      lblDesNme = "No Current Employee"
      'lblSsn = ""
   End If
   Set RdoEmp = Nothing
   Exit Sub

DiaErr1:
   sProcName = "getemploy"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub


'Private Sub GetEmployee(ByRef strEmpNo As String, ByRef strLabel As String)
'   Dim RdoEmp As ADODB.Recordset
'   On Error GoTo DiaErr1
'   sSql = "Qry_EmployeeName " & Val(strEmpNo)
'   bSqlRows = clsADOCon.GetDataSet(sSql,RdoEmp)
'   If bSqlRows Then
'      With RdoEmp
'         strEmpNo = Format(!PREMNUMBER, "000000")
'         strLabel = "" & Trim(!PREMLSTNAME) & ", " _
'                  & Trim(!PREMFSTNAME) & " " _
'                  & Trim(!PREMMINIT)
'         'lblSsn = "" & Trim(!PREMSOCSEC)
'         .Cancel
'         sCurrEmployee = cmbSrcEmp
'      End With
'   Else
'      MsgBox "Employee Wasn't Found.", vbExclamation, Caption
'      strLabel = "No Current Employee"
'      'lblSsn = ""
'   End If
'   Set RdoEmp = Nothing
'   Exit Sub
'
'DiaErr1:
'   sProcName = "getemploy"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub

Private Function GetCard(ByRef strCardNum As String) As Byte
   Dim RdoGet As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT TMCARD,TMEMP,TMDAY FROM TchdTable WHERE " _
          & "TMEMP=" & Val(cmbSrcEmp) & " AND TMDAY='" & txtScrDte & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet)
   If bSqlRows Then
      GetCard = True
      strCardNum = Trim(RdoGet!TMCARD)
      On Error Resume Next
      cmbSrcEmp.SetFocus
   Else
      strCardNum = ""
      sSql = "There Is No Time Card Recorded " & vbCrLf _
             & "For " & Trim(lblNme) & " On " & txtScrDte & "."
      'Beep
      MsgBox sSql, vbInformation, Caption
      GetCard = False
   End If
   Set RdoGet = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getcard"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub txtScrDte_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtScrDte_LostFocus()
   txtScrDte = CheckDate(txtScrDte)
   
End Sub

Private Sub txtDesDte_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtDesDte_LostFocus()
   txtDesDte = CheckDate(txtDesDte)
   
End Sub

Private Function CheckDateExists(empno As Long, strDesDate As String)

   Dim RdoCrd As ADODB.Recordset
   Dim strCardNum As String
   
   sSql = "SELECT TMCARD FROM TchdTable WHERE " _
          & "TMEMP=" & Val(empno) & " AND TMDAY='" & strDesDate & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCrd, ES_FORWARD)
   If bSqlRows Then
      If RdoCrd.EOF Then
         CheckDateExists = True
         strCardNum = "" & Trim(RdoCrd!TMCARD)
         ClearResultSet RdoCrd
      Else
         CheckDateExists = False
      End If
   Else
      CheckDateExists = False
   End If
   
   Set RdoCrd = Nothing
   
End Function

Private Function GetNewNumber() As String
   Dim S As Single
   Dim l As Long
   Dim m As Long
   Dim t As String
   On Error Resume Next
   '    m = DateValue(Format(ES_SYSDATE, "yyyy,mm,dd"))
   '    s = TimeValue(Format(ES_SYSDATE, "hh:mm:ss"))
   '    l = s * 1000000
   '    GetNewNumber = Format(m, "00000") & Format(l, "000000")
   Dim dt As Variant
   dt = GetServerDateTime()
   m = DateValue(Format(dt, "yyyy,mm,dd"))
   S = TimeValue(Format(dt, "hh:mm:ss"))
   l = S * 1000000
   GetNewNumber = Format(m, "00000") & Format(l, "000000")
   
End Function

Private Function GetWeekEnd(strDate As String) As String
   Dim RdoGet As ADODB.Recordset
   Dim A As Integer
   Dim iList As Integer
   Dim dDate As Date
   Dim sWeekEnds As String
   Dim strWndDate As String
   
   On Error GoTo DiaErr1
   sSql = "SELECT WEEKENDS FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet, ES_FORWARD)
   If bSqlRows Then
      With RdoGet
         sWeekEnds = "" & Trim(!WEEKENDS)
         .Cancel
      End With
      If sWeekEnds = "Sat" Then iList = 7 Else iList = 8
   End If
   A = Format(strDate, "w")
   strWndDate = Format(CDate(strDate) + (iList - A), "mm/dd/yy")
   GetWeekEnd = strWndDate
   Set RdoGet = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getweeken"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


