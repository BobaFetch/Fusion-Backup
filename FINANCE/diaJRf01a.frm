VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaJRf01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open Journal"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSug 
      Caption         =   "&Suggest"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5280
      TabIndex        =   19
      ToolTipText     =   "Suggest Default System Descrition And Number"
      Top             =   1440
      Width           =   875
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1560
      TabIndex        =   16
      Top             =   840
      Width           =   615
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   360
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2895
      FormDesignWidth =   6255
   End
   Begin VB.ComboBox cmbTyp 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Select Journal Type From List"
      Top             =   360
      Width           =   2895
   End
   Begin VB.TextBox txtNme 
      Height          =   285
      Left            =   1560
      MaxLength       =   30
      TabIndex        =   1
      ToolTipText     =   "Journal Name/Description (30 Chars Max)"
      Top             =   1440
      Width           =   3375
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Tag             =   "4"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   3840
      TabIndex        =   5
      Tag             =   "4"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtFis 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      MaxLength       =   4
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "Fiscal Year As 2005"
      Top             =   1800
      Width           =   645
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Open"
      Height          =   315
      Left            =   5280
      TabIndex        =   6
      ToolTipText     =   "Open Journal"
      Top             =   600
      Width           =   875
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Left            =   4305
      MaxLength       =   4
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "Journal Number (1 To 255)"
      Top             =   1800
      Width           =   630
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   8
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
      PictureUp       =   "diaJRf01a.frx":0000
      PictureDn       =   "diaJRf01a.frx":0146
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Default System Descriptions Will Be Used."
      Height          =   495
      Index           =   7
      Left            =   2280
      TabIndex        =   18
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Open All Journals Types"
      Height          =   495
      Index           =   6
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblTyp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4560
      TabIndex        =   15
      Top             =   360
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Journal Type"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Open"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   11
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fiscal Year"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Journal Number"
      Height          =   255
      Index           =   5
      Left            =   3000
      TabIndex        =   9
      Top             =   1800
      Width           =   1335
   End
End
Attribute VB_Name = "diaJRf01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'*************************************************************************************
' diaJRf01a - Open Journals
'
' Created (cjs)
' Revision:
'   09/10/01 (nth) Interface rework
'   12/10/01 (nth) Add sysmsg to OpenJournal
'   01/28/03 (nth) Fixed error allow opening of journal without fiscal year
'                  definition.
'   01/28/03 (nth) Added CheckFiscalForJrn
'
'*************************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodYear As Byte
Dim bThisYear As Byte
Dim bNoFy As Byte
Dim iNextJournal As Integer
Dim sJournals(13, 2) As String
Dim sMsg As String

Public bRemote As Byte
Public bIndex As Byte ' journal type to display

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*************************************************************************************

Private Sub chkAll_Click()
   If chkAll = vbChecked Then
      txtNme.enabled = False
      txtNum.enabled = False
   Else
      txtNme.enabled = True
      txtNum.enabled = True
   End If
End Sub

Private Sub cmbTyp_Click()
   
   lblTyp = sJournals(cmbTyp.ListIndex, 1)
   iNextJournal = GetNextJournal()
   txtNum = Format(iNextJournal, "0000")
   
   ' Fill default
   FillDefDesc
      
   'txtFis.enabled = True
   txtNum.enabled = True
   txtBeg.enabled = True
   txtEnd.enabled = True
   'txtNme = ""
   
End Sub

Private Sub cmbTyp_LostFocus()
   Dim b As Byte
   Dim i As Integer
   For i = 0 To cmbTyp.ListCount - 1
      If cmbTyp = cmbTyp.List(i) Then b = True
   Next
   If Not b Then
      Beep
      cmbTyp = cmbTyp.List(0)
      lblTyp = sJournals(0, 1)
   End If
   
   'txtFis.enabled = True
   txtNum.enabled = True
   txtBeg.enabled = True
   txtEnd.enabled = True
   txtNum = Format(GetNextJournal(), "0000")
End Sub

Private Sub cmdAdd_Click()
   Dim dBeg As Date
   Dim dEnd As Date
   
   On Error Resume Next
   If Len(Trim(txtNme)) = 0 Then
      MsgBox "Requires A Journal Description.", _
         vbInformation, Caption
      txtNme.SetFocus
      Exit Sub
   End If
   If Val(txtNum) = 0 Then
      MsgBox "Requires A Valid Journal Number.", _
         vbInformation, Caption
      txtNum.SetFocus
      Exit Sub
   End If
   dBeg = Format(txtBeg, "mm/dd/yy")
   dEnd = Format(txtEnd, "mm/dd/yy")
   If dBeg > dEnd Then
      MsgBox "There Is A Date Mismatch.", _
         vbInformation, Caption
      txtBeg.SetFocus
      Exit Sub
   End If
   
   If (chkAll.Value = vbChecked) Then
      OpenAllJournal
   Else
      OpenJournal
   End If
End Sub

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
      MouseCursor 13
      FillCombo
      ' Fill default
      FillDefDesc
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub FillCombo()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim i As Integer
   
   On Error GoTo DiaErr1
   
   For i = 0 To 6
      AddComboStr cmbTyp.hWnd, sJournals(i, 0)
   Next
   cmbTyp = cmbTyp.List(bIndex)
   lblTyp = sJournals(bIndex, 1)
   
   'txtFis = Format(Now, "yyyy")
   SetFyForDate
   bGoodYear = CheckFiscalYear()
   If bGoodYear = 1 Then
      txtNum = Format(GetNextJournal(), "0000")
   Else
      sMsg = "Fiscal Years Have Not Been Initialized." & vbCr _
             & "Initialize Fiscal Years Now?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         bNoFy = 1
         diaGLe04a.Show
         Unload Me
      Else
         cmdAdd.enabled = False
      End If
   End If
   Exit Sub
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   
   sJournals(0, 0) = "Sales Journal"
   sJournals(0, 1) = "SJ"
   sJournals(1, 0) = "Purchases Journal"
   sJournals(1, 1) = "PJ"
   sJournals(2, 0) = "Cash Receipts Journal"
   sJournals(2, 1) = "CR"
   sJournals(3, 0) = "Computer Checks Journal"
   sJournals(3, 1) = "CC"
   sJournals(4, 0) = "External Checks Journal"
   sJournals(4, 1) = "XC"
'   sJournals(5, 0) = "Payroll Labor Journal"
'   sJournals(5, 1) = "PL"
'   sJournals(6, 0) = "Payroll Disbursements Journal"
'   sJournals(6, 1) = "PD"
   sJournals(5, 0) = "Time Journal"
   sJournals(5, 1) = "TJ"
   sJournals(6, 0) = "Inventory Journal"
   sJournals(6, 1) = "IJ"
   
   'txtFis = Format(Now, "yyyy")
   txtBeg = Format(Now, "mm/01/yy")
   txtEnd = GetMonthEnd(txtBeg)
   SetFyForDate
   
   sCurrForm = Caption
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If bRemote Then
      bNoFy = True
   End If
   FormUnload bNoFy
   Set diaJRf01a = Nothing
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Public Function CheckFiscalYear() As Byte
   Dim RdoFyr As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT FYYEAR FROM GlfyTable "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoFyr, ES_FORWARD)
   If bSqlRows Then
      CheckFiscalYear = 1
      RdoFyr.Cancel
   Else
      CheckFiscalYear = 0
   End If
   Set RdoFyr = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "checkfisc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub txtBeg_Change()
   Dim iNextJournal As Integer
   SetFyForDate
   iNextJournal = GetNextJournal()
   txtNum = Format(iNextJournal, "0000")
   FillDefDesc
End Sub

Private Sub txtBeg_Click()
   SetFyForDate
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendar Me
   SetFyForDate
End Sub

Private Sub txtBeg_LostFocus()
   txtBeg = CheckDate(txtBeg)
   SetFyForDate
End Sub

Private Sub txtend_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEnd_LostFocus()
   txtEnd = CheckDate(txtEnd)
End Sub

Private Sub txtNme_LostFocus()
   txtNme = CheckLen(txtNme, 30)
   txtNme = CheckComments(txtNme)
   txtNme = StrCase(txtNme)
End Sub

Private Sub txtNum_LostFocus()
   txtNum = CheckLen(txtNum, 4)
   'txtNum = Format(txtNum, "###0")
End Sub

Public Sub GetJrnlYear()
   On Error Resume Next
   If Len(Trim(txtFis)) = 2 Then
      If Val(txtFis) > 85 Then
         txtFis = "19" & Format(Val(txtFis), "00")
      Else
         txtFis = "20" & Format(Val(txtFis), "00")
      End If
   Else
      If Len(txtFis) = 3 Then
         Beep
         txtFis = Format(Now, "yyyy")
      End If
   End If
End Sub

Public Function GetNextJournal(Optional strType As String = "") As Integer
   Dim RdoNxt As ADODB.Recordset
   On Error GoTo DiaErr1
   
   If (strType = "") Then
      strType = lblTyp
   End If
   
   sSql = "SELECT MAX(MJNO) FROM JrhdTable WHERE " _
          & "(MJTYPE='" & strType & "' AND " _
          & "MJFY=" & Val(txtFis) & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoNxt, ES_FORWARD)
   
   If Not IsNull(RdoNxt.Fields(0)) Then
      GetNextJournal = RdoNxt.Fields(0) + 1
   Else
      GetNextJournal = 1
   End If
   
   Set RdoNxt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getnextjou"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Sub GetThisJournal(iIndex As Integer, iFY As Integer, iNo As Integer)
   Dim RdoGet As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT MJTYPE,MJFY,MJNO,MJSTART,MJEND," _
          & "MJDESCRIPTION FROM JrhdTable WHERE " _
          & "(MJFY=" & iFY & " AND MJNO=" & iNo & ") " _
          & "AND MJTYPE='" & lblTyp & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet, ES_FORWARD)
   If bSqlRows Then
      With RdoGet
         'txtFis.enabled = False
         txtNum.enabled = False
         txtBeg.enabled = False
         txtEnd.enabled = False
         
         txtFis = Format(!MJFY, "0000")
         txtNum = Format(!MJNO, "0000")
         txtNme = "" & Trim(!MJDESCRIPTION)
         txtBeg = Format(!MJSTART, "mm/dd/yy")
         txtEnd = Format(!MJEND, "mm/dd/yy")
         .Cancel
      End With
      On Error Resume Next
      txtNme.SetFocus
   Else
      
      'txtFis.enabled = True
      txtNum.enabled = True
      txtBeg.enabled = True
      txtEnd.enabled = True
      txtNme = ""
   End If
   Set RdoGet = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getthisjour"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub OpenAllJournal()
   Dim rdoJrn As ADODB.Recordset
   Dim rdoJrn1 As ADODB.Recordset
   Dim bResponse As Byte
   Dim sBDate As String
   Dim sJrnl As String
   Dim i As Integer
   Dim iJrNum As Integer
   Dim strMJType As String
   Dim strMonth As String
   
   Dim strDate As String
   
   Dim strBegDt As String
   Dim strEndDt As String
   
   
   On Error GoTo DiaErr1
   
   If CheckFiscalForJrn(CDate(txtBeg), CDate(txtEnd)) = 0 Then
      sMsg = "Fiscal Period Not Found For Journal" & vbCrLf _
             & "Date Range."
      MsgBox sMsg, vbInformation, Caption
      txtBeg.SetFocus
      Exit Sub
   End If
   
   strBegDt = Format(txtBeg, "mm/dd/yy")
   strEndDt = Format(txtEnd, "mm/dd/yy")
   
   sBDate = Format(txtBeg, "mmm") & " " & txtFis
   
   For i = 0 To cmbTyp.ListCount - 1
      
      strMJType = sJournals(i, 1)
      iJrNum = GetNextJournal(strMJType)
      
      sSql = "SELECT MJTYPE,MJFY,MJNO FROM JrhdTable WHERE " _
             & "MJTYPE='" & strMJType & "' AND MJFY=" & txtFis _
             & " AND MJSTART = '" & strBegDt & "' AND MJEND = '" & strEndDt & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn, ES_FORWARD)
      
      If Not bSqlRows Then
         
         sSql = "SELECT MJTYPE,MJSTART,MJEND FROM JrhdTable WHERE " _
                & "MJTYPE='" & strMJType & "' AND " _
                & "MJSTART IN('" & sBDate & "') AND MJCLOSED IS NULL"
         bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn1, ES_FORWARD)
         
         If Not bSqlRows Then
            
            If (IsNumeric(Left(txtBeg, 2))) Then
               strMonth = Left(txtBeg, 2)
            Else
               strMonth = Trim(txtNum)
            End If
            
            If (IsNumeric(Mid(txtBeg, 4, 2))) Then
               strDate = Mid(txtBeg, 4, 2)
            End If
            
            'sJrnl = strMJType & "-" & Trim(txtFis) & "-" & Trim(strMonth)
            sJrnl = strMJType & "-" & Trim(txtFis) & "-" & Trim(strMonth) & Trim(strDate)
                        
            sSql = "INSERT INTO JrhdTable (MJTYPE,MJFY,MJNO," _
                   & "MJSTART,MJEND,MJDESCRIPTION,MJGLJRNL) " _
                   & "VALUES('" & strMJType & "'," _
                   & txtFis & "," & iJrNum & ",'" _
                   & txtBeg & "','" & txtEnd & "','" _
                   & Trim(sJrnl) & "','" & Trim(sJrnl) & "')"
            On Error Resume Next
            clsADOCon.ADOErrNum = 0
            clsADOCon.ExecuteSQL sSql
         End If
         
      End If
      
      Set rdoJrn = Nothing
   Next
   
   If clsADOCon.ADOErrNum = 0 Then
      SysMsg "Successfully Opened All Journals.", True
   Else
      MsgBox "Could Not Open all Journals.", _
         vbExclamation, Caption
   End If
   
   
   Exit Sub
   
DiaErr1:
   sProcName = "openjournal"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub OpenJournal()
   Dim rdoJrn As ADODB.Recordset
   Dim bResponse As Byte
   Dim sBDate As String
   Dim sJrnl As String
   Dim strMonth As String
   Dim strDate As String
   Dim strBegDt As String
   Dim strEndDt As String
   Dim iJrNum As String
   
   Dim stDte1 As Date
   Dim stDte2 As Date
   Dim JrnlStatus As Boolean
   
   On Error GoTo DiaErr1
   
   If CheckFiscalForJrn(CDate(txtBeg), CDate(txtEnd)) = 0 Then
      sMsg = "Fiscal Period Not Found For Journal" & vbCrLf _
             & "Date Range."
      MsgBox sMsg, vbInformation, Caption
      txtBeg.SetFocus
      Exit Sub
   End If
   
   sBDate = Format(txtBeg, "mmm") & " " & txtFis
   strBegDt = Format(txtBeg, "mm/dd/yy")
   strEndDt = Format(txtEnd, "mm/dd/yy")
      
   sSql = "SELECT MJTYPE,MJFY,MJNO FROM JrhdTable WHERE " _
          & "MJTYPE='" & lblTyp & "' AND MJFY=" & txtFis _
         & " AND MJSTART = '" & strBegDt & "' AND MJEND = '" & strEndDt & "'" _
            & " AND MJNO= " & txtNum
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn, ES_FORWARD)
   If bSqlRows Then
      sMsg = "There Is A " & cmbTyp & " In The" & vbCr _
             & "Fiscal Year With That Number.."
      MsgBox sMsg, vbExclamation, Caption
      On Error Resume Next
      txtNum.SetFocus
      Set rdoJrn = Nothing
      Exit Sub
   End If
   Set rdoJrn = Nothing
   
   sSql = "SELECT MJTYPE,MJSTART,MJEND FROM JrhdTable WHERE " _
          & "MJTYPE='" & lblTyp & "' AND " _
          & "MJSTART IN('" & sBDate & "') AND MJCLOSED IS NULL"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn, ES_FORWARD)
   If bSqlRows Then
      'sMsg = "There Is An Open " & cmbTyp & vbCr _
      '       & "In The Selected Date Range.."
      'MsgBox sMsg, vbInformation, Caption
       Set rdoJrn = Nothing
       
       sSql = "SELECT Max(MJEND) FROM JrhdTable WHERE " _
          & "MJTYPE='" & lblTyp & "' AND " _
          & "MJSTART IN('" & sBDate & "') AND MJCLOSED IS NULL"
       bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn, ES_FORWARD)
       stDte1 = rdoJrn.Fields(0)
       stDte2 = Trim(txtBeg)
       If stDte1 > stDte2 Then
              sMsg = "There Is An Open " & cmbTyp & vbCr _
                     & "In The Selected Date Range.."
              MsgBox sMsg, vbInformation, Caption
              JrnlStatus = False
       Else
              JrnlStatus = True
       End If
   
   Else
      ' The Journal was not found
      JrnlStatus = True
   End If
   
   
   If JrnlStatus Then
      sMsg = "Open A New " & cmbTyp & vbCr _
             & "For The Selected Date Range.."
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         If (IsNumeric(Left(txtBeg, 2))) Then
            strMonth = Left(txtBeg, 2)
         Else
            strMonth = Trim(txtNum)
         End If
         
         If (IsNumeric(Mid(txtBeg, 4, 2))) Then
            strDate = Mid(txtBeg, 4, 2)
         End If
         'iJrNum = GetNextJournal()
         
         'sJrnl = lblTyp & "-" & Trim(txtFis) & "-" & Trim(strMonth)
         sJrnl = lblTyp & "-" & Trim(txtFis) & "-" & Trim(strMonth) & Trim(strDate)
         sSql = "INSERT INTO JrhdTable (MJTYPE,MJFY,MJNO," _
                & "MJSTART,MJEND,MJDESCRIPTION,MJGLJRNL) " _
                & "VALUES('" & lblTyp & "'," _
                & txtFis & "," & txtNum & ",'" _
                & txtBeg & "','" & txtEnd & "','" _
                & txtNme & "','" & Trim(sJrnl) & "')"
         On Error Resume Next
         clsADOCon.ADOErrNum = 0
         clsADOCon.ExecuteSQL sSql
         If clsADOCon.ADOErrNum = 0 Then
            SysMsg "Successfully Opened.", True
            txtNum = Format(GetNextJournal(), "0000")
         Else
            MsgBox "Could Not Open " & cmbTyp, _
               vbExclamation, Caption
         End If
      Else
         CancelTrans
      End If
   End If
   Set rdoJrn = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "openjournal"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Function CheckFiscalForJrn(dBeg As Date, dEnd As Date) As Byte
   
   Dim rdoFis As ADODB.Recordset
   Dim i As Integer
   On Error GoTo DiaErr1
'   sSql = "SELECT * FROM GlfyTable WHERE FYYEAR = " _
'          & CInt(Format(CStr(dBeg), "yyyy"))
'   bSqlRows = clsAdoCon.GetDataSet(sSql,rdoFis, ES_FORWARD)
'   If bSqlRows Then
'      With rdoFis
'         For i = 4 To (!FYPERIODS * 2) + 4 Step 2
'            If dBeg >= .Fields(i) And _
'                       dEnd <= .Fields(i + 1) Then
'               ' found it!
'               CheckFiscalForJrn = 1
'               Exit For
'            End If
'         Next
'      End With
'   End If

'   sSql = "SELECT * FROM GlfyTable WHERE FYYEAR = " _
'          & CInt(Format(CStr(dBeg), "yyyy"))
   sSql = "SELECT * FROM GlfyTable WHERE FYSTART <= '" & Format(dBeg, "mm/dd/yyyy") & "'" & vbCrLf _
          & "and FYEND >= '" & Format(dEnd, "mm/dd/yyyy") & "'"
   If clsADOCon.GetDataSet(sSql, rdoFis, ES_FORWARD) Then
      CheckFiscalForJrn = 1
   End If

   Set rdoFis = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "checkfiscalfor"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub SetFyForDate()
   If IsDate(txtBeg.Text) Then
      Dim rdo As ADODB.Recordset
      sSql = "select FYYEAR from GlfyTable WHERE FYSTART <= '" & Format(txtBeg, "mm/dd/yyyy") & "'" & vbCrLf _
             & "and FYEND >= '" & Format(txtBeg, "mm/dd/yyyy") & "'"
      If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
         txtFis = rdo.Fields(0)
         DoEvents
      Else
         MsgBox "No defined fiscal year includes the date " & txtBeg, vbInformation, Caption
      End If
   Else
      txtFis = ""
   End If
   Set rdo = Nothing
End Sub

Private Function FillDefDesc()
   Dim strMonth As String
   Dim strDate As String
   
   If (IsNumeric(Left(txtBeg, 2))) Then
      strMonth = Left(txtBeg, 2)
   End If
   If (IsNumeric(Mid(txtBeg, 4, 2))) Then
      strDate = Mid(txtBeg, 4, 2)
   End If
   
   'txtNme = lblTyp + "-" + txtFis + "-" + strMonth
   txtNme = lblTyp + "-" + txtFis + "-" + strMonth + "-" + strDate
End Function

