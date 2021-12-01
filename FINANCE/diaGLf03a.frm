VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaGLf03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancel Journal"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtTyp 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Tag             =   "0"
      ToolTipText     =   "Journal Prefix (12 Chars Max)"
      Top             =   720
      Width           =   1605
   End
   Begin VB.ComboBox cmbFyr 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Fiscal Year Filter"
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   5520
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin VB.ComboBox cmbjrn 
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Tag             =   "2"
      ToolTipText     =   "Journal Name (12 Chars Max)"
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5520
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   90
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   6
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
      PictureUp       =   "diaGLf03a.frx":0000
      PictureDn       =   "diaGLf03a.frx":0146
   End
   Begin VB.Label P 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   720
      Width           =   945
   End
   Begin VB.Label P 
      BackStyle       =   0  'Transparent
      Caption         =   "(Optional)"
      Height          =   285
      Index           =   2
      Left            =   3240
      TabIndex        =   13
      Top             =   720
      Width           =   945
   End
   Begin VB.Label P 
      BackStyle       =   0  'Transparent
      Caption         =   "Fiscal Year"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   360
      Width           =   1185
   End
   Begin VB.Label lblPost 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3360
      TabIndex        =   11
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblOpen 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Post"
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   8
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Opened"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Journal"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
End
Attribute VB_Name = "diaGLf03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'**********************************************************************************
' diaGLf03a - Cancel A GL Journal (Posted and Unposted)
'
' Notes: Same form serves both purposes.
'
' Created: 09/30/01 (nth)
' Revisions:
' 11/15/02 (nth) Revised up to current specs.
' 11/24/03 (nth) fixed incident 18561
' 09/29/04 (nth) Added fiscal year filter.
'
'*********************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodYear As Byte
Dim sMsg As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Public Sub GetJrn()
   Dim RdoJid As ADODB.Recordset
   Dim sJrn As String
   On Error GoTo DiaErr1
   
   sJrn = Trim(cmbJrn)
   sSql = "SELECT GJDESC,GJOPEN,GJPOST FROM GjhdTable WHERE GJNAME = '" _
          & sJrn & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoJid)
   
   If bSqlRows Then
      With RdoJid
         lblDsc = "" & Trim(!GJDESC)
         lblOpen = "" & Format(!GJOPEN, "mm/dd/yy")
         lblPost = "" & Format(!GJPOST, "mm/dd/yy")
         .Cancel
      End With
   End If
   Set RdoJid = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "GetJrn"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub



Private Sub UnPostJrn()
   'Dim rdoJrn As ADODB.Recordset
   
   clsADOCon.ADOErrNum = 0
   sSql = "UPDATE GjhdTable SET GJPOSTED=0 WHERE GJNAME = '" _
          & cmbJrn & "'"
   clsADOCon.ExecuteSQL sSql
   
   If clsADOCon.ADOErrNum = 0 Then
      
      'Check is journal is a subsiderary journal
      'if so also cancel.
      'sSql = "SELECT MJGLJRNL FROM JrhdTable WHERE MJGLJRNL = '" & cmbJrn & "'"
      'bSqlRows = clsAdoCon.GetDataSet(sSql,rdoJrn)
      'If bSqlRows Then
      '    CnclJrn
      'End If
      sMsg = "Successfully Unposted Journal " & cmbJrn
      FillJournals
      
      MsgBox sMsg, vbInformation
   Else
      sMsg = "Could Not Successfully Unpost Journal " & cmbJrn
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
      MouseCursor 0
      MsgBox sMsg, vbExclamation
   End If
End Sub

Private Sub CnclJrn()
   Dim sJrn As String
   Dim sMsg As String
   Dim iResponse As Integer
   
   
   On Error GoTo DiaErr1
   MouseCursor 13
   sJrn = Trim(cmbJrn)
   
   sMsg = "Cancel Journal Entry " & sJrn & " ?"
   MsgBox sMsg, ES_YESQUESTION, Caption
   If iResponse = vbNo Then Exit Sub
   
   
   On Error Resume Next
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0

   sSql = "DELETE FROM GjitTable WHERE JINAME = '" & sJrn & "'"
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      
      clsADOCon.BeginTrans
      sSql = "DELETE FROM GjhdTable WHERE GJNAME = '" & sJrn & "'"
      clsADOCon.ExecuteSQL sSql
      
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         MouseCursor 0
         FillJournals
      Else
         sMsg = "Could Not Successfully Delete Journal " & sJrn
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         MouseCursor 0
         MsgBox sMsg, vbExclamation
      End If
      
   Else
      MouseCursor 0
      sMsg = "Could Not Successfully Delete Journal " & sJrn
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
      MsgBox sMsg, vbExclamation
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "CnclJrn"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Function CheckFiscalYear() As Byte
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

Private Sub FillCombo()
   Dim sMsg As String
   Dim bResponse As Byte
   
   bGoodYear = CheckFiscalYear()
   On Error GoTo DiaErr1
   If bGoodYear Then
      sProcName = "fillfiscalyea"
      FillFiscalYears Me
      sProcName = "fillJournals"
      FillJournals
   Else
      sMsg = "Fiscal Years Have Not Been Initialized." & vbCr _
             & "Initialize Fiscal Years Now?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         diaGLf03a.Show
         Unload Me
      Else
         Unload Me
      End If
   End If
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillJournals()
   Dim rdoJrn As ADODB.Recordset
   On Error GoTo DiaErr1
   cmbJrn.Clear
   
   sSql = "SELECT GJNAME" & vbCrLf _
      & "FROM GjhdTable" & vbCrLf _
      & "JOIN GlfyTable ON GJPOST BETWEEN FYSTART AND FYEND" & vbCrLf _
      & "AND FYYEAR = " & cmbFyr & vbCrLf
   If Me.Caption = "Cancel A Posted Journal Entry" Then
      sSql = sSql & "WHERE GJPOSTED <> 0" & vbCrLf
   Else
      sSql = sSql & "WHERE GJPOSTED = 0" & vbCrLf
   End If
   If Trim(txtTyp) <> "" Then
      sSql = sSql & "AND GJNAME LIKE '" & txtTyp & "%'"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn, ES_FORWARD)
   If bSqlRows Then
      With rdoJrn
         Do Until .EOF
            AddComboStr cmbJrn.hWnd, "" & Trim(!GJNAME)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   lblDsc = ""
   If cmbJrn.ListCount > 0 Then
      cmbJrn.ListIndex = 0
   End If
   Set rdoJrn = Nothing
   Exit Sub
DiaErr1:
   Set rdoJrn = Nothing
   sProcName = "filljournals"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cmbFyr_Click()
   FillJournals
End Sub

Private Sub cmbFyr_GotFocus()
   SelectFormat Me
End Sub

Private Sub cmbFyr_LostFocus()
   FillJournals
End Sub

Private Sub cmbjrn_Click()
   GetJrn
End Sub

Private Sub cmbJrn_GotFocus()
   SelectFormat Me
End Sub

Private Sub cmbjrn_LostFocus()
   cmbJrn = CheckLen(cmbJrn, 12)
   GetJrn
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCancel_Click()
   If Me.Caption = "Cancel A Posted Journal Entry" Then
      UnPostJrn
   Else
      CnclJrn
   End If
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "No Subject Help"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      If Me.Caption = "Cancel A Posted Journal Entry" Then
         cmdCancel.Caption = "&Unpost"
      End If
      FillCombo
      GetJrn
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   Set diaGLf03a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub txtTyp_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtTyp_LostFocus()
   txtTyp = CheckLen(txtTyp, 12)
   FillJournals
End Sub
