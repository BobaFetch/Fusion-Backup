VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaGLf01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Post General Journal"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtcrd 
      BackColor       =   &H8000000F&
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtDeb 
      BackColor       =   &H8000000F&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtPost 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   10
      Tag             =   "4"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtCreated 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   8
      Tag             =   "4"
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "&Post"
      Height          =   315
      Left            =   5280
      TabIndex        =   7
      ToolTipText     =   "Post This Journal To General Ledger"
      Top             =   600
      Width           =   855
   End
   Begin VB.ComboBox cmbjrn 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Tag             =   "2"
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox txtcmt 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   3
      Tag             =   "2"
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   4
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
      PictureUp       =   "diaGLf01a.frx":0000
      PictureDn       =   "diaGLf01a.frx":0146
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Amt."
      Height          =   255
      Index           =   6
      Left            =   2880
      TabIndex        =   13
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Debit Amt."
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   11
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Post"
      Height          =   255
      Index           =   4
      Left            =   2880
      TabIndex        =   9
      Top             =   960
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Created"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Journal ID"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "diaGLf01a"
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

Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Public Sub PostJrnl()
   Dim sMsg As String
   Dim bResponse As Byte
   Dim sJrn As String, PostThisJournal As Boolean
   
  
   sJrn = Trim(cmbjrn)
   sMsg = "Do You Wish To Post General Journal " & sJrn & " ?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   
   If bResponse = vbNo Then
      Exit Sub
   End If
   
   On Error Resume Next
   Err.Clear
   
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   sSql = "UPDATE GjhdTable SET " _
          & "GJPOSTED = 1" _
          & " WHERE GJNAME = '" & sJrn & "'"
   clsADOCon.ExecuteSql sSql
   
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      MouseCursor 0
      MsgBox Trim(cmbjrn) & " Successfully Posted.", vbInformation, Caption
      FillJournals
      
   Else
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
      MouseCursor 0
      MsgBox "Couldn't Successfuly Post " & Trim(cmbjrn) & ".", _
         vbExclamation, Caption
      GoTo DiaErr1
'   End If
   End If
   Exit Sub
DiaErr1:
   sProcName = "PostJrnl"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Public Sub FillJournals()
   Dim rdoJrn As ADODB.Recordset
   cmbjrn.Clear
   
   On Error GoTo DiaErr1
   sSql = "SELECT GJNAME FROM GjhdTable WHERE GJPOSTED = 0"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn, ES_FORWARD)
   
   If bSqlRows Then
      With rdoJrn
         Do Until .EOF
            AddComboStr cmbjrn.hWnd, "" & Trim(!GJNAME)
            .MoveNext
         Loop
         .Cancel
      End With
      cmbjrn.ListIndex = 0
      GetJrn
   Else
      MsgBox "No Open Journal Entries Found.", vbInformation, Caption
      Set rdoJrn = Nothing
      Unload Me
   End If
   Set rdoJrn = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "filljournals"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   Set rdoJrn = Nothing
   
End Sub

Public Function GetJrn() As Byte
   Dim RdoSum As ADODB.Recordset
   Dim RdoJid As ADODB.Recordset
   Dim sJrn As String
   On Error GoTo DiaErr1
   sJrn = Trim(cmbjrn)
   
   sSql = "SELECT GJDESC,GJOPEN,GJPOST FROM GjhdTable WHERE GJNAME = '" & sJrn & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoJid, ES_KEYSET)
   
   If bSqlRows Then
      With RdoJid
         txtcmt = "" & Trim(!GJDESC)
         txtCreated = "" & Format(!GJOPEN, "mm/dd/yy")
         txtPost = "" & Format(!GJPOST, "mm/dd/yy")
         .Cancel
      End With
   End If
   Set RdoJid = Nothing
   
   ' Now calc the sums
   sSql = "SELECT Sum(JIDEB) AS SumOfDCDEBIT, Sum(JICRD) AS SumOfDCCREDIT FROM GjitTable WHERE JINAME = '" & sJrn & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSum, ES_FORWARD)
   
   If bSqlRows Then
      With RdoSum
         txtDeb = Format(!SumOfDCDEBIT, "######0.00")
         txtcrd = Format(!SumOfDCCREDIT, "######0.00")
      End With
      
      If txtDeb = "" Then txtDeb = Format("0", "######0.00")
      If txtcrd = "" Then txtcrd = Format("0", "######0.00")
   End If
   
   If txtDeb = txtcrd Then
      If Val(txtDeb) <> 0 And Val(txtcrd) <> 0 Then
         cmdPost.enabled = True
         cmdPost.SetFocus
      Else
         cmdPost.enabled = False
      End If
   Else
      cmdPost.enabled = False
   End If
   Set RdoSum = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "GetJrn"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   Set RdoJid = Nothing
   
End Function

Private Sub cmbjrn_Click()
   If bOnLoad = False Then GetJrn
End Sub

Private Sub cmbjrn_LostFocus()
   If bOnLoad = False Then GetJrn
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "No Subject Help"
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdPost_Click()
   PostJrnl
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillJournals
      bOnLoad = False
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   bOnLoad = True
   txtDeb = Format("0", "######0.00")
   txtcrd = Format("0", "######0.00")
End Sub


Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   Set diaGLf01a = Nothing
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub
