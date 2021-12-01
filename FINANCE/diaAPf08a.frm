VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaAPf08a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Check Number"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbChkAcc 
      Height          =   315
      Left            =   2280
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Check Account"
      Top             =   960
      Width           =   1555
   End
   Begin VB.TextBox txtFromChk 
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Tag             =   "3"
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Update"
      Height          =   315
      Left            =   5160
      TabIndex        =   3
      ToolTipText     =   "Change Account Number And Update References"
      Top             =   600
      Width           =   875
   End
   Begin VB.TextBox txtToChk 
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Tag             =   "2"
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5160
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   7
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
      PictureUp       =   "diaAPf08a.frx":0000
      PictureDn       =   "diaAPf08a.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3240
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2370
      FormDesignWidth =   6195
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Account "
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   8
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "To the New Check Number"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Change Check From"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "diaAPf08a"
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
Dim bGoodCheck As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

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


Private Sub cmdUpd_Click()
   UpdateCheckNumber
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then bOnLoad = False
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   bOnLoad = True
   bGoodCheck = False
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaAPf08a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub


Private Sub txtFromChk_LostFocus()
   If (Trim(txtFromChk) <> "") Then
      PopulateChkAcc Trim(txtFromChk.Text)
   End If
   
End Sub


Private Function PopulateChkAcc(strChk As String) As Boolean

   Dim RdoChk As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT CHKNUMBER,CHKACCT " _
          & "FROM chkstable WHERE CHKNUMBER= '" & strChk & "' ORDER BY CHKACCT"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk, ES_FORWARD)
   If bSqlRows Then
      With RdoChk
         Do Until .EOF
            AddComboStr cmbChkAcc.hWnd, "" & Trim(!CHKACCT)
            .MoveNext
         Loop
         .Cancel
      End With
      If (cmbChkAcc.ListCount > 0) Then cmbChkAcc = cmbChkAcc.List(0)
      
      bGoodCheck = True
   Else
      MsgBox "Check Number doesn't exist. Please provide the valid CheckNumber.", vbExclamation, Caption
      bGoodCheck = False
   End If
   Set RdoChk = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "ValidateChkAcc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function

Private Function ValidateChk(strChk As String, strAcct As String) As Byte
   Dim RdoChk As ADODB.Recordset
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT DISTINCT CHKNUMBER " _
          & "FROM chkstable WHERE CHKNUMBER= '" & strChk & "' AND CHKACCT = '" & strAcct & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk, ES_FORWARD)
   If bSqlRows Then
      ValidateChk = 1
   Else
      ValidateChk = 0
   End If
   Set RdoChk = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "PopulateChkAcc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Sub UpdateCheckNumber()
   Dim i As Integer
   Dim bResponse As Byte
   Dim strMsg As String
   Dim strFromChk As String
   Dim strToChk As String
   Dim strChkAcc As String
   Dim bToCheck As String
   
   On Error GoTo DiaErr1
   strFromChk = Trim(Compress(txtFromChk))
   strToChk = Trim(Compress(txtToChk))
   strChkAcc = Trim(cmbChkAcc.Text)
   
   If (strChkAcc = "") Then
      MsgBox "Check Number doesn't have a valid Check Account.", vbExclamation, Caption
      Exit Sub
   End If
   
   ' Make sure that the check number/Acct exists
   bGoodCheck = ValidateChk(strFromChk, strChkAcc)
   
   If (bGoodCheck = False) Then
      MsgBox "Check Number doesn't exist. Please provide the valid CheckNumber.", vbExclamation, Caption
      Exit Sub
   End If
   
   
   If (strToChk = "") Then
      MsgBox "Please provide new Check Number.", vbExclamation, Caption
      Exit Sub
   End If
   
   ' Make sure that the check number/Acct exists
   bToCheck = ValidateChk(strToChk, strChkAcc)
   If (bToCheck = True) Then
      MsgBox "The new Check Number Exists. Please provide new Check Number.", vbExclamation, Caption
      Exit Sub
   End If
   
   
   strMsg = "Do You Really Want To Change Check # " & strFromChk & vbCrLf _
          & "To Check # " & strToChk & "... "
          
   bResponse = MsgBox(strMsg, ES_NOQUESTION, Caption)
   
   If bResponse = vbYes Then
      MouseCursor 13
      Err = 0
      'On Error Resume Next
      Err.Clear
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      
      'Vendor Invoice table
      sSql = "UPDATE vihdTable SET VICHECKNO ='" & strToChk & "' " _
             & "WHERE VICHECKNO ='" & strFromChk & "' AND VICHKACCT = '" & strChkAcc & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      'Jurnal table
      sSql = "UPDATE jritTable SET DCCHECKNO ='" & strToChk & "' " _
             & "WHERE DCCHECKNO ='" & strFromChk & "' AND DCCHKACCT = '" & strChkAcc & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      'The actual check table
      sSql = "UPDATE chkstable SET CHKNUMBER ='" & strToChk & "' " _
             & "WHERE CHKNUMBER ='" & strFromChk & "' AND CHKACCT = '" & strChkAcc & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      MouseCursor 0
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         MsgBox "Check Number Was Successfully Changed.", _
            vbInformation, Caption
            
         txtFromChk = ""
         txtToChk = ""
         cmbChkAcc.Clear
      Else
         MsgBox "Update Check Number failed.", _
            vbInformation, Caption
         clsADOCon.RollbackTrans
      End If
   Else
      CancelTrans
   End If
   
   Exit Sub
   
DiaErr1:
   sProcName = "UpdateCheckNumber"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   Unload Me
   
End Sub

