VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaGLf03a
   BorderStyle = 3 'Fixed Dialog
   Caption = "Cancel Journal"
   ClientHeight = 2265
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 5550
   ClipControls = 0 'False
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 2265
   ScaleWidth = 5550
   ShowInTaskbar = 0 'False
   Begin VB.CommandButton cmdCancel
      Caption = "&Cancel"
      Height = 315
      Left = 4560
      TabIndex = 1
      Top = 600
      Width = 855
   End
   Begin VB.ComboBox cmbjrn
      Height = 315
      Left = 1440
      TabIndex = 0
      Tag = "2"
      Top = 240
      Width = 1935
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 4560
      TabIndex = 4
      TabStop = 0 'False
      Top = 90
      Width = 875
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 5
      ToolTipText = "Subject Help"
      Top = 0
      Width = 255
      _Version = 65536
      _ExtentX = 450
      _ExtentY = 397
      _StockProps = 65
      BackColor = 12632256
      Autosize = 2
      RoundedCorners = 0 'False
      BevelWidth = 0
      Outline = 0 'False
      PictureUp = "diaCnclGJ.frx":0000
      PictureDn = "diaCnclGJ.frx":0146
   End
   Begin VB.Label lblPost
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 255
      Left = 1440
      TabIndex = 10
      Top = 1440
      Width = 1095
   End
   Begin VB.Label lblOpen
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 255
      Left = 1440
      TabIndex = 9
      Top = 1080
      Width = 1095
   End
   Begin VB.Label lbldsc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 255
      Left = 1440
      TabIndex = 8
      Top = 600
      Width = 2775
   End
   Begin VB.Label z1
      Caption = "Post"
      Height = 255
      Index = 3
      Left = 360
      TabIndex = 7
      Top = 1440
      Width = 735
   End
   Begin VB.Label z1
      Caption = "Opened"
      Height = 255
      Index = 2
      Left = 360
      TabIndex = 6
      Top = 1080
      Width = 1095
   End
   Begin VB.Label z1
      Caption = "Description"
      Height = 255
      Index = 1
      Left = 360
      TabIndex = 3
      Top = 600
      Width = 855
   End
   Begin VB.Label z1
      Caption = "Journal ID"
      Height = 255
      Index = 0
      Left = 360
      TabIndex = 2
      Top = 240
      Width = 855
   End
End
Attribute VB_Name = "diaGLf03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001, ES/2002) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit


'**********************************************************************************
' diaGLf03a - Cancel A GL Journal (Posted and Unposted)
'
' Notes: Same form serves both purposes.
'
' Created: 09/30/01 (nth)
' Revisions:
'   11/15/02 (nth) Revised up to current specs.
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bGoodYear As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Public Sub GetJrn()
   Dim RdoJid As rdoResultset
   Dim sJrn As String
   On Error GoTo DiaErr1
   
   sJrn = Trim(cmbjrn)
   sSql = "SELECT GJDESC,GJOPEN,GJPOST FROM GjhdTable WHERE GJNAME = '" _
          & sJrn & "'"
   bSqlRows = GetDataSet(RdoJid)
   
   If bSqlRows Then
      With RdoJid
         lbldsc = "" & Trim(!GJDESC)
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



Public Sub UnPostJrn()
   Dim sJrn As String
   Dim smsg As String
   sJrn = cmbjrn
   
   On Error Resume Next
   sSql = "UPDATE GjhdTable SET GJPOSTED=0 WHERE GJNAME = '" & sJrn & "'"
   RdoCon.Execute sSql, rdExecDirect
   If Err = 0 Then
      RdoCon.CommitTrans
      FillJournals
      smsg = "Successfully Unposted Journal " & sJrn
      MsgBox smsg, vbInformation
   Else
      smsg = "Could Not Successfully Unpost Journal " & sJrn
      RdoCon.RollbackTrans
      MouseCursor 0
      MsgBox smsg, vbExclamation
   End If
End Sub

Public Sub CnclJrn()
   Dim sJrn As String
   Dim smsg As String
   Dim iResponse As Integer
   
   
   On Error GoTo DiaErr1
   MouseCursor 13
   sJrn = Trim(cmbjrn)
   
   smsg = "Cancel Journal Entry " & sJrn & " ?"
   MsgBox smsg, ES_YESQUESTION, Caption
   If iResponse = vbNo Then Exit Sub
   
   
   On Error Resume Next
   RdoCon.BeginTrans
   sSql = "DELETE FROM GjitTable WHERE JINAME = '" & sJrn & "'"
   RdoCon.Execute sSql, rdExecDirect
   If Err = 0 Then
      RdoCon.CommitTrans
      
      RdoCon.BeginTrans
      sSql = "DELETE FROM GjhdTable WHERE GJNAME = '" & sJrn & "'"
      RdoCon.Execute sSql, rdExecDirect
      
      If Err = 0 Then
         RdoCon.CommitTrans
         MouseCursor 0
         FillJournals
      Else
         smsg = "Could Not Successfully Delete Journal " & sJrn
         RdoCon.RollbackTrans
         MouseCursor 0
         MsgBox smsg, vbExclamation
      End If
      
   Else
      MouseCursor 0
      smsg = "Could Not Successfully Delete Journal " & sJrn
      RdoCon.RollbackTrans
      MsgBox smsg, vbExclamation
   End If
   Exit Sub
   
   DiaErr1:
   sProcName = "CnclJrn"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Public Function CheckFiscalYear() As Byte
   Dim RdoFyr As rdoResultset
   On Error GoTo DiaErr1
   sSql = "SELECT FYYEAR FROM GlfyTable "
   bSqlRows = GetDataSet(RdoFyr, ES_FORWARD)
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

Public Sub FillCombo()
   Dim smsg As String
   Dim bResponse As Byte
   
   bGoodYear = CheckFiscalYear()
   If bGoodYear Then
      FillJournals
   Else
      smsg = "Fiscal Years Have Not Been Initialized." & vbCr _
             & "Initialize Fiscal Years Now?"
      bResponse = MsgBox(smsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         diaGLf03a.Show
         Unload Me
      Else
         Unload Me
      End If
   End If
   
End Sub


Public Sub FillJournals()
   Dim RdoJrn As rdoResultset
   On Error GoTo DiaErr1
   On Error GoTo 0
   cmbjrn.Clear
   
   If Me.Caption = "Cancel A Posted Journal Entry" Then
      sSql = "SELECT GJNAME FROM GjhdTable WHERE GJPOSTED <> 0"
   Else
      sSql = "SELECT GJNAME FROM GjhdTable WHERE GJPOSTED = 0"
   End If
   
   bSqlRows = GetDataSet(RdoJrn, ES_FORWARD)
   If bSqlRows Then
      With RdoJrn
         Do Until .EOF
            AddComboStr cmbjrn.hWnd, "" & Trim(!GJNAME)
            .MoveNext
         Loop
         .Cancel
      End With
      cmbjrn.ListIndex = 0
   Else
      MsgBox "No Journals Found.", vbExclamation, Caption
      Unload Me
   End If
   Set RdoJrn = Nothing
   Exit Sub
   DiaErr1:
   Set RdoJrn = Nothing
   sProcName = "filljournals"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cmbjrn_Click()
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
   SetDiaPos Me
   'cmdCancel.Enabled = False
   
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

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub
