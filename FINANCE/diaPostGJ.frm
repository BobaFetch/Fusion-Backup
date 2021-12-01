VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaPostGJ
   BorderStyle = 3 'Fixed Dialog
   Caption = "Post General Journal"
   ClientHeight = 2400
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 6480
   ClipControls = 0 'False
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 2400
   ScaleWidth = 6480
   ShowInTaskbar = 0 'False
   Begin VB.TextBox txtcrd
      BackColor = &H8000000F&
      ForeColor = &H0000C000&
      Height = 285
      Left = 3720
      Locked = -1 'True
      TabIndex = 14
      Top = 1920
      Width = 1095
   End
   Begin VB.TextBox txtDeb
      BackColor = &H8000000F&
      ForeColor = &H000000FF&
      Height = 285
      Left = 1560
      Locked = -1 'True
      TabIndex = 12
      Top = 1920
      Width = 1095
   End
   Begin VB.TextBox txtPost
      BackColor = &H8000000F&
      Height = 285
      Left = 3720
      Locked = -1 'True
      TabIndex = 10
      Tag = "4"
      Top = 1080
      Width = 1095
   End
   Begin VB.TextBox txtCreated
      BackColor = &H8000000F&
      Height = 285
      Left = 1560
      Locked = -1 'True
      TabIndex = 8
      Tag = "4"
      Top = 1080
      Width = 1095
   End
   Begin VB.CommandButton cmdPost
      Caption = "&Post"
      Height = 375
      Left = 5520
      TabIndex = 7
      ToolTipText = "Post This Journal To General Ledger"
      Top = 600
      Width = 855
   End
   Begin VB.ComboBox cmbjrn
      Height = 315
      Left = 1560
      TabIndex = 5
      Tag = "2"
      Top = 240
      Width = 1815
   End
   Begin VB.TextBox txtcmt
      BackColor = &H8000000F&
      Height = 285
      Left = 1560
      Locked = -1 'True
      TabIndex = 3
      Tag = "2"
      Top = 720
      Width = 3315
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 5520
      TabIndex = 2
      TabStop = 0 'False
      Top = 90
      Width = 875
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 4
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
      PictureUp = "diaPostGJ.frx":0000
      PictureDn = "diaPostGJ.frx":0146
   End
   Begin Threed.SSFrame z2
      Height = 30
      Index = 0
      Left = -360
      TabIndex = 15
      Top = 1680
      Width = 6765
      _Version = 65536
      _ExtentX = 11933
      _ExtentY = 53
      _StockProps = 14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
   End
   Begin VB.Label z1
      Caption = "Credit"
      Height = 255
      Index = 6
      Left = 2880
      TabIndex = 13
      Top = 1920
      Width = 855
   End
   Begin VB.Label z1
      Caption = "Debit"
      Height = 255
      Index = 5
      Left = 360
      TabIndex = 11
      Top = 1920
      Width = 855
   End
   Begin VB.Label z1
      Caption = "Post"
      Height = 255
      Index = 4
      Left = 2880
      TabIndex = 9
      Top = 1080
      Width = 855
   End
   Begin VB.Label z1
      Caption = "Created"
      Height = 255
      Index = 2
      Left = 360
      TabIndex = 6
      Top = 1080
      Width = 855
   End
   Begin VB.Label z1
      Caption = "Description"
      Height = 255
      Index = 1
      Left = 360
      TabIndex = 1
      Top = 720
      Width = 855
   End
   Begin VB.Label z1
      Caption = "Journal ID"
      Height = 255
      Index = 0
      Left = 360
      TabIndex = 0
      Top = 240
      Width = 855
   End
End
Attribute VB_Name = "diaPostGJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bOnload As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Public Sub PostJrnl()
   Dim sMsg As String
   Dim bResponse As Byte
   Dim sJrn As String
   
   On Error GoTo DiaErr1
   sJrn = Trim(cmbjrn)
   sMsg = "Do You Wish To Post General Journal " & Trim(cmbjrn) & " ?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   
   If bResponse = vbNo Then
      Exit Sub
   End If
   
   On Error Resume Next
   Err = 0
   RdoCon.BeginTrans
   
   sSql = "UPDATE GjhdTable SET " _
          & "GJPOSTED = 1" _
          & " WHERE GJNAME = '" & sJrn & "'"
   RdoCon.Execute sSql, rdExecDirect
   
   If Err = 0 Then
      RdoCon.CommitTrans
      FillJournals
      
      MouseCursor 0
      MsgBox Trim(cmbjrn) & " Successfully Posted.", _
                  vbInformation, Caption
      
   Else
      RdoCon.RollbackTrans
      MouseCursor 0
      MsgBox "Couldn't Successfuly Post " & Trim(cmbjrn) & ".", _
         vbExclamation, Caption
      
   End If
   Exit Sub
   DiaErr1:
   sProcName = "PostJrnl"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Public Sub FillJournals()
   Dim RdoJrn As rdoResultset
   cmbjrn.Clear
   
   On Error GoTo DiaErr1
   sSql = "SELECT GJNAME FROM GjhdTable WHERE GJPOSTED = 0"
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
      MsgBox "No Open General Journals Found.", vbExclamation, Caption
      Unload Me
   End If
   
   Set RdoJrn = Nothing
   Exit Sub
   DiaErr1:
   sProcName = "filljournals"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   Set RdoJrn = Nothing
   
End Sub

Public Function GetJrn() As Byte
   Dim RdoSum As rdoResultset
   Dim RdoJid As rdoResultset
   Dim sJrn As String
   On Error GoTo DiaErr1
   sJrn = Trim(cmbjrn)
   
   sSql = "SELECT GJDESC,GJOPEN,GJPOST FROM GjhdTable WHERE GJNAME = '" & sJrn & "'"
   bSqlRows = GetDataSet(RdoJid, ES_KEYSET)
   
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
   bSqlRows = GetDataSet(RdoSum, ES_FORWARD)
   
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
         cmdPost.Enabled = True
         cmdPost.SetFocus
      Else
         cmdPost.Enabled = False
      End If
   Else
      cmdPost.Enabled = False
   End If
   
   Exit Function
   
   DiaErr1:
   sProcName = "GetJrn"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   Set RdoJid = Nothing
   
End Function



Private Sub cmbjrn_Click()
   If bOnload = False Then GetJrn
End Sub

Private Sub cmbjrn_LostFocus()
   If bOnload = False Then GetJrn
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
   If bOnload Then
      FillJournals
      GetJrn
      bOnload = False
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   bOnload = True
   txtDeb = Format("0", "######0.00")
   txtcrd = Format("0", "######0.00")
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   Set diaPostGJ = Nothing
   
End Sub



Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub
