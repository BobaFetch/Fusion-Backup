VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaGLe04a
   BorderStyle = 3 'Fixed Dialog
   Caption = "Fiscal Years"
   ClientHeight = 5565
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 5280
   ClipControls = 0 'False
   ControlBox = 0 'False
   ForeColor = &H00000000&
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 5565
   ScaleWidth = 5280
   ShowInTaskbar = 0 'False
   Visible = 0 'False
   Begin VB.TextBox txtPer
      Height = 285
      Index = 10
      Left = 4080
      TabIndex = 42
      Tag = "2"
      Top = 4800
      Width = 615
   End
   Begin VB.TextBox txtEnd
      Height = 285
      Index = 10
      Left = 3000
      TabIndex = 41
      Tag = "4"
      Top = 4800
      Width = 915
   End
   Begin VB.TextBox txtBeg
      Height = 285
      Index = 10
      Left = 1920
      TabIndex = 40
      Tag = "4"
      Top = 4800
      Width = 915
   End
   Begin VB.TextBox txtYer
      Enabled = 0 'False
      Height = 285
      Index = 10
      Left = 960
      TabIndex = 39
      Tag = "2"
      Top = 4800
      Width = 645
   End
   Begin VB.TextBox txtPer
      Height = 285
      Index = 9
      Left = 4080
      TabIndex = 38
      Tag = "2"
      Top = 4440
      Width = 615
   End
   Begin VB.TextBox txtEnd
      Height = 285
      Index = 9
      Left = 3000
      TabIndex = 37
      Tag = "4"
      Top = 4440
      Width = 915
   End
   Begin VB.TextBox txtBeg
      Height = 285
      Index = 9
      Left = 1920
      TabIndex = 36
      Tag = "4"
      Top = 4440
      Width = 915
   End
   Begin VB.TextBox txtYer
      Enabled = 0 'False
      Height = 285
      Index = 9
      Left = 960
      TabIndex = 35
      Tag = "2"
      Top = 4440
      Width = 645
   End
   Begin VB.TextBox txtYer
      Enabled = 0 'False
      Height = 285
      Index = 8
      Left = 960
      TabIndex = 31
      Tag = "2"
      Top = 4080
      Width = 645
   End
   Begin VB.TextBox txtBeg
      Height = 285
      Index = 8
      Left = 1920
      TabIndex = 32
      Tag = "4"
      Top = 4080
      Width = 915
   End
   Begin VB.TextBox txtEnd
      Height = 285
      Index = 8
      Left = 3000
      TabIndex = 33
      Tag = "4"
      Top = 4080
      Width = 915
   End
   Begin VB.TextBox txtPer
      Height = 285
      Index = 8
      Left = 4080
      TabIndex = 34
      Tag = "2"
      Top = 4080
      Width = 615
   End
   Begin VB.TextBox txtYer
      Enabled = 0 'False
      Height = 285
      Index = 7
      Left = 960
      TabIndex = 27
      Tag = "2"
      Top = 3720
      Width = 645
   End
   Begin VB.TextBox txtBeg
      Height = 285
      Index = 7
      Left = 1920
      TabIndex = 28
      Tag = "4"
      Top = 3720
      Width = 915
   End
   Begin VB.TextBox txtEnd
      Height = 285
      Index = 7
      Left = 3000
      TabIndex = 29
      Tag = "4"
      Top = 3720
      Width = 915
   End
   Begin VB.TextBox txtPer
      Height = 285
      Index = 7
      Left = 4080
      TabIndex = 30
      Tag = "2"
      Top = 3720
      Width = 615
   End
   Begin VB.TextBox txtYer
      Enabled = 0 'False
      Height = 285
      Index = 6
      Left = 960
      TabIndex = 23
      Tag = "2"
      Top = 3360
      Width = 645
   End
   Begin VB.TextBox txtBeg
      Height = 285
      Index = 6
      Left = 1920
      TabIndex = 24
      Tag = "4"
      Top = 3360
      Width = 915
   End
   Begin VB.TextBox txtEnd
      Height = 285
      Index = 6
      Left = 3000
      TabIndex = 25
      Tag = "4"
      Top = 3360
      Width = 915
   End
   Begin VB.TextBox txtPer
      Height = 285
      Index = 6
      Left = 4080
      TabIndex = 26
      Tag = "2"
      Top = 3360
      Width = 615
   End
   Begin VB.TextBox txtYer
      Enabled = 0 'False
      Height = 285
      Index = 5
      Left = 960
      TabIndex = 19
      Tag = "2"
      Top = 3000
      Width = 645
   End
   Begin VB.TextBox txtBeg
      Height = 285
      Index = 5
      Left = 1920
      TabIndex = 20
      Tag = "4"
      Top = 3000
      Width = 915
   End
   Begin VB.TextBox txtEnd
      Height = 285
      Index = 5
      Left = 3000
      TabIndex = 21
      Tag = "4"
      Top = 3000
      Width = 915
   End
   Begin VB.TextBox txtPer
      Height = 285
      Index = 5
      Left = 4080
      TabIndex = 22
      Tag = "2"
      Top = 3000
      Width = 615
   End
   Begin VB.TextBox txtYer
      Enabled = 0 'False
      Height = 285
      Index = 4
      Left = 960
      TabIndex = 15
      Tag = "2"
      Top = 2640
      Width = 645
   End
   Begin VB.TextBox txtBeg
      Height = 285
      Index = 4
      Left = 1920
      TabIndex = 16
      Tag = "4"
      Top = 2640
      Width = 915
   End
   Begin VB.TextBox txtEnd
      Height = 285
      Index = 4
      Left = 3000
      TabIndex = 17
      Tag = "4"
      Top = 2640
      Width = 915
   End
   Begin VB.TextBox txtPer
      Height = 285
      Index = 4
      Left = 4080
      TabIndex = 18
      Tag = "2"
      Top = 2640
      Width = 615
   End
   Begin VB.TextBox txtYer
      Enabled = 0 'False
      Height = 285
      Index = 3
      Left = 960
      TabIndex = 12
      Tag = "2"
      Top = 2280
      Width = 645
   End
   Begin VB.TextBox txtBeg
      Height = 285
      Index = 3
      Left = 1920
      TabIndex = 13
      Tag = "4"
      Top = 2280
      Width = 915
   End
   Begin VB.TextBox txtEnd
      Height = 285
      Index = 3
      Left = 3000
      TabIndex = 43
      Tag = "4"
      Top = 2280
      Width = 915
   End
   Begin VB.TextBox txtPer
      Height = 285
      Index = 3
      Left = 4080
      TabIndex = 14
      Tag = "2"
      Top = 2280
      Width = 615
   End
   Begin VB.TextBox txtYer
      Enabled = 0 'False
      Height = 285
      Index = 2
      Left = 960
      TabIndex = 8
      Tag = "2"
      Top = 1920
      Width = 645
   End
   Begin VB.TextBox txtBeg
      Height = 285
      Index = 2
      Left = 1920
      TabIndex = 9
      Tag = "4"
      Top = 1920
      Width = 915
   End
   Begin VB.TextBox txtEnd
      Height = 285
      Index = 2
      Left = 3000
      TabIndex = 10
      Tag = "4"
      Top = 1920
      Width = 915
   End
   Begin VB.TextBox txtPer
      Height = 285
      Index = 2
      Left = 4080
      TabIndex = 11
      Tag = "2"
      Top = 1920
      Width = 615
   End
   Begin VB.TextBox txtYer
      Enabled = 0 'False
      Height = 285
      Index = 1
      Left = 960
      TabIndex = 4
      Tag = "2"
      Top = 1560
      Width = 645
   End
   Begin VB.TextBox txtBeg
      Height = 285
      Index = 1
      Left = 1920
      TabIndex = 5
      Tag = "4"
      Top = 1560
      Width = 915
   End
   Begin VB.TextBox txtEnd
      Height = 285
      Index = 1
      Left = 3000
      TabIndex = 6
      Tag = "4"
      Top = 1560
      Width = 915
   End
   Begin VB.TextBox txtPer
      Height = 285
      Index = 1
      Left = 4080
      TabIndex = 7
      Tag = "2"
      Top = 1560
      Width = 615
   End
   Begin VB.CommandButton cmdUpd
      Caption = "&Update"
      Height = 315
      Left = 4320
      TabIndex = 58
      ToolTipText = "Update Structure And Associated Entries"
      Top = 560
      Width = 875
   End
   Begin VB.TextBox txtPer
      Height = 285
      Index = 0
      Left = 4080
      TabIndex = 3
      Tag = "2"
      Top = 1200
      Width = 615
   End
   Begin VB.TextBox txtEnd
      Height = 285
      Index = 0
      Left = 3000
      TabIndex = 2
      Tag = "4"
      Top = 1200
      Width = 915
   End
   Begin VB.TextBox txtBeg
      Height = 285
      Index = 0
      Left = 1920
      TabIndex = 1
      Tag = "4"
      Top = 1200
      Width = 915
   End
   Begin VB.TextBox txtYer
      Enabled = 0 'False
      Height = 285
      Index = 0
      Left = 960
      TabIndex = 0
      Tag = "2"
      Top = 1200
      Width = 645
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 4320
      TabIndex = 44
      TabStop = 0 'False
      Top = 90
      Width = 875
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 3720
      Top = 120
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 5565
      FormDesignWidth = 5280
   End
   Begin Threed.SSCommand cmdDn
      Height = 375
      Left = 4800
      TabIndex = 56
      TabStop = 0 'False
      ToolTipText = "Next Page (Page Down)"
      Top = 4800
      Width = 375
      _Version = 65536
      _ExtentX = 661
      _ExtentY = 661
      _StockProps = 78
      ForeColor = -2147483630
      Enabled = 0 'False
      RoundedCorners = 0 'False
      Outline = 0 'False
      Picture = "diaGlfyr.frx":0000
   End
   Begin Threed.SSCommand cmdUp
      Height = 375
      Left = 4800
      TabIndex = 57
      TabStop = 0 'False
      ToolTipText = "Last Page (Page Up)"
      Top = 4440
      Width = 375
      _Version = 65536
      _ExtentX = 661
      _ExtentY = 661
      _StockProps = 78
      ForeColor = -2147483630
      Enabled = 0 'False
      RoundedCorners = 0 'False
      Outline = 0 'False
      Picture = "diaGlfyr.frx":0502
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 60
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
      PictureUp = "diaGlfyr.frx":0A04
      PictureDn = "diaGlfyr.frx":0B4A
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Years     "
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = -1 'True
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 255
      Index = 15
      Left = 960
      TabIndex = 65
      Top = 960
      Width = 675
   End
   Begin VB.Label lblMsg
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 700
      Underline = 0 'False
      Italic = -1 'True
      Strikethrough = 0 'False
      EndProperty
      ForeColor = &H00800000&
      Height = 255
      Left = 960
      TabIndex = 64
      Top = 5200
      Width = 3000
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Enter Or Revise General Ledger Fiscal Years"
      Height = 255
      Index = 14
      Left = 240
      TabIndex = 63
      Top = 600
      Width = 3975
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Year"
      Height = 255
      Index = 13
      Left = 240
      TabIndex = 62
      Top = 4800
      Width = 800
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Year"
      Height = 255
      Index = 12
      Left = 240
      TabIndex = 61
      Top = 4440
      Width = 800
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Year"
      Height = 255
      Index = 0
      Left = 240
      TabIndex = 59
      Top = 1200
      Width = 800
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Year"
      Height = 255
      Index = 11
      Left = 240
      TabIndex = 55
      Top = 4080
      Width = 800
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Year"
      Height = 255
      Index = 10
      Left = 240
      TabIndex = 54
      Top = 3720
      Width = 800
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Year"
      Height = 255
      Index = 9
      Left = 240
      TabIndex = 53
      Top = 3360
      Width = 800
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Year"
      Height = 255
      Index = 8
      Left = 240
      TabIndex = 52
      Top = 3000
      Width = 800
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Year"
      Height = 255
      Index = 7
      Left = 240
      TabIndex = 51
      Top = 2640
      Width = 800
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Year"
      Height = 255
      Index = 6
      Left = 240
      TabIndex = 50
      Top = 2280
      Width = 800
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Year"
      Height = 255
      Index = 5
      Left = 240
      TabIndex = 49
      Top = 1920
      Width = 800
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Periods  "
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = -1 'True
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 255
      Index = 4
      Left = 4080
      TabIndex = 48
      Top = 960
      Width = 800
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Ending         "
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = -1 'True
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 255
      Index = 3
      Left = 3000
      TabIndex = 47
      Top = 960
      Width = 915
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Starting        "
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = -1 'True
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 255
      Index = 2
      Left = 1920
      TabIndex = 46
      Top = 960
      Width = 915
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Year"
      Height = 255
      Index = 1
      Left = 240
      TabIndex = 45
      Top = 1560
      Width = 800
   End
End
Attribute VB_Name = "diaGle04a"
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

Dim bOnLoad As Byte
Dim bGoodYears As Byte
Dim bDataChanged As Boolean

Dim sMonths(14, 3) As String
Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

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


Private Sub cmdUpd_Click()
   ' UnderCons.Show
   UpDateFiscalYears
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillYears
      bOnLoad = False
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   sCurrForm = Caption
   bOnLoad = True
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim bResponse As Byte
   Dim smsg As String
   If bDataChanged Then
      smsg = "The Fiscal Year Data Has Changed Or Just Initialized." & vbCr _
             & "Do You Wish To Leave Without Updating?"
      bResponse = MsgBox(smsg, ES_NOQUESTION, Caption)
      If bResponse = vbNo Then Cancel = True
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaGLe03a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Public Sub FillYears()
   Dim RdoFsc As rdoResultset
   Dim a As Integer
   Dim i As Integer
   
   On Error GoTo DiaErr1
   bGoodYears = CheckYears
   If bGoodYears = 0 Then
      a = Year(Now)
      For i = 0 To 9
         txtYer(i) = a
         a = a + 1
      Next
      txtYer(i) = a
      InitializeYears
   Else
      sProcName = "fillyears"
      sSql = "SELECT FYYEAR,FYSTART,FYEND,FYPERIODS " _
             & "FROM GlfyTable "
      bSqlRows = GetDataSet(RdoFsc, ES_FORWARD)
      If bSqlRows Then
         With RdoFsc
            i = 0
            lblMsg.Visible = True
            lblMsg = "Loading Fiscal Years."
            lblMsg.Refresh
            Sleep 500
            Do Until .EOF
               txtYer(i) = Format(!FYYEAR, "0000")
               txtBeg(i) = Format(!FYSTART, "mm/dd/yy")
               txtEnd(i) = Format(!FYEND, "mm/dd/yy")
               txtPer(i) = Format(!FYPERIODS, "#0")
               i = i + 1
               If i > 10 Then Exit Do
               .MoveNext
            Loop
            .Cancel
         End With
         lblMsg = "Fiscal Years Loaded."
         lblMsg.Refresh
         Sleep 1000
         lblMsg.Visible = False
         bDataChanged = False
      End If
   End If
   Set RdoFsc = Nothing
   Exit Sub
   
   DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   
End Sub



Private Sub txtBeg_Change(Index As Integer)
   bDataChanged = True
   
End Sub

Private Sub txtBeg_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtBeg_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtBeg_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyDate KeyAscii
   
End Sub


Private Sub txtBeg_LostFocus(Index As Integer)
   txtBeg(Index) = CheckDate(txtBeg(Index))
   
End Sub

Private Sub txtEnd_Change(Index As Integer)
   bDataChanged = True
   
End Sub

Private Sub txtEnd_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtEnd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtEnd_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyDate KeyAscii
   
End Sub


Private Sub txtEnd_LostFocus(Index As Integer)
   txtEnd(Index) = CheckDate(txtEnd(Index))
   
End Sub

Private Sub txtPer_Change(Index As Integer)
   bDataChanged = True
   
End Sub

Private Sub txtPer_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtPer_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtPer_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyDate KeyAscii
   
End Sub


Private Sub txtPer_LostFocus(Index As Integer)
   txtPer(Index) = CheckLen(txtPer(Index), 2)
   txtPer(Index) = Format(Abs(Val(txtPer(Index))), "#0")
   If Val(txtPer(Index)) > 13 Or Val(txtPer(Index)) < 12 Then
      Beep
      txtPer(Index) = "12"
   End If
   
End Sub

Private Sub txtYer_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtYer_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtYer_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub



Public Function CheckYears() As Byte
   Dim RdoFyr As rdoResultset
   sSql = "SELECT FYYEAR FROM GlfyTable "
   bSqlRows = GetDataSet(RdoFyr, ES_FORWARD)
   If bSqlRows Then
      CheckYears = 1
      RdoFyr.Cancel
   Else
      CheckYears = 0
   End If
   Set RdoFyr = Nothing
   sProcName = "checkyears"
   
End Function

Public Sub InitializeYears()
   Dim a As Integer
   Dim i As Integer
   MouseCursor 13
   lblMsg.Visible = True
   lblMsg = "Initializing Fiscal Years."
   lblMsg.Refresh
   For i = 0 To 9
      txtBeg(i) = "01/01/" & Right(txtYer(i), 2)
      txtEnd(i) = "12/31/" & Right(txtYer(i), 2)
      txtPer(i) = 12
   Next
   txtBeg(i) = "01/01/" & Right(txtYer(i), 2)
   txtEnd(i) = "12/31/" & Right(txtYer(i), 2)
   txtPer(i) = 12
   Err = 0
   RdoCon.BeginTrans
   For i = 0 To 9
      sSql = "INSERT INTO GlfyTable (FYYEAR,FYSTART," _
             & "FYEND) VALUES(" & Val(txtYer(i)) & ",'" _
             & Format(txtBeg(i), "mm/dd/yyyy") & "','" _
             & Format(txtEnd(i), "mm/dd/yyyy") & "')"
      lblMsg = lblMsg & "."
      lblMsg.Refresh
      RdoCon.Execute sSql, rdExecDirect
   Next
   sSql = "INSERT INTO GlfyTable (FYYEAR,FYSTART," _
          & "FYEND) VALUES(" & Val(txtYer(i)) & ",'" _
          & Format(txtBeg(i), "mm/dd/yyyy") & "','" _
          & Format(txtEnd(i), "mm/dd/yyyy") & "')"
   lblMsg = lblMsg & "."
   lblMsg.Refresh
   RdoCon.Execute sSql, rdExecDirect
   If Err = 0 Then
      RdoCon.CommitTrans
      bDataChanged = True
      lblMsg = "Fiscal Years Initialized."
      lblMsg.Refresh
      Sleep 500
      MouseCursor 0
      lblMsg.Visible = False
   Else
      MouseCursor 0
      RdoCon.RollbackTrans
      MsgBox "Unable To Establish Fiscal Years.", _
         vbExclamation, Caption
      lblMsg = "No Current Fiscal Years Established."
   End If
   sProcName = "initializeye"
   
End Sub


Public Sub UpDateFiscalYears()
   Dim bResponse As Byte
   Dim iDay As Byte
   Dim i As Integer
   Dim iYear As Integer
   Dim smsg As String
   On Error GoTo DiaErr1
   smsg = "This Function Establishes Monthly Data " & vbCr _
          & "Based On Selected Periods. Continue?"
   bResponse = MsgBox(smsg, ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      CancelTrans
   Else
      cmdUpd.Enabled = False
      MouseCursor 13
      On Error Resume Next
      Err = 0
      RdoCon.BeginTrans
      For i = 0 To 10
         iYear = Val(txtYer(i))
         If iYear = 2000 Or iYear = 2004 Or iYear = 2008 Or iYear = 20012 Then
            iDay = 29
         Else
            iDay = 28
         End If
         If Val(txtPer(i)) = 12 Then
            sSql = "UPDATE GlfyTable SET " _
                   & "FYSTART='" & Left(txtBeg(i), 6) & iYear & "'," _
                   & "FYEND='" & Left(txtEnd(i), 6) & iYear & "'," _
                   & "FYPERIODS=" & Val(txtPer(i)) & "," _
                   & "FYPERSTART1='" & Left(txtBeg(i), 6) & iYear & "'," _
                   & "FYPEREND1='01/31/" & iYear & "'," _
                   & "FYPERSTART2='02/01/" & iYear & "'," _
                   & "FYPEREND2='02/" & iDay & "/" & iYear & "'," _
                   & "FYPERSTART3='03/01/" & iYear & "'," _
                   & "FYPEREND3='03/31/" & iYear & "'," _
                   & "FYPERSTART4='04/01/" & iYear & "'," _
                   & "FYPEREND4='04/30/" & iYear & "'," _
                   & "FYPERSTART5='05/01/" & iYear & "'," _
                   & "FYPEREND5='05/31/" & iYear & "'," _
                   & "FYPERSTART6='06/01/" & iYear & "'," _
                   & "FYPEREND6='06/30/" & iYear & "'," _
                   & "FYPERSTART7='07/01/" & iYear & "'," _
                   & "FYPEREND7='07/31/" & iYear & "'," _
                   & "FYPERSTART8='08/01/" & iYear & "'," _
                   & "FYPEREND8='08/31/" & iYear & "'," _
                   & "FYPERSTART9='09/01/" & iYear & "'," _
                   & "FYPEREND9='09/30/" & iYear & "',"
            sSql = sSql _
                   & "FYPERSTART10='10/01/" & iYear & "'," _
                   & "FYPEREND10='10/31/" & iYear & "'," _
                   & "FYPERSTART11='11/01/" & iYear & "'," _
                   & "FYPEREND11='11/30/" & iYear & "'," _
                   & "FYPERSTART12='12/01/" & iYear & "'," _
                   & "FYPEREND12='" & Left(txtEnd(i), 6) & iYear & "'," _
                   & "FYPERSTART13=NULL," _
                   & "FYPEREND13=NULL " _
                   & "WHERE FYYEAR=" & iYear & " "
            RdoCon.Execute sSql, rdExecDirect
         Else
            GetMonth i
            sSql = "UPDATE GlfyTable SET " _
                   & "FYSTART='" & Left(txtBeg(i), 6) & iYear & "'," _
                   & "FYEND='" & Left(txtEnd(i), 6) & iYear & "'," _
                   & "FYPERIODS=" & Val(txtPer(i)) & "," _
                   & "FYPERSTART1='" & sMonths(1, 0) & "'," _
                   & "FYPEREND1='" & sMonths(1, 1) & "'," _
                   & "FYPERSTART2='" & sMonths(2, 0) & "'," _
                   & "FYPEREND2='" & sMonths(2, 1) & "'," _
                   & "FYPERSTART3='" & sMonths(3, 0) & "'," _
                   & "FYPEREND3='" & sMonths(3, 1) & "'," _
                   & "FYPERSTART4='" & sMonths(4, 0) & "'," _
                   & "FYPEREND4='" & sMonths(4, 1) & "'," _
                   & "FYPERSTART5='" & sMonths(5, 0) & "'," _
                   & "FYPEREND5='" & sMonths(5, 1) & "'," _
                   & "FYPERSTART6='" & sMonths(6, 0) & "'," _
                   & "FYPEREND6='" & sMonths(6, 1) & "'," _
                   & "FYPERSTART7='" & sMonths(7, 0) & "'," _
                   & "FYPEREND7='" & sMonths(7, 1) & "'," _
                   & "FYPERSTART8='" & sMonths(8, 0) & "'," _
                   & "FYPEREND8='" & sMonths(8, 1) & "'," _
                   & "FYPERSTART9='" & sMonths(9, 0) & "'," _
                   & "FYPEREND9='" & sMonths(9, 1) & "',"
            sSql = sSql _
                   & "FYPERSTART10='" & sMonths(10, 0) & "'," _
                   & "FYPEREND10='" & sMonths(10, 1) & "'," _
                   & "FYPERSTART11='" & sMonths(11, 0) & "'," _
                   & "FYPEREND11='" & sMonths(11, 1) & "'," _
                   & "FYPERSTART12='" & sMonths(12, 0) & "'," _
                   & "FYPEREND12='" & sMonths(12, 1) & "'," _
                   & "FYPERSTART13='" & sMonths(13, 0) & "'," _
                   & "FYPEREND13='" & Left(txtEnd(i), 6) & iYear & "'" _
                   & "WHERE FYYEAR=" & iYear & " "
            RdoCon.Execute sSql, rdExecDirect
         End If
      Next
      If Err = 0 Then
         MouseCursor 0
         RdoCon.CommitTrans
         MsgBox "Fiscal Years Successfully Updated.", _
            vbInformation, Caption
         bDataChanged = False
      Else
         MouseCursor 0
         RdoCon.RollbackTrans
         MsgBox "Couldn't Successfuly Update Fiscal Years.", _
            vbExclamation, Caption
         bDataChanged = False
      End If
   End If
   cmdUpd.Enabled = True
   Exit Sub
   
   DiaErr1:
   sProcName = "updatefisc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   
End Sub

Public Sub GetMonth(iIndex As Integer)
   Dim i As Integer
   Dim sBeg As Date
   Dim sEnd As Date
   Erase sMonths
   
   sBeg = Format(txtBeg(iIndex), "mm/dd/yyyy")
   sEnd = Format(sBeg + 28, "mm/dd/yyyy")
   sMonths(1, 0) = Format(sBeg, "mm/dd/yyyy")
   sMonths(1, 1) = Format(sEnd, "mm/dd/yyyy")
   For i = 2 To 12
      sBeg = Format(sEnd + 1, "mm/dd/yyyy")
      sEnd = Format(sBeg + 28, "mm/dd/yyyy")
      sMonths(i, 0) = Format(sBeg, "mm/dd/yyyy")
      sMonths(i, 1) = Format(sEnd, "mm/dd/yyyy")
   Next
   sBeg = Format(sEnd + 1, "mm/dd/yyyy")
   sEnd = Format(sBeg + 28, "mm/dd/yyyy")
   sMonths(i, 0) = Format(sBeg, "mm/dd/yyyy")
   sMonths(i, 1) = Format(txtEnd(iIndex), "mm/dd/yyyy")
   
End Sub
