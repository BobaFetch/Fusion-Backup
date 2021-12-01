VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaJcncl
   BorderStyle = 3 'Fixed Dialog
   Caption = "Close A Manufacturing Order"
   ClientHeight = 2505
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 6075
   ClipControls = 0 'False
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 2505
   ScaleWidth = 6075
   ShowInTaskbar = 0 'False
   Begin VB.ComboBox txtDte
      Height = 315
      Left = 1320
      TabIndex = 2
      Tag = "4"
      Top = 1820
      Width = 1095
   End
   Begin VB.ComboBox cmbPrt
      Height = 315
      Left = 1320
      TabIndex = 0
      Tag = "3"
      ToolTipText = "Contains Qualified Part Numbers (CO)"
      Top = 720
      Width = 3545
   End
   Begin VB.ComboBox cmbRun
      Height = 315
      Left = 1320
      TabIndex = 1
      Tag = "1"
      ToolTipText = "Contains Qualified Runs"
      Top = 1440
      Width = 1095
   End
   Begin VB.CommandButton cmdDel
      Caption = "M&O Close"
      Height = 315
      Left = 5040
      TabIndex = 3
      ToolTipText = "Press To Close thel MO"
      Top = 480
      Width = 875
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 5040
      TabIndex = 4
      TabStop = 0 'False
      Top = 0
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
      PictureUp = "diaJcncl.frx":0000
      PictureDn = "diaJcncl.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 5400
      Top = 1920
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 2505
      FormDesignWidth = 6075
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Date Closed"
      Height = 255
      Index = 4
      Left = 240
      TabIndex = 13
      Top = 1820
      Width = 1095
   End
   Begin VB.Label lblDte
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 5040
      TabIndex = 12
      Top = 1440
      Width = 855
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Complete"
      Height = 255
      Index = 1
      Left = 4080
      TabIndex = 11
      Top = 1440
      Width = 1215
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Part Number"
      Height = 255
      Index = 3
      Left = 240
      TabIndex = 10
      Top = 765
      Width = 1095
   End
   Begin VB.Label lblDsc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1320
      TabIndex = 9
      Top = 1080
      Width = 3135
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Run"
      Height = 255
      Index = 2
      Left = 240
      TabIndex = 8
      Top = 1440
      Width = 1095
   End
   Begin VB.Label lblStat
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 3240
      TabIndex = 7
      Top = 1440
      Width = 615
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Status"
      Height = 255
      Index = 0
      Left = 2520
      TabIndex = 6
      Top = 1440
      Width = 1095
   End
End
Attribute VB_Name = "diaJcncl"
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

Dim RdoQry As rdoQuery
Dim bOnLoad As Byte
Dim bGoodRun As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbPrt_Click()
   FindPart Me
   GetRuns
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If Len(cmbPrt) Then
      FindPart Me
      GetRuns
   End If
   
End Sub


Private Sub cmbRun_Click()
   bGoodRun = GetCurrRun()
   
End Sub


Private Sub cmbRun_LostFocus()
   cmbRun = CheckLen(cmbRun, 5)
   cmbRun = Format(Abs(Val(cmbRun)), "####0")
   bGoodRun = GetCurrRun()
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmbPrt = ""
   
End Sub


Private Sub cmdDel_Click()
   If bGoodRun = 0 Then
      MsgBox "Requires A Valid Run. See Help.", _
         vbInformation, Caption
   Else
      CloseMO
   End If
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   Dim l&
   If cmdHlp Then
      MouseCursor 13
      l& = WinHelp(Me.hWnd, sReportPath & "Esiprod.hlp", HELP_KEY, Caption)
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   sCurrForm = Caption
   sSql = "SELECT RUNREF,RUNNO FROM RunsTable WHERE " _
          & "RUNREF= ? AND RUNSTATUS='CO' "
   Set RdoQry = RdoCon.CreateQuery("", sSql)
   bOnLoad = True
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaJcncl = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDte = Format(Now, "mm/dd/yy")
   
End Sub

Public Sub FillCombo()
   Dim Cmb As rdoResultset
   On Error GoTo DiaErr1
   cmbPrt.Clear
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,RUNREF FROM " _
          & "RunsTable,PartTable WHERE PARTREF=RUNREF AND " _
          & "RUNSTATUS='CO' "
   bSqlRows = GetDataSet(Cmb, ES_FORWARD)
   If bSqlRows Then
      With Cmb
         Do Until .EOF
            'cmbPrt.AddItem "" & Trim(!PARTNUM)
            AddComboStr cmbPrt.hWnd, "" & Trim(!PARTNUM)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set Cmb = Nothing
   If cmbPrt.ListCount > 0 Then
      cmbPrt = cmbPrt.List(0)
      FindPart Me
      GetRuns
   End If
   Exit Sub
   
   DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub GetRuns()
   Dim RdoRns As rdoResultset
   Dim SPartRef As String
   
   On Error GoTo DiaErr1
   cmbRun.Clear
   SPartRef = Compress(cmbPrt)
   RdoQry(0) = SPartRef
   bSqlRows = GetQuerySet(RdoRns, RdoQry, ES_FORWARD)
   If bSqlRows Then
      With RdoRns
         Do Until .EOF
            AddComboStr cmbRun.hWnd, Format$(!RunNo, "####0")
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoRns = Nothing
   If cmbRun.ListCount > 0 Then
      cmbRun = cmbRun.List(0)
      bGoodRun = GetCurrRun()
   End If
   Exit Sub
   
   DiaErr1:
   sProcName = "closemo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Function GetCurrRun() As Byte
   Dim RdoRun As rdoResultset
   Dim lRunno As Long
   Dim sPart As String
   
   lRunno = Val(cmbRun)
   sPart = Compress(cmbPrt)
   On Error GoTo DiaErr1
   sSql = "SELECT RUNREF,RUNNO,RUNSTATUS,RUNCOMPLETE FROM RunsTable " _
          & "WHERE RUNREF='" & sPart & "' AND RUNNO=" & lRunno & " "
   bSqlRows = GetDataSet(RdoRun, ES_FORWARD)
   If bSqlRows Then
      lblStat = "" & Trim(RdoRun!RUNSTATUS)
      lblDte = Format(RdoRun!RUNCOMPLETE, "mm/dd/yy")
   Else
      lblStat = "**"
      lblDte = ""
   End If
   If lblStat = "CO" Then
      GetCurrRun = 1
   Else
      GetCurrRun = 0
      lblDte = ""
      lblStat = "**"
   End If
   Set RdoRun = Nothing
   Exit Function
   
   DiaErr1:
   sProcName = "closemo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** Part" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = vbBlack
   End If
   
End Sub

Private Sub lblStat_Change()
   If lblStat = "**" Then
      lblStat.ForeColor = ES_RED
   Else
      lblStat.ForeColor = vbBlack
   End If
   
End Sub


Public Sub CloseMO()
   Dim bResponse As Byte
   Dim lRunno As Long
   Dim sMsg As String
   Dim sPart As String
   
   sPart = Compress(cmbPrt)
   lRunno = Val(cmbRun)
   On Error GoTo DiaErr1
   sMsg = "This Closes The MO To All Functions." & vbCrLf _
          & "Do You Really Want To Close This MO?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      sSql = "UPDATE RunsTable SET RUNSTATUS='CL'," _
             & "RUNCLOSED='" & txtDte & "' " _
             & "WHERE RUNREF='" & sPart & "' AND " _
             & "RUNNO=" & lRunno & " "
      RdoCon.Execute sSql, rdExecDirect
      
      If Err = 0 Then
         sMsg = "The Status Was Changed From CO To CL." & vbCrLf _
                & "No Additional Action Can Be Executed."
         MsgBox sMsg, vbInformation, Caption
         FillCombo
      Else
         MsgBox "Couldn't Change The Run To Closed (CL).", vbInformation, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
   DiaErr1:
   sProcName = "closemo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtDte_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtDte_LostFocus()
   txtDte = CheckDate(txtDte)
   
End Sub
