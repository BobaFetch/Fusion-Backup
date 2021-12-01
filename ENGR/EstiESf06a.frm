VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form EstiESf06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy An Estimate Routing"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDsc 
      Height          =   288
      Left            =   1920
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "Use The Description Or Over Write The Entry"
      Top             =   2520
      Width           =   2772
   End
   Begin VB.TextBox txtRte 
      Height          =   288
      Left            =   1920
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Use The Part Number Or Over Write The Entry"
      Top             =   2160
      Width           =   2772
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "EstiESf06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdCbd 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4800
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Copy The Current Bid Routing To The Manufacturing Routing"
      Top             =   2160
      Width           =   875
   End
   Begin VB.ComboBox cmbBid 
      Height          =   288
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Select Or Enter A Bid Number (Full Bids Only)"
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4800
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5400
      Top             =   2880
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3045
      FormDesignWidth =   5730
   End
   Begin VB.Label lblDsc 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1920
      TabIndex        =   16
      Top             =   1680
      Width           =   2772
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copy An Estimate Routing To A Manufacturing Routing"
      Height          =   252
      Index           =   2
      Left            =   360
      TabIndex        =   15
      Top             =   240
      Width           =   4452
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Routing"
      Height          =   252
      Index           =   3
      Left            =   360
      TabIndex        =   14
      Top             =   2160
      Width           =   1212
   End
   Begin VB.Label lblPrt 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1920
      TabIndex        =   12
      Top             =   1320
      Width           =   2772
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   252
      Index           =   5
      Left            =   360
      TabIndex        =   11
      Top             =   1320
      Width           =   1212
   End
   Begin VB.Label lblDate 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3720
      TabIndex        =   10
      Top             =   950
      Width           =   950
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   252
      Index           =   4
      Left            =   3240
      TabIndex        =   9
      Top             =   960
      Width           =   732
   End
   Begin VB.Label lblTyp 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1920
      TabIndex        =   8
      Top             =   960
      Width           =   732
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   252
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   960
      Width           =   732
   End
   Begin VB.Label lblCls 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   312
      Left            =   1920
      TabIndex        =   6
      Top             =   600
      Width           =   252
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimate Number"
      Height          =   252
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   600
      Width           =   1572
   End
End
Attribute VB_Name = "EstiESf06a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'New 10/12/06
Option Explicit
Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bDelOld As Byte
Dim bGoodBid As Byte
'Dim bGoodOld As Byte
Dim bUnload As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmbBid_Click()
   If cmbBid.ListCount > 0 Then bGoodBid = GetTheBid()
   
End Sub


Private Sub cmbBid_LostFocus()
   cmbBid = CheckLen(cmbBid, 6)
   cmbBid = Format(Abs(Val(cmbBid)), "000000")
   If bCancel = 1 Then Exit Sub
   If cmbBid.ListCount > 0 Then bGoodBid = GetTheBid()
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = 1
   
End Sub


Private Sub cmdCbd_Click()
   Dim bResponse As Byte
   Dim sMsg As String
'   bGoodOld = GetManufacturingRouting()
'   If bGoodOld = 1 Then
   If GetManufacturingRouting() Then
      sMsg = "Note:" & vbCr & "This Part Number Has A Routing. Do You" & vbCrLf _
             & "Want To Continue And Replace The Routing?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbYes Then
         bDelOld = 1
      Else
         bDelOld = 0
         CancelTrans
         Exit Sub
      End If
   Else
      bDelOld = 0
   End If
   
   If bDelOld = 0 Then
      sMsg = "This Will Copy A Bid Routing To A New Routing." & vbCrLf _
         & "Are You Sure That You Wish Copy The Routing?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   End If
   If bResponse = vbYes Then
      CopyThisRouting
   Else
      CancelTrans
   End If
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3554
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
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set EstiESf06a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   Dim RdoCmb As ADODB.Recordset
   Dim iList As Integer
   
   'Get rid of Nulls-Easier Transport
   On Error Resume Next
   sSql = "UPDATE EsrtTable SET BIDRTENOTES='' WHERE BIDRTENOTES IS NULL"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   On Error GoTo DiaErr1
   cmbBid.Clear
   sSql = "SELECT BIDREF,BIDPRE,BIDCANCELED,BIDACCEPTED FROM EstiTable " _
          & "WHERE (BIDCANCELED=0 AND BIDCLASS='FULL') ORDER BY BIDREF DESC"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         cmbBid = Format(!BIDREF, "000000")
         lblCls = "" & Trim(!BIDPRE)
         Do Until .EOF
            iList = iList + 1
            If iList > 300 Then Exit Do
            AddComboStr cmbBid.hwnd, Format(!BIDREF, "000000")
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   End If
   If cmbBid.ListCount > 0 Then bGoodBid = GetTheBid()
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetTheBid() As Byte
   Dim RdoBid As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT BIDREF,BIDNUM,BIDPRE,BIDCLASS,BIDPART,BIDCUST," _
          & "BIDDATE,BIDRFQ,PARTREF,PARTNUM,PADESC FROM EstiTable," _
          & "PartTable WHERE (BIDPART=PARTREF AND BIDREF=" _
          & Val(cmbBid) & " AND BIDCANCELED=0 AND BIDCLASS='FULL')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBid, ES_FORWARD)
   If bSqlRows Then
      With RdoBid
         GetTheBid = 1
         lblCls = "" & Trim(!BIDPRE)
         lblTyp = "" & Trim(!BidClass)
         lblDate = "" & Format(!BIDDATE, "mm/dd/yyyy")
         lblPrt = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         txtRte = lblPrt
         txtDsc = lblDsc
         cmdCbd.Enabled = True
         ClearResultSet RdoBid
      End With
   Else
      GetTheBid = 0
      lblCls = ""
      lblTyp = ""
      lblDate = ""
      lblDsc = ""
      lblPrt = "*** Matching Full Bid Wasn't Found ***"
      MsgBox "This Bid Does Not Exist Or Doesn't Qualify." & vbCrLf _
         & "Only Full Bids May Have Routings.", _
         vbInformation, Caption
      cmdCbd.Enabled = False
      On Error Resume Next
      cmbBid.SetFocus
   End If
   Set RdoBid = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getthebid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Sub CopyThisRouting()
   Dim RdoCopy As ADODB.Recordset
   Dim lOldBid As Long
   
   lOldBid = Val(cmbBid)
   MouseCursor 13
   'On Error Resume Next
   On Error GoTo whoops
   clsADOCon.BeginTrans
   'If bDelOld = 1 Then
      'Opted to overwrite Routing
      sSql = "DELETE FROM RtopTable WHERE OPREF='" & Compress(txtRte) & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      'Clear it totally
      sSql = "DELETE FROM RthdTable WHERE RTREF='" & Compress(txtRte) & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
   'End If
   
   'Err.Clear
   'Insert New Header first
   sSql = "INSERT INTO RthdTable (RTREF,RTNUM,RTDESC)" & vbCrLf _
      & "VALUES('" & Compress(txtRte) & "','" & Trim(txtRte) & "','" & Trim(txtDsc) & "')"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   sSql = "SELECT BIDRTEOPNO,BIDRTESHOP,BIDRTECENTER,BIDRTESETUP," & vbCrLf _
      & "BIDRTEUNIT,BIDRTEQHRS,BIDRTEMHRS,BIDRTENOTES" & vbCrLf _
      & "FROM EsrtTable WHERE BIDRTEREF=" & lOldBid
''   If clsADOCon.GetDataSet(sSql,RdoCopy, ES_FORWARD) Then
''      With RdoCopy
''         Do Until .EOF
''            sSql = "INSERT INTO RtopTable" & vbCrLf _
''               & "(OPREF,OPNO,OPSHOP," & vbCrLf _
''               & "OPCENTER,OPSETUP,OPUNIT,OPQHRS,OPMHRS,OPCOMT)" & vbCrLf _
''               & "VALUES('" & Compress(txtRte) & "'," & !BIDRTEOPNO & ",'" & Trim(!BIDRTESHOP) & "'," & vbCrLf _
''               & "'" & Trim(!BIDRTECENTER) & "'," _
''               & !BIDRTESETUP & "," _
''               & !BIDRTEUNIT & "," _
''               & !BIDRTEQHRS & "," _
''               & !BIDRTEMHRS & ",'" _
''               & Trim(!BIDRTENOTES) & "')"
''            clsADOCon.ExecuteSQL sSql 'rdExecDirect
''            .MoveNext
''         Loop
''      End With
      
   sSql = "INSERT INTO RtopTable" & vbCrLf _
      & "(OPREF,OPNO,OPSHOP," & vbCrLf _
      & "OPCENTER,OPSETUP,OPUNIT,OPQHRS,OPMHRS,OPCOMT)" & vbCrLf _
      & "SELECT '" & Compress(txtRte) & "',BIDRTEOPNO,RTRIM(BIDRTESHOP)," & vbCrLf _
      & "RTRIM(BIDRTECENTER),BIDRTESETUP,BIDRTEUNIT,BIDRTEQHRS,BIDRTEMHRS,RTRIM(BIDRTENOTES)" & vbCrLf _
      & "FROM EsrtTable WHERE BIDRTEREF=" & lOldBid
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
   MouseCursor 0
   If clsADOCon.RowsAffected > 0 Then
      clsADOCon.CommitTrans
      SysMsg "Routing successfully copied.", True
      txtRte = ""
      txtDsc = ""
   Else
      clsADOCon.RollbackTrans
      MsgBox "No Routing Steps Were Found For This Estimate.", _
         vbInformation, Caption
   End If
   Exit Sub
   
whoops:
   clsADOCon.RollbackTrans
   sProcName = "CopyThisRouting"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub lblPrt_Change()
   If Left(lblPrt, 6) = "*** Ma" Then lblPrt.ForeColor = ES_RED _
           Else lblPrt.ForeColor = vbBlack
   
End Sub

Private Sub txtRte_LostFocus()
   txtRte = CheckLen(txtRte, 30)
   
End Sub

Private Function GetManufacturingRouting() As Boolean
   Dim RdoMrte As ADODB.Recordset
   sSql = "SELECT RTREF FROM RthdTable WHERE RTREF='" _
          & Compress(txtRte) & "'"
   GetManufacturingRouting = clsADOCon.GetDataSet(sSql, RdoMrte, ES_FORWARD)
   
End Function

