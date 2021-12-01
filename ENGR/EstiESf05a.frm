VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form EstiESf05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy An Estimate"
   ClientHeight    =   2880
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
   ScaleHeight     =   2880
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optEdit 
      Alignment       =   1  'Right Justify
      Caption         =   "Edit After Copy"
      Height          =   192
      Left            =   360
      TabIndex        =   2
      ToolTipText     =   "Return To Full Or Qwik Bid Feature After The Estimate Has Been Copied"
      Top             =   2520
      Width           =   1750
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   4200
      Top             =   0
   End
   Begin VB.TextBox txtBid 
      Height          =   288
      Left            =   1920
      TabIndex        =   1
      Tag             =   "1"
      Top             =   2160
      Width           =   732
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "EstiESf05a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
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
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Copy The Current Bid To The New Estimate"
      Top             =   2160
      Width           =   875
   End
   Begin VB.ComboBox cmbBid 
      Height          =   288
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Select Or Enter A Bid Number"
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4800
      TabIndex        =   3
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
      FormDesignHeight=   2880
      FormDesignWidth =   5730
   End
   Begin VB.Label lblNxt 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1920
      TabIndex        =   19
      Top             =   120
      Width           =   792
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Next Estimate"
      Height          =   252
      Index           =   31
      Left            =   360
      TabIndex        =   18
      Top             =   120
      Width           =   1452
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Estimate"
      Height          =   252
      Index           =   3
      Left            =   360
      TabIndex        =   17
      Top             =   2160
      Width           =   1212
   End
   Begin VB.Label lblPrt 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1920
      TabIndex        =   15
      Top             =   1680
      Width           =   2772
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   252
      Index           =   5
      Left            =   360
      TabIndex        =   14
      Top             =   1680
      Width           =   1212
   End
   Begin VB.Label lblDate 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   3960
      TabIndex        =   13
      Top             =   960
      Width           =   950
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   252
      Index           =   4
      Left            =   3240
      TabIndex        =   12
      Top             =   960
      Width           =   732
   End
   Begin VB.Label lblCust 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1920
      TabIndex        =   10
      Top             =   1296
      Width           =   3372
   End
   Begin VB.Label lblNik 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1920
      TabIndex        =   9
      Top             =   960
      Width           =   1212
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   252
      Index           =   2
      Left            =   360
      TabIndex        =   8
      Top             =   960
      Width           =   1212
   End
   Begin VB.Label lblTyp 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   3960
      TabIndex        =   7
      Top             =   600
      Width           =   732
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   252
      Index           =   1
      Left            =   3240
      TabIndex        =   6
      Top             =   600
      Width           =   732
   End
   Begin VB.Label lblCls 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   312
      Left            =   1920
      TabIndex        =   5
      Top             =   600
      Width           =   252
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimate Number"
      Height          =   252
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   600
      Width           =   1572
   End
End
Attribute VB_Name = "EstiESf05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'New 10/2/06
Option Explicit
Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodBid As Byte
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
   
   bResponse = GetNewBid()
   If bResponse = 0 Then
      sMsg = "This Copies An Entire Bid To The New Bid." & vbCrLf _
             & "Are You Sure That You Wish Copy The Bid?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         CopyThisBid
      Else
         CancelTrans
      End If
   End If
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3553
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      GetNextBid Me
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   GetOptions
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   If bUnload = 0 Then FormUnload
   Set EstiESf05a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   Dim RdoCmb As ADODB.Recordset
   Dim iList As Integer
   On Error GoTo DiaErr1
   cmbBid.Clear
   sSql = "SELECT BIDREF,BIDPRE,BIDCANCELED,BIDACCEPTED FROM EstiTable " _
          & "WHERE BIDCANCELED=0 ORDER BY BIDREF DESC"
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
          & "BIDDATE,BIDRFQ,CUREF,CUNICKNAME,CUNAME,PARTREF,PARTNUM " _
          & "FROM EstiTable,CustTable,PartTable WHERE (BIDCUST=CUREF " _
          & "AND BIDPART=PARTREF) AND BIDREF=" & Val(cmbBid) & " " _
          & "AND BIDCANCELED=0 "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBid, ES_FORWARD)
   If bSqlRows Then
      With RdoBid
         GetTheBid = 1
         lblCls = "" & Trim(!BIDPRE)
         lblNik = "" & Trim(!CUNICKNAME)
         lblTyp = "" & Trim(!BidClass)
         lblCust = "" & Trim(!CUNAME)
         lblDate = "" & Format(!BIDDATE, "mm/dd/yyyy")
         lblPrt = "" & Trim(!PartNum)
         cmdCbd.Enabled = True
         ClearResultSet RdoBid
      End With
   Else
      GetTheBid = 0
      lblCls = ""
      lblNik = ""
      lblTyp = ""
      lblCust = "*** Estimate Wasn't Found Or Doesn't Qualify ***"
      lblDate = ""
      lblPrt = ""
      MsgBox "This Bid Does Not Exist Or Doesn't Qualify." & vbCrLf _
         & "See Help For Instructions On Canceling A Bid.", _
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


Private Sub CopyThisBid()
   Dim bResponse As Byte
   Dim lOldBid As Long
   Dim lNewBid As Long
   Dim sMsg As String
   
   Timer1.Enabled = False
   GetNextBid Me
   lOldBid = Val(cmbBid)
   lNewBid = Val(txtBid)
   
   'Drop them in case they linger (should never be)
   On Error Resume Next
   sSql = "DROP TABLE #Esti"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   sSql = "DROP TABLE #Esbm"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   sSql = "DROP TABLE #Esos"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   sSql = "DROP TABLE #Esrt"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   Err.Clear
   MouseCursor 13
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   'EstiTable
   sSql = "SELECT * INTO #Esti from EstiTable where BIDREF=" & lOldBid & " "
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   sSql = "UPDATE #Esti SET BIDREF=" & lNewBid & " "
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   sSql = "INSERT INTO EstiTable SELECT * FROM #Esti "
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   sSql = "UPDATE EstiTable SET BIDNUM='" & Format(lNewBid, "000000") & "'," _
          & "BIDDATE='" & Format(Now, "mm/dd/yyyy") & "',BIDACCEPTED=0,BIDLOCKED=0 " _
          & "WHERE BIDREF=" & lNewBid & " "
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   'EsbmTable
   sSql = "SELECT * INTO #Esbm from EsbmTable where BIDBOMREF=" & lOldBid & " "
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   sSql = "UPDATE #Esbm SET BIDBOMREF=" & lNewBid & " "
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   sSql = "INSERT INTO EsbmTable SELECT * FROM #Esbm "
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   'EsosTable
   sSql = "SELECT * INTO #Esos from EsosTable where BIDOSREF=" & lOldBid & " "
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   sSql = "UPDATE #Esos SET BIDOSREF=" & lNewBid & " "
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   sSql = "INSERT INTO EsosTable SELECT * FROM #Esos "
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   'EsrtTable
   sSql = "SELECT * INTO #Esrt from EsrtTable where BIDRTEREF=" & lOldBid & " "
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   sSql = "UPDATE #Esrt SET BIDRTEREF=" & lNewBid & " "
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   sSql = "INSERT INTO EsrtTable SELECT * FROM #Esrt "
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   Sleep 500
   MouseCursor 0
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      If optEdit = vbChecked Then
         MsgBox "Estimate " & cmbBid & " Has Been Copied to " & txtBid & ".", _
            vbInformation, Caption
         bUnload = 1
         If RunningBeta Then
            If Left(lblTyp, 1) = "F" Then
               ' MM TODO: ppiESe02a.Show
            Else
               ' MM TODO: ppiESe01a.Show
            End If
         Else
            If Left(lblTyp, 1) = "F" Then
               EstiESe02a.Show
            Else
               EstiESe01a.Show
            End If
         End If
         Unload Me
      Else
         sMsg = "Estimate " & cmbBid & " Has Been Copied to " & txtBid & "." & vbCrLf _
                & "Would You Like To Print the New Estimate."
         bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
         If bResponse = vbNo Then
            FillCombo
            GetNextBid Me
            Timer1.Enabled = True
         Else
            bUnload = 1
            If RunningBeta Then
               ' MM TODO: ppiESp01a.Show
            Else
               EstiESp01a.Show
            End If
         End If
         Unload Me
      End If
   Else
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
      MsgBox "Could Not Copy The Requested Estimate.", _
         vbExclamation, Caption
      FillCombo
      GetNextBid Me
      Timer1.Enabled = True
   End If
   
End Sub

Private Sub lblCust_Change()
   If Left(lblCust, 7) = "*** Est" Then lblCust.ForeColor = ES_RED _
           Else lblCust.ForeColor = vbBlack
   
End Sub


Private Sub lblNxt_Change()
   txtBid = lblNxt
   
End Sub

Private Sub optEdit_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub txtBid_LostFocus()
   txtBid = Format(Abs(Val(txtBid)), "000000")
End Sub



Private Function GetNewBid() As Byte
   Dim RdoBid As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT BIDREF FROM EstiTable WHERE BIDREF=" & Val(txtBid) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBid, ES_FORWARD)
   If bSqlRows Then
      With RdoBid
         GetNewBid = 1
         ClearResultSet RdoBid
      End With
   Else
      GetNewBid = 0
   End If
   If GetNewBid = 1 Then MsgBox "That Estimate Number Is In Use." & vbCrLf _
                  & "Try The Next Number In The Sequence.", vbInformation, Caption
   Set RdoBid = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getthebid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiEngr", "Esf05a", optEdit
   
End Sub

Public Sub GetOptions()
   optEdit = GetSetting("Esi2000", "EsiEngr", "Esf05a", optEdit)
   
End Sub
