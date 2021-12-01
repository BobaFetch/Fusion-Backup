VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form EstiESf02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancel An Estimate"
   ClientHeight    =   2700
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
   ScaleHeight     =   2700
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "EstiESf02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtdmy 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1840
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   75
   End
   Begin VB.CommandButton cmdCbd 
      Caption         =   "C&ancel"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4800
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Cancel The Current Bid"
      Top             =   600
      Width           =   875
   End
   Begin VB.ComboBox cmbBid 
      Height          =   315
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
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5160
      Top             =   2040
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2700
      FormDesignWidth =   5730
   End
   Begin VB.Label lblPrt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   16
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   15
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblDate 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3960
      TabIndex        =   14
      Top             =   1080
      Width           =   950
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   13
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label lblRfq 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblCust 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Top             =   1410
      Width           =   3375
   End
   Begin VB.Label lblNik 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "RFQ"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   8
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblTyp 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3960
      TabIndex        =   6
      Top             =   600
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   5
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblCls 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      Top             =   600
      Width           =   255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimate Number"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "EstiESf02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodBid As Byte

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
   
   sMsg = "This Function Permanently Cancels A Bid." & vbCrLf _
          & "Are You Sure That You Wish Cancel The Bid?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      clsADOCon.ADOErrNum = 0
      sSql = "UPDATE EstiTable SET BIDCANCELED=1,BIDDATECANCELED='" _
             & Format(ES_SYSDATE, "mm/dd/yy") & "' WHERE BIDREF=" & Val(cmbBid) & " "
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      If clsADOCon.ADOErrNum = 0 Then
         MsgBox "Bid Was Successfully Marked Canceled.", _
            vbInformation, Caption
         FillCombo
      Else
         MsgBox "Bid Not Was Successfully Marked Canceled.", _
            vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3551
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
   FormLoad Me, ES_DONTLIST
   FormatControls
   
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set EstiESf02a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtdmy.BackColor = Es_FormBackColor
   
End Sub

Private Sub FillCombo()
   Dim RdoCmb As ADODB.Recordset
   Dim iList As Integer
   On Error GoTo DiaErr1
   cmbBid.Clear
   sSql = "SELECT BIDREF,BIDPRE,BIDCANCELED,BIDACCEPTED FROM EstiTable " _
          & "WHERE BIDCANCELED=0 AND BIDACCEPTED=0 ORDER BY BIDREF DESC"
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
          & "AND BIDCANCELED=0 AND BIDACCEPTED=0"
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
      lblCust = ""
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
