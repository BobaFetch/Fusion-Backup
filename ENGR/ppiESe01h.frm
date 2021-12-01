VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ppiESe01h 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Qwik Bid"
   ClientHeight    =   2505
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   5685
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   4680
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "Unit Costs"
      Top             =   960
      Width           =   915
   End
   Begin VB.ComboBox cmbCls 
      Height          =   288
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Class A-Z"
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   315
      Left            =   4680
      TabIndex        =   4
      ToolTipText     =   "Add This New Estimate"
      Top             =   480
      Width           =   915
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4680
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin VB.ComboBox cmbEst 
      Height          =   288
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   3
      Tag             =   "2"
      Top             =   1680
      Width           =   2772
   End
   Begin VB.ComboBox cmbCst 
      Height          =   288
      Left            =   1680
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select A Customer"
      Top             =   960
      Width           =   1555
   End
   Begin VB.TextBox txtBid 
      Height          =   288
      Left            =   2160
      TabIndex        =   0
      Tag             =   "1"
      Top             =   600
      Width           =   852
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   4920
      Top             =   2040
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5280
      Top             =   2040
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2505
      FormDesignWidth =   5685
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      Height          =   252
      Index           =   8
      Left            =   3600
      TabIndex        =   13
      Top             =   960
      Width           =   972
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimator"
      Height          =   252
      Index           =   32
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Width           =   1572
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   1332
   End
   Begin VB.Label lblNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1680
      TabIndex        =   8
      Top             =   1320
      Width           =   3920
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimate Number"
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   1572
   End
   Begin VB.Label lblNxt 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1680
      TabIndex        =   6
      Top             =   120
      Width           =   852
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Next Estimate"
      Height          =   252
      Index           =   31
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   1452
   End
End
Attribute VB_Name = "ppiESe01h"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bOnLoad As Byte
Dim cGna As Currency
Dim cProfit As Currency
Dim cOldQty As Currency
Dim cScrap As Currency

'Bid stuff
Dim cBFoh As Currency
Dim cBGna As Currency
Dim cBprofit As Currency
Dim cBScrap As Currency

Dim cBGnaRate As Currency
Dim cBProfitRate As Currency
Dim cBScrapRate As Currency

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub AddBid()
   Dim bResponse As Byte
   Dim lBidNo As Long
   Dim sMsg As String
   Dim vDate As String
   
   On Error GoTo DiaErr1
   Timer1.Enabled = False
   GetNextBid Me
   vDate = Format(ES_SYSDATE, "mm/dd/yy")
   lBidNo = Val(lblNxt)
   sMsg = "Add New Estimate " & lblNxt & "?... "
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      'Add It Here
      lBidNo = Val(lblNxt)
      GetRates
      sSql = "INSERT INTO EstiTable(BIDREF,BIDNUM,BIDPRE," _
             & "BIDCUST,BIDCLASS,BIDDATE,BIDGNARATE,BIDPROFITRATE," _
             & "BIDSCRAPRATE,BIDQUANTITY,BIDESTIMATOR) VALUES(" _
             & lBidNo & ",'" _
             & Format$(lBidNo, "000000") & "','" _
             & cmbCls & "','" _
             & Compress(cmbCst) & "'," _
             & "'QWIK','" _
             & vDate & "'," _
             & cGna & "," _
             & cProfit & "," _
             & cScrap & "," _
             & Val(txtQty) & ",'" _
             & cmbEst & "')"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      If Err = 0 Then
         SaveSetting "Esi2000", "EsiEngr", "Estimator", cmbEst
         SysMsg "Estimate " & txtBid & " Was Added.", True
         On Error Resume Next
         ppiESe01a.cmbBid.AddItem txtBid
         ppiESe01a.cmbBid = txtBid
         bGoodBid = ppiESe01a.GetThisPPIQBid(True)
         ppiESe01a.txtPrt.SetFocus
         Form_Deactivate
      Else
         MsgBox "Could Not Add Estimate " & txtBid & ".", _
            vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   On Error Resume Next
   MsgBox "Estimate " & lBidNo & " Has been Recorded. Try Again.", _
      vbInformation, Caption
   Timer1.Enabled = True
   GetNextBid Me
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   GetNextBid Me
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   For b = 65 To 88
      cmbCls.AddItem Chr$(b)
   Next
   cmbCls.AddItem Chr$(b)
   cmbCls = ppiESe01a.cmbCls
   txtQty = Format(1, "####0.00")
   
End Sub


Private Sub cmbCst_Click()
   FindCustomer Me, cmbCst
   
End Sub


Private Sub cmbCst_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   For iList = 0 To cmbCst.ListCount - 1
      If cmbCst.List(iList) = cmbCst Then
         b = 1
         Exit For
      End If
   Next
   If b = 0 Then
      Beep
      cmbCst = cmbCst.List(0)
   End If
   
   
End Sub


Private Sub cmbEst_LostFocus()
   cmbEst = CheckLen(cmbEst, 30)
   cmbEst = StrCase(cmbEst)
   
End Sub


Private Sub cmdAdd_Click()
   Dim bByte As Byte
   If Trim(cmbEst) = "" Then
      MsgBox "Requires The Estimator's Name.", _
         vbInformation, Caption
      Exit Sub
   End If
   If Trim(cmbCst) = "" Then
      MsgBox "Requires A Valid Customer Name.", _
         vbInformation, Caption
      Exit Sub
   End If
   bByte = GetThisBidNumber()
   If bByte = 1 Then
      MsgBox "That Estimate Number Has Been Used. Try Again.", _
         vbInformation, Caption
      GetNextBid Me
      txtBid = lblNxt
   Else
      AddBid
   End If
   
End Sub

Private Sub cmdCan_Click()
   Form_Deactivate
   
   
End Sub

Private Sub Form_Activate()
   If bOnLoad Then
      FillCustomers
      cmbCst = ppiESe01a.cmbCst
      cmbCls = ppiESe01a.cmbCls
      FindCustomer Me, cmbCst
      FillEstimators
   End If
   bOnLoad = 0
   
End Sub

Private Sub GetRates()
   Dim RdoPar As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_EstParameters"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPar, ES_FORWARD)
   If bSqlRows Then
      With RdoPar
         cGna = Format(!EstGenAdmnExp / 100, ES_QuantityDataFormat)
         cProfit = Format(!EstProfitOfSale / 100, ES_QuantityDataFormat)
         cScrap = Format(!EstScrapRate / 100, ES_QuantityDataFormat)
         'Defaults
         cBGnaRate = cGna
         cBProfitRate = cProfit
         cBScrapRate = cScrap
         ClearResultSet RdoPar
      End With
   End If
   Set RdoPar = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getrates"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Deactivate()
   Unload Me
   
End Sub


Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   Move 2000, (Screen.Height - (Me.Height + 3000)) / 2
   FormatControls
   bOnLoad = 1
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set ppiESe01h = Nothing
   
End Sub


Private Sub lblNxt_Change()
   txtBid = lblNxt
   
End Sub

Private Sub Timer1_Timer()
   GetNextBid Me
   
End Sub



Private Sub FillEstimators()
   sSql = "SELECT DISTINCT BIDESTIMATOR from EstiTable WHERE BIDESTIMATOR<>''"
   LoadComboBox cmbEst, -1
   cmbEst = GetSetting("Esi2000", "EsiEngr", "Estimator", cmbEst)
   
End Sub

Public Function GetThisBidNumber() As Byte
   Dim RdoBid As ADODB.Recordset
   sSql = "SELECT BIDREF From EstiTable WHERE BIDREF=" & Val(txtBid) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBid, ES_FORWARD)
   If bSqlRows Then GetThisBidNumber = 1 Else GetThisBidNumber = 0
   
End Function

Private Sub txtBid_LostFocus()
   txtBid = Format(Abs(Val(txtBid)), "000000")
   If Val(txtBid) = 0 Then
      Beep
      txtBid = lblNxt
   End If
   
End Sub


Private Sub txtQty_LostFocus()
   txtQty = Format(Abs(Val(txtQty)), "####0.00")
   
End Sub
