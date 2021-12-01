VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ppiESe02e 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Other Estimate Charges"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ppiESe02e.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtPackage 
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "Unit Costs"
      Top             =   1560
      Width           =   792
   End
   Begin VB.TextBox txtMask 
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "Unit Costs"
      Top             =   1200
      Width           =   792
   End
   Begin VB.TextBox txtHoleCost 
      Height          =   285
      Left            =   3480
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Unit Costs"
      Top             =   840
      Width           =   552
   End
   Begin VB.TextBox txtHoles 
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Unit Costs"
      Top             =   840
      Width           =   552
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "ppiESe02e.frx":07AE
      DownPicture     =   "ppiESe02e.frx":1120
      Height          =   350
      Left            =   1920
      Picture         =   "ppiESe02e.frx":1A92
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Standard Comments"
      Top             =   3120
      Width           =   350
   End
   Begin VB.TextBox txtSet 
      Height          =   285
      Left            =   4320
      TabIndex        =   9
      Tag             =   "1"
      Top             =   4080
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox txtEpd 
      Height          =   285
      Left            =   4320
      TabIndex        =   10
      Tag             =   "1"
      Top             =   4440
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox txtTlg 
      Height          =   285
      Left            =   4320
      TabIndex        =   11
      Tag             =   "1"
      Top             =   4800
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox txtCmt 
      Height          =   765
      Left            =   2400
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Tag             =   "2"
      Top             =   3120
      Width           =   3795
   End
   Begin VB.ComboBox txtFst 
      Height          =   288
      Left            =   2400
      TabIndex        =   4
      Tag             =   "4"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtEst 
      Height          =   285
      Left            =   2400
      TabIndex        =   7
      Tag             =   "2"
      ToolTipText     =   "Profit Calculated On The Subtotal Of All (Including G&A)"
      Top             =   2760
      Width           =   2595
   End
   Begin VB.TextBox txtByr 
      Height          =   285
      Left            =   2400
      TabIndex        =   6
      Tag             =   "2"
      ToolTipText     =   "Profit Calculated On The Subtotal Of All (Including G&A)"
      Top             =   2412
      Width           =   2595
   End
   Begin VB.ComboBox txtDue 
      Height          =   288
      Left            =   4680
      TabIndex        =   5
      Tag             =   "4"
      ToolTipText     =   "Due Date"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5400
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   1200
      Top             =   3960
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4200
      FormDesignWidth =   6345
   End
   Begin VB.Label lblHoles 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   4440
      TabIndex        =   29
      Top             =   840
      Width           =   912
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "="
      Height          =   252
      Index           =   4
      Left            =   4080
      TabIndex        =   28
      Top             =   840
      Width           =   252
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "@"
      Height          =   252
      Index           =   3
      Left            =   3000
      TabIndex        =   27
      Top             =   840
      Width           =   492
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Packaging"
      Height          =   252
      Index           =   2
      Left            =   600
      TabIndex        =   26
      Top             =   1560
      Width           =   1572
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Masking"
      Height          =   252
      Index           =   1
      Left            =   600
      TabIndex        =   25
      Top             =   1200
      Width           =   1572
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Blind Holes"
      Height          =   252
      Index           =   0
      Left            =   600
      TabIndex        =   24
      Top             =   840
      Width           =   1572
   End
   Begin VB.Label lblBid 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   240
      TabIndex        =   22
      ToolTipText     =   "Total Services"
      Top             =   3720
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tooling Charges"
      Height          =   252
      Index           =   19
      Left            =   2400
      TabIndex        =   21
      Top             =   4800
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Expediting Charges"
      Height          =   252
      Index           =   20
      Left            =   2400
      TabIndex        =   20
      Top             =   4440
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extra Setup Charges"
      Height          =   252
      Index           =   21
      Left            =   2400
      TabIndex        =   19
      Top             =   4080
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Bid Costs (Unit Costs):"
      Height          =   252
      Index           =   22
      Left            =   240
      TabIndex        =   18
      Top             =   480
      Width           =   4332
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "First Delivery Date"
      Height          =   252
      Index           =   23
      Left            =   240
      TabIndex        =   17
      Top             =   2040
      Width           =   1572
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments:"
      Height          =   255
      Index           =   24
      Left            =   240
      TabIndex        =   16
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimator"
      Height          =   252
      Index           =   32
      Left            =   240
      TabIndex        =   15
      Top             =   2760
      Width           =   1572
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Buyer"
      Height          =   252
      Index           =   33
      Left            =   240
      TabIndex        =   14
      Top             =   2412
      Width           =   1572
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bid Is Due"
      Height          =   252
      Index           =   34
      Left            =   3600
      TabIndex        =   13
      Top             =   2040
      Width           =   1572
   End
End
Attribute VB_Name = "ppiESe02e"
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
Dim bGoodBid As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdComments_Click()
   If cmdComments Then
      'See List For Index
      txtCmt.SetFocus
      SysComments.lblListIndex = 7
      SysComments.Show
      cmdComments = False
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 8512
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      cmdComments.Enabled = True
      Caption = Caption & " - Estimate " & lblBid
      
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Deactivate()
   Unload Me
   
End Sub

Private Sub Form_Load()
   Move 1000, 1000
   FormatControls
   GetOptions
   lblBid = ppiESe02a.cmbBid
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   ppiESe02a.lblOther = Format(Val(lblHoles) + Val(txtMask) + Val(txtPackage), "#####0.00")
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set ppiESe02e = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   If Trim(txtFst) = "" Then txtFst = Format(Now + 60, "mm/dd/yy")
   txtHoles = "0"
   txtMask = "0.00"
   txtPackage = "0.00"
   lblHoles = "0.00"
   
End Sub


Private Sub lblBid_Change()
   If Val(lblBid) > 0 Then bGoodBid = 1
   
End Sub


Private Sub lblHoles_Change()
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDHOLESTOTAL = Val(lblHoles)
      RdoFull.Update
   End If
   
End Sub

Private Sub txtByr_LostFocus()
   txtByr = CheckLen(txtByr, 30)
   txtByr = StrCase(txtByr)
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDBUYER = Trim(txtByr)
      RdoFull.Update
   End If
   
End Sub


Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 1024)
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDCOMMENT = Trim(txtCmt)
      RdoFull.Update
   End If
   
End Sub


Private Sub txtDue_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtDue_LostFocus()
   txtDue = CheckDate(txtDue)
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDDUE = Format(txtDue, "mm/dd/yy")
      RdoFull.Update
   End If
   
End Sub


Private Sub txtEpd_LostFocus()
   txtEpd = CheckLen(txtEpd, 9)
   txtEpd = Format(Abs(Val(txtEpd)), "####0.00")
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDEXPEDITE = Val(txtEpd)
      RdoFull.Update
   End If
   
End Sub


Private Sub txtEst_LostFocus()
   txtEst = CheckLen(txtEst, 30)
   txtEst = StrCase(txtEst)
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDESTIMATOR = Trim(txtEst)
      RdoFull.Update
   End If
   If Trim(txtEst) <> "" Then
      SaveSetting "Esi2000", "EsiEngr", "Estimator", txtEst
      sCurrEstimator = txtEst
   End If
   ppiESe02a.lblEstimator = txtEst
   ppiESe02a.Estimator = sCurrEstimator
   
End Sub


Private Sub txtFst_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub txtFst_LostFocus()
   If Trim(txtFst) = "" Then
      If bGoodBid Then
         On Error Resume Next
         'RdoFull.Edit
         RdoFull!BIDFIRSTDELIVERY = Null
         RdoFull.Update
      End If
   Else
      txtFst = CheckDate(txtFst)
      If bGoodBid Then
         On Error Resume Next
         'RdoFull.Edit
         RdoFull!BIDFIRSTDELIVERY = Format(txtFst, "mm/dd/yy")
         RdoFull.Update
      End If
   End If
   
End Sub


Private Sub txtHoleCost_LostFocus()
   txtHoleCost = Format(Abs(Val(txtHoleCost)), "##0.00")
   lblHoles = Format(Val(txtHoles) * Val(txtHoleCost), "####0.00")
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDHOLESCOST = Val(txtHoleCost)
      RdoFull.Update
   End If
   
End Sub


Private Sub txtHoles_LostFocus()
   txtHoles = Format(Abs(Val(txtHoles)), "##0")
   lblHoles = Format(Val(txtHoles) * Val(txtHoleCost), "####0.00")
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDHOLES = Val(txtHoles)
      RdoFull.Update
   End If
   
End Sub


Private Sub txtMask_LostFocus()
   txtMask = Format(Abs(Val(txtMask)), "#####0.00")
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDMASKING = Val(txtMask)
      RdoFull.Update
   End If
   
End Sub


Private Sub txtPackage_LostFocus()
   txtPackage = Format(Abs(Val(txtPackage)), "#####0.00")
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDPACKAGING = Val(txtPackage)
      RdoFull.Update
   End If
   
End Sub


Private Sub txtSet_LostFocus()
   txtSet = CheckLen(txtSet, 9)
   txtSet = Format(Abs(Val(txtSet)), "####0.00")
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDSETUP = Val(txtSet)
      RdoFull.Update
   End If
   
End Sub


Private Sub txtTlg_LostFocus()
   txtTlg = CheckLen(txtTlg, 7)
   txtTlg = Format(Abs(Val(txtTlg)), "####0.00")
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDTOOLING = Val(txtTlg)
      RdoFull.Update
   End If
   
End Sub



Private Sub GetOptions()
   Dim sOptions As String
   sOptions = GetSetting("Esi2000", "EsiEngr", "ppiESe02e", sOptions)
   If Len(sOptions) > 0 Then
      txtHoleCost = sOptions
   Else
      txtHoleCost = "0.15"
   End If
   sOptions = GetSetting("Esi2000", "EsiEngr", "ppiESe02e2", sOptions)
   If Len(sOptions) > 0 Then
      txtEst = sOptions
   End If
   
End Sub

Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiEngr", "ppiESe02e", txtHoleCost
   SaveSetting "Esi2000", "EsiEngr", "ppiESe02e2", txtEst
   
End Sub
