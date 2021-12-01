VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form EstiESe01b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estimating Parameters"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdUpdate 
      Cancel          =   -1  'True
      Caption         =   "Update"
      Height          =   435
      Left            =   4920
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   540
      Width           =   875
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "EstiESe01b.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optProfit 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      ToolTipText     =   "Allow Adjusting Estimate Percentages"
      Top             =   3360
      Width           =   715
   End
   Begin VB.CheckBox optGna 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      ToolTipText     =   "Allow Adjusting Estimate Percentages"
      Top             =   3120
      Width           =   715
   End
   Begin VB.CheckBox optScrap 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      ToolTipText     =   "Allow Adjusting Estimate Percentages"
      Top             =   2880
      Width           =   715
   End
   Begin VB.TextBox txtEst 
      Height          =   285
      Left            =   3120
      TabIndex        =   10
      Tag             =   "2"
      ToolTipText     =   "Estimators Name - Stored In This Computers Registry (30 Char Max)"
      Top             =   3960
      Width           =   2595
   End
   Begin VB.TextBox txtScr 
      Height          =   285
      Left            =   3120
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Percentage Of Material And Labor"
      Top             =   1800
      Width           =   915
   End
   Begin VB.TextBox txtRte 
      Height          =   285
      Left            =   3120
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Labor Rate In Dollars"
      Top             =   1080
      Width           =   915
   End
   Begin VB.CheckBox optWcn 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      ToolTipText     =   "Work Center Percentages"
      Top             =   3600
      Width           =   715
   End
   Begin VB.TextBox txtPrf 
      Height          =   285
      Left            =   3120
      TabIndex        =   5
      Tag             =   "3"
      ToolTipText     =   "Percentage Of Total"
      Top             =   2520
      Width           =   915
   End
   Begin VB.TextBox txtGna 
      Height          =   285
      Left            =   3120
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Percent"
      Top             =   2160
      Width           =   915
   End
   Begin VB.TextBox txtFoh 
      Height          =   285
      Left            =   3120
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Percentage Of Labor Rate"
      Top             =   1440
      Width           =   915
   End
   Begin VB.TextBox txtMtl 
      Height          =   285
      Left            =   3120
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Percentage of Material"
      Top             =   720
      Width           =   915
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   4920
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5400
      Top             =   4440
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4560
      FormDesignWidth =   5910
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Allow Over Writing Profit"
      Height          =   252
      Index           =   15
      Left            =   240
      TabIndex        =   27
      ToolTipText     =   "Allow Adjusting Estimate Percentages"
      Top             =   3360
      Width           =   2892
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Allow Over Writing General/Admin"
      Height          =   252
      Index           =   14
      Left            =   240
      TabIndex        =   26
      ToolTipText     =   "Allow Adjusting Estimate Percentages"
      Top             =   3120
      Width           =   2892
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Allow Over Writing Scrap"
      Height          =   252
      Index           =   4
      Left            =   240
      TabIndex        =   25
      ToolTipText     =   "Allow Adjusting Estimate Percentages"
      Top             =   2880
      Width           =   2892
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimator (Stored Locally)"
      Height          =   252
      Index           =   13
      Left            =   240
      TabIndex        =   24
      ToolTipText     =   "Profit Calculated On The Subtotal Of All (Including G&A)"
      Top             =   3960
      Width           =   2412
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reduction (Scrap)"
      Height          =   252
      Index           =   12
      Left            =   240
      TabIndex        =   23
      ToolTipText     =   "Percentage Calculated On The Subtotal Of All"
      Top             =   1800
      Width           =   2892
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   252
      Index           =   11
      Left            =   4080
      TabIndex        =   22
      Top             =   1800
      Width           =   492
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Default Labor Rate"
      Height          =   252
      Index           =   10
      Left            =   240
      TabIndex        =   21
      ToolTipText     =   "Percentage Of Labor Used"
      Top             =   1080
      Width           =   2892
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   252
      Index           =   9
      Left            =   4080
      TabIndex        =   20
      Top             =   720
      Width           =   492
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Use Overhead From Work Centers"
      Height          =   252
      Index           =   8
      Left            =   240
      TabIndex        =   19
      ToolTipText     =   "Work Center Percentages"
      Top             =   3600
      Width           =   2652
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   252
      Index           =   7
      Left            =   4080
      TabIndex        =   18
      Top             =   2520
      Width           =   492
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   252
      Index           =   6
      Left            =   4080
      TabIndex        =   17
      Top             =   2160
      Width           =   492
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   252
      Index           =   5
      Left            =   4080
      TabIndex        =   16
      Top             =   1440
      Width           =   492
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Profit Of Sale"
      Height          =   252
      Index           =   3
      Left            =   240
      TabIndex        =   15
      ToolTipText     =   "Profit Calculated On The Subtotal Of All (Including G&A)"
      Top             =   2520
      Width           =   2892
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "General And Admin Expense"
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   14
      ToolTipText     =   "Percentage Calculated On The Subtotal Of All"
      Top             =   2160
      Width           =   2892
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Default Factory Overhead Expense"
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   13
      ToolTipText     =   "Percentage Of Labor Used"
      Top             =   1440
      Width           =   2892
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Material Burden Expense"
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   12
      ToolTipText     =   "Percentage Of Material Used"
      Top             =   720
      Width           =   2892
   End
End
Attribute VB_Name = "EstiESe01b"
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

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmdCan_Click()
   'UpdateRows
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 1103
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdUpdate_Click()
   UpdateRows
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      GetParameters
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
   Set EstiESe01b = Nothing
   
End Sub



Private Sub FormatControls()
   On Error Resume Next
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtRte = "0.000"
   txtFoh = "0.000"
   txtGna = "0.000"
   txtPrf = "0.000"
   txtMtl = "0.000"
   txtScr = "0.000"
   
End Sub


Private Sub GetParameters()
   Dim RdoPar As ADODB.Recordset
   'On Error Resume Next
   sSql = "select * from preferences"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPar, ES_FORWARD)
   If bSqlRows Then
      With RdoPar
         txtMtl = Format(!EstMatlBurden, ES_QuantityDataFormat)
         txtFoh = Format(!EstFactoryOverHead, ES_QuantityDataFormat)
         txtRte = Format(!EstLaborRate, ES_QuantityDataFormat)
         txtGna = Format(!EstGenAdmnExp, ES_QuantityDataFormat)
         txtPrf = Format(!EstProfitOfSale, ES_QuantityDataFormat)
         txtScr = Format(!EstScrapRate, ES_QuantityDataFormat)
         optWcn.value = Format(!EstUseWCOverhead, "0")
         optScrap.value = Format(!ESTOVERWRITESCRAP, "0")
         optGna.value = Format(!ESTOVERWRITEGNA, "0")
         optProfit.value = Format(!ESTOVERWRITEPROFIT, "0")
         ClearResultSet RdoPar
      End With
   End If
   txtEst = GetSetting("Esi2000", "EsiEngr", "Estimator", txtEst)
   sCurrEstimator = GetSetting("Esi2000", "EsiEngr", "Estimator", sCurrEstimator)
   Set RdoPar = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getparameters"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub optGna_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optProfit_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optScrap_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optWcn_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub txtEst_LostFocus()
   txtEst = CheckLen(txtEst, 30)
   txtEst = StrCase(txtEst)
   SaveSetting "Esi2000", "EsiEngr", "Estimator", txtEst
   sCurrEstimator = txtEst
   
End Sub


Private Sub txtFoh_LostFocus()
   txtFoh = CheckLen(txtFoh, 7)
   txtFoh = Format(Abs(Val(txtFoh)), ES_QuantityDataFormat)
   
End Sub


Private Sub txtGna_LostFocus()
   txtGna = CheckLen(txtGna, 7)
   txtGna = Format(Abs(Val(txtGna)), ES_QuantityDataFormat)
   
End Sub


Private Sub txtMtl_LostFocus()
   txtMtl = CheckLen(txtMtl, 7)
   txtMtl = Format(Abs(Val(txtMtl)), ES_QuantityDataFormat)
   
   
End Sub


Private Sub txtPrf_LostFocus()
   txtPrf = CheckLen(txtPrf, 7)
   txtPrf = Format(Abs(Val(txtPrf)), ES_QuantityDataFormat)
   
End Sub



Private Sub UpdateRows()
   'On Error Resume Next
   sSql = "UPDATE Preferences SET " _
          & "EstMatlBurden=" & Val(txtMtl) & ", " & vbCrLf _
          & "EstFactoryOverHead=" & Val(txtFoh) & "," & vbCrLf _
          & "EstGenAdmnExp=" & Val(txtGna) & "," & vbCrLf _
          & "EstProfitOfSale=" & Val(txtPrf) & "," & vbCrLf _
          & "EstUseWCOverhead=" & optWcn.value & "," & vbCrLf _
          & "EstScrapRate=" & Val(txtScr) & "," & vbCrLf _
          & "EstLaborRate=" & Val(txtRte) & "," & vbCrLf _
          & "ESTOVERWRITESCRAP=" & optScrap.value & "," & vbCrLf _
          & "ESTOVERWRITEGNA=" & optGna.value & "," & vbCrLf _
          & "ESTOVERWRITEPROFIT=" & optProfit.value & " " & vbCrLf _
          & "WHERE PreRecord=1"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   MsgBox "Estimating parameters updated"
   
End Sub

Private Sub txtRte_LostFocus()
   txtRte = CheckLen(txtRte, 7)
   txtRte = Format(Abs(Val(txtRte)), ES_QuantityDataFormat)
   
End Sub


Private Sub txtScr_LostFocus()
   txtScr = CheckLen(txtScr, 7)
   txtScr = Format(Abs(Val(txtScr)), ES_QuantityDataFormat)
   If Val(txtScr) > 9.5 Then
      MsgBox txtScr & " Seems High. May Wish To Check.", _
         vbInformation, Caption
   End If
   
End Sub
