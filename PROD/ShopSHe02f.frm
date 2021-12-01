VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ShopSHe02f 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manufacturing Order Budgets"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHe02f.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtHrs 
      Height          =   285
      Left            =   2640
      TabIndex        =   4
      Tag             =   "1"
      Top             =   2760
      Width           =   1035
   End
   Begin VB.TextBox txtFoh 
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      Tag             =   "1"
      Top             =   2400
      Width           =   1035
   End
   Begin VB.TextBox txtExp 
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Tag             =   "1"
      Top             =   2040
      Width           =   1035
   End
   Begin VB.TextBox txtMat 
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Tag             =   "1"
      Top             =   1680
      Width           =   1035
   End
   Begin VB.TextBox txtLab 
      Height          =   285
      Left            =   2640
      TabIndex        =   0
      Tag             =   "1"
      Top             =   1320
      Width           =   1035
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4440
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   90
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5160
      Top             =   2520
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3315
      FormDesignWidth =   5370
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Budget Hours"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   14
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Budget Overhead"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   13
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Budget Expense"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label lblMon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label lblRun 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Top             =   720
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO Number"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   9
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Budget Material"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Budget Labor"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   2055
   End
End
Attribute VB_Name = "ShopSHe02f"
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
Dim RdoBud As ADODB.Recordset
Dim bGoodMo As Byte
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      'SelectHelpTopic Me, Caption
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillBudget
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Initialize()
   Move 400, 600
   
End Sub


Private Sub Form_Load()
   SetFormSize Me
   lblMon = ShopSHe02a.cmbPrt
   lblRun = ShopSHe02a.cmbRun
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   ShopSHe02a.optBud.Value = vbUnchecked
   Set RdoBud = Nothing
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set ShopSHe02f = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillBudget()
   On Error GoTo DiaErr1
   bGoodMo = 0
   sSql = "SELECT RUNREF,RUNNO,RUNBUDLAB,RUNBUDMAT," _
          & "RUNBUDEXP,RUNBUDOH,RUNBUDHRS FROM " _
          & "RunsTable WHERE RUNREF='" & Compress(lblMon) & "' " _
          & "AND RUNNO=" & Val(lblRun) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBud, ES_KEYSET)
   If bSqlRows Then
      With RdoBud
         txtLab = Format(!RUNBUDLAB, ES_QuantityDataFormat)
         txtMat = Format(!RUNBUDMAT, ES_QuantityDataFormat)
         txtExp = Format(!RUNBUDEXP, ES_QuantityDataFormat)
         txtFoh = Format(!RUNBUDOH, ES_QuantityDataFormat)
         txtHrs = Format(!RUNBUDHRS, ES_QuantityDataFormat)
         bGoodMo = 1
      End With
   Else
      MsgBox "No Active Manufacturing Order.", _
         vbExclamation, Caption
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillbudget"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtExp_LostFocus()
   txtExp = CheckLen(txtExp, 10)
   txtExp = Format(Abs(Val(txtExp)), ES_QuantityDataFormat)
   If bGoodMo = 1 Then
      On Error Resume Next
      RdoBud!RUNBUDEXP = Val(txtExp)
      RdoBud.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtFoh_LostFocus()
   txtFoh = CheckLen(txtFoh, 10)
   txtFoh = Format(Abs(Val(txtFoh)), ES_QuantityDataFormat)
   If bGoodMo = 1 Then
      On Error Resume Next
      RdoBud!RUNBUDOH = Val(txtFoh)
      RdoBud.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtHrs_LostFocus()
   txtHrs = CheckLen(txtHrs, 10)
   txtHrs = Format(Abs(Val(txtHrs)), ES_QuantityDataFormat)
   If bGoodMo = 1 Then
      On Error Resume Next
      RdoBud!RUNBUDHRS = Val(txtHrs)
      RdoBud.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtLab_LostFocus()
   txtLab = CheckLen(txtLab, 10)
   txtLab = Format(Abs(Val(txtLab)), ES_QuantityDataFormat)
   If bGoodMo = 1 Then
      On Error Resume Next
      RdoBud!RUNBUDLAB = Val(txtLab)
      RdoBud.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtMat_LostFocus()
   txtMat = CheckLen(txtMat, 10)
   txtMat = Format(Abs(Val(txtMat)), ES_QuantityDataFormat)
   If bGoodMo = 1 Then
      On Error Resume Next
      RdoBud!RUNBUDMAT = Val(txtMat)
      RdoBud.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub
