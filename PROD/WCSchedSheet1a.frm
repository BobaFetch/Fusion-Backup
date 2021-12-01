VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form WCSchedSheet1a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Run, Work Center"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtSetup 
      Height          =   285
      Left            =   1320
      TabIndex        =   17
      Tag             =   "1"
      Text            =   "0.0000"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtCmt 
      Height          =   1545
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Tag             =   "9"
      Text            =   "WCSchedSheet1a.frx":0000
      Top             =   2520
      Width           =   4335
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "WCSchedSheet1a.frx":0007
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtPri 
      Height          =   285
      Left            =   5760
      TabIndex        =   5
      Tag             =   "1"
      Text            =   "0"
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   5520
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Update And Apply Changes"
      Top             =   600
      Width           =   875
   End
   Begin VB.ComboBox cmbWcn 
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select Work Center Or leave Blank For All"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   5520
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5880
      Top             =   1920
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4320
      FormDesignWidth =   6630
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Setup Hrs"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   19
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "OP Comments"
      Height          =   285
      Index           =   10
      Left            =   120
      TabIndex        =   18
      Top             =   2400
      Width           =   1080
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Op No"
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   15
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblOpno 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1320
      TabIndex        =   14
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblWcn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4440
      TabIndex        =   12
      Top             =   1560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblQty 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblRun 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   240
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run Qty"
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   9
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lblMon 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   5
      Left            =   4080
      TabIndex        =   7
      Top             =   240
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mo Priority"
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
End
Attribute VB_Name = "WCSchedSheet1a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'4/20/05 Corrected Work Center from calling dialog
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbWcn_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   cmbWcn = CheckLen(cmbWcn, 12)
   For iList = 0 To cmbWcn.ListCount - 1
      If cmbWcn = cmbWcn.List(iList) Then b = 1
   Next
   If b = 0 Then
      Beep
      cmbWcn = lblWcn
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      ' SelectHelpTopic Me, CustSh01.Caption
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdUpd_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim strWCShop As String
   Dim strSetup As String
   Dim strCmt As String
   
   sMsg = "Update the WorkCenter/Priority for the Selected MO?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
   
      strWCShop = GetWCShop(Compress(cmbWcn))
      strSetup = txtSetup.Text
      strCmt = txtCmt.Text
      
      If (strWCShop = "") Then
         MsgBox "Could not find Work Shop for the Selected WC.", vbExclamation, Caption
         Exit Sub
      End If
      
      On Error Resume Next
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      sSql = "UPDATE RnopTable SET OPCENTER='" & Compress(cmbWcn) & "'," _
             & " OPSHOP = '" & strWCShop & "', " _
             & " OPSUHRS = '" & strSetup & "', " _
             & " OPCOMT = '" & strCmt & "' " _
             & " WHERE RUNREF='" & Compress(lblMon) & "' AND " _
             & "RUNNO=" & Val(lblRun) & " "
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
'      sSql = "UPDATE RnopTable SET OPCENTER='" & Compress(cmbWcn) & "'," _
'             & " OPSHOP = '" & strWCShop & "' " _
'             & " WHERE (OPREF='" & Compress(lblMon) & "' AND " _
'             & "OPRUN=" & Val(lblRun) & " AND OPNO=" & Val(lblOpno) & ") "
'      clsADOCon.ExecuteSQL sSql ', rdExecDirect
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         SysMsg "Updated.", True
         
         WCSchedSheet.optFrom = 1
         WCSchedSheet.bOnLoad = 1
         Unload Me
      Else
         clsADOCon.RollbackTrans
         MsgBox "Could Not Complete The Transaction.", _
            vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   
End Sub


Public Function GetWCShop(strWC As String)
   Dim rdoTmp As ADODB.Recordset
   On Error Resume Next
   sSql = "select WCNSHOP from wcntTable where WCNREF = '" & strWC & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoTmp)
   If bSqlRows Then
      With rdoTmp
         GetWCShop = !WCNSHOP
         .Cancel
      End With
      ClearResultSet rdoTmp
   Else
      GetWCShop = ""
   End If
   rdoTmp.Close
   Set rdoTmp = Nothing
   
End Function


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
   FormUnload 1
   Set WCSchedSheet1a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   ES_TimeFormat = GetTimeFormat()
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_FillWorkCentersAll"
   LoadComboBox cmbWcn
   If cmbWcn.ListCount > 0 Then cmbWcn = lblWcn
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub optCom_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub txtPri_LostFocus()
   txtPri = CheckLen(txtPri, 2)
   txtPri = Format(Abs(Val(txtPri)), "#0")
   
End Sub

