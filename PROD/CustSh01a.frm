VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form CustSh01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Run, Work Center"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CustSh01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdRel 
      Cancel          =   -1  'True
      Caption         =   "&Release"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2160
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Mark MO As RL (Released)"
      Top             =   3000
      Width           =   875
   End
   Begin VB.TextBox txtUnt 
      Height          =   285
      Left            =   4440
      TabIndex        =   10
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   2280
      Width           =   825
   End
   Begin VB.TextBox txtSet 
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   2280
      Width           =   825
   End
   Begin VB.CheckBox optCom 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   2640
      Width           =   715
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Tag             =   "1"
      Text            =   "0"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtPri 
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Tag             =   "1"
      Text            =   "0"
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   5280
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Update And Apply Changes"
      Top             =   600
      Width           =   875
   End
   Begin VB.ComboBox cmbWcn 
      Height          =   315
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select Work Center Or leave Blank For All"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
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
      FormDesignHeight=   3660
      FormDesignWidth =   6255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Release MO"
      Enabled         =   0   'False
      Height          =   255
      Index           =   10
      Left            =   360
      TabIndex        =   25
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label lblSta 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4440
      TabIndex        =   24
      Top             =   600
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Time"
      Height          =   255
      Index           =   9
      Left            =   3360
      TabIndex        =   23
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Setup Time"
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   22
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblWcn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4440
      TabIndex        =   21
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblQty 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3120
      TabIndex        =   20
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblRun 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4440
      TabIndex        =   19
      Top             =   240
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run Qty"
      Height          =   255
      Index           =   7
      Left            =   2160
      TabIndex        =   18
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblOpno 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1080
      TabIndex        =   17
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblMon 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1080
      TabIndex        =   16
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Op No"
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   15
      Top             =   600
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   5
      Left            =   3840
      TabIndex        =   14
      Top             =   240
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   13
      Top             =   240
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operation Complete"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   6
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Completed Qty"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mo Priority"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "CustSh01a"
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



Private Sub cmdRel_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   sMsg = "Mark This MO As RL (Released To Production)?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      clsADOCon.ADOErrNum = 0
      
      sSql = "UPDATE RunsTable SET RUNSTATUS='RL',RUNRELEASED=1 " _
             & "WHERE RUNREF='" & Compress(lblMon) & "' " _
             & "AND RUNNO=" & Val(lblRun) & " "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then
         CustSh01.optData.Value = vbChecked
         lblSta = "RL"
         MsgBox "Run Status Updated.", vbInformation, Caption
      Else
         MsgBox "Run Status Could Not Be Updated.", vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   
End Sub

Private Sub cmdUpd_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   sMsg = "Update The Manufacturing Order And Operation?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      
      sSql = "UPDATE RunsTable SET RUNPRIORITY=" & Val(txtPri) & " " _
             & "WHERE RUNREF='" & Compress(lblMon) & "' AND " _
             & "RUNNO=" & Val(lblRun) & " "
      clsADOCon.ExecuteSQL sSql
      
      sSql = "UPDATE RnopTable SET OPCOMPLETE=" & Val(optCom) & "," _
             & "OPCOMPDATE='" & Format(ES_SYSDATE, "mm/dd/yy") & "'," _
             & "OPYIELD=" & Val(txtQty) & ",OPCENTER='" & Compress(cmbWcn) & "'," _
             & "OPSUHRS=" & Val(txtSet) & ",OPUNITHRS=" & txtUnt & " " _
             & "WHERE (OPREF='" & Compress(lblMon) & "' AND " _
             & "OPRUN=" & Val(lblRun) & " AND OPNO=" & Val(lblOpno) & ") "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then
         CustSh01.optData.Value = vbChecked
         clsADOCon.CommitTrans
         SysMsg "Updated.", True
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


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      GetHours
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   Move CustSh01.Left + 800, CustSh01.Top + 1000
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   CustSh01.optfrom = vbUnchecked
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload 1
   Set CustSh01a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   ES_TimeFormat = GetTimeFormat()
   txtQty = "0"
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_FillWorkCentersAll"
   LoadComboBox cmbWcn
   'If cmbWcn.ListCount > 0 Then cmbWcn = cmbWcn.List(0)
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


Private Sub txtQty_LostFocus()
   txtQty = Format(Abs(Val(txtQty)), "####0")
   
End Sub


Private Sub txtSet_LostFocus()
   txtSet = CheckLen(txtSet, 7)
   txtSet = Format(Abs(Val(txtSet)), "##0.000")
   
End Sub


Private Sub txtUnt_LostFocus()
   txtUnt = CheckLen(txtUnt, 8)
   txtUnt = Format(Abs(Val(txtUnt)), ES_TimeFormat)
   
End Sub



Private Sub GetHours()
   Dim RdoHrs As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT OPREF,OPRUN,OPNO,OPSUHRS,OPUNITHRS FROM " _
          & "RnopTable WHERE (OPREF='" & Compress(lblMon) & "' AND " _
          & "OPRUN=" & Val(lblRun) & "AND OPNO=" & Val(lblOpno) & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoHrs, ES_FORWARD)
   If bSqlRows Then
      With RdoHrs
         txtSet = Format(!OPSUHRS, "#0.000")
         txtUnt = Format(!OPUNITHRS, ES_TimeFormat)
         If lblSta = "SC" Then
            z1(10).Enabled = True
            cmdRel.Enabled = True
         Else
            z1(10).Enabled = False
            cmdRel.Enabled = False
         End If
         ClearResultSet RdoHrs
      End With
   Else
      txtSet = "0.000"
      txtUnt = ES_TimeFormat
   End If
   Set RdoHrs = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "gethours"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
