VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form SaleSLf10a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lock/Unlock Sales Orders"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLf10a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optStatus 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   252
      Left            =   2830
      TabIndex        =   10
      Top             =   1200
      Width           =   252
   End
   Begin VB.ComboBox cmbSon 
      Height          =   288
      Left            =   1560
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Select or Enter Sales Order Number (Contains 300 Max)"
      Top             =   720
      Width           =   975
   End
   Begin VB.CheckBox optLock 
      Caption         =   "Show Locked"
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      Top             =   720
      Width           =   1932
   End
   Begin VB.CommandButton cmdLock 
      Cancel          =   -1  'True
      Caption         =   "&Lock"
      Height          =   315
      Left            =   4920
      TabIndex        =   9
      ToolTipText     =   "Lock Or Unlock The Sales Order Depending On The Current Statis"
      Top             =   600
      Width           =   875
   End
   Begin VB.TextBox txtCst 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtNme 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1560
      Width           =   3855
   End
   Begin VB.TextBox txtSon 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Tag             =   "1"
      Text            =   "00000"
      ToolTipText     =   "Sales Order Number (No Class)"
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   4920
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   4200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2550
      FormDesignWidth =   5850
   End
   Begin VB.Label lblStatus 
      Height          =   252
      Left            =   3120
      TabIndex        =   11
      Top             =   1200
      Width           =   1812
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "Sales Order Number (No Class)"
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order "
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "Sales Order Number (No Class)"
      Top             =   720
      Width           =   1065
   End
   Begin VB.Label lblSoType 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   720
      Width           =   255
   End
End
Attribute VB_Name = "SaleSLf10a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'8/11/04 new
'10/17/05 Added Combo and switch
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FillSalesOrds()
   cmbSon.Clear
   If optLock.Value = vbChecked Then
      optLock.Caption = "Show Unlocked"
   Else
      optLock.Caption = "Show Unlocked"
   End If
   sSql = "SELECT SONUMBER FROM SohdTable WHERE SOLOCKED=" & optLock.Value & " " _
          & "ORDER BY SONUMBER DESC"
   LoadNumComboBox cmbSon, SO_NUM_FORMAT
   If cmbSon.ListCount > 0 Then cmbSon = cmbSon.List(0)
   txtSon = cmbSon
   lblSoType = GetSalesOrderType()
   
End Sub

Private Function GetSalesOrderType()
   Dim RdoTyp As ADODB.Recordset
   'On Error GoTo DiaErr1
   sSql = "SELECT SONUMBER,SOTYPE,SOCUST,SOLOCKED,CUREF," _
          & "CUNICKNAME,CUNAME FROM SohdTable,CustTable " _
          & "WHERE (SOCUST=CUREF) AND SONUMBER=" & Val(txtSon) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTyp, ES_FORWARD)
   If bSqlRows Then
      With RdoTyp
         GetSalesOrderType = "" & Trim(!SOTYPE)
         txtCst = "" & Trim(!CUNICKNAME)
         txtNme = "" & Trim(!CUNAME)
         optStatus.Value = !SOLOCKED
         ClearResultSet RdoTyp
      End With
   Else
      GetSalesOrderType = ""
      txtCst = ""
      txtNme = "*** Sales Order Wasn't Found ****"
      optLock.Value = vbUnchecked
   End If
   If optStatus.Value = vbChecked Then
      lblStatus.Caption = " SO Locked"
      cmdLock.Caption = "&Unlock"
   Else
      lblStatus.Caption = " SO Not Locked"
      cmdLock.Caption = "&Lock"
   End If
   
   Set RdoTyp = Nothing
   Exit Function
   
DiaErr1:
   GetSalesOrderType = ""
   txtCst = ""
   txtNme = ""
   optLock.Value = vbUnchecked
   
End Function

Private Sub cmbSon_Click()
   txtSon = cmbSon
   lblSoType = GetSalesOrderType()
   
End Sub


Private Sub cmbSon_LostFocus()
   cmbSon = CheckLen(cmbSon, SO_NUM_SIZE)
   cmbSon = Format(Abs(Val(cmbSon)), SO_NUM_FORMAT)
   txtSon = cmbSon
   lblSoType = GetSalesOrderType()
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2160
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdLock_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   If Val(txtSon) = 0 Then Exit Sub
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   
   If txtNme.ForeColor = ES_RED Then
      MsgBox "Requires A Valid Sales Order.", _
         vbInformation, Caption
   Else
      If optLock.Value = vbUnchecked Then
         'lock
         sMsg = "Lock This Sales Order To Prevent " & vbCrLf _
                & "Further Editing?"
         bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
         If bResponse = vbYes Then
            sSql = "UPDATE SohdTable SET SOLOCKED=1 WHERE " _
                   & "SONUMBER=" & Val(txtSon) & " "
            clsADOCon.ExecuteSQL sSql 'rdExecDirect
            If clsADOCon.ADOErrNum = 0 Then
               SysMsg "Sales Order Locked.", True
               
            Else
               MsgBox "Transaction Failed.", vbInformation, _
                  Caption
            End If
         Else
            CancelTrans
         End If
      Else
         'unlock
         sMsg = "Unlock This Sales Order To Allow " & vbCrLf _
                & "Further Editing?"
         bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
         If bResponse = vbYes Then
            sSql = "UPDATE SohdTable SET SOLOCKED=0 WHERE " _
                   & "SONUMBER=" & Val(txtSon) & " "
            clsADOCon.ExecuteSQL sSql 'rdExecDirect
            If clsADOCon.ADOErrNum = 0 Then
               SysMsg "Sales Order Unlocked.", True
            Else
               MsgBox "Transaction Failed.", vbInformation, _
                  Caption
            End If
         Else
            CancelTrans
         End If
      End If
   End If
   lblSoType = GetSalesOrderType()
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillSalesOrds
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
   Set SaleSLf10a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtCst.BackColor = Es_FormBackColor
   txtNme.BackColor = Es_FormBackColor
   txtSon = SO_NUM_FORMAT
   
End Sub

Private Sub optLock_Click()
   FillSalesOrds
   
End Sub

Private Sub txtNme_Change()
   If Left(txtNme, 9) = "*** Sales" Then txtNme.ForeColor = ES_RED _
           Else txtNme.ForeColor = vbBlack
   
End Sub

Private Sub txtSon_LostFocus()
   txtSon = Format(Abs(Val(txtSon)), SO_NUM_FORMAT)
   If Val(txtSon) > 0 Then
      lblSoType = GetSalesOrderType()
   Else
      lblSoType = ""
   End If
   
End Sub
