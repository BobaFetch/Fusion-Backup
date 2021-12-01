VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form EstiESf03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancel An RFQ"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "EstiESf03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdCbd 
      Caption         =   "C&ancel"
      Height          =   315
      Left            =   5040
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Cancel The Current RFQ"
      Top             =   1440
      Width           =   875
   End
   Begin VB.ComboBox cmbRfq 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select Or Enter RFQ Number"
      Top             =   1440
      Width           =   2040
   End
   Begin VB.TextBox txtDsc 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   2
      Tag             =   "2"
      Top             =   1800
      Width           =   3795
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select A Customer"
      Top             =   720
      Width           =   1555
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5040
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5760
      Top             =   2040
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2550
      FormDesignWidth =   5985
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "RFQ Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "Enter Or Select An RFQ"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Enter Or Select An RFQ"
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label txtNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   1080
      Width           =   3255
   End
End
Attribute VB_Name = "EstiESf03a"
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
Dim bGoodRfq As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Function GetThisRfq() As Byte
   Dim RdoRfq As ADODB.Recordset
   
   Dim sCust As String
   sCust = Compress(cmbCst)
   sSql = "SELECT * FROM RfqsTable WHERE RFQREF='" & cmbRfq & "' " _
          & "AND RFQREF<>'NONE' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRfq, ES_FORWARD)
   If bSqlRows Then
      With RdoRfq
         If sCust <> "" & Trim(!RFQCUST) Then
            MsgBox "That RFQ BelOngs To Another Customer." _
               & vbExclamation, Caption
            GetThisRfq = 2
         Else
            txtDsc = "" & Trim(!RFQDESC)
            GetThisRfq = 1
         End If
         ClearResultSet RdoRfq
      End With
   Else
      txtDsc = ""
      GetThisRfq = 0
   End If
   Set RdoRfq = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getthisrfq"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub cmbCst_Click()
   FindCustomer Me, cmbCst
   FillCustomerRFQs Me, cmbCst
   If cmbRfq.ListCount > 0 Then
      bGoodRfq = GetThisRfq()
   Else
      txtDsc = ""
   End If
   
End Sub


Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   FindCustomer Me, cmbCst
   FillCustomerRFQs Me, cmbCst
   If cmbRfq.ListCount > 0 Then
      bGoodRfq = GetThisRfq()
   Else
      txtDsc = ""
   End If
   
End Sub


Private Sub cmbRfq_Click()
   bGoodRfq = GetThisRfq()
   If bGoodRfq = 0 Then txtDsc = ""
   
End Sub


Private Sub cmbRfq_LostFocus()
   cmbRfq = CheckLen(cmbRfq, 14)
   If bCancel = 1 Then Exit Sub
   bGoodRfq = GetThisRfq()
   If cmbRfq.ListCount > 0 Then bGoodRfq = GetThisRfq()
   If bGoodRfq = 0 Then txtDsc = ""
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = 1
   
End Sub


Private Sub cmdCbd_Click()
   If cmbCst <> "" Then
      If bGoodRfq <> 1 Or cmbRfq.ListCount = 0 Then
         MsgBox "Must Select A Valid RFQ.", _
            vbInformation, Caption
      Else
         CancelTheRfq
      End If
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3552
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillCustomers
      If cmbCst.ListCount > 0 Then
         FindCustomer Me, cmbCst
         FillCustomerRFQs Me, cmbCst
         If cmbRfq.ListCount > 0 Then bGoodRfq = GetThisRfq()
      End If
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
   Set EstiESf03a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub



Private Sub CancelTheRfq()
   Dim bResponse As Byte
   Dim sMsg As String
   
   sMsg = "Warning. This Function Cancels The RFQ " & vbCrLf _
          & "And Resets All Assigned Estimates To No RFQ." & vbCrLf _
          & "Do You Wish To Continue?.."
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbNo Then Exit Sub
   
   sMsg = "Do You Wish To Cancel RFQ " & cmbRfq & " For " _
          & "Customer " & cmbCst & "?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      Err = 0
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      
      sSql = "DELETE FROM RfqsTable WHERE RFQREF='" & cmbRfq & "' "
      clsADOCon.ExecuteSQL sSql ', rdExecDirect
      
      sSql = "UPDATE EstiTable SET BIDRFQ='NONE' " _
             & "WHERE BIDRFQ='" & cmbRfq & "' "
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         MsgBox "RFQ Was Successfully Canceled.", _
            vbInformation, Caption
         FillCustomerRFQs Me, cmbCst
         If cmbRfq.ListCount > 0 Then
            bGoodRfq = GetThisRfq()
         Else
            txtDsc = ""
         End If
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         MsgBox "RFQ Was Not Successfully Canceled.", _
            vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   
End Sub
