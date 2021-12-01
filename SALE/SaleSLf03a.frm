VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form SaleSLf03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete A Customer"
   ClientHeight    =   2175
   ClientLeft      =   3000
   ClientTop       =   1710
   ClientWidth     =   6240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLf03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdDel 
      Cancel          =   -1  'True
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5280
      TabIndex        =   5
      ToolTipText     =   "Delete The Current Customer"
      Top             =   480
      Width           =   915
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Customers With No Sales Orders"
      Top             =   1080
      Width           =   1555
   End
   Begin VB.TextBox txtNme 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   2
      Tag             =   "2"
      Top             =   1440
      Width           =   3475
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5760
      Top             =   1680
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2175
      FormDesignWidth =   6240
   End
   Begin VB.Label lblWrn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label lblWrn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Please Close All Other Sections Before Proceeding"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nickname"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1095
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1425
   End
End
Attribute VB_Name = "SaleSLf03a"
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

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtNme.BackColor = BackColor
   
End Sub

Private Function CheckWindows() As Byte
   Dim b As Byte
   b = Val(GetSetting("Esi2000", "Sections", "admn", 0))
   b = b + Val(GetSetting("Esi2000", "Sections", "prod", 0))
   b = b + Val(GetSetting("Esi2000", "Sections", "engr", 0))
   b = b + Val(GetSetting("Esi2000", "Sections", "fina", 0))
   b = b + Val(GetSetting("Esi2000", "Sections", "qual", 0))
   If b > 0 Then
      lblWrn(0) = sSysCaption & " Has Determined " & b & " Other Open Section(s)"
      lblWrn(0).Visible = True
      lblWrn(1).Visible = True
      cmdDel.Enabled = False
   End If
   CheckWindows = b
   
End Function

Private Sub cmbCst_Click()
   GetDelCustomer
   
End Sub


Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If Len(Trim(cmbCst)) Then GetDelCustomer
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdDel_Click()
   If txtNme.ForeColor = ES_RED Then
      MsgBox "Requires A Valid Customer.", _
         vbInformation, Caption
   Else
      DeleteTheCustomer
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2153
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      CheckWindows
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   
   lblWrn(0).ForeColor = ES_RED
   lblWrn(1).ForeColor = ES_RED
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set SaleSLf03a = Nothing
   
End Sub



Private Sub FillCombo()
   On Error GoTo DiaErr1
   MouseCursor 13
   cmbCst.Clear
   sSql = "SELECT CUREF,CUNICKNAME FROM CustTable LEFT " _
          & "JOIN SohdTable ON CUREF = SohdTable.SOCUST " _
          & "WHERE (SohdTable.SOCUST Is Null)"
   LoadComboBox cmbCst
   MouseCursor 0
   If cmbCst.ListCount > 0 Then
      cmbCst = cmbCst.List(0)
      GetDelCustomer
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub GetDelCustomer()
   Dim RdoCst As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_GetCustomerBasics '" & Compress(cmbCst) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCst, ES_FORWARD)
   If bSqlRows Then
      With RdoCst
         cmbCst = "" & Trim(!CUNICKNAME)
         txtNme = "" & Trim(!CUNAME)
         ClearResultSet RdoCst
      End With
   Else
      txtNme = "*** Customer Wasn't Found ***"
   End If
   Set RdoCst = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getdelcust"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub txtNme_Change()
   If Left(txtNme, 6) = "*** Cu" Then
      txtNme.ForeColor = ES_RED
      cmdDel.Enabled = False
   Else
      txtNme.ForeColor = Es_TextForeColor
      cmdDel.Enabled = True
   End If
   
End Sub



Private Sub DeleteTheCustomer()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sCust As String
   sCust = Compress(cmbCst)
   
   sMsg = "It Is Not A Good Idea To Delete A Customer " & vbCrLf _
          & "If There Is Any Chance That It Is In Use Right Now."
   MsgBox sMsg, vbExclamation, Caption
   
   sMsg = "This Function Permanently Removes The Customer." & vbCrLf _
          & "Are You Sure That You Want To Continue?      "
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      'start checking
      'Sales Orders
      On Error Resume Next
      sSql = "SELECT DISTINCT SOCUST FROM SohdTable WHERE " _
             & "SOCUST='" & sCust & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      If clsADOCon.RowsAffected <> 0 Then
         MsgBox "That Customer Has At Least One Sales " & vbCrLf _
            & "Order And Cannot Be Deleted.", vbExclamation, Caption
         Exit Sub
      End If
      'Shouldn't be any, but test PS anyway
      sSql = "SELECT DISTINCT PSCUST FROM PshdTable WHERE " _
             & "PSCUST='" & sCust & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      If clsADOCon.RowsAffected <> 0 Then
         MsgBox "That Customer Has At Least One Packing " & vbCrLf _
            & "Slip And Cannot Be Deleted.", vbExclamation, Caption
         Exit Sub
      End If
      'Invoice?
      sSql = "SELECT DISTINCT INVCUST FROM CihdTable WHERE " _
             & "INVCUST='" & sCust & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      If clsADOCon.RowsAffected <> 0 Then
         MsgBox "That Customer Has At Least One Invoice " & vbCrLf _
            & "And Cannot Be Deleted.", vbExclamation, Caption
         Exit Sub
      End If
      
      'Jounral?
      sSql = "SELECT DISTINCT DCCUST FROM JritTable WHERE " _
             & "DCCUST='" & sCust & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      If clsADOCon.RowsAffected <> 0 Then
         MsgBox "That Customer Has At Least One Journal " & vbCrLf _
            & "Entry And Cannot Be Deleted.", vbExclamation, Caption
         Exit Sub
      End If
      
      'RejTag?
      sSql = "SELECT DISTINCT REJCUST FROM RjhdTable WHERE " _
             & "REJCUST='" & sCust & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      If clsADOCon.RowsAffected <> 0 Then
         MsgBox "That Customer Has At Least One Rejection " & vbCrLf _
            & "Tag And Cannot Be Deleted.", vbExclamation, Caption
         Exit Sub
      End If
      'Document?
      sSql = "SELECT DISTINCT DOCUST FROM DdocTable WHERE " _
             & "DOCUST='" & sCust & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      If clsADOCon.RowsAffected <> 0 Then
         MsgBox "That Customer Has At Least One Document " & vbCrLf _
            & "Attached And Cannot Be Deleted.", vbExclamation, Caption
         Exit Sub
      End If
      
      'Esimating
      sSql = "SELECT DISTINCT BIDCUST FROM EstiTable WHERE " _
             & "BIDCUST='" & sCust & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      If clsADOCon.RowsAffected <> 0 Then
         MsgBox "That Customer Has At Least One Esimate " & vbCrLf _
            & "Attached And Cannot Be Deleted.", vbExclamation, Caption
         Exit Sub
      End If
      
      'RFQ?
      sSql = "SELECT DISTINCT RFQCUST FROM RfqsTable WHERE " _
             & "RFQCUST='" & sCust & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      If clsADOCon.RowsAffected <> 0 Then
         MsgBox "That Customer Has At Least One RFQ " & vbCrLf _
            & "Attached And Cannot Be Deleted.", vbExclamation, Caption
         Exit Sub
      End If
      If clsADOCon.RowsAffected <> 0 Then
         MsgBox "That Customer Has At Least One RFQ " & vbCrLf _
            & "Attached And Cannot Be Deleted.", vbExclamation, Caption
         Exit Sub
      End If
      
      'Tools
      sSql = "SELECT DISTINCT TOOL_CUST FROM TohdTable WHERE " _
             & "TOOL_CUST='" & sCust & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      If clsADOCon.RowsAffected <> 0 Then
         MsgBox "That Customer Has At Least One Tool " & vbCrLf _
            & "Attached And Cannot Be Deleted.", vbExclamation, Caption
         Exit Sub
      End If
      
      'Lots
      sSql = "SELECT DISTINCT LOICUST FROM LoitTable WHERE " _
             & "LOICUST='" & sCust & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      If clsADOCon.RowsAffected <> 0 Then
         MsgBox "That Customer Has At Least One Lot Row " & vbCrLf _
            & "Attached And Cannot Be Deleted.", vbExclamation, Caption
         Exit Sub
      End If
      
      sMsg = "Last Chance. Are You Sure That You Want" & vbCrLf _
             & "To Delete Customer " & cmbCst & "?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbYes Then
         clsADOCon.ADOErrNum = 0
         sSql = "DELETE FROM CustTable WHERE " _
                & "CUREF='" & sCust & "'"
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
         If clsADOCon.ADOErrNum = 0 Then
            SysMsg "Customer Was Deleted.", True
            cUR.CurrentCustomer = ""
            FillCombo
         Else
            MsgBox "Could Not Delete The Customer.", _
               vbExclamation, Caption
         End If
      Else
         CancelTrans
      End If
   Else
      CancelTrans
   End If
   
End Sub
