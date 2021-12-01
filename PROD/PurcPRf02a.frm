VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PurcPRf02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete A Vendor"
   ClientHeight    =   2010
   ClientLeft      =   3000
   ClientTop       =   1710
   ClientWidth     =   6345
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRf02a.frx":0000
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
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5280
      TabIndex        =   5
      ToolTipText     =   "Delete The Current Vendor"
      Top             =   600
      Width           =   915
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Vendors With No Purchase Orders"
      Top             =   960
      Width           =   1555
   End
   Begin VB.TextBox txtNme 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   2
      Tag             =   "2"
      Top             =   1320
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
      Left            =   5880
      Top             =   1440
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2010
      FormDesignWidth =   6345
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
      Left            =   120
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
      Left            =   120
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
      Top             =   960
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor Name"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1425
   End
End
Attribute VB_Name = "PurcPRf02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'5/26/06 Added Checks
Option Explicit
Dim bOnLoad As Byte

Private Function CheckWindows() As Byte
   Dim b As Byte
   b = Val(GetSetting("Esi2000", "Sections", "admn", 0))
   b = b + Val(GetSetting("Esi2000", "Sections", "fina", 0))
   b = b + Val(GetSetting("Esi2000", "Sections", "qual", 0))
   b = b + Val(GetSetting("Esi2000", "Sections", "invc", 0))
   If b > 0 Then
      lblWrn(0) = sSysCaption & " Has Determined " & b & " Other Open Section(s)"
      lblWrn(0).Visible = True
      lblWrn(1).Visible = True
      cmdDel.Enabled = False
   End If
   CheckWindows = b
   
End Function

Private Sub cmbCst_Click()
   GetDelVendor
   
End Sub


Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If Len(Trim(cmbCst)) Then GetDelVendor
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdDel_Click()
   If txtNme.ForeColor = ES_RED Then
      MsgBox "Requires A Valid Vendor.", _
         vbInformation, Caption
   Else
      DeleteTheVendor
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4351
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      CheckWindows
      txtNme.BackColor = BackColor
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   
   lblWrn(0).ForeColor = ES_RED
   lblWrn(1).ForeColor = ES_RED
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set PurcPRf02a = Nothing
   
End Sub



Private Sub FillCombo()
   On Error GoTo DiaErr1
   MouseCursor 13
   cmbCst.Clear
   sSql = "SELECT VndrTable.VEREF,VENICKNAME FROM VndrTable LEFT " _
          & "JOIN PohdTable ON VEREF = PohdTable.POVENDOR " _
          & "WHERE (PohdTable.POVENDOR Is Null) AND VEREF<>'NONE'"
   LoadComboBox cmbCst
   If cmbCst.ListCount > 0 Then
      cmbCst = cmbCst.List(0)
      'GetDelVendor
   Else
      MsgBox "No Vendors Available To Delete.", _
         vbInformation, Caption
      txtNme = " "
   End If
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub GetDelVendor()
   Dim RdoCst As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT VEREF,VENICKNAME,VEBNAME FROM VndrTable WHERE " _
          & "VEREF='" & Compress(cmbCst) & "' AND VEREF<>'NONE'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCst, ES_FORWARD)
   If bSqlRows Then
      With RdoCst
         cmbCst = "" & Trim(!VENICKNAME)
         txtNme = "" & Trim(!VEBNAME)
         ClearResultSet RdoCst
      End With
   Else
      txtNme = "*** Vendor Wasn't Found ***"
   End If
   Set RdoCst = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getdelvend"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub txtNme_Change()
   If Left(txtNme, 6) = "*** Ve" Or Trim(txtNme) = "" Then
      txtNme.ForeColor = ES_RED
      cmdDel.Enabled = False
   Else
      txtNme.ForeColor = Es_TextForeColor
      cmdDel.Enabled = True
   End If
   
End Sub



Private Sub DeleteTheVendor()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sVend As String
   sVend = Compress(cmbCst)
   
   sMsg = "It Is Not A Good Idea To Delete A Vendor If " & vbCr _
          & "There Is Any Chance That It Is In Use Right Now."
   MsgBox sMsg, vbExclamation, Caption
   
   sMsg = "This Function Permanently Removes The Vendor." & vbCr _
          & "Are You Sure That You Want To Continue?      "
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      'start checking
      'Purchase Orders
      On Error Resume Next
      sSql = "SELECT DISTINCT POVENDOR FROM PohdTable WHERE " _
             & "POVENDOR='" & sVend & "'"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.RowsAffected <> 0 Then
         MsgBox "That Vendor Has At Least One Purchase " & vbCr _
            & "Order And Cannot Be Deleted.", vbExclamation, Caption
         Exit Sub
      End If
      'Invoice?
      sSql = "SELECT DISTINCT VIVENDOR FROM VihdTable WHERE " _
             & "VIVENDOR='" & sVend & "'"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.RowsAffected <> 0 Then
         MsgBox "That Vendor Has At Least One Invoice " & vbCr _
            & "And Cannot Be Deleted.", vbExclamation, Caption
         Exit Sub
      End If
      
      'Journal
      sSql = "SELECT DISTINCT DCVENDOR FROM JritTable WHERE " _
             & "DCVENDOR='" & sVend & "'"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.RowsAffected <> 0 Then
         MsgBox "That Vendor Has At Least One Journal " & vbCr _
            & "Entry And Cannot Be Deleted.", vbExclamation, Caption
         Exit Sub
      End If
      
      'RejTag?
      sSql = "SELECT DISTINCT REJVENDOR FROM RjhdTable WHERE " _
             & "REJVENDOR='" & sVend & "'"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.RowsAffected <> 0 Then
         MsgBox "That Vendor Has At Least One Rejection " & vbCr _
            & "Tag And Cannot Be Deleted.", vbExclamation, Caption
         Exit Sub
      End If
      
      'Commissions
      sSql = "SELECT DISTINCT SPVENDOR FROM SprsTable WHERE " _
             & "SPVENDOR='" & sVend & "'"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.RowsAffected <> 0 Then
         MsgBox "That Vendor Is Used In A Sales Person's " & vbCr _
            & "Commissions And Cannot Be Deleted.", vbExclamation, Caption
         Exit Sub
      End If
      
      'Checks
      sSql = "SELECT DISTINCT CHKVENDOR FROM ChksTable WHERE " _
             & "CHKVENDOR='" & sVend & "'"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.RowsAffected <> 0 Then
         MsgBox "A Check Has Been Written To That Vendor" & vbCr _
            & "And Cannot Be Deleted.", vbExclamation, Caption
         Exit Sub
      End If
      
      sMsg = "Last Chance. Are You Sure That You Want" & vbCr _
             & "To Delete Vendor " & cmbCst & "?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      
      If bResponse = vbYes Then
         
         clsADOCon.ADOErrNum = 0
         
         sSql = "DELETE FROM VndrTable WHERE VEREF='" & sVend & "' "
         clsADOCon.ExecuteSQL sSql
         
         sSql = "DELETE FROM BuyvTable WHERE BYVENDOR='" & sVend & "' "
         clsADOCon.ExecuteSQL sSql
         
         If clsADOCon.ADOErrNum = 0 Then
            SysMsg "Vendor Was Deleted.", True
            cUR.CurrentVendor = ""
            txtNme = ""
            FillCombo
         Else
            MsgBox "Could Not Delete The Vendor.", _
               vbExclamation, Caption
         End If
      Else
         CancelTrans
      End If
   Else
      CancelTrans
   End If
   
End Sub
