VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form InspRTe06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inspectors"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "InspRTe06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbDiv 
      Height          =   315
      Left            =   1500
      Sorted          =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Select A Division (2 char)"
      Top             =   2160
      Width           =   860
   End
   Begin VB.TextBox txtNte 
      Height          =   855
      Left            =   1500
      MultiLine       =   -1  'True
      TabIndex        =   8
      Tag             =   "9"
      Top             =   3600
      Width           =   4095
   End
   Begin VB.TextBox txtDpt 
      Height          =   285
      Left            =   1500
      TabIndex        =   7
      Tag             =   "2"
      Top             =   3240
      Width           =   2085
   End
   Begin VB.TextBox txtStp 
      Height          =   285
      Left            =   1500
      TabIndex        =   6
      Tag             =   "31"
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox txtBdg 
      Height          =   285
      Left            =   1500
      TabIndex        =   5
      Tag             =   "3"
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txtLst 
      Height          =   285
      Left            =   1500
      TabIndex        =   3
      Tag             =   "2"
      Top             =   1800
      Width           =   2085
   End
   Begin VB.TextBox txtMid 
      Height          =   285
      Left            =   1500
      TabIndex        =   2
      Tag             =   "3"
      Top             =   1440
      Width           =   375
   End
   Begin VB.ComboBox cmbIns 
      Height          =   315
      Left            =   1500
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Or Enter Inspector ID (No Spaces O8r Dashes) "
      Top             =   720
      Width           =   1665
   End
   Begin VB.TextBox txtFst 
      Height          =   285
      Left            =   1500
      TabIndex        =   1
      Tag             =   "2"
      Top             =   1080
      Width           =   1275
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4800
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   4440
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4695
      FormDesignWidth =   5745
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Active"
      Height          =   285
      Index           =   8
      Left            =   3360
      TabIndex        =   21
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblAct 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4200
      TabIndex        =   20
      Top             =   720
      Width           =   300
   End
   Begin VB.Label lblDiv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2500
      TabIndex        =   19
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      Height          =   285
      Index           =   16
      Left            =   240
      TabIndex        =   18
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   17
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   16
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Stamp Number"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   15
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Badge/Clock "
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   14
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   13
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Middle Initial"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inspector Id"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "InspRTe06a"
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
Dim RdoIns As ADODB.Recordset
Dim bOnLoad As Byte
Dim bGoodInsp As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Function GetInspector() As Byte
   Dim sInsp As String
   sInsp = Compress(cmbIns)
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM RinsTable WHERE " _
          & "INSID='" & sInsp & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoIns, ES_KEYSET)
   If bSqlRows Then
      With RdoIns
         cmbIns = "" & Trim(!INSID)
         txtFst = "" & Trim(!INSFIRST)
         txtMid = "" & Trim(!INSMIDD)
         txtLst = "" & Trim(!INSLAST)
         txtBdg = "" & Trim(!INSBADGE)
         cmbDiv = "" & Trim(!INSDIVISION)
         txtStp = Format(0 + !INSSTAMP, "####0")
         txtDpt = "" & Trim(!INSDEPT)
         txtNte = "" & Trim(!INSNOTES)
         If !INSACTIVE = 1 Then
            lblAct = "Y"
         Else
            lblAct = "N"
         End If
      End With
      GetInspector = True
   Else
      txtFst = ""
      txtMid = ""
      txtLst = ""
      txtBdg = ""
      txtStp = ""
      txtDpt = ""
      txtNte = ""
      lblAct = ""
      GetInspector = False
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getinspect"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub cmbDiv_Click()
   GetDivision
   
End Sub


Private Sub cmbDiv_LostFocus()
   cmbDiv = Compress(cmbDiv)
   cmbDiv = CheckLen(cmbDiv, 4)
   GetDivision
   If bGoodInsp Then
      On Error Resume Next
      If lblDiv.ForeColor = ES_RED Then
         RdoIns!INSDIVISION = ""
      Else
         RdoIns!INSDIVISION = "" & cmbDiv
      End If
      RdoIns.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub cmbIns_Click()
   bGoodInsp = GetInspector()
   
End Sub

Private Sub cmbIns_LostFocus()
   cmbIns = CheckLen(cmbIns, 12)
   cmbIns = Compress(cmbIns)
   If Len(cmbIns) Then
      bGoodInsp = GetInspector()
      If Not bGoodInsp Then AddInspector
   Else
      bGoodInsp = False
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 6106
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillDivisions
      FillInspectors
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set RdoIns = Nothing
   Set InspRTe06a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillInspectors()
   On Error GoTo DiaErr1
   sSql = "Qry_FillInspectorsAll"
   LoadComboBox cmbIns, -1
   If cmbIns.ListCount > 0 Then cmbIns = cmbIns.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillinspe"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub AddInspector()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sInsp As String
   On Error GoTo DiaErr1
   sInsp = Compress(cmbIns)
   sMsg = cmbIns & " Doesn't Exist. Add The Inspector?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      sSql = "INSERT INTO RinsTable (INSID,INSDIVISION) " _
             & "VALUES('" & sInsp & "','" & Compress(cmbDiv) & "')"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.RowsAffected > 0 Then
         AddComboStr cmbIns.hwnd, cmbIns
         SysMsg "Inspector Was Added.", True
         bGoodInsp = GetInspector()
      Else
         MsgBox "Couldn't Add The Inspector.", vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "addinspect"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub lblDiv_Change()
   If Left(lblDiv, 8) = "*** Divi" Then
      lblDiv.ForeColor = ES_RED
   Else
      lblDiv.ForeColor = vbBlack
   End If
   
End Sub

Private Sub txtBdg_LostFocus()
   txtBdg = CheckLen(txtBdg, 10)
   If bGoodInsp Then
      On Error Resume Next
      RdoIns!INSBADGE = "" & txtBdg
      RdoIns.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtDpt_LostFocus()
   txtDpt = CheckLen(txtDpt, 20)
   txtDpt = StrCase(txtDpt)
   If bGoodInsp Then
      On Error Resume Next
      RdoIns!INSDEPT = "" & txtDpt
      RdoIns.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtFst_LostFocus()
   txtFst = CheckLen(txtFst, 10)
   txtFst = StrCase(txtFst)
   If bGoodInsp Then
      On Error Resume Next
      RdoIns!INSFIRST = "" & txtFst
      RdoIns.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtLst_LostFocus()
   txtLst = CheckLen(txtLst, 20)
   txtLst = StrCase(txtLst)
   If bGoodInsp Then
      On Error Resume Next
      RdoIns!INSLAST = "" & txtLst
      RdoIns.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtMid_LostFocus()
   txtMid = CheckLen(txtMid, 1)
   If bGoodInsp Then
      On Error Resume Next
      RdoIns!INSMIDD = "" & txtMid
      RdoIns.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtNte_LostFocus()
   txtNte = CheckLen(txtNte, 255)
   txtNte = StrCase(txtNte, ES_FIRSTWORD)
   If bGoodInsp Then
      On Error Resume Next
      RdoIns!INSNOTES = "" & txtNte
      RdoIns.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtStp_LostFocus()
   txtStp = CheckLen(txtStp, 5)
   txtStp = Format(Abs(Val(txtStp)), "####0")
   If bGoodInsp Then
      On Error Resume Next
      RdoIns!INSSTAMP = Val(txtStp)
      RdoIns.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub



Private Sub GetDivision()
   Dim RdoDiv As ADODB.Recordset
   If Trim(cmbDiv) = "" Then
      lblDiv = ""
      Exit Sub
   End If
   sSql = "SELECT DIVREF,DIVDESC FROM " _
          & "CdivTable WHERE DIVREF='" & cmbDiv & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDiv, ES_FORWARD)
   If bSqlRows Then
      With RdoDiv
         cmbDiv = "" & Trim(!DIVREF)
         lblDiv = "" & Trim(!DIVDESC)
         ClearResultSet RdoDiv
      End With
   Else
      lblDiv = "*** Division Wasn't Found ***"
   End If
   Set RdoDiv = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getdivision"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
