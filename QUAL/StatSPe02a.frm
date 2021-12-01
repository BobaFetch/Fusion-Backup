VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form StatSPe02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Team Members"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   Tag             =   "3"
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "StatSPe02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtCmt 
      Height          =   735
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   8
      Tag             =   "9"
      Top             =   3360
      Width           =   3495
   End
   Begin VB.TextBox txtDpt 
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Tag             =   "3"
      Top             =   3000
      Width           =   1605
   End
   Begin VB.ComboBox cmbDiv 
      Height          =   315
      Left            =   4440
      Sorted          =   -1  'True
      TabIndex        =   6
      Tag             =   "3"
      ToolTipText     =   "Select Division From List"
      Top             =   2640
      Width           =   860
   End
   Begin VB.ComboBox cmbShp 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   5
      Tag             =   "3"
      ToolTipText     =   "Select Shop From List"
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox txtLst 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Tag             =   "2"
      Top             =   2280
      Width           =   2085
   End
   Begin VB.TextBox txtMid 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Tag             =   "3"
      Top             =   1920
      Width           =   285
   End
   Begin VB.TextBox txtFst 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Tag             =   "2"
      Top             =   1560
      Width           =   2085
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Tag             =   "1"
      Top             =   1200
      Width           =   875
   End
   Begin VB.ComboBox cmbMem 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Enter New Team Member  (15 Char) Or Select From List (15 Char Max)"
      Top             =   720
      Width           =   1875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4440
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5280
      Top             =   3960
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4425
      FormDesignWidth =   5460
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   18
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   17
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      Height          =   255
      Index           =   6
      Left            =   3720
      TabIndex        =   16
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shop"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   15
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   14
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Middle Init"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   13
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Team Member"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "StatSPe02a"
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
Dim AdoMem As ADODB.Recordset

Dim bCanceled As Byte
Dim bOnLoad As Byte
Dim bGoodMem As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbDiv_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   
   cmbDiv = CheckLen(cmbDiv, 4)
   For iList = 0 To cmbDiv.ListCount - 1
      If cmbDiv.List(iList) = cmbDiv Then b = 1
   Next
   If b = 0 Then
      Beep
      cmbDiv = ""
   End If
   If bGoodMem Then
      On Error Resume Next
      AdoMem!TMMDIVISION = cmbDiv
      AdoMem.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub cmbMem_Change()
   If Len(cmbMem) > 15 Then cmbMem = Left(cmbMem, 15)
   
End Sub

Private Sub cmbMem_Click()
   bGoodMem = GetMember()
   
End Sub


Private Sub cmbMem_LostFocus()
   cmbMem = Compress(cmbMem)
   cmbMem = CheckLen(cmbMem, 15)
   bGoodMem = GetMember()
   If bCanceled Then Exit Sub
   If Len(Trim(cmbMem)) Then
      If bGoodMem = 0 Then AddMember
   End If
   
End Sub


Private Sub cmbShp_Click()
   GetShop
   
End Sub

Private Sub cmbShp_LostFocus()
   cmbShp = CheckLen(cmbShp, 12)
   GetShop
   If bGoodMem Then
      On Error Resume Next
      AdoMem!TMMSHOP = Compress(cmbShp)
      AdoMem.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCanceled = True
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 6302
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillDivisions
      FillCombo
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
   Set AdoMem = Nothing
   Set StatSPe02a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_FillSPTeam"
   LoadComboBox cmbMem, -1
   
   sSql = "SELECT SHPREF,SHPNUM FROM ShopTable "
   LoadComboBox cmbShp
   cmbShp.ToolTipText = "*** No Valid Shop Selected ***"
   If cmbMem.ListCount > 0 Then cmbMem = cmbMem.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetMember() As Byte
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM RjtmTable WHERE TMMID='" _
          & Compress(cmbMem) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoMem, ES_KEYSET)
   If bSqlRows Then
      With AdoMem
         cmbMem = "" & Trim(!TMMID)
         If !TMMNUMBER > 0 Then
            txtNum = Format(!TMMNUMBER, "000000")
         Else
            txtNum = ""
         End If
         txtLst = "" & Trim(!TMMLSTNAME)
         txtMid = "" & Trim(!TMMMINIT)
         txtFst = "" & Trim(!TMMFSTNAME)
         cmbShp = "" & Trim(!TMMSHOP)
         cmbDiv = "" & Trim(!TMMDIVISION)
         txtDpt = "" & Trim(!TMMDEPARTMENT)
         txtCmt = "" & Trim(!TMMNOTES)
         GetMember = 1
         GetShop
      End With
   Else
      txtNum = ""
      txtLst = ""
      txtMid = ""
      txtFst = ""
      txtCmt = ""
      GetMember = 0
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getmember"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 255)
   If bGoodMem Then
      On Error Resume Next
      AdoMem!TMMNOTES = Trim(txtCmt)
      AdoMem.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtDpt_LostFocus()
   txtDpt = CheckLen(txtDpt, 12)
   If bGoodMem Then
      On Error Resume Next
      AdoMem!TMMDEPARTMENT = Trim(txtDpt)
      AdoMem.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtFst_LostFocus()
   txtFst = CheckLen(txtFst, 20)
   txtFst = StrCase(txtFst)
   If bGoodMem Then
      On Error Resume Next
      AdoMem!TMMFSTNAME = txtFst
      AdoMem.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtLst_LostFocus()
   txtLst = CheckLen(txtLst, 20)
   txtLst = StrCase(txtLst)
   If bGoodMem Then
      On Error Resume Next
      AdoMem!TMMLSTNAME = txtLst
      AdoMem.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtMid_LostFocus()
   txtMid = CheckLen(txtMid, 1)
   If bGoodMem Then
      On Error Resume Next
      AdoMem!TMMMINIT = txtMid
      AdoMem.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtNum_LostFocus()
   txtNum = CheckLen(txtNum, 6)
   If Val(txtNum) <> 0 Then
      txtNum = Format(Abs(Val(txtNum)), "000000")
   Else
      Beep
      txtNum = ""
   End If
   If bGoodMem Then
      On Error Resume Next
      AdoMem!TMMNUMBER = Val(txtNum)
      AdoMem.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub



Private Sub GetShop()
   Dim AdoShp As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "SELECT SHPREF,SHPNUM,SHPDESC FROM " _
          & "ShopTable WHERE SHPREF='" & Compress(cmbShp) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoShp, ES_FORWARD)
   If bSqlRows Then
      With AdoShp
         cmbShp = "" & Trim(!SHPNUM)
         cmbShp.ToolTipText = "" & Trim(!SHPDESC)
         ClearResultSet AdoShp
      End With
   Else
      Beep
      cmbShp = ""
      cmbShp.ToolTipText = "*** No Valid Shop Selected ***"
   End If
   Set AdoShp = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getshop"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub AddMember()
   Dim bResponse As Byte
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   sMsg = "Team Member " & cmbMem & " Wasn't Found." & vbCr _
          & "Add The New Team Member?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      clsADOCon.ADOErrNum = 0
      
      sSql = "INSERT INTO RjtmTable(TMMID,TMMSHOP," _
             & "TMMDIVISION,TMMDEPARTMENT) " _
             & "VALUES('" & cmbMem & "','" _
             & Compress(cmbShp) & "','" _
             & cmbDiv & "','" _
             & txtDpt & "')"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then
         MsgBox "Team Member Was Successfully Added.", _
            vbInformation, Caption
         AddComboStr cmbMem.hwnd, cmbMem
         bGoodMem = GetMember()
      Else
         MsgBox "Could Not Successfully Add Team Member.", _
            vbExclamation, Caption
         bGoodMem = 0
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "addmember"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
