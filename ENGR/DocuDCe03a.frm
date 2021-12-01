VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form DocuDCe03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Document Classes"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "DocuDCe03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optAdcn 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   3000
      Width           =   855
   End
   Begin VB.CheckBox optSht 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox txtNte 
      Height          =   1005
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   3475
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Enter/Revise A Class (16 char)"
      Top             =   600
      Width           =   2000
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   960
      Width           =   3475
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4440
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
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
      FormDesignWidth =   5385
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Boeing Drawings)"
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   11
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Track ADCN's"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Track Multiple Sheets?"
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Notes"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Document Class"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "DocuDCe03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'4/26/06 Corrected GetSomeClass query
Option Explicit
Dim RdoDoc As ADODB.Recordset

Dim bOnLoad As Byte
Dim bGoodClass As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCls_Change()
   If Len(cmbCls) > 16 Then cmbCls = Left(cmbCls, 16)
   
End Sub

Private Sub cmbCls_Click()
   bGoodClass = GetSomeClass()
   
End Sub


Private Sub cmbCls_LostFocus()
   cmbCls = CheckLen(cmbCls, 16)
   If Len(cmbCls) Then
      bGoodClass = GetSomeClass()
      If Not bGoodClass Then AddSomeClass
   Else
      bGoodClass = False
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmbCls = ""
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3303
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = 0
      FillClasses
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   sLastDocClass = cmbCls
   SaveSetting "Esi2000", "EsiEngr", "DocClass", sLastDocClass
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set RdoDoc = Nothing
   Set DocuDCe03a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillClasses()
   On Error GoTo DiaErr1
   sSql = "Qry_FillDocumentCombo"
   LoadComboBox cmbCls
   If cmbCls.ListCount > 0 Then cmbCls = cmbCls.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillclasses"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetSomeClass() As Byte
   Dim sClass As String
   sClass = Compress(cmbCls)
   On Error GoTo DiaErr1
   sSql = "SELECT DCLREF,DCLNAME,DCLDESC,DCLNOTES,DCLSHEETS,DCLADCN " _
          & "FROM DclsTable WHERE DCLREF='" & Compress(sClass) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDoc, ES_KEYSET)
   If bSqlRows Then
      With RdoDoc
         cmbCls = "" & Trim(!DCLNAME)
         txtDsc = "" & Trim(!DCLDESC)
         txtNte = "" & Trim(!DCLNOTES)
         If !DCLSHEETS Then
            optSht.value = vbChecked
         Else
            optSht.value = vbUnchecked
         End If
         If !DCLADCN Then
            optAdcn.value = vbChecked
         Else
            optAdcn.value = vbUnchecked
         End If
      End With
      GetSomeClass = True
   Else
      txtDsc = ""
      txtNte = ""
      GetSomeClass = False
   End If
   
   Exit Function
   
DiaErr1:
   sProcName = "getsomeclass"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AddSomeClass()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sClass As String
   
   If Len(Trim(cmbCls)) = 0 Then Exit Sub
   sClass = Compress(cmbCls)
   sMsg = cmbCls & " Wasn't Found. Add The Class?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error GoTo DiaErr1
      sSql = "INSERT INTO DclsTable (DCLREF,DCLNAME) " _
             & "VALUES('" & sClass & "','" & cmbCls & "')"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      If clsADOCon.RowsAffected > 0 Then
         AddComboStr cmbCls.hwnd, cmbCls
         SysMsg "Class Added.", True
         bGoodClass = GetSomeClass()
      Else
         MsgBox "Couldn't Add The Class.", vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "addsomecl"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub optAdcn_Click()
   If bGoodClass Then
      On Error Resume Next
      'RdoDoc.Edit
      RdoDoc!DCLADCN = optAdcn.value
      RdoDoc.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub optSht_Click()
   If bGoodClass Then
      On Error Resume Next
      'RdoDoc.Edit
      RdoDoc!DCLSHEETS = optSht.value
      RdoDoc.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 40)
   txtDsc = StrCase(txtDsc)
   If bGoodClass Then
      On Error Resume Next
      'RdoDoc.Edit
      RdoDoc!DCLDESC = "" & txtDsc
      RdoDoc.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtNte_LostFocus()
   txtNte = CheckLen(txtNte, 255)
   txtNte = StrCase(txtNte, ES_FIRSTWORD)
   If bGoodClass Then
      On Error Resume Next
      'RdoDoc.Edit
      RdoDoc!DCLNOTES = "" & txtNte
      RdoDoc.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub
