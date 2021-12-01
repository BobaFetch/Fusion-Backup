VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PadmPRe02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Product Classes"
   ClientHeight    =   2550
   ClientLeft      =   2055
   ClientTop       =   1065
   ClientWidth     =   5685
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PadmPRe02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtNte 
      Height          =   975
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1320
      Width           =   3735
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Enter/Revise Product Class (4 Char)"
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   960
      Width           =   3075
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4680
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   90
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   2040
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2550
      FormDesignWidth =   5685
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Notes:"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Class"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "PadmPRe02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Dim RdoCls As ADODB.Recordset
Dim bGoodClass As Byte
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd






Private Sub cmbCls_Change()
   If Len(cmbCls) > 4 Then cmbCls = Left(cmbCls, 4)
   
End Sub

Private Sub cmbCls_Click()
   bGoodClass = GetClass()
   
End Sub


Private Sub cmbCls_LostFocus()
   cmbCls = CheckLen(cmbCls, 4)
   If Len(cmbCls) Then
      bGoodClass = GetClass()
      If Not bGoodClass Then AddProductClass
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 1302
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillProductClasses
      If cmbCls.ListCount > 0 Then
         cmbCls = cmbCls.List(0)
         bGoodClass = GetClass()
      End If
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
   Set RdoCls = Nothing
   Set PadmPRe01a = Nothing
End Sub



Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 30)
   txtDsc = StrCase(txtDsc)
   If bGoodClass Then
      On Error Resume Next
      'RdoCls.Edit
      RdoCls!CCDESC = "" & txtDsc
      RdoCls.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub



Private Function GetClass() As Byte
   Dim sPcode As String
   
   sPcode = Compress(cmbCls)
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM PclsTable WHERE CCREF='" & sPcode & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCls, ES_KEYSET)
   If bSqlRows Then
      With RdoCls
         cmbCls = "" & Trim(!CCCODE)
         txtDsc = "" & Trim(!CCDESC)
         txtNte = "" & Trim(!CCNOTES)
      End With
      GetClass = True
   Else
      txtDsc = ""
      txtNte = ""
      GetClass = False
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getclass"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AddProductClass()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sCode As String
   
   On Error GoTo DiaErr1
   sCode = Compress(cmbCls)
   If sCode = "ALL" Then
      MsgBox "Illegal Product Class Name.", vbExclamation, Caption
      Exit Sub
   End If
   bResponse = IllegalCharacters(cmbCls)
   If bResponse > 0 Then
      MsgBox "The Product Class Contains An Illegal " & Chr$(bResponse) & ".", _
         vbExclamation, Caption
      Exit Sub
   End If
   sMsg = cmbCls & " Wasn't Found. Add The Product Class?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      sSql = "INSERT INTO PclsTable (CCREF,CCCODE) " _
             & "VALUES('" & sCode & "','" & cmbCls & "')"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.RowsAffected Then
         SysMsg "Product Class Added.", True
         AddComboStr cmbCls.hwnd, cmbCls
         bGoodClass = GetClass()
         On Error Resume Next
         txtDsc.SetFocus
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getclass"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtNte_LostFocus()
   txtNte = CheckLen(txtNte, 255)
   txtNte = StrCase(txtNte, ES_FIRSTWORD)
   If bGoodClass Then
      On Error Resume Next
      'RdoCls.Edit
      RdoCls!CCNOTES = "" & txtNte
      RdoCls.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub
