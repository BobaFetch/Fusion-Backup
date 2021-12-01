VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form SadmSLe01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Divisions"
   ClientHeight    =   2895
   ClientLeft      =   1200
   ClientTop       =   855
   ClientWidth     =   5625
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   1201
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SadmSLe01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbDiv 
      Height          =   315
      Left            =   1360
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Enter/Revise A Division (4 char)"
      Top             =   600
      Width           =   860
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4800
      Top             =   2640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2895
      FormDesignWidth =   5625
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1360
      TabIndex        =   1
      Top             =   960
      Width           =   3495
   End
   Begin VB.CheckBox optUse 
      Alignment       =   1  'Right Justify
      Caption         =   "Force Divisions"
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.TextBox txtCmt 
      Height          =   975
      Left            =   1360
      MultiLine       =   -1  'True
      TabIndex        =   2
      Tag             =   "9"
      Top             =   1320
      Width           =   3475
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4680
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Forces Use Of Fixed Divisions In All Sections)"
      Height          =   255
      Index           =   3
      Left            =   1680
      TabIndex        =   8
      Top             =   2445
      Width           =   3975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "SadmSLe01a"
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
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter
Dim rdoDiv As ADODB.Recordset

Dim bOnLoad As Byte
Dim bGoodDiv As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbDiv_Click()
   bGoodDiv = GetDivision()
   
End Sub


Private Sub cmbDiv_LostFocus()
   cmbDiv = CheckLen(cmbDiv, 4)
   If Len(cmbDiv) Then
      cmbDiv = Compress(cmbDiv)
      bGoodDiv = GetDivision()
      If bGoodDiv = 0 Then AddDivision
   Else
      bGoodDiv = 0
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmbDiv = ""
   
End Sub


Private Sub cmdHlp_Click()
   Dim l&
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 1201
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub





Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillDivisions
      If cmbDiv.ListCount > 0 Then
         cmbDiv = cmbDiv.List(0)
         bGoodDiv = GetDivision()
      End If
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   sSql = "SELECT TOP 1 DIVREF,DIVDESC,DIVCOMT FROM " _
          & "CdivTable WHERE DIVREF= ? "
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Size = 6
   AdoParameter.Type = adChar
   AdoQry.Parameters.Append AdoParameter
   
   
'   Set RdoQry = RdoCon.CreateQuery("", sSql)
'   RdoQry.MaxRows = 1
       
   Dim sString$
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   Set SadmSLe01a = Nothing
   
End Sub




Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub optUse_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 255)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   On Error Resume Next
   If bGoodDiv Then
      rdoDiv!DIVCOMT = "" & txtCmt
      rdoDiv.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 40)
   txtDsc = StrCase(txtDsc, vbProperCase)
   On Error Resume Next
   If bGoodDiv Then
      rdoDiv!DIVDESC = "" & txtDsc
      rdoDiv.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub



Private Function GetDivision() As Byte
   Dim sDivision As String
   sDivision = cmbDiv
   
   On Error GoTo DiaErr1
   'RdoQry(0) = sDivision
   AdoQry.Parameters(0).value = sDivision
   bSqlRows = clsADOCon.GetQuerySet(rdoDiv, AdoQry, ES_KEYSET, True)
   If bSqlRows Then
      With rdoDiv
         cmbDiv = "" & Trim(!DIVREF)
         txtDsc = "" & Trim(!DIVDESC)
         txtCmt = "" & Trim(!DIVCOMT)
      End With
      GetDivision = 1
   Else
      txtDsc = ""
      txtCmt = ""
      GetDivision = 0
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getdivision"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AddDivision()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sDivision As String
   
   sDivision = cmbDiv
   If Trim(cmbDiv) < 2 Then
      MsgBox "Divisions Must Be At Least (2) Characters.", _
         vbInformation, Caption
      Exit Sub
   End If
   bResponse = IllegalCharacters(cmbDiv)
   If bResponse > 0 Then
      MsgBox "The Division Contains An Illegal " & Chr$(bResponse) & ".", _
         vbExclamation, Caption
      Exit Sub
   End If
   sMsg = sDivision & " Wasn't Found. Add The Division?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error GoTo DiaErr1
      sSql = "INSERT INTO CdivTable (DIVREF) " _
             & "VALUES('" & sDivision & "')"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.RowsAffected Then
         SysMsg "Division Added.", True
         cmbDiv = sDivision
         AddComboStr cmbDiv.hwnd, sDivision
         bGoodDiv = GetDivision()
         On Error Resume Next
         txtDsc.SetFocus
      Else
         MsgBox "Couldn't The Add Division.", vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "adddivisi"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
