VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form InspRTe03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Discrepancy Codes"
   ClientHeight    =   2190
   ClientLeft      =   2385
   ClientTop       =   1680
   ClientWidth     =   5745
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   Icon            =   "InspRTe03a.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2190
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "InspRTe03a.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   840
      Width           =   3475
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5400
      Top             =   1800
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2190
      FormDesignWidth =   5745
   End
   Begin VB.TextBox txtCmt 
      Height          =   555
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   4245
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Select Or Add Characteristic Code (12 Char Max)"
      Top             =   450
      Width           =   1675
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4770
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Code Id"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "InspRTe03a"
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
Dim RdoCde As ADODB.Recordset
Dim bOnLoad As Byte
Dim bGoodCode As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCde_Click()
   bGoodCode = GetCode()
   
End Sub

Private Sub cmbCde_LostFocus()
   cmbCde = CheckLen(cmbCde, 12)
   cmbCde = Trim(cmbCde)
   If Len(cmbCde) = 0 Then
      cmdCan.SetFocus
      Exit Sub
   Else
      bGoodCode = GetCode()
   End If
   If Not bGoodCode Then AddCode
   
End Sub


Private Sub cmdCan_Click()
   txtDsc_LostFocus
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 6103
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = 0
      FillCombo
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
   Set RdoCde = Nothing
   Set InspRTe03a = Nothing
   
End Sub


Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 255)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   If bGoodCode Then
      On Error Resume Next
      RdoCde!CDENOTES = "" & txtCmt
      RdoCde.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 40)
   txtDsc = StrCase(txtDsc)
   If bGoodCode Then
      On Error Resume Next
      RdoCde!CDEDESC = "" & txtDsc
      RdoCde.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub



Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_FillDescripancyCodes"
   LoadComboBox cmbCde
   If cmbCde.ListCount > 0 Then cmbCde = cmbCde.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub AddCode()
   Dim bResponse As Byte
   Dim sNewCode As String
   bResponse = MsgBox(cmbCde & " Wasn't Found. Add It?", ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      bGoodCode = False
      MouseCursor 0
      cmbCde = ""
      On Error Resume Next
      cmdCan.SetFocus
      Exit Sub
   End If
   MouseCursor 13
   sNewCode = Compress(cmbCde)
   On Error Resume Next
   'RdoCde.Close
   '10/29/02
   ' sSql = "Select * FROM RjcdTable"
   ' bSqlRows = GetDataSet(RdoCde, ES_KEYSET)
   ' RdoCde.AddNew
   'RdoCde!CDEREF = "" & sNewCode
   'RdoCde!CDENUM = "" & cmbCde
   'RdoCde.Update
   Err.Clear
   clsADOCon.ADOErrNum = 0
   sSql = "INSERT INTO RjcdTable (CDEREF,CDENUM) " _
          & "VALUES('" & sNewCode & "','" _
          & cmbCde & "')"
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then
      AddComboStr cmbCde.hwnd, cmbCde
      MouseCursor 0
      SysMsg cmbCde & " Successfully Added.", True
      bGoodCode = GetCode()
   Else
      MsgBox "Was Unable To Add the Code.", _
         vbExclamation, Caption
   End If
   
   MouseCursor 0
   AddComboStr cmbCde.hwnd, cmbCde
   SysMsg cmbCde & " Was Successfully Added.", True
   bGoodCode = GetCode()
   
   On Error Resume Next
   txtDsc.SetFocus
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "addcode"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume DiaErr2
DiaErr2:
   MsgBox "Couldn't Add Discrepancy Code.", _
      vbExclamation, Caption
   DoModuleErrors Me
   
End Sub

Private Function GetCode() As Byte
   Dim sRejCode As String
   sRejCode = Compress(cmbCde)
   MouseCursor 13
   On Error GoTo DiaErr1
   sSql = "SELECT *  FROM RjcdTable WHERE CDEREF='" & sRejCode & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCde, ES_KEYSET)
   If bSqlRows Then
      With RdoCde
         cmbCde = "" & Trim(!CDENUM)
         txtDsc = "" & Trim(!CDEDESC)
         txtCmt = "" & Trim(!CDENOTES)
      End With
      GetCode = True
   Else
      GetCode = False
      txtDsc = ""
      txtCmt = ""
      'RdoCde.Close
   End If
   MouseCursor 0
   Exit Function
   
DiaErr1:
   sProcName = "getcode"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub
