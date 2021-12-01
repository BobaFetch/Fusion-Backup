VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form SadmSLe08a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ship Via"
   ClientHeight    =   3030
   ClientLeft      =   1200
   ClientTop       =   855
   ClientWidth     =   5700
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   1202
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SadmSLe08a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   1800
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1360
      TabIndex        =   1
      Top             =   960
      Width           =   3495
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
   Begin VB.ComboBox cmbShv 
      Height          =   315
      Left            =   1360
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Enter/Revise A Location  (4 char)"
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4800
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   90
      Width           =   875
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
      FormDesignHeight=   3030
      FormDesignWidth =   5700
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ship Via"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "SadmSLe08a"
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
'8/5/05 New
Dim RdoLoc As ADODB.Recordset

Dim bOnLoad As Byte
Dim bGoodLoc As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT SHPVIA FROM CshpviaTable "
   LoadComboBox cmbShv, -1
   If cmbShv.ListCount > 0 Then cmbShv = cmbShv.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub cmbShv_Click()
   bGoodLoc = GetShipVia()
   
End Sub


Private Sub cmbShv_LostFocus()
   cmbShv = CheckLen(cmbShv, 8)
   If Len(cmbShv) Then
      cmbShv = Compress(cmbShv)
      bGoodLoc = GetShipVia()
      If bGoodLoc = 0 Then AddShipVia
   Else
      bGoodLoc = 0
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   cmbShv = ""
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 1207
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   If bOnLoad Then FillCombo
   bOnLoad = 0
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   cUR.CurrentRegion = cmbShv
   SaveCurrentSelections
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set RdoLoc = Nothing
   Set SadmSLe08a = Nothing
   
End Sub




Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub



Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 255)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   On Error Resume Next
   If bGoodLoc Then
      'RdoLoc.Edit
      RdoLoc!SHPVCOMT = "" & txtCmt
      RdoLoc.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 30)
   txtDsc = StrCase(txtDsc)
   On Error Resume Next
   If bGoodLoc Then
      'RdoLoc.Edit
      RdoLoc!SHPVDESC = "" & txtDsc
      RdoLoc.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub



Private Function GetShipVia() As Byte
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM CshpviaTable WHERE SHPVIA='" _
          & Compress(cmbShv) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLoc, ES_KEYSET)
   If bSqlRows Then
      With RdoLoc
         cmbShv = "" & Trim(!SHPVIA)
         txtDsc = "" & Trim(!SHPVDESC)
         txtCmt = "" & Trim(!SHPVCOMT)
      End With
      GetShipVia = 1
   Else
      txtDsc = ""
      txtCmt = ""
      GetShipVia = 0
   End If
   Exit Function
   
DiaErr1:
   sProcName = "GetShipVia"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AddShipVia()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sShipVia As String
   
   sShipVia = cmbShv
   bResponse = IllegalCharacters(cmbShv)
   If bResponse > 0 Then
      MsgBox "The Ship Via Contains An Illegal " & Chr$(bResponse) & ".", _
         vbExclamation, Caption
      Exit Sub
   End If
   sMsg = sShipVia & " Wasn't Found. Add The Ship Via Type.?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error GoTo DiaErr1
      sSql = "INSERT INTO CshpviaTable (SHPVIA) " _
             & "VALUES('" & sShipVia & "')"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.RowsAffected Then
         SysMsg "Ship Via Added.", True
         cmbShv = sShipVia
         AddComboStr cmbShv.hwnd, sShipVia
         bGoodLoc = GetShipVia()
         On Error Resume Next
         txtDsc.SetFocus
      Else
         MsgBox "Couldn't Add Ship Via.", vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "addLocation"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
