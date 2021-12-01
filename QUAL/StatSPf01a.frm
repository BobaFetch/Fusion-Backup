VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form StatSPf01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete A Reasoning Code"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "StatSPf01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   4920
      TabIndex        =   7
      ToolTipText     =   "Delete This Reasoning Code"
      Top             =   600
      Width           =   875
   End
   Begin VB.TextBox txtCmt 
      BackColor       =   &H8000000F&
      Height          =   735
      Left            =   1680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "9"
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox txtDsc 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   1
      Tag             =   "2"
      Top             =   1200
      Width           =   3475
   End
   Begin VB.ComboBox cmbRes 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Contains Only Codes Not In Use"
      Top             =   840
      Width           =   1875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4920
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5520
      Top             =   2280
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2625
      FormDesignWidth =   5850
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Notes"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reasoning Code"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1515
   End
End
Attribute VB_Name = "StatSPf01a"
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
Dim AdoRcd As ADODB.Recordset
Dim bCanceled As Byte
Dim bOnLoad As Byte
Dim bGoodRes As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd



Private Sub cmbRes_Click()
   bGoodRes = GetCode()
   
End Sub


Private Sub cmbRes_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   cmbRes = CheckLen(cmbRes, 15)
   If bCanceled Then Exit Sub
   For iList = 0 To cmbRes.ListCount - 1
      If cmbRes = cmbRes.List(iList) Then bByte = 1
   Next
   If bByte = 0 Then
      Beep
      If cmbRes.ListCount > 0 Then cmbRes = cmbRes.List(0)
   End If
   bGoodRes = GetCode()
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCanceled = True
   
End Sub


Private Sub cmdDel_Click()
   If bGoodRes Then
      DeleteCode
   Else
      MsgBox "Requires A Valid Reasoning Code.", _
         vbExclamation, Caption
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 6350
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST, ES_RESIZE
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdoRcd = Nothing
   Set StatSPf01a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDsc.BackColor = BackColor
   txtCmt.BackColor = BackColor
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   cmbRes.Clear
   txtDsc = ""
   sSql = "SELECT RjrcTable.RCOREF, RjrcTable.RCOID " _
          & "FROM RjrcTable LEFT JOIN RjkyTable ON RjrcTable.RCOREF" _
          & "= RjkyTable.KEYREASON Where (RjkyTable.KEYREASON Is Null)"
   LoadComboBox cmbRes
   If cmbRes.ListCount > 0 Then cmbRes = cmbRes.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetCode() As Byte
   On Error GoTo DiaErr1
   sSql = "SELECT RCOREF,RCOID,RCODESC,RCONOTES FROM " _
          & "RjrcTable WHERE RCOREF='" & Compress(cmbRes) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoRcd, ES_FORWARD)
   If bSqlRows Then
      With AdoRcd
         cmbRes = "" & Trim(!RCOID)
         txtDsc = "" & Trim(!RCODESC)
         txtCmt = "" & Trim(!RCONOTES)
         ClearResultSet AdoRcd
         GetCode = 1
      End With
   Else
      GetCode = 0
      txtDsc = ""
      txtCmt = ""
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getcode"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub DeleteCode()
   Dim bResponse As Byte
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   sMsg = "This Function Permanently Removes All Records" & vbCr _
          & "Of " & cmbRes & ". Are You Sure That You Want To?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      
      
      sSql = "SELECT KEYREASON FROM RjkyTable WHERE " _
             & "KEYREASON='" & Compress(cmbRes) & "'"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "DELETE FROM RjrcTable WHERE RCOREF='" _
             & Compress(cmbRes) & "' "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum > 0 Then
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         MsgBox "Could Not Successfully Delete The Code.", _
            vbExclamation, Caption
      Else
         clsADOCon.CommitTrans
         MsgBox "Successfully Deleted " & cmbRes & ".", _
            vbInformation, Caption
         FillCombo
         cmbRes.SetFocus
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "deletecode"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 40)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   
End Sub


Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 40)
   txtDsc = StrCase(txtDsc)
   
End Sub
