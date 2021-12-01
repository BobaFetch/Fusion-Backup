VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form StatSPe04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Process ID's"
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
      Picture         =   "StatSPe04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtCmt 
      Height          =   735
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   2
      Tag             =   "9"
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Tag             =   "2"
      Top             =   1200
      Width           =   3475
   End
   Begin VB.ComboBox cmbPrc 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Enter NewProcess ID (15 Char) Or Select From List (15 Char Max)"
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
      Top             =   90
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
      Caption         =   "Process ID"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1395
   End
End
Attribute VB_Name = "StatSPe04a"
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
Dim AdoPrc As ADODB.Recordset
Dim bCanceled As Byte
Dim bOnLoad As Byte
Dim bGoodPrc As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbPrc_Change()
   If Len(cmbPrc) > 15 Then cmbPrc = Left(cmbPrc, 15)
   
End Sub

Private Sub cmbPrc_Click()
   bGoodPrc = GetProcess()
   
End Sub


Private Sub cmbPrc_LostFocus()
   cmbPrc = CheckLen(cmbPrc, 15)
   If bCanceled Then Exit Sub
   bGoodPrc = GetProcess()
   If Len(Trim(cmbPrc)) Then
      If bGoodPrc = 0 Then AddProcess
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
      OpenHelpContext 6304
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
   FormLoad Me
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdoPrc = Nothing
   Set StatSPe04a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_FillSPProcessID"
   LoadComboBox cmbPrc
   If cmbPrc.ListCount > 0 Then cmbPrc = cmbPrc.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetProcess() As Byte
   On Error GoTo DiaErr1
   sSql = "SELECT PROREF,PROID,PRODESC,PRONOTES FROM " _
          & "RjprTable WHERE PROREF='" & Compress(cmbPrc) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoPrc, ES_KEYSET)
   If bSqlRows Then
      With AdoPrc
         cmbPrc = "" & Trim(!PROID)
         txtDsc = "" & Trim(!PRODESC)
         txtCmt = "" & Trim(!PRONOTES)
         GetProcess = 1
      End With
   Else
      GetProcess = 0
      txtDsc = ""
      txtCmt = ""
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getprocess"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AddProcess()
   Dim bResponse As Byte
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   bResponse = IllegalCharacters(cmbPrc)
   If bResponse > 0 Then
      MsgBox "The Process ID Contains An Illegal " & Chr$(bResponse) & ".", _
         vbExclamation, Caption
      Exit Sub
   End If
   sMsg = "Process ID " & cmbPrc & " Wasn't Found." & vbCr _
          & "Add The New Process ID?.."
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      clsADOCon.ADOErrNum = 0
      sSql = "INSERT INTO RjprTable (PROREF,PROID) " _
             & "VALUES('" & Compress(cmbPrc) & "','" _
             & cmbPrc & "')"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then
         AddComboStr cmbPrc.hwnd, cmbPrc
         MsgBox "Successfully Added Process ID.", _
            vbInformation, Caption
         bGoodPrc = GetProcess()
      Else
         MsgBox "Couldn't Add Process ID.", _
            vbExclamation, Caption
         bGoodPrc = 0
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "addprocess"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 40)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   If bGoodPrc Then
      On Error Resume Next
      AdoPrc!PRONOTES = "" & txtCmt
      AdoPrc.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 40)
   txtDsc = StrCase(txtDsc)
   If bGoodPrc Then
      On Error Resume Next
      AdoPrc!PRODESC = "" & txtDsc
      AdoPrc.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub
