VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form StatSPe05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Family ID's"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   6305
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
      Picture         =   "StatSPe05a.frx":0000
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
   Begin VB.ComboBox cmbFam 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Enter New Family ID (15 Char) Or Select From List (15 Char Max)"
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
      Caption         =   "Family ID"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1395
   End
End
Attribute VB_Name = "StatSPe05a"
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
Dim AdoFam As ADODB.Recordset
Dim bCanceled As Byte
Dim bOnLoad As Byte
Dim bGoodFam As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbFam_Change()
   If Len(cmbFam) > 15 Then cmbFam = Left(cmbFam, 15)
   
End Sub

Private Sub cmbFam_Click()
   bGoodFam = GetFamily()
   
End Sub


Private Sub cmbFam_LostFocus()
   cmbFam = CheckLen(cmbFam, 15)
   If bCanceled Then Exit Sub
   bGoodFam = GetFamily()
   If Len(Trim(cmbFam)) Then
      If bGoodFam = 0 Then AddFamily
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
      OpenHelpContext 6305
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
   Set AdoFam = Nothing
   Set StatSPe05a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_FillSPFamily"
   LoadComboBox cmbFam
   If cmbFam.ListCount > 0 Then cmbFam = cmbFam.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetFamily() As Byte
   On Error GoTo DiaErr1
   sSql = "SELECT FAMREF,FAMID,FAMDESC,FAMNOTES FROM " _
          & "RjfmTable WHERE FAMREF='" & Compress(cmbFam) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoFam, ES_KEYSET)
   If bSqlRows Then
      With AdoFam
         cmbFam = "" & Trim(!FAMID)
         txtDsc = "" & Trim(!FAMDESC)
         txtCmt = "" & Trim(!FAMNOTES)
         GetFamily = 1
      End With
   Else
      GetFamily = 0
      txtDsc = ""
      txtCmt = ""
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getfamily"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AddFamily()
   Dim bResponse As Byte
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   bResponse = IllegalCharacters(cmbFam)
   If bResponse > 0 Then
      MsgBox "The Family Contains An Illegal " & Chr$(bResponse) & ".", _
         vbExclamation, Caption
      Exit Sub
   End If
   
   sMsg = "Family ID " & cmbFam & " Wasn't Found." & vbCr _
          & "Add The New Family ID?.."
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      clsADOCon.ADOErrNum = 0
      sSql = "INSERT INTO RjfmTable (FAMREF,FAMID) " _
             & "VALUES('" & Compress(cmbFam) & "','" _
             & cmbFam & "')"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then
         AddComboStr cmbFam.hwnd, cmbFam
         MsgBox "Successfully Added Family ID.", _
            vbInformation, Caption
         bGoodFam = GetFamily()
      Else
         MsgBox "Couldn't Add Family ID.", _
            vbExclamation, Caption
         bGoodFam = 0
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "addfamily"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 40)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   If bGoodFam Then
      On Error Resume Next
      AdoFam!FAMNOTES = "" & txtCmt
      AdoFam.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 40)
   txtDsc = StrCase(txtDsc)
   If bGoodFam Then
      On Error Resume Next
      AdoFam!FAMDESC = "" & txtDsc
      AdoFam.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub
