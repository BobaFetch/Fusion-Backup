VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PurcPRf10a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change A Manufacturer's Nickname"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRf10a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5520
      TabIndex        =   10
      ToolTipText     =   "Apply New Nickname"
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtNew 
      Height          =   288
      Left            =   1680
      TabIndex        =   1
      Tag             =   "3"
      Top             =   1560
      Width           =   1332
   End
   Begin VB.ComboBox cmbMfr 
      Height          =   288
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Manufacturer or Enter a New Vendor (10 Char Max)"
      Top             =   600
      Width           =   1555
   End
   Begin VB.TextBox txtManu 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   4
      Tag             =   "2"
      ToolTipText     =   "Manufacturer's Name (30)"
      Top             =   960
      Width           =   3360
   End
   Begin VB.TextBox txtType 
      Height          =   288
      Left            =   5940
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Tag             =   "3"
      ToolTipText     =   "Type (2 Char)"
      Top             =   960
      Width           =   372
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5520
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   4200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2460
      FormDesignWidth =   6540
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Nickname"
      Height          =   288
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   288
      Index           =   32
      Left            =   5280
      TabIndex        =   8
      Top             =   960
      Width           =   1632
   End
   Begin VB.Label lblNum 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   3360
      TabIndex        =   7
      Top             =   600
      Width           =   324
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nickname"
      Height          =   288
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   288
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1428
   End
End
Attribute VB_Name = "PurcPRf10a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'10/17/06 New
Option Explicit
Dim bCanceled As Byte
Dim bOnLoad As Byte
Dim bGoodMfr As Byte
Dim bGoodNew As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Function GetManufacturer() As Byte
   Dim RdoMfr As ADODB.Recordset
   cmdUpd.Enabled = False
   sSql = "SELECT MFGR_NICKNAME,MFGR_NUMBER,MFGR_TYPE,MFGR_NAME " _
          & "FROM MfgrTable WHERE MFGR_REF='" & Compress(cmbMfr) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoMfr, ES_FORWARD)
   If bSqlRows Then
      With RdoMfr
         cmbMfr = "" & Trim(!MFGR_NICKNAME)
         lblNum = Format(!MFGR_NUMBER, "##0")
         txtType = "" & Trim(!MFGR_TYPE)
         txtManu = "" & Trim(!MFGR_NAME)
         GetManufacturer = 1
      End With
      ClearResultSet RdoMfr
   Else
      GetManufacturer = 0
      lblNum = ""
      txtType = ""
      txtManu = ""
   End If
   Set RdoMfr = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getmanufac"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub cmbMfr_Click()
   bGoodMfr = GetManufacturer()
   
   
End Sub


Private Sub cmbMfr_LostFocus()
   If bCanceled = 1 Then Exit Sub
   If Len(cmbMfr) Then
      bGoodMfr = GetManufacturer()
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   bCanceled = 1
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4359
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdUpd_Click()
   Dim bResponse As Byte
   If Len(Trim(txtNew)) < 4 Then
      Beep
      MsgBox "(4) Characters Or More Please.", vbInformation, Caption
      Exit Sub
   End If
   
   If txtNew = "ALL" Then
      Beep
      MsgBox "ALL Is An Illegal Nickname.", vbExclamation, Caption
      Exit Sub
   End If
   bResponse = IllegalCharacters(cmbMfr)
   If bResponse > 0 Then
      MsgBox "The Nickname Contains An Illegal " & Chr$(bResponse) & ".", _
         vbExclamation, Caption
      Exit Sub
   End If
   
   bResponse = MsgBox("Change The Manufacturer's Nickname?", _
               ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      cmdUpd.Enabled = False
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      
      sSql = "UPDATE MfgrTable SET MFGR_REF='" & Compress(txtNew) _
             & "',MFGR_NICKNAME='" & txtNew & "' WHERE MFGR_REF='" _
             & Compress(cmbMfr) & "'"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "UPDATE PartTable SET PAMANUFACTURER='" & Compress(txtNew) _
             & "' WHERE PAMANUFACTURER='" & Compress(cmbMfr) & "'"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         SysMsg "Manufacturer Changed.", True
         bCanceled = 1
         cmbMfr = txtNew
         txtNew = ""
         FillCombo
         'bGoodMfr = GetManufacturer()
      Else
         clsADOCon.RollbackTrans
         MsgBox "Could Not Successfully Change The Nickname.", _
            vbInformation, Caption
      End If
      
   Else
      CancelTrans
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then FillCombo
   bOnLoad = 0
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set PurcPRf10a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtType.BackColor = BackColor
   txtManu.BackColor = BackColor
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   cmbMfr.Clear
   sSql = "SELECT MFGR_REF,MFGR_NICKNAME FROM MfgrTable "
   LoadComboBox cmbMfr
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub




Private Function CheckNewMfr() As Byte
   Dim iList As Integer
   CheckNewMfr = 1
   For iList = 0 To cmbMfr.ListCount - 1
      If Compress(txtNew) = Compress(cmbMfr.List(iList)) _
                  Then CheckNewMfr = 0
      Next
      
   End Function
   
   Private Sub txtNew_LostFocus()
      txtNew = CheckLen(txtNew, 10)
      If Len(txtNew) Then
         bGoodNew = CheckNewMfr()
      Else
         bGoodNew = 0
      End If
      If bGoodNew Then cmdUpd.Enabled = True _
                                        Else cmdUpd.Enabled = False
      
   End Sub
