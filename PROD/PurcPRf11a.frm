VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PurcPRf11a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete A Manufacturer"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRf11a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
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
      TabIndex        =   8
      ToolTipText     =   "Delete This Manufacturer"
      Top             =   720
      Width           =   855
   End
   Begin VB.ComboBox cmbMfr 
      Height          =   288
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Manufacturer or Enter a New Vendor (10 Char Max)"
      Top             =   720
      Width           =   1555
   End
   Begin VB.TextBox txtManu 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   3
      Tag             =   "2"
      ToolTipText     =   "Manufacturer's Name (30)"
      Top             =   1080
      Width           =   3360
   End
   Begin VB.TextBox txtType 
      Height          =   288
      Left            =   5940
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "3"
      ToolTipText     =   "Type (2 Char)"
      Top             =   1080
      Width           =   372
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5520
      TabIndex        =   1
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
      FormDesignHeight=   2325
      FormDesignWidth =   6540
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   288
      Index           =   32
      Left            =   5280
      TabIndex        =   7
      Top             =   1080
      Width           =   1632
   End
   Begin VB.Label lblNum 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   3360
      TabIndex        =   6
      Top             =   720
      Width           =   324
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nickname"
      Height          =   288
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   288
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1428
   End
End
Attribute VB_Name = "PurcPRf11a"
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
      cmdUpd.Enabled = True
   Else
      cmdUpd.Enabled = False
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
      OpenHelpContext 4360
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdUpd_Click()
   Dim bResponse As Byte
   bResponse = MsgBox("Are You Certain The Manufacturer Is To Be Deleted?", _
               ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      
      sSql = "DELETE FROM MfgrTable WHERE MFGR_REF='" _
             & Compress(cmbMfr) & "'"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "UPDATE PartTable SET PAMANUFACTURER='' WHERE " _
             & "PAMANUFACTURER='" & Compress(cmbMfr) & "'"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         SysMsg "Manufacturer Deleted.", True
         bCanceled = 1
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
   Set PurcPRf11a = Nothing
   
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
