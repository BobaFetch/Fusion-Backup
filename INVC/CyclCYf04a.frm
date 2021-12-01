VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form CyclCYf04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete A Cycle Count"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CyclCYf04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   4320
      TabIndex        =   8
      ToolTipText     =   "Delete This Cycle Count ID"
      Top             =   480
      Width           =   875
   End
   Begin VB.ComboBox cmbCid 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "List Includes Cycle ID's Not Locked Or Completed"
      Top             =   960
      Width           =   2115
   End
   Begin VB.TextBox txtDsc 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   1320
      Width           =   3375
   End
   Begin VB.CheckBox optSaved 
      Enabled         =   0   'False
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4320
      TabIndex        =   0
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
      FormDesignHeight=   2340
      FormDesignWidth =   5295
   End
   Begin VB.Label lblCabc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      ToolTipText     =   "ABC Code Selected"
      Top             =   960
      Width           =   405
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cycle Count ID"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Count Saved"
      Height          =   255
      Index           =   13
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
End
Attribute VB_Name = "CyclCYf04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
Option Explicit
Dim bOnLoad As Byte
Dim bGoodCount As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Function GetCycleCount() As Byte
   Dim RdoCid As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_GetCycleCountNotLocked '" & Trim(cmbCid) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCid, ES_FORWARD)
   If bSqlRows Then
      With RdoCid
         lblCabc = "" & Trim(!CCABCCODE)
         txtDsc = "" & Trim(!CCDESC)
         optSaved.value = !CCCOUNTSAVED
         GetCycleCount = 1
         ClearResultSet RdoCid
      End With
   Else
      GetCycleCount = 0
      MsgBox "That Count ID Wasn't Found Or Is Locked.", _
         vbInformation, Caption
   End If
   Set RdoCid = Nothing
   Exit Function
   
DiaErr1:
   GetCycleCount = 0
   sProcName = "getcycleco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub cmbCid_Click()
   bGoodCount = GetCycleCount()
   
End Sub


Private Sub cmbCid_LostFocus()
   cmbCid = CheckLen(cmbCid, 20)
   bGoodCount = GetCycleCount()
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdDel_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   If optSaved.value = vbChecked Then
      sMsg = "This Cycle Count Has Been Saved." & vbCr _
             & "Delete All Record Of It Anyway?"
   Else
      sMsg = "This Function Deletes All Records Of The Count ID." & vbCr _
             & "Continue To Delete All Record Of It?"
   End If
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then DeleteCount Else CancelTrans
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "5454"
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
   FormLoad Me, ES_DONTLIST
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set CyclCYf04a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDsc.BackColor = Me.BackColor
   
End Sub

Private Sub FillCombo()
   'Dim RdoCmb As rdoResultset
   On Error GoTo DiaErr1
   cmbCid.Clear
   sSql = "SELECT CCREF FROM CchdTable WHERE CCCOUNTLOCKED=0"
   LoadComboBox cmbCid, -1
   If cmbCid.ListCount > 0 Then
      If Trim(cmbCid) = "" Then cmbCid = cmbCid.List(0)
      'bGoodCount = GetCycleCount()
   Else
      MsgBox "There Are No Unlocked Counts Recorded.", _
         vbInformation, Caption
      Unload Me
   End If
   'Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub DeleteCount()
   'On Error Resume Next
   On Error GoTo whoops
   
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   
   sSql = "DELETE FROM CcLog WHERE CCREF='" & Trim(cmbCid) & "'"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "DELETE FROM CcLotAlloc WHERE CCREF='" & Trim(cmbCid) & "'"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "DELETE FROM CcltTable WHERE CLREF='" & Trim(cmbCid) & "'"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "DELETE FROM CcitTable WHERE CIREF='" & Trim(cmbCid) & "'"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "DELETE FROM CchdTable WHERE CCREF='" & Trim(cmbCid) & "'"
   clsADOCon.ExecuteSQL sSql
   
   clsADOCon.CommitTrans
   MsgBox cmbCid & " Was Successfully Deleted.", _
      vbInformation, Caption
   FillCombo
   Exit Sub
   
whoops:
   clsADOCon.RollbackTrans
   sProcName = "DeleteCount"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
   
End Sub
