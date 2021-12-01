VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form CyclCYf06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mark Cycle Count Audited"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CyclCYf06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtAud 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Tag             =   "2"
      ToolTipText     =   "Person Or Firm Auditing"
      Top             =   2040
      Width           =   3075
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4320
      TabIndex        =   8
      ToolTipText     =   "Save Changes To The Cycle Count"
      Top             =   480
      Width           =   875
   End
   Begin VB.ComboBox cmbCid 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "List Includes Cycle ID's Not Locked Or Completed"
      Top             =   960
      Width           =   2115
   End
   Begin VB.TextBox txtDsc 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   1320
      Width           =   3375
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4320
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   2520
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2655
      FormDesignWidth =   5280
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Auditor"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Audit Date"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   1335
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
End
Attribute VB_Name = "CyclCYf06a"
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
   sSql = "SELECT CCREF,CCDESC,CCPLANDATE,CCABCCODE,CCCOUNTSAVED," _
          & "CCAUDITOR,CCAUDITDATE FROM CchdTable WHERE " _
          & "(CCREF='" & Trim(cmbCid) & "' AND " _
          & "CCCOUNTLOCKED=1 AND CCUPDATED=1)"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCid, ES_FORWARD)
   If bSqlRows Then
      With RdoCid
         lblCabc = "" & Trim(!CCABCCODE)
         txtDsc = "" & Trim(!CCDESC)
         If Not IsNull(!CCAUDITDATE) Then
            txtEnd = Format(!CCAUDITDATE, "mm/dd/yyyy")
         Else
            txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
         End If
         txtAud = "" & Trim(!CCAUDITOR)
         If Len(txtAud) Then cmdDel.Enabled = True
                'Else txtAud.Enabled = False
         GetCycleCount = 1
         ClearResultSet RdoCid
      End With
   Else
      GetCycleCount = 0
      MsgBox "That Count ID Wasn't Found Or Isn't Reconciled.", _
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
   
   On Error Resume Next
   cmdDel.Enabled = False
   If Trim(txtAud) = "" Then
      MsgBox "Requires An Auditor.", _
         vbInformation, Caption
      txtAud.SetFocus
      Exit Sub
   End If
   sMsg = "Continue To Mark The Count Audited?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   clsADOCon.ADOErrNum = 0
   If bResponse = vbYes Then
      sSql = "UPDATE CchdTable SET CCAUDITED=1, " _
             & "CCAUDITOR='" & txtAud & "'," _
             & "CCAUDITDATE='" & txtEnd & "' WHERE " _
             & "CCREF='" & cmbCid & "'"
      clsADOCon.ExecuteSql sSql
      If clsADOCon.ADOErrNum = 0 Then
         SysMsg "Count Marked Audited.", True
         FillCombo
      Else
         MsgBox "Could Not Update The Count.", _
            vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "5456"
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
   Set CyclCYf06a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDsc.BackColor = Me.BackColor
   txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   cmbCid.Clear
   sSql = "SELECT CCCOUNTLOCKED,CCREF FROM CchdTable WHERE " _
          & "CCCOUNTLOCKED=1 AND CCUPDATED=1 ORDER BY CCREF"
   LoadComboBox cmbCid
   If cmbCid.ListCount > 0 Then
      If Trim(cmbCid) = "" Then cmbCid = cmbCid.List(0)
      'bGoodCount = GetCycleCount()
   Else
      MsgBox "There Are No Reconciled Counts Recorded.", _
         vbInformation, Caption
      Unload Me
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtAud_LostFocus()
   txtAud = CheckLen(txtAud, 30)
   txtAud = StrCase(txtAud)
   If Len(txtAud) Then cmdDel.Enabled = True
   
End Sub

Private Sub txtEnd_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtEnd_LostFocus()
   txtEnd = CheckDateEx(txtEnd)
   
End Sub
