VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form CyclCYe02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Initialize a Cycle Count"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CyclCYe02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtDsc 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Tag             =   "2"
      ToolTipText     =   "Description (40 Chars Max)"
      Top             =   1800
      Width           =   3375
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4080
      TabIndex        =   4
      ToolTipText     =   "Add This Cycle Count ID"
      Top             =   480
      Width           =   915
   End
   Begin VB.TextBox txtAbc 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Suggested Count ID (20 Char Max, 5 Min)"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.ComboBox txtDate 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Tag             =   "4"
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox cmbAbc 
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "CyclCYe02a.frx":07AE
      Left            =   1560
      List            =   "CyclCYe02a.frx":07BE
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select From List"
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   4080
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
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
      FormDesignHeight=   2490
      FormDesignWidth =   5085
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "ABC Class"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Count Date"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cycle Count ID"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   1335
   End
End
Attribute VB_Name = "CyclCYe02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'New 2/26/04
Option Explicit
Dim bOnLoad As Byte
Dim bNewCount As Byte
Dim sOldId As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbAbc_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   If cmbAbc.ListCount > 0 Then
      For iList = 0 To cmbAbc.ListCount - 1
         If cmbAbc = cmbAbc.List(iList) Then b = 1
      Next
      If b = 0 Then
         Beep
         cmbAbc = cmbAbc.List(0)
      End If
      txtAbc.Enabled = True
      txtDsc.Enabled = True
      txtAbc = Format(txtDate, "yyyymmdd") & "-" & cmbAbc
      sOldId = txtAbc
      On Error Resume Next
      txtAbc.SetFocus
   Else
      MsgBox "ABC Classes Are Not Properly Set.", _
         vbInformation, Caption
   End If
   
End Sub


Private Sub cmdAdd_Click()
   Dim bResponse As Byte
   If Len(Trim(txtAbc)) < 5 Then
      MsgBox "The Cycle Count ID Must Be At Least (5) Chars.", _
         vbInformation, Caption
      Exit Sub
   End If
   
   bResponse = MsgBox("Add Cycle Count ID " & txtAbc & "?", _
               ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      AddCycleCount
   Else
      CancelTrans
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "5402"
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub



Private Sub Form_Activate()
   Dim b As Byte
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      b = GetABCSetup()
      If b = 0 Then
         MsgBox "ABC Cycle Counting Has Not Been Initialized.", _
            vbInformation, Caption
         Unload Me
      Else
         FillCombo
      End If
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
   If bNewCount = 0 Then FormUnload
   Set CyclCYe02a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDate = Format(ES_SYSDATE, "mm/dd/yyyy")
   
End Sub

Private Sub FillCombo()
   cmbAbc.Clear
   sSql = "Qry_FillABCCombo"
   LoadComboBox cmbAbc
   If cmbAbc.ListCount > 0 Then cmbAbc = cmbAbc.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub txtAbc_LostFocus()
   txtAbc = CheckLen(txtAbc, 20)
   If Len(txtAbc) < 5 Then
      Beep
      MsgBox "5 Characters Minimum.", _
         vbInformation, Caption
      txtAbc = sOldId
   End If
   
End Sub


Private Sub txtDate_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtDate_LostFocus()
   txtDate = CheckDateEx(txtDate)
   txtAbc = Format(txtDate, "yyyymmdd")
   cmdAdd.Enabled = True
   
End Sub


Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 40)
   txtDsc = StrCase(txtDsc)
   
End Sub



'Private Sub TableDocs()
'
'End Sub
'
Private Sub AddCycleCount()
   On Error Resume Next
   Dim RdoCyc As ADODB.Recordset
   sSql = "SELECT * FROM CchdTable WHERE CCREF='FOOBAR'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCyc, ES_KEYSET)
   With RdoCyc
      .AddNew
      !CCREF = txtAbc
      !CCDESC = txtDsc
      !CCPLANDATE = Format(txtDate, "mm/dd/yyyy")
      !CCCREATEDBY = sInitials
      !CCABCCODE = cmbAbc
      .Update
   End With
   Set RdoCyc = Nothing
   If Err > 0 Then
      MsgBox "That Cycle ID Has Already Been Recorded.", _
         vbInformation
   Else
      bNewCount = 1
      CyclCYe04.cmbCid = txtAbc
      CyclCYe04.txtDsc = txtDsc
      CyclCYe04.lblCabc = cmbAbc
      CyclCYe04.txtPlan = txtDate
      SysMsg "Cycle Count Has Been Added.", True
      CyclCYe04.Show
      Unload Me
   End If
End Sub

Private Function GetABCSetup() As Byte
   Dim RdoSet As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_GetABCPreference"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSet, ES_FORWARD)
   If bSqlRows Then
      With RdoSet
         If Not IsNull(.Fields(0)) Then
            GetABCSetup = .Fields(0)
         Else
            GetABCSetup = 0
         End If
         ClearResultSet RdoSet
      End With
   End If
   Set RdoSet = Nothing
   Exit Function
   
DiaErr1:
   GetABCSetup = 0
   sProcName = "getabcsetup"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
   
End Function
