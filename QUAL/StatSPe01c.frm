VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form StatSPe01c 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data Sources For A Key"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "StatSPe01c.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   37
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtDte6 
      Height          =   315
      Left            =   240
      TabIndex        =   20
      Tag             =   "4"
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox txtDrw6 
      Height          =   285
      Left            =   1440
      TabIndex        =   21
      Tag             =   "3"
      ToolTipText     =   "Drawing (Optional)"
      Top             =   4680
      Width           =   2895
   End
   Begin VB.ComboBox cmbTag6 
      Height          =   315
      Left            =   4440
      Sorted          =   -1  'True
      TabIndex        =   22
      Tag             =   "3"
      ToolTipText     =   "Select Or Enter Tag Number (Optional)"
      Top             =   4680
      Width           =   1935
   End
   Begin VB.TextBox txtCmt6 
      Height          =   285
      Left            =   1440
      TabIndex        =   23
      Tag             =   "2"
      ToolTipText     =   "Notes 60 Chars (Optional)"
      Top             =   5040
      Width           =   4695
   End
   Begin VB.ComboBox txtDte5 
      Height          =   315
      Left            =   240
      TabIndex        =   16
      Tag             =   "4"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox txtDrw5 
      Height          =   285
      Left            =   1440
      TabIndex        =   17
      Tag             =   "3"
      ToolTipText     =   "Drawing (Optional)"
      Top             =   3960
      Width           =   2895
   End
   Begin VB.ComboBox cmbTag5 
      Height          =   315
      Left            =   4440
      Sorted          =   -1  'True
      TabIndex        =   18
      Tag             =   "3"
      ToolTipText     =   "Select Or Enter Tag Number (Optional)"
      Top             =   3960
      Width           =   1935
   End
   Begin VB.TextBox txtCmt5 
      Height          =   285
      Left            =   1440
      TabIndex        =   19
      Tag             =   "2"
      ToolTipText     =   "Notes 60 Chars (Optional)"
      Top             =   4320
      Width           =   4695
   End
   Begin VB.ComboBox txtDte4 
      Height          =   315
      Left            =   240
      TabIndex        =   12
      Tag             =   "4"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txtDrw4 
      Height          =   285
      Left            =   1440
      TabIndex        =   13
      Tag             =   "3"
      ToolTipText     =   "Drawing (Optional)"
      Top             =   3240
      Width           =   2895
   End
   Begin VB.ComboBox cmbTag4 
      Height          =   315
      Left            =   4440
      Sorted          =   -1  'True
      TabIndex        =   14
      Tag             =   "3"
      ToolTipText     =   "Select Or Enter Tag Number (Optional)"
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox txtCmt4 
      Height          =   285
      Left            =   1440
      TabIndex        =   15
      Tag             =   "2"
      ToolTipText     =   "Notes 60 Chars (Optional)"
      Top             =   3600
      Width           =   4695
   End
   Begin VB.ComboBox txtDte3 
      Height          =   315
      Left            =   240
      TabIndex        =   8
      Tag             =   "4"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtDrw3 
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Tag             =   "3"
      ToolTipText     =   "Drawing (Optional)"
      Top             =   2520
      Width           =   2895
   End
   Begin VB.ComboBox cmbTag3 
      Height          =   315
      Left            =   4440
      Sorted          =   -1  'True
      TabIndex        =   10
      Tag             =   "3"
      ToolTipText     =   "Select Or Enter Tag Number (Optional)"
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox txtCmt3 
      Height          =   285
      Left            =   1440
      TabIndex        =   11
      Tag             =   "2"
      ToolTipText     =   "Notes 60 Chars (Optional)"
      Top             =   2880
      Width           =   4695
   End
   Begin VB.ComboBox txtDte2 
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Tag             =   "4"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtDrw2 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Tag             =   "3"
      ToolTipText     =   "Drawing (Optional)"
      Top             =   1800
      Width           =   2895
   End
   Begin VB.ComboBox cmbTag2 
      Height          =   315
      Left            =   4440
      Sorted          =   -1  'True
      TabIndex        =   6
      Tag             =   "3"
      ToolTipText     =   "Select Or Enter Tag Number (Optional)"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtCmt2 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Tag             =   "2"
      ToolTipText     =   "Notes 60 Chars (Optional)"
      Top             =   2160
      Width           =   4695
   End
   Begin VB.TextBox txtCmt1 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Tag             =   "2"
      ToolTipText     =   "Notes 60 Chars (Optional)"
      Top             =   1440
      Width           =   4695
   End
   Begin VB.ComboBox cmbTag1 
      Height          =   315
      Left            =   4440
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Select Or Enter Tag Number (Optional)"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txtDrw1 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Drawing (Optional)"
      Top             =   1080
      Width           =   2895
   End
   Begin VB.ComboBox txtDte1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Tag             =   "4"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5760
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6480
      Top             =   5160
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5460
      FormDesignWidth =   6735
   End
   Begin VB.Label z2 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Sources For:"
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   36
      Top             =   260
      Width           =   1575
   End
   Begin VB.Label z2 
      BackStyle       =   0  'Transparent
      Caption         =   "Notes:"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   35
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label z2 
      BackStyle       =   0  'Transparent
      Caption         =   "Notes:"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   34
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label z2 
      BackStyle       =   0  'Transparent
      Caption         =   "Notes:"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   33
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label z2 
      BackStyle       =   0  'Transparent
      Caption         =   "Notes:"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   32
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label z2 
      BackStyle       =   0  'Transparent
      Caption         =   "Notes:"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   31
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label z2 
      BackStyle       =   0  'Transparent
      Caption         =   "Notes:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   30
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label z2 
      BackStyle       =   0  'Transparent
      Caption         =   "Inspection Report         "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4440
      TabIndex        =   29
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label z2 
      BackStyle       =   0  'Transparent
      Caption         =   "Drawing                                                   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   28
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label z2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date                "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   27
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblKey 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3240
      TabIndex        =   26
      Top             =   500
      Width           =   1335
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   240
      TabIndex        =   25
      Top             =   500
      Width           =   2895
   End
End
Attribute VB_Name = "StatSPe01c"
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
Dim AdoSrc As ADODB.Recordset

Dim bGoodScr As Byte
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbTag1_LostFocus()
   cmbTag1 = CheckLen(cmbTag1, 12)
   If bGoodScr = 1 Then
      On Error Resume Next
      AdoSrc!DATTAG1 = cmbTag1
      AdoSrc.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub cmbTag2_LostFocus()
   cmbTag2 = CheckLen(cmbTag2, 12)
   If bGoodScr = 1 Then
      On Error Resume Next
      AdoSrc!DATTAG2 = cmbTag2
      AdoSrc.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub cmbTag3_LostFocus()
   cmbTag3 = CheckLen(cmbTag3, 12)
   If bGoodScr = 1 Then
      On Error Resume Next
      AdoSrc!DATTAG3 = cmbTag3
      AdoSrc.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub cmbTag4_LostFocus()
   cmbTag4 = CheckLen(cmbTag4, 12)
   If bGoodScr = 1 Then
      On Error Resume Next
      AdoSrc!DATTAG4 = cmbTag4
      AdoSrc.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub cmbTag5_LostFocus()
   cmbTag5 = CheckLen(cmbTag5, 12)
   If bGoodScr = 1 Then
      On Error Resume Next
      AdoSrc!DATTAG5 = cmbTag5
      AdoSrc.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub cmbTag6_LostFocus()
   cmbTag6 = CheckLen(cmbTag6, 12)
   If bGoodScr = 1 Then
      On Error Resume Next
      AdoSrc!DATTAG6 = cmbTag6
      AdoSrc.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 6301
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
   Move StatSPe01a.Left + 400, StatSPe01a.Top + 400
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   StatSPe01a.optDsc.value = vbUnchecked
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set AdoSrc = Nothing
   Set StatSPe01c = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   Dim AdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   AddComboStr cmbTag1.hwnd, "NONE"
   AddComboStr cmbTag2.hwnd, "NONE"
   AddComboStr cmbTag3.hwnd, "NONE"
   AddComboStr cmbTag4.hwnd, "NONE"
   AddComboStr cmbTag5.hwnd, "NONE"
   AddComboStr cmbTag6.hwnd, "NONE"
   
   sSql = "Qry_FillRejectionTags"
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoCmb, ES_FORWARD)
   If bSqlRows Then
      With AdoCmb
         Do Until .EOF
            AddComboStr cmbTag1.hwnd, "" & Trim(!REJNUM)
            AddComboStr cmbTag2.hwnd, "" & Trim(!REJNUM)
            AddComboStr cmbTag3.hwnd, "" & Trim(!REJNUM)
            AddComboStr cmbTag4.hwnd, "" & Trim(!REJNUM)
            AddComboStr cmbTag5.hwnd, "" & Trim(!REJNUM)
            AddComboStr cmbTag6.hwnd, "" & Trim(!REJNUM)
            .MoveNext
         Loop
         ClearResultSet AdoCmb
      End With
   End If
   cmbTag1 = cmbTag1.List(0)
   cmbTag2 = cmbTag2.List(0)
   cmbTag3 = cmbTag3.List(0)
   cmbTag4 = cmbTag4.List(0)
   cmbTag5 = cmbTag5.List(0)
   cmbTag6 = cmbTag6.List(0)
   bGoodScr = GetSource()
   Set AdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtCmt1_LostFocus()
   txtCmt1 = CheckLen(txtCmt1, 60)
   If bGoodScr = 1 Then
      On Error Resume Next
      AdoSrc!DATNOTE1 = txtCmt1
      AdoSrc.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtCmt2_LostFocus()
   txtCmt2 = CheckLen(txtCmt2, 60)
   If bGoodScr = 1 Then
      On Error Resume Next
      AdoSrc!DATNOTE2 = txtCmt2
      AdoSrc.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtCmt3_LostFocus()
   txtCmt3 = CheckLen(txtCmt3, 60)
   If bGoodScr = 1 Then
      On Error Resume Next
      AdoSrc!DATNOTE3 = txtCmt3
      AdoSrc.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtCmt4_LostFocus()
   txtCmt4 = CheckLen(txtCmt4, 60)
   If bGoodScr = 1 Then
      On Error Resume Next
      AdoSrc!DATNOTE4 = txtCmt4
      AdoSrc.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtCmt5_LostFocus()
   txtCmt5 = CheckLen(txtCmt5, 60)
   If bGoodScr = 1 Then
      On Error Resume Next
      AdoSrc!DATNOTE5 = txtCmt5
      AdoSrc.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtCmt6_LostFocus()
   txtCmt6 = CheckLen(txtCmt6, 60)
   If bGoodScr = 1 Then
      On Error Resume Next
      AdoSrc!DATNOTE6 = txtCmt6
      AdoSrc.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtDrw1_LostFocus()
   txtDrw1 = CheckLen(txtDrw1, 30)
   If bGoodScr = 1 Then
      On Error Resume Next
      AdoSrc!DATDRAW1 = txtDrw1
      AdoSrc.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtDrw2_LostFocus()
   txtDrw2 = CheckLen(txtDrw2, 30)
   If bGoodScr = 1 Then
      On Error Resume Next
      AdoSrc!DATDRAW2 = txtDrw2
      AdoSrc.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtDrw3_LostFocus()
   txtDrw3 = CheckLen(txtDrw3, 30)
   If bGoodScr = 1 Then
      On Error Resume Next
      AdoSrc!DATDRAW3 = txtDrw3
      AdoSrc.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtDrw4_LostFocus()
   txtDrw4 = CheckLen(txtDrw4, 30)
   If bGoodScr = 1 Then
      On Error Resume Next
      AdoSrc!DATDRAW4 = txtDrw4
      AdoSrc.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtDrw5_LostFocus()
   txtDrw5 = CheckLen(txtDrw5, 30)
   If bGoodScr = 1 Then
      On Error Resume Next
      AdoSrc!DATDRAW5 = txtDrw5
      AdoSrc.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtDrw6_LostFocus()
   txtDrw6 = CheckLen(txtDrw6, 30)
   If bGoodScr = 1 Then
      On Error Resume Next
      AdoSrc!DATDRAW6 = txtDrw6
      AdoSrc.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtDte1_DropDown()
   If Trim(txtDte1) = "" Then txtDte1 = Format(ES_SYSDATE, "mm/dd/yy")
   ShowCalendar Me
   
End Sub


Private Sub txtDte1_LostFocus()
   If Len(Trim(txtDte1)) Then txtDte1 = CheckDate(txtDte1)
   If bGoodScr = 1 Then
      On Error Resume Next
      If Len(Trim(txtDte1)) = 8 Then
         AdoSrc!DATDATE1 = Format(txtDte1, "mm/dd/yy")
         AdoSrc.Update
      Else
         AdoSrc!DATDATE1 = Null
         AdoSrc.Update
      End If
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtDte2_DropDown()
   If Trim(txtDte2) = "" Then txtDte2 = Format(ES_SYSDATE, "mm/dd/yy")
   ShowCalendar Me
   
End Sub


Private Sub txtDte2_LostFocus()
   If Len(Trim(txtDte2)) Then txtDte2 = CheckDate(txtDte2)
   If bGoodScr = 1 Then
      On Error Resume Next
      If Len(Trim(txtDte2)) = 8 Then
         AdoSrc!DATDATE2 = Format(txtDte2, "mm/dd/yy")
         AdoSrc.Update
      Else
         AdoSrc!DATDATE2 = Null
         AdoSrc.Update
      End If
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtDte3_DropDown()
   If Trim(txtDte3) = "" Then txtDte3 = Format(ES_SYSDATE, "mm/dd/yy")
   ShowCalendar Me
   
End Sub


Private Sub txtDte3_LostFocus()
   If Len(Trim(txtDte3)) Then txtDte3 = CheckDate(txtDte3)
   If bGoodScr = 1 Then
      On Error Resume Next
      If Len(Trim(txtDte3)) = 8 Then
         AdoSrc!DATDATE3 = Format(txtDte3, "mm/dd/yy")
         AdoSrc.Update
      Else
         AdoSrc!DATDATE3 = Null
         AdoSrc.Update
      End If
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtDte4_DropDown()
   If Trim(txtDte4) = "" Then txtDte4 = Format(ES_SYSDATE, "mm/dd/yy")
   ShowCalendar Me
   
End Sub


Private Sub txtDte4_LostFocus()
   If Len(Trim(txtDte4)) Then txtDte4 = CheckDate(txtDte4)
   If bGoodScr = 1 Then
      On Error Resume Next
      If Len(Trim(txtDte4)) = 8 Then
         AdoSrc!DATDATE4 = Format(txtDte4, "mm/dd/yy")
         AdoSrc.Update
      Else
         AdoSrc!DATDATE4 = Null
         AdoSrc.Update
      End If
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtDte5_DropDown()
   If Trim(txtDte5) = "" Then txtDte5 = Format(ES_SYSDATE, "mm/dd/yy")
   ShowCalendar Me
   
End Sub


Private Sub txtDte5_LostFocus()
   If Len(Trim(txtDte5)) Then txtDte5 = CheckDate(txtDte5)
   If bGoodScr = 1 Then
      On Error Resume Next
      If Len(Trim(txtDte5)) = 8 Then
         AdoSrc!DATDATE5 = Format(txtDte5, "mm/dd/yy")
         AdoSrc.Update
      Else
         AdoSrc!DATDATE5 = Null
         AdoSrc.Update
      End If
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtDte6_DropDown()
   If Trim(txtDte6) = "" Then txtDte6 = Format(ES_SYSDATE, "mm/dd/yy")
   ShowCalendar Me
   
End Sub


Private Sub txtDte6_LostFocus()
   If Len(Trim(txtDte6)) Then txtDte6 = CheckDate(txtDte6)
   If bGoodScr = 1 Then
      On Error Resume Next
      If Len(Trim(txtDte6)) = 8 Then
         AdoSrc!DATDATE6 = Format(txtDte6, "mm/dd/yy")
         AdoSrc.Update
      Else
         AdoSrc!DATDATE6 = Null
         AdoSrc.Update
      End If
      If Err > 0 Then ValidateEdit
   End If
   
End Sub



Private Function GetSource() As Byte
   Static bByte As Byte
   Dim sMsg As String
   bByte = bByte + 1
   sSql = "SELECT * FROM RjdtTable WHERE " _
          & "DATREF='" & Compress(lblPrt) & "' AND " _
          & "DATKEY='" & lblKey & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoSrc, ES_KEYSET)
   If bSqlRows Then
      With AdoSrc
         txtDte1 = Format(!DATDATE1, "mm/dd/yy")
         txtDrw1 = "" & Trim(!DATDRAW1)
         cmbTag1 = "" & Trim(!DATTAG1)
         txtCmt1 = "" & Trim(!DATNOTE1)
         
         txtDte2 = Format(!DATDATE2, "mm/dd/yy")
         txtDrw2 = "" & Trim(!DATDRAW2)
         cmbTag2 = "" & Trim(!DATTAG2)
         txtCmt2 = "" & Trim(!DATNOTE2)
         
         txtDte3 = Format(!DATDATE3, "mm/dd/yy")
         txtDrw3 = "" & Trim(!DATDRAW3)
         cmbTag3 = "" & Trim(!DATTAG3)
         txtCmt3 = "" & Trim(!DATNOTE3)
         
         txtDte4 = Format(!DATDATE4, "mm/dd/yy")
         txtDrw4 = "" & Trim(!DATDRAW4)
         cmbTag4 = "" & Trim(!DATTAG4)
         txtCmt4 = "" & Trim(!DATNOTE4)
         
         txtDte5 = Format(!DATDATE5, "mm/dd/yy")
         txtDrw5 = "" & Trim(!DATDRAW5)
         cmbTag5 = "" & Trim(!DATTAG5)
         txtCmt5 = "" & Trim(!DATNOTE5)
         
         txtDte6 = Format(!DATDATE6, "mm/dd/yy")
         txtDrw6 = "" & Trim(!DATDRAW6)
         cmbTag6 = "" & Trim(!DATTAG6)
         txtCmt6 = "" & Trim(!DATNOTE6)
         
         txtDte1 = Format(!DATDATE1, "mm/dd/yy")
         txtDrw1 = "" & Trim(!DATDRAW1)
         cmbTag1 = "" & Trim(!DATTAG1)
         txtCmt1 = "" & Trim(!DATNOTE1)
         GetSource = 1
      End With
      If cmbTag1 = "" Then cmbTag1 = "NONE"
      If cmbTag2 = "" Then cmbTag2 = "NONE"
      If cmbTag3 = "" Then cmbTag3 = "NONE"
      If cmbTag4 = "" Then cmbTag4 = "NONE"
      If cmbTag5 = "" Then cmbTag5 = "NONE"
      If cmbTag6 = "" Then cmbTag6 = "NONE"
      
   Else
      On Error Resume Next
      sSql = "INSERT INTO RjdtTable (DATREF,DATKEY) " _
             & "VALUES('" & Compress(lblPrt) & "','" & lblKey & "')"
      clsADOCon.ExecuteSQL sSql
      If Err > 0 Then
         sMsg = Trim(str(Err.Number)) & " " & Left(Err.Description, 30) & vbCr
         GetSource = 0
         MsgBox sMsg & "Couldn't Add The Data Sources For Key.", _
            vbExclamation, Caption
         Unload Me
         Exit Function
      End If
   End If
   If GetSource = 0 Then
      If bByte = 1 Then bGoodScr = GetSource()
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getsource"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function
