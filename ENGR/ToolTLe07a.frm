VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form ToolTLe07a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Tool to Part"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   3402
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPartFil 
      Caption         =   "&FilterPart"
      Height          =   315
      Left            =   3480
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Update The Current Tool List To Contain These Tools"
      Top             =   2160
      Width           =   870
   End
   Begin VB.TextBox txtPartFilter 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Tag             =   "2"
      ToolTipText     =   "Description - Up To (30) Chars"
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ToolTLe07a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "ToolTLe07a.frx":07AE
      Height          =   350
      Left            =   5760
      Picture         =   "ToolTLe07a.frx":0C88
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Show The Tool List"
      Top             =   1440
      Width           =   350
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Tag             =   "2"
      ToolTipText     =   "Description - Up To (30) Chars"
      Top             =   960
      Width           =   3135
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6240
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Update The Current Tool List To Contain These Tools"
      Top             =   1440
      Width           =   870
   End
   Begin VB.ListBox lstCur 
      Height          =   2205
      ItemData        =   "ToolTLe07a.frx":1162
      Left            =   4080
      List            =   "ToolTLe07a.frx":1164
      TabIndex        =   7
      ToolTipText     =   "Current Assignments - Double Click To Change Quantities"
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CommandButton cmdLst 
      Caption         =   "<=>"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Highlite Selection And Press To Move"
      Top             =   3600
      Width           =   495
   End
   Begin VB.ListBox lstPart 
      Height          =   2205
      ItemData        =   "ToolTLe07a.frx":1166
      Left            =   360
      List            =   "ToolTLe07a.frx":116D
      Sorted          =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "Right Mouse Button For Description Of Selection"
      Top             =   2760
      Width           =   3015
   End
   Begin VB.ComboBox cmbCls 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "8"
      ToolTipText     =   "Select Class From List"
      Top             =   1440
      Width           =   2000
   End
   Begin VB.ComboBox cmbLst 
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Tool List Number"
      Top             =   600
      Width           =   3345
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6240
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   4560
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5220
      FormDesignWidth =   7305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Parts Filter"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   16
      Top             =   2160
      Width           =   975
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   360
      X2              =   3360
      Y1              =   2865
      Y2              =   2865
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Parts"
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   13
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Parts"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   12
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Class"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Name"
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "ToolTLe07a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
Option Explicit
Dim RdoLst As ADODB.Recordset
Dim bDataChg As Byte
Dim bCancel As Byte
Dim bOnLoad As Byte
Dim bGoodList As Byte

Dim iCurList As Integer
Dim iTolList As Integer
Dim sOldClass As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCls_Click()
   Dim bResponse As Byte
   If bDataChg Then
      bResponse = MsgBox("The Data Has Changed. Save First?", _
                  ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         cmbCls = sOldClass
         Exit Sub
      End If
   End If
   FillCurrentParts
   
End Sub

Private Sub cmbLst_Click()
   FillCurrentParts    'BBS Fixed this on 9/3/2010 for Ticket #37441
   bGoodList = GetThisTool()
    
End Sub


Private Sub cmbLst_LostFocus()
   cmbLst = CheckLen(cmbLst, 30)
   If bCancel = 0 Then
      bGoodList = GetThisTool()
      If bGoodList = 0 Then MsgBox "The Tool number is not found.", vbInformation, Caption
      
   End If
End Sub


Private Sub cmdCan_Click()
   Dim bResponse As Byte
   If bDataChg Then
      bResponse = MsgBox("The Part tool list has Changed. Save First?", _
                  ES_YESQUESTION, Caption)
      If bResponse = vbYes Then Exit Sub
   End If
   Unload Me
   
End Sub



Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = 1
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3402
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdLst_Click()
   If iCurList >= 0 Then
      lstPart.AddItem lstCur.List(iCurList)
      lstCur.RemoveItem (iCurList)
   Else
      If iTolList >= 0 Then
         lstCur.AddItem lstPart.List(iTolList)
         lstPart.RemoveItem (iTolList)
      End If
   End If
   cmdUpd.Enabled = True
   cmdLst.Enabled = False
   bDataChg = 1
   
End Sub

Private Sub cmdPartFil_Click()
GetAllParts
End Sub

Private Sub cmdUpd_Click()
   Dim iList As Integer
   Dim bResponse As Byte
   bResponse = MsgBox("Update The List With The Current Class?", _
               ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      Err = 0
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      
      sSql = "DELETE FROM TlitTableNew where (TOOL_CLASS ='" _
             & Trim(cmbCls) & "' AND TOOL_NUMREF ='" _
             & Compress(cmbLst) & "')"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      
      For iList = 0 To lstCur.ListCount - 1
         sSql = "INSERT INTO TlitTableNew (TOOL_NUMREF," _
                & "TOOL_PARTREF,TOOL_CLASS) " _
                & " VALUES('" & Compress(cmbLst) & "','" _
                & Compress(lstCur.List(iList)) & "','" _
                & Trim(cmbCls) & "')"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
      Next
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         SysMsg "List Updated.", True
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         MsgBox "Couldn't Update The List.", _
            vbInformation, Caption
      End If
      bDataChg = 0
   Else
      CancelTrans
   End If
   
End Sub

Private Sub cmdVew_Click()
   ViewTool.lblLst = cmbLst
   ViewTool.Show
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then FillCombo
   txtPartFilter = ""
   bOnLoad = 0
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
   Set RdoLst = Nothing
   Set ToolTLe07a = Nothing
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT TOOL_NUM FROM TlnhdTableNew ORDER BY TOOL_NUM "
   LoadComboBox cmbLst, -1
   
   sSql = "SELECT DISTINCT TOOL_CLASS FROM TlnhdTableNew ORDER BY TOOL_CLASS "
   LoadComboBox cmbCls, -1
   If cmbLst.ListCount > 0 Then cmbLst = cmbLst.List(0)
   If cmbCls.ListCount > 0 Then
      cmbCls = cmbCls.List(0)
      sOldClass = cmbCls
      FillCurrentParts
   Else
      MsgBox "No Tool Classes Found (Required) .", _
         vbInformation, Caption
      cmbCls.Enabled = False
      lstPart.Enabled = False
      lstCur.Enabled = False
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetThisTool() As Byte
   On Error GoTo DiaErr1
   
   sSql = "SELECT TOOL_NUM,TOOL_CLASS  FROM TlnhdTableNew where TOOL_NUMREF = '" & Compress(cmbLst) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_KEYSET)
   
   If bSqlRows Then
      With RdoLst
         cmbLst = "" & Trim(!TOOL_NUM)
         cmbCls = "" & Trim(!TOOL_CLASS)
         GetThisTool = 1
      End With
   Else
      GetThisTool = 0
   End If
   
   Exit Function
   
DiaErr1:
   sProcName = "GetThisTool"
   GetThisTool = 0
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub GetAllParts()
   Dim bByte As Byte
   Dim iList As Integer
   Dim iRow As Integer
   Dim filter As String
   Dim RdoTls As ADODB.Recordset
   lstPart.Clear
   
   filter = txtPartFilter
   If (filter <> "") Then
      sSql = "SELECT DISTINCT TOP(32767) PartNum FROM PartTable WHERE Partref like '%" + filter + "%'" _
             & " AND paobsolete = 0 and PAINACTIVE = 0 ORDER BY PartNum"
   
   Else
      sSql = "SELECT DISTINCT TOP(32767) PartNum FROM PartTable WHERE " _
             & "paobsolete = 0 and PAINACTIVE = 0 ORDER BY PartNum"
   End If
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTls, ES_FORWARD)
   If bSqlRows Then
      With RdoTls
         Do Until .EOF
            bByte = 0
            For iList = 0 To lstCur.ListCount - 1
               If Trim(lstCur.List(iList)) = Trim(!PartNum) Then
                  bByte = 1
                  Exit For
               End If
            Next
            If bByte = 0 Then lstPart.AddItem Trim(!PartNum)
            .MoveNext
         Loop
         ClearResultSet RdoTls
      End With
   End If
   Set RdoTls = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "GetAllParts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub lstCur_Click()
   On Error Resume Next
   iTolList = -1
   cmdUpd.Enabled = False
   If lstCur.Selected(lstCur.ListIndex) Then
      iCurList = lstCur.ListIndex
      cmdLst.Enabled = True
   Else
      cmdLst.Enabled = False
   End If
   
End Sub

Private Sub lstPart_Click()
   On Error Resume Next
   iCurList = -1
   cmdUpd.Enabled = False
   If lstPart.Selected(lstPart.ListIndex) Then
      iTolList = lstPart.ListIndex
      cmdLst.Enabled = True
   Else
      iTolList = -1
      cmdLst.Enabled = False
   End If
   
End Sub

Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 30)
   txtDsc = StrCase(txtDsc)
   If bGoodList Then
      On Error Resume Next
      With RdoLst
         '.Edit
         !TOOLLIST_DESC = Trim(txtDsc)
         .Update
      End With
   End If
   
End Sub



Private Sub FillCurrentParts()
   Dim RdoCur As ADODB.Recordset
   
   lstCur.Clear
   sSql = "SELECT TOOL_PARTREF " _
          & "FROM TlitTableNew where (TOOL_CLASS ='" _
          & Trim(cmbCls) & "' AND TOOL_NUMREF ='" & Compress(cmbLst) _
          & "') ORDER BY TOOL_NUMREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCur, ES_FORWARD)
   If bSqlRows Then
      With RdoCur
         Do Until .EOF
            lstCur.AddItem Trim(!TOOL_PARTREF)
            .MoveNext
         Loop
         ClearResultSet RdoCur
      End With
   End If
   sOldClass = cmbCls
   Set RdoCur = Nothing
   GetAllParts
   Exit Sub
   
DiaErr1:
   sProcName = "FillCurrentParts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


