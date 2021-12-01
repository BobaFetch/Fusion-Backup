VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form ToolTLe02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tool Lists"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   3402
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Search Available Tools for"
      Height          =   1095
      Left            =   1200
      TabIndex        =   14
      Top             =   1920
      Width           =   5535
      Begin VB.ComboBox cmbCls 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   17
         Tag             =   "8"
         ToolTipText     =   "Select Class From List"
         Top             =   240
         Width           =   2000
      End
      Begin VB.TextBox txtSearchDoc 
         Height          =   285
         Left            =   1080
         TabIndex        =   15
         ToolTipText     =   "Type in a partial document number to narrow down your search"
         Top             =   660
         Width           =   4215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tool Class"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   300
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Number"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   660
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ToolTLe02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "ToolTLe02a.frx":07AE
      Height          =   350
      Left            =   5760
      Picture         =   "ToolTLe02a.frx":0C88
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Show The Tool List"
      Top             =   1440
      Width           =   350
   End
   Begin VB.ListBox lstQty 
      Height          =   2205
      ItemData        =   "ToolTLe02a.frx":1162
      Left            =   6360
      List            =   "ToolTLe02a.frx":1169
      TabIndex        =   10
      ToolTipText     =   "Current Quantity Assigned"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1560
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
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Update The Current Tool List To Contain These Tools"
      Top             =   1440
      Width           =   870
   End
   Begin VB.ListBox lstCur 
      Height          =   2205
      ItemData        =   "ToolTLe02a.frx":1175
      Left            =   3960
      List            =   "ToolTLe02a.frx":1177
      TabIndex        =   3
      ToolTipText     =   "Current Assignments - Double Click To Change Quantities"
      Top             =   3360
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
      Left            =   3360
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Highlite Selection And Press To Move"
      Top             =   4200
      Width           =   495
   End
   Begin VB.ListBox lstTools 
      Height          =   2205
      ItemData        =   "ToolTLe02a.frx":1179
      Left            =   240
      List            =   "ToolTLe02a.frx":1180
      Sorted          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Right Mouse Button For Description Of Selection"
      Top             =   3360
      Width           =   3015
   End
   Begin VB.ComboBox cmbLst 
      Height          =   315
      Left            =   1560
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   5400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5790
      FormDesignWidth =   7635
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Description"
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   19
      Top             =   960
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      Height          =   255
      Index           =   4
      Left            =   6360
      TabIndex        =   11
      Top             =   3120
      Width           =   855
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   3960
      X2              =   7080
      Y1              =   3465
      Y2              =   3465
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   240
      X2              =   3240
      Y1              =   3465
      Y2              =   3465
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Tools"
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   9
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tools"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool List"
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "ToolTLe02a"
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
   FillCurrentTools
   
End Sub

Private Sub cmbCls_LostFocus()
   cmbCls_Click
End Sub

Private Sub cmbLst_Change()
   FillCurrentTools
End Sub

Private Sub cmbLst_Click()
   FillCurrentTools    'BBS Fixed this on 9/3/2010 for Ticket #37441
   bGoodList = GetThisToolList()
    
End Sub


Private Sub cmbLst_LostFocus()
   cmbLst = CheckLen(cmbLst, 30)
   If bCancel = 0 Then
      FillCurrentTools
      bGoodList = GetThisToolList()
      If bGoodList = 0 Then AddToolList
   End If
End Sub


Private Sub cmdCan_Click()
   Dim bResponse As Byte
   If bDataChg Then
      bResponse = MsgBox("The Tool List Data Has Changed. Save First?", _
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
      lstTools.AddItem lstCur.list(iCurList)
      lstCur.RemoveItem (iCurList)
      lstQty.RemoveItem (iCurList)
   Else
      If iTolList >= 0 Then
         lstCur.AddItem lstTools.list(iTolList)
         lstQty.AddItem "1"
         lstTools.RemoveItem (iTolList)
      End If
   End If
   cmdUpd.Enabled = True
   cmdLst.Enabled = False
   bDataChg = 1
   
End Sub

Private Sub cmdLst_KeyUp(KeyCode As Integer, Shift As Integer)
   FillCurrentTools
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
      
      sSql = "DELETE FROM TlitTable where (TOOLLISTIT_CLASS ='" _
             & Trim(cmbCls) & "' AND TOOLLISTIT_REF ='" _
             & Compress(cmbLst) & "')"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      
      For iList = 0 To lstCur.ListCount - 1
         sSql = "INSERT INTO TlitTable (TOOLLISTIT_REF," _
                & "TOOLLISTIT_NUM,TOOLLISTIT_TOOLREF," _
                & "TOOLLISTIT_CLASS,TOOLLISTIT_QUANTITYUSED) " _
                & "VALUES('" & Compress(cmbLst) & "','" _
                & Trim(cmbLst) & "','" _
                & Compress(lstCur.list(iList)) & "','" _
                & Trim(cmbCls) & "'," & Val(lstQty.list(iList)) & ")"
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
   Set ToolTLe02a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_FillToolListCombo"
   LoadComboBox cmbLst
   
   sSql = "Qry_FillToolClasses"
   LoadComboBox cmbCls, -1
   If cmbLst.ListCount > 0 Then cmbLst = cmbLst.list(0)
   If cmbCls.ListCount > 0 Then
      cmbCls = cmbCls.list(0)
      sOldClass = cmbCls
      FillCurrentTools
   Else
      MsgBox "No Tool Classes Found (Required) .", _
         vbInformation, Caption
      cmbCls.Enabled = False
      lstTools.Enabled = False
      lstCur.Enabled = False
      lstQty.Enabled = False
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetThisToolList() As Byte
   On Error GoTo DiaErr1
   
   sSql = "Qry_GetToolList '" & Compress(cmbLst) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_KEYSET)
   
   
   If bSqlRows Then
      With RdoLst
         cmbLst = "" & Trim(!TOOLLIST_NUM)
         txtDsc = "" & Trim(!TOOLLIST_DESC)
         GetThisToolList = 1
      End With
   Else
      GetThisToolList = 0
   End If
   
   
   
   
   Exit Function
   
DiaErr1:
   sProcName = "getthistoollist"
   GetThisToolList = 0
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AddToolList()
   Dim bResponse As Byte
   Dim sMsg As String
   On Error GoTo DiaErr1
   If Len(Trim(cmbLst)) < 3 Then Exit Sub
   sMsg = cmbLst & " Wasn't Found.  Add The Tool List?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      clsADOCon.ADOErrNum = 0
      sSql = "INSERT INTO TlhdTable (TOOLLIST_REF,TOOLLIST_NUM) " _
             & "VALUES('" & Compress(cmbLst) & "','" & cmbLst & "')"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      If clsADOCon.ADOErrNum = 0 Then
         txtDsc = ""
         SysMsg "The Tool List Was Added.", True
         lstTools.Clear
         lstCur.Clear
         bGoodList = GetThisToolList()
         FillCurrentTools   'BBS Fixed this on 9/3/2010 for Ticket #37441
      Else
         MsgBox "Couldn't Add Tool List.", _
            vbInformation, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "addtoollist"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetTools()
   Dim bByte As Byte
   Dim iList As Integer
   Dim iRow As Integer
   Dim RdoTls As ADODB.Recordset
   lstTools.Clear
   sSql = "SELECT TOOL_PARTREF,TOOL_NUM,TOOL_CLASS FROM TohdTable WHERE " _
          & "TOOL_CLASS='" & cmbCls & "'" & vbCrLf _
          & "AND TOOL_PARTREF like '" & Compress(txtSearchDoc.Text) & "%'" & vbCrLf _
          & "ORDER BY TOOL_PARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTls, ES_FORWARD)
   If bSqlRows Then
      With RdoTls
         Do Until .EOF
            bByte = 0
            For iList = 0 To lstCur.ListCount - 1
               If Trim(lstCur.list(iList)) = Trim(!TOOL_NUM) Then
                  bByte = 1
                  Exit For
               End If
            Next
            If bByte = 0 Then lstTools.AddItem Trim(!TOOL_NUM)
            .MoveNext
         Loop
         ClearResultSet RdoTls
      End With
   End If
   Set RdoTls = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "gettools"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub lstCur_Click()
   On Error Resume Next
   iTolList = -1
   cmdUpd.Enabled = False
   If lstCur.ListCount > 0 Then
      If lstCur.Selected(lstCur.ListIndex) Then
         iCurList = lstCur.ListIndex
         cmdLst.Enabled = True
      Else
         cmdLst.Enabled = False
      End If
   Else
      cmdLst.Enabled = False
      iCurList = -1
   End If
   
End Sub

Private Sub lstCur_DblClick()
   iCurList = lstCur.ListIndex
   ToolTLe02b.lblIndex = iCurList
   ToolTLe02b.lblTool = lstCur.list(iCurList)
   ToolTLe02b.txtQty = lstQty.list(iCurList)
   ToolTLe02b.Show
   
End Sub


Private Sub lstQty_DblClick()
   iCurList = lstQty.ListIndex
   ToolTLe02b.lblIndex = iCurList
   ToolTLe02b.lblTool = lstCur.list(iCurList)
   ToolTLe02b.txtQty = lstQty.list(iCurList)
   ToolTLe02b.Show
   
End Sub


Private Sub lstTools_Click()
   On Error Resume Next
   iCurList = -1
   cmdUpd.Enabled = False
   If lstTools.ListCount > 0 Then
      If lstTools.Selected(lstTools.ListIndex) Then
         iTolList = lstTools.ListIndex
         cmdLst.Enabled = True
      Else
         iTolList = -1
         cmdLst.Enabled = False
      End If
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



Private Sub FillCurrentTools()
   Dim RdoCur As ADODB.Recordset
   
   lstCur.Clear
   lstQty.Clear
   sSql = "SELECT TOOLLISTIT_REF,TOOLLISTIT_NUM,TOOLLISTIT_TOOLREF," _
          & "TOOLLISTIT_CLASS,TOOLLISTIT_QUANTITYUSED,TOOL_NUM,TOOL_PARTREF " _
          & "FROM TlitTable,TohdTable where (TOOLLISTIT_CLASS ='" _
          & Trim(cmbCls) & "' AND TOOLLISTIT_REF ='" & Compress(cmbLst) _
          & "') AND TOOLLISTIT_TOOLREF=TOOL_PARTREF" & vbCrLf _
          & "ORDER BY TOOL_NUM"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCur, ES_FORWARD)
   If bSqlRows Then
      With RdoCur
         Do Until .EOF
            lstCur.AddItem Trim(!TOOL_NUM)
            lstQty.AddItem !TOOLLISTIT_QUANTITYUSED
            .MoveNext
         Loop
         ClearResultSet RdoCur
      End With
   End If
   sOldClass = cmbCls
   Set RdoCur = Nothing
   GetTools
   Exit Sub
   
DiaErr1:
   sProcName = "FillCurrentTools"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub txtSearchDoc_Change()
   FillCurrentTools
End Sub

Private Sub txtSearchDoc_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
