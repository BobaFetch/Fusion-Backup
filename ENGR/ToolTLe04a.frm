VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form ToolTLe04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Assign A Tool List To A Manufacturing Order"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ToolTLe04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbRun 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   6240
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "8"
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox lblCurList 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   "2"
      Text            =   " "
      Top             =   3240
      Width           =   3075
   End
   Begin VB.TextBox lblLst 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   "2"
      Text            =   " "
      Top             =   3960
      Width           =   3075
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   6240
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Update The Selected Manufacturing Order To Include The Tool List"
      Top             =   3600
      Width           =   870
   End
   Begin VB.ComboBox cmbLst 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Select Tool List Number"
      Top             =   3600
      Width           =   3345
   End
   Begin VB.ComboBox cmbOps 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   6240
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "8"
      Top             =   1200
      Width           =   855
   End
   Begin VB.ComboBox cmbRte 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Add/Edit Routing"
      Top             =   810
      WhatsThisHelpID =   100
      Width           =   3345
   End
   Begin VB.TextBox txtDsc 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   "2"
      Text            =   " "
      Top             =   1170
      Width           =   3075
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
      Top             =   4440
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4530
      FormDesignWidth =   7185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Runs"
      Height          =   255
      Index           =   8
      Left            =   5040
      TabIndex        =   20
      Top             =   840
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Tool List"
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   19
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Lists"
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   15
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label lblComt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Left            =   1560
      TabIndex        =   14
      Top             =   1920
      Width           =   4305
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCenter 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5160
      TabIndex        =   13
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblShop 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   12
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center"
      Height          =   285
      Index           =   4
      Left            =   3720
      TabIndex        =   10
      Top             =   1560
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shop"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operations"
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   8
      Top             =   1200
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   1170
      Width           =   915
   End
End
Attribute VB_Name = "ToolTLe04a"
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
Dim bCancel As Byte
Dim bGoodMO As Byte
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbLst_Click()
   If Trim(cmbLst) = "< NONE >" Then _
           lblLst = "Set Routing Operation To No Tool List" Else _
           cmbLst = FindToolList(cmbLst, lblLst)
   
   
End Sub


Private Sub cmbLst_LostFocus()
   Dim bByte As Byte
   Dim iRow As Integer
   For iRow = 0 To cmbLst.ListCount - 1
      If cmbLst = cmbLst.list(iRow) Then bByte = 1
   Next
   If bByte = 0 Then
      Beep
      cmbLst = cmbLst.list(0)
      lblLst = "Set Routing Operation To No Tool List"
   End If
   
End Sub


Private Sub cmbOps_Click()
   If cmbOps.ListCount > 0 Then GetThisOperation
End Sub

Private Sub cmbRte_Click()
   bGoodMO = GetThisMo()
End Sub


Private Sub cmbRte_LostFocus()
   cmbRte = CheckLen(cmbRte, 30)
   If bCancel = 0 Then bGoodMO = GetThisMo()
End Sub


Private Sub cmbRun_Click()
   FillOperations
   
End Sub


Private Sub cmbRun_LostFocus()
   If cmbRun.ListCount > 0 Then FillOperations
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = 1
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3404
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdUpd_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sList As String
   
   If cmbLst = "" Or cmbLst = "< NONE >" Then
      sList = ""
      sMsg = "Set This MO Operation Tool List To Blank (None)?"
   Else
      sMsg = "Set This MO Operation List To The Selected List?"
      sList = Compress(cmbLst)
   End If
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      sSql = "UPDATE RnopTable SET OPTOOLLIST='" & sList _
             & "' WHERE (OPREF='" & Compress(cmbRte) & "' AND " _
             & "OPRUN=" & Val(cmbRun) & " AND " _
             & "OPNO=" & Val(cmbOps) & ")"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      'If sList <> "" Then lblCurList = sList
      lblCurList = cmbLst
      
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
   FormLoad Me
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set ToolTLe04a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDsc.BackColor = BackColor
   lblLst.BackColor = BackColor
   lblCurList.BackColor = BackColor
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_FillToolListCombo"
'   AddComboStr cmbLst.hwnd, "< NONE >"
   LoadComboBox cmbLst
   AddComboStr cmbLst.hwnd, "< NONE >"
   cmbLst.AddItem "< NONE >", 0
   If cmbLst.ListCount > 0 Then
      cmbLst = cmbLst.list(0)
      lblLst = "Set Routing Operation To No Tool List"
   End If
   sSql = "Qry_RunsNotLikeC"
   LoadComboBox cmbRte
   If cmbRte.ListCount > 0 Then
      cmbRte = cmbRte.list(0)
      bGoodMO = GetThisMo()
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetThisMo() As Byte
   Dim RdoRte As ADODB.Recordset
   cmbOps.Clear
   'cmdUpd.Enabled = False
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT RUNREF,PARTREF,PARTNUM,PADESC FROM RunsTable," _
          & "PartTable WHERE (RUNREF=PARTREF AND RUNREF='" & Compress(cmbRte) & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_FORWARD)
   If bSqlRows Then
      With RdoRte
         cmbRte = "" & Trim(!PartNum)
         txtDsc = "" & Trim(!PADESC)
         GetThisMo = 1
         ClearResultSet RdoRte
      End With
   Else
      GetThisMo = 0
      txtDsc = "*** MO Wasn't Found ***"
   End If
   Set RdoRte = Nothing
   If GetThisMo = 1 Then FillRuns
   Exit Function
   
DiaErr1:
   sProcName = "getthismo"
   GetThisMo = 0
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub txtDsc_Change()
   If Left(txtDsc, 6) = "*** MO" Then txtDsc.ForeColor = _
           ES_RED Else txtDsc.ForeColor = vbBlack
   
End Sub



Private Sub FillOperations()
   Dim RdoOps As ADODB.Recordset
   On Error GoTo DiaErr1
   cmbOps.Clear
   sSql = "SELECT OPREF,OPRUN,OPNO FROM RnopTable WHERE " _
          & "(OPREF='" & Compress(cmbRte) & "' AND OPRUN=" _
          & Val(cmbRun) & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoOps, ES_FORWARD)
   If bSqlRows Then
      With RdoOps
         Do Until .EOF
            cmbOps.AddItem Format$(!OPNO, "000")
            .MoveNext
         Loop
         ClearResultSet RdoOps
      End With
   Else
      lblShop = ""
      lblCenter = ""
      lblComt = ""
      lblCurList = ""
   End If
   If cmbOps.ListCount > 0 Then
      cmbOps = cmbOps.list(0)
      GetThisOperation
   End If
   Set RdoOps = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "filloperations"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub GetThisOperation()
   Dim RdoOpn As ADODB.Recordset
   sProcName = "getthisopera"
   sSql = "SELECT OPREF,OPSHOP,OPCENTER,OPTOOLLIST,OPCOMT,SHPREF,SHPNUM," _
          & "WCNREF,WCNNUM FROM RnopTable,ShopTable,WcntTable " _
          & "WHERE (OPREF='" & Compress(cmbRte) & "' " _
          & "AND OPRUN=" & Val(cmbRun) & " AND OPNO=" _
          & Val(cmbOps) & ") AND OPSHOP=SHPREF AND OPCENTER=WCNREF "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoOpn, ES_STATIC)
   If bSqlRows Then
      With RdoOpn
         lblShop = "" & Trim(!SHPNUM)
         lblCenter = "" & Trim(!WCNNUM)
         lblComt = "" & Trim(!OPCOMT)
         lblCurList = "" & Trim(!OPTOOLLIST)
         lblCurList = FindToolList(lblCurList, txtDsc, 1)
         ClearResultSet RdoOpn
      End With
   End If
'   If cmbLst.ListCount > 0 Then cmdUpd.Enabled = True _
'                                                 Else cmdUpd.Enabled = False
   Set RdoOpn = Nothing
   
End Sub


Private Sub FillRuns()
   Dim RdoRun As ADODB.Recordset
   cmbRun.Clear
   cmbOps.Clear
   lblShop = ""
   lblCenter = ""
   lblComt = ""
   lblCurList = ""
   On Error GoTo DiaErr1
   sSql = "SELECT RUNREF,RUNNO FROM RunsTable WHERE " _
          & "RUNREF='" & Compress(cmbRte) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRun, ES_FORWARD)
   If bSqlRows Then
      With RdoRun
         Do Until .EOF
            AddComboStr cmbRun.hwnd, Format$(!Runno, "#####0")
            .MoveNext
         Loop
         ClearResultSet RdoRun
      End With
   End If
   If cmbRun.ListCount > 0 Then
      cmbRun = cmbRun.list(0)
      FillOperations
   End If
   Set RdoRun = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
