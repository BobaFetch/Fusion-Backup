VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form ToolTLe03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Assign A Tool List To A Routing"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ToolTLe03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox lblCurList 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   "2"
      Text            =   " "
      Top             =   3120
      Width           =   3075
   End
   Begin VB.TextBox lblLst 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Tag             =   "2"
      Text            =   " "
      Top             =   3840
      Width           =   3075
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6120
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Update The Selected Routing To Include The Tool List"
      Top             =   3480
      Width           =   870
   End
   Begin VB.ComboBox cmbLst 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Select Tool List Number"
      Top             =   3480
      Width           =   3345
   End
   Begin VB.ComboBox cmbOps 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   6120
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "8"
      Top             =   720
      Width           =   855
   End
   Begin VB.ComboBox cmbRte 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Add/Edit Routing"
      Top             =   690
      WhatsThisHelpID =   100
      Width           =   3345
   End
   Begin VB.TextBox txtDsc 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Tag             =   "2"
      Text            =   " "
      Top             =   1050
      Width           =   3075
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6120
      TabIndex        =   3
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
      FormDesignHeight=   4410
      FormDesignWidth =   7065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Tool List"
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   18
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Lists"
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   14
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lblComt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Left            =   1560
      TabIndex        =   13
      Top             =   1800
      Width           =   4305
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCenter 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5160
      TabIndex        =   12
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblShop 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   11
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center"
      Height          =   285
      Index           =   4
      Left            =   3720
      TabIndex        =   9
      Top             =   1440
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shop"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   1440
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operations"
      Height          =   255
      Index           =   2
      Left            =   5160
      TabIndex        =   7
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Routing"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1050
      Width           =   915
   End
End
Attribute VB_Name = "ToolTLe03a"
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
Dim bGoodRouting As Byte
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
      ' MM cmbLst = cmbLst.List(0)
      lblLst = "Set Routing Operation To No Tool List"
   End If
   
End Sub


Private Sub cmbOps_Click()
   If cmbOps.ListCount > 0 Then GetThisOperation
End Sub

Private Sub cmbOps_LostFocus()
   If cmbOps.ListCount > 0 Then GetThisOperation
End Sub

Private Sub cmbRte_Click()
   bGoodRouting = GetRouting()
   
End Sub


Private Sub cmbRte_LostFocus()
   cmbRte = CheckLen(cmbRte, 30)
   If bCancel = 0 Then bGoodRouting = GetRouting()
   
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
      OpenHelpContext 3403
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdUpd_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sList As String
   
   If Trim(cmbLst) = "" Or cmbLst = "< NONE >" Then
      sList = ""
      sMsg = "Set This Routing Tool List To Blank (None)?"
   Else
      sMsg = "Set This Routing Tool List To The Selected List?"
      sList = Compress(cmbLst)
   End If
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      sSql = "UPDATE RtopTable SET OPTOOLLIST='" & sList _
             & "' WHERE OPREF='" & Compress(cmbRte) & "' AND " _
             & "OPNO=" & Val(cmbOps) & " "
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      If sList <> "" Then lblCurList = sList
      
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
   Set ToolTLe03a = Nothing
   
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
   LoadComboBox cmbLst
   If cmbLst.ListCount > 0 Then
      cmbLst = cmbLst.list(0)
      lblLst = "Set Routing Operation To No Tool List"
   End If
   FillRoutings
   If cmbRte.ListCount > 0 Then bGoodRouting = GetRouting()
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetRouting() As Byte
   Dim RdoRte As ADODB.Recordset
   cmbOps.Clear
   cmdUpd.Enabled = False
   On Error GoTo DiaErr1
   sSql = "SELECT RTREF,RTNUM,RTDESC FROM RthdTable " _
          & "WHERE RTREF='" & Compress(cmbRte) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_FORWARD)
   If bSqlRows Then
      With RdoRte
         cmbRte = "" & Trim(!RTNUM)
         txtDsc = "" & Trim(!RTDESC)
         GetRouting = 1
         ClearResultSet RdoRte
      End With
   Else
      GetRouting = 0
      txtDsc = "*** Routing Wasn't Found ***"
   End If
   Set RdoRte = Nothing
   If GetRouting = 1 Then FillOperations
   Exit Function
   
DiaErr1:
   sProcName = "getrouting"
   GetRouting = 0
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub lblLst_Change()
   If Trim(cmbLst) = "< NONE >" Then _
           lblLst = "Set Routing Operation To No Tool List"
   
End Sub

Private Sub txtDsc_Change()
   If Left(txtDsc, 6) = "*** Ro" Then txtDsc.ForeColor = _
           ES_RED Else txtDsc.ForeColor = vbBlack
   
End Sub



Private Sub FillOperations()
   Dim RdoOps As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT OPREF,OPNO FROM RtopTable WHERE " _
          & "OPREF='" & Compress(cmbRte) & "'"
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
          & "WCNREF,WCNNUM FROM RtopTable,ShopTable,WcntTable " _
          & "WHERE (OPREF='" & Compress(cmbRte) & "' AND OPNO=" _
          & Val(cmbOps) & ") AND OPSHOP=SHPREF AND OPCENTER=WCNREF "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoOpn, ES_STATIC)
   If bSqlRows Then
      With RdoOpn
         lblShop = "" & Trim(!SHPNUM)
         lblCenter = "" & Trim(!WCNNUM)
         lblComt = "" & Trim(!OPCOMT)
         lblCurList = "" & Trim(!OPTOOLLIST)
         lblCurList = FindToolList(lblCurList, "", 1)
         ClearResultSet RdoOpn
      End With
   End If
   If cmbLst.ListCount > 0 Then cmdUpd.Enabled = True _
                                                 Else cmdUpd.Enabled = False
   Set RdoOpn = Nothing
   
End Sub
