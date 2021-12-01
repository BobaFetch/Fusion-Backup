VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form RoutRTe01e 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fill From Library"
   ClientHeight    =   4440
   ClientLeft      =   2985
   ClientTop       =   2310
   ClientWidth     =   6150
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4440
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdUndo 
      Caption         =   "&Undo"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5220
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Undo The Most Recent Changes (This Session)"
      Top             =   1320
      Width           =   870
   End
   Begin VB.CommandButton cmdFil 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5220
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Fill The Current Operation"
      Top             =   960
      Width           =   870
   End
   Begin VB.CheckBox optSrv 
      Alignment       =   1  'Right Justify
      Caption         =   "Service Op?"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4470
      TabIndex        =   6
      Top             =   2070
      Width           =   1185
   End
   Begin VB.CheckBox optPck 
      Alignment       =   1  'Right Justify
      Caption         =   "Pick Op?"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4470
      TabIndex        =   5
      Top             =   1710
      Width           =   1185
   End
   Begin VB.ComboBox cmbShp 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1770
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   270
      Width           =   1815
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5220
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin VB.ComboBox cmbOpr 
      Height          =   315
      Left            =   1770
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   630
      Width           =   1815
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5760
      Top             =   3960
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4440
      FormDesignWidth =   6150
   End
   Begin VB.Label lblGridRow 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3720
      TabIndex        =   24
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPrt 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1800
      TabIndex        =   22
      Top             =   3960
      Width           =   3072
   End
   Begin VB.Label lblCmt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   1770
      TabIndex        =   21
      Top             =   2430
      Width           =   3885
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblUnt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3390
      TabIndex        =   20
      Top             =   2070
      Width           =   1005
   End
   Begin VB.Label lblSet 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1770
      TabIndex        =   19
      Top             =   2070
      Width           =   1095
   End
   Begin VB.Label lblMdy 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3390
      TabIndex        =   18
      Top             =   1710
      Width           =   1005
   End
   Begin VB.Label lblQdy 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1770
      TabIndex        =   17
      Top             =   1710
      Width           =   1095
   End
   Begin VB.Label lblWcn 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1770
      TabIndex        =   16
      Top             =   1350
      Width           =   1680
   End
   Begin VB.Label lblDsc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1770
      TabIndex        =   15
      Top             =   990
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
      Height          =   285
      Index           =   10
      Left            =   270
      TabIndex        =   14
      Top             =   2430
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Part"
      Height          =   285
      Index           =   7
      Left            =   270
      TabIndex        =   13
      Top             =   3960
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit"
      Height          =   285
      Index           =   9
      Left            =   2940
      TabIndex        =   12
      Top             =   2070
      Width           =   645
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Setup"
      Height          =   285
      Index           =   8
      Left            =   270
      TabIndex        =   11
      Top             =   2070
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Move"
      Height          =   285
      Index           =   6
      Left            =   2940
      TabIndex        =   10
      Top             =   1710
      Width           =   645
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Queue"
      Height          =   285
      Index           =   5
      Left            =   270
      TabIndex        =   9
      Top             =   1710
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center"
      Height          =   285
      Index           =   1
      Left            =   270
      TabIndex        =   8
      Top             =   1350
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   4
      Left            =   270
      TabIndex        =   7
      Top             =   990
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shop"
      Height          =   285
      Index           =   0
      Left            =   270
      TabIndex        =   4
      Top             =   270
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operation Name"
      Height          =   285
      Index           =   2
      Left            =   270
      TabIndex        =   3
      Top             =   630
      Width           =   1425
   End
End
Attribute VB_Name = "RoutRTe01e"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'11/2/04 Revised general structure and Fill button
'        Attempts to update Ops grid
'1/26/07 Undo
Option Explicit
Dim RdoLib As ADODB.Recordset

Dim bGoodOp As Byte
Dim bOnLoad As Byte
Dim sUndoCmt As String
Dim sUndoShp As String
Dim sUndoCnt As String

Dim sUndoQdy As String
Dim sUndoMdy As String
Dim sUndoSet As String
Dim sUndoUnt As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub cmbOpr_Click()
   GetOperation
   
End Sub

Private Sub cmbOpr_LostFocus()
   GetOperation
   
End Sub


Private Sub cmbShp_Click()
   FillOperations
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub




Private Sub cmdFil_Click()
   If Not bGoodOp Then
      MsgBox "No Information To Fill.", vbInformation, Caption
   Else
      MouseCursor 13
      sUndoCmt = RoutRTe01b.txtCmt
      sUndoShp = RoutRTe01b.cmbShp
      sUndoCnt = RoutRTe01b.cmbWcn
      
      sUndoQdy = RoutRTe01b.txtQdy
      sUndoMdy = RoutRTe01b.txtMdy
      sUndoSet = RoutRTe01b.txtSet
      sUndoUnt = RoutRTe01b.txtUnt
      
      If Val(lblGridRow) > 0 Then
         RoutRTe01b.Grd.Col = 1
         RoutRTe01b.Grd.Text = cmbShp
         RoutRTe01b.Grd.Col = 2
         RoutRTe01b.Grd.Text = lblWcn
         RoutRTe01b.Grd.Col = 3
         RoutRTe01b.Grd.Text = Left(lblCmt, 20)
      End If
      RoutRTe01b.cmbShp = cmbShp
      RoutRTe01b.cmbWcn = lblWcn
      RoutRTe01b.txtQdy = lblQdy
      RoutRTe01b.txtMdy = lblMdy
      RoutRTe01b.txtSet = lblSet
      RoutRTe01b.txtUnt = lblUnt
      RoutRTe01b.txtCmt = lblCmt
      RoutRTe01b.cmbPrt = lblPrt
      RoutRTe01b.optSrv.value = optSrv.value
      RoutRTe01b.optPck.value = optPck.value
      MouseCursor 0
      cmdUndo.Enabled = True
      'Unload Me
   End If
   
End Sub

Private Sub cmdUndo_Click()
   If Val(lblGridRow) > 0 Then
      RoutRTe01b.Grd.Col = 1
      RoutRTe01b.Grd.Text = sUndoShp
      RoutRTe01b.Grd.Col = 2
      RoutRTe01b.Grd.Text = sUndoCnt
      RoutRTe01b.Grd.Col = 3
      RoutRTe01b.Grd.Text = Left(sUndoCmt, 20)
   End If
   RoutRTe01b.txtCmt = sUndoCmt
   RoutRTe01b.cmbShp = sUndoShp
   RoutRTe01b.cmbWcn = sUndoCnt
   RoutRTe01b.txtQdy = sUndoQdy
   RoutRTe01b.txtMdy = sUndoMdy
   RoutRTe01b.txtSet = sUndoSet
   RoutRTe01b.txtUnt = sUndoUnt
   
End Sub


Private Sub Form_Activate()
   MouseCursor 13
   If bOnLoad Then
      ES_TimeFormat = GetTimeFormat()
      FillShops
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   SetFormSize Me
   Move 2000, 2000
   FormatControls
   cUR.CurrentShop = GetSetting("Esi2000", "Current", "Shop", cUR.CurrentShop)
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   RoutRTe01b.optLib = vbUnchecked
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   RdoLib.Close
   MouseCursor 0
   Set RdoLib = Nothing
   Set RoutRTe01e = Nothing
   
End Sub



Private Sub FillShops()
   On Error Resume Next
   sSql = "Qry_FillShops"
   LoadComboBox cmbShp
   If bSqlRows Then
      If Len(cUR.CurrentShop) > 0 Then
         cmbShp = cUR.CurrentShop
      Else
         cmbShp = cmbShp.List(0)
      End If
   End If
   FillOperations
   
End Sub

Private Sub FillOperations()
   Dim sShop As String
   sShop = Compress(cmbShp)
   On Error Resume Next
   cmbOpr.Clear
   cmdFil.Enabled = False
   sSql = "SELECT LIBREF,LIBNUM,LIBDESC FROM RlbrTable WHERE " _
          & "LIBSHOP='" & sShop & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLib, ES_FORWARD)
   If bSqlRows Then
      With RdoLib
         cmbOpr = "" & Trim(!LIBNUM)
         lblDsc = "" & Trim(!LIBDESC)
         Do Until .EOF
            AddComboStr cmbOpr.hwnd, "" & Trim(!LIBNUM)
            .MoveNext
         Loop
         ClearResultSet RdoLib
      End With
   Else
      lblDsc = ""
      lblWcn = ""
      lblQdy = ""
      lblMdy = ""
      lblSet = ""
      lblUnt = ""
      lblCmt = ""
      optSrv.value = 0
      optPck.value = 0
      lblPrt = ""
   End If
   RdoLib.Close
   GetOperation
   
End Sub

Private Sub GetOperation()
   Dim sOperation As String
   Dim sThisShop As String
   
   sThisShop = Compress(cmbShp)
   sOperation = Compress(cmbOpr)
   On Error Resume Next
   sSql = "SELECT * FROM RlbrTable WHERE LIBREF='" & sOperation & "' AND LIBSHOP='" & sThisShop & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLib, ES_STATIC)
   If bSqlRows Then
      With RdoLib
         cmbOpr = "" & Trim(!LIBNUM)
         lblDsc = "" & Trim(!LIBDESC)
         lblWcn = "" & Trim(!LIBCENTER)
         lblQdy = Format(!LIBQHRS, ES_QuantityDataFormat)
         lblMdy = Format(!LIBMHRS, ES_QuantityDataFormat)
         lblSet = Format(!LIBSETUP, ES_QuantityDataFormat)
         lblUnt = Format(!LIBUNIT, ES_TimeFormat)
         lblCmt = "" & Trim(!LIBCOMT)
         optSrv.value = 0 + !LIBSERVICE
         optPck.value = 0 + !LIBPICKOP
         lblPrt = "" & Trim(!LIBSERVPART)
         If lblPrt = "" Then lblPrt = "NONE"
         bGoodOp = True
      End With
      If Trim(lblWcn) = "" Then
         MsgBox "Please Add A Valid Work Center To This Library Operation.", _
            vbInformation, Caption
         cmdFil.Enabled = False
      Else
         cmdFil.Enabled = True
      End If
   Else
      bGoodOp = False
      cmbOpr = ""
      lblDsc = ""
      lblWcn = ""
      optSrv.value = 0
      optPck.value = 0
      lblPrt = ""
      lblCmt = ""
      cmdFil.Enabled = False
   End If
   RdoLib.Close
   GetCenter
   
End Sub

Private Sub GetCenter()
   On Error Resume Next
   sSql = "SELECT WCNREF,WCNUM FROM WcntTable WHERE WCNREF='" _
          & Compress(lblWcn) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLib)
   If bSqlRows Then
      lblWcn = "" & Trim(RdoLib!WCNNUM)
   Else
      lblWcn = ""
   End If
   RdoLib.Close
   
End Sub
