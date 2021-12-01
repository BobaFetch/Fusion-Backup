VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ShopSHe02g 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fill From Library"
   ClientHeight    =   4440
   ClientLeft      =   2985
   ClientTop       =   2310
   ClientWidth     =   6165
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
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdUndo 
      Caption         =   "&Undo"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5220
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Undo The Most Recent Changes (This Session)"
      Top             =   960
      Width           =   870
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHe02g.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdFil 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   5220
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Fill The Current Operation"
      Top             =   600
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
      FormDesignWidth =   6165
   End
   Begin VB.Label lblPrt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1770
      TabIndex        =   22
      Top             =   3960
      Width           =   3075
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
Attribute VB_Name = "ShopSHe02g"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
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
      sUndoCmt = ShopSHe02c.txtCmt
      sUndoShp = ShopSHe02c.cmbShp
      sUndoCnt = ShopSHe02c.cmbWcn
      
      sUndoQdy = ShopSHe02c.txtQdy
      sUndoMdy = ShopSHe02c.txtMdy
      sUndoSet = ShopSHe02c.txtSet
      sUndoUnt = ShopSHe02c.txtUnt
      
      ShopSHe02c.cmbShp = cmbShp
      ShopSHe02c.cmbWcn = lblWcn
      ShopSHe02c.txtQdy = lblQdy
      ShopSHe02c.txtMdy = lblMdy
      ShopSHe02c.txtSet = lblSet
      ShopSHe02c.txtUnt = lblUnt
      ShopSHe02c.txtCmt = lblCmt
      ShopSHe02c.cmbPrt = lblPrt
      ShopSHe02c.optSrv.Value = optSrv.Value
      ShopSHe02c.optPck.Value = optPck.Value
   End If
   MouseCursor 0
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3105
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdUndo_Click()
   ShopSHe02c.txtCmt = sUndoCmt
   ShopSHe02c.cmbShp = sUndoShp
   ShopSHe02c.cmbWcn = sUndoCnt
   ShopSHe02c.txtQdy = sUndoQdy
   ShopSHe02c.txtMdy = sUndoMdy
   ShopSHe02c.txtSet = sUndoSet
   ShopSHe02c.txtUnt = sUndoUnt
   
End Sub

Private Sub Form_Activate()
   MouseCursor 13
   If bOnLoad Then
      FillShops
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   SetFormSize Me
   Move 800, 800
   FormatControls
   cUR.CurrentShop = GetSetting("Esi2000", "Current", "Shop", cUR.CurrentShop)
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   ShopSHe02c.optLib = vbUnchecked
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   RdoLib.Close
   Set RdoLib = Nothing
   MouseCursor 0
   Set ShopSHe02g = Nothing
   
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
   Dim sShop
   sShop = Compress(cmbShp)
   On Error Resume Next
   cmbOpr.Clear
   sSql = "SELECT LIBREF,LIBNUM,LIBSHOP,LIBDESC FROM RlbrTable " _
          & "WHERE LIBSHOP='" & sShop & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLib)
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
      optSrv.Value = 0
      optPck.Value = 0
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
         lblQdy = Format(!LIBQHRS, "##0.000")
         lblMdy = Format(!LIBMHRS, "##0.000")
         lblSet = Format(!LIBSETUP, "##0.000")
         lblUnt = Format(!LIBUNIT, "##0.000")
         lblCmt = "" & Trim(!LIBCOMT)
         optSrv.Value = 0 + !LIBSERVICE
         optPck.Value = 0 + !LIBPICKOP
         lblPrt = "" & Trim(!LIBSERVPART)
         If lblPrt = "" Then lblPrt = "NONE"
         bGoodOp = True
      End With
   Else
      bGoodOp = False
      cmbOpr = ""
      lblDsc = ""
      lblWcn = ""
      optSrv.Value = 0
      optPck.Value = 0
      lblPrt = ""
   End If
   RdoLib.Close
   GetCenter
   
End Sub

Private Sub GetCenter()
   Dim sThisCenter As String
   On Error Resume Next
   sThisCenter = Compress(lblWcn)
   sSql = "SELECT * FROM WcntTable WHERE WCNREF='" & sThisCenter & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLib)
   If bSqlRows Then
      lblWcn = "" & Trim(RdoLib!WCNNUM)
   Else
      lblWcn = ""
   End If
   RdoLib.Close
   
End Sub
