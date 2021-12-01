VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form RoutRTf03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reassign Shops And Work Centers"
   ClientHeight    =   2370
   ClientLeft      =   3060
   ClientTop       =   1530
   ClientWidth     =   4905
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2370
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "RoutRTf03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbNwc 
      Height          =   288
      Left            =   1800
      TabIndex        =   3
      Top             =   1776
      Width           =   1815
   End
   Begin VB.ComboBox cmbNsh 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1416
      Width           =   1815
   End
   Begin VB.ComboBox cmbOwc 
      Height          =   288
      Left            =   1800
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.ComboBox cmbOSh 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton cmdAsn 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   3960
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Reassign To All Routings"
      Top             =   1440
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   3960
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   2280
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2370
      FormDesignWidth =   4905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Work Center"
      Height          =   288
      Index           =   3
      Left            =   180
      TabIndex        =   9
      Top             =   1776
      Width           =   1632
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Shop"
      Height          =   288
      Index           =   2
      Left            =   180
      TabIndex        =   8
      Top             =   1416
      Width           =   1548
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Old Work Center"
      Height          =   288
      Index           =   1
      Left            =   180
      TabIndex        =   7
      Top             =   960
      Width           =   1548
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Old Shop"
      Height          =   288
      Index           =   0
      Left            =   180
      TabIndex        =   6
      Top             =   600
      Width           =   1428
   End
End
Attribute VB_Name = "RoutRTf03a"
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
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub cmbNsh_Click()
   FillNewCenters
   
End Sub

Private Sub cmbOSh_Click()
   FillOldCenters
   
End Sub


Private Sub cmdAsn_Click()
   Dim iList As Integer
   Dim n As Integer
   Dim bResponse As Byte
   Dim sOldShop As String
   Dim sOldCenter As String
   Dim sNewShop As String
   Dim sNewCenter As String
   
   bResponse = MsgBox("Reassign Old Shops, Work Centers?", ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      MouseCursor 0
      On Error Resume Next
      cmdCan.SetFocus
      Width = Width + 10
      Exit Sub
   End If
    
   cmdAsn.Enabled = False
   cmdCan.Enabled = False
   sOldShop = Compress(cmbOSh)
   sOldCenter = Compress(cmbOwc)
   sNewShop = Compress(cmbNsh)
   sNewCenter = Compress(cmbNwc)
   
   MouseCursor 13
   On Error GoTo reAsgnErr1
   sSql = "UPDATE RtopTable SET OPSHOP='" & sNewShop & "' WHERE OPSHOP='" & sOldShop & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   iList = clsADOCon.RowsAffected
   sSql = "UPDATE RtopTable SET OPCENTER='" & sNewCenter & "' WHERE OPCENTER='" & sOldCenter & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   n = clsADOCon.RowsAffected
   MouseCursor 0
   MsgBox str(n) & " Operations Updated.", vbInformation, Caption
   cmdAsn.Enabled = True
   cmdCan.Enabled = True
   Exit Sub
   
reAsgnErr1:
   CurrError.Description = Err.Description
   Resume reAsgnErr2
reAsgnErr2:
   MsgBox CurrError.Description & " Couldn't Reassign.", vbExclamation, Caption
   cmdAsn.Enabled = True
   cmdCan.Enabled = True
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3152
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = 0
      FillCombos
      FillOldCenters
      FillNewCenters
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
   FormUnload
   Set RoutRTf03a = Nothing
   
End Sub



Private Sub FillCombos()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_FillShops"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb)
   If bSqlRows Then
      With RdoCmb
         cmbOSh = "" & Trim(!SHPNUM)
         cmbNsh = "" & Trim(!SHPNUM)
         Do Until .EOF
            AddComboStr cmbOSh.hwnd, "" & Trim(!SHPNUM)
            AddComboStr cmbNsh.hwnd, "" & Trim(!SHPNUM)
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   End If
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombos"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub FillOldCenters()
   On Error GoTo DiaErr1
   cmbOwc.Clear
   sSql = "Qry_FillWorkCenters '" & Compress(cmbOSh) & "'"
   LoadComboBox cmbOwc
   If cmbOwc.ListCount > 0 Then cmbOwc = cmbOwc.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "filloldce"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub FillNewCenters()
   On Error GoTo DiaErr1
   cmbNwc.Clear
   sSql = "Qry_FillWorkCenters '" & Compress(cmbNsh) & "'"
   LoadComboBox cmbNwc
   If cmbNwc.ListCount > 0 Then cmbNwc = cmbNwc.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillnewce"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
