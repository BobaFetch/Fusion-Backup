VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form CapaCPe03a 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Work Center Calendars"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   4201
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CapaCPe03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   1440
      TabIndex        =   9
      ToolTipText     =   "Times From Company Calendar Or Work Center Settings (Recommended)"
      Top             =   1320
      Width           =   2895
      Begin VB.OptionButton optWcn 
         Caption         =   "Work Center"
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         ToolTipText     =   "Times From Company Calendar Or Work Center Settings (Recommended)"
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optCom 
         Caption         =   "Calendar"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Times From Company Calendar Or Work Center Settings (Recommended)"
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "A&LL"
      Height          =   315
      Left            =   4080
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Feature Allows All Work Centers To Be Created From The Loaded One"
      Top             =   960
      Width           =   875
   End
   Begin VB.CheckBox optCal 
      Caption         =   "Calendar"
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   2280
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2295
      FormDesignWidth =   5055
   End
   Begin VB.CommandButton cmdCal 
      Caption         =   "C&alendar"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4080
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Press To Open Calendar"
      Top             =   600
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4080
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin VB.ComboBox cmbWcn 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      ToolTipText     =   "Select Work Center From List"
      Top             =   855
      Width           =   1815
   End
   Begin VB.ComboBox cmbShp 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      ToolTipText     =   "Select Shop From List"
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Times From"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shop"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "CapaCPe03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'9/22/03 Added Calendar options
Option Explicit
Dim AdoQry As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter
Dim ADOParameter2 As ADODB.Parameter
Dim bGoodCenter As Byte
Dim bOnLoad As Byte

Private txtKeyPress(2) As New EsiKeyBd
Private txtGotFocus(2) As New EsiKeyBd

Private Sub FormatControls()
   On Error Resume Next
   Set txtGotFocus(0).esCmbGotfocus = cmbShp
   Set txtGotFocus(1).esCmbGotfocus = cmbWcn
   
   Set txtKeyPress(0).esCmbKeylock = cmbShp
   Set txtKeyPress(1).esCmbKeyCase = cmbWcn
   
End Sub

Private Sub cmbShp_Click()
   FillWorkCenters
   
End Sub


Private Sub cmbWcn_Click()
   bGoodCenter = GetCenter()
   
End Sub


Private Sub cmbWcn_LostFocus()
   cmbWcn = CheckLen(cmbWcn, 12)
   bGoodCenter = GetCenter()
   
End Sub

Private Sub cmdAll_Click()
   On Error GoTo DiaErr1
   SaveSetting "Esi2000", "EsiProd", "WcCalOption", Trim(optWcn.Value)
   optCal.Value = vbChecked
   CapaCPe03b.lblFrom(0) = "ALL"
   CapaCPe03b.lblFrom(1) = "ALL"
   If optWcn.Value = True Then CapaCPe03b.optWcn.Value = vbChecked
   CapaCPe03b.optAll.Value = vbChecked
   CapaCPe03b.Show
   Exit Sub
   
DiaErr1:
   sProcName = "cmdall_click"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmdCal_Click()
   On Error GoTo DiaErr1
   SaveSetting "Esi2000", "EsiProd", "WcCalOption", Trim(optWcn.Value)
   If bGoodCenter Then
      MouseCursor 13
      optCal.Value = vbChecked
      If optWcn.Value = True Then CapaCPe03b.optWcn.Value = vbChecked
      CapaCPe03b.Show
      Unload Me
   Else
      MsgBox "No Such Shop and Work Center Combination.", vbExclamation, Caption
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "cmdcal_click"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4201
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub Form_Activate()
   Dim b As Boolean
   MouseCursor 13
   If bOnLoad Then
      b = GetSetting("Esi2000", "EsiProd", "WcCalOption", b)
      If b Then
         optWcn.Value = True
         optCom.Value = False
      Else
         optWcn.Value = False
         optCom.Value = True
      End If
      bOnLoad = 0
   End If
   MDISect.lblBotPanel = Caption
   optCal.Value = vbUnchecked
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   GetOptions
   On Error Resume Next
   sSql = "SELECT WCNREF,WCNNUM FROM WcntTable WHERE WCNREF= ? AND WCNSHOP= ? "
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adChar
   AdoParameter1.SIZE = 12
   Set ADOParameter2 = New ADODB.Parameter
   ADOParameter2.Type = adChar
   ADOParameter2.SIZE = 12
   
   AdoQry.Parameters.Append AdoParameter1
   AdoQry.Parameters.Append ADOParameter2
   FillShops
   Show
   
End Sub

Private Sub FillShops()
   'Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_FillShops "
   LoadComboBox cmbShp
   If cmbShp.ListCount > 0 Then cmbShp = cmbShp.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillshops"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub FillWorkCenters()
   cmbWcn.Clear
   On Error GoTo DiaErr1
   sSql = "Qry_FillWorkCenters '" & Compress(cmbShp) & "'"
   LoadComboBox cmbWcn
   If cmbWcn.ListCount > 0 Then cmbWcn = cmbWcn.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillworkcenters"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If optCal.Value = vbUnchecked Then FormUnload
   SaveSetting "Esi2000", "EsiProd", "WcCalOption", Trim(optWcn.Value)
   Set AdoParameter1 = Nothing
   Set ADOParameter2 = Nothing
   Set AdoQry = Nothing
   Set CapaCPe03a = Nothing
   
End Sub



Private Function GetCenter() As Byte
   Dim RdoCnt As ADODB.Recordset
   Dim sCenter As String
   Dim sShop As String
   
   sCenter = Compress(cmbWcn)
   sShop = Compress(cmbShp)
   
   On Error GoTo DiaErr1
   AdoQry.Parameters(0).Value = sCenter
   AdoQry.Parameters(1).Value = sShop
   bSqlRows = clsADOCon.GetQuerySet(RdoCnt, AdoQry, ES_KEYSET, False, 1)
   If bSqlRows Then
      With RdoCnt
         GetCenter = True
         cmbWcn = "" & Trim(!WCNNUM)
         cmdCal.Enabled = True
         ClearResultSet RdoCnt
      End With
   Else
      GetCenter = False
      cmdCal.Enabled = False
   End If
   Set RdoCnt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getcenter"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub GetOptions()
   On Error Resume Next
   cmbShp = GetSetting("Esi2000", "EsiProd", "Cwopc1", cmbShp)
   cmbWcn = GetSetting("Esi2000", "EsiProd", "Cwopc2", cmbWcn)
   
End Sub

Private Sub SaveOptions()
   'Save by Menu Option
   SaveSetting "Esi2000", "EsiProd", "Cwopc1", Trim(cmbShp)
   SaveSetting "Esi2000", "EsiProd", "Cwopc2", Trim(cmbWcn)
   
End Sub
