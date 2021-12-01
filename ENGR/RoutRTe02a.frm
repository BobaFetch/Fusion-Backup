VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form RoutRTe02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Routing Assignments"
   ClientHeight    =   2850
   ClientLeft      =   2130
   ClientTop       =   1455
   ClientWidth     =   6465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2850
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "RoutRTe02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "RoutRTe02a.frx":07AE
      Height          =   320
      Left            =   4920
      Picture         =   "RoutRTe02a.frx":0C88
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Parts Assigned To This Routing"
      Top             =   1800
      Width           =   350
   End
   Begin VB.CommandButton cmdAsn 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   5520
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Assign Routing to A Part"
      Top             =   1830
      Width           =   875
   End
   Begin VB.ComboBox cmbRte 
      Height          =   315
      Left            =   1600
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select Or Enter Routing"
      Top             =   1800
      Width           =   3255
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1600
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Or Enter Part"
      Top             =   750
      Width           =   3255
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5520
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   2160
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2850
      FormDesignWidth =   6465
   End
   Begin VB.Label txtDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1600
      TabIndex        =   12
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type"
      Height          =   285
      Index           =   4
      Left            =   5040
      TabIndex        =   11
      Top             =   750
      Width           =   1005
   End
   Begin VB.Label lblTyp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6000
      TabIndex        =   10
      Top             =   720
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   3
      Left            =   180
      TabIndex        =   9
      Top             =   1080
      Width           =   1365
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1600
      TabIndex        =   8
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Assigned"
      Height          =   285
      Index           =   2
      Left            =   180
      TabIndex        =   7
      Top             =   1470
      Width           =   1455
   End
   Begin VB.Label lblRout 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1605
      TabIndex        =   6
      Top             =   1470
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Routings"
      Height          =   285
      Index           =   1
      Left            =   180
      TabIndex        =   4
      Top             =   1830
      Width           =   1365
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   285
      Index           =   0
      Left            =   180
      TabIndex        =   5
      Top             =   750
      Width           =   1365
   End
End
Attribute VB_Name = "RoutRTe02a"
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
Dim bGoodPart As Byte
Dim bGoodRout As Byte
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub cmbPrt_Click()
   bGoodPart = GetPart()
   bGoodRout = GetRout(False, False)
   
End Sub


Private Sub cmbPrt_LostFocus()

   If (Not ValidPartNumber(cmbPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If

   bGoodPart = GetPart()
   bGoodRout = GetRout(False, False)
End Sub

Private Sub cmbRte_Click()
   bGoodRout = GetRout(0, 1)
   
End Sub


Private Sub cmbRte_LostFocus()
   cmbRte = CheckLen(cmbRte, 30)
   bGoodRout = GetRout(0, 1)
   
End Sub


Private Sub cmdAsn_Click()
   Dim sPartNumber As String
   Dim sNewRout As String
   
   bGoodPart = GetPart()
   If Not bGoodPart Then Exit Sub
   bGoodRout = GetRout(False, True)
   If Not bGoodRout Then Exit Sub
   
   On Error Resume Next
   MouseCursor 13
   sPartNumber = Compress(cmbPrt)
   sNewRout = Compress(cmbRte)
   sSql = "UPDATE PartTable SET PAROUTING='" & sNewRout & "' WHERE PARTREF='" & sPartNumber & "'"
   clsADOCon.ExecuteSQL sSql ' rdExecDirect
   MouseCursor 0
   If clsADOCon.RowsAffected > 0 Then
      lblRout = cmbRte
      SysMsg "Routing Assigned", True, Me
   Else
      MsgBox "Couldn't Assign Routing.", _
         vbExclamation, Caption
   End If
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3102
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdVew_Click()
   If cmdVew Then
      RteTree.Show
      cmdVew = False
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillParts
      FillRoutings
      If cmbRte.ListCount > 0 Then bGoodRout = GetRout(False, False)
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
   On Error Resume Next
   SaveCurrentSelections
   FormUnload
   Set RoutRTe02a = Nothing
   
End Sub





Private Sub FillParts()
   cmbPrt.Clear
   On Error GoTo DiaErr1
   sSql = "Qry_FillPartRoutings"
   LoadComboBox cmbPrt
   If bSqlRows Then
      'If cUR.CurrentPart <> "" Then
      '    cmbPrt = cUR.CurrentPart
      'Else
      cmbPrt = cmbPrt.List(0)
      'End If
   End If
   bGoodPart = GetPart()
   If lblRout = "" Then lblRout = "NONE"
   Exit Sub
   
DiaErr1:
   sProcName = "fillparts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Function GetRout(bFillLabel As Byte, bMessage As Byte) As Byte
   Dim RdoRte As ADODB.Recordset
   Dim sRout As String
   GetRout = False
   On Error GoTo DiaErr1
   sSql = "Qry_GetToolRout '" & Compress(cmbRte) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte)
   If bSqlRows Then
      With RdoRte
         GetRout = True
         If bFillLabel Then
            lblRout = "" & Trim(!RTNUM)
         Else
            cmbRte = "" & Trim(!RTNUM)
            txtDsc = "" & Trim(!RTDESC)
         End If
         ClearResultSet RdoRte
      End With
   Else
      If lblRout <> "NONE" Then
         If bOnLoad = 0 Then If bMessage Then MsgBox "Routing Wasn't Found.", vbExclamation, Caption
      End If
      GetRout = False
   End If
   Set RdoRte = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getrout"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function GetPart() As Byte
   Dim RdoPrt As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_GetPartRouting '" & Compress(cmbPrt) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
   If bSqlRows Then
      With RdoPrt
         Do Until .EOF
            cmbPrt = "" & Trim(!PartNum)
            lblDsc = "" & Trim(!PADESC)
            lblRout = "" & Trim(!PAROUTING)
            lblTyp = Format(0 + !PALEVEL, "0")
            cUR.CurrentPart = cmbPrt
            .MoveNext
         Loop
         ClearResultSet RdoPrt
      End With
      If lblRout = "" Then lblRout = "None"
      GetPart = True
   Else
      MsgBox "Couldn't Find Part.", vbExclamation, Caption
      GetPart = False
   End If
   If lblRout = "" Then lblRout = "NONE"
   Set RdoPrt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function
