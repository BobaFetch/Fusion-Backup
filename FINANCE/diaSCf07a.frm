VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaSCf07a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Restore Previous Standards"
   ClientHeight    =   5295
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   5880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5295
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optVew 
      Height          =   255
      Left            =   360
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4920
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "Restore"
      Height          =   315
      Left            =   4920
      TabIndex        =   5
      ToolTipText     =   "Update Standard Cost To Calculated Total"
      Top             =   600
      Width           =   875
   End
   Begin VB.TextBox cmbprt 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Tag             =   "3"
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton cmdVew 
      Height          =   320
      Left            =   4320
      Picture         =   "diaSCf07a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Show BOM Structure"
      Top             =   720
      UseMaskColor    =   -1  'True
      Width           =   350
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3240
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5295
      FormDesignWidth =   5880
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaSCf07a.frx":0342
      PictureDn       =   "diaSCf07a.frx":0488
   End
   Begin VB.Label lblPreRev 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3300
      TabIndex        =   34
      Top             =   1440
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pre Rev"
      Height          =   285
      Index           =   7
      Left            =   2520
      TabIndex        =   33
      Top             =   1440
      Width           =   1065
   End
   Begin VB.Label lblPre 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2520
      TabIndex        =   32
      Top             =   4680
      Width           =   1035
   End
   Begin VB.Label lblRev 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   31
      Top             =   1440
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Revised"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   30
      Top             =   1440
      Width           =   1065
   End
   Begin VB.Label lblTot 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   28
      Top             =   4200
      Width           =   1035
   End
   Begin VB.Label lblOh 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   27
      Top             =   3840
      Width           =   1035
   End
   Begin VB.Label lblExp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   26
      Top             =   3480
      Width           =   1035
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   25
      Top             =   3120
      Width           =   1035
   End
   Begin VB.Label lblLab 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   24
      Top             =   2760
      Width           =   1035
   End
   Begin VB.Label lblHrs 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   23
      Top             =   2400
      Width           =   1035
   End
   Begin VB.Label lblCur 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   22
      Top             =   4680
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Standard Cost"
      Height          =   405
      Index           =   3
      Left            =   120
      TabIndex        =   21
      Top             =   4680
      Width           =   1035
   End
   Begin VB.Label lblTot 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   20
      Top             =   4200
      Width           =   1035
   End
   Begin VB.Label lblOh 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   19
      Top             =   3840
      Width           =   1035
   End
   Begin VB.Label lblExp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   18
      Top             =   3480
      Width           =   1035
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   17
      Top             =   3120
      Width           =   1035
   End
   Begin VB.Label lblLab 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   16
      Top             =   2760
      Width           =   1035
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Current"
      Height          =   285
      Index           =   13
      Left            =   1440
      TabIndex        =   15
      Top             =   2040
      Width           =   1035
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Previous"
      Height          =   285
      Index           =   14
      Left            =   2520
      TabIndex        =   14
      Top             =   2040
      Width           =   915
   End
   Begin VB.Label lblHrs 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   13
      Top             =   2400
      Width           =   1035
   End
   Begin VB.Line Line2 
      X1              =   1320
      X2              =   1320
      Y1              =   1920
      Y2              =   5040
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4560
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Overhead Cost"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Material Cost"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Expense Cost"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Labor Cost"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Hours"
      Height          =   285
      Index           =   16
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   285
      Index           =   17
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   1035
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1065
   End
End
Attribute VB_Name = "diaSCf07a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'*********************************************************************************
' diaScf07a - Restore Previous Standards
'
' Notes:
'
' Created: 12/06/02 (nth)
' Revisions:
'   09/14/04 (nth) Fixed bCancel
'
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodPart As Byte
Dim RdoPrt As ADODB.Recordset
Dim sMsg As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Private Sub cmbPrt_GotFocus()
   SelectFormat Me
End Sub

Private Sub cmbPrt_LostFocus()
   If Not bCancel Then
      cmbprt = CheckLen(cmbprt, 30)
      bGoodPart = GetPart
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             x As Single, y As Single)
   bCancel = True
End Sub

Private Sub cmdUpd_Click()
   RestoreCost
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      MouseCursor 13
      If Len(cUR.CurrentPart) Then
         cmbprt = cUR.CurrentPart
         bGoodPart = GetPart
      End If
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   sCurrForm = Caption
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If bGoodPart = 1 Then
      cUR.CurrentPart = Trim(cmbprt)
      SaveCurrentSelections
   End If
   Set RdoPrt = Nothing
   FormUnload
   Set diaSCf07a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub optVew_Click()
   If optVew.Value = vbUnchecked Then
      ' Part search is closing refresh form
      cmbPrt_LostFocus
   End If
End Sub

Private Function GetPart() As Byte
   Dim SPartRef As String
   
   On Error GoTo DiaErr1
   MouseCursor 13
   
   SPartRef = Compress(cmbprt)
   sSql = "SELECT PADESC, PALEVEL, PAREVDATE, PAPREVDATE,PAPREVSTDCOST, " _
          & "PASTDCOST, PABOMLABOR, PAPREVLABOR, PAPREVEXP, PAPREVMATL, PAPREVOH," _
          & "PAPREVHRS, PATOTHRS, PATOTEXP, PATOTLABOR, PATOTMATL, PATOTOH " _
          & "FROM PartTable WHERE PARTREF = '" & SPartRef & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_KEYSET)
   
   If bSqlRows Then
      With RdoPrt
         ' Check for uncostable parts
         If RdoPrt!PALEVEL = "5" Or RdoPrt!PALEVEL = "6" Then
            sMsg = "Cannot Cost Part Types 5 and 6"
            MsgBox sMsg, vbInformation
            cmbprt = "NONE"
            lblDsc.ForeColor = ES_RED
            lblDsc = "*** Part Number Wasn't Found ***"
            Set RdoPrt = Nothing
            cmbprt.SetFocus
            GetPart = 0
            Exit Function
         End If
         
         lblHrs(0) = Format(!PATOTHRS, "#,###,###.000")
         lblMat(0) = Format(!PATOTMATL, "#,###,###.000")
         lblExp(0) = Format(!PATOTEXP, "#,###,###.000")
         lblLab(0) = Format(!PATOTLABOR, "#,###,###.000")
         lblOh(0) = Format(!PATOTOH, "#,###,###.000")
         lblTot(0) = Format(!PATOTHRS + !PATOTMATL + !PATOTEXP + _
                !PATOTLABOR + !PATOTOH, "#,###,###.000")
         
         lblHrs(1) = Format(!PAPREVHRS, "#,###,###.000")
         lblMat(1) = Format(!PAPREVMATL, "#,###,###.000")
         lblExp(1) = Format(!PAPREVEXP, "#,###,###.000")
         lblLab(1) = Format(!PAPREVLABOR, "#,###,###.000")
         lblOh(1) = Format(!PAPREVOH, "#,###,###.000")
         lblTot(1) = Format(!PAPREVHRS + !PAPREVMATL + !PAPREVEXP + _
                !PAPREVLABOR + !PAPREVOH, "#,###,###.000")
         
         lblCur = Format(!PASTDCOST, "#,###,###.00")
         lblPre = Format(!PAPREVSTDCOST, "#,###,###.00")
         
         lblDsc.ForeColor = Me.ForeColor
         lblDsc = "" & Trim(!PADESC)
         lblRev = Format("" & Trim(!PAREVDATE), "mm/dd/yy")
         lblPreRev = Format("" & Trim(!PAPREVDATE), "mm/dd/yy")
      End With
      
      cmdUpd.enabled = True
      
      GetPart = 1
   Else
      cmbprt = "NONE"
      lblDsc.ForeColor = ES_RED
      lblDsc = "*** Part Number Wasn't Found ***"
      lblHrs(0) = ""
      lblMat(0) = ""
      lblExp(0) = ""
      lblLab(0) = ""
      lblOh(0) = ""
      
      lblTot(0) = ""
      
      lblHrs(1) = ""
      lblMat(1) = ""
      lblExp(1) = ""
      lblLab(1) = ""
      lblOh(1) = ""
      lblTot(1) = ""
      
      lblCur = ""
      lblRev = ""
      
      cmdUpd.enabled = False
      
      Set RdoPrt = Nothing
      cmbprt.SetFocus
      GetPart = 0
   End If
   
   MouseCursor 0
   Exit Function
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Public Sub RestoreCost()
   On Error Resume Next
   Err = 0
   MouseCursor 13
   
   With RdoPrt
      !PATOTHRS = !PAPREVHRS
      !PATOTLABOR = !PAPREVLABOR
      !PATOTMATL = !PAPREVMATL
      !PATOTEXP = !PAPREVEXP
      !PATOTOH = !PAPREVOH
      
      !PAPREVHRS = CSng(lblHrs(0))
      !PAPREVLABOR = CSng(lblLab(0))
      !PAPREVMATL = CSng(lblMat(0))
      !PAPREVEXP = CSng(lblExp(0))
      !PAPREVOH = CSng(lblOh(0))
      
      !PAPREVDATE = lblRev
      !PAREVDATE = Format(Now, "mm/dd/yy")
      
      !PAPREVSTDCOST = CCur(lblCur)
      !PASTDCOST = CCur(lblPre)
      .Update
   End With
   
   If Err > 0 Then
      ValidateEdit Me
      cmbprt.SetFocus
   Else
      sMsg = "Standard Cost Restored"
      SysMsg sMsg, True
      Set RdoPrt = Nothing
      bGoodPart = GetPart() ' Refresh
   End If
End Sub
