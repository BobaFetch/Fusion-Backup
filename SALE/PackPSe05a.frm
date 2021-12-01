VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form PackPSe05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revise A Packing Slip - Not Shipped"
   ClientHeight    =   4935
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   7455
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   Begin VB.Frame tabFrame 
      Height          =   2892
      Index           =   1
      Left            =   8280
      TabIndex        =   64
      Top             =   1920
      Width           =   6972
      Begin VB.TextBox txtVia 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Tag             =   "6"
         Top             =   120
         Width           =   1875
      End
      Begin VB.TextBox txtSel 
         Height          =   285
         Left            =   5040
         TabIndex        =   6
         Tag             =   "3"
         Top             =   120
         Width           =   1675
      End
      Begin VB.TextBox txtCnt 
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Tag             =   "3"
         Top             =   480
         Width           =   1675
      End
      Begin VB.TextBox txtFrt 
         Height          =   285
         Left            =   5040
         TabIndex        =   8
         Tag             =   "1"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtBxs 
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Tag             =   "1"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtShn 
         Height          =   285
         Left            =   5040
         TabIndex        =   10
         Tag             =   "3"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtCrt 
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Tag             =   "1"
         Top             =   1200
         Width           =   735
      End
      Begin VB.ComboBox cmbTrm 
         ForeColor       =   &H00800000&
         Height          =   288
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   12
         Tag             =   "8"
         ToolTipText     =   "Select Terms From List"
         Top             =   1560
         Width           =   660
      End
      Begin VB.TextBox txtTrk 
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Tag             =   "3"
         ToolTipText     =   "Tracking Number (Free Form To 40 Chars)"
         Top             =   2040
         Width           =   3555
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ship Via"
         Height          =   252
         Index           =   26
         Left            =   120
         TabIndex        =   74
         Top             =   120
         Width           =   1092
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Seal"
         Height          =   252
         Index           =   7
         Left            =   5040
         TabIndex        =   73
         Top             =   480
         Width           =   1092
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Container"
         Height          =   252
         Index           =   8
         Left            =   120
         TabIndex        =   72
         Top             =   480
         Width           =   1092
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Freight"
         Height          =   252
         Index           =   9
         Left            =   3720
         TabIndex        =   71
         Top             =   480
         Width           =   1092
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Boxes"
         Height          =   252
         Index           =   10
         Left            =   120
         TabIndex        =   70
         Top             =   840
         Width           =   1092
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Shipper Number"
         Height          =   252
         Index           =   11
         Left            =   3720
         TabIndex        =   69
         Top             =   840
         Width           =   1332
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Carton"
         Height          =   252
         Index           =   12
         Left            =   120
         TabIndex        =   68
         Top             =   1200
         Width           =   1092
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Shipping Terms"
         Height          =   252
         Index           =   13
         Left            =   120
         TabIndex        =   67
         Top             =   1560
         Width           =   1212
      End
      Begin VB.Label lblTrm 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   2568
         TabIndex        =   66
         Top             =   1560
         Width           =   2652
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tracking No"
         Height          =   252
         Index           =   24
         Left            =   240
         TabIndex        =   65
         Top             =   2040
         Width           =   1212
      End
   End
   Begin VB.Frame tabFrame 
      Height          =   2892
      Index           =   2
      Left            =   7920
      TabIndex        =   54
      Top             =   1920
      Width           =   6972
      Begin VB.TextBox txtGrs 
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Tag             =   "1"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtNet 
         Height          =   285
         Left            =   3600
         TabIndex        =   16
         Tag             =   "1"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtLod 
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Tag             =   "3"
         Top             =   120
         Width           =   1675
      End
      Begin VB.TextBox txtCuf 
         Height          =   285
         Left            =   5400
         TabIndex        =   17
         Tag             =   "1"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtHgt 
         Height          =   285
         Left            =   1680
         TabIndex        =   18
         Tag             =   "1"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtWid 
         Height          =   285
         Left            =   5400
         TabIndex        =   20
         Tag             =   "1"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtGir 
         Height          =   285
         Left            =   1680
         TabIndex        =   21
         Tag             =   "1"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtRem 
         Height          =   855
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Tag             =   "9"
         Top             =   1800
         Width           =   5055
      End
      Begin VB.TextBox txtLen 
         Height          =   285
         Left            =   3600
         TabIndex        =   19
         Tag             =   "1"
         Top             =   840
         Width           =   975
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Gross Lbs"
         Height          =   252
         Index           =   14
         Left            =   120
         TabIndex        =   63
         Top             =   480
         Width           =   1092
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Lbs"
         Height          =   252
         Index           =   15
         Left            =   2880
         TabIndex        =   62
         Top             =   480
         Width           =   732
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Load No"
         Height          =   252
         Index           =   16
         Left            =   120
         TabIndex        =   61
         Top             =   120
         Width           =   1092
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cu Feet"
         Height          =   252
         Index           =   17
         Left            =   4680
         TabIndex        =   60
         Top             =   480
         Width           =   1092
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Height"
         Height          =   252
         Index           =   18
         Left            =   120
         TabIndex        =   59
         Top             =   840
         Width           =   972
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Length"
         Height          =   252
         Index           =   19
         Left            =   2880
         TabIndex        =   58
         Top             =   840
         Width           =   732
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Width"
         Height          =   252
         Index           =   20
         Left            =   4680
         TabIndex        =   57
         Top             =   840
         Width           =   852
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Girth"
         Height          =   252
         Index           =   21
         Left            =   120
         TabIndex        =   56
         Top             =   1200
         Width           =   732
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Load Remarks:"
         Height          =   252
         Index           =   22
         Left            =   120
         TabIndex        =   55
         Top             =   1560
         Width           =   1212
      End
   End
   Begin VB.Frame tabFrame 
      Height          =   2892
      Index           =   3
      Left            =   7560
      TabIndex        =   52
      Top             =   1920
      Width           =   6972
      Begin VB.TextBox txtBtName 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   240
         Width           =   3475
      End
      Begin VB.TextBox txtBtAdr 
         BackColor       =   &H00E0E0E0&
         Height          =   1095
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   600
         Width           =   3475
      End
      Begin VB.TextBox txtBtCity 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txtBtState 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtBtZip 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bill To:"
         Height          =   252
         Index           =   6
         Left            =   120
         TabIndex        =   53
         Top             =   120
         Width           =   1092
      End
   End
   Begin VB.Frame tabFrame 
      Height          =   2892
      Index           =   0
      Left            =   40
      TabIndex        =   45
      Top             =   1920
      Width           =   6972
      Begin VB.TextBox txtStn 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Tag             =   "2"
         Top             =   120
         Width           =   3475
      End
      Begin VB.TextBox txtSta 
         Height          =   975
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   3
         Tag             =   "9"
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox txtCmt 
         Height          =   1005
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Tag             =   "9"
         Top             =   1560
         Width           =   5175
      End
      Begin VB.CommandButton cmdComments 
         DisabledPicture =   "PackPSe05a.frx":0000
         DownPicture     =   "PackPSe05a.frx":0972
         Height          =   350
         Index           =   0
         Left            =   1200
         Picture         =   "PackPSe05a.frx":12E4
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Standard Comments"
         Top             =   1560
         Width           =   350
      End
      Begin VB.CommandButton cmdComments 
         DisabledPicture =   "PackPSe05a.frx":18E6
         DownPicture     =   "PackPSe05a.frx":2258
         Height          =   350
         Index           =   1
         Left            =   5280
         Picture         =   "PackPSe05a.frx":2BCA
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Standard Comments"
         Top             =   480
         Width           =   350
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ship To Name"
         Height          =   252
         Index           =   3
         Left            =   120
         TabIndex        =   51
         Top             =   120
         Width           =   1332
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ship To Address"
         Height          =   252
         Index           =   4
         Left            =   120
         TabIndex        =   50
         Top             =   480
         Width           =   1332
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   252
         Index           =   5
         Left            =   120
         TabIndex        =   49
         Top             =   1800
         Width           =   972
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "More >>>"
         Height          =   252
         Left            =   6000
         TabIndex        =   48
         Top             =   240
         Visible         =   0   'False
         Width           =   852
      End
   End
   Begin VB.Frame z2 
      Height          =   40
      Left            =   120
      TabIndex        =   43
      Top             =   1440
      Width           =   7212
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PackPSe05a.frx":31CC
      Style           =   1  'Graphical
      TabIndex        =   42
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdShip 
      Caption         =   "S&hip"
      Height          =   315
      Left            =   6480
      TabIndex        =   28
      ToolTipText     =   "Mark This Packing Slip As Shipped"
      Top             =   500
      Width           =   915
   End
   Begin VB.ComboBox txtShip 
      Height          =   315
      Left            =   3600
      TabIndex        =   1
      Tag             =   "4"
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton optPrn 
      DownPicture     =   "PackPSe05a.frx":397A
      Height          =   320
      Left            =   7000
      Picture         =   "PackPSe05a.frx":3F0C
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Print This Packing Slip"
      Top             =   960
      Width           =   350
   End
   Begin VB.TextBox txtDmy 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   960
      TabIndex        =   38
      Top             =   1035
      Width           =   75
   End
   Begin VB.ComboBox cmbPsl 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Form List Or Enter The Number (PS000099 or 99). Contains Not Shipped Packing Slips"
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6480
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7080
      Top             =   4920
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4935
      FormDesignWidth =   7455
   End
   Begin MSComctlLib.TabStrip tab1 
      Height          =   3372
      Left            =   12
      TabIndex        =   44
      Top             =   1560
      Width           =   7340
      _ExtentX        =   12965
      _ExtentY        =   5953
      TabWidthStyle   =   2
      TabFixedWidth   =   1587
      TabFixedHeight  =   473
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Shipping"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Load"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Billing"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shipped"
      Height          =   255
      Index           =   25
      Left            =   2880
      TabIndex        =   41
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblPson 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   480
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblPrn 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5400
      TabIndex        =   37
      Top             =   360
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Printed"
      Height          =   255
      Index           =   23
      Left            =   4800
      TabIndex        =   36
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3600
      TabIndex        =   35
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   34
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   33
      Top             =   1080
      Width           =   3555
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   32
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblCst 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   31
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Packing Slip"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   30
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "PackPSe05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'6/25/04 New
'1/7/05 Added optAppend and associated code
'8/10/06 Replaced Tab with TabStrip
'8/28/06 fixed Null (GetPackSlip)!PRIMARYSO
Option Explicit
Dim RdoPls As ADODB.Recordset
'Dim rdoQry As rdoQuery
 Dim cmdObj As ADODB.Command
Dim bCanceled As Byte
Dim bGoodPs As Byte
Dim bOnLoad As Byte
Dim bPrint As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub ClearCombos()
   Dim iControl As Integer
   For iControl = 0 To Controls.Count - 1
      If TypeOf Controls(iControl) Is ComboBox Then _
                         Controls(iControl).SelLength = 0
   Next
   
End Sub

Private Sub GetTerms()
   Dim RdoTrm As ADODB.Recordset
   Dim sTerms As String
   On Error GoTo DiaErr1
   sTerms = Compress(cmbTrm)
   sSql = "SELECT TRMDESC FROM StrmTable WHERE TRMREF='" & sTerms & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTrm, ES_FORWARD)
   If bSqlRows Then
      lblTrm = "" & Trim(RdoTrm!TRMDESC)
      ClearResultSet RdoTrm
   Else
      lblTrm = ""
   End If
   Set RdoTrm = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getterms"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDmy.BackColor = Es_FormBackColor
   txtBtName.BackColor = Es_TextDisabled
   txtBtAdr.BackColor = Es_TextDisabled
   txtBtCity.BackColor = Es_TextDisabled
   txtBtState.BackColor = Es_TextDisabled
   txtBtZip.BackColor = Es_TextDisabled
   txtShip = Format(ES_SYSDATE, "mm/dd/yyyy")
   
End Sub

Private Sub cmbPsl_Click()
   bGoodPs = GetPackslip()
   
End Sub


Private Sub cmbPsl_LostFocus()
   cmbPsl = CheckLen(cmbPsl, 8)
   If bCanceled Then Exit Sub
   If Val(cmbPsl) > 0 Then cmbPsl = Format(cmbPsl, "00000000")
   bGoodPs = GetPackslip()
   If Len(cmbPsl) Then
      If bGoodPs = 0 Then MsgBox "That Packing Slip Doesn't Exist " & vbCrLf _
                   & "Or Doesn't Qualify.", vbInformation, Caption
   End If
   
End Sub


Private Sub cmbTrm_Click()
   GetTerms
   
End Sub

Private Sub cmbTrm_LostFocus()
   cmbTrm = CheckLen(cmbTrm, 2)
   If Trim(cmbTrm) = "" Then
      If cmbTrm.ListCount > 0 Then
         Beep
         cmbTrm = cmbTrm.List(0)
      End If
   End If
   GetTerms
   If bGoodPs Then
      On Error Resume Next
      'RdoPls.Edit
      RdoPls!PSTERMS = cmbTrm
      RdoPls.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCanceled = 1
   
End Sub


Private Sub cmdComments_Click(Index As Integer)
   If Index = 0 Then
      If cmdComments(0) Then
         'See List For Index
         txtCmt.SetFocus
         SysComments.lblListIndex = 5
         SysComments.Show
         cmdComments(0) = False
      End If
   Else
      If cmdComments(1) Then
         'See List For Index
         txtSta.SetFocus
         SysComments.lblControl = "txtSta"
         SysComments.lblListIndex = 3
         SysComments.Show
         cmdComments(1) = False
      End If
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2205
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub



Private Sub cmdShip_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   If bGoodPs Then
      sMsg = "This Will Mark This Packing Slip As Having Been " & vbCrLf _
             & "Shipped. The Packing Slip Will Be Unavailable " & vbCrLf _
             & "For Further Edits. Continue?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         On Error Resume Next
         clsADOCon.ADOErrNum = 0
         'RdoPls.Edit
         RdoPls!PSSHIPPEDDATE = txtShip
         RdoPls!PSSHIPPED = 1
         RdoPls.Update
         sSql = "UPDATE SoitTable SET ITPSSHIPPED=1 " _
                & "WHERE ITPSNUMBER='" & Trim(cmbPsl) & "'"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
         If clsADOCon.ADOErrNum = 0 Then
            SysMsg "Item Marked As Shipped.", True
            bResponse = MsgBox("Reprint This Packing Slip?", _
                        ES_YESQUESTION, Caption)
            If bResponse = vbYes Then
               bPrint = 1
               PackPSp01a.optNot.Value = vbChecked
               PackPSp01a.Show
               Hide
               optPrn.Value = False
            End If
         Else
            MsgBox "Couldn't Mark As Shipped.", _
               vbInformation, Caption
         End If
      Else
         CancelTrans
      End If
   Else
      MsgBox "Requires A Valid Packing Slip.", _
         vbInformation, Caption
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      cmdComments(0).Enabled = True
      cmdComments(1).Enabled = True
      FillCombo
      FillTerms
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   Dim b As Byte
   FormLoad Me
   FormatControls
   
   For b = 0 To 3
      With tabFrame(b)
         .BorderStyle = 0
         .Visible = False
         .Left = 40
      End With
   Next
   tabFrame(0).Visible = True
   sSql = "SELECT * FROM PshdTable WHERE PSNUMBER= ? " _
          & "AND (PSPRINTED IS NOT NULL AND PSINVOICE=0 AND PSSHIPPED=0 )"
   Set cmdObj = New ADODB.Command
   cmdObj.CommandText = sSql
   
   'Set rdoQry = RdoCon.CreateQuery("", sSql)
   'rdoQry.MaxRows = 1
   Dim prmObj As ADODB.Parameter
   Set prmObj = New ADODB.Parameter
   prmObj.Type = adChar
   prmObj.Size = 8

   cmdObj.parameters.Append prmObj
   
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload bPrint
   Set cmdObj = Nothing
   Set RdoPls = Nothing
   Set PackPSe05a = Nothing
   
End Sub



Private Sub FillCombo()
'   On Error GoTo DiaErr1
'   cmbPsl.Clear
'   sSql = "SELECT PSNUMBER FROM PshdTable where  " _
'          & "(PSPRINTED IS NOT NULL AND PSINVOICE=0 AND PSSHIPPED=0)"
'   LoadComboBox cmbPsl, -1
'   If Not bSqlRows Then
'      MsgBox "No Qualifying Packing Slips Where Found.", _
'         vbInformation, Caption
'   Else
'      If cmbPsl.ListCount > 0 Then cmbPsl = cmbPsl.List(0)
'   End If
'   Exit Sub
'
'DiaErr1:
'   sProcName = "fillcombo"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'

   Dim ps As New ClassPackSlip
   ps.FillPSComboPrintedNotShipped cmbPsl

   If cmbPsl.ListCount = 0 Then
      MsgBox "No Qualifying Packing Slips Where Found.", vbInformation, Caption
   End If
End Sub



Private Function GetPackslip() As Byte
   On Error GoTo DiaErr1
   GetPackslip = False
  ' rdoQry.RowsetSize = 1
  ' rdoQry(0) = Compress(cmbPsl)
  ' bSqlRows = GetQuerySet(RdoPls, rdoQry, ES_KEYSET)
   
   cmdObj.parameters(0).Value = Compress(cmbPsl)
   bSqlRows = clsADOCon.GetQuerySet(RdoPls, cmdObj, ES_FORWARD, True)
 
   If bSqlRows Then
      With RdoPls
         lblDte = "" & Format(!PSDATE, "mm/dd/yyyy")
         lblPrn = "" & Format(!PSPRINTED, "mm/dd/yyyy")
         If Trim(!PSCUST) <> "" Then FindCustomer Me, Trim(!PSCUST), True
         txtStn = "" & Trim(!PSSTNAME)
         txtSta = "" & Trim(!PSSTADR)
         txtCmt = "" & Trim(!PSCOMMENTS)
         txtVia = "" & Trim(!PSVIA)
         txtBxs = "" & Trim(!PSCOMMENTS)
         txtCnt = "" & Trim(!PSCONTAINER)
         txtCrt = "" & Format(!PSCARTON, "###0")
         txtCuf = "" & Format(!PSCUFEET, "##0.000")
         txtFrt = "" & Format(!PSFREIGHT, ES_QuantityDataFormat)
         txtGir = "" & Format(!PSGIRTH, "##0.000")
         txtGrs = "" & Format(!PSGROSSLBS, ES_QuantityDataFormat)
         txtHgt = "" & Format(!PSHEIGHT, "##0.000")
         txtLen = "" & Format(!PSLENGTH, "##0.000")
         txtLod = "" & Trim(!PSLOADNO)
         txtNet = "" & Format(!PSNETLBS, ES_QuantityDataFormat)
         txtRem = "" & Trim(!PSLOADREM)
         txtSel = "" & Trim(!PSSEAL)
         txtShn = "" & Format(!PSSHIPNO, "###0")
         cmbTrm = "" & Trim(!PSTERMS)
         txtVia = "" & Trim(!PSVIA)
         txtWid = "" & Format(!PSWIDTH, "##0.000")
         txtTrk = "" & Trim(!PSTRACKING)
         If Not IsNull(!PSPRIMARYSO) Then _
                       lblPson = !PSPRIMARYSO Else lblPson = 0
         
      End With
      GetTerms
      GetPackslip = 1
   Else
      lblCst = ""
      lblNme = "*** Not Found Or Doesn't Qualify ***"
      lblDte = ""
      lblPrn = ""
      txtStn = ""
      txtSta = ""
      txtCmt = ""
      txtVia = ""
      txtBxs = ""
      txtCnt = ""
      txtCrt = ""
      txtCuf = ""
      txtFrt = ""
      txtGir = ""
      txtGrs = ""
      txtHgt = ""
      txtLen = ""
      txtLod = ""
      txtNet = ""
      txtRem = ""
      txtSel = ""
      txtShn = ""
      cmbTrm = ""
      txtVia = ""
      txtWid = ""
      FillCombo
      GetPackslip = 0
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getpacksl"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Sub lblNme_Change()
   If Left(lblNme, 10) = "*** Not Fo" Then
      lblNme.ForeColor = ES_RED
   Else
      lblNme.ForeColor = Es_TextForeColor
   End If
End Sub

Private Sub lblPson_Click()
   'tracks the primary sales order
   
End Sub



Private Sub optPrn_Click()
   If optPrn Then
      bPrint = 1
      PackPSp01a.optNot.Value = vbChecked
      PackPSp01a.Show
      Hide
      optPrn.Value = False
   End If
   
End Sub

Private Sub tab1_Click()
   Dim b As Byte
   On Error Resume Next
   ClearCombos
   For b = 0 To 3
      tabFrame(b).Visible = False
   Next
   tabFrame(tab1.SelectedItem.Index - 1).Visible = True
   Select Case tab1.SelectedItem.Index
      Case 1
         txtStn.SetFocus
      Case 2
         txtVia.SetFocus
      Case 3
         txtLod.SetFocus
   End Select
   
End Sub

Private Sub txtBxs_LostFocus()
   txtBxs = CheckLen(txtBxs, 4)
   txtBxs = Format(Abs(Val(txtBxs)), "###0")
   If bGoodPs Then
      On Error Resume Next
      'RdoPls.Edit
      RdoPls!PSBOXES = Val(txtBxs)
      RdoPls.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 2040)
   If bGoodPs Then
      On Error Resume Next
      'RdoPls.Edit
      RdoPls!PSCOMMENTS = "" & txtCmt
      RdoPls.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtCnt_LostFocus()
   txtCnt = CheckLen(txtCnt, 12)
   If bGoodPs Then
      On Error Resume Next
      'RdoPls.Edit
      RdoPls!PSCONTAINER = "" & txtCnt
      RdoPls.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtCrt_LostFocus()
   txtCrt = CheckLen(txtCrt, 4)
   txtCrt = Format(Abs(Val(txtCrt)), "###0")
   If bGoodPs Then
      On Error Resume Next
      'RdoPls.Edit
      RdoPls!PSCARTON = Val(txtCrt)
      RdoPls.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtCuf_LostFocus()
   txtCuf = CheckLen(txtCuf, 7)
   txtCuf = Format(Abs(Val(txtCuf)), "##0.000")
   If bGoodPs Then
      On Error Resume Next
      'RdoPls.Edit
      RdoPls!PSCUFEET = Val(txtCuf)
      RdoPls.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtFrt_LostFocus()
   txtFrt = CheckLen(txtFrt, 9)
   txtFrt = Format(Abs(Val(txtFrt)), ES_QuantityDataFormat)
   If bGoodPs Then
      On Error Resume Next
      'RdoPls.Edit
      RdoPls!PSFREIGHT = Val(txtFrt)
      RdoPls.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtGir_LostFocus()
   txtGir = CheckLen(txtGir, 7)
   txtGir = Format(Abs(Val(txtGir)), "##0.000")
   If bGoodPs Then
      On Error Resume Next
      'RdoPls.Edit
      RdoPls!PSGIRTH = Val(txtGir)
      RdoPls.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtGrs_LostFocus()
   txtGrs = CheckLen(txtGrs, 9)
   txtGrs = Format(Abs(Val(txtGrs)), ES_QuantityDataFormat)
   If bGoodPs Then
      On Error Resume Next
      'RdoPls.Edit
      RdoPls!PSGROSSLBS = Val(txtGrs)
      RdoPls.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtHgt_LostFocus()
   txtHgt = CheckLen(txtHgt, 7)
   txtHgt = Format(Abs(Val(txtHgt)), "##0.000")
   If bGoodPs Then
      On Error Resume Next
      'RdoPls.Edit
      RdoPls!PSHEIGHT = Val(txtHgt)
      RdoPls.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtLen_LostFocus()
   txtLen = CheckLen(txtLen, 7)
   txtLen = Format(Abs(Val(txtLen)), "##0.000")
   If bGoodPs Then
      On Error Resume Next
      'RdoPls.Edit
      RdoPls!PSLENGTH = Val(txtLen)
      RdoPls.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtLod_LostFocus()
   txtLod = CheckLen(txtLod, 12)
   If bGoodPs Then
      On Error Resume Next
      'RdoPls.Edit
      RdoPls!PSLOADNO = "" & txtLod
      RdoPls!PSREVISED = Format(ES_SYSDATE, "mm/dd/yyyy")
      RdoPls.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtNet_LostFocus()
   txtNet = CheckLen(txtNet, 9)
   txtNet = Format(Abs(Val(txtNet)), ES_QuantityDataFormat)
   If bGoodPs Then
      On Error Resume Next
      'RdoPls.Edit
      RdoPls!PSNETLBS = Val(txtNet)
      RdoPls.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtRem_LostFocus()
   txtRem = CheckLen(txtRem, 255)
   If bGoodPs Then
      On Error Resume Next
      'RdoPls.Edit
      RdoPls!PSLOADREM = "" & txtRem
      RdoPls.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtSel_LostFocus()
   txtSel = CheckLen(txtSel, 12)
   If bGoodPs Then
      On Error Resume Next
      'RdoPls.Edit
      RdoPls!PSSEAL = "" & txtSel
      RdoPls.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtShip_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtShip_LostFocus()
   txtShip = CheckDateEx(txtShip)
   
End Sub


Private Sub txtShn_LostFocus()
   txtShn = CheckLen(txtShn, 4)
   txtShn = Format(Abs(Val(txtShn)), "###0")
   If bGoodPs Then
      On Error Resume Next
      'RdoPls.Edit
      RdoPls!PSSHIPNO = Val(txtShn)
      RdoPls.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtSta_LostFocus()
   txtSta = CheckLen(txtSta, 255)
   If bGoodPs Then
      On Error Resume Next
      'RdoPls.Edit
      RdoPls!PSSTADR = "" & txtSta
      RdoPls.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtStn_LostFocus()
   txtStn = CheckLen(txtStn, 40)
   txtStn = StrCase(txtStn)
   If bGoodPs Then
      On Error Resume Next
      'RdoPls.Edit
      RdoPls!PSSTNAME = "" & txtStn
      RdoPls!PSREVISED = Format(ES_SYSDATE, "mm/dd/yyyy")
      RdoPls.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub



Private Sub txtTrk_LostFocus()
   txtTrk = CheckLen(txtTrk, 40)
   If bGoodPs Then
      On Error Resume Next
      'RdoPls.Edit
      RdoPls!PSTRACKING = txtTrk
      RdoPls.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtVia_LostFocus()
   txtVia = CheckLen(txtVia, 16)
   If bGoodPs Then
      On Error Resume Next
      'RdoPls.Edit
      RdoPls!PSVIA = txtVia
      RdoPls!PSREVISED = Format(ES_SYSDATE, "mm/dd/yyyy")
      RdoPls.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtWid_LostFocus()
   txtWid = CheckLen(txtWid, 7)
   txtWid = Format(Abs(Val(txtWid)), "##0.000")
   If bGoodPs Then
      On Error Resume Next
      'RdoPls.Edit
      RdoPls!PSWIDTH = Val(txtWid)
      RdoPls.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub



Private Sub GetSalesOrder()
   Dim RdoSon As ADODB.Recordset
   On Error GoTo DiaErr1
   If Val(lblPson) > 0 Then
      sSql = "Qry_GetSalesOrderText " & Val(lblPson) & " "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoSon, ES_FORWARD)
      If bSqlRows Then
         With RdoSon
            lblPson = "" & Trim(!SOTYPE) & Trim(!SOTEXT)
            ClearResultSet RdoSon
         End With
      End If
   End If
   Set RdoSon = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getsalesor"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
