VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form PackPSe02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revise A Packing Slip"
   ClientHeight    =   4980
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   7395
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   2202
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   Begin VB.Frame tabFrame 
      Height          =   2892
      Index           =   1
      Left            =   5000
      TabIndex        =   51
      Top             =   1920
      Width           =   6972
      Begin VB.TextBox txtTrk 
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Tag             =   "3"
         ToolTipText     =   "Tracking Number (Free Form To 40 Chars)"
         Top             =   2040
         Width           =   3555
      End
      Begin VB.ComboBox cmbTrm 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Tag             =   "8"
         ToolTipText     =   "Select Terms From List"
         Top             =   1560
         Width           =   780
      End
      Begin VB.TextBox txtCrt 
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Tag             =   "1"
         Top             =   1200
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
      Begin VB.TextBox txtBxs 
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Tag             =   "1"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtFrt 
         Height          =   285
         Left            =   5040
         TabIndex        =   8
         Tag             =   "1"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtCnt 
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Tag             =   "3"
         Top             =   480
         Width           =   1675
      End
      Begin VB.TextBox txtSel 
         Height          =   285
         Left            =   5040
         TabIndex        =   6
         Tag             =   "3"
         Top             =   120
         Width           =   1675
      End
      Begin VB.TextBox txtVia 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Tag             =   "6"
         Top             =   120
         Width           =   1875
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tracking No"
         Height          =   252
         Index           =   24
         Left            =   120
         TabIndex        =   61
         Top             =   2040
         Width           =   1212
      End
      Begin VB.Label lblTrm 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2685
         TabIndex        =   60
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Shipping Terms"
         Height          =   252
         Index           =   13
         Left            =   120
         TabIndex        =   59
         Top             =   1560
         Width           =   1212
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Carton"
         Height          =   252
         Index           =   12
         Left            =   120
         TabIndex        =   58
         Top             =   1200
         Width           =   1092
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Shipper Number"
         Height          =   252
         Index           =   11
         Left            =   3720
         TabIndex        =   57
         Top             =   840
         Width           =   1332
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Boxes"
         Height          =   252
         Index           =   10
         Left            =   120
         TabIndex        =   56
         Top             =   840
         Width           =   1092
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Freight"
         Height          =   252
         Index           =   9
         Left            =   3720
         TabIndex        =   55
         Top             =   480
         Width           =   1092
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Container"
         Height          =   252
         Index           =   8
         Left            =   120
         TabIndex        =   54
         Top             =   480
         Width           =   1092
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Seal"
         Height          =   252
         Index           =   7
         Left            =   5040
         TabIndex        =   53
         Top             =   480
         Width           =   1092
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ship Via"
         Height          =   252
         Index           =   6
         Left            =   120
         TabIndex        =   52
         Top             =   120
         Width           =   1092
      End
   End
   Begin VB.CheckBox cbRefreshPS 
      Caption         =   "cbRefreshPS"
      Height          =   255
      Left            =   5040
      TabIndex        =   76
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   3720
      TabIndex        =   75
      Tag             =   "4"
      Top             =   360
      Width           =   1215
   End
   Begin VB.Frame tabFrame 
      Height          =   2892
      Index           =   3
      Left            =   8160
      TabIndex        =   72
      Top             =   1920
      Width           =   6972
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
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bill To:"
         Height          =   252
         Index           =   25
         Left            =   120
         TabIndex        =   73
         Top             =   120
         Width           =   1092
      End
   End
   Begin VB.Frame tabFrame 
      Height          =   2892
      Index           =   2
      Left            =   120
      TabIndex        =   62
      Top             =   2040
      Width           =   6972
      Begin VB.TextBox txtLen 
         Height          =   285
         Left            =   3600
         TabIndex        =   19
         Tag             =   "1"
         Top             =   840
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
      Begin VB.TextBox txtGir 
         Height          =   285
         Left            =   1680
         TabIndex        =   21
         Tag             =   "1"
         Top             =   1200
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
      Begin VB.TextBox txtHgt 
         Height          =   285
         Left            =   1680
         TabIndex        =   18
         Tag             =   "1"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtCuf 
         Height          =   285
         Left            =   5400
         TabIndex        =   17
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
      Begin VB.TextBox txtNet 
         Height          =   285
         Left            =   3600
         TabIndex        =   16
         Tag             =   "1"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtGrs 
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Tag             =   "1"
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Load Remarks:"
         Height          =   252
         Index           =   22
         Left            =   120
         TabIndex        =   71
         Top             =   1560
         Width           =   1212
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Girth"
         Height          =   252
         Index           =   21
         Left            =   120
         TabIndex        =   70
         Top             =   1200
         Width           =   732
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Width"
         Height          =   252
         Index           =   20
         Left            =   4680
         TabIndex        =   69
         Top             =   840
         Width           =   852
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Length"
         Height          =   252
         Index           =   19
         Left            =   2880
         TabIndex        =   68
         Top             =   840
         Width           =   732
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Height"
         Height          =   252
         Index           =   18
         Left            =   120
         TabIndex        =   67
         Top             =   840
         Width           =   972
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cu Feet"
         Height          =   252
         Index           =   17
         Left            =   4680
         TabIndex        =   66
         Top             =   480
         Width           =   1092
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Load No"
         Height          =   252
         Index           =   16
         Left            =   120
         TabIndex        =   65
         Top             =   120
         Width           =   1092
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Lbs"
         Height          =   252
         Index           =   15
         Left            =   2880
         TabIndex        =   64
         Top             =   480
         Width           =   732
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
   End
   Begin VB.Frame tabFrame 
      Height          =   2892
      Index           =   0
      Left            =   5000
      TabIndex        =   47
      Top             =   1920
      Width           =   6972
      Begin VB.CommandButton cmdComments 
         DisabledPicture =   "PackPSe02a.frx":0000
         DownPicture     =   "PackPSe02a.frx":0972
         Height          =   350
         Index           =   1
         Left            =   5280
         Picture         =   "PackPSe02a.frx":12E4
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Standard Comments"
         Top             =   480
         Width           =   350
      End
      Begin VB.CommandButton cmdComments 
         DisabledPicture =   "PackPSe02a.frx":173C
         DownPicture     =   "PackPSe02a.frx":20AE
         Height          =   350
         Index           =   0
         Left            =   1200
         Picture         =   "PackPSe02a.frx":2A20
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Standard Comments"
         Top             =   1560
         Width           =   350
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
      Begin VB.TextBox txtSta 
         Height          =   975
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   2
         Tag             =   "9"
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox txtStn 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Tag             =   "2"
         Top             =   120
         Width           =   3475
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "More >>>"
         Height          =   252
         Left            =   6000
         TabIndex        =   74
         Top             =   240
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   252
         Index           =   5
         Left            =   120
         TabIndex        =   50
         Top             =   1800
         Width           =   972
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ship To Address"
         Height          =   252
         Index           =   4
         Left            =   120
         TabIndex        =   49
         Top             =   480
         Width           =   1332
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ship To Name"
         Height          =   252
         Index           =   3
         Left            =   120
         TabIndex        =   48
         Top             =   120
         Width           =   1332
      End
   End
   Begin MSComctlLib.TabStrip tab1 
      Height          =   3372
      Left            =   10
      TabIndex        =   46
      Top             =   1560
      Width           =   7092
      _ExtentX        =   12515
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
   Begin VB.Frame z2 
      ForeColor       =   &H8000000F&
      Height          =   30
      Left            =   0
      TabIndex        =   45
      Top             =   1440
      Width           =   6972
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PackPSe02a.frx":2E78
      Style           =   1  'Graphical
      TabIndex        =   44
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton optPrn 
      DownPicture     =   "PackPSe02a.frx":3626
      Height          =   320
      Left            =   6600
      Picture         =   "PackPSe02a.frx":37B0
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Print This Packing Slip"
      Top             =   960
      Width           =   350
   End
   Begin VB.TextBox txtDmy 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1200
      TabIndex        =   41
      Top             =   1035
      Width           =   75
   End
   Begin VB.CheckBox optItm 
      Caption         =   "Items"
      Height          =   195
      Left            =   1440
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox optNew 
      Caption         =   "New Slip"
      Height          =   255
      Left            =   360
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdItm 
      Caption         =   "&Items"
      Height          =   315
      Left            =   6120
      TabIndex        =   34
      TabStop         =   0   'False
      ToolTipText     =   "Retrieve/Add Pack Slip Items"
      Top             =   500
      Width           =   915
   End
   Begin VB.ComboBox cmbPsl 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Form List Or Enter The Number (PS000099 or 99) Contains Top 1000 Desc"
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6120
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   -120
      Top             =   960
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4980
      FormDesignWidth =   7395
   End
   Begin VB.Label lblPson 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5040
      TabIndex        =   43
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblPrn 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3720
      TabIndex        =   40
      Top             =   720
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Printed"
      Height          =   255
      Index           =   23
      Left            =   3120
      TabIndex        =   39
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3720
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   35
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   33
      Top             =   1035
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
      Left            =   1320
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
      Width           =   975
   End
End
Attribute VB_Name = "PackPSe02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'11/5/03 Added tracking
'4/11/05 Descending sort and last 300 Pack Slips
'8/28/06 Fixed invalid Null (GetPackSlip)
Option Explicit
Dim RdoPls As ADODB.Recordset

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
   
End Sub

Private Sub cmbPsl_Click()
   bGoodPs = GetPackslip()
   
End Sub


Private Sub cmbPsl_LostFocus()
   cmbPsl = CheckLen(cmbPsl, 8)
   If bCanceled Then Exit Sub
   'If Val(cmbPsl) > 0 Then cmbPsl = "PS" & Format(cmbPsl, "000000")
   bGoodPs = GetPackslip()
   If Len(cmbPsl) Then
      If bGoodPs = 0 Then
         MsgBox "That Packing Slip Doesn't Exist " & vbCrLf _
            & "Or It Has Been Canceled.", vbInformation, Caption
      Else
         GetBillTo
      End If
   End If
   
End Sub


Private Sub cmbTrm_Click()
   GetTerms
   
End Sub

Private Sub cmbTrm_DropDown()
   bDataHasChanged = True
   
End Sub

Private Sub cmbTrm_LostFocus()
   ' MM cmbTrm = CheckLen(cmbTrm, 2)
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
      OpenHelpContext 2202
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdItm_Click()
   optItm.Value = vbChecked
   PackPSe02b.lblCst = lblCst
   PackPSe02b.lblNme = lblNme
   PackPSe02b.Show
   
End Sub

Private Sub Form_Activate()
   If bOnLoad Then
      FillCombo
   End If
   
   If optNew.Value = vbChecked Then
      Caption = "Edit A New Packing Slip"
      cmbPsl = PackPSe01a.lblPrefix & PackPSe01a.txtPsl
      Unload PackPSe01a
      bGoodPs = GetPackslip()
      optNew.Value = vbUnchecked
   Else
      Caption = "Revise A Packing Slip"
   End If
   If optItm.Value = vbChecked Then
      Unload PackPSe02b
      optItm.Value = vbUnchecked
      bGoodPs = GetPackslip()
   End If
   If bOnLoad Then
      cmdComments(0).Enabled = True
      cmdComments(1).Enabled = True
      'FillCombo  - do it before setting value
      FillTerms
      bOnLoad = 0
   End If
   If cbRefreshPS.Value = vbChecked Then
        bGoodPs = GetPackslip()
        cbRefreshPS.Value = vbUnchecked
        SysMsg "Pack Slip Boxes and Gross Pounds have been updated", 1
   End If
   MdiSect.lblBotPanel = Caption
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
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If bGoodPs And bDataHasChanged Then
      'RdoPls.Edit
      RdoPls!PSTERMS = cmbTrm
      RdoPls.Update
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload bPrint
   Set RdoPls = Nothing
   Set PackPSe02a = Nothing
   
End Sub



Private Sub FillCombo()
   Dim RdoCmb As ADODB.Recordset
   Dim iRows As Integer

'   On Error Resume Next
'   sSql = "UPDATE PshdTable set PSPRIMARYSO=0 WHERE PSPRIMARYSO IS NULL"
'   RdoCon.Execute sSql, rdExecDirect

'   sSql = "SELECT PSNUMBER FROM PshdTable" & vbCrLf _
'      & "WHERE PSDATE > '" & DateAdd("yyyy", -2, Now) & "'" & vbCrLf _
'      & "ORDER BY PSNUMBER DESC"
'
'   On Error GoTo DiaErr1
'   sSql = "SELECT PSNUMBER FROM PshdTable" & vbCrLf _
'      & "WHERE PSDATE > '" & DateAdd("yyyy", -2, Now) & "'" & vbCrLf _
'      & "ORDER BY PSNUMBER DESC"
'   'sSql = "Qry_FillPackSlips '" & DateAdd("yyyy", -3, Now) & "'"
'   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
'   If bSqlRows Then
'      With RdoCmb
'         iRows = 0
'         Do Until .EOF
'            iRows = iRows + 1
'            AddComboStr cmbPsl.hWnd, "" & Trim(.Fields(0))
'            'If iRows > 999 Then Exit Do
'            .MoveNext
'         Loop
'         ClearResultSet RdoCmb
'      End With
'   End If
'   Set RdoCmb = Nothing
'   If Trim(cmbPsl) = "" Then _
'           If cmbPsl.ListCount > 0 Then cmbPsl = cmbPsl.List(0)
'   If cmbPsl.ListCount > 0 Then bGoodPs = GetPackslip()
'   Exit Sub
'
'DiaErr1:
'   sProcName = "fillcombo"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me

   Dim ps As New ClassPackSlip
   'ps.FillPSComboUnprinted cmbPsl
   ps.FillPSComboAll cmbPsl
   If cmbPsl.ListCount > 0 Then bGoodPs = GetPackslip()

'   Dim ps As New ClassPackSlip
'   ps.FillPSComboAll cmbPsl
'   If cmbPsl.ListCount > 0 Then bGoodPs = GetPackslip()
   
End Sub



Private Function GetPackslip() As Byte
   On Error Resume Next
   If bGoodPs And bDataHasChanged Then
      'RdoPls.Edit
      RdoPls!PSTERMS = cmbTrm
      RdoPls.Update
      bDataHasChanged = False
   End If
   On Error GoTo DiaErr1
   Dim Index As Integer
   GetPackslip = False
   sSql = "SELECT * FROM PshdTable WHERE PSNUMBER='" & Trim(cmbPsl) & "' " _
          & "AND (PSTYPE=1 AND PSCANCELED=0)"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPls, ES_KEYSET)
   If bSqlRows Then
      With RdoPls
         'lblDte = "" & Format(!PSDATE, "mm/dd/yy")
         txtDte = "" & Format(!PSDATE, "mm/dd/yyyy")
         lblPrn = "" & Format(!PSPRINTED, "mm/dd/yyyy")
         If Trim(!PSCUST) <> "" Then FindCustomer Me, Trim(!PSCUST), True
         txtStn = "" & Trim(!PSSTNAME)
         txtSta = "" & Trim(!PSSTADR)
         txtCmt = "" & Trim(!PSCOMMENTS)
         txtVia = "" & Trim(!PSVIA)
         txtBxs = "" & Trim(!PSBOXES)
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
         Index = GetListIndex(cmbTrm, Trim(!PSTERMS))
         
         cmbTrm.ListIndex = Index
         
         txtVia = "" & Trim(!PSVIA)
         txtWid = "" & Format(!PSWIDTH, "##0.000")
         txtTrk = "" & Trim(!PSTRACKING)
         If Not IsNull(!PSPRIMARYSO) Then _
                       lblPson = !PSPRIMARYSO Else lblPson = 0
         GetSalesOrder
         If Trim(lblPrn) = "" Then
            txtDmy.Enabled = False
            cmdItm.Enabled = True
         Else
            txtDmy.Enabled = True
            cmdItm.Enabled = False
         End If
         ClearResultSet RdoPls
      End With
      GetTerms
      GetPackslip = 1
   Else
      lblCst = ""
      lblNme = "*** No Current Packslip ***"
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
      cmbTrm.ListIndex = -1
      txtVia = ""
      txtWid = ""
      GetPackslip = 0
   End If
   bDataHasChanged = False
   Exit Function
   
DiaErr1:
   sProcName = "getpacksl"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub lblNme_Change()
   If Left(lblNme, 10) = "*** No Cur" Then
      lblNme.ForeColor = ES_RED
   Else
      lblNme.ForeColor = Es_TextForeColor
   End If
End Sub

Private Sub lblPson_Click()
   'tracks the primary sales order
   
End Sub

Private Sub optItm_Click()
   'Never visible. Used to flag PackPSe02b as open
   
End Sub

Private Sub optNew_Click()
   'Never visible. Checked if new PS
   
End Sub

Private Sub optPrn_Click()
   If optPrn Then
      bPrint = 1
      PackPSp01a.Show
      Hide
      optPrn.Value = False
      PackPSp01a.optRev.Value = vbChecked
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
      RdoPls!PSBOXES = Val(txtBxs)
      RdoPls.Update
      If Err > 0 Then ValidateEdit
       
      PackPSe02d.lblPSNumber.Caption = cmbPsl
      If Val(txtBxs.Text) = 0 Then PackPSe02d.lblTotalBoxes = "1" Else PackPSe02d.lblTotalBoxes.Caption = txtBxs.Text
      PackPSe02d.lblFreight = txtFrt
      PackPSe02d.Show

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

Private Sub txtDte_DropDown()
   ShowCalendarEx Me
   bDataHasChanged = True
End Sub

Private Sub txtDte_LostFocus()

   Dim strJournalID As String
   txtDte = CheckDateEx(txtDte)
   '   Larry 5/13/00
   '   If Format(txtDte, "mm/dd/yy") < format(es_sysdate, "mm/dd/yy") Then
   '       MsgBox "The Date May Not Be Retroactive.", vbExclamation, Caption
   '       txtDte = format(es_sysdate, "mm/dd/yy")
   '   End If
' 4/16/2010 Users are allowed to create PS for previous week if journal is open.
   'CheckPeriodDate
   
   strJournalID = GetOpenJournal("IJ", Format(txtDte, "mm/dd/yy"))
   If Left(strJournalID, 4) = "None" Or (strJournalID = "") Then
      MsgBox "There Is No Open Inventory Journal For This" & vbCrLf _
         & "Period. Cannot Set Pack Slip date.", _
         vbExclamation, Caption
      
      txtDte = Format(ES_SYSDATE, "mm/dd/yyyy")
   End If
   
   If bGoodPs And bDataHasChanged Then
      'RdoPls.Edit
      RdoPls!PSDATE = txtDte
      RdoPls.Update
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
      RdoPls!PSREVISED = Format(ES_SYSDATE, "mm/dd/yy")
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
      RdoPls!PSREVISED = Format(ES_SYSDATE, "mm/dd/yy")
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
   txtVia = CheckLen(txtVia, 30)
   If bGoodPs Then
      On Error Resume Next
      'RdoPls.Edit
      RdoPls!PSVIA = txtVia
      RdoPls!PSREVISED = Format(ES_SYSDATE, "mm/dd/yy")
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
            ' MM no need to append the type
            'lblPson = "" & Trim(!SOTYPE) & Trim(!SOTEXT)
            lblPson = "" & Trim(!SOTEXT)
            
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

Private Function GetListIndex(cmbTrm As ComboBox, strSearch As String) As Double
   
   Dim bFound As Boolean
   Dim i As Integer
   Dim strCur As String
   bFound = False
   For i = 0 To cmbTrm.ListCount - 1
      If (cmbTrm.List(i) = strSearch) Then
         bFound = True
         Exit For
      End If
   Next
   
   If (bFound) Then
      GetListIndex = i
   Else
      GetListIndex = -1
   End If
   
End Function


Private Sub GetBillTo()
   Dim RdoBill As ADODB.Recordset
   sSql = "SELECT CUREF,CUBTNAME,CUBTADR,CUBCITY,CUBSTATE," _
          & "CUBZIP FROM CustTable WHERE CUREF='" & Compress(lblCst) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBill, ES_FORWARD)
   If bSqlRows Then
      With RdoBill
         txtBtName = "" & Trim(!CUBTNAME)
         txtBtAdr = "" & Trim(!CUBTADR)
         txtBtCity = "" & Trim(!CUBCITY)
         txtBtState = "" & Trim(!CUBSTATE)
         txtBtZip = "" & Trim(!CUBZIP)
         ClearResultSet RdoBill
      End With
   Else
      txtBtName = ""
      txtBtAdr = ""
      txtBtCity = ""
      txtBtState = ""
      txtBtZip = ""
   End If
   
   Set RdoBill = Nothing
End Sub
