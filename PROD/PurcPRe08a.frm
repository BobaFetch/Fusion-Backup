VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form PurcPRe08a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manufacturers"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Parts"
      Height          =   315
      Left            =   6240
      TabIndex        =   49
      ToolTipText     =   "Show Part Numbers To Be Assigned This Manufacturer"
      Top             =   480
      Width           =   875
   End
   Begin VB.TextBox txtType 
      Height          =   288
      Left            =   6720
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Type (2 Char)"
      Top             =   840
      Width           =   372
   End
   Begin VB.Frame tabFrame 
      Height          =   4092
      Index           =   1
      Left            =   7440
      TabIndex        =   28
      Top             =   1560
      Width           =   7116
      Begin VB.TextBox txtCmt 
         Height          =   855
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   18
         Tag             =   "9"
         Top             =   2160
         Width           =   5160
      End
      Begin VB.TextBox txtPDue 
         Height          =   285
         Left            =   3000
         TabIndex        =   17
         Tag             =   "1"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtPDate 
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Tag             =   "1"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtDday 
         Height          =   285
         Left            =   3000
         TabIndex        =   15
         Tag             =   "1"
         Text            =   "0"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtNDays 
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Tag             =   "1"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtDisc 
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Tag             =   "1"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Comments:"
         Height          =   288
         Index           =   12
         Left            =   120
         TabIndex        =   46
         Top             =   2160
         Width           =   1032
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Prox Due"
         Height          =   288
         Index           =   19
         Left            =   2160
         TabIndex        =   45
         Top             =   1680
         Width           =   1032
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Prox Date"
         Height          =   288
         Index           =   18
         Left            =   120
         TabIndex        =   44
         Top             =   1680
         Width           =   1032
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Or Terms:"
         Height          =   288
         Index           =   17
         Left            =   120
         TabIndex        =   43
         Top             =   1320
         Width           =   1032
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Due By"
         Height          =   288
         Index           =   16
         Left            =   2160
         TabIndex        =   42
         Top             =   720
         Width           =   1032
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Days"
         Height          =   288
         Index           =   15
         Left            =   120
         TabIndex        =   41
         Top             =   720
         Width           =   1032
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   288
         Index           =   14
         Left            =   2040
         TabIndex        =   40
         Top             =   240
         Width           =   432
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         Height          =   288
         Index           =   13
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   1032
      End
   End
   Begin VB.Frame tabFrame 
      Height          =   4092
      Index           =   0
      Left            =   40
      TabIndex        =   27
      Top             =   1560
      Width           =   7116
      Begin VB.TextBox txtSEmail 
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Tag             =   "2"
         ToolTipText     =   "Double Click Here To Send E-Mail (Requires An Entry)"
         Top             =   3360
         Width           =   5300
      End
      Begin VB.TextBox txtBEmail 
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Tag             =   "2"
         ToolTipText     =   "Double Click Here To Send E-Mail (Requires An Entry)"
         Top             =   2040
         Width           =   5300
      End
      Begin VB.TextBox txtSExt 
         Height          =   285
         Left            =   3480
         TabIndex        =   11
         Tag             =   "1"
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox txtSPhone 
         Height          =   288
         Left            =   1320
         TabIndex        =   10
         Tag             =   "1"
         Top             =   3000
         Width           =   1692
      End
      Begin VB.TextBox txtSCont 
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Tag             =   "2"
         ToolTipText     =   "Service Contact (20)"
         Top             =   2640
         Width           =   2535
      End
      Begin VB.TextBox txtFax 
         Height          =   288
         Left            =   4920
         TabIndex        =   7
         Tag             =   "1"
         Top             =   1680
         Width           =   1692
      End
      Begin VB.TextBox txtBExt 
         Height          =   285
         Left            =   3480
         TabIndex        =   6
         Tag             =   "1"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtBPhone 
         Height          =   288
         Left            =   1320
         TabIndex        =   5
         Tag             =   "1"
         Top             =   1680
         Width           =   1692
      End
      Begin VB.TextBox txtBCont 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Tag             =   "2"
         ToolTipText     =   "Buyer Contact (20)"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtAdr 
         Height          =   855
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   3
         Tag             =   "9"
         Top             =   240
         Width           =   3475
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail"
         Height          =   288
         Index           =   21
         Left            =   240
         TabIndex        =   48
         Top             =   3360
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail"
         Height          =   288
         Index           =   20
         Left            =   240
         TabIndex        =   47
         Top             =   2040
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         Height          =   288
         Index           =   11
         Left            =   240
         TabIndex        =   38
         Top             =   3000
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ext"
         Height          =   288
         Index           =   10
         Left            =   3120
         TabIndex        =   37
         Top             =   3000
         Width           =   432
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact"
         Height          =   288
         Index           =   5
         Left            =   240
         TabIndex        =   36
         Top             =   2640
         Width           =   1032
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Service:"
         Height          =   288
         Index           =   4
         Left            =   120
         TabIndex        =   35
         Top             =   2400
         Width           =   1032
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Buyer:"
         Height          =   288
         Index           =   3
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         Width           =   1032
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
         Height          =   288
         Index           =   7
         Left            =   4320
         TabIndex        =   33
         Top             =   1680
         Width           =   708
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         Height          =   288
         Index           =   6
         Left            =   240
         TabIndex        =   32
         Top             =   1680
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ext"
         Height          =   288
         Index           =   8
         Left            =   3120
         TabIndex        =   31
         Top             =   1680
         Width           =   432
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact"
         Height          =   288
         Index           =   9
         Left            =   240
         TabIndex        =   30
         Top             =   1320
         Width           =   1032
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   288
         Index           =   2
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1032
      End
   End
   Begin VB.Frame z2 
      Height          =   70
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   1080
      Width           =   7035
   End
   Begin MSComctlLib.TabStrip tab1 
      Height          =   4572
      Left            =   0
      TabIndex        =   25
      Top             =   1200
      Width           =   7212
      _ExtentX        =   12726
      _ExtentY        =   8070
      TabWidthStyle   =   2
      TabFixedWidth   =   2290
      TabFixedHeight  =   473
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Other"
            ImageVarType    =   2
         EndProperty
      EndProperty
      Enabled         =   0   'False
   End
   Begin VB.TextBox txtManu 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Tag             =   "2"
      ToolTipText     =   "Manufacturer's Name (30)"
      Top             =   840
      Width           =   3360
   End
   Begin VB.ComboBox cmbMfr 
      Height          =   288
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Manufacturer or Enter a New Manufacturer (10 Char Max)"
      Top             =   480
      Width           =   1555
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRe08a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6240
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7200
      Top             =   5640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5790
      FormDesignWidth =   7230
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   24
      Top             =   720
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nickname"
      Height          =   288
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   480
      Width           =   1428
   End
   Begin VB.Label lblNum 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   3120
      TabIndex        =   22
      Top             =   480
      Width           =   324
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturer Type"
      Height          =   288
      Index           =   32
      Left            =   5160
      TabIndex        =   21
      Top             =   840
      Width           =   1632
   End
End
Attribute VB_Name = "PurcPRe08a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'10/17/06 New
Option Explicit
Dim RdoMfr As ADODB.Recordset

Dim bCancel As Byte
Dim bOnLoad As Byte
Dim bGoodMfr As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd



Private Sub cmbMfr_Change()
   If Len(cmbMfr) > 10 Then cmbMfr = Left(cmbMfr, 10)
   
End Sub

Private Sub cmbMfr_Click()
   bGoodMfr = GetManufacturer()
   
End Sub

Private Sub cmbMfr_LostFocus()
   cmbMfr = CheckLen(cmbMfr, 10)
   If bCancel = 0 Then
      bGoodMfr = GetManufacturer()
      If bGoodMfr = 0 Then AddManufacturer
   End If
   
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
      OpenHelpContext 4309
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdUpd_Click()
   PurcPRe08b.Show
   
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
   tabFrame(0).BorderStyle = 0
   tabFrame(1).BorderStyle = 0
   tabFrame(0).Visible = True
   tabFrame(1).Visible = False
   tabFrame(0).Left = 40
   tabFrame(1).Left = 40
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set RdoMfr = Nothing
   Set PurcPRe08a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT MFGR_REF,MFGR_NICKNAME FROM MfgrTable "
   LoadComboBox cmbMfr
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub tab1_Click()
   On Error Resume Next
   If tab1.SelectedItem.Index = 1 Then
      tabFrame(0).Visible = True
      tabFrame(1).Visible = False
      txtAdr.SetFocus
   Else
      If bGoodMfr Then
         RdoMfr!MFGR_REVISED = Format(Now, "mm/dd/yy")
         RdoMfr.Update
      End If
      tabFrame(1).Visible = True
      tabFrame(0).Visible = False
      txtDisc.SetFocus
   End If
   
End Sub



Private Sub txtAdr_LostFocus()
   txtAdr = CheckLen(txtAdr, 30)
   If bGoodMfr Then
      On Error Resume Next
      RdoMfr!MFGR_ADDRESS = Trim(txtAdr)
      RdoMfr.Update
   End If
   
End Sub


Private Sub txtBCont_LostFocus()
   txtBCont = CheckLen(txtBCont, 20)
   txtBCont = StrCase(txtBCont)
   If bGoodMfr Then
      On Error Resume Next
      RdoMfr!MFGR_BCONTACT = Trim(txtBCont)
      RdoMfr!MFGR_REVISED = Format(Now, "mm/dd/yy")
      RdoMfr.Update
   End If
   
End Sub


Private Sub txtBEmail_DblClick()
   If Trim(txtBEmail) <> "" Then SendEMail Trim(txtBEmail)
   
End Sub


Private Sub txtBEmail_LostFocus()
   txtBEmail = CheckLen(txtBEmail, 60)
   If bGoodMfr Then
      On Error Resume Next
      RdoMfr!MFGR_BEMAIL = Trim(txtBEmail)
      RdoMfr.Update
   End If
   
End Sub


Private Sub txtBExt_LostFocus()
   txtBExt = CheckLen(txtBExt, 4)
   txtBExt = Format(Abs(Val(txtBExt)), "###0")
   If bGoodMfr Then
      On Error Resume Next
      RdoMfr!MFGR_BEXT = Val(txtBExt)
      RdoMfr.Update
   End If
   
End Sub


Private Sub txtBPhone_LostFocus()
   txtBPhone = CheckLen(txtBPhone, 20)
   If bGoodMfr Then
      On Error Resume Next
      RdoMfr!MFGR_BPHONE = Trim(txtBPhone)
      RdoMfr.Update
   End If
   
End Sub


Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 255)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   If bGoodMfr Then
      On Error Resume Next
      RdoMfr!MFGR_COMT = Trim(txtCmt)
      RdoMfr.Update
   End If
   
End Sub


Private Sub txtDday_LostFocus()
   txtDday = Format(Abs(Val(txtDday)), "###0")
   If bGoodMfr Then
      On Error Resume Next
      RdoMfr!MFGR_DDAYS = Val(txtDday)
      RdoMfr.Update
   End If
   
End Sub


Private Sub txtDisc_LostFocus()
   txtDisc = Format(Abs(Val(txtDisc)), "###0.00")
   If bGoodMfr Then
      On Error Resume Next
      RdoMfr!MFGR_DISCOUNT = Val(txtDisc)
      RdoMfr.Update
   End If
   
End Sub


Private Sub txtFax_LostFocus()
   txtFax = CheckLen(txtFax, 20)
   If bGoodMfr Then
      On Error Resume Next
      RdoMfr!MFGR_FAX = Trim(txtFax)
      RdoMfr.Update
   End If
   
End Sub






Private Sub txtManu_LostFocus()
   txtManu = CheckLen(txtManu, 30)
   txtManu = StrCase(txtManu)
   If bGoodMfr Then
      On Error Resume Next
      RdoMfr!MFGR_NAME = Trim(txtManu)
      RdoMfr.Update
   End If
   
End Sub


Private Sub txtNDays_LostFocus()
   txtNDays = Format(Abs(Val(txtNDays)), "###0")
   If bGoodMfr Then
      On Error Resume Next
      RdoMfr!MFGR_NETDAYS = Val(txtNDays)
      RdoMfr.Update
   End If
   
End Sub


Private Sub txtPDate_LostFocus()
   txtPDate = Format(Abs(Val(txtPDate)), "###0")
   If bGoodMfr Then
      On Error Resume Next
      RdoMfr!MFGR_PROXDATE = Val(txtPDate)
      RdoMfr.Update
   End If
   
End Sub


Private Sub txtPdue_LostFocus()
   txtPDue = Format(Abs(Val(txtPDue)), "###0")
   If bGoodMfr Then
      On Error Resume Next
      RdoMfr!MFGR_PROXDUE = Val(txtPDue)
      RdoMfr!MFGR_REVISED = Format(Now, "mm/dd/yy")
      RdoMfr.Update
   End If
   
End Sub


Private Sub txtSCont_LostFocus()
   txtSCont = CheckLen(txtSCont, 20)
   txtSCont = StrCase(txtSCont)
   If bGoodMfr Then
      On Error Resume Next
      RdoMfr!MFGR_SCONTACT = Trim(txtSCont)
      RdoMfr.Update
   End If
   
End Sub


Private Sub txtSEmail_DblClick()
   If Trim(txtSEmail) <> "" Then SendEMail Trim(txtSEmail)
   
End Sub


Private Sub txtSEmail_LostFocus()
   txtSEmail = CheckLen(txtSEmail, 60)
   If bGoodMfr Then
      On Error Resume Next
      RdoMfr!MFGR_SEMAIL = Trim(txtSEmail)
      RdoMfr.Update
   End If
   
End Sub


Private Sub txtSext_LostFocus()
   txtSExt = CheckLen(txtSExt, 4)
   txtSExt = Format(Abs(Val(txtSExt)), "###0")
   If bGoodMfr Then
      On Error Resume Next
      RdoMfr!MFGR_SEXT = Trim(txtSExt)
      RdoMfr.Update
   End If
   
End Sub


Private Sub txtSPhone_LostFocus()
   txtSPhone = CheckLen(txtSPhone, 20)
   If bGoodMfr Then
      On Error Resume Next
      RdoMfr!MFGR_SPHONE = Trim(txtSPhone)
      RdoMfr.Update
   End If
   
End Sub



Private Function GetManufacturer() As Byte
   sSql = "SELECT * FROM MfgrTable WHERE MFGR_REF='" & Compress(cmbMfr) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoMfr, ES_KEYSET)
   If bSqlRows Then
      With RdoMfr
         cmbMfr = "" & Trim(!MFGR_NICKNAME)
         lblNum = Format(!MFGR_NUMBER, "##0")
         txtType = "" & Trim(!MFGR_TYPE)
         txtManu = "" & Trim(!MFGR_NAME)
         txtAdr = "" & Trim(!MFGR_ADDRESS)
         txtBCont = "" & Trim(!MFGR_BCONTACT)
         txtBPhone = "" & Trim(!MFGR_BPHONE)
         txtBEmail = "" & Trim(!MFGR_BEMAIL)
         txtBExt = Format(!MFGR_BEXT, "###0")
         txtFax = "" & Trim(!MFGR_FAX)
         txtSCont = "" & Trim(!MFGR_SCONTACT)
         txtSExt = Format(!MFGR_SEXT, "###0")
         txtSEmail = "" & Trim(!MFGR_SEMAIL)
         txtCmt = "" & Trim(!MFGR_COMT)
         txtDisc = Format(!MFGR_DISCOUNT, "###0.00")
         txtNDays = Format(!MFGR_NETDAYS, "###0")
         txtDday = Format(!MFGR_DDAYS, "###0")
         txtPDate = Format(!MFGR_PROXDATE, "###0")
         txtPDue = Format(!MFGR_PROXDUE, "###0")
         tab1.Enabled = True
         GetManufacturer = 1
      End With
   Else
      tab1.Enabled = False
      GetManufacturer = 0
      lblNum = ""
      txtType = ""
      txtManu = ""
      txtAdr = ""
      txtBCont = ""
      txtBPhone = ""
      txtBExt = ""
      txtFax = ""
      txtSCont = ""
      txtSExt = ""
      txtCmt = ""
      txtDisc = ""
      txtNDays = ""
      txtDday = ""
      txtPDate = ""
      txtPDue = ""
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getmanufac"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub txtType_LostFocus()
   txtType = CheckLen(txtType, 2)
   If bGoodMfr Then
      On Error Resume Next
      RdoMfr!MFGR_TYPE = Trim(txtType)
      RdoMfr.Update
   End If
   
End Sub



Private Function GetNextMFR() As Long
   Dim RdoNext As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT MAX(MFGR_NUMBER) AS HighestOne FROM MfgrTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoNext, ES_FORWARD)
   If bSqlRows Then
      With RdoNext
         If Not IsNull(!HighestOne) Then
            GetNextMFR = !HighestOne + 1
         Else
            GetNextMFR = 1
         End If
      End With
      ClearResultSet RdoNext
   Else
      GetNextMFR = 1
   End If
   Set RdoNext = Nothing
   Exit Function
DiaErr1:
   GetNextMFR = 1
   
End Function

Private Sub AddManufacturer()
   Dim bResponse As Byte
   Dim lNewNumber As Long
   Dim sMsg As String
   
   If Len(Trim(cmbMfr)) < 4 Then
      Beep
      MsgBox "(4) Characters Or More Please.", vbInformation, Caption
      Exit Sub
   End If
   
   If cmbMfr = "ALL" Then
      Beep
      MsgBox "ALL Is An Illegal Nickname.", vbExclamation, Caption
      Exit Sub
   End If
   
   If cmbMfr = "NONE" Then
      Beep
      MsgBox "NONE Is An Illegal Nickname.", vbExclamation, Caption
      Exit Sub
   End If
   
   bResponse = IllegalCharacters(cmbMfr)
   If bResponse > 0 Then
      MsgBox "The Nickname Contains An Illegal " & Chr$(bResponse) & ".", _
         vbExclamation, Caption
      Exit Sub
   End If
   
   sMsg = cmbMfr & " Wasn't Found. Add The Manufacturer?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      lNewNumber = GetNextMFR()
      On Error Resume Next
      clsADOCon.ADOErrNum = 0
      sSql = "INSERT INTO MfgrTable (MFGR_REF,MFGR_NICKNAME,MFGR_NUMBER,MFGR_USER) " _
             & "VALUES('" & Compress(cmbMfr) & "','" _
             & Trim(cmbMfr) & "'," _
             & str$(lNewNumber) & ",'" _
             & sInitials & "')"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then
         SysMsg "Added Manufacturer.", True
         cmbMfr.AddItem cmbMfr
         bGoodMfr = GetManufacturer()
      Else
         MsgBox "Could Not Successfully Add " & cmbMfr & ".", _
            vbInformation, Caption
      End If
   Else
      CancelTrans
   End If
   
End Sub
