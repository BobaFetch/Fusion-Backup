VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form diaPcode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Product Codes"
   ClientHeight    =   4875
   ClientLeft      =   2055
   ClientTop       =   330
   ClientWidth     =   7290
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdUpd 
      Caption         =   "P&arts"
      Height          =   315
      Left            =   6240
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Update Part Accounts"
      Top             =   480
      Width           =   875
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Enter/Revise Product Code (6 Char)"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Tag             =   "2"
      Top             =   960
      Width           =   3075
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6240
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   17
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaPcode.frx":0000
      PictureDn       =   "diaPcode.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   4800
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4875
      FormDesignWidth =   7290
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   3375
      Left            =   0
      TabIndex        =   19
      Top             =   1440
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   5953
      _Version        =   393216
      Style           =   1
      TabHeight       =   476
      Enabled         =   0   'False
      TabCaption(0)   =   "&Accounts        "
      TabPicture(0)   =   "diaPcode.frx":028C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "z1(6)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "z1(7)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblTrv"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblTcg"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "z1(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "z1(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblRev"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblDis"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "z1(16)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblAccounts"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtTrv"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtTcg"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtRev"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtDis"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "&Inventory/Expense"
      TabPicture(1)   =   "diaPcode.frx":02A8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "z1(9)"
      Tab(1).Control(1)=   "z1(10)"
      Tab(1).Control(2)=   "z1(11)"
      Tab(1).Control(3)=   "z1(12)"
      Tab(1).Control(4)=   "lblWla"
      Tab(1).Control(5)=   "lblWma"
      Tab(1).Control(6)=   "lblWoh"
      Tab(1).Control(7)=   "lblWex"
      Tab(1).Control(8)=   "z1(14)"
      Tab(1).Control(9)=   "txtWla"
      Tab(1).Control(10)=   "txtWma"
      Tab(1).Control(11)=   "txtWoh"
      Tab(1).Control(12)=   "txtWex"
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "&Cost Of Goods"
      TabPicture(2)   =   "diaPcode.frx":02C4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "z1(4)"
      Tab(2).Control(1)=   "z1(5)"
      Tab(2).Control(2)=   "z1(8)"
      Tab(2).Control(3)=   "z1(13)"
      Tab(2).Control(4)=   "z1(15)"
      Tab(2).Control(5)=   "lblGla"
      Tab(2).Control(6)=   "lblGma"
      Tab(2).Control(7)=   "lblGoh"
      Tab(2).Control(8)=   "lblGex"
      Tab(2).Control(9)=   "txtGla"
      Tab(2).Control(10)=   "txtGma"
      Tab(2).Control(11)=   "txtGoh"
      Tab(2).Control(12)=   "txtGex"
      Tab(2).ControlCount=   13
      Begin VB.ComboBox txtGex 
         Height          =   315
         Left            =   -72840
         Sorted          =   -1  'True
         TabIndex        =   14
         Tag             =   "3"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.ComboBox txtGoh 
         Height          =   315
         Left            =   -72840
         Sorted          =   -1  'True
         TabIndex        =   13
         Tag             =   "3"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.ComboBox txtGma 
         Height          =   315
         Left            =   -72840
         Sorted          =   -1  'True
         TabIndex        =   12
         Tag             =   "3"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.ComboBox txtGla 
         Height          =   315
         Left            =   -72840
         Sorted          =   -1  'True
         TabIndex        =   11
         Tag             =   "3"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.ComboBox txtDis 
         Height          =   315
         Left            =   2160
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1440
         Width           =   1935
      End
      Begin VB.ComboBox txtRev 
         Height          =   315
         Left            =   2160
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1080
         Width           =   1935
      End
      Begin VB.ComboBox txtTcg 
         Height          =   315
         Left            =   2160
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   2160
         Width           =   1935
      End
      Begin VB.ComboBox txtTrv 
         Height          =   315
         Left            =   2160
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1800
         Width           =   1935
      End
      Begin VB.ComboBox txtWex 
         Height          =   315
         Left            =   -72840
         Sorted          =   -1  'True
         TabIndex        =   10
         Tag             =   "3"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.ComboBox txtWoh 
         Height          =   315
         Left            =   -72840
         Sorted          =   -1  'True
         TabIndex        =   9
         Tag             =   "3"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.ComboBox txtWma 
         Height          =   315
         Left            =   -72840
         Sorted          =   -1  'True
         TabIndex        =   8
         Tag             =   "3"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.ComboBox txtWla 
         Height          =   315
         Left            =   -72840
         Sorted          =   -1  'True
         TabIndex        =   7
         Tag             =   "3"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblAccounts 
         BackStyle       =   0  'Transparent
         Caption         =   "No Accounts Have Been Entered"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   2640
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Label lblGex 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -70800
         TabIndex        =   46
         Top             =   2160
         Width           =   2775
      End
      Begin VB.Label lblGoh 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -70800
         TabIndex        =   45
         Top             =   1800
         Width           =   2775
      End
      Begin VB.Label lblGma 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -70800
         TabIndex        =   44
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label lblGla 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -70800
         TabIndex        =   43
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Miscellaneous Accounts:"
         Height          =   255
         Index           =   16
         Left            =   240
         TabIndex        =   42
         Top             =   720
         Width           =   3075
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cost Of Goods Sold Accounts:"
         Height          =   255
         Index           =   15
         Left            =   -74760
         TabIndex        =   41
         Top             =   720
         Width           =   3075
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory/Expense Accounts:"
         Height          =   255
         Index           =   14
         Left            =   -74760
         TabIndex        =   40
         Top             =   720
         Width           =   3075
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Expense Account"
         Height          =   255
         Index           =   13
         Left            =   -74760
         TabIndex        =   39
         Top             =   2160
         Width           =   1995
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Overhead Account"
         Height          =   255
         Index           =   8
         Left            =   -74760
         TabIndex        =   38
         Top             =   1800
         Width           =   1995
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Material Account"
         Height          =   255
         Index           =   5
         Left            =   -74760
         TabIndex        =   37
         Top             =   1440
         Width           =   1995
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Labor Account"
         Height          =   255
         Index           =   4
         Left            =   -74760
         TabIndex        =   36
         Top             =   1080
         Width           =   1995
      End
      Begin VB.Label lblDis 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4200
         TabIndex        =   35
         Top             =   1470
         Width           =   2775
      End
      Begin VB.Label lblRev 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4200
         TabIndex        =   34
         Top             =   1110
         Width           =   2775
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Account"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   33
         Top             =   1470
         Width           =   1995
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Revenue Account"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   32
         Top             =   1080
         Width           =   1995
      End
      Begin VB.Label lblTcg 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4200
         TabIndex        =   31
         Top             =   2160
         Width           =   2775
      End
      Begin VB.Label lblTrv 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4200
         TabIndex        =   30
         Top             =   1800
         Width           =   2775
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transfer CGS Acct"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   29
         Top             =   2160
         Width           =   1995
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transfer Revenue Acct"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   28
         Top             =   1800
         Width           =   1995
      End
      Begin VB.Label lblWex 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -70800
         TabIndex        =   27
         Top             =   2160
         Width           =   2775
      End
      Begin VB.Label lblWoh 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -70800
         TabIndex        =   26
         Top             =   1800
         Width           =   2775
      End
      Begin VB.Label lblWma 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -70800
         TabIndex        =   25
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label lblWla 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -70800
         TabIndex        =   24
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Expense Account"
         Height          =   375
         Index           =   12
         Left            =   -74760
         TabIndex        =   23
         Top             =   2160
         Width           =   1995
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Overhead Account"
         Height          =   375
         Index           =   11
         Left            =   -74760
         TabIndex        =   22
         Top             =   1800
         Width           =   1995
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Material Account"
         Height          =   375
         Index           =   10
         Left            =   -74760
         TabIndex        =   21
         Top             =   1440
         Width           =   1995
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Labor Account"
         Height          =   255
         Index           =   9
         Left            =   -74760
         TabIndex        =   20
         Top             =   1080
         Width           =   1995
      End
   End
   Begin Threed.SSFrame fra2 
      Height          =   30
      Left            =   0
      TabIndex        =   47
      Top             =   1320
      Width           =   7215
      _Version        =   65536
      _ExtentX        =   12726
      _ExtentY        =   53
      _StockProps     =   14
      ForeColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   15
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "diaPcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Dim RdoCde As ADODB.Recordset
Dim bCancel As Boolean
Dim bGoodCode As Byte
Dim bOnLoad As Byte
Dim bNoAccts As Boolean

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCde_Click()
   bGoodCode = GetCode()
   Tab1.enabled = False
   
End Sub

Private Sub cmbCde_LostFocus()
   cmbCde = CheckLen(cmbCde, 6)
   If bCancel Then Exit Sub
   If Len(cmbCde) Then
      bGoodCode = GetCode(True)
      If Not bGoodCode Then
         AddProductCode
      Else
         Tab1.enabled = True
      End If
   End If
   
End Sub


Private Sub cmdCan_Click()
   bCancel = True
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   bCancel = True
   
End Sub


Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "hs1301"
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdUpd_Click()
   Dim bResponse As Byte
   Dim sPcode As String
   Dim sMsg As String
   Dim sPart As String
   Dim sProdAccount(14) As String
   
   sMsg = "Do You Want To Update Accounts Of " & vbCr _
          & "All Parts With Product Code " & cmbCde & "?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      cmdUpd.enabled = False
      sPcode = Compress(cmbCde)
      sProdAccount(0) = Compress(txtRev)
      sProdAccount(1) = Compress(txtDis)
      sProdAccount(4) = Compress(txtTrv)
      sProdAccount(5) = Compress(txtTcg)
      'INV/EXP
      sProdAccount(6) = Compress(txtWla)
      sProdAccount(7) = Compress(txtWma)
      sProdAccount(8) = Compress(txtWoh)
      sProdAccount(9) = Compress(txtWex)
      'CGS
      sProdAccount(10) = Compress(txtGla)
      sProdAccount(11) = Compress(txtGma)
      sProdAccount(12) = Compress(txtGoh)
      sProdAccount(13) = Compress(txtGex)
      
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      
      sSql = "UPDATE PartTable SET " _
             & "PAREVACCT='" & sProdAccount(0) & "'," _
             & "PADISACCT='" & sProdAccount(1) & "'," _
             & "PACGSACCT='" & sProdAccount(2) & "'," _
             & "PAACCTNO='" & sProdAccount(3) & "'," _
             & "PATFRREVACCT='" & sProdAccount(4) & "'," _
             & "PATFRCGSACCT='" & sProdAccount(5) & "'," _
             & "PAINVLABACCT='" & sProdAccount(6) & "'," _
             & "PAINVMATACCT='" & sProdAccount(7) & "'," _
             & "PAINVOHDACCT='" & sProdAccount(8) & "'," _
             & "PAINVEXPACCT='" & sProdAccount(9) & "'," _
             & "PACGSLABACCT='" & sProdAccount(10) & "'," _
             & "PACGSMATACCT='" & sProdAccount(11) & "'," _
             & "PACGSOHDACCT='" & sProdAccount(12) & "'," _
             & "PACGSEXPACCT='" & sProdAccount(13) & "' " _
             & "WHERE PAPRODCODE='" & sPcode & "' "
      clsADOCon.ExecuteSQL sSql
      MouseCursor 0
      cmdUpd.enabled = True
      If clsADOCon.RowsAffected > 0 Then
         sMsg = Trim(str(clsADOCon.RowsAffected)) & " Parts Are Selected To Be Updated." & vbCr _
                & "You Wish To Continue Updating?"
         bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
         If bResponse = vbYes Then
            clsADOCon.CommitTrans
            MsgBox str(clsADOCon.RowsAffected) & " Parts Were Updated.", _
                       vbInformation, Caption
         Else
            clsADOCon.RollbackTrans
            clsADOCon.ADOErrNum = 0
            CancelTrans
         End If
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         MsgBox "No Parts Were Updated.", _
            vbInformation, Caption
      End If
   Else
      CancelTrans
   End If
   
End Sub





Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillProductCodes Me
      FillAccounts
      If cmbCde.ListCount > 0 Then
         cmbCde = cmbCde.List(0)
         bGoodCode = GetCode()
      End If
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   Tab1.Tab = 0
   bOnLoad = 1
   Show
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set RdoCde = Nothing
   Set diaPcode = Nothing
   
End Sub







Private Sub tab1_Click(PreviousTab As Integer)
   On Error Resume Next
   Select Case Tab1.Tab
      Case 1
         txtWla.SetFocus
      Case 2
         txtGla.SetFocus
      Case Else
         txtRev.SetFocus
   End Select
   
End Sub

Private Sub txtDis_Click()
   GetAccount txtDis, "txtDis"
   
End Sub

Private Sub txtDis_LostFocus()
   txtDis = CheckLen(txtDis, 12)
   GetAccount txtDis, "txtDis"
   If bGoodCode Then
      On Error Resume Next
      RdoCde!PCDISCACCT = "" & Compress(txtDis)
      RdoCde.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 30)
   txtDsc = StrCase(txtDsc)
   If bGoodCode Then
      On Error Resume Next
      RdoCde!PCDESC = "" & txtDsc
      RdoCde.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub





Private Sub txtGex_Click()
   GetAccount txtGex, "txtGex"
   
End Sub


Private Sub txtGex_LostFocus()
   txtGex = CheckLen(txtGex, 12)
   GetAccount txtGex, "txtGex"
   If bGoodCode Then
      On Error Resume Next
      RdoCde!PCCGSEXPACCT = "" & Compress(txtGex)
      RdoCde.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtGla_Click()
   GetAccount txtGla, "txtGla"
   
End Sub


Private Sub txtGla_LostFocus()
   txtGla = CheckLen(txtGla, 12)
   GetAccount txtGla, "txtGla"
   If bGoodCode Then
      On Error Resume Next
      RdoCde!PCCGSLABACCT = "" & Compress(txtGla)
      RdoCde.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtGma_Click()
   GetAccount txtGma, "txtGma"
   
End Sub


Private Sub txtGma_LostFocus()
   txtGma = CheckLen(txtGma, 12)
   GetAccount txtGma, "txtGma"
   If bGoodCode Then
      On Error Resume Next
      RdoCde!PCCGSMATACCT = "" & Compress(txtGma)
      RdoCde.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtGoh_Click()
   GetAccount txtGoh, "txtGoh"
   
End Sub


Private Sub txtGoh_LostFocus()
   txtGoh = CheckLen(txtGoh, 12)
   GetAccount txtGoh, "txtGoh"
   If bGoodCode Then
      On Error Resume Next
      RdoCde!PCCGSOHDACCT = "" & Compress(txtGoh)
      RdoCde.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtRev_Click()
   GetAccount txtRev, "txtRev"
   
End Sub

Private Sub txtRev_LostFocus()
   txtRev = CheckLen(txtRev, 12)
   GetAccount txtRev, "txtRev"
   If bGoodCode Then
      On Error Resume Next
      RdoCde!PCREVACCT = "" & Compress(txtRev)
      RdoCde.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtTcg_Click()
   GetAccount txtTcg, "txtTcg"
   
End Sub

Private Sub txtTcg_LostFocus()
   txtTcg = CheckLen(txtTcg, 12)
   GetAccount txtTcg, "txtTcg"
   If bGoodCode Then
      On Error Resume Next
      RdoCde!PCDCGSXFERAC = "" & Compress(txtTcg)
      RdoCde.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtTrv_Click()
   GetAccount txtTrv, "txtTrv"
   
End Sub

Private Sub txtTrv_LostFocus()
   txtTrv = CheckLen(txtTrv, 12)
   GetAccount txtTrv, "txtTrv"
   If bGoodCode Then
      On Error Resume Next
      RdoCde!PCDREVXFERAC = "" & Compress(txtTrv)
      RdoCde.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   lblAccounts.ForeColor = ES_BLUE
   
End Sub



Private Function GetCode(Optional MoveTab As Byte) As Byte
   Dim sPcode As String
   sPcode = Compress(cmbCde)
   If MoveTab Then Tab1.Tab = 0
   On Error GoTo DiaErr1
   MouseCursor 13
   sSql = "SELECT * FROM PcodTable WHERE PCREF='" & sPcode & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCde, ES_KEYSET)
   If bSqlRows Then
      With RdoCde
         cmbCde = "" & Trim(!PCCODE)
         txtDsc = "" & Trim(!PCDESC)
         
         txtRev = "" & Trim(!PCREVACCT)
         GetAccount txtRev, "txtRev"
         
         txtDis = "" & Trim(!PCDISCACCT)
         GetAccount txtDis, "txtDis"
         
         txtTrv = "" & Trim(!PCDREVXFERAC)
         GetAccount txtTrv, "txtTrv"
         
         txtTcg = "" & Trim(!PCDCGSXFERAC)
         GetAccount txtTcg, "txtTcg"
         'Inv/Exp
         txtWla = "" & Trim(!PCINVLABACCT)
         GetAccount txtWla, "txtWla"
         
         txtWma = "" & Trim(!PCINVMATACCT)
         GetAccount txtWma, "txtWma"
         
         txtWoh = "" & Trim(!PCINVOHDACCT)
         GetAccount txtWoh, "txtWoh"
         
         txtWex = "" & Trim(!PCINVEXPACCT)
         GetAccount txtWex, "txtWex"
         'CGS
         txtGla = "" & Trim(!PCCGSLABACCT)
         GetAccount txtGla, "txtGla"
         
         txtGma = "" & Trim(!PCCGSMATACCT)
         GetAccount txtGma, "txtGma"
         
         txtGoh = "" & Trim(!PCCGSOHDACCT)
         GetAccount txtGoh, "txtGoh"
         
         txtGex = "" & Trim(!PCCGSEXPACCT)
         GetAccount txtGex, "txtGex"
         
      End With
      GetCode = True
   Else
      txtDsc = ""
      txtRev = ""
      txtDis = ""
      txtTrv = ""
      txtTcg = ""
      txtWla = ""
      txtWma = ""
      txtWoh = ""
      txtWex = ""
      txtGla = ""
      txtGma = ""
      txtGoh = ""
      txtGex = ""
      GetCode = False
   End If
   MouseCursor 0
   Exit Function
   
DiaErr1:
   sProcName = "getcode"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AddProductCode()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sCode As String
   
   sCode = Compress(cmbCde)
   If sCode = "ALL" Then
      MsgBox "Illegal Product Code Name.", vbExclamation, Caption
      Exit Sub
   End If
   bResponse = IllegalCharacters(cmbCde)
   If bResponse > 0 Then
      MsgBox "The Product Code Contains An Illegal " & Chr$(bResponse) & ".", _
         vbExclamation, Caption
      Exit Sub
   End If
   sMsg = cmbCde & " Wasn't Found. Add The Product Code?"
   On Error GoTo DiaErr1
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      sSql = "INSERT INTO PcodTable (PCREF,PCCODE) " _
             & "VALUES('" & sCode & "','" & cmbCde & "')"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.RowsAffected Then
         SysMsg "Product Code Added.", True
         AddComboStr cmbCde.hWnd, cmbCde
         bGoodCode = GetCode()
         Tab1.enabled = True
         On Error Resume Next
         txtDsc.SetFocus
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "addproduc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillAccounts()
   Dim RdoGlm As ADODB.Recordset
   On Error GoTo DiaErr1
   MouseCursor 13
   sSql = "Qry_FillLowAccounts"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGlm, ES_FORWARD)
   If bSqlRows Then
      With RdoGlm
         Do Until .EOF
            AddComboStr txtRev.hWnd, "" & Trim(!GLACCTNO)
            AddComboStr txtDis.hWnd, "" & Trim(!GLACCTNO)
            AddComboStr txtTrv.hWnd, "" & Trim(!GLACCTNO)
            AddComboStr txtTcg.hWnd, "" & Trim(!GLACCTNO)
            
            'Inv/Exp
            AddComboStr txtWla.hWnd, "" & Trim(!GLACCTNO)
            AddComboStr txtWma.hWnd, "" & Trim(!GLACCTNO)
            AddComboStr txtWoh.hWnd, "" & Trim(!GLACCTNO)
            AddComboStr txtWex.hWnd, "" & Trim(!GLACCTNO)
            
            'Inv/Exp
            AddComboStr txtGla.hWnd, "" & Trim(!GLACCTNO)
            AddComboStr txtGma.hWnd, "" & Trim(!GLACCTNO)
            AddComboStr txtGoh.hWnd, "" & Trim(!GLACCTNO)
            AddComboStr txtGex.hWnd, "" & Trim(!GLACCTNO)
            .MoveNext
         Loop
         .Cancel
      End With
   Else
      bNoAccts = True
      CloseBoxes
   End If
   Set RdoGlm = Nothing
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   bNoAccts = True
   CloseBoxes
   
End Sub

Private Sub GetAccount(sAccount As String, sBox As String)
   Dim RdoGlm As ADODB.Recordset
   On Error GoTo DiaErr1
   If bNoAccts Then Exit Sub
   MouseCursor 13
   '1/7/04
   ' sSql = "SELECT GLACCTREF,GLACCTNO,GLDESCR FROM GlacTable " _
   '     & "WHERE GLACCTREF='" & sAccount & "' "
   sSql = "Qry_GetAccount '" & Compress(sAccount) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGlm, ES_FORWARD)
   If bSqlRows Then
      With RdoGlm
         Select Case sBox
            Case "txtRev"
               txtRev = "" & Trim(!GLACCTNO)
               lblRev = "" & Trim(!GLDESCR)
            Case "txtDis"
               txtDis = "" & Trim(!GLACCTNO)
               lblDis = "" & Trim(!GLDESCR)
            Case "txtTrv"
               txtTrv = "" & Trim(!GLACCTNO)
               lblTrv = "" & Trim(!GLDESCR)
            Case "txtTcg"
               txtTcg = "" & Trim(!GLACCTNO)
               lblTcg = "" & Trim(!GLDESCR)
               'Inv/Exp
            Case "txtWla"
               txtWla = "" & Trim(!GLACCTNO)
               lblWla = "" & Trim(!GLDESCR)
            Case "txtWma"
               txtWma = "" & Trim(!GLACCTNO)
               lblWma = "" & Trim(!GLDESCR)
            Case "txtWoh"
               txtWoh = "" & Trim(!GLACCTNO)
               lblWoh = "" & Trim(!GLDESCR)
            Case "txtWex"
               txtWex = "" & Trim(!GLACCTNO)
               lblWex = "" & Trim(!GLDESCR)
               'Cogs
            Case "txtGla"
               txtGla = "" & Trim(!GLACCTNO)
               lblGla = "" & Trim(!GLDESCR)
            Case "txtGma"
               txtGma = "" & Trim(!GLACCTNO)
               lblGma = "" & Trim(!GLDESCR)
            Case "txtGoh"
               txtGoh = "" & Trim(!GLACCTNO)
               lblGoh = "" & Trim(!GLDESCR)
            Case "txtGex"
               txtGex = "" & Trim(!GLACCTNO)
               lblGex = "" & Trim(!GLDESCR)
         End Select
         .Cancel
      End With
   Else
      Select Case sBox
         Case "txtRev"
            txtRev = ""
            lblRev = ""
         Case "txtDis"
            txtDis = ""
            lblDis = ""
         Case "txtTrv"
            txtTrv = ""
            lblTrv = ""
         Case "txtTcg"
            txtTcg = ""
            lblTcg = ""
         Case "txtWla"
            txtWla = ""
            lblWla = ""
         Case "txtWma"
            txtWma = ""
            lblWma = ""
         Case "txtWoh"
            txtWoh = ""
            lblWoh = ""
         Case "txtWex"
            txtWex = ""
            lblWex = ""
      End Select
   End If
   Set RdoGlm = Nothing
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "getaccount"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub CloseBoxes()
   txtRev.enabled = False
   txtDis.enabled = False
   txtTrv.enabled = False
   txtTcg.enabled = False
   txtWex.enabled = False
   txtWla.enabled = False
   txtWma.enabled = False
   txtWoh.enabled = False
   txtGex.enabled = False
   txtGla.enabled = False
   txtGma.enabled = False
   txtGoh.enabled = False
   cmdUpd.enabled = False
   lblAccounts.Visible = True
   
End Sub

Private Sub txtWex_Click()
   GetAccount txtWex, "txtWex"
   
End Sub


Private Sub txtWex_LostFocus()
   txtWex = CheckLen(txtWex, 12)
   GetAccount txtWex, "txtWex"
   If bGoodCode Then
      On Error Resume Next
      RdoCde!PCINVEXPACCT = "" & Compress(txtWex)
      RdoCde.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtWla_Click()
   GetAccount txtWla, "txtWla"
   
End Sub


Private Sub txtWla_LostFocus()
   txtWla = CheckLen(txtWla, 12)
   GetAccount txtWla, "txtWla"
   If bGoodCode Then
      On Error Resume Next
      RdoCde!PCINVLABACCT = "" & Compress(txtWla)
      RdoCde.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtWma_Click()
   GetAccount txtWma, "txtWma"
   
End Sub


Private Sub txtWma_LostFocus()
   txtWma = CheckLen(txtWma, 12)
   GetAccount txtWma, "txtWma"
   If bGoodCode Then
      On Error Resume Next
      RdoCde!PCINVMATACCT = "" & Compress(txtWma)
      RdoCde.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtWoh_Click()
   GetAccount txtWoh, "txtWoh"
   
End Sub


Private Sub txtWoh_LostFocus()
   txtWoh = CheckLen(txtWoh, 12)
   GetAccount txtWoh, "txtWoh"
   If bGoodCode Then
      On Error Resume Next
      RdoCde!PCINVOHDACCT = "" & Compress(txtWoh)
      RdoCde.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub
