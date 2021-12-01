VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form PadmPRe01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Product Codes"
   ClientHeight    =   4875
   ClientLeft      =   2055
   ClientTop       =   330
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.Frame tabFrame 
      Height          =   2892
      Index           =   2
      Left            =   7680
      TabIndex        =   42
      Top             =   1800
      Width           =   7092
      Begin VB.ComboBox txtGex 
         Height          =   288
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   13
         Tag             =   "3"
         Top             =   1560
         Width           =   1935
      End
      Begin VB.ComboBox txtGoh 
         Height          =   288
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   12
         Tag             =   "3"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.ComboBox txtGma 
         Height          =   288
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   11
         Tag             =   "3"
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox txtGla 
         Height          =   288
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   10
         Tag             =   "3"
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblGex 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4080
         TabIndex        =   51
         Top             =   1560
         Width           =   2772
      End
      Begin VB.Label lblGoh 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4080
         TabIndex        =   50
         Top             =   1200
         Width           =   2772
      End
      Begin VB.Label lblGma 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4080
         TabIndex        =   49
         Top             =   840
         Width           =   2772
      End
      Begin VB.Label lblGla 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4080
         TabIndex        =   48
         Top             =   480
         Width           =   2772
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cost Of Goods Sold Accounts:"
         Height          =   252
         Index           =   15
         Left            =   120
         TabIndex        =   47
         Top             =   120
         Width           =   3072
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Expense Account"
         Height          =   252
         Index           =   13
         Left            =   120
         TabIndex        =   46
         Top             =   1560
         Width           =   1992
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Overhead Account"
         Height          =   252
         Index           =   8
         Left            =   120
         TabIndex        =   45
         Top             =   1200
         Width           =   1992
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Material Account"
         Height          =   252
         Index           =   5
         Left            =   120
         TabIndex        =   44
         Top             =   840
         Width           =   1992
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Labor Account"
         Height          =   252
         Index           =   4
         Left            =   120
         TabIndex        =   43
         Top             =   480
         Width           =   1992
      End
   End
   Begin VB.Frame tabFrame 
      Height          =   2892
      Index           =   1
      Left            =   7320
      TabIndex        =   32
      Top             =   1800
      Width           =   7092
      Begin VB.ComboBox txtWex 
         Height          =   288
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   9
         Tag             =   "3"
         Top             =   1560
         Width           =   1935
      End
      Begin VB.ComboBox txtWoh 
         Height          =   288
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   8
         Tag             =   "3"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.ComboBox txtWma 
         Height          =   288
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   7
         Tag             =   "3"
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox txtWla 
         Height          =   288
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   6
         Tag             =   "3"
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory/Expense Accounts:"
         Height          =   252
         Index           =   14
         Left            =   120
         TabIndex        =   41
         Top             =   120
         Width           =   3072
      End
      Begin VB.Label lblWex 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4080
         TabIndex        =   40
         Top             =   1560
         Width           =   2772
      End
      Begin VB.Label lblWoh 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4080
         TabIndex        =   39
         Top             =   1200
         Width           =   2772
      End
      Begin VB.Label lblWma 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4080
         TabIndex        =   38
         Top             =   840
         Width           =   2772
      End
      Begin VB.Label lblWla 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4080
         TabIndex        =   37
         Top             =   480
         Width           =   2772
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Expense Account"
         Height          =   372
         Index           =   12
         Left            =   120
         TabIndex        =   36
         Top             =   1560
         Width           =   1992
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Overhead Account"
         Height          =   372
         Index           =   11
         Left            =   120
         TabIndex        =   35
         Top             =   1200
         Width           =   1992
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Material Account"
         Height          =   372
         Index           =   10
         Left            =   120
         TabIndex        =   34
         Top             =   840
         Width           =   1992
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Labor Account"
         Height          =   252
         Index           =   9
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   1992
      End
   End
   Begin VB.Frame tabFrame 
      Height          =   2892
      Index           =   0
      Left            =   40
      TabIndex        =   21
      Top             =   1800
      Width           =   7092
      Begin VB.ComboBox txtTcg 
         Height          =   288
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1560
         Width           =   1935
      End
      Begin VB.ComboBox txtDis 
         Height          =   288
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox txtRev 
         Height          =   288
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
      Begin VB.ComboBox txtTrv 
         Height          =   288
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblAccounts 
         BackStyle       =   0  'Transparent
         Caption         =   "No Accounts Have Been Entered"
         Height          =   252
         Left            =   120
         TabIndex        =   31
         Top             =   2040
         Visible         =   0   'False
         Width           =   3852
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Miscellaneous Accounts:"
         Height          =   252
         Index           =   16
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   3072
      End
      Begin VB.Label lblDis 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4080
         TabIndex        =   29
         Top             =   876
         Width           =   2772
      End
      Begin VB.Label lblRev 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4080
         TabIndex        =   28
         Top             =   516
         Width           =   2772
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Account"
         Height          =   252
         Index           =   3
         Left            =   120
         TabIndex        =   27
         Top             =   876
         Width           =   1992
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Revenue Account"
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   1992
      End
      Begin VB.Label lblTcg 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4080
         TabIndex        =   25
         Top             =   1560
         Width           =   2772
      End
      Begin VB.Label lblTrv 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4080
         TabIndex        =   24
         Top             =   1200
         Width           =   2772
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transfer CGS Acct"
         Height          =   252
         Index           =   7
         Left            =   120
         TabIndex        =   23
         Top             =   1560
         Width           =   1992
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transfer Revenue Acct"
         Height          =   252
         Index           =   6
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   1992
      End
   End
   Begin MSComctlLib.TabStrip tab1 
      Height          =   3372
      Left            =   0
      TabIndex        =   20
      Top             =   1440
      Width           =   7212
      _ExtentX        =   12726
      _ExtentY        =   5953
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      TabFixedWidth   =   2558
      TabFixedHeight  =   473
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Accounts"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Inventory/Expense"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cost Of Goods"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame z2 
      Height          =   30
      Left            =   0
      TabIndex        =   19
      Top             =   1320
      Width           =   7260
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PadmPRe01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "P&arts"
      Height          =   315
      Left            =   6360
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Update Part Accounts"
      Top             =   600
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
      Left            =   6360
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
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
      FormDesignWidth =   7260
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
      TabIndex        =   14
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "PadmPRe01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'8/10/06 Replaced Tab with TabStrip
Option Explicit
Dim RdoCde As ADODB.Recordset
Dim bCancel As Boolean
Dim bGoodCode As Byte
Dim bOnLoad As Byte
Dim bNoAccts As Boolean

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCde_Change()
   If Len(cmbCde) > 6 Then cmbCde = Left(cmbCde, 6)
   
End Sub

Private Sub cmbCde_Click()
   bGoodCode = GetCode()
   
End Sub

Private Sub cmbCde_LostFocus()
   cmbCde = CheckLen(cmbCde, 6)
   If bCancel Then Exit Sub
   If Len(cmbCde) Then
      bGoodCode = GetCode(True)
      If Not bGoodCode Then AddProductCode
   End If
End Sub


Private Sub cmdCan_Click()
   bCancel = True
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = True
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 1301
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
      cmdUpd.Enabled = False
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
      cmdUpd.Enabled = True
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
            CancelTrans
         End If
      Else
         clsADOCon.RollbackTrans
         MsgBox "No Parts Were Updated.", _
            vbInformation, Caption
      End If
   Else
      CancelTrans
   End If
    
End Sub





Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillProductCodes
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
   Dim b As Byte
   FormLoad Me
   FormatControls
   For b = 0 To 2
      With tabFrame(b)
         .Left = 40
         .BorderStyle = 0
         .Visible = False
      End With
   Next
   tabFrame(0).Visible = True
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set RdoCde = Nothing
   Set PadmPRe01a = Nothing
   
End Sub







Private Sub tab1_Click()
   Dim b As Byte
   On Error Resume Next
   For b = 0 To 2
      tabFrame(b).Visible = False
   Next
   tabFrame(tab1.SelectedItem.Index - 1).Visible = True
   Select Case tab1.SelectedItem.Index
      Case 2
         txtWla.SetFocus
      Case 3
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
      'RdoCde.Edit
      RdoCde!PCDISCACCT = "" & Compress(txtDis)
      RdoCde.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 30)
   txtDsc = StrCase(txtDsc)
   If bGoodCode Then
      On Error Resume Next
      'RdoCde.Edit
      RdoCde!PCDESC = "" & txtDsc
      RdoCde.Update
      If Err > 0 Then ValidateEdit
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
      'RdoCde.Edit
      RdoCde!PCCGSEXPACCT = "" & Compress(txtGex)
      RdoCde.Update
      If Err > 0 Then ValidateEdit
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
      'RdoCde.Edit
      RdoCde!PCCGSLABACCT = "" & Compress(txtGla)
      RdoCde.Update
      If Err > 0 Then ValidateEdit
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
      'RdoCde.Edit
      RdoCde!PCCGSMATACCT = "" & Compress(txtGma)
      RdoCde.Update
      If Err > 0 Then ValidateEdit
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
      'RdoCde.Edit
      RdoCde!PCCGSOHDACCT = "" & Compress(txtGoh)
      RdoCde.Update
      If Err > 0 Then ValidateEdit
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
      'RdoCde.Edit
      RdoCde!PCREVACCT = "" & Compress(txtRev)
      RdoCde.Update
      If Err > 0 Then ValidateEdit
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
      'RdoCde.Edit
      RdoCde!PCDCGSXFERAC = "" & Compress(txtTcg)
      RdoCde.Update
      If Err > 0 Then ValidateEdit
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
      'RdoCde.Edit
      RdoCde!PCDREVXFERAC = "" & Compress(txtTrv)
      RdoCde.Update
      If Err > 0 Then ValidateEdit
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
         AddComboStr cmbCde.hwnd, cmbCde
         bGoodCode = GetCode()
         tab1.Enabled = True
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
            AddComboStr txtRev.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtDis.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtTrv.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtTcg.hwnd, "" & Trim(!GLACCTNO)
            
            'Inv/Exp
            AddComboStr txtWla.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtWma.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtWoh.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtWex.hwnd, "" & Trim(!GLACCTNO)
            
            'Inv/Exp
            AddComboStr txtGla.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtGma.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtGoh.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtGex.hwnd, "" & Trim(!GLACCTNO)
            .MoveNext
         Loop
         ClearResultSet RdoGlm
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
         ClearResultSet RdoGlm
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
   txtRev.Enabled = False
   txtDis.Enabled = False
   txtTrv.Enabled = False
   txtTcg.Enabled = False
   txtWex.Enabled = False
   txtWla.Enabled = False
   txtWma.Enabled = False
   txtWoh.Enabled = False
   txtGex.Enabled = False
   txtGla.Enabled = False
   txtGma.Enabled = False
   txtGoh.Enabled = False
   cmdUpd.Enabled = False
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
      'RdoCde.Edit
      RdoCde!PCINVEXPACCT = "" & Compress(txtWex)
      RdoCde.Update
      If Err > 0 Then ValidateEdit
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
      'RdoCde.Edit
      RdoCde!PCINVLABACCT = "" & Compress(txtWla)
      RdoCde.Update
      If Err > 0 Then ValidateEdit
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
      'RdoCde.Edit
      RdoCde!PCINVMATACCT = "" & Compress(txtWma)
      RdoCde.Update
      If Err > 0 Then ValidateEdit
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
      'RdoCde.Edit
      RdoCde!PCINVOHDACCT = "" & Compress(txtWoh)
      RdoCde.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub
