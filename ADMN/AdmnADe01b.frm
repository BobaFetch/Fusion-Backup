VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form AdmnADe01b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inventory/Expense and Cost of Goods"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   Begin VB.Frame tabFrame 
      Height          =   4452
      Index           =   1
      Left            =   20
      TabIndex        =   38
      Top             =   1320
      Width           =   6732
      Begin VB.ComboBox cmbWipExp 
         Height          =   288
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   12
         Tag             =   "3"
         Top             =   1560
         Width           =   1935
      End
      Begin VB.ComboBox cmbWipOhd 
         Height          =   288
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   11
         Tag             =   "3"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.ComboBox cmbWipMat 
         Height          =   288
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   10
         Tag             =   "3"
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox cmbWipLab 
         Height          =   288
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   9
         Tag             =   "3"
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblWipExp 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   3960
         TabIndex        =   47
         Top             =   1560
         Width           =   2604
      End
      Begin VB.Label lblWipOhd 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   3960
         TabIndex        =   46
         Top             =   1200
         Width           =   2604
      End
      Begin VB.Label lblWipMat 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   3960
         TabIndex        =   45
         Top             =   840
         Width           =   2604
      End
      Begin VB.Label lblWipLab 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   3960
         TabIndex        =   44
         Top             =   480
         Width           =   2604
      End
      Begin VB.Label I 
         BackStyle       =   0  'Transparent
         Caption         =   "WIP Accounts:"
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   43
         Top             =   120
         Width           =   1992
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Labor Account"
         Height          =   252
         Index           =   8
         Left            =   240
         TabIndex        =   42
         Top             =   480
         Width           =   1992
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Material Account"
         Height          =   252
         Index           =   7
         Left            =   240
         TabIndex        =   41
         Top             =   840
         Width           =   1992
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Overhead Account"
         Height          =   252
         Index           =   6
         Left            =   240
         TabIndex        =   40
         Top             =   1200
         Width           =   1992
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Expense Account"
         Height          =   252
         Index           =   5
         Left            =   240
         TabIndex        =   39
         Top             =   1560
         Width           =   1992
      End
   End
   Begin VB.Frame tabFrame 
      Height          =   4452
      Index           =   0
      Left            =   8000
      TabIndex        =   19
      Top             =   1320
      Width           =   6732
      Begin VB.ComboBox txtGla 
         Height          =   288
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   5
         Tag             =   "3"
         Top             =   2400
         Width           =   1935
      End
      Begin VB.ComboBox txtGma 
         Height          =   288
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   6
         Tag             =   "3"
         Top             =   2760
         Width           =   1935
      End
      Begin VB.ComboBox txtGoh 
         Height          =   288
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   7
         Tag             =   "3"
         Top             =   3120
         Width           =   1935
      End
      Begin VB.ComboBox txtGex 
         Height          =   288
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   8
         Tag             =   "3"
         Top             =   3480
         Width           =   1935
      End
      Begin VB.ComboBox txtWex 
         Height          =   288
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   4
         Tag             =   "3"
         Top             =   1560
         Width           =   1935
      End
      Begin VB.ComboBox txtWoh 
         Height          =   288
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   3
         Tag             =   "3"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.ComboBox txtWma 
         Height          =   288
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   2
         Tag             =   "3"
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox txtWla 
         Height          =   288
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   1
         Tag             =   "3"
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "More >>>"
         Height          =   255
         Left            =   5280
         TabIndex        =   48
         Top             =   200
         Width           =   1092
      End
      Begin VB.Label I 
         BackStyle       =   0  'Transparent
         Caption         =   "Cost Of Goods Sold:"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   37
         Top             =   2040
         Width           =   1992
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Labor Account"
         Height          =   372
         Index           =   1
         Left            =   240
         TabIndex        =   36
         Top             =   2400
         Width           =   1992
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Material Account"
         Height          =   372
         Index           =   2
         Left            =   240
         TabIndex        =   35
         Top             =   2760
         Width           =   1992
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Overhead Account"
         Height          =   372
         Index           =   3
         Left            =   240
         TabIndex        =   34
         Top             =   3120
         Width           =   1992
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Expense Account"
         Height          =   372
         Index           =   4
         Left            =   240
         TabIndex        =   33
         Top             =   3480
         Width           =   1992
      End
      Begin VB.Label lblGla 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   3960
         TabIndex        =   32
         Top             =   2400
         Width           =   2604
      End
      Begin VB.Label lblGma 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   3960
         TabIndex        =   31
         Top             =   2760
         Width           =   2604
      End
      Begin VB.Label lblGoh 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   3960
         TabIndex        =   30
         Top             =   3120
         Width           =   2604
      End
      Begin VB.Label lblGex 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   3960
         TabIndex        =   29
         Top             =   3480
         Width           =   2604
      End
      Begin VB.Label I 
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory/Expense:"
         Height          =   252
         Index           =   8
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   1992
      End
      Begin VB.Label lblWex 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   3960
         TabIndex        =   27
         Top             =   1560
         Width           =   2604
      End
      Begin VB.Label lblWoh 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   3960
         TabIndex        =   26
         Top             =   1200
         Width           =   2604
      End
      Begin VB.Label lblWma 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   3960
         TabIndex        =   25
         Top             =   840
         Width           =   2604
      End
      Begin VB.Label lblWla 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   3960
         TabIndex        =   24
         Top             =   480
         Width           =   2604
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Expense Account"
         Height          =   372
         Index           =   12
         Left            =   240
         TabIndex        =   23
         Top             =   1560
         Width           =   1992
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Overhead Account"
         Height          =   372
         Index           =   11
         Left            =   240
         TabIndex        =   22
         Top             =   1200
         Width           =   1992
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Material Account"
         Height          =   372
         Index           =   10
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   1992
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Labor Account"
         Height          =   372
         Index           =   9
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   1992
      End
   End
   Begin MSComctlLib.TabStrip tab1 
      Height          =   4932
      Left            =   40
      TabIndex        =   18
      Top             =   960
      Width           =   6972
      _ExtentX        =   12303
      _ExtentY        =   8705
      TabWidthStyle   =   2
      TabFixedWidth   =   1411
      TabFixedHeight  =   473
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Inventory"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Wip"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "AdmnADe01b.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "P&arts"
      Height          =   315
      Left            =   6120
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Update Part WIP Accounts For This Part Type"
      Top             =   600
      Width           =   855
   End
   Begin VB.ComboBox cmbLvl 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Select Level From List"
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   5880
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5955
      FormDesignWidth =   7065
   End
   Begin VB.Label lblLvl 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2760
      TabIndex        =   15
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   600
      Width           =   1995
   End
End
Attribute VB_Name = "AdmnADe01b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'1/22/05 Added UpdateCurrent
'11/17/05 Added Wip Accounts - changes throughout
'8/10/06 Replaced Tab with TabStrip
Dim bOnLoad As Byte
Dim bDataChanged As Byte

Dim sLevel(9) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbWipExp_Change()
   bDataChanged = 1
   
End Sub

Private Sub cmbWipExp_Click()
   GetAccount cmbWipExp, "lblWipExp"
   
End Sub


Private Sub cmbWipLab_Change()
   bDataChanged = 1
   
End Sub

Private Sub cmbWipLab_Click()
   GetAccount cmbWipLab, "lblWipLab"
   
End Sub


Private Sub cmbWipMat_Change()
   bDataChanged = 1
   
End Sub


Private Sub cmbWipMat_Click()
   GetAccount cmbWipMat, "lblWipMat"
   
End Sub


Private Sub cmbWipOhd_Change()
   bDataChanged = 1
   
End Sub


Private Sub cmbWipOhd_Click()
   GetAccount cmbWipOhd, "lblWipOhd"
   
End Sub



Private Sub cmbLvl_Click()
   If cmbLvl.ListIndex >= 0 Then
      lblLvl = sLevel(cmbLvl.ListIndex)
      FillAccounts cmbLvl.ListIndex
   End If
   CheckPartType
   
End Sub


Private Sub cmdCan_Click()
   UpdateWipAccounts
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 923
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdUpd_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sILabAcct As String
   Dim sIMatAcct As String
   Dim sIOhdAcct As String
   Dim sIExpAcct As String
   Dim sGLabAcct As String
   Dim sGMatAcct As String
   Dim sGOhdAcct As String
   Dim sGExpAcct As String
   
   sMsg = "Do You Wish To Continue To Update  " & vbCr _
          & "All Part Types " & cmbLvl & "?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      sILabAcct = Compress(txtWla)
      sIMatAcct = Compress(txtWma)
      sIOhdAcct = Compress(txtWoh)
      sIExpAcct = Compress(txtWex)
      
      sGLabAcct = Compress(txtGla)
      sGMatAcct = Compress(txtGma)
      sGOhdAcct = Compress(txtGoh)
      sGExpAcct = Compress(txtGex)
      clsADOCon.BeginTrans
      sSql = "UPDATE PartTable SET " _
             & "PAINVLABACCT='" & sILabAcct _
             & "',PAINVMATACCT='" & sIMatAcct _
             & "',PAINVOHDACCT='" & sIOhdAcct _
             & "',PAINVEXPACCT='" & sIExpAcct _
             & "',PACGSLABACCT='" & sGLabAcct _
             & "',PACGSMATACCT='" & sGMatAcct _
             & "',PACGSOHDACCT='" & sGOhdAcct _
             & "',PACGSEXPACCT='" & sGExpAcct _
             & "' WHERE PALEVEL=" & cmbLvl & " "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.RowsAffected > 0 Then
         sMsg = Trim(str(clsADOCon.RowsAffected)) & " Parts Are Selected To Be Updated." & vbCr _
                & "You Wish To Continue Updating?"
         bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
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
         MsgBox "No Parts Selected.", vbInformation, Caption
      End If
   Else
      CancelTrans
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      FillAccounts 0
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_DblClick()
   Unload Me
   
End Sub


Private Sub Form_Load()
   Dim iLevel As Integer
   Move 400, 600
   FormatControls
   
   tabFrame(0).BorderStyle = 0
   tabFrame(0).Left = 60
   tabFrame(1).BorderStyle = 0
   tabFrame(1).Left = 60
   tabFrame(1).Visible = False
   
   sLevel(0) = "Top Assembly"
   sLevel(1) = "Intermediate Assembly"
   sLevel(2) = "Base Assembly"
   sLevel(3) = "Raw Material"
   sLevel(4) = "Expendable"
   sLevel(5) = "Service"
   sLevel(6) = "Outside Service"
   sLevel(7) = "Project"
   
   cmbLvl = " 1"
   For iLevel = 1 To 7
      AddComboStr cmbLvl.hwnd, Format$(iLevel)
   Next
   AddComboStr cmbLvl.hwnd, Format$(iLevel)
   lblLvl = sLevel(0)
   bOnLoad = 1
   Show
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   AdmnADe01a.optWip.Value = vbUnchecked
   If bDataChanged Then UpDateCurrent
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   WindowState = 1
   On Error Resume Next
   Set AdmnADe01b = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   Dim RdoGlm As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "Qry_FillLowAccounts"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGlm, ES_FORWARD)
   If bSqlRows Then
      With RdoGlm
         Do Until .EOF
            'Inv/Exp
            AddComboStr txtWla.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtWma.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtWoh.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtWex.hwnd, "" & Trim(!GLACCTNO)
            'Cgs
            AddComboStr txtGla.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtGma.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtGoh.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtGex.hwnd, "" & Trim(!GLACCTNO)
            
            'Wip
            AddComboStr cmbWipLab.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr cmbWipMat.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr cmbWipExp.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr cmbWipOhd.hwnd, "" & Trim(!GLACCTNO)
            
            .MoveNext
         Loop
         ClearResultSet RdoGlm
      End With
   End If
   Set RdoGlm = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillAccounts(iIndex As Integer)
   Dim RdoFil As ADODB.Recordset
   Dim slvl As String
   On Error GoTo DiaErr1
   slvl = Trim$(str$(iIndex + 1))
   sSql = "SELECT COINVLABACCT" & slvl _
          & ",COINVMATACCT" & slvl _
          & ",COINVOHDACCT" & slvl _
          & ",COINVEXPACCT" & slvl _
          & ",COCGSLABACCT" & slvl _
          & ",COCGSMATACCT" & slvl _
          & ",COCGSOHDACCT" & slvl _
          & ",COCGSEXPACCT" & slvl _
          & ",WIPLABACCT" _
          & ",WIPMATACCT" _
          & ",WIPEXPACCT" _
          & ",WIPOHDACCT " _
          & "FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoFil, ES_KEYSET)
   If bSqlRows Then
      With RdoFil
         txtWla = "" & Trim(.Fields(0))
         txtWma = "" & Trim(.Fields(1))
         txtWoh = "" & Trim(.Fields(2))
         txtWex = "" & Trim(.Fields(3))
         
         txtGla = "" & Trim(.Fields(4))
         txtGma = "" & Trim(.Fields(5))
         txtGoh = "" & Trim(.Fields(6))
         txtGex = "" & Trim(.Fields(7))
         
         If bOnLoad = 1 Then
            cmbWipLab = "" & Trim(.Fields(8))
            cmbWipMat = "" & Trim(.Fields(9))
            cmbWipExp = "" & Trim(.Fields(10))
            cmbWipOhd = "" & Trim(.Fields(11))
         End If
         ClearResultSet RdoFil
      End With
      GetAccount txtWla, "txtWla"
      GetAccount txtWma, "txtWma"
      GetAccount txtWoh, "txtWoh"
      GetAccount txtWex, "txtWex"
      
      GetAccount txtGla, "txtGla"
      GetAccount txtGma, "txtGma"
      GetAccount txtGoh, "txtGoh"
      GetAccount txtGex, "txtGex"
      
      If bOnLoad = 1 Then
         GetAccount cmbWipLab, "lblWipLab"
         GetAccount cmbWipMat, "lblWipMat"
         GetAccount cmbWipExp, "lblWipExp"
         GetAccount cmbWipOhd, "lblWipOhd"
      End If
   Else
      txtWla = ""
      txtWma = ""
      txtWoh = ""
      txtWex = ""
      lblWla = ""
      lblWma = ""
      lblWoh = ""
      lblWex = ""
      
      txtGla = ""
      txtGma = ""
      txtGoh = ""
      txtGex = ""
      lblGla = ""
      lblGma = ""
      lblGoh = ""
      lblGex = ""
   End If
   Set RdoFil = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillaccounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetAccount(sAccount As String, sBox As String)
   Dim RdoGlm As ADODB.Recordset
   
   On Error GoTo DiaErr1
   If bDataChanged = 1 Then UpDateCurrent
   sSql = "Qry_GetAccount '" & Compress(sAccount) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGlm, ES_FORWARD)
   If bSqlRows Then
      With RdoGlm
         Select Case sBox
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
            Case "lblWipLab"
               cmbWipLab = "" & Trim(!GLACCTNO)
               lblWipLab = "" & Trim(!GLDESCR)
            Case "lblWipMat"
               cmbWipMat = "" & Trim(!GLACCTNO)
               lblWipMat = "" & Trim(!GLDESCR)
            Case "lblWipExp"
               cmbWipExp = "" & Trim(!GLACCTNO)
               lblWipExp = "" & Trim(!GLDESCR)
            Case "lblWipOhd"
               cmbWipOhd = "" & Trim(!GLACCTNO)
               lblWipOhd = "" & Trim(!GLDESCR)
         End Select
         ClearResultSet RdoGlm
      End With
   Else
      Select Case sBox
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
         Case "txtGla"
            txtGla = ""
            lblGla = ""
         Case "txtGma"
            txtGma = ""
            lblGma = ""
         Case "txtGoh"
            txtGoh = ""
            lblGoh = ""
         Case "txtGex"
            txtGex = ""
            lblGex = ""
      End Select
   End If
   bDataChanged = 0
   Set RdoGlm = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getaccount"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub





Private Sub tab1_Click()
   On Error Resume Next
   If Tab1.SelectedItem.Index = 1 Then
      tabFrame(0).Visible = True
      tabFrame(1).Visible = False
      txtWla.SetFocus
      cmdUpd.Enabled = False
      cmbLvl.Enabled = False
   Else
      tabFrame(0).Visible = False
      tabFrame(1).Visible = True
      cmbWipLab.SetFocus
      cmdUpd.Enabled = True
      cmbLvl.Enabled = True
   End If
   
End Sub


Private Sub txtGex_Change()
   bDataChanged = 1
   
End Sub

Private Sub txtGex_Click()
   GetAccount txtGex, "txtGex"
   
End Sub

Private Sub txtGex_LostFocus()
   Dim sAccount As String
   txtGex = CheckLen(txtGex, 12)
   GetAccount txtGex, "txtGex"
   sAccount = Compress(txtGex)
   sSql = "UPDATE ComnTable SET COCGSEXPACCT" & Trim(cmbLvl) _
          & "='" & sAccount & "' "
   On Error Resume Next
   clsADOCon.ExecuteSQL sSql
   
End Sub


Private Sub txtGla_Change()
   bDataChanged = 1
   
End Sub

Private Sub txtGla_Click()
   GetAccount txtGla, "txtGla"
   
End Sub


Private Sub txtGla_LostFocus()
   Dim sAccount As String
   txtGla = CheckLen(txtGla, 12)
   GetAccount txtGla, "txtGla"
   sAccount = Compress(txtGla)
   sSql = "UPDATE ComnTable SET COCGSLABACCT" & Trim(cmbLvl) _
          & "='" & sAccount & "' "
   On Error Resume Next
   clsADOCon.ExecuteSQL sSql
   
End Sub


Private Sub txtGma_Change()
   bDataChanged = 1
   
End Sub

Private Sub txtGma_Click()
   GetAccount txtGma, "txtGma"
   
End Sub


Private Sub txtGma_LostFocus()
   Dim sAccount As String
   txtGma = CheckLen(txtGma, 12)
   GetAccount txtGma, "txtGma"
   sAccount = Compress(txtGma)
   sSql = "UPDATE ComnTable SET COCGSMATACCT" & Trim(cmbLvl) _
          & "='" & sAccount & "' "
   On Error Resume Next
   clsADOCon.ExecuteSQL sSql
   
End Sub


Private Sub txtGoh_Change()
   bDataChanged = 1
   
End Sub

Private Sub txtGoh_Click()
   GetAccount txtGoh, "txtGoh"
   
End Sub


Private Sub txtGoh_LostFocus()
   Dim sAccount As String
   txtGoh = CheckLen(txtGoh, 12)
   GetAccount txtGoh, "txtGoh"
   sAccount = Compress(txtGoh)
   sSql = "UPDATE ComnTable SET COCGSOHDACCT" & Trim(cmbLvl) _
          & "='" & sAccount & "' "
   On Error Resume Next
   clsADOCon.ExecuteSQL sSql
   
End Sub


Private Sub txtWex_Change()
   bDataChanged = 1
   
End Sub

Private Sub txtWex_Click()
   GetAccount txtWex, "txtWex"
   
End Sub


Private Sub txtWex_LostFocus()
   Dim sAccount As String
   txtWex = CheckLen(txtWex, 12)
   GetAccount txtWex, "txtWex"
   sAccount = Compress(txtWex)
   sSql = "UPDATE ComnTable SET COINVEXPACCT" & Trim(cmbLvl) _
          & "='" & sAccount & "' "
   On Error Resume Next
   clsADOCon.ExecuteSQL sSql
   
End Sub


Private Sub txtWla_Change()
   bDataChanged = 1
   
End Sub

Private Sub txtWla_Click()
   GetAccount txtWla, "txtWla"
   
End Sub

Private Sub txtWla_LostFocus()
   Dim sAccount As String
   txtWla = CheckLen(txtWla, 12)
   GetAccount txtWla, "txtWla"
   sAccount = Compress(txtWla)
   sSql = "UPDATE ComnTable SET COINVLABACCT" & Trim(cmbLvl) _
          & "='" & sAccount & "' "
   On Error Resume Next
   clsADOCon.ExecuteSQL sSql
   
End Sub


Private Sub txtWma_Change()
   bDataChanged = 1
   
End Sub

Private Sub txtWma_Click()
   GetAccount txtWma, "txtWma"
   
End Sub


Private Sub txtWma_LostFocus()
   Dim sAccount As String
   txtWma = CheckLen(txtWma, 12)
   GetAccount txtWma, "txtWma"
   sAccount = Compress(txtWma)
   sSql = "UPDATE ComnTable SET COINVMATACCT" & Trim(cmbLvl) _
          & "='" & sAccount & "' "
   On Error Resume Next
   clsADOCon.ExecuteSQL sSql
   
End Sub


Private Sub txtWoh_Change()
   bDataChanged = 1
   
End Sub

Private Sub txtWoh_Click()
   GetAccount txtWoh, "txtWoh"
   
End Sub


Private Sub txtWoh_LostFocus()
   Dim sAccount As String
   txtWoh = CheckLen(txtWoh, 12)
   GetAccount txtWoh, "txtWoh"
   sAccount = Compress(txtWoh)
   sSql = "UPDATE ComnTable SET COINVOHDACCT" & Trim(cmbLvl) _
          & "='" & sAccount & "' "
   On Error Resume Next
   clsADOCon.ExecuteSQL sSql
   
End Sub



Private Sub CheckPartType()
   On Error Resume Next
   Select Case Val(cmbLvl)
      Case 4
         txtWma.Enabled = True
         txtGma.Enabled = True
         
         txtWla.Enabled = False
         txtWex.Enabled = False
         txtWoh.Enabled = False
         
         txtGla.Enabled = False
         txtGex.Enabled = False
         txtGoh.Enabled = False
      Case 5, 6, 7
         txtWex.Enabled = True
         txtGex.Enabled = True
         
         txtWla.Enabled = False
         txtWma.Enabled = False
         txtWoh.Enabled = False
         
         txtGla.Enabled = False
         txtGma.Enabled = False
         txtGoh.Enabled = False
      Case Else
         txtWla.Enabled = True
         txtWma.Enabled = True
         txtWoh.Enabled = True
         txtWex.Enabled = True
         
         txtGla.Enabled = True
         txtGma.Enabled = True
         txtGoh.Enabled = True
         txtGex.Enabled = True
   End Select
   
End Sub

Private Sub UpDateCurrent()
   Dim sAccount As String
   On Error Resume Next
   MouseCursor 13
   sAccount = Compress(txtWla)
   sSql = "UPDATE ComnTable SET COINVLABACCT" & Trim(cmbLvl) _
          & "='" & sAccount & "' "
   clsADOCon.ExecuteSQL sSql
   
   sAccount = Compress(txtWma)
   sSql = "UPDATE ComnTable SET COINVMATACCT" & Trim(cmbLvl) _
          & "='" & sAccount & "' "
   clsADOCon.ExecuteSQL sSql
   
   sAccount = Compress(txtWoh)
   sSql = "UPDATE ComnTable SET COINVOHDACCT" & Trim(cmbLvl) _
          & "='" & sAccount & "' "
   clsADOCon.ExecuteSQL sSql
   
   sAccount = Compress(txtWex)
   sSql = "UPDATE ComnTable SET COINVEXPACCT" & Trim(cmbLvl) _
          & "='" & sAccount & "' "
   clsADOCon.ExecuteSQL sSql
   
   sAccount = Compress(txtGla)
   sSql = "UPDATE ComnTable SET COCGSLABACCT" & Trim(cmbLvl) _
          & "='" & sAccount & "' "
   clsADOCon.ExecuteSQL sSql
   
   sAccount = Compress(txtGma)
   sSql = "UPDATE ComnTable SET COCGSMATACCT" & Trim(cmbLvl) _
          & "='" & sAccount & "' "
   clsADOCon.ExecuteSQL sSql
   
   sAccount = Compress(txtGoh)
   sSql = "UPDATE ComnTable SET COCGSOHDACCT" & Trim(cmbLvl) _
          & "='" & sAccount & "' "
   clsADOCon.ExecuteSQL sSql
   
   sAccount = Compress(txtGex)
   sSql = "UPDATE ComnTable SET COCGSEXPACCT" & Trim(cmbLvl) _
          & "='" & sAccount & "' "
   clsADOCon.ExecuteSQL sSql
   
   bDataChanged = 0
   MouseCursor 0
   
End Sub

Private Sub UpdateWipAccounts()
   MouseCursor 13
   On Error Resume Next
   sSql = "UPDATE ComnTable SET WIPLABACCT='" & Compress(cmbWipLab) & "', " _
          & "WIPMATACCT='" & Compress(cmbWipMat) & "', " _
          & "WIPEXPACCT='" & Compress(cmbWipExp) & "', " _
          & "WIPOHDACCT='" & Compress(cmbWipOhd) & "' "
   clsADOCon.ExecuteSQL sSql
   
End Sub
