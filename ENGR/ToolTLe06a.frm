VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form ToolTLe06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add/Delete Custom Tools"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   1875
   ClientWidth     =   8895
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   3401
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMO 
      Height          =   285
      Left            =   4560
      TabIndex        =   66
      Tag             =   "2"
      ToolTipText     =   "Tool Description (40 chars)"
      Top             =   8040
      Width           =   2505
   End
   Begin VB.ComboBox cmbLastPhyInv 
      Height          =   315
      Left            =   1920
      TabIndex        =   29
      Tag             =   "4"
      ToolTipText     =   "Don't Use After"
      Top             =   8040
      Width           =   1095
   End
   Begin VB.TextBox txtUnitNo 
      Height          =   285
      Left            =   6960
      TabIndex        =   15
      Tag             =   "2"
      ToolTipText     =   "Tool Description (40 chars)"
      Top             =   3480
      Width           =   345
   End
   Begin VB.TextBox txtCode 
      Height          =   285
      Left            =   5040
      TabIndex        =   14
      Tag             =   "2"
      ToolTipText     =   "Tool Description (40 chars)"
      Top             =   3480
      Width           =   705
   End
   Begin VB.ComboBox cmbCategory 
      Height          =   315
      Left            =   6960
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Tag             =   "3"
      ToolTipText     =   "Select Customer From List"
      Top             =   3000
      Width           =   600
   End
   Begin VB.TextBox txtGovPrimeCont 
      Height          =   285
      Left            =   3840
      TabIndex        =   11
      Tag             =   "2"
      ToolTipText     =   "Tool Description (40 chars)"
      Top             =   3000
      Width           =   2025
   End
   Begin VB.CommandButton cmdDelete 
      Cancel          =   -1  'True
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   435
      Left            =   7200
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   4200
      Width           =   875
   End
   Begin VB.CheckBox ChkITAR 
      Caption         =   "Check1"
      Height          =   255
      Left            =   5280
      TabIndex        =   28
      Top             =   7560
      Width           =   255
   End
   Begin VB.CheckBox chkGovOwned 
      Caption         =   "Check1"
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   3000
      Width           =   255
   End
   Begin VB.ComboBox cmbAcctTo 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   19
      Tag             =   "2"
      ToolTipText     =   "12 Char Class - Retrieved From Previous Entries"
      Top             =   5655
      Width           =   1935
   End
   Begin VB.ComboBox cmbCst 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   17
      Tag             =   "2"
      ToolTipText     =   "12 Char Class - Retrieved From Previous Entries"
      Top             =   5280
      Width           =   1935
   End
   Begin VB.ComboBox cmbDisp 
      Height          =   315
      Left            =   5280
      TabIndex        =   26
      Tag             =   "3"
      ToolTipText     =   "Select Customer From List"
      Top             =   7200
      Width           =   1555
   End
   Begin VB.ComboBox cmbServ 
      Height          =   315
      Left            =   1920
      TabIndex        =   27
      Tag             =   "3"
      ToolTipText     =   "Select Customer From List"
      Top             =   7560
      Width           =   1555
   End
   Begin VB.TextBox txtWeight 
      Height          =   285
      Left            =   6480
      TabIndex        =   24
      Tag             =   "2"
      ToolTipText     =   "Standard Or Actual Cost"
      Top             =   6480
      Width           =   1935
   End
   Begin VB.TextBox txtPN 
      Height          =   285
      Left            =   1920
      TabIndex        =   23
      Tag             =   "2"
      ToolTipText     =   "Standard Or Actual Cost"
      Top             =   6480
      Width           =   2415
   End
   Begin VB.TextBox txtBlankedPO 
      Height          =   285
      Left            =   6480
      TabIndex        =   22
      Tag             =   "2"
      ToolTipText     =   "Standard Or Actual Cost"
      Top             =   6120
      Width           =   1935
   End
   Begin VB.TextBox txtSlNo 
      Height          =   285
      Left            =   1920
      TabIndex        =   21
      Tag             =   "2"
      ToolTipText     =   "Standard Or Actual Cost"
      Top             =   6120
      Width           =   2415
   End
   Begin VB.TextBox txtDim 
      Height          =   285
      Left            =   6480
      TabIndex        =   20
      Tag             =   "2"
      ToolTipText     =   "Standard Or Actual Cost"
      Top             =   5760
      Width           =   1935
   End
   Begin VB.TextBox txtCav 
      Height          =   285
      Left            =   6480
      TabIndex        =   18
      Tag             =   "2"
      ToolTipText     =   "Standard Or Actual Cost"
      Top             =   5400
      Width           =   1935
   End
   Begin VB.TextBox txtCGSPOP 
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Tag             =   "2"
      ToolTipText     =   "Tool Description (40 chars)"
      Top             =   2520
      Width           =   1065
   End
   Begin VB.TextBox txtCustPO 
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Tag             =   "2"
      ToolTipText     =   "Tool Description (40 chars)"
      Top             =   2040
      Width           =   2745
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Left            =   6120
      TabIndex        =   9
      Tag             =   "2"
      ToolTipText     =   "Tool Description (40 chars)"
      Top             =   2520
      Width           =   1305
   End
   Begin VB.TextBox txtHmGrid 
      Height          =   285
      Left            =   6120
      TabIndex        =   7
      Tag             =   "2"
      ToolTipText     =   "Tool Description (40 chars)"
      Top             =   2040
      Width           =   2025
   End
   Begin VB.TextBox txtHmShelf 
      Height          =   285
      Left            =   6120
      TabIndex        =   5
      Tag             =   "2"
      ToolTipText     =   "Tool Description (40 chars)"
      Top             =   1560
      Width           =   2025
   End
   Begin VB.TextBox txtHomeAisle 
      Height          =   285
      Left            =   6120
      TabIndex        =   3
      Tag             =   "2"
      ToolTipText     =   "Tool Description (40 chars)"
      Top             =   1095
      Width           =   2025
   End
   Begin VB.TextBox txtHmBldg 
      Height          =   285
      Left            =   6120
      TabIndex        =   1
      Tag             =   "2"
      ToolTipText     =   "Tool Description (40 chars)"
      Top             =   660
      Width           =   1995
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ToolTLe06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   40
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbToolMat 
      Height          =   315
      Left            =   1920
      TabIndex        =   25
      Tag             =   "3"
      ToolTipText     =   "Select Customer From List"
      Top             =   7080
      Width           =   1555
   End
   Begin VB.ComboBox txtDtAdded 
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Tag             =   "4"
      ToolTipText     =   "Don't Use After"
      Top             =   1095
      Width           =   1095
   End
   Begin VB.TextBox txtCmt 
      Height          =   1215
      Left            =   1920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      ToolTipText     =   "1000 Chars Max"
      Top             =   3960
      Width           =   4695
   End
   Begin VB.CheckBox optExp 
      Alignment       =   1  'Right Justify
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   9480
      TabIndex        =   31
      Top             =   8760
      Width           =   715
   End
   Begin VB.ComboBox cmbCls 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   13
      Tag             =   "2"
      ToolTipText     =   "12 Char Class - Retrieved From Previous Entries"
      Top             =   3480
      Width           =   1935
   End
   Begin VB.ComboBox cmbTol 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Enter a New Tool or Select From List (30 chars)"
      Top             =   660
      Width           =   2775
   End
   Begin VB.TextBox txtCGPO 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Tag             =   "2"
      ToolTipText     =   "Tool Description (40 chars)"
      Top             =   1560
      Width           =   2745
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   7920
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7680
      Top             =   6960
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   8505
      FormDesignWidth =   8895
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO #"
      Height          =   285
      Index           =   10
      Left            =   4080
      TabIndex        =   67
      Top             =   8040
      Width           =   435
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Inventory"
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   65
      Top             =   8040
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit #"
      Height          =   255
      Index           =   7
      Left            =   6000
      TabIndex        =   64
      Top             =   3480
      Width           =   795
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Code"
      Height          =   255
      Index           =   5
      Left            =   4080
      TabIndex        =   63
      Top             =   3480
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      Height          =   285
      Index           =   4
      Left            =   6120
      TabIndex        =   62
      Top             =   3000
      Width           =   795
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Gov Prime Contract"
      Height          =   285
      Index           =   3
      Left            =   2280
      TabIndex        =   61
      Top             =   3000
      Width           =   1515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "ITAR"
      Height          =   255
      Index           =   2
      Left            =   4830
      TabIndex        =   59
      Top             =   7605
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Disposition Status"
      Height          =   255
      Index           =   38
      Left            =   3960
      TabIndex        =   58
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Status"
      Height          =   255
      Index           =   37
      Left            =   240
      TabIndex        =   57
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Weight"
      Height          =   255
      Index           =   36
      Left            =   4560
      TabIndex        =   56
      Top             =   6480
      Width           =   1515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Also Makes PNs"
      Height          =   255
      Index           =   35
      Left            =   240
      TabIndex        =   55
      Top             =   6480
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Blanket PO / Contract #"
      Height          =   255
      Index           =   34
      Left            =   4560
      TabIndex        =   54
      Top             =   6120
      Width           =   1755
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Serial Number"
      Height          =   255
      Index           =   33
      Left            =   240
      TabIndex        =   53
      Top             =   6120
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dimensions (WDH)"
      Height          =   255
      Index           =   32
      Left            =   4560
      TabIndex        =   52
      Top             =   5760
      Width           =   1515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Accountable to"
      Height          =   255
      Index           =   31
      Left            =   240
      TabIndex        =   51
      Top             =   5760
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "# of Cavities"
      Height          =   255
      Index           =   30
      Left            =   4560
      TabIndex        =   50
      Top             =   5400
      Width           =   1635
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Government Owned"
      Height          =   285
      Index           =   29
      Left            =   240
      TabIndex        =   49
      Top             =   3000
      Width           =   1515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "CG SO #"
      Height          =   285
      Index           =   28
      Left            =   240
      TabIndex        =   48
      Top             =   2520
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer PO#"
      Height          =   285
      Index           =   27
      Left            =   240
      TabIndex        =   47
      Top             =   2040
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Loc #"
      Height          =   285
      Index           =   26
      Left            =   5040
      TabIndex        =   46
      Top             =   2520
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Home Grid"
      Height          =   285
      Index           =   25
      Left            =   5040
      TabIndex        =   45
      Top             =   2040
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Home Shelf"
      Height          =   285
      Index           =   24
      Left            =   5040
      TabIndex        =   44
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Home Aisle"
      Height          =   285
      Index           =   23
      Left            =   5040
      TabIndex        =   43
      Top             =   1095
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Home Bldg"
      Height          =   285
      Index           =   18
      Left            =   5040
      TabIndex        =   42
      Top             =   720
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "CG PO#"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   41
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Material "
      Height          =   255
      Index           =   22
      Left            =   240
      TabIndex        =   39
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "In Service"
      Height          =   255
      Index           =   19
      Left            =   240
      TabIndex        =   38
      Top             =   1095
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments:"
      Height          =   255
      Index           =   17
      Left            =   240
      TabIndex        =   37
      Top             =   3960
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
      Height          =   255
      Index           =   15
      Left            =   7680
      TabIndex        =   36
      Top             =   6960
      Width           =   15
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
      Height          =   255
      Index           =   14
      Left            =   7800
      TabIndex        =   35
      Top             =   6960
      Width           =   15
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Owner"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   34
      Top             =   5400
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Class"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   33
      Top             =   3480
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   32
      Top             =   660
      Width           =   1155
   End
End
Attribute VB_Name = "ToolTLe06a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'8/19/04 New
Option Explicit
Dim RdoTool As ADODB.Recordset

Dim bCancel As Byte
Dim bGoodTool As Byte
Dim bOnLoad As Byte
Dim bPartExists As Byte

Dim sOldClass As String
Dim sOldLast As Variant
Dim sOldNext As Variant

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub chkGovOwned_Click()
   
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
'         !TOOL_GOVOWNED = IIf(chkGovOwned.Value = vbChecked, 1, 0)
         !TOOL_GOVOWNED = chkGovOwned.Value
         .Update
      End With
   End If

End Sub


Private Sub cmbCategory_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         !TOOL_CATEGORY = Trim(cmbCategory)
         .Update
      End With
   End If
End Sub

Private Sub cmbCls_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   cmbCls = CheckLen(cmbCls, 12)
   cmbCls = StrCase(cmbCls)
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         !TOOL_CLASS = Trim(cmbCls)
         .Update
      End With
   End If
End Sub



Private Sub cmbLastPhyInv_DropDown()
   ShowCalendar Me
End Sub

Private Sub cmbLastPhyInv_LostFocus()
   If Len(Trim(cmbLastPhyInv)) > 0 Then cmbLastPhyInv = CheckDate(cmbLastPhyInv)
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         If Len(Trim(cmbLastPhyInv)) Then
            !TOOL_LASTINVDATE = Format(cmbLastPhyInv, "mm/dd/yy")
         Else
            !TOOL_LASTINVDATE = Null
         End If
         .Update
      End With
   End If

End Sub

Private Sub cmbTol_Click()
   bGoodTool = GetThisTool()
   
End Sub

Private Sub cmbTol_LostFocus()
   cmbTol = CheckLen(cmbTol, 30)
   If bCancel = 1 Then Exit Sub
   bGoodTool = GetThisTool()
   If bGoodTool = 0 Then AddNewTool

End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = 1
   
End Sub


Private Sub cmdDelete_Click()
   If cmbTol = "" Then Exit Sub
   Dim Tool As String
   Tool = cmbTol
   If MsgBox("Delete tool " & Tool & "?", ES_YESQUESTION, Caption) Then
      If DeleteCurrentTool() Then
         MsgBox "Deletion of tool " & Tool & " succeeded"
      Else
         MsgBox "Deletion of tool " & Tool & " failed"
      End If
   End If
End Sub

Private Function DeleteCurrentTool() As Boolean
   
   clsADOCon.BeginTrans
   Dim success As Boolean
   success = True
   
   sSql = "DELETE FROM TlitTableNew WHERE TOOL_NUMREF = '" & Compress(cmbTol) & "'"
   success = clsADOCon.ExecuteSql(sSql)
   
   If success Then
      sSql = "DELETE FROM TlnhdTableNew WHERE TOOL_NUMREF = '" & Compress(cmbTol) & "'"
      success = clsADOCon.ExecuteSql(sSql)
   End If
   
   If success Then
      clsADOCon.CommitTrans
      FillCombo
      BlankOutFields
   Else
      clsADOCon.RollbackTrans
   End If
   
   DeleteCurrentTool = success

End Function

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3401
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub BlankOutFields()
      txtDtAdded = ""
      txtCGPO = ""
      txtCustPO = ""
      txtCGSPOP = ""
      chkGovOwned = vbUnchecked
      cmbCls = ""
      txtHmBldg = ""
      txtHomeAisle = ""
      txtHmShelf = ""
      txtHmGrid = ""
      txtLoc = ""
      txtCmt = ""
      cmbCst = ""
      cmbAcctTo = ""
      txtSlNo = ""
      txtPN = ""
      txtCav = ""
      txtDim = ""
      txtBlankedPO = ""
      txtWeight = ""
      cmbToolMat = ""
      chkITAR = vbUnchecked
      cmbServ = ""
      cmbDisp = ""
      
      txtGovPrimeCont = ""
      cmbCategory.ListIndex = 0
      txtCode.Text = ""
      txtUnitNo = ""
      cmbLastPhyInv = ""

End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillToolClass
      'FillToolOwner
      FillCustomers
      FillToolAcct
      FillToolMat
      FillService
      FillDisp
      FillToolCategory
      FillCombo
   End If
   bOnLoad = 0
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
   Set RdoTool = Nothing
   Set ToolTLe06a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub
Private Sub FillDisp()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT TOOL_DISPSTAT FROM TlnhdTableNew ORDER BY TOOL_DISPSTAT"
   LoadComboBox cmbDisp, -1
   Exit Sub
   
DiaErr1:
   sProcName = "FillDisp"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FillService()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT TOOL_SRVSTAT FROM TlnhdTableNew ORDER BY TOOL_SRVSTAT"
   LoadComboBox cmbServ, -1
   Exit Sub
   
DiaErr1:
   sProcName = "FillService"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Sub FillToolMat()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT TOOL_TOOLMATSTAT FROM TlnhdTableNew ORDER BY TOOL_TOOLMATSTAT"
   LoadComboBox cmbToolMat, -1
   Exit Sub
   
DiaErr1:
   sProcName = "FillToolMat"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FillToolAcct()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT TOOL_ACCTTO FROM TlnhdTableNew ORDER BY TOOL_ACCTTO"
   LoadComboBox cmbAcctTo, -1
   Exit Sub
   
DiaErr1:
   sProcName = "FillToolAcct"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Sub FillToolOwner()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT TOOL_OWNER FROM TlnhdTableNew ORDER BY TOOL_OWNER"
   LoadComboBox cmbCst, -1
   Exit Sub
   
DiaErr1:
   sProcName = "FillToolOwner"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FillToolCategory()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT ToolCategory from ToolNewCategories order by ToolCategory"
   LoadComboBox cmbCategory, -1
   Exit Sub
   
DiaErr1:
   sProcName = "FillToolCategory"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub



Private Sub FillToolClass()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT TOOL_CLASS FROM TlnhdTableNew WHERE TOOL_CLASS IS NOT NULL ORDER BY TOOL_CLASS "
   LoadComboBox cmbCls, -1
   Exit Sub
   
DiaErr1:
   sProcName = "FillToolClass"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Sub FillCombo()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT TOOL_NUM FROM TlnhdTableNew ORDER BY TOOL_NUM "
   LoadComboBox cmbTol, -1, False
'   If cmbTol.ListCount > 0 Then
'      cmbTol = cmbTol.List(0)
'      sSql = "SELECT DISTINCT TOOL_CLASS FROM TlnhdTableNew WHERE TOOL_CLASS<>'' ORDER BY TOOL_CLASS "
'      LoadComboBox cmbCls, -1
'      bGoodTool = GetThisTool()
'   End If

   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetThisTool() As Byte
   On Error GoTo DiaErr1
   
   cmdDelete.Enabled = False
   
   sSql = "SELECT TOOL_NUM, TOOL_DTADDED, TOOL_CGPONUM, TOOL_CUSTPONUM, TOOL_CGSOPONUM, TOOL_GOVOWNED, TOOL_CLASS, TOOL_HOMEBLDG," & vbCrLf _
          & "TOOL_HOMEAISLE, TOOL_SHELFNUM, TOOL_GRID, TOOL_LOCNUM, TOOL_COMMENTS, TOOL_OWNER, TOOL_ACCTTO, TOOL_SN, TOOL_MAKEPN," & vbCrLf _
          & "TOOL_CAVNUM, TOOL_DIM, TOOL_BLANKPONUM, TOOL_MONUM, TOOL_WEIGHT, TOOL_TOOLMATSTAT, TOOL_SRVSTAT, TOOL_DISPSTAT, TOOL_ITAR, " & vbCrLf _
          & "TOOL_CODE, TOOL_UNITNUM, TOOL_GOVPRIMECONTRACT, TOOL_CATEGORY, TOOL_LASTINVDATE " & vbCrLf _
          & "FROM TlnhdTableNew " & vbCrLf _
          & "WHERE TOOL_NUMREF='" & Compress(cmbTol) & "'"

   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTool, ES_KEYSET)
   If bSqlRows Then
      With RdoTool
         'txtQty = Format(!TOOL_QOH, "#######0")
         cmdDelete.Enabled = True
         cmbTol = "" & Trim(!TOOL_NUM)
         txtDtAdded = "" & Trim(!TOOL_DTADDED)
         txtCGPO = "" & Trim(!TOOL_CGPONUM)
         txtCustPO = "" & Trim(!TOOL_CUSTPONUM)
         txtCGSPOP = "" & Trim(!TOOL_CGSOPONUM)
         chkGovOwned = IIf(!TOOL_GOVOWNED, vbChecked, vbUnchecked)
'         If IsNull(!TOOL_GOVOWNED) Then
'            chkGovOwned = vbUnchecked
'         Else
'            chkGovOwned = !TOOL_GOVOWNED
'         End If
         cmbCls = "" & Trim(!TOOL_CLASS)
         txtHmBldg = "" & Trim(!TOOL_HOMEBLDG)
         txtHomeAisle = "" & Trim(!TOOL_HOMEAISLE)
         txtHmShelf = "" & Trim(!TOOL_SHELFNUM)
         txtHmGrid = "" & Trim(!TOOL_GRID)
         txtLoc = "" & Trim(!TOOL_LOCNUM)
         txtCmt = "" & Trim(!TOOL_COMMENTS)
         cmbCst = "" & Trim(!TOOL_OWNER)
         cmbAcctTo = "" & Trim(!TOOL_ACCTTO)
         txtSlNo = "" & Trim(!TOOL_SN)
         txtPN = "" & Trim(!TOOL_MAKEPN)
         txtCav = "" & Trim(!TOOL_CAVNUM)
         txtDim = "" & Trim(!TOOL_DIM)
         txtBlankedPO = "" & Trim(!TOOL_BLANKPONUM)
         txtMO = "" & !TOOL_MONUM
         txtWeight = "" & Trim(!TOOL_WEIGHT)
         cmbToolMat = "" & Trim(!TOOL_TOOLMATSTAT)
         chkITAR = IIf(!TOOL_ITAR, vbChecked, vbUnchecked)
'         If IsNull(!TOOL_ITAR) Then
'            ChkITAR = vbUnchecked
'         Else
'            ChkITAR = !TOOL_ITAR
'         End If
         cmbServ = "" & Trim(!TOOL_SRVSTAT)
         cmbDisp = "" & Trim(!TOOL_DISPSTAT)
         
         'Added Sept 2017: TOOL_CODE, TOOL_UNITNUM, TOOL_GOVPRIMECONTRACT, TOOL_CATEGORY, TOOL_LASTINVDATE
         txtGovPrimeCont = "" & !TOOL_GOVPRIMECONTRACT
         
         'find item in the dropdown list
         Dim i As Integer
         For i = 0 To cmbCategory.ListCount - 1
            If cmbCategory.List(i) = !TOOL_CATEGORY Then
               cmbCategory.ListIndex = i
               Exit For
            End If
         Next
         
         'cmbCategory.Text = !TOOL_CATEGORY
         'cmbCategory.Text = "3"
         txtCode.Text = "" & !TOOL_CODE
         txtUnitNo = "" & !TOOL_UNITNUM
         cmbLastPhyInv = IIf(IsNull(!TOOL_LASTINVDATE), "", !TOOL_LASTINVDATE)
         
      End With
      GetThisTool = 1
   Else
      GetThisTool = 0
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getthisto"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AddNewTool()
   Dim bResponse As Byte
   Dim sMsg As String
   If Len(cmbTol) < 3 Then Exit Sub
   'On Error Resume Next
   bResponse = MsgBox("Add New Tool " & cmbTol & "?", _
               ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      'bPartExists = GetPartNumber()
      
      BlankOutFields
      
      bResponse = IllegalCharacters(cmbTol)
      If bResponse > 0 Then
         MsgBox "The Part Number Contains An Illegal " & Chr$(bResponse) & ".", _
            vbExclamation, Caption
         Exit Sub
      Else
      'Add it
         sSql = "INSERT INTO TlnhdTableNew (TOOL_NUMREF, TOOL_NUM) " _
                & "VALUES('" & Compress(Trim(cmbTol)) & "', '" & Trim(cmbTol) & "')"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
         
         SysMsg "The Tool Was Created.", True
         cmbTol.AddItem cmbTol
         bGoodTool = GetThisTool()

      End If
   Else
      CancelTrans
   End If
   
End Sub

Private Function GetPartNumber() As Byte
   Dim RdoPrt As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF FROM PartTable WHERE PARTREF='" _
          & Compress(cmbTol) & " '"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
   If bSqlRows Then GetPartNumber = 1 Else GetPartNumber = 0
   Set RdoPrt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpartnum"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Sub txtCode_LostFocus()
   If bGoodTool = 1 Then
      txtCode = Trim(txtCode)
      txtCode = CheckLen(txtCode, 6)
      With RdoTool
         !TOOL_CODE = txtCode
         .Update
      End With
   End If

End Sub

Private Sub txtDtAdded_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtDtAdded_LostFocus()
   If Len(Trim(txtDtAdded)) > 0 Then txtDtAdded = CheckDate(txtDtAdded)
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         If Len(Trim(txtDtAdded)) Then
            !TOOL_DTADDED = Format(txtDtAdded, "mm/dd/yy")
         Else
            !TOOL_DTADDED = Null
         End If
         .Update
      End With
   End If
   
End Sub

Private Sub txtCGPO_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_CGPONUM = txtCGPO
         .Update
      End With
   End If
   
End Sub

Private Sub txtCustPO_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_CUSTPONUM = txtCustPO
         .Update
      End With
      clsADOCon.ExecuteSql sSql 'rdExecDirect
   End If
   
End Sub


Private Sub txtCGSPOP_LostFocus()
   'txtRev = CheckLen(txtCurRev, 6)
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_CGSOPONUM = txtCGSPOP
         .Update
      End With
   End If
   
End Sub

Private Sub txtGovPrimeCont_LostFocus()
   If bGoodTool = 1 Then
      txtGovPrimeCont = CheckLen(txtGovPrimeCont, 20)
      With RdoTool
         '.Edit
         !TOOL_GOVPRIMECONTRACT = txtGovPrimeCont
         .Update
      End With
   End If
End Sub


Private Sub txtHmBldg_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         !TOOL_HOMEBLDG = txtHmBldg
         .Update
      End With
   End If
   
End Sub

Private Sub txtHomeAisle_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_HOMEAISLE = txtHomeAisle
         .Update
      End With
   End If
   
End Sub

Private Sub txtHmShelf_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_SHELFNUM = txtHmShelf
         .Update
      End With
   End If
   
End Sub

Private Sub txtHmGrid_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_GRID = txtHmGrid
         .Update
      End With
   End If
   
End Sub



Private Sub txtLoc_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_LOCNUM = txtLoc
         .Update
      End With
   End If
   
End Sub


Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 4080)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   txtCmt = ReplaceString(txtCmt)
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_COMMENTS = Trim(txtCmt)
         .Update
      End With
   End If
   
End Sub

Private Sub cmbCst_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_OWNER = cmbCst
         .Update
      End With
   End If

End Sub


Private Sub cmbAcctTo_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_ACCTTO = cmbAcctTo
         .Update
      End With
   End If

End Sub

Private Sub txtSlNo_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         !TOOL_SN = txtSlNo
         .Update
      End With
   End If
   
End Sub

Private Sub txtMO_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         !TOOL_MONUM = txtMO
         .Update
      End With
   End If
   
End Sub

Private Sub txtPN_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         !TOOL_MAKEPN = txtPN
         .Update
      End With
   End If
   
End Sub


Private Sub txtCav_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_CAVNUM = txtCav
         .Update
      End With
   End If
   
End Sub


Private Sub txtDim_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_DIM = txtDim
         .Update
      End With
   End If
   
End Sub


Private Sub txtBlankedPO_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_BLANKPONUM = txtBlankedPO
         .Update
      End With
   End If
   
End Sub


Private Sub txtUnitNo_LostFocus()
   If bGoodTool = 1 Then
      txtUnitNo = Trim(txtUnitNo)
      If Len(txtUnitNo) > 1 Or Not IsNumeric(txtUnitNo) Then
         MsgBox "unit number must be a single digit integer or a blank"
         txtCode.SetFocus
         Exit Sub
      End If
      On Error Resume Next
      With RdoTool
         !TOOL_UNITNUM = txtUnitNo
         .Update
      End With
   End If
End Sub

Private Sub txtWeight_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_WEIGHT = txtWeight
         .Update
      End With
   End If
   
End Sub

Private Sub chkITAR_Click()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         '!TOOL_ITAR = IIf(ChkITAR.Value = vbChecked, 1, 0)
         !TOOL_ITAR = chkITAR.Value
         .Update
      End With
   End If
   
End Sub


Private Sub cmbToolMat_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_TOOLMATSTAT = cmbToolMat
         .Update
      End With
   End If
   
End Sub

Private Sub cmbServ_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_SRVSTAT = cmbServ
         .Update
      End With
   End If
   
End Sub

Private Sub cmbDisp_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_DISPSTAT = cmbDisp
         .Update
      End With
   End If
   
End Sub


