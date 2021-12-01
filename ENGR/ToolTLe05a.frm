VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ToolTLe05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Tools"
   ClientHeight    =   8640
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
   ScaleHeight     =   8640
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbDisp 
      Height          =   315
      Left            =   5040
      TabIndex        =   23
      Tag             =   "3"
      ToolTipText     =   "Select Customer From List"
      Top             =   7560
      Width           =   1555
   End
   Begin VB.ComboBox cmbServ 
      Height          =   315
      Left            =   1680
      TabIndex        =   22
      Tag             =   "3"
      ToolTipText     =   "Select Customer From List"
      Top             =   7920
      Width           =   1555
   End
   Begin VB.TextBox txtWO 
      Height          =   285
      Left            =   6240
      TabIndex        =   20
      Tag             =   "2"
      ToolTipText     =   "Standard Or Actual Cost"
      Top             =   6840
      Width           =   1935
   End
   Begin VB.TextBox txtPN 
      Height          =   285
      Left            =   1680
      TabIndex        =   16
      Tag             =   "2"
      ToolTipText     =   "Standard Or Actual Cost"
      Top             =   6840
      Width           =   2415
   End
   Begin VB.TextBox txtPO 
      Height          =   285
      Left            =   6240
      TabIndex        =   19
      Tag             =   "2"
      ToolTipText     =   "Standard Or Actual Cost"
      Top             =   6480
      Width           =   1935
   End
   Begin VB.TextBox txtSlNo 
      Height          =   285
      Left            =   1680
      TabIndex        =   15
      Tag             =   "2"
      ToolTipText     =   "Standard Or Actual Cost"
      Top             =   6480
      Width           =   2415
   End
   Begin VB.TextBox txtDim 
      Height          =   285
      Left            =   6240
      TabIndex        =   18
      Tag             =   "2"
      ToolTipText     =   "Standard Or Actual Cost"
      Top             =   6120
      Width           =   1935
   End
   Begin VB.TextBox txtAcctTo 
      Height          =   285
      Left            =   1680
      TabIndex        =   14
      Tag             =   "2"
      ToolTipText     =   "Standard Or Actual Cost"
      Top             =   6120
      Width           =   2415
   End
   Begin VB.TextBox txtCav 
      Height          =   285
      Left            =   6240
      TabIndex        =   17
      Tag             =   "2"
      ToolTipText     =   "Standard Or Actual Cost"
      Top             =   5760
      Width           =   1935
   End
   Begin VB.TextBox txtCurRev 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Tag             =   "2"
      ToolTipText     =   "Tool Description (40 chars)"
      Top             =   3360
      Width           =   1065
   End
   Begin VB.TextBox txtUnit 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Tag             =   "2"
      ToolTipText     =   "Tool Description (40 chars)"
      Top             =   2880
      Width           =   1065
   End
   Begin VB.TextBox txtPrtName 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Tag             =   "2"
      ToolTipText     =   "Tool Description (40 chars)"
      Top             =   2400
      Width           =   2745
   End
   Begin VB.TextBox txtESI 
      Height          =   285
      Left            =   5520
      TabIndex        =   11
      Tag             =   "2"
      ToolTipText     =   "Tool Description (40 chars)"
      Top             =   2880
      Width           =   1305
   End
   Begin VB.TextBox txtGrd4 
      Height          =   285
      Left            =   5520
      TabIndex        =   10
      Tag             =   "2"
      ToolTipText     =   "Tool Description (40 chars)"
      Top             =   2400
      Width           =   2025
   End
   Begin VB.TextBox txtGrd3 
      Height          =   285
      Left            =   5520
      TabIndex        =   9
      Tag             =   "2"
      ToolTipText     =   "Tool Description (40 chars)"
      Top             =   1920
      Width           =   2025
   End
   Begin VB.TextBox txtGrd2 
      Height          =   285
      Left            =   5520
      TabIndex        =   8
      Tag             =   "2"
      ToolTipText     =   "Tool Description (40 chars)"
      Top             =   1455
      Width           =   2025
   End
   Begin VB.TextBox txtGrd1 
      Height          =   285
      Left            =   5520
      TabIndex        =   7
      Tag             =   "2"
      ToolTipText     =   "Tool Description (40 chars)"
      Top             =   1020
      Width           =   1995
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ToolTLe05a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   34
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbCTL 
      Height          =   315
      Left            =   1680
      TabIndex        =   21
      Tag             =   "3"
      ToolTipText     =   "Select Customer From List"
      Top             =   7440
      Width           =   1555
   End
   Begin VB.ComboBox txtDtAdded 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Tag             =   "4"
      ToolTipText     =   "Don't Use After"
      Top             =   1455
      Width           =   1095
   End
   Begin VB.TextBox txtCmt 
      Height          =   1215
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      ToolTipText     =   "1000 Chars Max"
      Top             =   4320
      Width           =   4695
   End
   Begin VB.CheckBox optExp 
      Alignment       =   1  'Right Justify
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   9480
      TabIndex        =   25
      Top             =   8760
      Width           =   715
   End
   Begin VB.TextBox txtOwner 
      Height          =   285
      Left            =   1680
      TabIndex        =   13
      Tag             =   "2"
      ToolTipText     =   "Standard Or Actual Cost"
      Top             =   5760
      Width           =   2415
   End
   Begin VB.ComboBox cmbCls 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   6
      Tag             =   "2"
      ToolTipText     =   "12 Char Class - Retrieved From Previous Entries"
      Top             =   3840
      Width           =   1935
   End
   Begin VB.ComboBox cmbTol 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Enter a New Tool or Select From List (30 chars)"
      Top             =   1020
      Width           =   2775
   End
   Begin VB.TextBox txtPrtFamily 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Tag             =   "2"
      ToolTipText     =   "Tool Description (40 chars)"
      Top             =   1920
      Width           =   2745
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   7920
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7800
      Top             =   7800
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   8640
      FormDesignWidth =   8895
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Disposition Status"
      Height          =   255
      Index           =   38
      Left            =   3720
      TabIndex        =   52
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Status"
      Height          =   255
      Index           =   37
      Left            =   360
      TabIndex        =   51
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Order #"
      Height          =   255
      Index           =   36
      Left            =   4800
      TabIndex        =   50
      Top             =   6840
      Width           =   1515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Also Makes PNs"
      Height          =   255
      Index           =   35
      Left            =   360
      TabIndex        =   49
      Top             =   6840
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Acquisition PO #"
      Height          =   255
      Index           =   34
      Left            =   4800
      TabIndex        =   48
      Top             =   6480
      Width           =   1755
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Serial Number"
      Height          =   255
      Index           =   33
      Left            =   360
      TabIndex        =   47
      Top             =   6480
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dimensions (WDH)"
      Height          =   255
      Index           =   32
      Left            =   4800
      TabIndex        =   46
      Top             =   6120
      Width           =   1515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Accountable to"
      Height          =   255
      Index           =   31
      Left            =   360
      TabIndex        =   45
      Top             =   6120
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "# of Cavities"
      Height          =   255
      Index           =   30
      Left            =   4800
      TabIndex        =   44
      Top             =   5760
      Width           =   1635
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Rev"
      Height          =   285
      Index           =   29
      Left            =   240
      TabIndex        =   43
      Top             =   3360
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Number"
      Height          =   285
      Index           =   28
      Left            =   240
      TabIndex        =   42
      Top             =   2880
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Name"
      Height          =   285
      Index           =   27
      Left            =   240
      TabIndex        =   41
      Top             =   2400
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "ESI #"
      Height          =   285
      Index           =   26
      Left            =   4440
      TabIndex        =   40
      Top             =   2880
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Grid4"
      Height          =   285
      Index           =   25
      Left            =   4440
      TabIndex        =   39
      Top             =   2400
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Grid3"
      Height          =   285
      Index           =   24
      Left            =   4440
      TabIndex        =   38
      Top             =   1920
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Grid2"
      Height          =   285
      Index           =   23
      Left            =   4440
      TabIndex        =   37
      Top             =   1455
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Grid1"
      Height          =   285
      Index           =   18
      Left            =   4440
      TabIndex        =   36
      Top             =   1080
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Family"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   35
      Top             =   1920
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "CTL Status"
      Height          =   255
      Index           =   22
      Left            =   360
      TabIndex        =   33
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Added"
      Height          =   255
      Index           =   19
      Left            =   240
      TabIndex        =   32
      Top             =   1455
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments:"
      Height          =   255
      Index           =   17
      Left            =   240
      TabIndex        =   31
      Top             =   4320
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
      Height          =   255
      Index           =   15
      Left            =   7440
      TabIndex        =   30
      Top             =   7320
      Width           =   15
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
      Height          =   255
      Index           =   14
      Left            =   7560
      TabIndex        =   29
      Top             =   7320
      Width           =   15
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Owner"
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   28
      Top             =   5760
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Class"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   27
      Top             =   3840
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   26
      Top             =   1020
      Width           =   1155
   End
End
Attribute VB_Name = "ToolTLe05a"
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

Private Sub cmbCls_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   cmbCls = CheckLen(cmbCls, 12)
   cmbCls = StrCase(cmbCls)
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_CLASS = Trim(cmbCls)
         .Update
      End With
   End If
   If cmbCls.ListCount > 0 Then
      For iList = 0 To cmbCls.ListCount - 1
         If UCase$(cmbCls) = UCase$(cmbCls.List(iList)) Then bByte = 1
      Next
   End If
   If bByte = 0 Then cmbCls.AddItem cmbCls
   If sOldClass <> Trim(cmbCls) Then
      sSql = "UPDATE TlnhdTable SET TOOLLISTIT_CLASS='" & Trim(cmbCls) _
             & "' WHERE TOOLLISTIT_TOOLREF='" & Compress(cmbTol) & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
   End If
   sOldClass = Trim(cmbCls)
   
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


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3401
      MouseCursor 0
      cmdHlp = False
   End If
   
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
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set RdoTool = Nothing
   Set ToolTLe05a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT TOOL_NUM FROM TlnhdTable ORDER BY TOOL_NUM "
   LoadComboBox cmbTol, -1
   If cmbTol.ListCount > 0 Then
      cmbTol = cmbTol.List(0)
      sSql = "SELECT DISTINCT TOOL_CLASS FROM TlnhdTable WHERE TOOL_CLASS<>'' ORDER BY TOOL_CLASS "
      LoadComboBox cmbCls, -1
      bGoodTool = GetThisTool()
   End If
   
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetThisTool() As Byte
   On Error GoTo DiaErr1
   
   sSql = "SELECT TOOL_NUM,TOOL_DTADDED,TOOL_PARTFAMILY,TOOL_PARTREF,TOOL_UNITNO," _
          & "TOOL_CURREV,TOOL_CLASS,TOOL_GRD1,TOOL_GRD2,TOOL_GRD3," _
          & "TOOL_GRD4,TOOL_ESI,TOOL_COMMENTS,TOOL_OWNER,TOOL_ACCTTO," _
          & "TOOL_SN,TOOL_MAKEPN,TOOL_CAVNUM,TOOL_DIM," _
          & "TOOL_PONUM,TOOL_WONUM,TOOL_CTLSTAT,TOOL_SRVSTAT,TOOL_DISPSTAT " _
          & "FROM TlnhdTable " & vbCrLf _
          & "WHERE TOOL_PARTREF='" & Compress(cmbTol) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTool, ES_KEYSET)
   If bSqlRows Then
      With RdoTool
         'txtQty = Format(!TOOL_QOH, "#######0")
         cmbTol = "" & Trim(!TOOL_NUM)
         txtDtAdded = "" & Trim(!TOOL_DTADDED)
         txtPrtFamily = "" & Trim(!TOOL_PARTFAMILY)
         txtPrtName = "" & Trim(!TOOL_PARTREF)
         txtUnit = "" & Trim(!TOOL_UNITNO)
         txtCurRev = "" & Trim(!TOOL_CURREV)
         cmbCls = "" & Trim(!TOOL_CLASS)
         txtGrd1 = "" & Trim(!TOOL_GRD1)
         txtGrd2 = "" & Trim(!TOOL_GRD2)
         txtGrd3 = "" & Trim(!TOOL_GRD3)
         txtGrd4 = "" & Trim(!TOOL_GRD4)
         txtESI = "" & Trim(!TOOL_ESI)
         txtCmt = "" & Trim(!TOOL_COMMENTS)
         txtOwner = "" & Trim(!TOOL_OWNER)
         txtAcctTo = "" & Trim(!TOOL_ACCTTO)
         txtSlNo = "" & Trim(!TOOL_SN)
         txtPN = "" & Trim(!TOOL_MAKEPN)
         txtCav = "" & Trim(!TOOL_CAVNUM)
         txtDim = "" & Trim(!TOOL_DIM)
         txtPO = "" & Trim(!TOOL_PONUM)
         txtWO = "" & Trim(!TOOL_WONUM)
         cmbCTL = "" & Trim(!TOOL_CTLSTAT)
         cmbServ = "" & Trim(!TOOL_SRVSTAT)
         cmbDisp = "" & Trim(!TOOL_DISPSTAT)
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
      bPartExists = GetPartNumber()
      
      txtDtAdded = ""
      txtPrtFamily = ""
      txtPrtName = ""
      txtUnit = ""
      txtCurRev = ""
      cmbCls = ""
      txtGrd1 = ""
      txtGrd2 = ""
      txtGrd3 = ""
      txtGrd4 = ""
      txtESI = ""
      txtCmt = ""
      txtOwner = ""
      txtAcctTo = ""
      txtSlNo = ""
      txtPN = ""
      txtCav = ""
      txtDim = ""
      txtPO = ""
      txtWO = ""
      cmbCTL = ""
      cmbServ = ""
      cmbDisp = ""
      
      If bPartExists = 1 Then
         MsgBox cmbTol & " Is In Use As A Part Number " & vbCrLf _
            & "And Cannot Be Duplicated. Pick Another.", _
            vbInformation, Caption
         Exit Sub
      Else
         bResponse = IllegalCharacters(cmbTol)
         If bResponse > 0 Then
            MsgBox "The Part Number Contains An Illegal " & Chr$(bResponse) & ".", _
               vbExclamation, Caption
            Exit Sub
         Else
            'Add it
            'Err = 0
            clsADOCon.BeginTrans
            
            Dim part As New ClassPart
            If part.CreateNewPart(cmbTol, 8, "Tool", "M") Then
            
'            sSql = "INSERT INTO PartTable (PARTREF,PARTNUM,PALEVEL,PATOOL," _
'                   & "PAQOH) VALUES('" & Compress(cmbTol) & "','" & cmbTol & "',8,1,1)"
'            clsADOCon.ExecuteSQL sSql 'rdExecDirect
            
               sSql = "UPDATE PartTable" & vbCrLf _
                  & "SET PATOOL = 1, PAQOH = 1" & vbCrLf _
                 & "WHERE PARTREF = '" & Compress(cmbTol) & "'"
               clsADOCon.ExecuteSQL sSql 'rdExecDirect
               
               sSql = "INSERT INTO TlnhdTable (TOOL_NUM,TOOL_PARTREF,TOOL_CLASS) " _
                      & "VALUES('" & Trim(cmbTol) & "','" & Compress(cmbTol) & "','" _
                      & Trim(cmbCls) & "')"
               clsADOCon.ExecuteSQL sSql 'rdExecDirect
               
               clsADOCon.CommitTrans
               SysMsg "The Tool Was Created.", True
               cmbTol.AddItem cmbTol
               bGoodTool = GetThisTool()

            Else
               clsADOCon.RollbackTrans
               MsgBox "Could Not Create The Tool.", _
                  vbExclamation, Caption
            End If
         End If
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

Private Sub txtPrtFamily_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_PARTFAMILY = txtPrtFamily
         .Update
      End With
   End If
   
End Sub


Private Sub TOOL_DTADDED_DropDown()
   ShowCalendar Me
   
End Sub



Private Sub txtUnit_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_UNITNO = txtUnit
         .Update
      End With
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
   End If
   
End Sub


Private Sub txtRev_LostFocus()
   'txtRev = CheckLen(txtCurRev, 6)
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_CURREV = txtCurRev
         .Update
      End With
   End If
   
End Sub


Private Sub txtGrd1_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_GRD1 = txtGrd1
         .Update
      End With
   End If
   
End Sub

Private Sub txtGrd2_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_GRD2 = txtGrd2
         .Update
      End With
   End If
   
End Sub

Private Sub txtGrd3_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_GRD3 = txtGrd3
         .Update
      End With
   End If
   
End Sub

Private Sub txtGrd4_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_GRD4 = txtGrd4
         .Update
      End With
   End If
   
End Sub



Private Sub txtESI_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_ESI = txtESI
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

Private Sub txtOwner_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_OWNER = txtOwner
         .Update
      End With
   End If
   
End Sub


Private Sub txtAcctTo_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_ACCTTO = txtAcctTo
         .Update
      End With
   End If
   
End Sub



Private Sub txtSlNo_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_SN = txtSlNo
         .Update
      End With
   End If
   
End Sub


Private Sub txtPN_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
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


Private Sub txtPO_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_PONUM = txtPO
         .Update
      End With
   End If
   
End Sub


Private Sub txtWO_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_WONUM = txtWO
         .Update
      End With
   End If
   
End Sub

Private Sub cmbCTL_LostFocus()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_CTLSTAT = cmbCTL
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




