VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PurcPRf07a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Split A Purchase Order Item"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   Begin VB.Frame z3 
      Height          =   40
      Left            =   240
      TabIndex        =   23
      Top             =   1800
      Width           =   6132
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRf07a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdSplit 
      Cancel          =   -1  'True
      Caption         =   "S&plit"
      Height          =   315
      Left            =   5520
      TabIndex        =   6
      ToolTipText     =   "Split This Item"
      Top             =   2880
      Visible         =   0   'False
      Width           =   875
   End
   Begin VB.Frame z2 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Visible         =   0   'False
      Width           =   6375
      Begin VB.TextBox txtRev 
         Height          =   285
         Left            =   4320
         TabIndex        =   5
         Tag             =   "3"
         ToolTipText     =   "Item Rev (1 Ucase Letter)"
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox txtItm 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "1"
         ToolTipText     =   "New Item Number"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtQty 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Tag             =   "1"
         ToolTipText     =   "Enter Quantity To Be Split From The PO Item"
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox cmbAll 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   240
         TabIndex        =   19
         Tag             =   "8"
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cmbItm 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Tag             =   "8"
         ToolTipText     =   "Contains Only Qualifying Items"
         Top             =   120
         Width           =   855
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Number"
         Height          =   255
         Index           =   7
         Left            =   2640
         TabIndex        =   21
         Top             =   840
         Width           =   975
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Split Qty"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblQty 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3720
         TabIndex        =   18
         Top             =   120
         Width           =   975
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity "
         Height          =   255
         Index           =   5
         Left            =   2520
         TabIndex        =   17
         Top             =   120
         Width           =   975
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Part Number"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblPart 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Items"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "S&elect"
      Height          =   315
      Left            =   5520
      TabIndex        =   1
      ToolTipText     =   "Retrieve Purchase Order Items"
      Top             =   720
      Width           =   875
   End
   Begin VB.ComboBox cmbPon 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Select Or Enter PO (Contains PO's With Open Items)"
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   5520
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   4200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3510
      FormDesignWidth =   6465
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Orders With Open Items Qualify"
      ForeColor       =   &H00800000&
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   360
      Width           =   3972
   End
   Begin VB.Label lblNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   11
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label cmbVnd 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "PurcPRf07a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'7/2/04 New
Option Explicit
Dim bOnLoad As Byte

Dim sPoItems(500, 5) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbItm_Click()
   On Error Resume Next
   lblPart = sPoItems(cmbItm.ListIndex, 2)
   lblQty = sPoItems(cmbItm.ListIndex, 3)
   txtQty = "0.000"
   txtItm = Val(cmbItm)
   txtRev = ""
   
End Sub


Private Sub cmbItm_LostFocus()
   On Error Resume Next
   If Trim(cmbItm) = "" Then
      cmbItm = cmbItm.List(0)
      lblPart = sPoItems(0, 2)
      lblQty = sPoItems(0, 3)
   Else
      lblPart = sPoItems(cmbItm.ListIndex, 2)
      lblQty = sPoItems(cmbItm.ListIndex, 3)
   End If
   txtItm = Val(cmbItm)
   
End Sub


Private Sub cmbPon_Click()
   GetVendor
   
End Sub


Private Sub cmbPon_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   cmbPon = CheckLen(cmbPon, 6)
   cmbPon = Format(Abs(Val(cmbPon)), "000000")
   For iList = 0 To cmbPon.ListCount - 1
      If cmbPon = cmbPon.List(iList) Then bByte = 1
   Next
   If bByte = 1 Then
      GetVendor
   Else
      Beep
      cmbVnd = ""
      lblNme = "*** Qualifying Purchase Order Wasn't Found ***"
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4356
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdSel_Click()
   If lblNme.ForeColor = vbBlack Then GetItems
   
End Sub

Private Sub cmdSplit_Click()
   Dim bByte As Byte
   Dim iList As Integer
   Dim sItem As String
   
   If Val(txtQty) >= Val(lblQty) Then
      MsgBox "The Split Quantity Must Be Less The Original.", _
         vbInformation, Caption
      Exit Sub
   End If
   If Val(txtQty) = 0 Then
      MsgBox "The Split Quantity Must Be More Than Zero.", _
         vbInformation, Caption
      Exit Sub
   End If
   If Val(txtItm) = 0 Then
      MsgBox "The Requires A Valid Item Number.", _
         vbInformation, Caption
      Exit Sub
   End If
   sItem = Trim$(txtItm) & Trim$(txtRev)
   For iList = 0 To cmbAll.ListCount - 1
      If sItem = cmbAll.List(iList) Then bByte = 1
   Next
   If bByte = 1 Then
      MsgBox "That Item Number Is In Use.", _
         vbInformation, Caption
      Exit Sub
   End If
   bByte = MsgBox("Spit This Item As Noted?", _
           ES_YESQUESTION, Caption)
   If bByte = vbYes Then SplitItem Else CancelTrans
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
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
   FormUnload
   Set PurcPRf07a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   lblNme.ForeColor = vbBlack
   txtQty = "0.000"
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT PINUMBER FROM PoitTable WHERE " _
          & "PITYPE=14 ORDER BY PINUMBER DESC"
   LoadNumComboBox cmbPon, "000000"
   If cmbPon.ListCount > 0 Then cmbPon = cmbPon.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetVendor() As Byte
   Dim RdoVdr As ADODB.Recordset
   On Error GoTo DiaErr1
   lblPart = ""
   lblQty = ""
   'Height = 2310
   z2.Visible = False
   cmdSplit.Visible = False
   sSql = "SELECT POVENDOR FROM PohdTable WHERE PONUMBER=" & Val(cmbPon) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVdr, ES_FORWARD)
   If bSqlRows Then
      With RdoVdr
         cmbVnd = "" & Trim(!POVENDOR)
         ClearResultSet RdoVdr
      End With
      FindVendor cmbVnd, lblNme
   Else
      cmbVnd = ""
      lblNme = "*** Qualifying Purchase Order Wasn't Found ***"
   End If
   Set RdoVdr = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getvendor"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub lblNme_Change()
   If Left(lblNme, 8) = "*** Qual" Then
      lblNme.ForeColor = ES_RED
   Else
      lblNme.ForeColor = vbBlack
   End If
   
End Sub


Private Sub GetItems()
   Dim iItem As Integer
   Dim RdoPoi As ADODB.Recordset
   Erase sPoItems
   cmbItm.Clear
   cmbAll.Clear
   iItem = -1
   On Error GoTo DiaErr1
   sSql = "SELECT PINUMBER,PIITEM,PIREV,PITYPE,PIPART,PIPQTY,PARTREF,PARTNUM " _
          & "FROM PoitTable,PartTable WHERE (PIPART=PARTREF AND PINUMBER=" _
          & Val(cmbPon) & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPoi, ES_FORWARD)
   If bSqlRows Then
      With RdoPoi
         Do Until .EOF
            If !PITYPE = 14 Then
               iItem = iItem + 1
               sPoItems(iItem, 0) = Trim$(!PIITEM)
               sPoItems(iItem, 1) = Trim(!PIREV)
               sPoItems(iItem, 2) = Trim(!PartNum)
               sPoItems(iItem, 3) = Format$(!PIPQTY, ES_QuantityDataFormat)
               AddComboStr cmbItm.hwnd, "" & Format$(!PIITEM, "##0") & Trim(!PIREV)
            End If
            AddComboStr cmbAll.hwnd, "" & Format$(!PIITEM, "##0") & Trim(!PIREV)
            .MoveNext
         Loop
         ClearResultSet RdoPoi
      End With
   End If
   txtQty = "0.000"
   txtItm = ""
   txtRev = ""
   If cmbItm.ListCount > 0 Then
      On Error Resume Next
      cmbItm = cmbItm.List(0)
      cmbAll = cmbAll.List(0)
      cmbItm.ListIndex = 0
      txtItm = Val(cmbItm)
      lblPart = sPoItems(0, 2)
      lblQty = sPoItems(0, 3)
      cmdSplit.Visible = True
      z2.Visible = True
      'Height = 4215
   Else
      cmdSplit.Visible = False
      lblPart = ""
      lblQty = "0.000"
      z2.Visible = False
      'Height = 2310
   End If
   Set RdoPoi = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub




Private Sub txtItm_LostFocus()
   txtItm = CheckLen(txtItm, 3)
   txtItm = Format(Abs(Val(txtItm)), "##0")
   
End Sub


Private Sub txtQty_LostFocus()
   txtQty = CheckLen(txtQty, 9)
   txtQty = Format(Abs(Val(txtQty)), ES_QuantityDataFormat)
   
End Sub


Private Sub txtRev_LostFocus()
   txtRev = CheckLen(txtRev, 1)
   If Val(txtRev) > 0 Then
      MsgBox "Must Be A Character A - Z", _
         vbInformation, Caption
   End If
   
End Sub



Private Sub SplitItem()
   Dim RdoSplit As ADODB.Recordset
   Dim sComment As String
   On Error GoTo DiaErr1
   sSql = "SELECT PINUMBER,PIITEM,PIPART,PIPDATE,PIPQTY," _
          & "PIESTUNIT,PIADDERS,PILOT,PIRUNPART,PIRUNNO,PIRUNOPNO," _
          & "PICOMT FROM PoitTable WHERE (PINUMBER=" & Val(cmbPon) & " " _
          & "AND PIITEM=" & sPoItems(cmbItm.ListIndex, 0) & " AND PIREV='" _
          & sPoItems(cmbItm.ListIndex, 1) & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSplit, ES_STATIC)
   If bSqlRows Then
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      
      On Error Resume Next
      Err = 0
      
      With RdoSplit
         sComment = ReplaceString("" & Trim(!PICOMT))
         sSql = "INSERT INTO PoitTable (PINUMBER,PIITEM,PIREV,PITYPE,PIPART," _
                & "PIPDATE,PIPQTY,PIESTUNIT,PIADDERS,PILOT,PIRUNPART,PIRUNNO,PIRUNOPNO," _
                & "PICOMT) Values(" & !PINUMBER & "," _
                & Val(txtItm) & ",'" & Trim(txtRev) & "',14,'" _
                & Trim(!PIPART) & "','" & !PIPDATE & "'," _
                & Val(txtQty) & "," & !PIESTUNIT & "," _
                & !PIADDERS & "," & !PILOT & ",'" _
                & Trim(!PIRUNPART) & "'," & !PIRUNNO & "," _
                & !PIRUNOPNO & ",'" & sComment & "')"
         clsADOCon.ExecuteSQL sSql
         ClearResultSet RdoSplit
      End With
      sSql = "UPDATE PoitTable SET PIPQTY=PIPQTY-" & Val(txtQty) & " " _
             & "WHERE (PINUMBER=" & Val(cmbPon) & " " _
             & "AND PIITEM=" & sPoItems(cmbItm.ListIndex, 0) & " AND PIREV='" _
             & sPoItems(cmbItm.ListIndex, 1) & "')"
      clsADOCon.ExecuteSQL sSql
      
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         Sleep 500
         sSql = "UPDATE PoitTable SET PIPRESPLITFROM='" & cmbItm _
                & "' WHERE (PINUMBER=" & Val(cmbPon) & " " _
                & "AND PIITEM=" & Val(txtItm) & " AND PIREV='" _
                & Trim(txtRev) & "')"
         clsADOCon.ExecuteSQL sSql
         
         SysMsg "Item Was Split.", True
         GetVendor
      Else
         clsADOCon.RollbackTrans
         MsgBox "Couldn't Split This Item." & vbCr & Left(Err.Description, 30), _
            vbExclamation, Caption
      End If
   Else
      MsgBox "Couldn't Resolve The Item Number. Reselect.", _
         vbInformation, Caption
   End If
   Set RdoSplit = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
