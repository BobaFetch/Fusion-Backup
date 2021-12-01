VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form diaAPf09a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clear AP Aging Invoices"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   6120
      TabIndex        =   1
      Tag             =   "4"
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdcnl 
      Caption         =   "&Reselect"
      Height          =   315
      Left            =   6000
      TabIndex        =   5
      Top             =   1680
      Width           =   855
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7320
      Top             =   120
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6375
      FormDesignWidth =   8820
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4095
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Click On Check Number To Select A Check"
      Top             =   2160
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7223
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColor       =   16777215
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "&Select"
      Height          =   315
      Left            =   7800
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   8655
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      ToolTipText     =   "Nicknames"
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear Invoice"
      Height          =   315
      Left            =   7200
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   7800
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   240
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   8
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
      PictureUp       =   "diaAPf09a.frx":0000
      PictureDn       =   "diaAPf09a.frx":0146
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "CutOff Date"
      Height          =   255
      Index           =   6
      Left            =   5160
      TabIndex        =   11
      Top             =   480
      Width           =   855
   End
   Begin VB.Image imgInc 
      Height          =   180
      Left            =   3960
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgdInc 
      Height          =   180
      Left            =   3600
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   960
      TabIndex        =   10
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "diaAPf09a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Per IMAINC request, this function could have bad consequences and needed to be removed. 7/25/2019

'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          imp

'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'*************************************************************************************
' diaAPf09a - Void AP Checks
'
' Notes:
'
' Created: (nth)
' Revisons:
'
'*************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte

Dim sXC As String
Dim sCC As String
Dim sMsg As String
Dim iInvoices As Integer
Dim cCheckTotal As Currency

Dim sInvoices(100) As String
Dim vInvoice(100, 5) As Variant

Dim sChecks() As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*************************************************************************************

Private Sub FillCombo()
   Dim RdoVnd As ADODB.Recordset
   On Error GoTo DiaErr1
   
   cmbVnd.Clear
   
   sSql = "SELECT DISTINCT VENICKNAME FROM VndrTable INNER JOIN " _
          & "ChksTable ON VEREF = CHKVENDOR " _
          & "WHERE CHKVOIDDATE IS NULL ORDER BY VENICKNAME"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVnd)
   
   If bSqlRows Then
      With RdoVnd
         While Not .EOF
            AddComboStr cmbVnd.hWnd, "" & Trim(.Fields(0))
            .MoveNext
         Wend
         .Cancel
      End With
   Else
      'MsgBox "No Vendors Found.", vbInformation, Caption
   End If
   
   AddComboStr cmbVnd.hWnd, "ALL"
   
   Set RdoVnd = Nothing
   If cmbVnd.ListCount > 0 Then
      cmbVnd.ListIndex = 0
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub GetInvoices()
   Dim RdoChk As ADODB.Recordset
   Dim sGridItem As String
   Dim strDuedate As String
   Dim i As Integer
   
   On Error GoTo DiaErr1
   
   Grid1.Clear
   Grid1.rows = 1
   SetUpGrid
   strDuedate = Format(txtDte, "mm/dd/yy")
   
   sSql = "select VINO, vitype,VIDATE, VIDUEDATE, VIDUE, VIPAY, VIVENDOR" _
            & " FROM VihdTable where VIDATE <= '" & strDuedate & "' and vipif = 0 and " _
            & " VIVENDOR LIKE '" & Compress(cmbVnd) & "%' ORDER BY viduedate DESC"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk, ES_FORWARD)

   If bSqlRows Then
      i = 1
      With RdoChk
         While Not .EOF
            sGridItem = Chr(9) _
                        & Trim(!VIVENDOR) & Chr(9) _
                        & Trim(!VINO) & Chr(9) _
                        & Trim(!vitype) & Chr(9) _
                        & Format(!VIDATE, "mm/dd/yy") & Chr(9) _
                        & Format(!VIDUEDATE, "mm/dd/yy") & Chr(9) _
                        & Format(!VIDUE, "0.00") & Chr(9) _
                        & Format(!VIPAY, "0.00")
                        
                        Grid1.AddItem (sGridItem)
            Grid1.Row = i
            Grid1.Col = 0
            Grid1.CellPictureAlignment = flexAlignCenterCenter
            Set Grid1.CellPicture = imgdInc
            .MoveNext
            i = i + 1
         Wend
         .Cancel
      End With
   Else
      FillCombo
      Grid1.enabled = False
      cmdClear.enabled = False
      cmdSel.enabled = True
      cmdcnl.enabled = False
      cmbVnd.enabled = True
      cmbVnd.SetFocus
   End If
   Set RdoChk = Nothing
   
   Exit Sub
   
DiaErr1:
   sProcName = "GetInvoices"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub ClearInvoiceAmt()
   Dim cInvAmt As Currency
   Dim sVndName As String
   Dim sVino As String
   Dim iResponse As Integer
   Dim i As Integer
   Dim strCompt  As String
   
   
   On Error GoTo DiaErr1
   
   iResponse = MsgBox("Clear AP Invoice Amount to zero ?", ES_YESQUESTION, Caption)
   If iResponse = vbNo Then
      Exit Sub
   End If
   
   MouseCursor 13
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   For i = 1 To Grid1.rows - 1
      
      Grid1.Row = i
      Grid1.Col = 0
      
      If Grid1.CellPicture = imgInc Then
         
         With Grid1
            .Col = 1
            sVndName = Trim(Grid1)
            .Col = 2
            sVino = Trim(Grid1)
            .Col = 6
            cInvAmt = Val(Grid1)
         End With
         
         strCompt = "Cleared - " & CStr(cInvAmt)
         
         sSql = "UPDATE VihdTable SET VIPIF = 1, VIDUE = 0, VIPAY  = 0, VICOMT = '" & strCompt & "' " & vbCrLf _
                  & "WHERE VINO = '" & Trim(sVino) & "'" & "AND VIVENDOR = '" & Trim(sVndName) & "'"
         clsADOCon.ExecuteSql sSql
                  
         sSql = "UPDATE jritTable SET DCCREDIT = 0 " & vbCrLf _
                  & "WHERE DCVENDORINV = '" & Trim(sVino) & "'" & vbCrLf _
                  & "AND DCVENDOR = '" & Trim(sVndName) & "' AND DCDEBIT = 0"
         clsADOCon.ExecuteSql sSql
         
         sSql = "UPDATE jritTable SET DCDEBIT = 0 " & vbCrLf _
                  & "WHERE DCVENDORINV = '" & Trim(sVino) & "'" & vbCrLf _
                  & "AND DCVENDOR = '" & Trim(sVndName) & "' AND DCCREDIT = 0"
         clsADOCon.ExecuteSql sSql
         
      End If
   Next
   
   MouseCursor 0
   
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      SysMsg "Cleared Successfully AP Invoices Amount.", True
   Else
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
      MsgBox "Could Not Successfully Clear AP Invoices Amount." _
         , vbInformation, Caption
   End If
   
   Exit Sub
   
DiaErr1:
   sProcName = "ClearInvoiceAmt"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cmbVnd_Click()
   FindVendor Me
End Sub

Private Function GetVndForCheckNumber(strChkNum As String) As String
   Dim rdoChkNum As ADODB.Recordset
   Dim sVendor As String
   On Error GoTo DiaErr1
   sSql = "SELECT CHKVENDOR FROM ChksTable WHERE " _
          & "CHKNUMBER='" & strChkNum & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoChkNum)
   If bSqlRows Then
      GetVndForCheckNumber = rdoChkNum!CHKVENDOR
   Else
      GetVndForCheckNumber = ""
   End If
   On Error Resume Next
   rdoChkNum.Close
   Set rdoChkNum = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "checkinvoice"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function
Private Sub cmbVnd_LostFocus()
   'FindVendor Me
   'lblNum = NumberOfChecks(cmbVnd, True)
   
   If cmbVnd <> "" Then
      FindVendor Me
   End If
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCnl_Click()
   cmdClear.enabled = False
   cmdcnl.enabled = False
   cmdSel.enabled = True
   cmbVnd.enabled = True
   txtDte.enabled = True
   Grid1.rows = 1
   Grid1.Clear
   SetUpGrid
   'cmbVnd.SetFocus
   If (cmbVnd <> "") Then
      cmbVnd.SetFocus
   End If
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "No Subject Help"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub cmdSel_Click()
   GetInvoices
   cmdClear.enabled = True
   cmdcnl.enabled = True
   cmdSel.enabled = False
   txtDte.enabled = False
   Grid1.enabled = True
   cmbVnd.enabled = False
   Grid1.SetFocus
End Sub

Private Sub cmdClear_Click()
   ClearInvoiceAmt
   cmdClear.enabled = True
   cmdSel.enabled = False
   Grid1.SetFocus
   GetInvoices
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      MouseCursor 13
      SetUpGrid
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   imgdInc.Picture = Resources.imgdInc.Picture
   imgInc.Picture = Resources.imgInc.Picture
   txtDte = Format(ES_SYSDATE, "mm/dd/yy")
   cmdClear.enabled = False
   cmdcnl.enabled = False
   bOnLoad = True
End Sub

Private Sub SetUpGrid()
   With Grid1
      .Cols = 8
      
      .ColAlignment(0) = 0
      .ColAlignment(1) = 2
      .ColAlignment(2) = 2
      .ColAlignment(3) = 2
      .ColAlignment(4) = 2
      .ColAlignment(5) = 2
      .ColAlignment(6) = 2
      .ColAlignment(7) = 2
      
      .rows = 1
      .ColWidth(0) = 500
      .ColWidth(1) = 1200
      .ColWidth(2) = 1500
      .ColWidth(3) = 550
      .ColWidth(4) = 1000
      .ColWidth(5) = 1000
      .ColWidth(6) = 1000
      .ColWidth(7) = 1000
      .Row = 0
      .Col = 0
      .Text = "Clear"
      .Col = 1
      .Text = "Vendor"
      .Col = 2
      .Text = "Invoice Number"
      .Col = 3
      .Text = "Type"
      .Col = 4
      .Text = "Invoice Date"
      .Col = 5
      .Text = "Due Date"
      .Col = 6
      .Text = "Due Amt"
      .Col = 7
      .Text = "Payed Amt"
   End With
End Sub


Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   Set diaAPf09a = Nothing
   
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub Grid1_Click()
   With Grid1
      If .Row > 0 Then
         .Col = 0
         .Row = .RowSel
         If .CellPicture = imgdInc Then
            Set .CellPicture = imgInc
         Else
            Set .CellPicture = imgdInc
         End If
      End If
   End With
End Sub

Private Sub txtDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtDte_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtDte_LostFocus()
   txtDte = CheckDate(txtDte)
End Sub
