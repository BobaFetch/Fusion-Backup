VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form diaAPf05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reprint Computer Checks"
   ClientHeight    =   4560
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4560
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdReprt 
      Caption         =   "&Reprint"
      Height          =   315
      Left            =   5280
      TabIndex        =   9
      ToolTipText     =   "Reprint Select Checks"
      Top             =   600
      Width           =   875
   End
   Begin VB.TextBox txtEnd 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Tag             =   "1"
      Top             =   840
      Width           =   1000
   End
   Begin VB.TextBox txtBeg 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Tag             =   "1"
      Top             =   480
      Width           =   1000
   End
   Begin VB.CheckBox optAll 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   3
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
      PictureUp       =   "diaAPf05a.frx":0000
      PictureDn       =   "diaAPf05a.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4320
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4560
      FormDesignWidth =   6240
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Click On Check Number To Select A Check"
      Top             =   1560
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5106
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColor       =   16777215
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Select All"
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   840
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Check #"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1635
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Check #"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1545
   End
   Begin VB.Image imgdInc 
      Height          =   180
      Left            =   3600
      Picture         =   "diaAPf05a.frx":028C
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgInc 
      Height          =   180
      Left            =   3960
      Picture         =   "diaAPf05a.frx":053E
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "diaAPf05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'*********************************************************************************
' diaAPf05a - Reprint Computer Checks
'
' Notes:
'
' Created: 06/25/03 (nth)
' Revisions:
'   02/10/04 (nth) Do not allow alpha numeric check numbers.
'   02/13/04 (jcw) Fixed columns to conform to Design Standard.
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim sMsg As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, _
                             Shift As Integer, X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub cmdCan_MouseUp(Button As Integer, Shift As Integer, _
                           X As Single, Y As Single)
   bCancel = False
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      SetUpGrid
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   ' Clean the check setup
   sSql = "DELETE FROM ChseTable WHERE CHKREPRINTNO <> 0"
   clsADOCon.ExecuteSQL sSql
   'Grid1.Rows = 1
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaAPf05a = Nothing
End Sub

Private Sub optAll_Click()
   If Grid1.Rows > 1 Then
      FillGrid
   End If
End Sub

Private Sub cmdReprt_Click()
   If IsCheckSelected Then
      PrintReport
   Else
      sMsg = "No Checks Selected To Reprint."
      MsgBox sMsg, vbInformation, Caption
      Grid1.SetFocus
   End If
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub PrintReport()
   On Error GoTo DiaErr1
   MouseCursor 13
   ReprintChecks
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub SetUpGrid()
   On Error GoTo DiaErr1
   
   With Grid1
      .Rows = 1
      .Cols = 6
      .ColWidth(0) = 500
      .ColWidth(1) = 1000
      .ColWidth(2) = 1000
      .ColWidth(3) = 1000
      .ColWidth(4) = 1000
      .ColWidth(5) = 1000
      .Row = 0
      .Col = 0
      .Text = "Inc"
      .Col = 1
      .Text = "Check#"
      .Col = 2
      .Text = "Vendor"
      .Col = 3
      .Text = "Print Date"
      .Col = 4
      .Text = "Amount"
      .Col = 5
      .Text = "Account"
   End With
   
   Exit Sub
DiaErr1:
   sProcName = "SetUpGrid"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FillGrid()
   Dim RdoChk As ADODB.Recordset
   Dim sItem As String
   Dim i As Integer
   Dim imgUse As Image
   
   On Error GoTo DiaErr1
   
   MouseCursor 13
   
   Grid1.Clear
   
   If optAll.Value = vbChecked Then
      Set imgUse = imgInc
   Else
      Set imgUse = imgdInc
   End If
   
   sSql = " SELECT CHKNUMBER,CHKVENDOR,CHKPRINTDATE,CHKAMOUNT,CHKACCT " _
          & "FROM ChksTable WHERE ISNUMERIC(CHKNUMBER) = 1 AND CONVERT(INT, CHKNUMBER, 101) >=" & txtBeg _
          & " AND CONVERT(INT, CHKNUMBER, 101) <= " & txtEnd _
          & " AND CHKVOID = 0 AND CHKPRINTED = 1 AND CHKTYPE = 2"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk)
   If bSqlRows Then
      SetUpGrid
      With RdoChk
         i = 1
         While Not .EOF
            sItem = Chr(9) _
                    & " " & !CHKNUMBER & Chr(9) _
                    & " " & Trim(!CHKVENDOR) & Chr(9) _
                    & " " & Format(!CHKPRINTDATE, DATEMASK) & Chr(9) _
                    & Format(!CHKAMOUNT, CURRENCYMASK) & Chr(9) _
                    & " " & Trim(!CHKACCT)
            Grid1.AddItem sItem
            Grid1.Row = i
            Grid1.Col = 0
            Grid1.CellPictureAlignment = flexAlignCenterCenter
            Set Grid1.CellPicture = imgUse
            .MoveNext
            i = i + 1
         Wend
      End With
   Else
      With Grid1
         .Rows = 1
         .Cols = 1
         .ColWidth(0) = .Width
         .Text = "*** No Checks Found ***"
      End With
   End If
   
   Set RdoChk = Nothing
   
   MouseCursor 0
   
   Exit Sub
DiaErr1:
   sProcName = "fillgrid"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
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



Private Sub txtBeg_LostFocus()
   If Not bCancel Then
      If Val(txtEnd) > 0 _
             And Val(txtBeg) > 0 Then
         FillGrid
      End If
   End If
End Sub

Private Sub txtEnd_LostFocus()
   If Not bCancel Then
      If Val(txtBeg) > 0 And _
             Val(txtEnd) > 0 Then
         FillGrid
      End If
   End If
End Sub

Private Function IsCheckSelected() As Byte
   Dim i As Integer
   With Grid1
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 0
         If .CellPicture = imgInc Then
            IsCheckSelected = True
            Exit For
         End If
      Next
   End With
   
End Function

Private Sub ReprintChecks()
   
   Dim i As Integer
   Dim K As Integer
   
   Dim RdoInv As ADODB.Recordset
   Dim RdoChk As ADODB.Recordset
   'Dim RdoSum As ADODB.Recordset
   
   Dim sCheck() As String
   Dim sCheckAcct() As String
   
   Dim sVendor As String
   Dim sInvoice As String
   Dim sDate As String
   Dim sAccount As String
   
   Dim lCheck As Long
   
   Dim cDiscount As Currency
   Dim cPay As Currency
   Dim cDue As Currency
   
   On Error GoTo DiaErr1
   
   ' Load up checks to reprint
   With Grid1
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 0
         If .CellPicture = imgInc Then
            .Col = 1
            ReDim Preserve sCheck(K)
            ReDim Preserve sCheckAcct(K)
            sCheck(K) = Trim(.Text)
            
            .Col = 5
            sCheckAcct(K) = Trim(.Text)
            
            K = K + 1
         End If
      Next
   End With
   
   ' Recreate check setup then Q-up the
   ' check printing transaction
   sSql = "SELECT CHKNUMBER,CHKAMOUNT,CHKPRINTDATE,CHKACCT,DCDEBIT,DCCREDIT," _
          & "DCREF,VIVENDOR,VINO,VIDUE,VIPAY FROM ChksTable INNER JOIN " _
          & "JritTable ON ChksTable.CHKNUMBER = JritTable.DCCHECKNO " _
          & " AND ChksTable.CHKACCT = JritTable.DCCHKACCT INNER JOIN " _
          & "VihdTable ON JritTable.DCVENDORINV = VihdTable.VINO AND " _
          & "JritTable.DCVENDOR = VihdTable.VIVENDOR "
   For i = 0 To UBound(sCheck)
   '   sSql = sSql & sCheck(i) & "','"
      If (i = 0) Then sSql = sSql & " Where "
      
      sSql = sSql & "(CHKNUMBER = '" & sCheck(i) & "' AND CHKACCT = '" & sCheckAcct(i) & "')"
      
      If (i <> UBound(sCheck)) Then sSql = sSql & " OR "
   Next
   
   sSql = sSql & " ORDER BY CONVERT(INT, CHKNUMBER, 101),DCTRAN,DCREF "
   
   Debug.Print sSql
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk)
   If bSqlRows Then
      With RdoChk
         sInvoice = Trim(!VINO)
         lCheck = CLng(!CHKNUMBER)
         sDate = Format(!CHKPRINTDATE, DATEMASK)
         sVendor = Trim(!VIVENDOR)
         sAccount = "" & Trim(!CHKACCT)
         i = 1
         
         On Error Resume Next
         Err.Clear
         clsADOCon.ADOErrNum = 0
         While Not .EOF
            If sInvoice <> Trim(!VINO) Or _
                                sInvoice = Trim(!VINO) And sVendor <> Trim(!VIVENDOR) Then
               
               sSql = "INSERT INTO ChseTable(CHKNUM,CHKVND,CHKINV,CHKPAMT," _
                      & "CHKAMT,CHKDIS,CHKDATE,CHKBY,CHKREPRINTNO,CHKACCT)" _
                      & " VALUES( " _
                      & "'" & CStr(i) & "'," _
                      & "'" & sVendor & "'," _
                      & "'" & sInvoice & "'," _
                      & cPay - cDiscount & "," _
                      & cPay & "," _
                      & cDiscount & "," _
                      & "'" & sDate & "'," _
                      & "'" & sInitials & "'," _
                      & lCheck & ",'" & sAccount & "')"
               clsADOCon.ExecuteSQL sSql
               
               ' Reset
               If lCheck <> CLng(!CHKNUMBER) Then
                  i = i + 1
                  lCheck = CLng(!CHKNUMBER)
               End If
               cPay = 0
               cDiscount = 0
               sInvoice = Trim(!VINO)
               sDate = Format(!CHKPRINTDATE, DATEMASK)
               sVendor = Trim(!VIVENDOR)
               sAccount = "" & Trim(!CHKACCT)
            End If
            
            If !DCREF = 3 Then
               ' Grab the discount.  Its ALWAYS DCREF 3
               cDiscount = !DCCREDIT
            Else
               cPay = cPay + !DCDEBIT
               If !VIDUE < 0 Then
                  ' Credit Memo
                  ' Note ABS is needed to support vendor credits
                  ' in the VITYPE = PO format.
                  cPay = Abs(cPay) * -1
               End If
            End If
            
            .MoveNext
         Wend
         
         sSql = "INSERT INTO ChseTable(CHKNUM,CHKVND,CHKINV,CHKPAMT," _
                & "CHKAMT,CHKDIS,CHKDATE,CHKBY,CHKREPRINTNO,CHKACCT)" _
                & " VALUES( " _
                & "'" & CStr(i) & "'," _
                & "'" & sVendor & "'," _
                & "'" & sInvoice & "'," _
                & cPay - cDiscount & "," _
                & cPay & "," _
                & cDiscount & "," _
                & "'" & sDate & "'," _
                & "'" & sInitials & "'," _
                & lCheck & ",'" & sAccount & "')"
         clsADOCon.ExecuteSQL sSql
      End With
   End If
   Set RdoChk = Nothing
   
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      
      ' Que up check printing diaARf03a
      diaAPf03a.bReprint = True
      diaAPf03a.LoadCheckArray sCheck()
      Unload Me
      diaAPf03a.Show
   Else
      clsADOCon.RollbackTrans
      sMsg = "Cannot Reprint Checks" _
             & vbCrLf & "Transaction Canceled By User."
      MsgBox sMsg, vbExclamation, Caption
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "reprintchecks"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub
