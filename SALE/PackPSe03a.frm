VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form PackPSe03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ship Packaged Goods"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PackPSe03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Tag             =   "4"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CheckBox optNot 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      ToolTipText     =   "Shows A List Of Shipped Items To Mark Not Shipped"
      Top             =   1560
      Width           =   715
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4980
      TabIndex        =   15
      ToolTipText     =   "Cancel This List"
      Top             =   2160
      Width           =   875
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5880
      TabIndex        =   14
      ToolTipText     =   "Update Selected Pack Slips And Apply Changes"
      Top             =   2160
      Width           =   875
   End
   Begin VB.Frame z2 
      Height          =   40
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   6615
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "&Select"
      Height          =   315
      Left            =   5880
      TabIndex        =   5
      ToolTipText     =   "Retrieves List (Up To 150 Rows)"
      Top             =   1560
      Width           =   875
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Customers With Packing Slips"
      Top             =   360
      Width           =   1555
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   3720
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5880
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6840
      Top             =   3840
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5355
      FormDesignWidth =   6855
   End
   Begin MSFlexGridLib.MSFlexGrid grd1 
      Height          =   2535
      Left            =   240
      TabIndex        =   7
      ToolTipText     =   "Select The Item To Ship"
      Top             =   2520
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4471
      _Version        =   393216
      Rows            =   10
      Cols            =   6
      FixedCols       =   0
      Enabled         =   -1  'True
      HighLight       =   2
      GridLinesFixed  =   1
      ScrollBars      =   2
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Shipped"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   19
      Top             =   1560
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mark Not Shipped"
      Height          =   285
      Index           =   5
      Left            =   2760
      TabIndex        =   18
      ToolTipText     =   "Shows A List Of Shipped Items To Mark Not Shipped"
      Top             =   1560
      Width           =   1905
   End
   Begin VB.Label Selected 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   1440
      TabIndex        =   17
      ToolTipText     =   "Row Count"
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   16
      Top             =   2160
      Width           =   915
   End
   Begin VB.Label lblNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   11
      ToolTipText     =   "Select A Customer Or Leave Blank"
      Top             =   360
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   2
      Left            =   3840
      TabIndex        =   10
      Top             =   360
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Printed Dates"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   1080
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      Height          =   285
      Index           =   3
      Left            =   2760
      TabIndex        =   8
      Top             =   1080
      Width           =   915
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   0
      Picture         =   "PackPSe03a.frx":07AE
      Top             =   5160
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   360
      Picture         =   "PackPSe03a.frx":0B38
      Top             =   5160
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "PackPSe03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'New 10/2/03
'8/5/04 Added SoitTable.ITPSSHIPPED
'4/26/05 Decreased array size and selection size
'        Added GetOptions (save date)
Option Explicit
Dim bOnLoad As Byte
Dim iGridRows As Integer

Dim sPackSlip() As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub GetPsDates()
   Dim RdoGdt As ADODB.Recordset
   On Error GoTo DiaErr1
   If Trim(txtBeg) = "" Then
      sSql = "SELECT MIN(PSPRINTED) FROM PshdTable WHERE PSPRINTED " _
             & "IS NOT NULL"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoGdt, ES_FORWARD)
      If bSqlRows Then
         txtBeg = Format(RdoGdt.Fields(0), "mm/dd/yyyy")
      Else
         txtBeg = Format(ES_SYSDATE, "mm/dd/yyyy")
      End If
   End If
   txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   Exit Sub
   
DiaErr1:
   txtBeg = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   
End Sub

Private Sub cmbCst_Click()
   GetCustomer
   
End Sub


Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If cmbCst = "" Then cmbCst = "ALL"
   GetCustomer
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdEnd_Click()
   Dim bResponse As Byte
   bResponse = MsgBox("Cancel Without Saving?", _
               ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      grd1.Rows = 1
      cmbCst.Enabled = True
      txtBeg.Enabled = True
      txtEnd.Enabled = True
      cmdSel.Enabled = True
      txtDte.Enabled = True
      optNot.Enabled = True
      cmdEnd.Enabled = False
      cmdUpd.Enabled = False
      Selected = 0
      grd1.Enabled = False
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2203
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdSel_Click()
   GetPackingSlips
   
End Sub

Private Sub cmdUpd_Click()
   UpdateList
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   ReDim sPackSlip(151, 3)
   FormLoad Me
   FormatControls
   With grd1
      .Rows = 1
      .Row = 0
      .Col = 0
      .Text = "Ship"
      .Col = 1
      .Text = "Pack Slip"
      .Col = 2
      .Text = "Customer"
      .Col = 3
      .Text = "SO Number"
      .Col = 4
      .Text = "PO Number"
      .Col = 5
      .Text = "Printed"
      .ColWidth(0) = 500
      .ColWidth(1) = 1000
      .ColWidth(2) = 1200
      .ColWidth(3) = 1000
      .ColWidth(4) = 1800
      .ColWidth(5) = 900
   End With
   GetOptions
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set PackPSe03a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   GetPsDates
   txtDte = Format(ES_SYSDATE, "mm/dd/yyyy")
   cmdEnd.ToolTipText = "Cancel Work Not Updated And Return To Selection"
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_FillPSNotPrinted"
   LoadComboBox cmbCst, 2
   GetCustomer
   cmbCst = "ALL"
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetPackingSlips()
   Dim RdoPsl As ADODB.Recordset
   Dim bShip As Byte
   Dim iRows As Integer
   Dim sCust As String
   Dim sPsNo As String
   Dim sPoNo As String
   Dim sSoNo As String
   
   Dim sBegDate As String
   Dim sEndDate As String
   If txtBeg <> "ALL" Then
      sBegDate = txtBeg
   Else
      sBegDate = "12/31/1995"
   End If
   
   If txtEnd <> "ALL" Then
      sEndDate = txtEnd
   Else
      sEndDate = "12/31/2025"
   End If
   
   If cmbCst <> "ALL" Then sCust = Compress(cmbCst)
   bShip = optNot.Value
   iGridRows = 1
   grd1.Rows = 1
   Erase sPackSlip
   ReDim sPackSlip(151, 3)
   On Error GoTo DiaErr1
   sSql = "SELECT PSNUMBER,PSCUST,PSPRINTED,PSSHIPPED,CUREF,CUNICKNAME " _
          & "FROM PshdTable,CustTable WHERE (PSCUST=CUREF AND " _
          & "PSCUST LIKE '" & sCust & "%' AND PSPRINTED BETWEEN  '" _
          & sBegDate & " 00:00' AND '" & sEndDate & " 23:59' AND PSSHIPPED=" & bShip & ")" _
          & " ORDER BY PSNUMBER"
          
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPsl, ES_STATIC)

   'Debug.Print "Rows = " & RdoPsl.RecordCount
   ReDim sPackSlip(0 To RdoPsl.RecordCount + 1, 3)
   
   
   If bSqlRows Then
      With RdoPsl
         MouseCursor 13
         Do Until .EOF
            GetSalesOrder !PsNumber, sSoNo, sPoNo
            If iGridRows > 300 Then Exit Do
            iRows = iRows + 1
'            If iRows > 150 Then Exit Do
            grd1.Rows = grd1.Rows + 1
            grd1.Row = grd1.Rows - 1
            sPackSlip(grd1.Row, 0) = "" & Trim(!PsNumber)
            sPackSlip(grd1.Row, 1) = ""
            grd1.Col = 0
            Set grd1.CellPicture = Chkno.Picture
            grd1.Col = 1
            grd1.Text = "" & Trim(!PsNumber)
            grd1.Col = 2
            grd1.Text = "" & Trim(!CUNICKNAME)
            grd1.Col = 3
            grd1.Text = sSoNo
            grd1.Col = 4
            grd1.Text = sPoNo
            grd1.Col = 5
            grd1.Text = "" & Format(!PSPRINTED, "mm/dd/yy")
            .MoveNext
         Loop
         ClearResultSet RdoPsl
      End With
      MouseCursor 0
   Else
      MsgBox "No Matching Pack Slips Items Found.", _
         vbInformation, Caption
   End If
   iGridRows = grd1.Rows
   Selected = iGridRows - 1
   If iGridRows > 1 Then
      cmbCst.Enabled = False
      txtBeg.Enabled = False
      txtEnd.Enabled = False
      txtDte.Enabled = False
      optNot.Enabled = False
      cmdSel.Enabled = False
      cmdUpd.Enabled = True
      cmdEnd.Enabled = True
      grd1.Enabled = True
   End If
   Set RdoPsl = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getpackingslips"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiSale", "shpkg", Trim(sOptions))
   If Len(sOptions) > 0 Then txtBeg = sOptions Else txtBeg = ""
   
End Sub

Private Sub GetSalesOrder(PackSlip As String, SalesOrder As String, _
                          PurchaseOrder As String)
   
   Dim RdoSon As ADODB.Recordset
   sSql = "SELECT DISTINCT PIPACKSLIP,PISONUMBER,SONUMBER,SOTYPE," _
          & "SOPO FROM PsitTable,SohdTable WHERE (PIPACKSLIP='" _
          & PackSlip & "' AND PISONUMBER=SONUMBER)"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSon, ES_FORWARD)
   If bSqlRows Then
      With RdoSon
         SalesOrder = "" & Trim(!SOTYPE) & Format(!SoNumber, SO_NUM_FORMAT)
         PurchaseOrder = "" & Trim(!SOPO)
         ClearResultSet RdoSon
      End With
   Else
      PurchaseOrder = ""
   End If
   Set RdoSon = Nothing
   Exit Sub
   
DiaErr1:
   SalesOrder = ""
   PurchaseOrder = ""
End Sub

Private Sub grd1_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      grd1.Col = 0
      If grd1.Row = 0 Then grd1.Row = 1
      If grd1.CellPicture = Chkyes.Picture Then
         Set grd1.CellPicture = Chkno.Picture
         sPackSlip(grd1.Row, 1) = ""
      Else
         Set grd1.CellPicture = Chkyes.Picture
         sPackSlip(grd1.Row, 1) = "X"
      End If
   End If
   
End Sub


Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   grd1.Col = 0
   If grd1.Row = 0 Then grd1.Row = 1
   If grd1.CellPicture = Chkyes.Picture Then
      Set grd1.CellPicture = Chkno.Picture
      sPackSlip(grd1.Row, 1) = ""
   Else
      Set grd1.CellPicture = Chkyes.Picture
      sPackSlip(grd1.Row, 1) = "X"
   End If
   
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDateEx(txtBeg)
   
End Sub


Private Sub txtDte_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtDte_LostFocus()
   txtDte = CheckDateEx(txtDte)
   
End Sub


Private Sub txtend_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If Trim(txtEnd) <> "ALL" Then txtEnd = CheckDateEx(txtEnd)
   
End Sub



Private Sub UpdateList()
   Dim bByte As Byte
   Dim iRow As Integer
   Dim sMsg As String
   For iRow = 1 To iGridRows
      If sPackSlip(iRow, 1) = "X" Then bByte = 1
   Next
   If bByte = 0 Then
      MsgBox "No Items Have Been Selected For Update To Update.", _
         vbInformation, Caption
      Exit Sub
   End If
   On Error Resume Next
   sMsg = "Do You Wish To Continue And Mark Selected Items As "
   If optNot.Value = vbUnchecked Then sMsg = sMsg & "Shipped?" _
                     Else sMsg = sMsg & "Not Shipped?"
   bByte = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bByte = vbYes Then
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      If optNot.Value = vbUnchecked Then
         'Mark Shipped
         For iRow = 1 To iGridRows
            If sPackSlip(iRow, 1) = "X" Then
               sSql = "UPDATE PshdTable SET PSSHIPPEDDATE='" _
                      & txtDte & "',PSSHIPPED=1 WHERE PSNUMBER='" _
                      & sPackSlip(iRow, 0) & "'"
               clsADOCon.ExecuteSql sSql 'rdExecDirect
               
'               sSql = "UPDATE SoitTable SET ITACTUAL='" & Format(Now, "mm/dd/yyyy") & "'," _
'                      & "ITPSSHIPPED=1 WHERE (ITPSNUMBER='" & sPackSlip(iRow, 0) & "' AND ITCANCELED=0)"
               sSql = "UPDATE SoitTable SET ITACTUAL='" & Format(txtDte, "mm/dd/yyyy") & "'," _
                      & "ITPSSHIPPED=1 WHERE (ITPSNUMBER='" & sPackSlip(iRow, 0) & "' AND ITCANCELED=0)"
               clsADOCon.ExecuteSql sSql ' rdExecDirect
            End If
         Next
         sMsg = " Shipped"
      Else
         For iRow = 1 To iGridRows
            If sPackSlip(iRow, 1) = "X" Then
               sSql = "UPDATE PshdTable SET PSSHIPPEDDATE=Null," _
                      & "PSSHIPPED=0 WHERE PSNUMBER='" _
                      & sPackSlip(iRow, 0) & "'"
              clsADOCon.ExecuteSql sSql 'rdExecDirect
               
               sSql = "UPDATE SoitTable SET ITPSSHIPPED=0 " _
                      & "WHERE ITPSNUMBER='" & sPackSlip(iRow, 0) & "'"
               clsADOCon.ExecuteSql sSql 'rdExecDirect
            End If
         Next
         sMsg = " Not Shipped"
      End If
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         SysMsg "Packing Slips" & sMsg, True
      Else
         clsADOCon.RollbackTrans
         MsgBox "Could Not Mark Pack Slips Shipped.", _
            vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   Erase sPackSlip
   ReDim sPackSlip(151, 3)

   grd1.Rows = 1
   cmbCst.Enabled = True
   txtBeg.Enabled = True
   txtEnd.Enabled = True
   cmdSel.Enabled = True
   txtDte.Enabled = True
   optNot.Enabled = True
   cmdEnd.Enabled = False
   cmdUpd.Enabled = False
   Selected = 0
   grd1.Enabled = False
   
End Sub

Private Sub GetCustomer()
   Dim RdoCst As ADODB.Recordset
   On Error GoTo DiaErr1
   If cmbCst = "" Then cmbCst = "ALL"
   If cmbCst = "ALL" Then
      lblNme = "All Customers Selected."
   Else
      sSql = "Qry_GetCustomerBasics '" & Compress(cmbCst) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCst, ES_FORWARD)
      If bSqlRows Then
         With RdoCst
            lblNme = "" & Trim(.Fields(1))
            ClearResultSet RdoCst
         End With
      End If
   End If
   Set RdoCst = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getcustomer"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub SaveOptions()
   On Error Resume Next
   SaveSetting "Esi2000", "EsiSale", "shpkg", Trim(txtBeg)
   
End Sub
