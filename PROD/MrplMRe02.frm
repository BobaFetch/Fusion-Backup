VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form MrplMRe02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create PO's from MRP Exceptions By Part(s)"
   ClientHeight    =   7890
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   12780
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7890
   ScaleWidth      =   12780
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "MrplMRe02.frx":0000
      Height          =   315
      Left            =   5160
      Picture         =   "MrplMRe02.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   720
      Width           =   350
   End
   Begin VB.TextBox txtPrt 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Tag             =   "3"
      Top             =   720
      Width           =   3495
   End
   Begin VB.CheckBox OptAutoPO 
      Caption         =   "Option AutoPO"
      Height          =   195
      Left            =   7320
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CheckBox OptAddComm 
      Caption         =   "Option to Add Comment"
      Height          =   195
      Left            =   240
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2280
      Width           =   2775
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   8280
      Sorted          =   -1  'True
      TabIndex        =   23
      Tag             =   "3"
      ToolTipText     =   "Select Product Code From List"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7800
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdPO 
      Caption         =   "&Create PO from MRP"
      Height          =   435
      Left            =   10800
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Create PO from MRP"
      Top             =   2520
      Width           =   1755
   End
   Begin VB.PictureBox picUnchecked 
      Height          =   285
      Left            =   8280
      Picture         =   "MrplMRe02.frx":0684
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picChecked 
      Height          =   285
      Left            =   8280
      Picture         =   "MrplMRe02.frx":09C6
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   19
      Top             =   600
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.CommandButton cmdMRP 
      Caption         =   "&Get MRP exception (PO)"
      Height          =   435
      Left            =   5280
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Get MRP exception"
      Top             =   2040
      Width           =   2115
   End
   Begin VB.ComboBox cmbPart 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   720
      Width           =   3495
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "MrplMRe02.frx":0D08
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Tag             =   "4"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4200
      TabIndex        =   6
      Tag             =   "4"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Select Product Code From List"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   4200
      Sorted          =   -1  'True
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Select Product Class From List"
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   11280
      Top             =   6960
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   7890
      FormDesignWidth =   12780
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   5055
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   8916
      _Version        =   393216
      Rows            =   3
      Cols            =   8
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   315
      FocusRect       =   2
      ScrollBars      =   2
      MergeCells      =   1
      AllowUserResizing=   3
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   6960
      Picture         =   "MrplMRe02.frx":14B6
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   6960
      Picture         =   "MrplMRe02.frx":1840
      Stretch         =   -1  'True
      Top             =   480
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   13
      Left            =   5640
      TabIndex        =   15
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   11
      Left            =   5640
      TabIndex        =   14
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   10
      Left            =   5640
      TabIndex        =   13
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   9
      Left            =   2880
      TabIndex        =   12
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   11
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Classes"
      Height          =   285
      Index           =   0
      Left            =   2880
      TabIndex        =   10
      Top             =   1200
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Codes"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   1425
   End
   Begin VB.Label p 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Numbers"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   765
      Width           =   1425
   End
End
Attribute VB_Name = "MrplMRe02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'5/19/06 Revised report and selections. Removed extra report.
Option Explicit
Dim bOnLoad As Byte

'Passed document stuff
Dim iDocEco As Integer
Dim strDocName As String
Dim strDocClass As String
Dim strDocSheet As String
Dim strDocDesc As String
Dim strDocAdcn As String
Dim sListRef As String
Dim strListRev As String
Dim UsingMouse As Boolean
Dim bGenMRP As Boolean
Dim strLastVendor As String
Dim strLastPoNum As String
Dim bView As Integer

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'Least to greatest dates 10/12/01

Private Sub GetMRPDates()
   Dim RdoDte As ADODB.Recordset
    sSql = "SELECT MIN(MRP_PARTDATERQD) FROM MrplTable WHERE " _
           & "MRP_TYPE>" & MRPTYPE_BeginningBalance
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDte, ES_FORWARD)
   If bSqlRows Then
      With RdoDte
         If Not IsNull(.Fields(0)) Then
            txtBeg = Format(.Fields(0), "mm/dd/yyyy")
         Else
            txtBeg = Format(ES_SYSDATE, "mm/dd/yyyy")
         End If
         ClearResultSet RdoDte
      End With
   End If
   txtBeg.ToolTipText = "Earliest Date By Default"
   
    sSql = "SELECT MAX(MRP_PARTDATERQD) FROM MrplTable WHERE " _
           & "MRP_TYPE>" & MRPTYPE_BeginningBalance
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDte, ES_FORWARD)
   If bSqlRows Then
      With RdoDte
         If Not IsNull(.Fields(0)) Then
            txtEnd = Format(.Fields(0), "mm/dd/yyyy")
         Else
            txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
         End If
         ClearResultSet RdoDte
      End With
   End If
   txtEnd.ToolTipText = "Latest Date By Default"
   
End Sub


Private Sub cmbCde_LostFocus()
   cmbCde = CheckLen(cmbCde, 6)
   If cmbCde = "" Then cmbCde = "ALL"
   
End Sub


Private Sub cmbCls_LostFocus()
   cmbCls = CheckLen(cmbCls, 6)
   If cmbCls = "" Then cmbCls = "ALL"
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdFnd_Click()
   If txtPrt.Visible Then
      cmbPart = txtPrt
      ViewParts.lblControl = "TXTPRT"
   Else
      ViewParts.lblControl = "CMBPRT"
   End If
   ViewParts.txtPrt = cmbPart
   ViewParts.Show
   bView = 0
End Sub

Private Sub Text1_GotFocus()
   Grd.Text = Text1.Text
   If Grd.Col >= Grd.Cols Then Grd.Col = 1
   ChangeCellText
End Sub

Private Sub cmbVnd_GotFocus()
   Grd.Text = cmbVnd.Text
   If Grd.Col >= Grd.Cols Then Grd.Col = 1
   ChangeCellCombo
End Sub


Private Sub Grd_EnterCell()  ' Assign cell value to the textbox
   If (bGenMRP = True) Then
      Text1.Text = Grd.Text
      cmbVnd.Text = Grd.Text
   End If
End Sub

Private Sub Grd_LeaveCell()
   ' Assign textbox value to Grd
   If (bGenMRP = True) And (Text1.Visible = True) Then
      Grd.Text = Text1.Text
      Text1.Text = ""
      Text1.Visible = False
   End If
   
   If (bGenMRP = True) And (cmbVnd.Visible = True) Then
      Grd.Text = cmbVnd.Text
      cmbVnd.Visible = False
   End If

End Sub

Private Sub cmbVnd_LostFocus()

   If (cmbVnd.Visible = True) Then
      Grd.Text = cmbVnd.Text
      cmbVnd.Text = ""
      cmbVnd.Visible = False
   End If
   
   If UsingMouse = True Then
      UsingMouse = False
      Exit Sub
   End If
End Sub

Private Sub Text1_LostFocus()

   If (Text1.Visible = True) Then
      Grd.Text = Text1.Text
      Text1.Text = ""
      Text1.Visible = False
   End If
   
   If UsingMouse = True Then
      UsingMouse = False
      Exit Sub
   End If
End Sub

Public Sub ChangeCellText() ' Move Textbox to active cell.
   Text1.Move Grd.Left + Grd.CellLeft, _
   Grd.Top + Grd.CellTop, _
   Grd.CellWidth, Grd.CellHeight
   'Text1.SetFocus
   'Text1.ZOrder 0
End Sub

Public Sub ChangeCellCombo() ' Move Textbox to active cell.
   cmbVnd.Move Grd.Left + Grd.CellLeft, _
   Grd.Top + Grd.CellTop, _
   Grd.CellWidth
   'Text1.SetFocus
   'Text1.ZOrder 0
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 907
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub FillCombos()
    On Error Resume Next
    sSql = "SELECT DISTINCT PARTREF,PARTNUM " _
        & "FROM PartTable  " _
        & "INNER JOIN MrplTable ON MrplTable.MRP_PARTREF=PartTable.PARTREF " _
        & " WHERE PAINACTIVE = 0 AND PAOBSOLETE = 0 " _
        & "ORDER BY PARTREF"
    LoadComboBox cmbPart, 0
    cmbPart = "ALL"
    If Trim(cmbPart) = "" Then cmbPart = "ALL"
End Sub


Private Sub cmdPO_Click()
   ' First create PO for individual items
   
   Err.Clear
   clsADOCon.ADOErrNum = 0
   
   CreateNewPO ("SINGLE_POITEM")
   If clsADOCon.ADOErrNum <> 0 Then
      Exit Sub
   End If
   
   ' First create PO for individual items
   CreateNewPO ("GRP_POITEM")
   
   If clsADOCon.ADOErrNum = 0 Then
      MsgBox "Successfully created PO's", _
            vbInformation, Caption
      
      ' Open the Revise Purchase Order form
'      PurcPRe02a.Show
'      PurcPRe02a.OptAutoPO = vbChecked
'      PurcPRe02a.cmbPon = strLastPoNum
'      PurcPRe02a.cmbVnd = strLastVendor
'      PurcPRe02a.SetFocus
'      PurcPRe02a.cmbPon.SetFocus
   End If
   
End Sub

Private Sub cmdMRP_Click()

    Dim sParts As String
    Dim sCode As String
    Dim sClass As String
    Dim sBDate As String
    Dim sEDate As String
    Dim sBegDate As String
    Dim sEndDate As String
    Dim sUnitPrice As String
    Grd.Clear
    GrdAddHeader
    
    GetMRPCreateDates sBegDate, sEndDate
    
    If Trim(txtBeg) = "" Then txtBeg = "ALL"
    If Trim(txtEnd) = "" Then txtEnd = "ALL"
    If Not IsDate(txtBeg) Then
       sBDate = "1/1/2000"
    Else
       sBDate = Format(txtBeg, "mm/dd/yyyy")
    End If
    If Not IsDate(txtEnd) Then
       sEDate = "12/31/2024"
    Else
       sEDate = Format(txtEnd, "mm/dd/yyyy")
    End If
    
    If Trim(cmbPart) = "" Then cmbPart = "ALL"
    If Trim(cmbCde) = "" Then cmbCde = "ALL"
    If Trim(cmbCls) = "" Then cmbCls = "ALL"
    If Trim(cmbPart) = "ALL" Then sParts = "" Else sParts = Compress(cmbPart)
    If Trim(cmbCde) = "ALL" Then sCode = "" Else sCode = Compress(cmbCde)
    If Trim(cmbCls) = "ALL" Then sClass = "" Else sClass = Compress(cmbCls)
   
   
   Dim RdoMrpEx As ADODB.Recordset
   
   sSql = "SELECT MRP_PARTREF, MRP_PARTNUM,MRP_PARTQTYRQD," & vbCrLf _
            & "CONVERT(varchar(12), MRP_PARTDATERQD,101) MRP_PARTDATERQD," & vbCrLf _
            & "CONVERT(varchar(12), MRP_ACTIONDATE, 101) MRP_ACTIONDATE," & vbCrLf _
            & "(select TOP 1 PIVENDOR from poitTable " & vbCrLf _
            & "WHERE PIPART = MRP_PARTREF AND PIADATE IS NOT NULL ORDER BY PIADATE DESC) as POVENDOR1," & vbCrLf _
            & "(select TOP 1 PIAMT from poitTable" & vbCrLf _
            & "WHERE PIPART = MRP_PARTREF AND PIADATE IS NOT NULL ORDER BY PIADATE DESC) as UNITPRICE " & vbCrLf _
      & "  FROM MrplTable, PartTable " & vbCrLf _
      & " WHERE MRP_PARTREF = PartRef" & vbCrLf _
            & "AND MrplTable.MRP_PARTREF LIKE '" & sParts & "%'" & vbCrLf _
            & "AND MrplTable.MRP_PARTPRODCODE LIKE '" & sCode & "%'" & vbCrLf _
             & "AND MrplTable.MRP_PARTCLASS LIKE '" & sClass & "%'" & vbCrLf _
            & "AND MrplTable.MRP_PARTDATERQD BETWEEN '" & sBDate & "' AND '" & sEDate & "'" & vbCrLf _
            & "AND MrplTable.MRP_TYPE IN (6, 5)" & vbCrLf _
            & "AND PartTable.PAMAKEBUY ='B'" & vbCrLf _
            & " order by MRP_PARTREF"

   Debug.Print sSql
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoMrpEx)
   If bSqlRows Then
      With RdoMrpEx
         Do Until .EOF
            Grd.Rows = Grd.Rows + 1
            Grd.row = Grd.Rows - 1
            
            Grd.Col = 0
            Set Grd.CellPicture = Chkno.Picture
            Grd.Col = 1
            Grd.Text = Trim(!MRP_PARTNUM)
            Grd.Col = 2
            Grd.Text = Trim(!MRP_PARTQTYRQD)
            Grd.Col = 3
            sUnitPrice = IIf(IsNull(!UnitPrice), "0.0000", !UnitPrice)
            Grd.Text = Format(Val(sUnitPrice), ES_PurchasedDataFormat)
            Grd.Col = 4
            Grd.Text = Trim(!MRP_PARTDATERQD)
            Grd.Col = 5
            Grd.Text = Trim(!MRP_ACTIONDATE)
            Grd.Col = 6
            Grd.Text = Trim(IIf(IsNull(!POVENDOR1), "", !POVENDOR1))
            .MoveNext
         Loop
      End With
   End If
   
   Set RdoMrpEx = Nothing
   bGenMRP = True
   Exit Sub
   
End Sub

Private Sub grd_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      Dim iCurCol As Integer
      iCurCol = Grd.Col
      If Grd.row >= 1 Then
         If Grd.row = 0 Then Grd.row = 1
               
         If (Grd.Col = 0) Then
            If Grd.CellPicture = Chkyes.Picture Then
               Set Grd.CellPicture = Chkno.Picture
            Else
               Set Grd.CellPicture = Chkyes.Picture
            End If
         ElseIf (Grd.Col = 7) Then
            
            If Grd.CellPicture = picChecked.Picture Then
               Set Grd.CellPicture = picUnchecked.Picture
            Else
               Set Grd.CellPicture = picChecked.Picture
            End If
         ElseIf ((Grd.Col = 2) Or (Grd.Col = 4) Or (Grd.Col = 3)) Then
            UsingMouse = True
            Grd.Text = Text1.Text
            Text1.Visible = True
            ChangeCellText
         ElseIf ((Grd.Col = 6)) Then
            UsingMouse = True
            Grd.Text = cmbVnd.Text
            cmbVnd.Visible = True
            ChangeCellCombo
         End If
      
      End If
   End If

End Sub

Private Sub Grd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   Dim iCurCol As Integer
   iCurCol = Grd.Col
   If Grd.row >= 1 Then
      If Grd.row = 0 Then Grd.row = 1
            
      If (Grd.Col = 0) Then
         If Grd.CellPicture = Chkyes.Picture Then
            Set Grd.CellPicture = Chkno.Picture
         Else
            Set Grd.CellPicture = Chkyes.Picture
         End If
      ElseIf (Grd.Col = 7) Then
         
         If Grd.CellPicture = picChecked.Picture Then
            Set Grd.CellPicture = picUnchecked.Picture
         Else
            Set Grd.CellPicture = picChecked.Picture
         End If
      ElseIf ((Grd.Col = 2) Or (Grd.Col = 4) Or (Grd.Col = 3)) Then
         UsingMouse = True
         Grd.Text = Text1.Text
         Text1.Visible = True
         ChangeCellText
      ElseIf ((Grd.Col = 6)) Then
         UsingMouse = True
         Grd.Text = cmbVnd.Text
         cmbVnd.Visible = True
         ChangeCellCombo
      End If
   
   End If
End Sub


Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      GetLastMrp
      GetMRPDates
      GetOptions
      cmbCde.AddItem "ALL"
      FillProductCodes
      If Trim(cmbCde) = "" Then cmbCde = cmbCde.List(0)
      cmbCls.AddItem "ALL"
      FillProductClasses
      If Trim(cmbCls) = "" Then cmbCls = cmbCls.List(0)
      
      Dim bPartSearch As Boolean
      
      bPartSearch = GetPartSearchOption
      SetPartSearchOption (bPartSearch)
      
      If (Not bPartSearch) Then FillCombos
      
      FillVendors
      'MM OptAutoPO = vbUnchecked
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   ' Add headers
   GrdAddHeader
   bGenMRP = False
   bOnLoad = 1
   
End Sub

Private Sub GrdAddHeader()
     
     With Grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 1
      .ColAlignment(6) = 1
   
      .Rows = 1
      .row = 0
      .Col = 0
      .Text = "Sel"
      .Col = 1
      .Text = "PartNumber"
      .Col = 2
      .Text = "Qty"
      .Col = 3
      .Text = "UnitPrice"
      .Col = 4
      .Text = "Required Date"
      .Col = 5
      .Text = "Action Date"
      .Col = 6
      .Text = "Vendor"
      .Col = 7
      .Text = "Group PO"
      
      .ColWidth(0) = 500
      .ColWidth(1) = 3200
      .ColWidth(2) = 1000
      .ColWidth(3) = 1000
      .ColWidth(4) = 1200
      .ColWidth(5) = 1200
      .ColWidth(6) = 1200
      .ColWidth(7) = 900
      
      
      .ScrollBars = flexScrollBarBoth
      .AllowUserResizing = flexResizeColumns
      
   End With

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' If the form is closed from Revise purchase,
   ' then don't call Unload as this will open the tab dialog.
   'If (OptAutoPO = vbUnchecked) Then FormUnload
   Set MrplMRe02 = Nothing
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
'   txtPrt = "ALL"
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   Dim sCode As String * 6
   Dim sClass As String * 4
   sCode = cmbCde
   sClass = cmbCls
   SaveSetting "Esi2000", "EsiProd", "MrplMRe02", sOptions
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "Prdmr02", sOptions)
   If Len(Trim(sOptions)) > 0 Then
      cmbCde = Mid$(sOptions, 1, 6)
      cmbCls = Mid$(sOptions, 7, 4)
   End If
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDateEx(txtBeg)
   
End Sub


Private Sub txtend_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If txtEnd <> "ALL" Then txtEnd = CheckDateEx(txtEnd)
   
End Sub

Private Sub txtPrt_LostFocus()
   If Trim(txtPrt) = "" Or Trim(txtPrt) = "ALL" Then txtPrt = "ALL"
   cmbPart = txtPrt
End Sub

Private Sub cmbPart_LostFocus()
    cmbPart = CheckLen(cmbPart, 30)
    If Trim(cmbPart) = "" Then cmbPart = "ALL"
End Sub


Private Sub CreateNewPO(strIterType As String)

   Dim iList As Integer
   Dim strPartNum As String
   Dim strQty As String
   Dim strPartRqd As String
   Dim strActDate As String
   Dim strVendor As String
   Dim bGrpChked As String
   Dim strPrevPart As String
   Dim strPrevPONum As String
   Dim strUnitPrice As String
   
    On Error GoTo DiaErr1
    MouseCursor 13
    Err.Clear
    
   strPrevPart = ""
   strPrevPONum = ""
   
    ' Go throught all the record in the grid and create MO
   For iList = 1 To Grd.Rows - 1
      Grd.Col = 0
      Grd.row = iList
      ' Only if the part is checked
      If Grd.CellPicture = Chkyes.Picture Then
            
         Grd.Col = 1
         strPartNum = Grd.Text
         Grd.Col = 2
         strQty = Grd.Text
         Grd.Col = 3
         strUnitPrice = Grd.Text
         Grd.Col = 4
         strPartRqd = Grd.Text
         Grd.Col = 5
         strActDate = Grd.Text
         Grd.Col = 6
         strVendor = Grd.Text
         
         ' Default va;ue
         bGrpChked = False
         Grd.Col = 7
         If (Grd.CellPicture = picChecked.Picture) Then
            bGrpChked = True
         End If
         
         
         Dim strPartRef As String
         Dim strPONum As String
         
         Dim clsPO  As ClassPO
         Set clsPO = New ClassPO
         ' Open the transcation
         
         clsADOCon.ADOErrNum = 0
         clsADOCon.BeginTrans
         
         If (strIterType = "GRP_POITEM") Then
            ' If grouped create one PO for all the item selected for a Partnumber
            If (bGrpChked = True) Then
               If (strPrevPart <> strPartNum) Then
                  strPONum = clsPO.AddNewPo(strVendor)
                  strPrevPart = strPartNum
                  strPrevPONum = strPONum
                  strLastVendor = strVendor
                  strLastPoNum = strPONum
               Else
                  strPONum = strPrevPONum
               End If
            Else
               ' The Group is not selected
               ' It could be indivdual one.
               strPONum = ""
            End If
         Else
            ' First create PO for each item selected.
            If (bGrpChked = False) Then
               strPONum = clsPO.AddNewPo(strVendor)
               
'               If (strPoNum <> "") Then
'                  MrplMRe02b.txtPONumber = strPoNum
'                  MrplMRe02b.Show vbModal
'               End If
               
            Else
               strPONum = ""
            End If
         End If
         
         If (strPONum <> "") Then
            clsPO.AddPOItem strPONum, Compress(strPartNum), strQty, strUnitPrice, _
                     strPartRqd, strActDate, strVendor
            'AddPOItemComment strPoNum, Compress(strPartNum)
         End If
               
         If (clsADOCon.ADOErrNum = 0) And (OptAddComm = vbChecked) _
               And (strPONum <> "") Then
            
            Dim iTotItems As Integer
            clsPO.GetLastItem strPONum, iTotItems
            
            MrplMRe02b.txtPONumber = strPONum
            MrplMRe02b.txtTotItems = iTotItems - 1
            MrplMRe02b.Show vbModal
            
         End If
         
         If clsADOCon.ADOErrNum <> 0 Then
            MsgBox "Couldn't Successfully Update..", _
               vbInformation, Caption
            
            clsADOCon.RollbackTrans
         Else
            clsADOCon.CommitTrans
         End If
         Set clsPO = Nothing
      
      End If
   Next
    
   MouseCursor 0
   Exit Sub

DiaErr1:
   clsADOCon.RollbackTrans
   sProcName = "cmdUpdate"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub

Private Function AddPOItemComment(strPONum As String, strPartRef As String)


End Function

Function SetPartSearchOption(bPartSearch As Boolean)
   
   If (bPartSearch = True) Then
      cmbPart.Visible = False
      txtPrt.Visible = True
      cmdFnd.Visible = True
   Else
      cmbPart.Visible = True
      txtPrt.Visible = False
      cmdFnd.Visible = False
   End If
End Function

