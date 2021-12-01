VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PadmPRf03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Product Codes or Classes"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10320
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Change Selected Parts Field"
      Height          =   975
      Left            =   360
      TabIndex        =   17
      Top             =   2040
      Width           =   7575
      Begin VB.ComboBox cmbChangeTo 
         Height          =   315
         Left            =   2880
         TabIndex        =   23
         Top             =   360
         Width           =   3135
      End
      Begin VB.OptionButton optPartField 
         Caption         =   "Part Class"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   21
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optPartField 
         Caption         =   "Product Code"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdUpdateParts 
         Caption         =   "Change"
         Height          =   375
         Left            =   6240
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Change to"
         Height          =   255
         Left            =   1920
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCheckAll 
      Caption         =   "Check All"
      Height          =   375
      Left            =   1440
      TabIndex        =   15
      ToolTipText     =   "Check All Parts in Grid"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter Part Information to Search for Below"
      Height          =   1935
      Left            =   360
      TabIndex        =   9
      ToolTipText     =   "Search for Parts to Change by Entering fields below and hitting ""Search"""
      Top             =   120
      Width           =   7575
      Begin VB.ComboBox cmbPrt 
         Height          =   315
         Left            =   1560
         TabIndex        =   24
         Top             =   240
         Width           =   4935
      End
      Begin VB.CommandButton cmdFindPart 
         Height          =   315
         Left            =   6720
         Picture         =   "PadmPRf03a.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Find A Part Number"
         Top             =   240
         Visible         =   0   'False
         Width           =   350
      End
      Begin VB.TextBox txtPartDesc 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         ToolTipText     =   "Enter Partial/Full Part Description or Leave Blank"
         Top             =   600
         Width           =   4935
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   375
         Left            =   6240
         TabIndex        =   13
         ToolTipText     =   "Fill Grid with Parts Based on Search Criteria"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ComboBox cmbCde 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Tag             =   "3"
         ToolTipText     =   "Select Product Code"
         Top             =   960
         Width           =   1215
      End
      Begin VB.ComboBox cmbTyp 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Tag             =   "8"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtPrt 
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         ToolTipText     =   "Enter Partial/Full Part Number or Leave Blank"
         Top             =   240
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.Label Label2 
         Caption         =   "Part Description"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Part Number"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Code"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblPTDsc 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3000
         TabIndex        =   11
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Part Type"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdCheckNone 
      Caption         =   "Check None"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Uncheck All Parts in Grid"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update Part Type"
      Height          =   315
      Left            =   8040
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Enter Updated Time Card"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PadmPRf03a.frx":043A
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   435
      Left            =   9360
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   8280
      Top             =   1440
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   7005
      FormDesignWidth =   10320
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   3255
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Click To Select Or Scroll And Press Enter (Also Page Up And Page Down)"
      Top             =   3600
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   5741
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      FocusRect       =   2
      ScrollBars      =   2
      AllowUserResizing=   1
   End
   Begin VB.Image Image1 
      Height          =   210
      Left            =   8640
      Picture         =   "PadmPRf03a.frx":0BE8
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   8760
      Picture         =   "PadmPRf03a.frx":0F72
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   0
      Picture         =   "PadmPRf03a.frx":12FC
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "PadmPRf03a"
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

Dim bOnLoad As Byte
Dim bGoodCode As Byte
Dim bShowParts As Byte
Dim bGoodPart As Byte
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter

Dim sType(8, 2) As String

Dim strProdCode, strPartNo, strPartDesc, strPartType As String
Dim intSortColumn As Integer
Dim strSortOrder As String
Dim arrKeyValue() As String


Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FillChangeFieldCombo(Index As Integer)
    Dim RdoCde As ADODB.Recordset
    Dim intArrPos, intMax As Integer
    
    On Error GoTo fcfc1
    cmbChangeTo.Clear
    
    Select Case Index
    Case 0:
        sSql = "SELECT PCREF,PCCODE,PCDESC From PcodTable  WHERE (PCREF<>'BID' AND PCREF<>'TOOL') ORDER BY PCREF"
    Case 1:
        sSql = "SELECT CCREF,CCCODE,CCDESC From PclsTable ORDER BY CCREF"
    End Select
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoCde, ES_STATIC)
    If bSqlRows Then
        Erase arrKeyValue
       
        If RdoCde.RecordCount > 0 Then intMax = RdoCde.RecordCount Else intMax = 1
        ReDim arrKeyValue(1 To intMax)
        intArrPos = 1
        
      With RdoCde
         Do Until .EOF
            Select Case Index
            Case 0:
                cmbChangeTo.AddItem "" & Trim(!PCCODE) & " - " & "" & Trim(!PCDESC)
                arrKeyValue(intArrPos) = "" & Trim(!PCREF)
                cmbChangeTo.ItemData(cmbChangeTo.NewIndex) = intArrPos
                intArrPos = intArrPos + 1
'                cmbChangeTo.ItemData(cmbChangeTo.NewIndex) = "" & Trim(!PCREF)
            Case 1:
                cmbChangeTo.AddItem "" & Trim(!CCCODE) & " - " & "" & Trim(!CCDESC)
                arrKeyValue(intArrPos) = "" & Trim(!CCREF)
                cmbChangeTo.ItemData(cmbChangeTo.NewIndex) = intArrPos
                intArrPos = intArrPos + 1

'                cmbChangeTo.ItemData(cmbChangeTo.NewIndex) = "" & Trim(!CCREF)

            End Select
            .MoveNext
         Loop
         ClearResultSet RdoCde
      End With
      cmbChangeTo.ListIndex = 0
   End If
   Set RdoCde = Nothing
   Exit Sub
   
fcfc1:
   sProcName = "fillchangefieldcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
End Sub


Private Sub cmbCde_Click()
'   strPartNo = Compress(txtPartNumber)
'   strPartType = Compress(cmbTyp)
   bGoodCode = GetCode()
'   If (bGoodCode) Then
'      strProdCode = Compress(cmbCde)
'      FillGrid strProdCode, strPartNo, strPartType
'   End If
'    intSortColumn = 1
'    FillGrid
    cmdSearch.Enabled = True
End Sub


Private Sub cmbCde_LostFocus()
   cmbCde = CheckLen(cmbCde, 6)
   If Len(cmbCde) Then bGoodCode = GetCode()
'   intSortColumn = 1
'   FillGrid
End Sub





Private Sub cmbTyp_Click()
   On Error Resume Next
   cmbTyp.ToolTipText = sType(cmbTyp.ListIndex, 1)
'    intSortColumn = 1
'    FillGrid
    cmdSearch.Enabled = True


End Sub


Private Sub cmbTyp_LostFocus()
   Dim iList As Integer
   On Error Resume Next
   
'   If Val(cmbTyp) < 1 Or Val(cmbTyp) > 8 Then
'      'Beep
'      cmbTyp = "1"
'   End If
   For iList = 0 To 7
      If Val(cmbTyp) = iList + 1 Then cmbTyp.ToolTipText = sType(iList, 1)
   Next
'   intSortColumn = 1
'   FillGrid
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCheckAll_Click()
   Dim iList As Integer
   For iList = 1 To Grd.Rows - 1
      Grd.Col = 0
      Grd.row = iList
      ' Only if the part is checked
      If Grd.CellPicture = Chkno.Picture Then
          Set Grd.CellPicture = Chkyes.Picture
            cmdUpdateParts.Enabled = True

      End If
   Next

End Sub


Private Sub cmdCheckNone_Click()
   Dim iList As Integer
   For iList = 1 To Grd.Rows - 1
      Grd.Col = 0
      Grd.row = iList
      ' Only if the part is checked
      If Grd.CellPicture = Chkyes.Picture Then
          Set Grd.CellPicture = Chkno.Picture
      End If
   Next
    cmdUpdateParts.Enabled = False

End Sub

'Private Sub cmdDel_Click()
'   If bGoodCode Then UpdateParts
'End Sub


Private Sub cmdFindPart_Click()
   ViewParts.lblControl = "TXTPRT"
   ViewParts.txtPrt = txtPrt
   ViewParts.Show
   bShowParts = 0
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 1351
      cmdHlp = False
      MouseCursor 0
   End If
End Sub



'Private Sub cmdUpdate_Click()
'   UpdatePartsType
'   strPartNo = Compress(txtPartNumber)
'   strPartType = Compress(cmbTyp)
'   ' Re populate the grid
'   ' get Product code
'   strProdCode = Compress(cmbCde)
'   FillGrid strProdCode, strPartNo, strPartType
'
'End Sub

Private Sub cmdSearch_Click()
    intSortColumn = 1
    FillGrid
    cmdSearch.Enabled = False
    
End Sub




Private Sub Form_Activate()
   Dim iType As Integer
      
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      cmbCde.AddItem "ALL"
      FillProductCodes
      FillPartCombo cmbPrt
      
      If cmbCde.ListCount > 0 Then
         cmbCde = cmbCde.List(0)
         bGoodCode = GetCode()
           cmbTyp.AddItem "ALL"
         For iType = 0 To 6
            AddComboStr cmbTyp.hwnd, sType(iType, 0)
         Next
         cmbTyp = cmbTyp.List(0)
         cmbTyp.ToolTipText = sType(0, 1)
         
         ' get Product code
'         strProdCode = Compress(cmbCde.List(0))
'         txtPartNumber = ""
'         strPartNo = ""
'         strPartType = cmbTyp
'         FillGrid strProdCode, strPartNo, strPartType, intSortColumn
         
'         intSortColumn = 1
' FillGrid
    cmdSearch.Enabled = True
    cmdCheckAll.Enabled = False
    cmdCheckNone.Enabled = False
    optPartField(0).Value = True
    


      End If
      

      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
      sSql = "SELECT TOP 1 PARTREF,PARTNUM,PADESC,PALEVEL," _
          & "PAQOH FROM PartTable WHERE PARTREF= ? "

     Set AdoQry = New ADODB.Command
     AdoQry.CommandText = sSql
  
     Set AdoParameter = New ADODB.Parameter
     AdoParameter.Type = adChar
     AdoParameter.Size = 30
  
     AdoQry.Parameters.Append AdoParameter
   
   sType(0, 0) = "1"
   sType(0, 1) = "Top Assembly Unit"
   sType(1, 0) = "2"
   sType(1, 1) = "Intermediate Assembly Unit"
   sType(2, 0) = "3"
   sType(2, 1) = "Base Manufacturing Unit"
   sType(3, 0) = "4"
   sType(3, 1) = "Raw Material Unit"
   sType(4, 0) = "5"
   sType(4, 1) = "Expense Item"
   sType(5, 0) = "6"
   sType(5, 1) = "Expense Item"
   sType(6, 0) = "7"
   sType(6, 1) = "Service Expense Item"
   sType(7, 0) = "8"
   sType(7, 1) = "Project"
   
   With Grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 1
      .Rows = 1
      .row = 0
      .Col = 0
      .Text = "Change"
      .Col = 1
      .Text = "Part Number"
      .Col = 2
      .Text = "Part Description"
      .Col = 3
      .Text = "Prod Code"
      .Col = 4
      .Text = "Part Type"
      .Col = 5
      .Text = "Part Class"
      .Col = 6
      .Text = "QOH"
      
      .ColWidth(0) = 650
      .ColWidth(1) = 2040
      .ColWidth(2) = 3700
      .ColWidth(3) = 1150
      .ColWidth(4) = 950
      .ColWidth(5) = 800
      .ColWidth(6) = 560
      
   End With
   
   
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set PadmPRf03a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub


Private Function GetCode() As Byte
   Dim RdoCde As ADODB.Recordset
   Dim sPcode As String
   sPcode = Compress(cmbCde)
   On Error GoTo DiaErr1
   If Len(sPcode) > 0 Then
      sSql = "Qry_GetProductCode '" & sPcode & "' "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCde, ES_FORWARD)
      If bSqlRows Then
         With RdoCde
            cmbCde = "" & Trim(!PCCODE)
            lblPTDsc = "" & Trim(!PCDESC)
         End With
         GetCode = True
      Else
         If sPcode = "ALL" Then lblPTDsc = "** ALL **" Else lblPTDsc = "*** Product Code Wasn't Found ***"
         GetCode = False
      End If
   End If
   Set RdoCde = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getcode"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Sub cmbPrt_LostFocus()
   If Len(Trim(cmbPrt)) Then bGoodPart = GetPart()
   
End Sub


Private Sub cmbPrt_Click()
   If Len(Trim(cmbPrt)) Then bGoodPart = GetPart()
   
End Sub



Private Sub lblPTDsc_Change()
   If Left(lblPTDsc, 7) = "*** Pro" Then
      lblPTDsc.ForeColor = ES_RED
   Else
      lblPTDsc.ForeColor = vbBlack
   End If
   
End Sub

Private Function FillGrid() As Integer

'Private Function FillGrid(ByVal strProdCode As String, ByVal strPartNo As String, ByVal strPartType As String, ByVal intSortCol As Integer) As Integer
    Dim RdoGrd As ADODB.Recordset
    Dim sSortCol As String
    
    
    On Error Resume Next
    Grd.Rows = 1
    On Error GoTo DiaErr1
    
    If Compress(txtPrt) = "ALL" Then strPartNo = "" Else strPartNo = Trim(txtPrt)
    If Compress(txtPartDesc) = "ALL" Then strPartDesc = "" Else strPartDesc = Trim(txtPartDesc)
    If Compress(cmbCde) = "ALL" Then strProdCode = "" Else strProdCode = Trim(cmbCde)
    If Compress(cmbTyp) = "ALL" Then strPartType = "" Else strPartType = Trim(cmbTyp)
    
    Select Case intSortColumn
    Case 1: sSortCol = "PARTNUM"
    Case 2: sSortCol = "PADESC"
    Case 3: sSortCol = "PAPRODCODE"
    Case 4: sSortCol = "PALEVEL"
    Case 5: sSortCol = "PACLASS"
    Case 6: sSortCol = "PAQOH"
    Case Else
            sSortCol = "PARTNUM"
            strSortOrder = "ASC"
    End Select
    If strSortOrder = "" Or strSortOrder = "DESC" Then strSortOrder = "ASC" Else strSortOrder = "DESC"
        
    
    sSql = "SELECT PARTNUM, PADESC, PAPRODCODE, PALEVEL, PAQOH, PACLASS " & _
         "FROM PartTable " & _
        "WHERE PAPRODCODE LIKE '" & strProdCode & "%' AND PARTNUM LIKE '" & strPartNo & "%' " & _
        "AND PADESC LIKE '" & strPartDesc & "%' " & _
        "AND PALEVEL LIKE '" & strPartType & "%' ORDER BY " & sSortCol & " " & strSortOrder
    
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoGrd, ES_FORWARD)
    If bSqlRows Then
      MouseCursor 13
        With RdoGrd
            Do Until .EOF
            Grd.Rows = Grd.Rows + 1
            Grd.row = Grd.Rows - 1
            Grd.Col = 0
            Set Grd.CellPicture = Chkno.Picture
            Grd.Col = 1
            Grd.Text = "" & Trim(!PartNum)
            Grd.Col = 2
            Grd.Text = "" & Trim(!PADESC)
            Grd.Col = 3
            Grd.Text = "" & Trim(!PAPRODCODE)
            Grd.Col = 4
            Grd.Text = "" & Trim(!PALEVEL)
            Grd.Col = 5
            Grd.Text = "" & Trim(!PACLASS)
            Grd.Col = 6
            Grd.Text = "" & Trim(!PAQOH)
            .MoveNext
         Loop
         ClearResultSet RdoGrd
      End With
      cmdCheckAll.Enabled = True
      cmdCheckNone.Enabled = True
      
   Else
        cmdCheckAll.Enabled = False
        cmdCheckNone.Enabled = False
        
   End If
   MouseCursor 0
   Set RdoGrd = Nothing
   cmdUpdateParts.Enabled = False
   Exit Function
   
DiaErr1:
   MouseCursor 0
   sProcName = "fillgrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub grd_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      Grd.Col = 0
      If Grd.row = 0 Then Grd.row = 1
      If Grd.CellPicture = Chkyes.Picture Then
         Set Grd.CellPicture = Chkno.Picture
      Else
         Set Grd.CellPicture = Chkyes.Picture
         cmdUpdateParts.Enabled = True
         
      End If
    End If
   

End Sub

Private Sub Grd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Grd.row = 1 And Grd.Col > 0 Then
'        strPartNo = txtPartNumber
'        strPartType = cmbTyp
'        strProdCode = cmbCde
'        FillGrid strProdCode, strPartNo, strPartType, Grd.Col
    intSortColumn = Grd.Col
    FillGrid
        Exit Sub
    End If
    
    Grd.Col = 0
    If Grd.Rows = 1 Then Exit Sub
    If Grd.row = 0 Then Grd.row = 1
    If Grd.CellPicture = Chkyes.Picture Then
       Set Grd.CellPicture = Chkno.Picture
    Else
       Set Grd.CellPicture = Chkyes.Picture
        cmdUpdateParts.Enabled = True

       
    End If
End Sub


Private Sub optPartField_Click(Index As Integer)
    FillChangeFieldCombo Index
End Sub

Private Sub txtPartDesc_Change()
    cmdSearch.Enabled = True
    
End Sub


Private Sub txtPrt_Change()
    cmdSearch.Enabled = True
    
End Sub



Private Function GetPart()
   Dim RdoGet As ADODB.Recordset
   On Error GoTo DiaErr1
   'RdoQry(0) = Compress(txtPrt)
   AdoQry.Parameters(0).Value = Compress(cmbPrt)
   bSqlRows = clsADOCon.GetQuerySet(RdoGet, AdoQry, ES_KEYSET)
   If bSqlRows Then
      With RdoGet
         GetPart = True
         cmbPrt = "" & Trim(!PartNum)
         txtPartDesc = "" & Trim(!PADESC)
      '   lblLvl = "" & Format(!PALEVEL, "0")
      '   lblQoh = "" & Format(!PAQOH, ES_QuantityDataFormat)
         ClearResultSet RdoGet
      End With
   Else
      GetPart = False
      txtPartDesc = "*** Invalid Part Number ***"
     ' lblLvl = ""
     ' lblQoh = "0.000"
   End If
   Set RdoGet = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Sub cmdUpdateParts_Click()
    Dim iList, lRows As Long
    Dim strPartsToUpdate, strFieldDesc, strPartRef, strFieldNme As String
    Dim bResponse As Byte

    strPartsToUpdate = ""

    If optPartField(0).Value = True Then
        strFieldNme = "PAPRODCODE"
        strFieldDesc = "Product Code"
    Else
        If optPartField(1).Value = True Then
            strFieldNme = "PACLASS"
            strFieldDesc = "Part Class"
        End If
    End If

    For iList = 1 To Grd.Rows - 1
        Grd.Col = 0
        Grd.row = iList
'        RdoCon.BeginTrans
        If Grd.CellPicture = Chkyes.Picture Then
            Grd.Col = 1
            strPartRef = Compress(Trim(Grd.Text))
            If strPartsToUpdate = "" Then strPartsToUpdate = "'" & strPartRef & "'" Else strPartsToUpdate = strPartsToUpdate & "," & "'" & strPartRef & "'"
            sSql = "Update PartTable SET " & strFieldNme & "='" & arrKeyValue(cmbChangeTo.ItemData(cmbChangeTo.ListIndex)) & "' " & _
                  "WHERE PARTREF IN (" & strPartsToUpdate & ")"
            ' updated part numbers
'            RdoCon.Execute sSql, rdExecDirect
        End If
    Next
    If strPartsToUpdate = "" Then
        MsgBox "You haven't selected any parts to update", vbOKOnly
        Exit Sub
    End If


    On Error GoTo updateparterror
    clsADOCon.BeginTrans
    clsADOCon.ExecuteSQL sSql
    lRows = clsADOCon.RowsAffected

    bResponse = MsgBox("You Are About to Update " & LTrim(str(lRows)) & " Part(s) With " & strFieldDesc & " " & arrKeyValue(cmbChangeTo.ItemData(cmbChangeTo.ListIndex)) & ". Continue?", ES_NOQUESTION, Caption)
    If bResponse = vbYes Then
        clsADOCon.CommitTrans
        MsgBox "Parts Successfully Updated.", vbInformation, Caption
        intSortColumn = 1
        FillGrid
    Else
        clsADOCon.RollbackTrans
    End If

    strPartsToUpdate = ""
    Exit Sub

updateparterror:

    'remember to re-fill grid

End Sub
'

