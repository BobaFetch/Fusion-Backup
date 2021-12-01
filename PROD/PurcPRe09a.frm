VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PurcPRe09a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Re-Schedule Purchase Order Item"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   5100
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   8520
      TabIndex        =   10
      Top             =   0
      Width           =   915
   End
   Begin VB.ComboBox cmbDueDate 
      Height          =   315
      Left            =   6000
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid gridPOItem 
      Height          =   2895
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Double Click on the Due Date to Edit"
      Top             =   2040
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   5106
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      ScrollBars      =   2
      AllowUserResizing=   1
   End
   Begin VB.CheckBox cbServicePO 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   1200
      Width           =   615
   End
   Begin VB.CheckBox cbTaxable 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4320
      TabIndex        =   2
      Top             =   1200
      Width           =   495
   End
   Begin VB.ComboBox cmbPon 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblPODate 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3960
      TabIndex        =   13
      Top             =   335
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "PO Date"
      Height          =   255
      Index           =   5
      Left            =   3240
      TabIndex        =   12
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Double-Click on the Due Date Column to Change/Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   9375
   End
   Begin VB.Label Label1 
      Caption         =   "Service PO?"
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Taxable?"
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Vendor"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblVendor 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "PO Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "PurcPRe09a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'10/29/2010 - CREATED!!

Option Explicit
Dim RdoPon As ADODB.Recordset
Dim bEditingCell As Byte
Dim sOrigDate As String



Dim bGoodPo As Byte
Dim bOnLoad As Byte
Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   'cmbReq.Tag = 2
   
End Sub


Private Sub cmbDueDate_DropDown()
    ShowCalendarEx Me
End Sub

Private Sub cmbPon_Click()
   bGoodPo = GetPurchaseOrder(0)
   
End Sub


Private Sub cmbPon_GotFocus()
   bGoodPo = GetPurchaseOrder(0)
   
End Sub


Private Sub cmbPon_LostFocus()
   cmbPon = CheckLen(cmbPon, 6)
   cmbPon = Format(Abs(Val(cmbPon)), "000000")

   bGoodPo = GetPurchaseOrder(1)
   
End Sub


Private Sub cmdCan_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
   If bOnLoad Then
      SetupGrid
      FillPOs cmbPon, False
      If cmbPon.ListCount > 0 Then cmbPon.ListIndex = 0
      bOnLoad = 0
   Else
      MouseCursor 0
      bGoodPo = GetPurchaseOrder(0)
   End If
End Sub

Private Sub Form_Load()
   bEditingCell = 0
   FormLoad Me
   FormatControls
   'Tab1.Tab = 0
'   MouseCursor 13
'   tabFrame(0).BorderStyle = 0
'   tabFrame(1).BorderStyle = 0
'   tabFrame(0).Visible = True
'   tabFrame(1).Visible = False
'   tabFrame(0).Left = 10
'   tabFrame(1).Left = 10
'   Set RdoPon = RdoCon.CreateQuery("", sSql)
   
'   sSql = "SELECT PINUMBER,PITYPE FROM PoitTable " _
'          & "WHERE PINUMBER= ? AND PITYPE<>14"
'   Set RdoRcd = RdoCon.CreateQuery("", sSql)
'   RdoRcd.MaxRows = 1
   bOnLoad = 1

End Sub


Private Sub Form_Resize()
   Refresh
End Sub


Private Sub Form_Unload(Cancel As Integer)
'   Set RdoPso = Nothing
   Set RdoPon = Nothing
'   Set RdoRcd = Nothing
   Set PurcPRe09a = Nothing
   
End Sub



Private Function GetPurchaseOrder(Optional bMessage As Byte) As Byte
   'Dim RdoVnd As ADODB.Recordset
   On Error GoTo DiaErr1
'   If bGoodPo = 1 And bDataHasChanged Then UpdatePurchaseOrder
   MouseCursor 13
'   ClearBoxes
   
   
   sSql = "SELECT TOP 1 PONUMBER,POTYPE,POVENDOR,PODATE," _
          & "POBUYER,POREQBY,PODIVISION,POSHIP,POSERVICE,POSHIPTO," _
          & "POBCONTACT,POFOB,POPRINTED,POREMARKS,POSTERMS," _
          & "PONETDAYS,PODDAYS,PODISCOUNT,POPROXDT,POPROXDUE,POBUYER," _
          & "POVIA,POTAXABLE,VEREF,VENICKNAME,VEBNAME FROM PohdTable " _
          & " INNER JOIN VndrTable ON POVENDOR=VEREF " _
          & "WHERE PONUMBER = " & cmbPon & " AND POCAN=0"


'   RdoPon.RowsetSize = 1
'   RdoPon(0) = Val(cmbPon)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPon)
   
   If bSqlRows Then
   
        If Len(Trim("" & RdoPon!VEBNAME)) = 0 Then lblVendor.Caption = "" & RdoPon!VEREF Else lblVendor.Caption = "" & RdoPon!VEBNAME
        
FillGridWithPOItems Trim("" & RdoPon!PONUMBER)
'      With RdoPso

         cmbPon = Format(RdoPon!PONUMBER, "000000")
         If RdoPon!POSERVICE Then cbServicePO.Value = 1 Else cbServicePO.Value = vbUnchecked
         If RdoPon!POTAXABLE Then cbTaxable.Value = 1 Else cbTaxable.Value = vbUnchecked
         lblPODate = Format(RdoPon!PODATE, "mm/dd/YYYY")
         
         
         
         
'         cmbVnd = "" & Trim(!VENICKNAME)
'         lblNme = "" & !VEBNAME
'         lblPdt = "" & Format(!PODATE, "mm/dd/yy")
'         cmbTyp = "" & Trim(!POTYPE)
'         txtSdt = "" & Format(!POSHIP, "mm/dd/yy")
'         cmbDiv = "" & Trim(!PODIVISION)
'         cmbByr = "" & Trim(!POBUYER)
'         txtCnt = "" & Trim(!POBCONTACT)
'         txtFob = "" & Trim(!POFOB)
'         cmbTrm = "" & Trim(!POSTERMS)
'         txtShp = "" & Trim(!POSHIPTO)
'         txtVia = "" & Trim(!POVIA)
'2         txtCmt = "" & !POREMARKS
'         lblPrn = "" & Format(!POPRINTED, "mm/dd/yy")
'         If txtSdt = "" Then txtSdt = Format(ES_SYSDATE, "mm/dd/yy")
'         txtPdt = lblPdt
'         cmbReq = "" & Trim(!POREQBY)
'      End With
'      If Len(Trim(cmbByr)) > 0 Then GetCurrentBuyer cmbByr
      GetPurchaseOrder = 1
'      cmdTrm.Enabled = True
'      cmdItm.Enabled = True
   Else
'      If bMessage Then
'         If Val(cmbPon) > 0 Then MsgBox "Purchase Order " & cmbPon & " Was     " & vbCr _
'                & "Canceled Or Doesn't Exist.", _
'                vbExclamation, Caption
'         If cmbPon.ListCount > 0 Then cmbPon = cmbPon.List(0)
'      End If
      GetPurchaseOrder = 0
'      cmdTrm.Enabled = False
'      cmdItm.Enabled = False
      On Error Resume Next
      cmbPon.SetFocus
   End If
'   If GetPurchaseOrder = 1 Then
'      RdoRcd.RowsetSize = 1
'      RdoRcd(0) = Val(cmbPon)
'      bSqlRows = clsAdoCon.GetQuerySet(RdoVnd, RdoRcd, ES_KEYSET)
'      If bSqlRows Then
'         z1(2).Enabled = False
'         cmbVnd.Enabled = False
'      Else
'         z1(2).Enabled = True
'         cmbVnd.Enabled = True
'      End If
'   End If
'   bDataHasChanged = False
   MouseCursor 0
   Set RdoPon = Nothing
'   Set RdoVnd = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpurcha"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function




Public Sub FillPOs(ByRef cmbPO As ComboBox, Optional IncludeCancelled As Boolean = False)
   Dim RdoPO As ADODB.Recordset
   On Error GoTo FillPOErr1
   sSql = "SELECT PONUMBER FROM PohdTable "
   If Not IncludeCancelled Then sSql = sSql & "WHERE POCAN=0 "
   sSql = sSql & "ORDER BY PONUMBER DESC "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPO, ES_FORWARD)
   If bSqlRows Then
         Do Until RdoPO.EOF
            cmbPO.AddItem Trim("" & RdoPO!PONUMBER)
            RdoPO.MoveNext
         Loop
         ClearResultSet RdoPO
   End If
'   If MDISect.ActiveForm.cmbRte.ListCount > 0 Then _
'      MDISect.ActiveForm.cmbRte.Text = MDISect.ActiveForm.cmbRte.List(0)
   Set RdoPO = Nothing
   Exit Sub
   
FillPOErr1:
   sProcName = "FillPOs"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
End Sub


Private Sub FillGridWithPOItems(ByVal PONUMBER As String)
    Dim RdoPOItem As ADODB.Recordset
    Dim iRow As Integer
    
    On Error GoTo fillgriderror1
    
    'gridPOItem.Clear
    iRow = 1
   
    
    
    sSql = "SELECT PIPART, PIITEM, PIREV, PIPDATE, PADESC, PIESTUNIT, PIPQTY, PAUNITS, PIRUNPART, PIRUNNO from PoitTable INNER JOIN PartTable ON PIPART=PARTREF WHERE PINUMBER=" & Trim(PONUMBER) & " ORDER BY PIITEM"
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoPOItem, ES_FORWARD)
    If bSqlRows Then
        Do Until RdoPOItem.EOF
            iRow = iRow + 1
            gridPOItem.Rows = iRow
            gridPOItem.row = iRow - 1
            gridPOItem.RowHeight(iRow - 1) = 315
            
            gridPOItem.Col = 0
            gridPOItem.Text = Trim("" & RdoPOItem!PIITEM)
            
            gridPOItem.Col = 1
            gridPOItem.Text = Trim("" & RdoPOItem!PIREV)
            
            gridPOItem.Col = 2
            gridPOItem.Text = Trim("" & RdoPOItem!PIPART)
            
            gridPOItem.Col = 3
            gridPOItem.Text = Trim("" & RdoPOItem!PADESC)
            
            gridPOItem.Col = 4
            gridPOItem.Text = Format(RdoPOItem!PIPQTY, "###0.0000")
            gridPOItem.Col = 5
            gridPOItem.Text = Format(RdoPOItem!PIESTUNIT, "$#,##0.0000")
            gridPOItem.Col = 6
            gridPOItem.Text = Trim("" & RdoPOItem!PAUNITS)
            
            gridPOItem.Col = 7
            gridPOItem.Text = Format(RdoPOItem!PIPDATE, "mm/dd/YY")
            
        
            RdoPOItem.MoveNext
        Loop
        ClearResultSet RdoPOItem
    End If
    Set RdoPOItem = Nothing
    Exit Sub
    
fillgriderror1:
    sProcName = "FillGridWithPOItems"
    CurrError.Number = Err
    CurrError.Description = Err.Description
    DoModuleErrors MDISect.ActiveForm
End Sub


Private Sub SetupGrid()
   With gridPOItem
      .Rows = 2
      .Cols = 8
      .RowHeight(0) = 315

      .ColWidth(0) = 425
      .ColWidth(1) = 400
      .ColWidth(2) = 1500
      .ColWidth(3) = 3600
      .ColWidth(4) = 800
      .ColWidth(5) = 900
      .ColWidth(6) = 300
      .ColWidth(7) = 1200

      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .ColAlignment(2) = 0
      .ColAlignment(3) = 0
      .ColAlignment(4) = 0
      .ColAlignment(5) = 0
      .ColAlignment(6) = 0
      .ColAlignment(7) = 0
      .row = 0
      .Col = 0
      .Text = "Item"
      .Col = 1
      .Text = "Rev"
      .Col = 2
      .Text = "Part Number"
      .Col = 3
      .Text = "Part Description"
      .Col = 4
      .Text = "Qty"
      .Col = 5
      .Text = "Unit Price"
      .Col = 6
      .Text = "Unit"
      .Col = 7
      .Text = "Due Date"
   End With

End Sub

Private Sub gridPOItem_Click()
'MsgBox "click"
End Sub



Private Sub gridPOItem_DblClick()
    If gridPOItem.Col <> 7 Then Exit Sub

   'position the edit box
   cmbDueDate.Left = gridPOItem.CellLeft + gridPOItem.Left
   cmbDueDate.Top = gridPOItem.CellTop + gridPOItem.Top
   cmbDueDate.Width = gridPOItem.CellWidth
'   cmbDueDate.Height = gridPOItem.CellHeight
   cmbDueDate.Visible = True
   cmbDueDate = gridPOItem.Text
   sOrigDate = gridPOItem.Text
   
   cmbDueDate.SetFocus
   bEditingCell = 1
End Sub

Private Sub gridPOItem_LeaveCell()

    If (bEditingCell = 1) And (sOrigDate <> cmbDueDate) Then
        bEditingCell = 0
        'gridPOItem.Col = 0
        sSql = "Update PoitTable SET PIPDATE='" & cmbDueDate & "' "
        If MsgBox("Would you also like to also change the original due date to " & cmbDueDate & " ?", vbYesNo) = vbYes Then
            sSql = sSql & ",PIPORIGDATE='" & cmbDueDate & "' "
        End If
        gridPOItem.Col = 0
        sSql = sSql & "WHERE PINUMBER=" & cmbPon & " AND PIITEM=" & gridPOItem.Text
        gridPOItem.Col = 1
        If Len(Trim(gridPOItem)) > 0 Then sSql = sSql & " AND PIREV='" & Trim(gridPOItem.Text) & "' "
                
        cmbDueDate.Visible = False
        clsADOCon.ExecuteSQL sSql
    
        'MsgBox "Row = " & gridPOItem.row & "  COl=" & gridPOItem.Col & "  gridpoitem.text=" & gridPOItem.Text
        gridPOItem.Col = 7
        gridPOItem.Text = cmbDueDate
        
        
        'FillGridWithPOItems cmbPon
        
    
    Else
        cmbDueDate.Visible = False
    End If
End Sub


