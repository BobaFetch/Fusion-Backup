VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form SaleSLe05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Re-Schedule Sales Order Item"
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
      TabIndex        =   6
      Top             =   0
      Width           =   915
   End
   Begin VB.ComboBox cmbDueDate 
      Height          =   315
      Left            =   6840
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid gridSOItem 
      Height          =   2895
      Left            =   120
      TabIndex        =   4
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
   Begin VB.ComboBox cmbSon 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblSOType 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1440
      TabIndex        =   10
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblSODate 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1440
      TabIndex        =   9
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "SO Date"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   8
      Top             =   720
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
      TabIndex        =   7
      Top             =   1680
      Width           =   9375
   End
   Begin VB.Label Label1 
      Caption         =   "Customer"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblCustomer 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "SO Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "SaleSLe05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***


Option Explicit
Dim RdoSon As ADODB.Recordset
Dim bEditingCell As Byte
Dim sOrigDate As String

Dim bGoodSO As Byte
Dim bOnLoad As Byte
Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub


Private Sub cmbDueDate_DropDown()
    ShowCalendarEx Me
End Sub

Private Sub cmbSon_Click()
   bGoodSO = GetSalesOrder(0)
End Sub

Private Sub cmbSon_GotFocus()
   bGoodSO = GetSalesOrder(0)
End Sub

Private Sub cmbSon_LostFocus()
   cmbSon = CheckLen(cmbSon, SO_NUM_SIZE)
   cmbSon = Format(Abs(Val(cmbSon)), SO_NUM_FORMAT)
   bGoodSO = GetSalesOrder(1)
End Sub


Private Sub cmdCan_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
   If bOnLoad Then
      SetupGrid
      FillSOs cmbSon, False
      If cmbSon.ListCount > 0 Then cmbSon.ListIndex = 0
      bOnLoad = 0
   Else
      MouseCursor 0
      bGoodSO = GetSalesOrder(0)
   End If
End Sub

Private Sub Form_Load()
   bEditingCell = 0
   FormLoad Me
   FormatControls
   bOnLoad = 1
End Sub


Private Sub Form_Resize()
   Refresh
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set RdoSon = Nothing
   Set SaleSLe05a = Nothing
End Sub


Private Function GetSalesOrder(Optional bMessage As Byte) As Byte
   On Error GoTo DiaErr1
   GetSalesOrder = 0
   If Len(Compress(cmbSon)) = 0 Then Exit Function
   MouseCursor 13
   
   sSql = "SELECT TOP 1 SONUMBER, SOTYPE, SODATE, SOCUST, SOTEXT, CUNAME " _
          & "FROM SohdTable " _
          & " INNER JOIN CustTable ON SOCUST=CUREF " _
          & "WHERE SONUMBER = " & cmbSon
   'Debug.Print sSql
          '& " AND POCAN=0"

   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSon)
   If bSqlRows Then
        If Len(Trim("" & RdoSon!CUNAME)) = 0 Then lblCustomer.Caption = "" & RdoSon!SOCUST Else lblCustomer.Caption = "" & RdoSon!CUNAME
        lblSOType.Caption = Trim("" & RdoSon!SOTYPE)
        FillGridWithSOItems Trim("" & RdoSon!SoNumber)
        cmbSon = Format(RdoSon!SoNumber, SO_NUM_FORMAT)
        lblSODate = Format(RdoSon!SODATE, "mm/dd/YYYY")
        GetSalesOrder = 1
   Else
        GetSalesOrder = 0
        On Error Resume Next
        cmbSon.SetFocus
   End If
   MouseCursor 0
   Set RdoSon = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getsaleso"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function




Public Sub FillSOs(ByRef cmbSO As ComboBox, Optional IncludeCancelled As Boolean = False)
   Dim RdoSO As ADODB.Recordset
   On Error GoTo FillSOErr1
   sSql = "SELECT SONUMBER FROM SohdTable "
   If Not IncludeCancelled Then sSql = sSql & "WHERE SOCANCELED=0 "
   sSql = sSql & "ORDER BY SONUMBER DESC "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSO, ES_FORWARD)
   If bSqlRows Then
         Do Until RdoSO.EOF
            cmbSO.AddItem Right(SO_NUM_FORMAT & Trim("" & RdoSO!SoNumber), SO_NUM_SIZE)
            RdoSO.MoveNext
         Loop
         ClearResultSet RdoSO
   End If
   Set RdoSO = Nothing
   Exit Sub
   
FillSOErr1:
   sProcName = "FillSOs"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
End Sub


Private Sub FillGridWithSOItems(ByVal SoNumber As String)
    Dim RdoSOItem As ADODB.Recordset
    Dim iRow As Integer
    
    On Error GoTo fillgriderror1
    iRow = 1
   
    sSql = "SELECT ITNUMBER, ITPART, ITREV, ITSCHED, PADESC, ITQTY, ITDOLLARS from SoitTable INNER JOIN PartTable ON ITPART=PARTREF WHERE ITSO=" & Trim(SoNumber) & " ORDER BY ITNUMBER"
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoSOItem, ES_FORWARD)
    If bSqlRows Then
        Do Until RdoSOItem.EOF
            iRow = iRow + 1
            gridSOItem.Rows = iRow
            gridSOItem.Row = iRow - 1
            gridSOItem.RowHeight(iRow - 1) = 315
            
            gridSOItem.Col = 0
            gridSOItem.Text = Trim("" & RdoSOItem!ITNUMBER)
            
            gridSOItem.Col = 1
            gridSOItem.Text = Trim("" & RdoSOItem!ITREV)
            
            gridSOItem.Col = 2
            gridSOItem.Text = Trim("" & RdoSOItem!ITPART)
            
            gridSOItem.Col = 3
            gridSOItem.Text = Trim("" & RdoSOItem!PADESC)
            
            gridSOItem.Col = 4
            gridSOItem.Text = Format(RdoSOItem!ITQTY, "###0.0000")
            gridSOItem.Col = 5
            gridSOItem.Text = Format(RdoSOItem!ITDOLLARS, "$#,##0.0000")
            
            gridSOItem.Col = 6
            gridSOItem.Text = Format(RdoSOItem!ITSCHED, "mm/dd/YY")
            
        
            RdoSOItem.MoveNext
        Loop
        ClearResultSet RdoSOItem
    End If
    Set RdoSOItem = Nothing
    Exit Sub
    
fillgriderror1:
    sProcName = "FillGridWithSOItems"
    CurrError.Number = Err
    CurrError.Description = Err.Description
    DoModuleErrors MdiSect.ActiveForm
End Sub


Private Sub SetupGrid()
   With gridSOItem
      .Rows = 2
      .Cols = 7
      .RowHeight(0) = 315

      .ColWidth(0) = 425
      .ColWidth(1) = 400
      .ColWidth(2) = 1500
      .ColWidth(3) = 3600
      .ColWidth(4) = 800
      .ColWidth(5) = 900
      .ColWidth(6) = 1200


      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .ColAlignment(2) = 0
      .ColAlignment(3) = 0
      .ColAlignment(4) = 0
      .ColAlignment(5) = 0
      .ColAlignment(6) = 0

      .Row = 0
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
      .Text = "Ship Date"

   End With

End Sub



Private Sub gridSOItem_DblClick()
    If gridSOItem.Col <> 6 Then Exit Sub

   'position the edit box
   cmbDueDate.Left = gridSOItem.CellLeft + gridSOItem.Left
   cmbDueDate.Top = gridSOItem.CellTop + gridSOItem.Top
   cmbDueDate.Width = gridSOItem.CellWidth
   cmbDueDate.Visible = True
   cmbDueDate = gridSOItem.Text
   sOrigDate = gridSOItem.Text
   
   cmbDueDate.SetFocus
   bEditingCell = 1
End Sub

Private Sub gridSOItem_LeaveCell()

    If (bEditingCell = 1) And (sOrigDate <> cmbDueDate) Then
        bEditingCell = 0
        sSql = "Update SoitTable SET ITSCHED='" & cmbDueDate & "' "
        'If MsgBox("Would you also like to also change the original due date to " & cmbDueDate & " ?", vbYesNo) = vbYes Then
        '    sSql = sSql & ",PIPORIGDATE='" & cmbDueDate & "' "
        'End If
        gridSOItem.Col = 0
        sSql = sSql & "WHERE ITSO=" & cmbSon & " AND ITNUMBER=" & gridSOItem.Text
        gridSOItem.Col = 1
        If Len(Trim(gridSOItem)) > 0 Then sSql = sSql & " AND ITREV='" & Trim(gridSOItem.Text) & "' "
                
        cmbDueDate.Visible = False
        clsADOCon.ExecuteSQL sSql 'rdExecDirect
    
        'MsgBox "Row = " & gridPOItem.row & "  COl=" & gridPOItem.Col & "  gridpoitem.text=" & gridPOItem.Text
        gridSOItem.Col = 6
        gridSOItem.Text = cmbDueDate
    Else
        cmbDueDate.Visible = False
    End If
End Sub


