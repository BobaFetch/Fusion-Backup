VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form PickMCe08 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sheet Restock"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10725
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboDate 
      Height          =   315
      Left            =   660
      TabIndex        =   18
      Tag             =   "4"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel this pick"
      Height          =   540
      Left            =   180
      TabIndex        =   17
      Top             =   1920
      Width           =   2040
   End
   Begin VB.TextBox txtLocation 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      MaxLength       =   4
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   120
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdRestock 
      Caption         =   "Restock"
      Height          =   375
      Left            =   8400
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtComments 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1440
      Width           =   8055
   End
   Begin VB.TextBox txtLeadingLotID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.ComboBox cboLotID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4440
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   9600
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid gridSheets 
      Height          =   2835
      Left            =   120
      TabIndex        =   4
      Top             =   3780
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   5001
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      RowHeightMin    =   285
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   19
      Top             =   1020
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Location"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   16
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblInstructions 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3 lines of text set in Form_Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   120
      TabIndex        =   14
      Top             =   2940
      Width           =   10455
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblSO 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   660
      TabIndex        =   13
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "SO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Comments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Part"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblPartNo 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Tracking # Leading Chars"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblPartDesc 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   7
      Top             =   600
      Width           =   5055
   End
End
Attribute VB_Name = "PickMCe08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Enum PICK_COL
        COL_LOIRECORD
        COL_OLD_HEIGHT
        COL_OLD_LENGTH
        COL_NEW_HEIGHT
        COL_NEW_LENGTH
        COL_COMMENTS
        COL_LOC
        COL_PK_BY
        COL_PK_ON
    End Enum
    
Dim priorLeadingCharacters As String
Private Sub cboDate_DropDown()
   ShowCalendarEx Me
End Sub

Private Sub cboDate_LostFocus()
   cboDate = CheckDateEx(cboDate)
End Sub

Private Sub cboLotID_Change()
    RefreshGrid
End Sub

'  note the funny CHECKED character -- it is not a "p"
'    Private Const Checked As String = "þ"
'    Private Const UNCHECKED As String = "q"

Private Sub cboLotID_Click()
    lblPartNo = ""
    lblPartDesc = ""
    txtComments = ""
    Dim Key As String
    Key = cboLotID.Text
    If Key = "" Then Exit Sub
    
    'get part description
    sSql = "select top 1 RTRIM(PARTNUM) AS PARTNUM, RTRIM(PADESC) AS PADESC, LOTCOMMENTS, LOISONUMBER, rtrim(LOTLOCATION) AS LOTLOCATION" & vbCrLf _
        & "from PartTable pt" & vbCrLf _
        & "join LohdTable lh on lh.LOTPARTREF = pt.PARTREF" & vbCrLf _
        & "join LoitTable li on li.LOINUMBER = lh.LOTNUMBER" & vbCrLf _
        & "where LOTUSERLOTID = '" & Key & "' and LOISHEETACTTYPE = 'PK'"
    Dim rdo As ADODB.Recordset
    If clsADOCon.GetDataSet(sSql, rdo, adUseClient) Then
        lblPartNo = rdo("PARTNUM")
        lblPartDesc = rdo("PADESC")
        txtComments = rdo("LOTCOMMENTS")
        txtLocation = rdo("LOTLOCATION")
        Me.lblSO = rdo("LOISONUMBER")
'        Me.lblHeightUnits = rdo("LOTMATLENGTHUM")
'        Me.lblWidthUnits = rdo("LOTMATHEIGHTHUM")
    End If
    
    RefreshGrid
End Sub

Private Sub RefreshGrid()
    Dim Key As String
    Key = cboLotID.Text
    gridSheets.Clear
    gridSheets.Rows = 0
   
    Dim rdo As ADODB.Recordset
    sSql = "select LOIRECORD," & vbCrLf _
        & "ISNULL(LOIHEIGHT,0) as [OLD HT], ISNULL(LOILENGTH,0) as [OLD LEN]," & vbCrLf _
        & "ISNULL(LOIHEIGHT,0) as [NEW HT], ISNULL(LOILENGTH,0) as [NEW LEN]," & vbCrLf _
        & "LOICOMMENT as COMMENTS, " & vbCrLf _
        & "LOTLOCATION as LOC,LOIUSER as [PK BY], ISNULL(CONVERT(varchar(8),LOIADATE,1),'') as [PK ON]" & vbCrLf _
        & "from LohdTable lh" & vbCrLf _
        & "join PartTable p on p.PARTREF = lh.LOTPARTREF" & vbCrLf _
        & "join LoitTable li on lh.LOTNUMBER = li.LOINUMBER" & vbCrLf _
        & "where LOTUSERLOTID = '" & Key & "' AND LOICLOSED is NULL" & vbCrLf _
        & "and LOISHEETACTTYPE = 'PK'" & vbCrLf _
        & "ORDER BY lh.LOTUSERLOTID, li.LOIAREA desc"
    bSqlRows = clsADOCon.GetDataSet(sSql, rdo, adUseClient)
    If bSqlRows Then
        'lblPart = rdo("PADESC")
        Dim iCol As Integer
        Dim iRow As Integer
        Dim fld As ADODB.Field
        
        'insert headers
        With gridSheets
            .Cols = rdo.Fields.count
            .Rows = rdo.RecordCount + 1
            .ColWidth(COL_LOIRECORD) = 0            'don't show
            .ColWidth(COL_OLD_HEIGHT) = 900
            .ColWidth(COL_OLD_LENGTH) = 900
            .ColWidth(COL_NEW_HEIGHT) = 900
            .ColWidth(COL_NEW_LENGTH) = 900
            .ColWidth(COL_COMMENTS) = 3000
            .ColWidth(COL_LOC) = 700
            .ColWidth(COL_PK_BY) = 700
            .ColWidth(COL_PK_ON) = 800
            
            Dim gridWidth As Integer
            gridWidth = 360 'allow for scrollbar
            Dim i As Integer
            For i = 0 To .Cols - 1
                gridWidth = gridWidth + .ColWidth(i)
            Next i
            .Width = gridWidth
            Dim minWidth As Integer
            minWidth = Me.txtComments.Left + txtComments.Width + 300
            If gridWidth > minWidth Then Me.Width = gridWidth + 600 Else Me.Width = minWidth

            .FixedRows = 1
            
            ' use winddings font for all column 0 (except header)
'            For iRow = 0 To rdo.RecordCount
'                .row = iRow
'                If iRow > 0 Then
'                    .Col = 0
'                    .FillStyle = flexFillRepeat
'                    .CellFontName = "Wingdings"
'                    .CellFontSize = 12
'                    .CellAlignment = flexAlignCenterCenter
'                    .FillStyle = flexFillSingle
'                End If
'
'            Next iRow
            
            iCol = 0
            For Each fld In rdo.Fields
'                If iCol = COL_CHECKBOX Then
'                    gridSheets.TextMatrix(0, iCol) = "PK"
'                Else
                    gridSheets.TextMatrix(0, iCol) = fld.Name
'                End If
                iCol = iCol + 1
            Next fld
        
            iRow = 1
            Do Until rdo.EOF
                iCol = 0
                For Each fld In rdo.Fields
'                    If iCol = 0 Then
'                        gridSheets.TextMatrix(iRow, iCol) = UNCHECKED
'                    Else
                        If IsNull(fld.Value) = True Then
                            gridSheets.TextMatrix(iRow, iCol) = vbNullString
                        Else
                            gridSheets.TextMatrix(iRow, iCol) = fld.Value
                        End If
'                    End If
                    iCol = iCol + 1
                Next fld
                iRow = iRow + 1
                rdo.MoveNext
            Loop
            
            'allow entry of new rows
            .Rows = .Rows + 3
            
        End With
    End If
End Sub

Private Sub cmdCancel_Click()
   SheetCancel
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRestock_Click()
    SheetRestock
End Sub

'Private Sub cmdZero_Click()
'Chuck says this is not required, so made the button invisible and left the incomplete code, just in case AWJ wants it later.
'    If MsgBox("Has this lot been fully consumed?", vbYesNo, "Zero the balance for this lot") <> vbYes Then
'      deactivateLot = True
'         SheetRestock
'      deactivateLot = False
'    End If
'End Sub

Private Sub Form_Activate()
    MouseCursor 0
End Sub

Private Sub Form_Load()
    lblInstructions = _
        "Set NEW HT and NEW LEN for all remaining rectangles (double-click to edit then press enter key to exit field.  " _
        & "Rectangles for which height and width are not entered will remain a part of the previous SO cost.  " _
        & "Click the Restock button to save the data."
    priorLeadingCharacters = "not initialized"
    
    cboDate = Format(ES_SYSDATE, "mm/dd/yyyy")
End Sub

Private Sub SheetCancel()
    With gridSheets
        If .Rows = 0 Then
            MsgBox "No items to cancel"
            Exit Sub
        Else
            If MsgBox("Cancel this pick?", vbYesNo, "Cancel Pick") <> vbYes Then
               Exit Sub
            End If
        End If
    End With
   
    sSql = "exec SheetCancelPick '" & Me.cboLotID & "'"
    clsADOCon.ExecuteSql sSql
    
    txtLeadingLotID_LostFocus
    RefreshGrid
    
    MsgBox "Pick has been canceled."
    
End Sub

Private Sub SheetRestock()
    With gridSheets

        If .Rows = 0 Then
            MsgBox "No items to restock"
            Exit Sub
        End If
        
        'check data
        Dim params As String, recNo As Integer, oldHt As Currency, oldLen As Currency, newHt As Currency, newLen As Currency
        Dim row As Integer, rectCount As Integer, Errors As Integer
        row = 1
        For row = 1 To gridSheets.Rows - 1
            recNo = CInt("0" & .TextMatrix(row, COL_LOIRECORD))
            oldHt = CCur("0" & .TextMatrix(row, COL_OLD_HEIGHT))
            oldLen = CCur("0" & .TextMatrix(row, COL_OLD_LENGTH))
            newHt = CCur("0" & .TextMatrix(row, COL_NEW_HEIGHT))
            newLen = CCur("0" & .TextMatrix(row, COL_NEW_LENGTH))
            
            If (newHt = 0 And newLen <> 0) Or (newHt <> 0 And newLen = 0) Then
                Errors = Errors + 1
            End If
            
            If newHt <> 0 Then
                rectCount = rectCount + 1
            End If
            
            If recNo <> 0 Or newHt <> 0 Then
                params = params & recNo & "," & newHt & "," & newLen & ","
            End If
        Next row
    End With
    
    If Errors > 0 Then
        MsgBox "Please fix " & Errors & " errors", vbOKOnly, "ERROR"
        Exit Sub
    End If
    
    If MsgBox("Restock " & rectCount & " rectangles?", vbYesNo, "Sheet Restock") <> vbYes Then
        Exit Sub
    End If
    
    sSql = "exec SheetRestock '" & Me.cboLotID & "', '" & sInitials & "', '" & Me.txtComments & "'," & vbCrLf _
      & "'" & Me.txtLocation & "', '" & cboDate & "', '" & params & "'"
    clsADOCon.ExecuteSql sSql
    
    txtLeadingLotID_LostFocus
    RefreshGrid
    
    MsgBox "Remainder of sheet has been restocked"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
End Sub

Private Sub gridSheets_LostFocus()
    'gridSheets_LeaveCell
End Sub

Private Sub Text1_LostFocus()
    gridSheets_LeaveCell
End Sub

Private Sub txtLeadingLotID_LostFocus()
    Dim leading As String
    leading = UCase$(Compress(txtLeadingLotID.Text))
    
    If leading = priorLeadingCharacters Then
        Exit Sub
    End If
    
    Me.cboLotID.Clear
    sSql = "select distinct RTRIM(LOTUSERLOTID) AS LOTUSERLOTID from LohdTable lh" & vbCrLf _
        & "join LoitTable li on li.LOINUMBER = lh.LOTNUMBER" & vbCrLf _
        & "where LOTUSERLOTID like '" & leading & "%'" & vbCrLf _
        & "and LOIADATE IS NOT NULL" & vbCrLf _
        & "and LOISHEETACTTYPE = 'PK'" & vbCrLf _
        & "and LOICLOSED IS NULL" & vbCrLf _
        & "order by LOTUSERLOTID desc"
    Dim rdo As ADODB.Recordset
    bSqlRows = clsADOCon.GetDataSet(sSql, rdo, adUseClient)
    If bSqlRows Then
        'ReDim combodata(rdo.RecordCount)
        With rdo
            Do Until .EOF
                cboLotID.AddItem CStr(!LOTUSERLOTID)
                'combodata(cboParts.NewIndex) = !PartRef
                rdo.MoveNext
            Loop
        End With
    End If
    priorLeadingCharacters = leading
End Sub

Private Sub gridSheets_DblClick()
    If gridSheets.Col = COL_NEW_LENGTH Or gridSheets.Col = COL_NEW_HEIGHT Then
        GridEdit Asc(" ")
    End If
End Sub

Private Sub gridSheets_KeyPress(KeyAscii As Integer)
    'GridEdit KeyAscii
End Sub

Sub GridEdit(KeyAscii As Integer)
    'use correct font
    Text1.FontName = gridSheets.FontName
    Text1.FontSize = gridSheets.FontSize
    
    '    Select Case KeyAscii
'
'       Case 0 To Asc(" ")
'          Text1 = gridSheets
'          Text1.Text = Trim(Text1.Text)
'          Text1.SelStart = 1000
'
'       Case Else
'           Text1 = gridSheets
'           Text1.Text = Trim(Text1.Text)
'          Text1.SelStart = 1000
'
'    End Select

    Text1 = gridSheets
    Text1.Text = Trim(Text1.Text)
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)


    'position the edit box
    Text1.Left = gridSheets.CellLeft + gridSheets.Left
    Text1.Top = gridSheets.CellTop + gridSheets.Top + 40
    Text1.Width = gridSheets.CellWidth
    Text1.Height = gridSheets.CellHeight - 50
    
    Text1.Visible = True
    Text1.SetFocus
    
End Sub

Private Sub gridSheets_LeaveCell()

  If Text1.Visible Then

  If gridSheets.Col = COL_NEW_LENGTH Or gridSheets.Col = COL_NEW_HEIGHT Then
        If Text1.Text = "" Then
            Text1.Text = " "
        End If
    End If
     gridSheets = Text1
     Text1.Visible = False

  End If

End Sub

'Private Sub gridSheets_GotFocus()
'
'  If Text1.Visible Then
'
'  If gridSheets.Col = COL_NEW_LENGTH Or gridSheets.Col = COL_NEW_HEIGHT Then
'
'        If Text1.Text = "" Then
'            Text1.Text = " "
'        End If
'
'    End If
'
'     gridSheets = Text1.Text
'     Text1.Visible = False
'
'  End If
'
'End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

   'noise suppression
   
   If gridSheets.Col <> COL_NEW_LENGTH And gridSheets.Col <> COL_NEW_HEIGHT Then
      KeyAscii = 0
   End If
   
   If KeyAscii = vbKeyReturn Then
      gridSheets.SetFocus
      KeyAscii = 0
   End If
   
   ' always allow backspace
   If KeyAscii = vbKeyBack Then
      Exit Sub
   
   '   ' allow 0 - 9 if currently no numbers after decimal
'   ElseIf (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
'      If Len(Text1.Text) = 0 Then
'         Exit Sub
'      ElseIf InStr(1, Text1.Text, ".") = 0 Or Right(Text1.Text, 1) = "." Then
'         Exit Sub
'      End If

   ' allow 0 - 9
   ElseIf (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
      Exit Sub
   
   ' allow zero or one decimal point
   ElseIf KeyAscii = 46 And InStr(1, Text1.Text, ".") = 0 Then
      Exit Sub
   End If
   
   'ignore other characters
   KeyAscii = 0

End Sub




