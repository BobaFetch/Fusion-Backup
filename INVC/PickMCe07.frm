VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form PickMCe07 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sheet Reservation and Pick"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11685
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   8280
      TabIndex        =   9
      Tag             =   "4"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtSO 
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
      Left            =   10560
      TabIndex        =   7
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdPick 
      Caption         =   "Pick sheets in selected lot to:"
      Height          =   375
      Left            =   8880
      TabIndex        =   6
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid gridSheets 
      Height          =   4695
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   8281
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      AllowBigSelection=   0   'False
      SelectionMode   =   1
   End
   Begin VB.ComboBox cboParts 
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
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
   Begin VB.TextBox txtLeading 
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
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Index           =   1
      Left            =   7800
      TabIndex        =   10
      Top             =   660
      Width           =   375
   End
   Begin VB.Label lblSO 
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
      Left            =   9960
      TabIndex        =   8
      Top             =   660
      Width           =   495
   End
   Begin VB.Label lblPart 
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
      TabIndex        =   5
      Top             =   600
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "Part Leading Characters"
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
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "PickMCe07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum RESCOL
    COL_CHECKBOX
    COL_LOTUSERID
    COL_HEIGHT
    COL_LENGTH
    COL_COMMENTS
    COL_CERT
    COL_LOC
    COL_STOCK_BY
    COL_STOCK_ON
    COL_RESTK_BY
    COL_RESTK_ON
    COL_RESTK_CANCEL
    COL_RES_BY
    COL_RES_ON
End Enum

'note the funny CHECKED character -- it is not a "p"
Private Const Checked As String = "þ"
Private Const UNCHECKED As String = "q"
Private so As String
    
Private Sub cboParts_Click()
    lblPart = ""
    Dim Key As String
    Key = UCase$(Compress(cboParts.Text))
    If Key = "" Then Exit Sub
    
    'get part description
'    sSql = "select PADESC from PartTable where PARTREF = '" & Key & "'" & vbCrLf _
'        & "and PAPURCONV <> 0 and PAPURCONV <> 1"
    sSql = "select PADESC from PartTable where PARTREF = '" & Key & "'"
    Dim rdo As ADODB.Recordset
    If clsADOCon.GetDataSet(sSql, rdo, adUseClient) Then
        lblPart = rdo("PADESC")
    End If
    
    RefreshGrid
End Sub

Private Sub RefreshGrid()
    Dim Key As String
    Key = UCase$(Compress(cboParts.Text))
    gridSheets.Clear
    gridSheets.Rows = 0
    
    Dim rdo As ADODB.Recordset
    sSql = "select PADESC, RTRIM(LOTUSERLOTID) as [TRACKING #], " & vbCrLf _
        & "ISNULL(cast(LOIHEIGHT as decimal(12,4)),0) as HT, ISNULL(cast(LOILENGTH as decimal(12,4)),0) as LEN," & vbCrLf _
        & "LOTCOMMENTS as COMMENTS, LOTCERT as [CERT / HT #], LOTLOCATION as LOC," & vbCrLf _
        & "LOTUSER as [STK BY], ISNULL(CONVERT(varchar(8),LOIADATE,1),'') as [STK ON]," & vbCrLf _
        & "case when LOISHEETACTTYPE = 'rs' then LOIUSER else '' end as [RSTK BY]," & vbCrLf _
        & "case when LOISHEETACTTYPE = 'rs' then CONVERT(varchar(8), LOIADATE,1) else '' end as [RSTK ON]," & vbCrLf _
        & "case when LOISHEETACTTYPE = 'rs' then 'Cancel' else '' end as RSTK," & vbCrLf _
        & "ISNULL(LOTRESERVEDBY,'') as [RES BY], ISNULL(CONVERT(varchar(8),LOTRESERVEDON,1),'') as [RES ON]" & vbCrLf _
        & "from LohdTable lh" & vbCrLf _
        & "join PartTable p on p.PARTREF = lh.LOTPARTREF" & vbCrLf _
        & "join LoitTable li on lh.LOTNUMBER = li.LOINUMBER" & vbCrLf _
        & "where PARTREF = '" & Key & "' and LOTREMAININGQTY > 0 and LOICLOSED is NULL" & vbCrLf _
        & "ORDER BY lh.LOTUSERLOTID, li.LOIAREA desc"
    bSqlRows = clsADOCon.GetDataSet(sSql, rdo, adUseClient)
    If bSqlRows Then
        Dim iCol As Integer
        Dim iRow As Integer
        Dim fld As ADODB.Field
        
        'insert headers
        With gridSheets
            .Cols = rdo.Fields.count
            .Rows = rdo.RecordCount + 1
            .ColWidth(COL_CHECKBOX) = 400
            .ColWidth(COL_LOTUSERID) = 3000
            .ColWidth(COL_HEIGHT) = 800
            .ColWidth(COL_LENGTH) = 800
            .ColWidth(COL_COMMENTS) = 3000
            .ColWidth(COL_CERT) = 1500
            .ColWidth(COL_LOC) = 700
            .ColWidth(COL_STOCK_BY) = 700
            .ColWidth(COL_STOCK_ON) = 800
            .ColWidth(COL_RESTK_BY) = 750
            .ColWidth(COL_RESTK_ON) = 820
            .ColWidth(COL_RESTK_CANCEL) = 640
            .ColWidth(COL_RES_BY) = 700
            
            Dim gridWidth As Integer
            gridWidth = 360 'allow for scrollbar
            Dim i As Integer
            For i = 0 To .Cols - 1
                gridWidth = gridWidth + .ColWidth(i)
            Next i
            .Width = gridWidth
            Me.Width = gridWidth + 600

            .FixedRows = 1
            
            ' use winddings font for all column 0 (except header)
            For iRow = 0 To rdo.RecordCount
                .row = iRow
                If iRow > 0 Then
                    .Col = 0
                    .FillStyle = flexFillRepeat
                    .CellFontName = "Wingdings"
                    .CellFontSize = 12
                    .CellAlignment = flexAlignCenterCenter
                    .FillStyle = flexFillSingle
                End If
                
                'force lot id to align left.  Sometimes it centers
                .Col = 1
                .CellAlignment = flexAlignLeftCenter
            Next iRow
            
            iCol = 0
            For Each fld In rdo.Fields
                If iCol = COL_CHECKBOX Then
                    'gridSheets.TextMatrix(0, iCol) = CHECKED    'checked box
                    gridSheets.TextMatrix(0, iCol) = "RES"
                Else
                    gridSheets.TextMatrix(0, iCol) = fld.Name
                End If
                iCol = iCol + 1
            Next fld
        
            iRow = 1
            Do Until rdo.EOF
                iCol = 0
                
                ' if a restock, create button to cancel
                If rdo.Fields(COL_RESTK_ON) <> "" Then
                    gridSheets.row = iRow
                    gridSheets.Col = COL_RESTK_CANCEL
                    gridSheets.CellBackColor = vbButtonFace
                End If
                
                For Each fld In rdo.Fields
                    If iCol = 0 Then
                        If rdo(COL_RES_ON) = "" Then
                            gridSheets.TextMatrix(iRow, iCol) = UNCHECKED
                        Else
                            gridSheets.TextMatrix(iRow, iCol) = Checked
                        End If
                        
                    ElseIf iCol = COL_HEIGHT Or iCol = COL_LENGTH Then
                        gridSheets.TextMatrix(iRow, iCol) = Format(fld.Value, "0.0000")
                    Else
                        If IsNull(fld.Value) = True Then
                            gridSheets.TextMatrix(iRow, iCol) = vbNullString
                        Else
                            gridSheets.TextMatrix(iRow, iCol) = fld.Value
                        End If
                    End If
                    iCol = iCol + 1
                Next fld
                iRow = iRow + 1
                rdo.MoveNext
            Loop
            
'            'make everything fit nicely    ' HAVE TO SET BEFORE POPULATING
'            Dim gridWidth As Integer
'            gridWidth = 100   '360 'allow for scrollbar
'            Dim i As Integer
'            For i = 0 To .Cols - 1
'                gridWidth = gridWidth + .ColWidth(i)
'            Next i
'
'            'is there a scrollbar?
'            If Not .RowIsVisible(.Rows - 1) Then
'                gridwWidth = gridWidth + 2000
'            End If
'            Me.Width = gridWidth + 600
'            .Width = gridWidth
'            .Refresh
            'unselect row
            .RowSel = 0
            .row = 0
        End With
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPick_Click()
    PickSheet
End Sub

Private Sub Form_Activate()
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   txtBeg = Format(ES_SYSDATE, "mm/dd/yyyy")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set PickMCe07 = Nothing
End Sub

Private Sub gridSheets_Click()
    With gridSheets
        If .row > 0 Then
            'MsgBox "mousecol = " & CStr(.MouseCol) & " cancel col = " & CStr(COL_RESTK_CANCEL)
            If .MouseCol = 0 Then
                If .TextMatrix(.row, 0) = "" Then
                    'do nothing
                ElseIf .TextMatrix(.row, 0) = UNCHECKED Then
                    'switch to checked box (note the funny character -- it is not a "p"
                    .TextMatrix(.row, 0) = Checked
                    Reserve True
                Else
                    'switch to unchecked box
                    .TextMatrix(.row, 0) = UNCHECKED
                    Reserve False
                End If
            
            ' cancel a restock
            ElseIf .MouseCol = COL_RESTK_CANCEL Then
               Dim rslot As String, so As String
               rslot = Trim(gridSheets.TextMatrix(.row, COL_LOTUSERID))
               
               .Col = COL_RESTK_CANCEL
               .CellBackColor = vbRed
               If MsgBox("Cancel restock of lot " & rslot & "?", vbYesNo) = vbYes Then
                  CancelRestock rslot
                  MsgBox "Restock Canceled"
                  RefreshGrid
               Else
                  .Col = COL_RESTK_CANCEL
                  .CellBackColor = vbButtonFace
               End If
            Else
                'select all rows for this sheet
                '.row = first row
                '.rowsel = last row
                Dim lot As String, firstRow As Integer, lastRow As Integer, row As Integer
                lot = .TextMatrix(.row, COL_LOTUSERID)
                firstRow = -1
                For row = 1 To .Rows - 1
                    If .TextMatrix(row, COL_LOTUSERID) = lot Then
                        If firstRow = -1 Then firstRow = row
                        lastRow = row
                    End If
                Next row
                .row = firstRow
                .Col = COL_LOTUSERID
                .RowSel = lastRow
                .ColSel = COL_RES_ON
            End If
        End If
    End With
End Sub

Private Sub Reserve(TurnOn As Boolean)
'reserve/unreserve a lot
'TurnOn = true to reserve
'       = false to un-reserve

    With gridSheets
        ' update the lot
        Dim lot As String
        Dim row As Integer
        row = .MouseRow
        lot = .TextMatrix(row, COL_LOTUSERID)
        If TurnOn Then
            sSql = "Update LohdTable set LOTRESERVEDBY = '" & sInitials & "', LOTRESERVEDON = '" & txtBeg.Text & "'" & vbCrLf _
                & "where LOTUSERLOTID = '" & lot & "'"
        Else
            sSql = "Update LohdTable set LOTRESERVEDBY = NULL, LOTRESERVEDON = NULL" & vbCrLf _
                & "where LOTUSERLOTID = '" & lot & "'"
        End If

        clsADOCon.ExecuteSql sSql
        
    End With
    
    RefreshGrid
    

End Sub

Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
End Sub

'Private Sub gridSheets_SelChange()
'    If gridSheets.row - gridSheets.RowSel <> 0 Then
'       'User selected more than one row
'       'So Make the row and selected row the same
'       gridSheets.row = gridSheets.RowSel
'
'       'To get highlight you must set focus to the control then back to whatever else
'       gridSheets.SetFocus
'       'cmdClose.SetFocus
'    End If
'End Sub

Private Sub txtLeading_LostFocus()
    Dim leading As String
    leading = UCase$(Compress(txtLeading.Text))
    cboParts.Clear
    sSql = "select RTRIM(PARTREF) AS PARTREF, RTRIM(PARTNUM) AS PARTNUM" & vbCrLf _
        & "from PartTable where PARTREF like '" & leading & "%' and PAPUNITS = 'SH' and PAUNITS <> PAPUNITS order by PARTNUM"
    Dim rdo As ADODB.Recordset
    bSqlRows = clsADOCon.GetDataSet(sSql, rdo, adUseClient)
    If bSqlRows Then
        'ReDim combodata(rdo.RecordCount)
        With rdo
            Do Until .EOF
                cboParts.AddItem CStr(!PartNum)
                'combodata(cboParts.NewIndex) = !PartRef
                rdo.MoveNext
            Loop
        End With
    End If
End Sub

Private Sub txtSO_Change()
    'allow up to 6 numeric digits only
    Dim textval As String
    Dim cursor_loc As Integer
    textval = txtSO.Text
    If (textval = "") Then Exit Sub
    cursor_loc = txtSO.SelStart
    If IsNumeric(textval) And Len(textval) <= 6 Then
      so = textval
    Else
      txtSO.Text = CStr(so)
      txtSO.SelStart = cursor_loc
    End If
End Sub

Private Sub PickSheet()
    Dim row As Integer
    row = gridSheets.RowSel
    If row < 1 Then
        MsgBox "You must click on a sheet to select it for picking"
        Exit Sub
    End If
        
    Dim lot As String, so As String
    lot = Trim(gridSheets.TextMatrix(row, COL_LOTUSERID))

    If IsNumeric(txtSO.Text) Then
        If MsgBox("Pick sheet " & lot & " to SO " & txtSO.Text & "?", vbYesNo) <> vbYes Then Exit Sub
    Else
        MsgBox "SO number required"
        Exit Sub
    End If

    sSql = "exec SheetPick '" & lot & "','" & sInitials & "'," & txtSO.Text & ",'" _
      & Format(txtBeg.Text, "mm/dd/yyyy") & "'"
    clsADOCon.ExecuteSql sSql
    
    RefreshGrid
    MsgBox "Sheets in lot " & lot & " have been picked to SO " & txtSO.Text
    
End Sub

Private Sub CancelRestock(UserLot As String)
   sSql = "exec SheetCancelRestock '" & UserLot & "'"
   clsADOCon.ExecuteSql sSql
End Sub



