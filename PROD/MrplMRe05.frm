VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form MrplMRe05 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter/Revise MRP Part Comments"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   10200
      TabIndex        =   9
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CheckBox chkExisting 
      Caption         =   "Show only parts with existing comments"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
   Begin VB.TextBox txtPrefix 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox txtComment 
      Height          =   1935
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   6480
      Width           =   9735
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   375
      Left            =   8520
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid gridComments 
      Height          =   4695
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   8281
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      AllowBigSelection=   0   'False
      SelectionMode   =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Part Leading characters: "
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Width           =   1785
   End
   Begin VB.Label Label1 
      Caption         =   "Part Number"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblDsc 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   1200
      Width           =   3375
   End
End
Attribute VB_Name = "MrplMRe05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions

Option Explicit

Private Enum COMMENT_COL
    COL_DELETE
    COL_ID
    COL_DATE
    COL_BY
    col_comment
End Enum

Private activated As Boolean

Private Sub chkExisting_Click()
   LoadPartCombo
End Sub

Private Sub cmbPrt_Change()
   GetPartDesc
   RefreshGrid
End Sub

Private Sub cmbPrt_Click()
   GetPartDesc
   RefreshGrid
End Sub

Private Sub cmdCan_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
   Dim txt As String, id As Integer, success As Boolean, msg As String
   txt = Replace(txtComment.Text, "'", "''")
   If txt <> "" Then
      With gridComments
         If .TextMatrix(.RowSel, COL_ID) = "" Then
            sSql = "INSERT INTO MrpPartComments (MrpPart,CreatedOn,CreatedBy,Comment)" & vbCrLf _
               & "VALUES ('" & Compress(cmbPrt.Text) & "', " & vbCrLf _
               & "getdate(), '" & sInitials & "','" & txt & "')"
            msg = ""
         Else
            If IsNumeric(.TextMatrix(.RowSel, COL_ID)) Then
               id = CInt(.TextMatrix(.RowSel, COL_ID))
               sSql = "UPDATE MrpPartComments set Comment = '" & txt & "' where CommentID = " & CStr(id)
               msg = "Comment successfully updated"
            Else
               MsgBox "No comment selected"
               Exit Sub
            End If
         End If
         
         If clsADOCon.ExecuteSql(sSql) Then
            RefreshGrid
            txtComment.Text = ""
            If msg <> "" Then MsgBox msg
         Else
            MsgBox "Update failed"
         End If
      End With
   End If
End Sub

Private Sub Form_Activate()
   cmbPrt = ""
   If Not activated Then
      txtPrefix.SetFocus
      activated = True
   End If
End Sub

Private Sub Form_Load()
   MouseCursor 0
End Sub

Private Sub Form_Resize()
   Refresh
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set MrplMRe05 = Nothing
End Sub



'Private Sub FormatControls()
'   Dim b As Byte
'   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
'End Sub
'
Private Sub txtPrt_Change()
    GetPartDesc
End Sub

Private Sub cmbPrt_LostFocus()
'   GetPartDesc
'   RefreshGrid
End Sub

Private Sub gridComments_Click()
   With gridComments
      If .row > 0 Then
         .RowSel = .row
         If .MouseCol = 0 Then
            Dim id As Integer
            If IsNumeric(.TextMatrix(.row, COL_ID)) Then
               id = .TextMatrix(.row, COL_ID)
               If MsgBox("Delete this comment?", vbYesNo) = vbYes Then
                  sSql = "delete from MrpPartComments where CommentID = " & CStr(id)
                  If clsADOCon.ExecuteSql(sSql) Then
                     RefreshGrid
                     MsgBox "Comment deleted"
                  Else
                     MsgBox "Deletion failed"
                  End If
               End If
               Exit Sub
            End If
         End If
         txtComment = .TextMatrix(.row, col_comment)
         txtComment.SetFocus
      Else
         txtComment = ""
      End If
   End With
End Sub

Private Sub txtPrefix_LostFocus()
   LoadPartCombo
End Sub

Sub GetPartDesc()
   sSql = "select RTRIM(PADESC) AS PADESC from PartTable where PARTREF = '" & Compress(cmbPrt.Text) & "'"
   Dim rdo As ADODB.Recordset
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      Me.lblDsc = rdo!PADESC
   Else
      lblDsc = ""
   End If
End Sub

Private Sub LoadPartCombo()
   cmbPrt.Clear
   lblDsc = ""
   
   'don't allow full search of entire part table
   If chkExisting.Value = 0 And Me.txtPrefix = "" Then Exit Sub
   
   sSql = "select RTRIM(PARTNUM) AS PARTNUM from PartTable" & vbCrLf _
      & "where PARTNUM like '" & txtPrefix & "%'" & vbCrLf
   If chkExisting.Value = 1 Then
      sSql = sSql & "and exists (select 1 from MrpPartComments where MrpPart = PARTREF)" & vbCrLf
   End If
   sSql = sSql & "order by PARTNUM"
   
   Dim rdo As ADODB.Recordset
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      With rdo
         Do Until .EOF
            cmbPrt.AddItem !PartNum
            .MoveNext
         Loop
      End With
   End If
   
End Sub

Private Sub RefreshGrid()
   Dim Key As String, iCol As Integer, iRow As Integer, fld As ADODB.Field

   Key = UCase$(Compress(cmbPrt.Text))
   gridComments.Clear
   gridComments.Rows = 0
    
   Dim rdo As ADODB.Recordset
   sSql = "select 'DEL' as [ADD/DEL], CommentID as ID, CONVERT(varchar(10),CreatedOn,101) as [Date], CreatedBy as [User], Comment from MrpPartComments" & vbCrLf _
      & "where MrpPart = '" & Key & "'" & vbCrLf _
      & "order by CommentID"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo, adUseClient)
   'insert headers
   With gridComments
      .Cols = rdo.Fields.Count
      .Rows = rdo.RecordCount + 2
      .ColWidth(COL_DELETE) = 800
      .ColWidth(COL_ID) = 600
      .ColWidth(COL_DATE) = 1000
      .ColWidth(COL_BY) = 800
      .ColWidth(col_comment) = 6000
      
      
      
      Dim gridWidth As Integer
      gridWidth = 360 'allow for scrollbar
      Dim I As Integer
      For I = 0 To .Cols - 1
          gridWidth = gridWidth + .ColWidth(I)
      Next I
      .Width = gridWidth
      txtComment.Width = gridWidth - 1440
      cmdSave.Left = txtComment.Left + txtComment.Width + 240
      Me.Width = gridWidth + 600
      Me.Height = .Top + .Height + 3000
      
      .FixedRows = 1
      .FixedCols = 0
      
      iCol = 0
      For Each fld In rdo.Fields
         gridComments.TextMatrix(0, iCol) = fld.Name
         iCol = iCol + 1
      Next fld
   
   End With
   
   If bSqlRows Then
        
      With gridComments
        
         iRow = 1
         Do Until rdo.EOF
             iCol = 0
             For Each fld In rdo.Fields
               If IsNull(fld.Value) = True Then
                   gridComments.TextMatrix(iRow, iCol) = vbNullString
               Else
                   gridComments.TextMatrix(iRow, iCol) = fld.Value
               End If
               If iCol = 0 Then
                  .row = iRow
                  .CellAlignment = flexAlignCenterCenter
               End If
               iCol = iCol + 1
             Next fld
             iRow = iRow + 1
             rdo.MoveNext
         Loop
         
         ' place a plus sign in the last row.  click here to create a new comment
         .TextMatrix(.Rows - 1, COL_DELETE) = "ADD"
         .row = .Rows - 1
         .Col = COL_DELETE
         .CellAlignment = flexAlignCenterCenter
         
         
         
         .RowSel = 0
         .row = 0
      End With
    End If
End Sub

