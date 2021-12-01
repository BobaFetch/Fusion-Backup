VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form DocuDCe01b 
   BackColor       =   &H80000018&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Document Lists"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "DocuDCe01b.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear All"
      Height          =   315
      Left            =   1320
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Clear all selected document list."
      Top             =   1560
      Width           =   915
   End
   Begin VB.CommandButton cmdSelAll 
      Caption         =   "S&elect All"
      Height          =   315
      Left            =   240
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Select all the document list."
      Top             =   1560
      Width           =   915
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "DocuDCe01b.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   5040
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   4560
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Record And Apply Changes"
      Top             =   840
      Width           =   915
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4560
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5880
      Top             =   4440
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5715
      FormDesignWidth =   5790
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Click The Row To Select A Document List To Update"
      Top             =   1920
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   5106
      _Version        =   393216
      Rows            =   10
      Cols            =   4
      FixedCols       =   0
      ForeColor       =   8404992
      HighLight       =   2
      GridLinesFixed  =   1
      ScrollBars      =   2
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   10
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblRev 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      ToolTipText     =   "Document Sheet"
      Top             =   600
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sheet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Document"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   6
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblSheet 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      ToolTipText     =   "Document Sheet"
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblDocument 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      ToolTipText     =   "Document Name/Number"
      Top             =   600
      Width           =   1950
   End
   Begin VB.Label lblClass 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      ToolTipText     =   "Document Class"
      Top             =   360
      Width           =   1950
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   360
      Picture         =   "DocuDCe01b.frx":0AB8
      Top             =   5040
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   600
      Picture         =   "DocuDCe01b.frx":0E42
      Top             =   5040
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Document Lists Using:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "DocuDCe01b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'3/3/05 New
Option Explicit
Dim bOnLoad As Byte
Dim bChanged As Byte
Dim iTotalDocs As Integer

'ReDim arrColumns(0 To rdoFlds.RowCount - 1) As String
'Dim sDocuments(500, 4) As String
Dim sDocuments() As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmdCan_Click()
   Form_Deactivate
   
End Sub

Private Sub cmdClear_Click()
   Dim iList As Integer
   For iList = 1 To grd.Rows - 1
       grd.col = 0
       grd.row = iList
       ' Only if the part is checked
       If grd.CellPicture = Chkyes.Picture Then
           Set grd.CellPicture = Chkno.Picture
           sDocuments(grd.row, 3) = ""
       End If
   Next
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2104
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub




Private Sub cmdSelAll_Click()
   Dim iList As Integer
   For iList = 1 To grd.Rows - 1
       grd.col = 0
       grd.row = iList
       ' Select all
      Set grd.CellPicture = Chkyes.Picture
      sDocuments(grd.row, 3) = "X"
      ' set the dirty flag
      bChanged = 1
   Next
End Sub

Private Sub cmdUpd_Click()
   If bChanged = 0 Then
      MsgBox "There Have Been No Changes.", _
         vbInformation, Caption
   Else
      UpdateLists
   End If
   
End Sub

Private Sub Form_Activate()
   If bOnLoad = 1 Then
      lblClass = Trim(DocuDCe01a.cboClass)
      lblDocument = Trim(DocuDCe01a.cboDoc)
      lblRev = Trim(DocuDCe01a.cboRev)
      lblSheet = Trim(DocuDCe01a.cboSheet)
      If lblClass = "SPECIFICATIONS" Then z1(3).Visible = False
      FillGrid
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Deactivate()
   Unload Me
   
End Sub

Private Sub Form_Load()
   Move 1000, 1500
   BackColor = Es_HelpBackGroundColor
   FormatControls
   With grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .Rows = 2
      .row = 0
      .col = 0
      .Text = "Update"
      .col = 1
      .Text = "Document List"
      .col = 2
      .Text = "List Rev"
      .col = 3
      .Text = "Document Rev"
      .ColWidth(0) = 800
      .ColWidth(1) = 2300
      .ColWidth(2) = 1100
      .ColWidth(3) = 1200
   End With
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'diaCsitm.optBok.Value = vbUnchecked
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set DocuDCe01b = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub




Private Sub FillGrid()
   Dim RdoGrd As ADODB.Recordset
   Dim iIndex As Integer
'   sSql = "SELECT DISTINCT DLSREF,DLSREV,DLSDOCREF,DLSDOCREV," _
'          & "DLSDOCSHEET,DLSDOCCLASS FROM DlstTable WHERE " _
'          & "(DLSDOCREF='" & Compress(lblDocument) & "' AND " _
'          & "DLSDOCCLASS='" & Trim(lblClass) & "' AND DLSDOCSHEET='" _
'          & Trim(lblSheet) & "') ORDER BY DLSREF,DLSREV"

   sSql = "SELECT DISTINCT DLSREF,DLSREV,DLSDOCREF,DLSDOCREV," _
          & "DLSDOCSHEET,DLSDOCCLASS" & vbCrLf _
          & "FROM DlstTable" & vbCrLf _
          & "WHERE (DLSDOCREF='" & Compress(lblDocument) & "' AND " _
          & "DLSDOCCLASS='" & Trim(lblClass) & "' AND DLSDOCSHEET='" _
          & Trim(lblSheet) & "') ORDER BY DLSREF,DLSREV"
          
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGrd, ES_STATIC)
   If bSqlRows Then
      With RdoGrd
                  
         ' Resize the Arrary for number of documents.
         ReDim sDocuments(0 To RdoGrd.RecordCount, 4)
         Do Until .EOF
            iIndex = iIndex + 1
            ' Remove the restriction
            ' MM Ticket# 42871: Select all option when updating document revisions
            'If iIndex > 498 Then Exit Do
            If iIndex > 1 Then grd.Rows = grd.Rows + 1
            grd.row = iIndex
            grd.col = 0
            Set grd.CellPicture = Chkno.Picture
            sDocuments(iIndex, 3) = ""
            grd.col = 1
            grd.Text = "" & Trim(!DLSREF)
            sDocuments(iIndex, 0) = "" & Trim(!DLSREF)
            grd.col = 2
            grd.Text = "" & Trim(!DLSREV)
            sDocuments(iIndex, 1) = "" & Trim(!DLSREV)
            grd.col = 3
            grd.Text = "" & Trim(!DLSDOCREV)
            sDocuments(iIndex, 2) = "" & Trim(!DLSDOCREV)
            .MoveNext
         Loop
         ClearResultSet RdoGrd
      End With
   End If
   iTotalDocs = iIndex
   If iTotalDocs = 0 Then cmdUpd.Enabled = False
   Set RdoGrd = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillgrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub UpdateLists()
   Dim bResponse As Byte
   Dim iList As Integer
   Dim sMsg As String
   
   For iList = 1 To iTotalDocs
      If sDocuments(iList, 3) = "X" Then bResponse = 1
   Next
   If bResponse = 0 Then
      MsgBox "No Document Lists Have Been Selected.", _
         vbInformation, Caption
      Exit Sub
   End If
   sMsg = "Do You Wish To Update The Selected Document Lists?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      CancelTrans
   Else
      On Error Resume Next
      Err.Clear
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      For iList = 1 To iTotalDocs
         ' only if the part is checked
         If (sDocuments(iList, 3) = "X") Then
            sSql = "UPDATE DlstTable SET " _
                   & "DLSDOCREV='" & Trim(lblRev) & "' " _
                   & "WHERE (DLSREF='" & sDocuments(iList, 0) & "' " _
                   & "AND DLSREV='" & sDocuments(iList, 1) & "' " _
                   & "AND DLSDOCREF='" & Compress(Trim(lblDocument)) & "' " _
                   & "AND DLSDOCCLASS='" & Trim(lblClass) & "')"
            clsADOCon.ExecuteSql sSql 'rdExecDirect
         End If
      Next
      If clsADOCon.ADOErrNum > 0 Then
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         MsgBox "Could Not Update Selections.", _
            vbInformation
      Else
         clsADOCon.CommitTrans
         MsgBox "Selections Were Successfully Updated.", _
            vbInformation
      End If
   End If
End Sub

Private Sub Grd_KeyPress(KeyAscii As Integer)
   On Error Resume Next
   If KeyAscii = 13 Then
      grd.col = 0
      If grd.CellPicture = Chkyes.Picture Then
         Set grd.CellPicture = Chkno.Picture
         sDocuments(grd.row, 3) = ""
      Else
         Set grd.CellPicture = Chkyes.Picture
         sDocuments(grd.row, 3) = "X"
         bChanged = 1
      End If
   End If
   
End Sub

Private Sub Grd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   grd.col = 0
   If grd.CellPicture = Chkyes.Picture Then
      Set grd.CellPicture = Chkno.Picture
      sDocuments(grd.row, 3) = ""
   Else
      Set grd.CellPicture = Chkyes.Picture
      sDocuments(grd.row, 3) = "X"
      bChanged = 1
   End If
   
End Sub
