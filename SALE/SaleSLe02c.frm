VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form SaleSLe02c 
   BackColor       =   &H80000018&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Price Book Parts"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "SaleSLe02c.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLe02c.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   3960
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      ToolTipText     =   "Leading Chars (Or Blank)"
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "S&elect"
      Height          =   315
      Left            =   5040
      TabIndex        =   4
      ToolTipText     =   "Select Parts"
      Top             =   480
      Width           =   915
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   3960
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4335
      FormDesignWidth =   7095
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Click The Row To Select A Partnumber And Book Price"
      Top             =   960
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   5106
      _Version        =   393216
      Rows            =   10
      Cols            =   4
      FixedCols       =   0
      ForeColor       =   8404992
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
   End
   Begin VB.Image Chkyes 
      Height          =   180
      Left            =   360
      Picture         =   "SaleSLe02c.frx":0AB8
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Chkno 
      Height          =   180
      Left            =   600
      Picture         =   "SaleSLe02c.frx":0CEA
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblTot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6120
      TabIndex        =   7
      ToolTipText     =   "Total Count (300 Max)"
      Top             =   480
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
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
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblBook 
      BackStyle       =   0  'Transparent
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
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Numbers For Price Book"
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
Attribute VB_Name = "SaleSLe02c"
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
Dim bOnLoad As Byte
Dim bChanged As Byte

Dim sCustomers(500, 2) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2104
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdSel_Click()
   FillGrid
   
End Sub

Private Sub Form_Activate()
   If bOnLoad = 1 Then
      UpdateColumns
      lblBook = Trim(SaleSLe02a.txtBook)
      FillGrid
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   'Move SaleSLe04a.Left + 800, SaleSLe04a.Top + 1200
   BackColor = Es_HelpBackGroundColor
   FormatControls
   With Grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .Rows = 1
      .Row = 0
      .Col = 0
      .Text = "Part Number"
      .Col = 1
      .Text = "Description"
      .Col = 2
      .Text = "Unit Price"
      .Col = 3
      .Text = "Book Price"
      .ColWidth(0) = 2300
      .ColWidth(1) = 2200
      .ColWidth(2) = 1100
      .ColWidth(3) = 1100
   End With
   bOnLoad = 1
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaleSLe02b.optBok.Value = vbUnchecked
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set SaleSLe02c = Nothing
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillGrid()
   Dim RdoGrd As ADODB.Recordset
   Dim sBookParts As String
   
   On Error Resume Next
   lblBook = Trim(SaleSLe02a.txtBook)
   sBookParts = Compress(txtPrt)
   Grd.Rows = 1
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PAPRICE,PBIREF,PBIPARTREF,PBIPRICE " _
          & "FROM PartTable,PbitTable WHERE (PARTREF=PBIPARTREF) AND (PBIREF='" _
          & Compress(lblBook) & "' AND PARTREF LIKE '" & sBookParts & "%')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGrd, ES_FORWARD)
   If bSqlRows Then
      With RdoGrd
         Do Until .EOF
            Grd.Rows = Grd.Rows + 1
            Grd.Row = Grd.Rows - 1
            Grd.Col = 0
            Grd.Text = "" & Trim(!PARTNUM)
            Grd.Col = 1
            Grd.Text = "" & Trim(!PADESC)
            Grd.Col = 2
            Grd.Text = Format(!PAPRICE, ES_QuantityDataFormat)
            Grd.Col = 3
            Grd.Text = Format(!PBIPRICE, ES_QuantityDataFormat)
            lblTot = Grd.Rows - 1
            If Grd.Rows > 300 Then Exit Do
            .MoveNext
         Loop
         ClearResultSet RdoGrd
      End With
   Else
      MsgBox "There Are No Parts Assigned To This Book.", _
         vbInformation, Caption
   End If
   Set RdoGrd = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillgrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub grd_Click()
   With Grd
      .Col = 0
      SaleSLe02b.cmbPrt = .Text
      .Col = 3
      SaleSLe02b.txtListPrice = .Text
      SaleSLe02b.CalculateDiscount
   End With
   
End Sub


Private Sub UpdateColumns()
   'In the case of nulls
   MouseCursor 13
   On Error Resume Next
   sSql = "UPDATE PartTable SET PAPRICE=0 WHERE PAPRICE IS NULL"
  clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   'In the case of nulls
   sSql = "UPDATE PbitTable SET PBIPRICE=0 WHERE PBIPRICE IS NULL"
  clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
End Sub
