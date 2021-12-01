VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form PurcPRe08b 
   BackColor       =   &H80000018&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manufacturer's Parts"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin VB.Frame z2 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   492
      Left            =   1080
      TabIndex        =   2
      Top             =   960
      Width           =   3372
      Begin VB.OptionButton optAll 
         Caption         =   "All Types"
         Height          =   252
         Left            =   1920
         TabIndex        =   5
         Top             =   120
         Width           =   972
      End
      Begin VB.OptionButton optType5 
         Caption         =   "Type 5"
         Height          =   252
         Left            =   1080
         TabIndex        =   4
         Top             =   120
         Width           =   852
      End
      Begin VB.OptionButton optType4 
         Caption         =   "Type 4"
         Height          =   252
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Value           =   -1  'True
         Width           =   852
      End
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRe08b.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "S&elect"
      Height          =   315
      Index           =   0
      Left            =   4920
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Selects A Maximum Of 300 Items"
      Top             =   600
      Width           =   852
   End
   Begin VB.TextBox txtPart 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Pattern Matches Leading Character >=  (Fills Up To 300 Part Numbers)"
      Top             =   600
      Width           =   3372
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   4920
      TabIndex        =   8
      ToolTipText     =   "Update Changes To The List"
      Top             =   1200
      Width           =   915
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4920
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5760
      Top             =   4680
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5235
      FormDesignWidth =   5925
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   3012
      Left            =   240
      TabIndex        =   7
      ToolTipText     =   "Click The Row To Select/Unselect A Customer (300 Customers Max)"
      Top             =   1680
      Width           =   5508
      _ExtentX        =   9710
      _ExtentY        =   5318
      _Version        =   393216
      Rows            =   10
      Cols            =   4
      FixedCols       =   0
      ForeColor       =   8404992
      HighLight       =   2
      GridLinesFixed  =   1
      ScrollBars      =   2
   End
   Begin VB.Label lblManu 
      Caption         =   "compressed Manufacturer"
      Height          =   252
      Left            =   3000
      TabIndex        =   14
      Top             =   4920
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   288
      Index           =   1
      Left            =   240
      TabIndex        =   13
      Top             =   1080
      Width           =   1512
   End
   Begin VB.Label lblSelected 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected"
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
      Left            =   240
      TabIndex        =   11
      Top             =   4800
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Parts"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   600
      Width           =   1515
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   480
      Picture         =   "PurcPRe08b.frx":07AE
      Top             =   5040
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   240
      Picture         =   "PurcPRe08b.frx":0B38
      Top             =   5040
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Numbers"
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
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "PurcPRe08b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'10/19/06 New
Option Explicit
Dim bOnLoad As Byte
Dim bChanged As Byte

Dim sPartNumbers(400, 3) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmdCan_Click()
   Form_Deactivate
   
End Sub

Private Sub cmdGo_Click(Index As Integer)
   FillGrid
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2104
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdUpd_Click()
   If Grd.Rows > 1 Then UpdateParts _
      Else MsgBox "Nothing To Update.", vbInformation, Caption
   
End Sub

Private Sub Form_Activate()
   If bOnLoad = 1 Then
      Caption = Caption & " - " & PurcPRe08a.cmbMfr
      lblManu = Compress(PurcPRe08a.cmbMfr)
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Deactivate()
   Unload Me
   
End Sub


Private Sub Form_Load()
   Move PurcPRe08a.Left + 800, PurcPRe08a.Top + 1200
   BackColor = Es_HelpBackGroundColor
   FormatControls
   With Grd
      .ColAlignment(1) = 0
      .ColAlignment(2) = 0
      .ColAlignment(3) = 0
      .Rows = 1
      .row = 0
      .Col = 0
      .Text = "Listed"
      .Col = 1
      .Text = "Part Number"
      .Col = 2
      .Text = "Description"
      .Col = 3
      .Text = "Type"
      .ColWidth(0) = 600
      .ColWidth(1) = 2000
      .ColWidth(2) = 2000
      .ColWidth(3) = 700
   End With
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim bResponse As Byte
   If bChanged = 1 Then
      bResponse = MsgBox("Exit Without Saving Changes?", ES_NOQUESTION, Caption)
      If bResponse = vbNo Then
         Cancel = True
         Exit Sub
      End If
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set PurcPRe08b = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   z2.BackColor = BackColor
   optType4.BackColor = BackColor
   optType5.BackColor = BackColor
   optAll.BackColor = BackColor
   
End Sub




Private Sub FillGrid()
   Dim RdoGrd As ADODB.Recordset
   Dim bLength As Byte
   Dim sParts As String
   On Error GoTo DiaErr1
   Grd.Rows = 1
   Erase sPartNumbers
   sParts = Trim(txtPart)
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PAMANUFACTURER " _
          & "FROM PartTable WHERE (PARTREF >='" & Compress(txtPart) & "'"
   If optType4 Then
      sSql = sSql & " AND PALEVEL=4"
   End If
   If optType5 Then
      sSql = sSql & " AND PALEVEL=5"
   End If
   sSql = sSql & ") ORDER BY PARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGrd, ES_FORWARD)
   If bSqlRows Then
      MouseCursor 13
      With RdoGrd
         Do Until .EOF
            Grd.Rows = Grd.Rows + 1
            If Grd.Rows > 300 Then Exit Do
            Grd.row = Grd.Rows - 1
            sPartNumbers(Grd.row, 0) = "" & Trim(!PartRef)
            Grd.Col = 0
            If "" & Trim(!PAMANUFACTURER) = lblManu Then
               Set Grd.CellPicture = Chkyes.Picture
               sPartNumbers(Grd.row, 1) = "X"
            Else
               Set Grd.CellPicture = Chkno.Picture
               sPartNumbers(Grd.row, 1) = ""
            End If
            sPartNumbers(Grd.row, 2) = "" & Trim(!PartNum)
            Grd.Col = 1
            Grd.Text = "" & Trim(!PartNum)
            Grd.Col = 2
            Grd.Text = "" & Trim(!PADESC)
            Grd.Col = 3
            Grd.Text = "" & Trim(!PALEVEL)
            .MoveNext
         Loop
         ClearResultSet RdoGrd
         bChanged = 0
      End With
      Grd.Col = 0
      Grd.row = 1
   End If
   lblSelected = "Selected " & Val(Grd.Rows - 1) & " Customers"
   MouseCursor 0
   Set RdoGrd = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillgrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'9/9/06 Revised Price Book trap

Private Sub UpdateParts()
   Dim iList As Integer
   Dim bResponse As Byte
   Dim sMsg As String
   
   On Error Resume Next
   sMsg = "This Procedure Will Add Or Remove Part Numbers" & vbCr _
          & "From This Manufacturer.  Continue?.."
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      'Clear all
      For iList = 1 To Grd.Rows - 1
         sSql = "UPDATE PartTable SET PAMANUFACTURER='' WHERE " _
                & "PARTREF='" & sPartNumbers(iList, 0) & "'"
         clsADOCon.ExecuteSQL sSql
      Next
      For iList = 1 To Grd.Rows - 1
         If sPartNumbers(iList, 1) = "X" Then
            sSql = "UPDATE PartTable SET PAMANUFACTURER='" & lblManu & "' WHERE " _
                   & "PARTREF='" & sPartNumbers(iList, 0) & "'"
            clsADOCon.ExecuteSQL sSql
         End If
      Next
      MouseCursor 0
      SysMsg "Changes Were Saved.", True
      bChanged = 0
      Unload Me
   Else
      CancelTrans
   End If
   
End Sub



Private Sub grd_KeyPress(KeyAscii As Integer)
   On Error Resume Next
   If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      Grd.Col = 0
      If Grd.CellPicture = Chkyes.Picture Then
         Set Grd.CellPicture = Chkno.Picture
         sPartNumbers(Grd.row, 1) = ""
      Else
         Set Grd.CellPicture = Chkyes.Picture
         sPartNumbers(Grd.row, 1) = "X"
      End If
      bChanged = 1
   End If
   
End Sub

Private Sub grd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   Grd.Col = 0
   If Grd.CellPicture = Chkyes.Picture Then
      Set Grd.CellPicture = Chkno.Picture
      sPartNumbers(Grd.row, 1) = ""
   Else
      Set Grd.CellPicture = Chkyes.Picture
      sPartNumbers(Grd.row, 1) = "X"
   End If
   bChanged = 1
   
End Sub



Private Sub txtPart_LostFocus()
   txtPart = CheckLen(txtPart, 30)
   
End Sub
