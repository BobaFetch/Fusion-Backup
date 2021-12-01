VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form MelterLog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Melters Log"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtBars 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      TabIndex        =   6
      ToolTipText     =   "Test Bars"
      Top             =   3360
      Width           =   2085
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   420
      Left            =   7320
      TabIndex        =   9
      ToolTipText     =   "Add A New User"
      Top             =   960
      Width           =   1080
   End
   Begin VB.Frame fraUser 
      Height          =   5775
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   6975
      Begin VB.TextBox txtRejQty 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1560
         TabIndex        =   5
         ToolTipText     =   "Test Bars"
         Top             =   2640
         Width           =   2085
      End
      Begin VB.TextBox txtNotes 
         Height          =   825
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Tag             =   "9"
         Top             =   4200
         Width           =   3435
      End
      Begin VB.ComboBox cmbHeatNum 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "MellterLog.frx":0000
         Left            =   1560
         List            =   "MellterLog.frx":0002
         TabIndex        =   3
         Tag             =   "8"
         ToolTipText     =   "Select User Class From List"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.ComboBox txtDte 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Tag             =   "4"
         Top             =   1310
         Width           =   1335
      End
      Begin VB.TextBox txtMOPartNum 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         TabIndex        =   0
         Tag             =   "2"
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtGCast 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1560
         TabIndex        =   4
         ToolTipText     =   "Enter Heat#"
         Top             =   2140
         Width           =   2055
      End
      Begin VB.TextBox txtMelNum 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1560
         TabIndex        =   7
         ToolTipText     =   "Case Sensitive Max (15) Char"
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox txtMORun 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Tag             =   "2"
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rejected Qty "
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   2745
         Width           =   975
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   21
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Number"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   20
         Top             =   3700
         Width           =   1575
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "MO PartNumber"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Accepted Qty"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   18
         Top             =   2190
         Width           =   1095
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Heat Num"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   17
         Top             =   1720
         Width           =   735
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Test Bar Qty"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   3220
         Width           =   975
      End
      Begin VB.Label a 
         BackStyle       =   0  'Transparent
         Caption         =   "MO Run"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   735
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cast Date"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "MellterLog.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   7320
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   360
      Width           =   1080
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7800
      Top             =   5040
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6270
      FormDesignWidth =   8655
   End
   Begin VB.Label lblUsers 
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "MelterLog"
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
'9/9/05 Corrected opening and closing files
Dim bOnLoad As Byte
Dim iOldrec As Integer
'Dim ParentFrm As frmMain


Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdSave_Click()

   On Error GoTo DiaErr1
   
   Dim bResponse As Byte
   
   Dim strMOPart As String
   Dim strMORun As String
   Dim strCastDate As String
   Dim strHeatNum As String
   
   Dim strGCast As String
   Dim strTestBars As String
   Dim strMelNum As String
   Dim strNotes As String
   Dim strRej As String
   
   Dim lMelterID As Integer
   
   strMOPart = Trim(txtMOPartNum.Text)
   strMORun = Trim(txtMORun.Text)
   strCastDate = Trim(txtDte.Text)
   strHeatNum = Trim(cmbHeatNum.Text)
   strGCast = Trim(txtGCast.Text)
   strTestBars = Trim(txtBars.Text)
   strMelNum = Trim(txtMelNum.Text)
   strNotes = Trim(txtNotes)
   strRej = Trim(txtRejQty.Text)
   
   If ((strMOPart = "") Or (strMORun = "")) Then
      MsgBox ("Please add MO Partnumber and Run. Please enter PartNumber and Run.")
      Exit Sub
   End If
   
   If ((strCastDate = "") Or (strHeatNum = "")) Then
      MsgBox ("The Castdate and Heat Number can't be empty.")
      Exit Sub
   End If
   
   Err.Clear
   
   lMelterID = GetNextMelterID
   
   If (lMelterID <> 0) Then
      
      sSql = "INSERT INTO MeltersLogTable (MELTERLOGID, MELTERDATE, HEATNUM, MOPARTNUM, MORUN, " _
               & " GOODCASTING, TESTBARS,SCRAPQTY, MELTERBY, NOTES) VALUES ('" & CStr(lMelterID) & "','" & strCastDate & "','" _
             & strHeatNum & "','" & strMOPart & "','" & strMORun & "','" _
             & strGCast & "','" & strTestBars & "','" _
             & strRej & "','" & strMelNum & "','" & strNotes & "')"
             
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
      If (Err.Number = 0) Then
         SysMsgBox.msg = "Successfully added MelterLog."
         SysMsgBox.Show vbModal
         If (strGCast <> "") Then
            'frmMain.Activate_Label frmMain.lblCom, True, False
            'frmMain.Activate_Label frmMain.lblRej, True, False
            frmMain.lblCom = strGCast
            frmMain.lblRej = strRej
            frmMain.txtNotes = strNotes
            
            frmMain.cmdCom.Caption = frmMain.lblCom
            frmMain.cmdRej.Caption = frmMain.lblRej
         End If
         
         Unload Me
      Else
         SysMsgBox.msg = "Could not add MelterLog."
         SysMsgBox.Show vbModal
      End If
      
           
   
   End If
   
   Exit Sub
   
DiaErr1:
   sProcName = "cmdSave"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me


End Sub

Private Function GetNextMelterID() As Integer
   Dim RdoRpt As ADODB.Recordset
   On Error GoTo DiaErr1
   
   
   sSql = "SELECT ISNULL(MAX(MELTERLOGID),0) + 1 MaxID FROM MeltersLogTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpt, ES_FORWARD)
   If bSqlRows Then
      
      With RdoRpt
         GetNextMelterID = Trim(!MaxID)
         ClearResultSet RdoRpt
      End With
   Else
      GetNextMelterID = 0
   End If
   Set RdoRpt = Nothing
   MouseCursor 0
   Exit Function
   
DiaErr1:
   sProcName = "GetNextMelterID"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description

End Function

Private Sub Form_Load()
   CenterForm Me
   txtDte.Text = Format(Now(), "mm/dd/yy")
   FillHeatNum
   bOnLoad = 1
End Sub

Private Sub FillHeatNum()

   ' Add Heat number
   cmbHeatNum.AddItem "P1-1"
   cmbHeatNum.AddItem "P1-2"
   cmbHeatNum.AddItem "P1-3"
   cmbHeatNum.AddItem "P2-1"
   cmbHeatNum.AddItem "P2-2"
   cmbHeatNum.AddItem "P2-3"
   cmbHeatNum.AddItem "P3-1"
   cmbHeatNum.AddItem "P3-2"
   cmbHeatNum.AddItem "P3-3"
   cmbHeatNum.AddItem "S1-1"
   cmbHeatNum.AddItem "S1-2"
   cmbHeatNum.AddItem "S1-3"
   cmbHeatNum.AddItem "S2-1"
   cmbHeatNum.AddItem "S2-2"
   cmbHeatNum.AddItem "S2-3"
   cmbHeatNum.AddItem "S3-1"
   cmbHeatNum.AddItem "S3-2"
   cmbHeatNum.AddItem "S3-2"
   cmbHeatNum.AddItem "S4-1"
   cmbHeatNum.AddItem "S4-2"
   cmbHeatNum.AddItem "S4-3"
   cmbHeatNum.AddItem "S5-1"
   cmbHeatNum.AddItem "S5-2"
   cmbHeatNum.AddItem "S5-3"
   cmbHeatNum.AddItem "S6-1"
   cmbHeatNum.AddItem "S6-2"
   cmbHeatNum.AddItem "S6-3"
   
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set MelterLog = Nothing
End Sub

Private Sub FormatControls()
   'Dim b As Byte
   'b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub txtDte_DropDown()
   'ShowCalendar Me
End Sub

