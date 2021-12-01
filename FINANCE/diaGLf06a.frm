VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form diaGLf06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open Close GL Accounting Periods"
   ClientHeight    =   4905
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   5160
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4905
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3735
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6588
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
   End
   Begin VB.ComboBox cmbYer 
      Height          =   315
      Left            =   960
      TabIndex        =   2
      Tag             =   "1"
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   4080
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   915
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaGLf06a.frx":0000
      PictureDn       =   "diaGLf06a.frx":0146
   End
   Begin VB.Image imgdInc 
      Height          =   180
      Left            =   360
      Picture         =   "diaGLf06a.frx":028C
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgInc 
      Height          =   180
      Left            =   720
      Picture         =   "diaGLf06a.frx":02E3
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periods"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   555
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      Height          =   255
      Index           =   16
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   555
   End
End
Attribute VB_Name = "diaGLf06a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'*********************************************************************************
' diaGLf06a.frm - Open Close GL Fiscal Periods
'
' Notes: Requested by THYPRE
'
' Created: (nth) 04/01/04
' Revisions:
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim iFY As Integer

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Private Sub cmbYer_Click()
   FillGrid
End Sub

Private Sub cmbYer_LostFocus()
   If Not bCancel Then
      FillGrid
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   sCurrForm = Caption
   With Grid1
      .Cols = 5
      .Row = 0
      .Col = 0
      .ColWidth(0) = 700
      .Text = "Period"
      .Col = 1
      .ColWidth(1) = 1000
      .Text = "Start"
      .Col = 2
      .ColWidth(2) = 1000
      .Text = "End"
      .Col = 3
      .ColWidth(3) = 700
      .Text = "Open"
      .Col = 4
      .ColWidth(4) = 700
      .Text = "Closed"
   End With
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaGLf06a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Public Sub FillCombo()
   Dim rdoYr As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT FYYEAR From GlfyTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoYr)
   With rdoYr
      While Not .EOF
         AddComboStr cmbYer.hwnd, Trim(!FYYEAR)
         .MoveNext
      Wend
   End With
   Set rdoYr = Nothing
   If cmbYer.ListCount > 0 Then
      cmbYer.ListIndex = 0
   End If
   Exit Sub
DiaErr1:
   sProcName = "fillcomb"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub FillGrid()
   Dim rdoPer As ADODB.Recordset
   Dim b As Byte
   Dim A As Byte
   On Error GoTo DiaErr1
   If CInt(cmbYer) = iFY Then
      Exit Sub
   End If
   iFY = CInt(cmbYer)
   sSql = "SELECT FYPERIODS,FYPERSTART1,FYPEREND1,FYCLOSED1,FYPERSTART2," _
          & "FYPEREND2,FYCLOSED2,FYPERSTART3,FYPEREND3,FYCLOSED3,FYPERSTART4," _
          & "FYPEREND4,FYCLOSED4,FYPERSTART5,FYPEREND5,FYCLOSED5,FYPERSTART6," _
          & "FYPEREND6,FYCLOSED6,FYPERSTART7,FYPEREND7,FYCLOSED7,FYPERSTART8," _
          & "FYPEREND8,FYCLOSED8,FYPERSTART9,FYPEREND9,FYCLOSED9,FYPERSTART10,FYPEREND10," _
          & "FYCLOSED10,FYPERSTART11,FYPEREND11,FYCLOSED11,FYPERSTART12,FYPEREND12,FYCLOSED12," _
          & "FYPERSTART13,FYPEREND13,FYCLOSED13 FROM GlfyTable WHERE FYYEAR = " & cmbYer
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoPer)
   With rdoPer
      A = 0
      Grid1.Rows = !FYPERIODS + 1
      For b = 1 To !FYPERIODS
         Grid1.Row = b
         Grid1.Col = 0
         Grid1.Text = b
         Grid1.Col = 1
         Grid1.Text = Format(.Fields(A + 1), "mm/dd/yy")
         Grid1.Col = 2
         Grid1.Text = Format(.Fields(A + 2), "mm/dd/yy")
         Grid1.Col = 3
         
         If IsNull(.Fields(A + 3)) Or .Fields(A + 3) = 0 Then
            Grid1.CellPictureAlignment = flexAlignCenterCenter
            Set Grid1.CellPicture = imgInc
            Grid1.Col = 4
            Grid1.CellPictureAlignment = flexAlignCenterCenter
            Set Grid1.CellPicture = imgdInc
         Else
            Grid1.CellPictureAlignment = flexAlignCenterCenter
            Set Grid1.CellPicture = imgdInc
            Grid1.Col = 4
            Grid1.CellPictureAlignment = flexAlignCenterCenter
            Set Grid1.CellPicture = imgInc
         End If
         
         A = A + 3
      Next
   End With
   Grid1.Row = 0
   Grid1.Col = 0
   Set rdoPer = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillgrid"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub Grid1_Click()
   Dim bPeriod As Byte
   Dim bClosed As Byte
   With Grid1
      If .Row > 0 Then
         bPeriod = CByte(.RowSel)
         If .ColSel >= 3 Then
            If .ColSel = 3 Then
               If .CellPicture = imgdInc Then
                  Set .CellPicture = imgInc
                  .Col = 4
                  Set .CellPicture = imgdInc
                  bClosed = 0
               End If
            Else
               If .CellPicture = imgdInc Then
                  Set .CellPicture = imgInc
                  .Col = 3
                  Set .CellPicture = imgdInc
                  bClosed = 1
               End If
            End If
            sSql = "UPDATE GlfyTable SET FYCLOSED" & bPeriod & "=" & bClosed _
                   & " WHERE FYYEAR = " & cmbYer
            clsADOCon.ExecuteSQL sSql
         End If
      End If
   End With
End Sub
