VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form MatlMMf03b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Previous Adjusted PQOH"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   Icon            =   "MatlMMf03b.frx":0000
   LinkTopic       =   "MatlMMf03b"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Click The Row To Select A Partnumber to adjust QOH"
      Top             =   240
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   5106
      _Version        =   393216
      Rows            =   10
      Cols            =   5
      FixedCols       =   0
      BackColorSel    =   -2147483640
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6480
      Top             =   3000
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3285
      FormDesignWidth =   6810
   End
End
Attribute VB_Name = "MatlMMf03b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress(4) As New EsiKeyBd
Private txtGotFocus(4) As New EsiKeyBd
Private txtKeyDown(2) As New EsiKeyBd


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub Form_Activate()
    
    If bOnLoad Then
        FillAjustedPAQOH
    End If
  bOnLoad = 0
  MouseCursor 0
  Me.Icon = Nothing
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   bOnLoad = 1
   FormatControls
    
      With grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .ColAlignment(2) = 1
      .Rows = 1
      .Row = 0
      .Col = 0
      .Text = "Part Number"
      .Col = 1
      .Text = "PAQOH"
      .Col = 2
      .Text = "Previous PAQOH"
      
      .ColWidth(0) = 3500
      .ColWidth(1) = 1250
      .ColWidth(2) = 1750
      
   End With
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set MatlMMf03b = Nothing
End Sub

Private Sub FillAjustedPAQOH()
    Dim RdoGrd As ADODB.Recordset
    Dim strPrt As String
    Dim strPartNum As String
    
    grd.Rows = 1
    On Error GoTo DiaErr1
    
    
    sSql = "SELECT PARTNUM, CURPAQOH, PREPAQOH " & _
                " FROM MaintPAQOH " & _
                " ORDER BY PARTNUM"

   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGrd, ES_FORWARD)
   If bSqlRows Then
      With RdoGrd
         Do Until .EOF
            
            grd.Rows = grd.Rows + 1
            grd.Row = grd.Rows - 1
            grd.Col = 0
            grd.Text = "" & Trim(!PartNum)
            grd.Col = 1
            grd.Text = "" & Trim(!CURPAQOH)
            grd.Col = 2
            grd.Text = "" & Trim(!PREPAQOH)
            
            .MoveNext
         Loop
         ClearResultSet RdoGrd
      End With
   End If
   Set RdoGrd = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "FillAjustedPAQOH"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

