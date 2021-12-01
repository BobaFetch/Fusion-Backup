VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form diaPrvew
   BackColor = &H80000018&
   BorderStyle = 1 'Fixed Single
   Caption = "View Fiscal Periods"
   ClientHeight = 3615
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 3735
   Icon = "diaPrvew.frx":0000
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 3615
   ScaleWidth = 3735
   Begin VB.ComboBox cmbYer
      ForeColor = &H00800000&
      Height = 315
      Left = 1560
      Sorted = -1 'True
      TabIndex = 7
      Tag = "8"
      Top = 120
      Width = 915
   End
   Begin MSFlexGridLib.MSFlexGrid Grd1
      Height = 2295
      Left = 240
      TabIndex = 1
      ToolTipText = "Double Click To Insert Dates"
      Top = 840
      Width = 3255
      _ExtentX = 5741
      _ExtentY = 4048
      _Version = 393216
      Rows = 13
      Cols = 3
      FixedRows = 0
      FixedCols = 0
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 315
      Left = 1320
      TabIndex = 0
      TabStop = 0 'False
      Top = 3240
      Width = 915
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 0
      Top = 3120
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 3615
      FormDesignWidth = 3735
   End
   Begin VB.Line Line1
      X1 = 3480
      X2 = 240
      Y1 = 800
      Y2 = 800
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "End"
      Height = 255
      Index = 3
      Left = 2085
      TabIndex = 6
      Top = 600
      Width = 1215
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Start"
      Height = 255
      Index = 2
      Left = 1035
      TabIndex = 5
      Top = 600
      Width = 855
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Period"
      Height = 255
      Index = 1
      Left = 240
      TabIndex = 4
      Top = 600
      Width = 855
   End
   Begin VB.Label lblYear
      BackStyle = 0 'Transparent
      Height = 255
      Left = 2400
      TabIndex = 3
      Top = 3240
      Visible = 0 'False
      Width = 1215
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Fiscal Year"
      Height = 255
      Index = 0
      Left = 240
      TabIndex = 2
      Top = 120
      Width = 1215
   End
End
Attribute VB_Name = "diaPrvew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bOnLoad As Byte

Private Sub cmbYer_Click()
   lblYear = cmbYer
   GetPeriods
   
End Sub


Private Sub cmbYer_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub Form_Activate()
   If bOnLoad Then
      bOnLoad = False
      FillYears
   End If
   
End Sub

Private Sub Form_Load()
   SetFormSize Me
   
   Grd1.ColWidth(0) = 700
   Grd1.ColWidth(1) = 1100
   Grd1.ColWidth(2) = 1100
   bOnLoad = True
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   WindowState = 1
   Set diaPrvew = Nothing
   
End Sub


Public Sub GetPeriods()
   Dim RdoPer As rdoResultset
   Dim A As Integer
   Dim i As Integer
   Grd1.Clear
   sSql = "SELECT * FROM GlfyTable WHERE FYYEAR=" & Trim(lblYear) & " "
   bSqlRows = GetDataSet(RdoPer, ES_FORWARD)
   If bSqlRows Then
      With RdoPer
         For A = 4 To 28 Step 2
            If i = 12 Then
               If IsNull(.rdoColumns(A)) Then Exit For
            End If
            Grd1.Row = i
            Grd1.Col = 0
            Grd1 = i + 1
            Grd1.Col = 1
            Grd1 = Format(.rdoColumns(A), "mm/dd/yy")
            Grd1.Col = 2
            Grd1 = Format(.rdoColumns(A + 1), "mm/dd/yy")
            i = i + 1
         Next
         .Cancel
      End With
   End If
   Set RdoPer = Nothing
   
End Sub

Public Sub FillYears()
   Dim RdoYer As rdoResultset
   Dim A As Integer
   Dim i As Integer
   sSql = "SELECT DISTINCT FYYEAR FROM GlfyTable "
   bSqlRows = GetDataSet(RdoYer, ES_FORWARD)
   If bSqlRows Then
      With RdoYer
         Do Until .EOF
            cmbYer.AddItem Format(!FYYEAR, "0000")
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoYer = Nothing
   cmbYer = lblYear
   If cmbYer.ListCount > 0 Then GetPeriods
   
End Sub
