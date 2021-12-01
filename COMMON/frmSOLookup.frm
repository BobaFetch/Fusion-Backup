VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form SOLookup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find a Sales Order"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   7785
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Search by"
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   3735
      Begin VB.OptionButton optSearchBy 
         Caption         =   "Customer"
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optSearchBy 
         Caption         =   "PO Number"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   960
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5741
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Label lblSearchField 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblSOType 
      Caption         =   "Label1"
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblSONumber 
      Caption         =   "Label1"
      Height          =   255
      Left            =   5040
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblControl 
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "SOLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'11/27/2010 Module Created
Option Explicit

Dim bOnLoad As Byte


Private Sub cmdSearch_Click()
    GetSalesOrders
End Sub

Private Sub Form_Activate()
    If bOnLoad = 1 Then
        GetSalesOrders
        bOnLoad = 0
    End If
End Sub

Private Sub Form_Deactivate()
   On Error Resume Next
   MdiSect.ActiveForm.Refresh
   Unload Me
End Sub

Private Sub Form_Load()
    AlwaysOnTop hwnd, True
    bOnLoad = 1
    SetupGrid
    optSearchBy(0).Value = True
    txtSearch.Text = ""
End Sub


Private Sub GetSalesOrders()
   Dim RdoGet As ADODB.Recordset
   Dim iRow As Integer
   Dim sSearchStr As String
   
   On Error Resume Next
   Grid1.Enabled = False
   Grid1.Rows = 1
   Grid1.Row = 1
   sSql = "SELECT SONUMBER, SOTYPE, SODATE, SOPO," & _
           " (SELECT COUNT(*) FROM SoitTable WHERE SoitTable.ITCANCELED=0 AND SoitTable.ITSO = SohdTable.SONUMBER) AS 'SOLineItems'," & _
           " SOCUST, CustTable.CUNAME " & _
           " From dbo.SohdTable " & _
           " INNER JOIN dbo.CustTable ON SohdTable.SOCUST = CustTable.CUREF " & _
           " Where SohdTable.SOCANCELED = 0 "
  ' sSql = sSql & " AND SOTYPE = '" & lblSOType.Caption & "' "
   If optSearchBy(0).Value = True Then sSql = sSql & " AND SOPO LIKE " Else sSql = sSql & " AND CustTable.CUNAME LIKE "
   sSql = sSql & "'" & txtSearch.Text & "%' "
   sSql = sSql & "ORDER BY SONUMBER DESC"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet, ES_FORWARD)
   If bSqlRows Then
      With RdoGet
         Do Until .EOF
            If iRow >= 300 Then Exit Do
            iRow = iRow + 1
            Grid1.Rows = iRow + 1
            Grid1.Row = iRow
            Grid1.Col = 0
            Grid1.Text = Trim(!SOTYPE)
            
            Grid1.Col = 1
            Grid1.Text = Trim(!SoNumber)
            
            Grid1.Col = 2
            Grid1.Text = Format("" & !SODATE, "mm/dd/YY")
            Grid1.Col = 3
            Grid1.Text = Trim("" & !SOPO)
            Grid1.Col = 4
            Grid1.Text = Trim("" & !SOLineItems)
            Grid1.Col = 5
            Grid1.Text = Trim("" & !CUNAME)
            
            .MoveNext
         Loop
         ClearResultSet RdoGet
      End With
      On Error Resume Next
      'lblSelected = iRow
      Grid1.Col = 0
      bOnLoad = 0
   End If
   If Grid1.Rows > 1 Then
      Grid1.Enabled = True
      Grid1.SetFocus
      Grid1.Row = 1
   End If
   Set RdoGet = Nothing

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    AlwaysOnTop hwnd, False
    If lblControl <> "" Then
        If lblControl.Caption = "txtBeg" Then MdiSect.ActiveForm.txtBeg.SetFocus Else MdiSect.ActiveForm.cmbSon.SetFocus
    End If
End Sub


Private Sub SetupGrid()
    Dim i As Integer
    
    Grid1.Cols = 6
    Grid1.Row = 0
    For i = 0 To 5
        Grid1.Col = i
        Grid1.ColAlignment(i) = 0
        Select Case i
            Case 0: Grid1.ColWidth(i) = 500
                    Grid1.Text = "Type"
            Case 1: Grid1.ColWidth(i) = 1000
                    Grid1.Text = "SO Number"
            Case 2: Grid1.ColWidth(i) = 800
                    Grid1.Text = "Date"
            Case 3: Grid1.ColWidth(i) = 1600
                    Grid1.Text = "PO Number"
            Case 4: Grid1.ColWidth(i) = 900
                    Grid1.Text = "Line Items"
            Case 5: Grid1.ColWidth(i) = 3000
                    Grid1.Text = "Customer"
        End Select
        
    
    Next i
    

End Sub


Private Sub Grid1_DblClick()
   Dim bByte As Byte
   On Error Resume Next
   If Grid1.Rows > 1 Then
       If lblControl.Caption = "txtBeg" Then
            Grid1.Col = 0
            MdiSect.ActiveForm.lblBeg.Text = Trim(Grid1.Text)
            Grid1.Col = 1
            MdiSect.ActiveForm.txtBeg.Text = Trim(Grid1.Text)
        Else
            Grid1.Col = 0
            MdiSect.ActiveForm.cmbPre = Trim(Grid1.Text)
            Grid1.Col = 1
            MdiSect.ActiveForm.cmbSon = Trim(Grid1.Text)
        End If
        Unload Me
   End If
   Exit Sub
DiaErr1:
   If bByte = 1 Then Unload Me
    
End Sub

Private Sub optSearchBy_Click(Index As Integer)
    If optSearchBy(0).Value = True Then lblSearchField.Caption = "PO Number" Else lblSearchField.Caption = "Customer"
End Sub

