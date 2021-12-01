VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ViewSales 
   BackColor       =   &H8000000C&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sales Order Items"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   Icon            =   "SalesSrch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Double Click To Insert Part Number Or Select And Press Enter"
      Top             =   600
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order For"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "ViewSales"
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


Private Sub Command1_Click()
   Unload Me
   
End Sub

Private Sub Form_Activate()
   If bOnLoad Then
      GetSalesOrders
      bOnLoad = 0
   End If
   
End Sub

Private Sub Form_DblClick()
   Unload Me
   
End Sub

Private Sub Form_Deactivate()
   On Error Resume Next
   MDISect.ActiveForm.optSle.value = vbUnchecked
   Unload Me
   
End Sub

Private Sub Form_Initialize()
   BackColor = ES_ViewBackColor
   
End Sub

Private Sub Form_Load()
   On Error Resume Next
   If MDISect.SideBar.Visible = False Then
      Move MDISect.Left + MDISect.ActiveForm.Left + 800, MDISect.Top + 3200
   Else
      Move MDISect.Left + MDISect.ActiveForm.Left + 2600, MDISect.Top + 3600
   End If
   With Grid1
      .Rows = 2
      .ColWidth(0) = 900
      .ColWidth(1) = 900
      .ColWidth(2) = 600
      .ColWidth(3) = 1200
      .ColWidth(4) = 900
      .ColWidth(5) = 900
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .Row = 0
      .Col = 0
      .Text = "Sales Order"
      .Col = 1
      .Text = "SO Date "
      .Col = 2
      .Text = "Item "
      .Col = 3
      .Text = "Customer"
      .Col = 4
      .Text = "Quantity"
      .Col = 5
      .Text = "Price"
   End With
   bOnLoad = 1
   
End Sub


Private Sub GetSalesOrders()
   Dim RdoGet As ADODB.Recordset
   Dim iList As Integer
   
   On Error Resume Next
   Grid1.Rows = 1
   Grid1.Row = 1
   sSql = "SELECT ITSO,ITNUMBER,ITREV,ITPART,ITQTY,ITDOLLARS," _
          & "SONUMBER,SOCUST,SODATE FROM SoitTable,SohdTable WHERE " _
          & "(ITSO=SONUMBER AND ITCANCELED=0) AND ITPART='" _
          & Compress(lblPrt) & "' ORDER BY SODATE DESC"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet, ES_FORWARD)
   If bSqlRows Then
      With RdoGet
         Do Until .EOF
            iList = iList + 1
            If iList > 300 Then Exit Do
            Grid1.Rows = iList + 1
            Grid1.Col = 0
            Grid1.Row = iList
            Grid1.Text = "" & Format(!ITSO, "00000")
            Grid1.Col = 1
            Grid1.Text = "" & Format(!SODATE, "mm/dd/yy")
            Grid1.Col = 2
            If Trim(!ITREV) = "" Then
               Grid1.Text = "" & Format(!ITNUMBER, "##0")
            Else
               Grid1.Text = "" & Format(!ITNUMBER, "##0") & "-" & Trim(!ITREV)
            End If
            Grid1.Col = 3
            Grid1.Text = "" & Trim(!SOCUST)
            Grid1.Col = 4
            Grid1.Text = Format(!ITQTY, ES_QuantityDataFormat)
            Grid1.Col = 5
            Grid1.Text = Format(!ITDOLLARS, ES_QuantityDataFormat)
            .MoveNext
         Loop
         ClearResultSet RdoGet
      End With
   End If
   Set RdoGet = Nothing
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error Resume Next
   MDISect.ActiveForm.optSle.value = vbUnchecked
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub



Private Sub Form_Unload(Cancel As Integer)
   Set ViewSales = Nothing
   
End Sub

Private Sub Grid1_DblClick()
   On Error Resume Next
   Grid1.Col = 0
   Unload Me
   
End Sub
