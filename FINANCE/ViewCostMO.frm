VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ViewMOCost 
   Caption         =   "View Detail MO Cost"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   LinkTopic       =   "View MMo Cost"
   ScaleHeight     =   7575
   ScaleWidth      =   13875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5040
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1200
   End
   Begin VB.TextBox txtCompletedThru 
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtCompletedFrom 
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   720
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   5415
      Left            =   360
      TabIndex        =   1
      ToolTipText     =   "Click To Select Or Scroll And Press Enter (Also Page Up And Page Down)"
      Top             =   1920
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   9551
      _Version        =   393216
      Rows            =   3
      Cols            =   12
      FixedRows       =   2
      FixedCols       =   0
      FocusRect       =   2
      ScrollBars      =   2
      MergeCells      =   1
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Recost MO's completed from"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "View Manufacturing Order Cost"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
End
Attribute VB_Name = "ViewMOCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bOnLoad As Byte

Private Sub cmbLvl_KeyPress(KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   If bOnLoad = 1 Then
      
      FillGrid
      bOnLoad = 0
            
   End If
   
End Sub

Private Sub Form_Load()
   AlwaysOnTop hwnd, True
   Dim bByte As Byte
   
   On Error Resume Next
   With Grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 1
      .ColAlignment(6) = 1
      .ColAlignment(7) = 1
      .ColAlignment(8) = 1
      .ColAlignment(9) = 1
      .ColAlignment(10) = 1
      .ColAlignment(11) = 1
      
      .Rows = 1
      .row = 0
      .Col = 0
      .Text = "MO PartNumber"
      .Col = 1
      .Text = "MO Run"
      .Col = 2
      .Text = "Lot Number"
      .Col = 3
      .Text = "InvaQty"
      .Col = 4
      .Text = "Inva UnitCost"
      .Col = 5
      .Text = "Lot UnitCost"
      .Col = 6
      .Text = "Inva TotMatl"
      .Col = 7
      .Text = "Inva TotLabor"
      .Col = 8
      .Text = "Inva TotExp"
      .Col = 9
      .Text = "Lot TotMatl"
      .Col = 10
      .Text = "Lot TotLabor"
      .Col = 11
      .Text = "Lot TotExp"
      
      .ColWidth(0) = 2300
      .ColWidth(1) = 700
      .ColWidth(2) = 1500
      .ColWidth(3) = 800
      .ColWidth(4) = 1000
      .ColWidth(5) = 1000
      .ColWidth(6) = 1000
      .ColWidth(7) = 1000
      .ColWidth(8) = 1000
      .ColWidth(9) = 1000
      .ColWidth(10) = 1000
      .ColWidth(11) = 1000
      
      .ScrollBars = flexScrollBarBoth
      .AllowUserResizing = flexResizeColumns
      
   End With
   bOnLoad = 1
   
End Sub

Function FillGrid() As Integer
   Dim RdoGrd As ADODB.Recordset
   Dim strEmp As String
   
   On Error Resume Next
   Grd.Rows = 1
   On Error GoTo DiaErr1
       
       
   Dim strComFrom, strComThru As String
   strComFrom = Format(txtCompletedFrom, "mm/dd/yyyy")
   strComThru = Format(txtCompletedThru, "mm/dd/yyyy")
   
   sSql = "SELECT INMOPART, INMORUN, LotNumber , " & vbCrLf _
            & " INAQTY,  INAMT, LotUnitCost, INTOTMATL,INTOTLABOR, INTOTEXP," & vbCrLf _
            & " LOTTOTMATL , LOTTOTLABOR, LOTTOTEXP" & vbCrLf _
         & " From RunsTable, INVATABLE, LohdTable, PartTable" & vbCrLf _
         & " Where PartRef = RUNREF And LOTMOPARTREF = RUNREF And LOTMORUNNO = Runno" & vbCrLf _
            & " AND INLOTNUMBER = LotNumber" & vbCrLf _
            & " AND  RUNREF = INMOPART AND RUNNO  = INMORUN" & vbCrLf _
            & " AND RUNCOMPLETE BETWEEN '" & strComFrom & "' AND '" & strComThru & "'" & vbCrLf _
            & " AND RUNSTATUS = 'CL'" & vbCrLf _
            & " AND INTYPE = 6" & vbCrLf _
            & " AND INAMT <> LotUnitCost " & vbCrLf _
         & " ORDER BY 1" & vbCrLf _

    
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGrd, ES_FORWARD)
   If bSqlRows Then
       With RdoGrd
           Do Until .EOF
            Grd.Rows = Grd.Rows + 1
            Grd.row = Grd.Rows - 1
            Grd.Col = 0
            Grd.Text = "" & Trim(!INMOPART)
            Grd.Col = 1
            Grd.Text = "" & Trim(!INMORUN)
            Grd.Col = 2
            Grd.Text = "" & Trim(!lotNumber)
            Grd.Col = 3
            Grd.Text = "" & Trim(!INAQTY)
            Grd.Col = 4
            Grd.Text = "" & Trim(!INAMT)
            Grd.Col = 5
            Grd.Text = "" & Trim(!LotUnitCost)
            Grd.Col = 6
            Grd.Text = "" & Trim(!INTOTMATL)
            Grd.Col = 7
            Grd.Text = "" & Trim(!INTOTLABOR)
            Grd.Col = 8
            Grd.Text = "" & Trim(!INTOTEXP)
            Grd.Col = 9
            Grd.Text = "" & Trim(!LOTTOTMATL)
            Grd.Col = 10
            Grd.Text = "" & Trim(!LOTTOTLABOR)
            Grd.Col = 11
            Grd.Text = "" & Trim(!LOTTOTEXP)
            .MoveNext
         Loop
      ClearResultSet RdoGrd
      End With
   End If
   Set RdoGrd = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "fillgrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


