VERSION 5.00
Begin VB.Form PackPSe02d 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pack Slip Box Detail"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4980
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPieces 
      Height          =   285
      Index           =   4
      Left            =   3240
      TabIndex        =   14
      ToolTipText     =   "Number of Pieces in this Box"
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox txtPieces 
      Height          =   285
      Index           =   3
      Left            =   3240
      TabIndex        =   12
      ToolTipText     =   "Number of Pieces in this Box"
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtPieces 
      Height          =   285
      Index           =   2
      Left            =   3240
      TabIndex        =   10
      ToolTipText     =   "Number of Pieces in this Box"
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtPieces 
      Height          =   285
      Index           =   1
      Left            =   3240
      TabIndex        =   8
      ToolTipText     =   "Number of Pieces in this Box"
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtPieces 
      Height          =   285
      Index           =   0
      Left            =   3240
      TabIndex        =   6
      ToolTipText     =   "Number of Pieces in this Box"
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   285
      Index           =   4
      Left            =   600
      Picture         =   "PackPSe02d.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Delete Box"
      Top             =   2400
      Width           =   285
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   285
      Index           =   3
      Left            =   600
      Picture         =   "PackPSe02d.frx":005E
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Delete Box"
      Top             =   2040
      Width           =   285
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   285
      Index           =   2
      Left            =   600
      Picture         =   "PackPSe02d.frx":00BC
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Delete Box"
      Top             =   1680
      Width           =   285
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   285
      Index           =   1
      Left            =   600
      Picture         =   "PackPSe02d.frx":011A
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Delete Box"
      Top             =   1320
      Width           =   285
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   285
      Index           =   0
      Left            =   600
      Picture         =   "PackPSe02d.frx":0178
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Delete Box"
      Top             =   960
      Width           =   285
   End
   Begin VB.CommandButton cmdPgDwn 
      Height          =   735
      Left            =   4320
      Picture         =   "PackPSe02d.frx":01D6
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdPgUp 
      Height          =   735
      Left            =   4320
      Picture         =   "PackPSe02d.frx":0269
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox txtWeight 
      Height          =   285
      Index           =   4
      Left            =   1920
      TabIndex        =   13
      Tag             =   "1"
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox txtWeight 
      Height          =   285
      Index           =   3
      Left            =   1920
      TabIndex        =   11
      Tag             =   "1"
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtWeight 
      Height          =   285
      Index           =   2
      Left            =   1920
      TabIndex        =   9
      Tag             =   "1"
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtWeight 
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   7
      Tag             =   "1"
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtWeight 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   5
      Tag             =   "1"
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   360
      Left            =   3840
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Label lblFreight 
      Caption         =   "lblFreight"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "# Pieces"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   27
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblBoxNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   960
      TabIndex        =   21
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblBoxNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   960
      TabIndex        =   20
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblBoxNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   960
      TabIndex        =   19
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblBoxNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   960
      TabIndex        =   18
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label lblBoxNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   960
      TabIndex        =   17
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Weight"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Box Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblPSNumber 
      Caption         =   "lblPSNumber"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblTotalBoxes 
      Caption         =   "lblTotalBoxes"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "PackPSe02d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'4/23/07 CJS New
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown()  As New EsiKeyBd

Dim frmCallingForm As Form
Dim bSomethingChanged As Boolean

Dim Weight() As Currency
Dim Pieces() As Currency
Dim iFirstBoxOnScreen As Integer
Dim iTotalScreens As Integer
Const iRecordsOnScreen As Integer = 5
Dim iTotalBoxes As Integer

Private Sub cmdCan_Click()
    Unload Me
End Sub


Private Sub FormatControls()
    Dim b As Byte
    b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())

    
End Sub

Private Sub cmdDelete_Click(Index As Integer)
    Dim iDelItem As Integer
    Dim i As Integer
    
    
    If MsgBox("You are about to delete a box from the pack slip. This will alter the number" & vbCrLf & " of boxes for this packslip. Are you sure?", vbYesNoCancel, Caption) <> vbYes Then
        Exit Sub
    End If
    
    bSomethingChanged = True
    
    iDelItem = (iFirstBoxOnScreen) + Index
    For i = iDelItem To UBound(Weight) - 1
        Weight(i) = Weight(i + 1)
        Pieces(i) = Pieces(i + 1)
    Next i

    iTotalBoxes = iTotalBoxes - 1
    ReDim Preserve Weight(1 To iTotalBoxes)
    ReDim Preserve Pieces(1 To iTotalBoxes)
    SetupScreenVariables
    LoadScreenGrid
End Sub

Private Sub cmdPgDwn_Click()
    If ((iFirstBoxOnScreen + iRecordsOnScreen) - 1) < iTotalBoxes Then
        iFirstBoxOnScreen = iFirstBoxOnScreen + iRecordsOnScreen
        If iFirstBoxOnScreen > iTotalBoxes Then iFirstBoxOnScreen = 1
        LoadScreenGrid
    End If
End Sub

Private Sub cmdPgUp_Click()
    If iFirstBoxOnScreen > 1 Then
        iFirstBoxOnScreen = iFirstBoxOnScreen - iRecordsOnScreen
        If iFirstBoxOnScreen < 1 Then iFirstBoxOnScreen = 1
        LoadScreenGrid
    End If

End Sub



Private Sub Form_Activate()
    On Error Resume Next
    If bOnLoad Then
        iTotalBoxes = Val(lblTotalBoxes)
        bOnLoad = 0
        SetupBoxTable
        LoadBoxArrays
        bSomethingChanged = False
        If iTotalBoxes = 1 Then
            If Val(Weight(1)) = 0 And Val(lblFreight) > 0 Then Weight(1) = lblFreight
            bSomethingChanged = True
        End If
        iFirstBoxOnScreen = 1
        SetupScreenVariables
        LoadScreenGrid
    End If
    MouseCursor 0
End Sub

Private Sub SetupScreenVariables()
    iTotalScreens = iTotalBoxes / iRecordsOnScreen
    If iTotalBoxes Mod iRecordsOnScreen > 0 Then iTotalScreens = iTotalScreens + 1

End Sub

Private Sub Form_Load()
    Dim fTemp As Form
    Set frmCallingForm = MdiSect.ActiveForm
    
    FormLoad Me
    FormatControls
    bOnLoad = 1
    frmCallingForm.Enabled = False
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveBoxData
    frmCallingForm.Enabled = True
    frmCallingForm.SetFocus
    
End Sub

Private Sub Form_Resize()
    Refresh
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set PackPSe02d = Nothing
   On Error Resume Next
   frmCallingForm.txtBxs.Text = LTrim(str(iTotalBoxes))
   If bSomethingChanged Then frmCallingForm.cbRefreshPS.Value = vbChecked
End Sub

Private Sub SetupBoxTable()
    If Not TableExists("PsibTable") Then
        sSql = "CREATE Table PsibTable (PIBPACKSLIP char(8) NOT NULL, PIBBOXNO int NULL, PIBWEIGHT decimal(7,2) NULL, PIBPIECES decimal(7,2) NULL) ON [PRIMARY]"
        clsADOCon.ExecuteSQL sSql 'rdExecDirect
    Else
        If Not ColumnExists("PsibTable", "PIBPIECES") Then
            sSql = "ALTER TABLE PsibTable ADD PIBPIECES decimal(7,2) NULL"
            clsADOCon.ExecuteSQL sSql 'rdExecDirect
        End If
        
    End If
End Sub

Private Sub SaveBoxData()
    Dim i As Integer
    Dim cGrossWeight As Currency

    cGrossWeight = 0
    
    sSql = "DELETE FROM PsibTable WHERE PIBPACKSLIP='" & lblPSNumber.Caption & "' "
    clsADOCon.ExecuteSQL sSql 'rdExecDirect
    
    For i = LBound(Weight) To UBound(Weight)
        sSql = "INSERT INTO PsibTable (PIBPACKSLIP, PIBBOXNO, PIBWEIGHT, PIBPIECES) VALUES ('" & lblPSNumber.Caption & "'," & LTrim(str(i)) & "," & Weight(i) & "," & Pieces(i) & ")" '
        'Debug.Print sSql
        
        clsADOCon.ExecuteSQL sSql 'rdExecDirect
        cGrossWeight = cGrossWeight + Weight(i)
        
    Next i
    sSql = "UPDATE PshdTable SET PSBOXES=" & LTrim(str(UBound(Weight))) & ", PSGROSSLBS = " & cGrossWeight & " WHERE PSNUMBER='" & lblPSNumber.Caption & "' "
    clsADOCon.ExecuteSQL sSql 'rdExecDirect
    
End Sub

Private Sub LoadBoxArrays()
    Dim RdoBoxes As ADODB.Recordset
    Dim iCurrentBox As Integer
    ReDim Weight(1 To Val(lblTotalBoxes))
    ReDim Pieces(1 To Val(lblTotalBoxes))
    
    For iCurrentBox = 1 To iTotalBoxes
    
        sSql = "SELECT * FROM PsibTable WHERE PIBPACKSLIP='" & lblPSNumber.Caption & "' AND PIBBOXNO=" & LTrim(str(iCurrentBox))
        bSqlRows = clsADOCon.GetDataSet(sSql, RdoBoxes, ES_FORWARD)
        If bSqlRows Then
            Weight(iCurrentBox) = RdoBoxes!PIBWEIGHT
            Pieces(iCurrentBox) = RdoBoxes!PIBPIECES
        Else
            Weight(iCurrentBox) = 0
            Pieces(iCurrentBox) = 0
        End If
        Set RdoBoxes = Nothing
    Next iCurrentBox
    
End Sub



Private Sub LoadScreenGrid()
    Dim i As Integer
    
    For i = 1 To iRecordsOnScreen
        If (i + iFirstBoxOnScreen) - 1 > iTotalBoxes Then
            lblBoxNumber(i - 1).Caption = ""
            txtWeight(i - 1).Text = "0.00"
            lblBoxNumber(i - 1).Visible = False
            txtWeight(i - 1).Visible = False
            cmdDelete(i - 1).Visible = False
            txtPieces(i - 1).Text = "0.00"
            txtPieces(i - 1).Visible = False
        Else
            lblBoxNumber(i - 1).Caption = LTrim(str((iFirstBoxOnScreen + i) - 1))
            txtWeight(i - 1).Text = Format(Weight((iFirstBoxOnScreen + i) - 1), ES_QuantityDataFormat)
            txtPieces(i - 1).Text = Format(Pieces((iFirstBoxOnScreen + i) - 1), ES_QuantityDataFormat)
            lblBoxNumber(i - 1).Visible = True
            txtWeight(i - 1).Visible = True
            cmdDelete(i - 1).Visible = True
            txtPieces(i - 1).Visible = True
        End If
    Next i
    If ((iFirstBoxOnScreen + iRecordsOnScreen) - 1) >= iTotalBoxes Then cmdPgDwn.Enabled = False Else cmdPgDwn.Enabled = True
    If iFirstBoxOnScreen = 1 Then cmdPgUp.Enabled = False Else cmdPgUp.Enabled = True

End Sub


Private Sub txtPieces_Change(Index As Integer)
    bSomethingChanged = True
End Sub

Private Sub txtPieces_LostFocus(Index As Integer)
    txtPieces(Index) = CheckLen(txtPieces(Index), 9)
    txtPieces(Index) = Format(Abs(Val(txtPieces(Index))), ES_QuantityDataFormat)
    Pieces((iFirstBoxOnScreen + Index)) = txtPieces(Index).Text
End Sub

Private Sub txtWeight_Change(Index As Integer)
    bSomethingChanged = True
    
End Sub

Private Sub txtWeight_LostFocus(Index As Integer)
    txtWeight(Index) = CheckLen(txtWeight(Index), 9)
    txtWeight(Index) = Format(Abs(Val(txtWeight(Index))), ES_QuantityDataFormat)
    Weight((iFirstBoxOnScreen + Index)) = txtWeight(Index).Text
End Sub
