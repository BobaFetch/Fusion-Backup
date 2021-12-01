VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form CyclCYe05 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Parts/Lots to a Cycle Count"
   ClientHeight    =   6180
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   9000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6180
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optAddLot 
      Caption         =   "Option1"
      Height          =   315
      Left            =   240
      TabIndex        =   13
      Top             =   2520
      Width           =   255
   End
   Begin VB.Frame fraAddLot 
      Caption         =   "Add a new available lot to existing part in the cycle count"
      Enabled         =   0   'False
      Height          =   3375
      Left            =   600
      TabIndex        =   10
      Top             =   2520
      Width           =   8115
      Begin VB.CommandButton cmdAddThisLot 
         Caption         =   "Add this lot"
         Height          =   315
         Left            =   3240
         TabIndex        =   19
         Top             =   2880
         Width           =   1455
      End
      Begin VB.ComboBox cboAvailableLots 
         DataSource      =   "rDt1"
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Tag             =   "99"
         ToolTipText     =   "Contains Part Numbers With Lots"
         Top             =   2400
         Width           =   3675
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   1395
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   7600
         _ExtentX        =   13414
         _ExtentY        =   2461
         _Version        =   393216
      End
      Begin VB.ComboBox cboCurrentParts 
         DataSource      =   "rDt1"
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Tag             =   "99"
         ToolTipText     =   "Contains Part Numbers With Lots"
         Top             =   360
         Width           =   3675
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Available Lots"
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   18
         Top             =   2460
         Width           =   1635
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lots for this part currently in the cycle count"
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   5355
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Part Number"
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   420
         Width           =   1635
      End
   End
   Begin VB.OptionButton optAddPart 
      Caption         =   "Option1"
      Height          =   315
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   255
   End
   Begin VB.ComboBox cboCycleCountID 
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Tag             =   "9"
      ToolTipText     =   "List Includes Cycle ID's Not Locked Or Completed"
      Top             =   480
      Width           =   2115
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CyclCYe05.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   7620
      TabIndex        =   2
      Top             =   480
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6840
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.Frame fraAddNewPart 
      Caption         =   "Add a new part and its nonzero lots to the cycle count"
      Enabled         =   0   'False
      Height          =   1275
      Left            =   600
      TabIndex        =   6
      Top             =   1080
      Width           =   8115
      Begin VB.CommandButton cmdAddPart 
         Caption         =   "Add this part"
         Height          =   315
         Left            =   2940
         TabIndex        =   14
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox cboPart 
         DataSource      =   "rDt1"
         Height          =   315
         Left            =   1860
         TabIndex        =   8
         Tag             =   "99"
         ToolTipText     =   "Contains Part Numbers With Lots"
         Top             =   360
         Width           =   3675
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Part Number"
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   9
         Top             =   420
         Width           =   1635
      End
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Locked, unreconciled)"
      Height          =   285
      Index           =   6
      Left            =   4140
      TabIndex        =   5
      Top             =   540
      Width           =   1935
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cycle Count ID"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   540
      Width           =   1335
   End
End
Attribute VB_Name = "CyclCYe05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'3/1/05 Changed date handling
'5/16/05 corrected group show/hide
'9/15/05 Added Inventory Transfer to report table (32)
Option Explicit
Dim bOnLoad As Byte

Dim iProg As Integer

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cboCurrentParts_Click()
   FillCurrentLots
   FillAvailableLots
End Sub

Private Sub FillCurrentLots()
   With Grid1
      .FixedRows = 0
      Grid1.Clear
      
      .RowHeightMin = 255
      '.FixedRows = 1
      .Rows = 1
      .FixedCols = 0
      .Cols = 4
      
      .Col = 0
      .Text = "System Lot ID"
      .ColWidth(.Col) = 2500
      .ColAlignment(.Col) = flexAlignLeftCenter
      
      .Col = 1
      .Text = "User Lot ID"
      .ColWidth(.Col) = 2500
      .ColAlignment(.Col) = flexAlignLeftCenter
      
      .Col = 2
      .Text = "Qty Remaining"
      .ColWidth(.Col) = 1200
      
      .Col = 3
      .Text = "Location"
      .ColWidth(.Col) = 1200
      
      sSql = "select CLLOTNUMBER, LOTUSERLOTID, LOTREMAININGQTY," & vbCrLf _
         & "case when LOTLOCATION = '' then PALOCATION else LOTLOCATION end as LOC" & vbCrLf _
         & "from CcltTable" & vbCrLf _
         & "join LohdTable on CLLOTNUMBER = LOTNUMBER" & vbCrLf _
         & "join PartTable on PARTREF = LOTPARTREF" & vbCrLf _
         & "and CLREF = '" & Me.cboCycleCountID & "'" & vbCrLf _
         & "and CLPARTREF = '" & Compress(Me.cboCurrentParts) & "'" & vbCrLf _
         & "order by CLLOTNUMBER"
      
      Dim rdo As ADODB.Recordset
      If clsADOCon.GetDataSet(sSql, rdo) Then
         Dim sItem As String
         Do While Not rdo.EOF
            sItem = Trim(rdo!CLLOTNUMBER) & Chr(9) & Trim(rdo!LOTUSERLOTID) _
               & Chr(9) & Format(rdo!LOTREMAININGQTY, ES_QuantityDataFormat) _
               & Chr(9) & Trim(rdo!Loc)
            rdo.MoveNext
            Grid1.AddItem sItem
         Loop
         .FixedRows = 1
      End If
      Set rdo = Nothing
   End With
End Sub

Private Sub FillAvailableLots()
   Me.cboAvailableLots.Clear
   sSql = "select RTRIM(LOTNUMBER) + ' (' + RTRIM(LOTUSERLOTID) + ')'" & vbCrLf _
      & "from LohdTable" & vbCrLf _
      & "where LOTPARTREF = '" & Compress(Me.cboCurrentParts) & "'" & vbCrLf _
      & "and LOTAVAILABLE = 1" & vbCrLf _
      & "and not exists (select CLLOTNUMBER from CcltTable" & vbCrLf _
      & "where CLLOTNUMBER = LOTNUMBER" & vbCrLf _
      & "and CLREF = '" & Me.cboCycleCountID & "'" & vbCrLf _
      & "and CLPARTREF = '" & Compress(Me.cboCurrentParts) & "')" & vbCrLf _
      & "order by LOTNUMBER"
   
   'Dim rdo As rdoResultset
   'If GetDataSet(rdo) Then
      LoadComboBox Me.cboAvailableLots, -1
   'End If
End Sub

Private Sub cboCycleCountID_Click()
   If Me.fraAddLot.Enabled Then
      FillCurrentLots
      FillAvailableLots
   ElseIf Me.fraAddNewPart.Enabled Then
      cboPart.Clear
   End If
End Sub

Private Sub cboPart_DropDown()
   
   ' if part exists in list, don't repopulate
   If cboPart.ListIndex <> -1 Then
      Exit Sub
   End If
   
   sSql = "select PARTNUM" & vbCrLf _
      & "from PartTable" & vbCrLf _
      & "left join LohdTable on PARTREF = LOTPARTREF " & vbCrLf _
      & "where PARTREF like '" & Compress(Me.cboPart) & "%'" & vbCrLf _
      & "and PAABC = (select CCABCCODE from CchdTable" & vbCrLf _
      & "where CCREF = '" & Me.cboCycleCountID & "')" & vbCrLf _
      & "and PALEVEL < 5" & vbCrLf _
      & "and PARTREF not in (select CIPARTREF from CcitTable" & vbCrLf _
      & "where CIREF = '" & cboCycleCountID & "')" & vbCrLf _
      & "And LOTAVAILABLE = 1" & vbCrLf _
      & "and LOTREMAININGQTY > 0" & vbCrLf _
      & "order by PARTNUM"

Debug.Print sSql

   LoadComboBox cboPart, -1
End Sub

Private Sub cboPart_KeyPress(KeyAscii As Integer)
   'if backspace or space, go to first entry (ALL)
   If (KeyAscii = 8 Or KeyAscii = 32) And cboPart.ListCount > 0 Then
      cboPart.ListIndex = 0
      cboPart.Text = cboPart.List(0)
   End If
End Sub

Private Sub cmdAddPart_Click()
   'make sure part exists, but is not on the list already
   Dim rdo As ADODB.Recordset
   sSql = "select PARTREF from PartTable where PARTREF = '" & Compress(Me.cboPart) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo)
   If bSqlRows = False Then
      MsgBox "Part " & cboPart & " does not exist."
      Set rdo = Nothing
      Exit Sub
   End If
   
   sSql = "select CIPARTREF from CcitTable where CIPARTREF = '" & Compress(Me.cboPart) & "'" & vbCrLf _
      & "and CIREF = '" & Me.cboCycleCountID & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo)
   If bSqlRows = True Then
      MsgBox "Part " & cboPart & " is already on this cycle count"
      Set rdo = Nothing
      Exit Sub
   End If
   
   'also make sure the part is not on any other locked cycle count
   sSql = "select CIREF" & vbCrLf _
      & "from CcitTable" & vbCrLf _
      & "join CchdTable on CCREF = CIREF" & vbCrLf _
      & "where CIREF <> '" & Me.cboCycleCountID & "'" & vbCrLf _
      & "and CIPARTREF = '" & Compress(Me.cboPart) & "'" & vbCrLf _
      & "and CCCOUNTLOCKED = 1" & vbCrLf _
      & "and CCUPDATED = 0"
      
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo)
   If bSqlRows = True Then
      MsgBox "Part " & cboPart & " is already on locked cycle count " & Trim(rdo!CIREF)
      Set rdo = Nothing
      Exit Sub
   End If
   
   
   clsADOCon.BeginTrans
   
   sSql = "insert into CcitTable" & vbCrLf _
      & "(CIREF,CIPARTREF,CILOTLOCATION,CIPADESC,CIPALOCATION,CIPASTDCOST,CIPAQOH,CILOTTRACK)" & vbCrLf _
      & "select DISTINCT '" & cboCycleCountID & "'," & vbCrLf _
      & "'" & Compress(cboPart) & "',LOTLOCATION," & vbCrLf _
      & "PADESC,PALOCATION,PASTDCOST,PAQOH,PALOTTRACK" & vbCrLf _
      & "from PartTable" & vbCrLf _
      & "left join LohdTable on PARTREF = LOTPARTREF" & vbCrLf _
      & "where PARTREF = '" & Compress(cboPart) & "'" & vbCrLf _
      & "and LOTAVAILABLE = 1" & vbCrLf _
      & "and LOTREMAININGQTY > 0"
      
   clsADOCon.ExecuteSql sSql
   
   Debug.Print sSql
   'add a CcltTable record.  If no lot found, add an empty CcltTable record
   
   sSql = "insert into CcltTable" & vbCrLf _
      & "(CLREF,CLPARTREF,CLLOTNUMBER,CLLOTREMAININGQTY)" & vbCrLf _
      & "select '" & cboCycleCountID & "'," & vbCrLf _
      & "'" & Compress(cboPart) & "'," & vbCrLf _
      & "isnull(LOTNUMBER,''),isnull(LOTREMAININGQTY,0)" & vbCrLf _
      & "from PartTable" & vbCrLf _
      & "left join LohdTable on PARTREF = LOTPARTREF" & vbCrLf _
      & "where LOTPARTREF = '" & Compress(cboPart) & "'" & vbCrLf _
      & "and LOTAVAILABLE = 1" & vbCrLf _
      & "and LOTREMAININGQTY > 0"
   
   Debug.Print sSql
   
   clsADOCon.ExecuteSql sSql
   
   clsADOCon.CommitTrans
   
   'remove part from dropdown list of available parts
   Dim i As Integer
   Dim NewPart As String
   NewPart = Compress(cboPart)
   For i = 0 To cboPart.ListCount - 1
      If Compress(cboPart.List(i)) = NewPart Then
         cboPart.RemoveItem i
         Exit For
      End If
   Next
   
   'now update the displays below and point them at the new part
   optAddLot.Value = True
   'Me.cboCurrentParts.Text = newPart
   For i = 0 To cboCurrentParts.ListCount - 1
      If Compress(cboCurrentParts.List(i)) = NewPart Then
         cboCurrentParts.ListIndex = i
         Exit For
      End If
   Next
   Set rdo = Nothing
End Sub

Private Sub cmdAddThisLot_Click()
   clsADOCon.BeginTrans
   
'   sSql = "insert into CcitTable" & vbCrLf _
'      & "(CIREF,CIPARTREF,CIPADESC,CIPALOCATION,CIPASTDCOST,CIPAQOH,CILOTTRACK)" & vbCrLf _
'      & "'" & cboCycleCountID & "'," & vbCrLf _
'      & "'" & Compress(cboCurrentParts) & "'," & vbCrLf _
'      & "PADESC,PALOCATION,PASTDCOST,PAQOH,PALOTTRACK" & vbCrLf _
'      & "from PartTable" & vbCrLf _
'      & "where PARTREF = '" & Compress(cboCurrentParts) & "'"
'   RdoCon.Execute sSql
   
'   'if adding a lot, delete possible non-lot CcltTable row
'   sSql = "delete from CcltTable" & vbCrLf _
'      & "where CLREF = '" & cboCycleCountID & "'" & vbCrLf _
'      & "and CLPARTREF = '" & Compress(cboCurrentParts) & "'" & vbCrLf _
'      & "and rtrim(CLLOTNUMBER) = ''"
'   RdoCon.Execute sSql
   
   sSql = "insert into CcltTable" & vbCrLf _
      & "(CLREF,CLPARTREF,CLLOTNUMBER,CLLOTREMAININGQTY)" & vbCrLf _
      & "select '" & cboCycleCountID & "'," & vbCrLf _
      & "'" & Compress(cboCurrentParts) & "'," & vbCrLf _
      & "LOTNUMBER,LOTREMAININGQTY" & vbCrLf _
      & "from LohdTable" & vbCrLf _
      & "where LOTNUMBER = '" & GetLotIdFromCombo & "'"
   clsADOCon.ExecuteSql sSql
   
   'if adding a lot, delete possible non-lot CcltTable row
   If clsADOCon.RowsAffected <> 0 Then
      sSql = "delete from CcltTable" & vbCrLf _
         & "where CLREF = '" & cboCycleCountID & "'" & vbCrLf _
         & "and CLPARTREF = '" & Compress(cboCurrentParts) & "'" & vbCrLf _
         & "and rtrim(CLLOTNUMBER) = ''"
      clsADOCon.ExecuteSql sSql
   End If

'   'if no lots found, add a non-lot CcltTable row back
'   If RdoCon.RowsAffected = 0 Then
'      sSql = "insert into CcltTable(CLREF,CLPARTREF,CLLOTNUMBER,CLLOTREMAININGQTY)" & vbCrLf _
'         & "values('" & cboCycleCountID & "', " _
'         & "'" & Compress(cboCurrentParts) & "', " _
'         & "'', 0)"
'      RdoCon.Execute sSql, rdExecDirect
'   End If
   
   clsADOCon.CommitTrans
   
   Dim msg As String
   msg = "Lot " & cboAvailableLots & " for part " & cboCurrentParts & " added."
   
   'now update the display
   FillCurrentLots
   FillAvailableLots
   
   MsgBox msg
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 907
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub




Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   MouseCursor 0
   If bOnLoad Then
      cboPart.SetFocus
      bOnLoad = 0
'      cboPart.Clear
'      cboPart.ListIndex = 0
   End If
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   bOnLoad = 1
   
   Dim cc As New ClassCycleCount
   cc.PopulateCycleCountCombo cboCycleCountID, 1, -1 ' locked and not reconcilled
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set LotsLTp01a = Nothing
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub SaveOptions()

End Sub

Private Sub GetOptions()

End Sub

Private Sub cboPart_LostFocus()
   cboPart = CheckLen(cboPart, 30)
'   If Trim(cboPart) = "" Then cboPart = "<ALL>"
   
End Sub


Private Sub optAddLot_Click()
   If Me.cboCycleCountID = "" Then
      MsgBox "First, you must select a cycle count"
      Exit Sub
   End If
   Me.fraAddLot.Enabled = True
   Me.fraAddNewPart.Enabled = False
   FillPartsInCycleCount
End Sub

Sub FillPartsInCycleCount()
   If Me.cboCycleCountID = "" Then
      MsgBox "First, you must select a cycle count"
      Exit Sub
   End If
   Me.cboPart.Clear
   sSql = "select PARTNUM" & vbCrLf _
      & "from PartTable" & vbCrLf _
      & "join CcitTable on CIPARTREF = PARTREF" & vbCrLf _
      & "where CIREF = '" & Me.cboCycleCountID & "'" & vbCrLf _
      & "order by PARTREF"
   LoadComboBox Me.cboCurrentParts, -1
End Sub

Private Sub optAddPart_Click()
   Me.fraAddLot.Enabled = False
   Me.fraAddNewPart.Enabled = True
   If Me.cboCycleCountID = "" Then
      MsgBox "First, you must select a cycle count"
      Exit Sub
   End If
   cboPart.Clear
End Sub

Private Function GetLotIdFromCombo() As String
   Dim firstBlank As Integer
   If Len(Me.cboAvailableLots) >= 1 Then
      firstBlank = InStr(1, Me.cboAvailableLots, " ")
      If firstBlank > 1 Then
         GetLotIdFromCombo = Left(Me.cboAvailableLots, firstBlank - 1)
         Exit Function
      End If
   End If
   GetLotIdFromCombo = ""
End Function

