VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form DockODf01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancel An On Dock Inspection"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8310
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvPOItems 
      Height          =   2655
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   4683
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CheckBox cbCancelAll 
      Caption         =   "Cancel ALL Items"
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      ToolTipText     =   "Only items that require On Dock with an Inspection Date equal to the one chosen are shown"
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "DockODf01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5160
      TabIndex        =   7
      ToolTipText     =   "Cancel This Inspection"
      Top             =   840
      Width           =   875
   End
   Begin VB.ComboBox cmbDte 
      Height          =   315
      Left            =   3840
      TabIndex        =   4
      Tag             =   "8"
      Top             =   840
      Width           =   1095
   End
   Begin VB.ComboBox cmbPon 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "Contains Qualifying Purchase Orders"
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   7320
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   4200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5460
      FormDesignWidth =   8310
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   " * Only items that require On Dock with an Inspection Date equal to the one chosen are shown"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   5160
      Width           =   8295
   End
   Begin VB.Label lblVendor 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   6
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inspection Date"
      Height          =   285
      Index           =   7
      Left            =   2520
      TabIndex        =   5
      Top             =   885
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Number"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblVendor 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "DockODf01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'12/2/04 New
Option Explicit
Dim bCancel As Byte
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cbCancelAll_Click()
    Dim i As Integer
    
    If cbCancelAll.Value = vbChecked Then
        For i = 1 To lvPOItems.ListItems.Count
            lvPOItems.ListItems(i).Checked = True
        Next i
        lvPOItems.Enabled = False
    Else
        lvPOItems.Enabled = True
    End If
End Sub

Private Sub cmbDte_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub cmbDte_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   If cmbDte.ListCount > 0 Then
      For iList = 0 To cmbDte.ListCount - 1
         If cmbDte = cmbDte.List(iList) Then bByte = 1
      Next
      If bByte = 0 Then
         Beep
         cmbDte = cmbDte.List(0)
      End If
   End If
   
End Sub


Private Sub cmbPon_Click()
   GetPurchaseOrder
   
End Sub


Private Sub cmbPon_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   
   If Len(Trim(cmbPon)) Then
      cmbPon = CheckLen(cmbPon, 6)
      cmbPon = Format(Abs(Val(cmbPon)), "000000")
      If bCancel = 1 Then Exit Sub
      If cmbPon.ListCount > 0 Then
         For iList = 0 To cmbPon.ListCount - 1
            If cmbPon = cmbPon.List(iList) Then bByte = 1
         Next
         If bByte = 0 Then
            Beep
            cmbPon = cmbPon.List(0)
         End If
         GetPurchaseOrder
      End If
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = 1
   
End Sub


Private Sub cmdEnd_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   If Trim(cmbDte) <> "" And lblVendor(1).ForeColor <> ES_RED Then
      sMsg = "You Have Selected To Cancel The On Dock" & vbCr _
             & "Inspection Of PO" & cmbPon & " Dated " & cmbDte & vbCr _
             & "Do You Wish To Continue?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbYes Then CancelInspection Else CancelTrans
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 6450
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub




Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   bOnLoad = 1
   
   lvPOItems.ColumnHeaders.Add , , "Item"
   lvPOItems.ColumnHeaders(1).Width = 750
   
   lvPOItems.ColumnHeaders.Add , , "Rev"
   lvPOItems.ColumnHeaders(2).Width = 650
   
   lvPOItems.ColumnHeaders.Add , , "Part Number"
   lvPOItems.ColumnHeaders(3).Width = 2235

   lvPOItems.ColumnHeaders.Add , , "Part Desc"
   lvPOItems.ColumnHeaders(4).Width = 3645
   
   lvPOItems.ColumnHeaders.Add , , "Qty"
   lvPOItems.ColumnHeaders(5).Width = 675
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set DockODf01a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   cmbDte.ToolTipText = "Contains Inspection Dates. Select From List"
   lblVendor(1).ForeColor = vbBlack
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   cmbPon.Clear
   cmbDte.Clear
   lblVendor(0) = ""
   lblVendor(1) = ""
   sSql = "SELECT DISTINCT PINUMBER FROM PoitTable WHERE " _
          & "(PIONDOCKINSPECTED=1 AND PITYPE=14 OR PITYPE=18) " _
          & " ORDER BY PINUMBER DESC"
   LoadNumComboBox cmbPon, "000000"
   If bSqlRows Then
      cmbPon = cmbPon.List(0)
      GetPurchaseOrder
   Else
      lblVendor(1) = "*** No Qualifying On Dock Inspections Found ***"
      lvPOItems.ListItems.Clear
      MsgBox "No Qualifying On Dock Inspections Where Found.", _
         vbInformation, Caption
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetPurchaseOrder()
   Dim RdoVnd As ADODB.Recordset
   
   Dim iPONumber As Long
   
   On Error GoTo DiaErr1
   cmbDte.Clear
   sSql = "select PONUMBER,POVENDOR,VEREF,VENICKNAME,VEBNAME " _
          & "FROM PohdTable,VndrTable WHERE (POVENDOR=VEREF AND " _
          & "PONUMBER=" & Val(cmbPon) & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVnd, ES_FORWARD)
   If bSqlRows Then
      With RdoVnd
         cmbPon = Format(!poNumber, "000000")
         lblVendor(0) = "" & Trim(!VENICKNAME)
         lblVendor(1) = "" & Trim(!VEBNAME)
         iPONumber = !poNumber
         ClearResultSet RdoVnd
      End With
   Else
      lblVendor(0) = ""
      lblVendor(1) = "*** PO Does Not Qualify ***"
      iPONumber = 0
   End If
   Set RdoVnd = Nothing
   If lblVendor(1).ForeColor <> ES_RED Then GetInspectionDates
   FillPOItems iPONumber
   Exit Sub
   
DiaErr1:
   sProcName = "getpurchaseor"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillPOItems(ByVal poNumber As Double)
    Dim RdoPOItems As ADODB.Recordset
    
    lvPOItems.ListItems.Clear
    lvPOItems.View = lvwReport
    'lvPOItems.ColumnHeaders = True
 

    sSql = "SELECT PIITEM, PIREV, PIPART, PIRUNNO, PARTNUM, PADESC, PIPQTY FROM PoitTable " & _
           " LEFT OUTER JOIN PartTable ON PIPART=PARTREF WHERE PINUMBER=" & poNumber & _
           " AND PIONDOCKINSPECTED=1 AND PITYPE=14 OR PITYPE=18 AND PIONDOCKINSPDATE = '" & cmbDte & "' "
    If poNumber > 0 Then bSqlRows = clsADOCon.GetDataSet(sSql, RdoPOItems, ES_FORWARD)
    If bSqlRows And poNumber > 0 Then
        With RdoPOItems
            Do Until .EOF
                lvPOItems.ListItems.Add , , LTrim(str(!PIITEM))
                lvPOItems.ListItems.item(lvPOItems.ListItems.Count).ListSubItems.Add , , Trim("" & !PIREV)
                lvPOItems.ListItems.item(lvPOItems.ListItems.Count).ListSubItems.Add , , Trim("" & !PartNum)
                lvPOItems.ListItems.item(lvPOItems.ListItems.Count).ListSubItems.Add , , Trim("" & !PADESC)
                lvPOItems.ListItems.item(lvPOItems.ListItems.Count).ListSubItems.Add , , Format(!PIPQTY, "#,###.00")
                lvPOItems.ListItems(lvPOItems.ListItems.Count).Checked = True
                .MoveNext
            Loop
            ClearResultSet RdoPOItems
        End With
        lvPOItems.Enabled = False
        cbCancelAll.Enabled = True
        cbCancelAll.Value = vbChecked
    Else
      cbCancelAll.Value = vbUnchecked
      cbCancelAll.Enabled = False
      lvPOItems.Enabled = False
    End If
    Set RdoPOItems = Nothing
    
End Sub

Private Sub lblVendor_Change(Index As Integer)
   If Left(lblVendor(1), 6) = "*** PO" Then _
           lblVendor(1).ForeColor = ES_RED Else lblVendor(1).ForeColor = vbBlack
   
End Sub


Private Sub GetInspectionDates()
   cmbDte.Clear
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT PINUMBER,PIONDOCKINSPDATE " _
          & "FROM PoitTable WHERE (PIONDOCKINSPECTED=1 AND " _
          & "PIONDOCKINSPDATE IS NOT NULL AND PITYPE=14 OR " _
          & "PITYPE=18) AND PINUMBER=" & Val(cmbPon) & ""
   LoadNumComboBox cmbDte, "mm/dd/yy", 1
   If bSqlRows Then
      cmbDte = cmbDte.List(0)
      cmdEnd.Enabled = True
   Else
      cmdEnd.Enabled = False
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getinspectionda"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub CancelInspection()
   Dim RdoIns As ADODB.Recordset
   
   Dim iList As Integer
   Dim iRows As Integer
   Dim sPoNumber() As String
   Dim sPoItem() As String
   Dim sPoRev() As String
   
   cmdEnd.Enabled = False
   On Error GoTo DiaErr1
   sSql = "SELECT PINUMBER,PIITEM,PIREV,PITYPE,PIONDOCKINSPECTED," _
          & "PIONDOCKINSPDATE FROM PoitTable WHERE (PIONDOCKINSPECTED=1" _
          & "AND PIONDOCKINSPDATE IS NOT NULL AND PITYPE=14 OR PITYPE=18) " _
          & "AND PINUMBER=" & Val(cmbPon) & " AND PIONDOCKINSPDATE='" & cmbDte & "'"
   If cbCancelAll.Value = vbUnchecked Then sSql = sSql & BuildAdditionalWhereClause
   
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoIns, ES_FORWARD)
   If bSqlRows Then
      With RdoIns
         Do Until .EOF
            iRows = iRows + 1
            ReDim Preserve sPoNumber(iRows)
            ReDim Preserve sPoItem(iRows)
            ReDim Preserve sPoRev(iRows)
            sPoNumber(iRows) = Format$(!PINUMBER)
            sPoItem(iRows) = Format$(!PIITEM)
            sPoRev(iRows) = "" & Trim(!PIREV)
            .MoveNext
         Loop
         ClearResultSet RdoIns
      End With
   End If
   If iRows > 0 Then
      On Error Resume Next
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      For iList = 1 To iRows
         sSql = "UPDATE PoitTable SET PITYPE=14,PIONDOCKINSPECTED=0," _
                & "PIONDOCKINSPDATE=Null,PIONDOCKQTYACC=0," _
                & "PIONDOCKQTYREJ=0,PIONDOCKINSPECTOR=''," _
                & "PIONDOCKCOMMENT='' WHERE (PINUMBER=" _
                & sPoNumber(iList) & " AND PIITEM=" & sPoItem(iList) _
                & " AND PIREV='" & sPoRev(iList) & " ')"
         clsADOCon.ExecuteSQL sSql
      Next
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         MsgBox iRows & " PO Items On Dock Inspections Were Canceled.", _
            vbInformation, Caption
      Else
         clsADOCon.RollbackTrans
         MsgBox "Could Not Successfully Cancel The Inspection.", _
            vbExclamation, Caption
      End If
   End If
   Erase sPoNumber
   Erase sPoItem
   Erase sPoRev
   Set RdoIns = Nothing
   FillCombo
   Exit Sub
   
DiaErr1:
   sProcName = "cancelinspect"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Function BuildAdditionalWhereClause() As String
    Dim i As Integer
    Dim sWhereClause As String
    
    'bFoundOne = False
    sWhereClause = ""
    
    For i = 1 To lvPOItems.ListItems.Count
        If lvPOItems.ListItems(i).Checked Then
            If Len(RTrim(sWhereClause)) = 0 Then sWhereClause = " AND ("
            sWhereClause = sWhereClause & " (PIITEM=" & lvPOItems.ListItems(i) & " AND PIREV='" & lvPOItems.ListItems(i).SubItems(1) & "') OR "
            'bFoundOne = True
        End If
    Next i
    If Len(RTrim(sWhereClause)) > 0 Then
        sWhereClause = Left(sWhereClause, Len(sWhereClause) - 3)
        sWhereClause = sWhereClause & ") "
    End If
    
    BuildAdditionalWhereClause = sWhereClause
End Function
