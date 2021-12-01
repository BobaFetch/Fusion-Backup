VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form CyclCYe04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Assign Parts to a Cycle Count"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9915
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtLotEndLoc 
      Height          =   285
      Left            =   3960
      MaxLength       =   4
      TabIndex        =   8
      Tag             =   "3"
      Top             =   1920
      Width           =   675
   End
   Begin VB.TextBox txtLotStartLoc 
      Height          =   285
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   7
      Tag             =   "3"
      Top             =   1920
      Width           =   675
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "CyclCYe04.frx":0000
      Height          =   350
      Left            =   8040
      Picture         =   "CyclCYe04.frx":04DA
      Style           =   1  'Graphical
      TabIndex        =   39
      TabStop         =   0   'False
      ToolTipText     =   "View Cycle Count Problems"
      Top             =   1380
      Width           =   360
   End
   Begin VB.CommandButton optDis 
      Height          =   330
      Left            =   7620
      Picture         =   "CyclCYe04.frx":09B4
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Display The Report"
      Top             =   2400
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton optPrn 
      Height          =   330
      Left            =   8220
      Picture         =   "CyclCYe04.frx":0B32
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Print The Report"
      Top             =   2400
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtStartLoc 
      Height          =   285
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   5
      Tag             =   "3"
      Top             =   1560
      Width           =   675
   End
   Begin VB.TextBox txtEndLoc 
      Height          =   285
      Left            =   3960
      MaxLength       =   4
      TabIndex        =   6
      Tag             =   "3"
      Top             =   1560
      Width           =   675
   End
   Begin VB.ComboBox cboClass 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Tag             =   "9"
      ToolTipText     =   "Contains Part Numbers With Lots"
      Top             =   2520
      Width           =   855
   End
   Begin VB.ComboBox cboProductCode 
      Height          =   315
      Left            =   4920
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Tag             =   "3"
      ToolTipText     =   "Product Code (Leading Characters Or Blank For All)"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtEnd 
      Height          =   285
      Left            =   3960
      TabIndex        =   4
      Tag             =   "3"
      Top             =   1200
      Width           =   1515
   End
   Begin VB.TextBox txtStart 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Tag             =   "3"
      Top             =   1200
      Width           =   1515
   End
   Begin VB.CommandButton cmdNone 
      Height          =   330
      Left            =   2640
      Picture         =   "CyclCYe04.frx":0CBC
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Deselect All"
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdAll 
      Height          =   330
      Left            =   2040
      Picture         =   "CyclCYe04.frx":0D36
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Select All"
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CyclCYe04.frx":0E48
      Style           =   1  'Graphical
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtPlan 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   5640
      TabIndex        =   1
      Tag             =   "4"
      ToolTipText     =   "Planned Inventory Date"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CheckBox optSaved 
      Enabled         =   0   'False
      Height          =   255
      Left            =   8340
      TabIndex        =   28
      Top             =   3015
      Width           =   375
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin VB.CommandButton cmdLock 
      Caption         =   "L&ock"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6060
      TabIndex        =   15
      ToolTipText     =   "Locks The Current Items And Values (No Further Editing)"
      Top             =   3000
      Width           =   875
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5100
      TabIndex        =   14
      ToolTipText     =   "Saves The Current List As Is. Will Update Quanities And Costs When Edited (Requires Clear To Add Items)"
      Top             =   3000
      Width           =   875
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4140
      TabIndex        =   13
      ToolTipText     =   "Clears The Current List And Removes All Settings"
      Top             =   3000
      Width           =   875
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "S&elect"
      Height          =   315
      Left            =   7800
      TabIndex        =   17
      ToolTipText     =   "Fill The Form With Qualifying Items"
      Top             =   840
      Width           =   875
   End
   Begin VB.Frame z2 
      Height          =   60
      Left            =   240
      TabIndex        =   22
      Top             =   2880
      Width           =   8445
   End
   Begin VB.ComboBox cmbCid 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "List Includes Cycle ID's Not Locked Or Completed"
      Top             =   480
      Width           =   2115
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   7800
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   60
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5760
      Top             =   -60
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   8340
      FormDesignWidth =   9915
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4875
      Left            =   180
      TabIndex        =   30
      Top             =   3360
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   8599
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      AllowBigSelection=   0   'False
      Enabled         =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "and <"
      Height          =   285
      Index           =   14
      Left            =   3420
      TabIndex        =   41
      Top             =   1980
      Width           =   555
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lot Locations >="
      Height          =   285
      Index           =   12
      Left            =   240
      TabIndex        =   40
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Locations >="
      Height          =   285
      Index           =   11
      Left            =   240
      TabIndex        =   36
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "and <"
      Height          =   285
      Index           =   4
      Left            =   3420
      TabIndex        =   35
      Top             =   1620
      Width           =   555
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "and <"
      Height          =   285
      Index           =   2
      Left            =   3420
      TabIndex        =   34
      Top             =   1260
      Width           =   555
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Numbers >="
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   33
      Top             =   1260
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Class(es)"
      Height          =   285
      Index           =   3
      Left            =   360
      TabIndex        =   32
      Top             =   2580
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code(s)"
      Height          =   285
      Index           =   0
      Left            =   3540
      TabIndex        =   31
      Top             =   2580
      Width           =   1275
   End
   Begin VB.Image imgInc 
      Height          =   180
      Left            =   3660
      Picture         =   "CyclCYe04.frx":15F6
      Top             =   3060
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgdInc 
      Height          =   180
      Left            =   3300
      Picture         =   "CyclCYe04.frx":18A8
      Top             =   3060
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Count Saved"
      Height          =   255
      Index           =   13
      Left            =   7020
      TabIndex        =   27
      Top             =   3015
      Width           =   1335
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   1080
      TabIndex        =   26
      ToolTipText     =   "Total Items Included"
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Items"
      Height          =   255
      Index           =   10
      Left            =   360
      TabIndex        =   25
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Items Must Have An Inventory Location To Be Included"
      Height          =   255
      Index           =   9
      Left            =   480
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Retrieve List"
      Height          =   255
      Index           =   8
      Left            =   5400
      TabIndex        =   23
      Top             =   1005
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Planned Date"
      Height          =   255
      Index           =   7
      Left            =   4560
      TabIndex        =   21
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblCabc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4080
      TabIndex        =   20
      ToolTipText     =   "ABC Code Selected"
      Top             =   480
      Width           =   405
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   19
      Top             =   880
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cycle Count ID"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   18
      Top             =   540
      Width           =   1335
   End
End
Attribute VB_Name = "CyclCYe04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'New 2/26/04
Option Explicit
Dim bChanging As Byte
Dim bOnLoad As Byte
Dim bGoodCount As Byte
Dim iCurrIdx As Integer
'Dim iIndex As Integer
'Dim iMaxPages As Integer
'Dim iPage As Integer
'Dim iTotalList As Integer

'grid columns
Private Const COL_Include = 0
Private Const COL_Location = 1
Private Const COL_PartRef = 2
Private Const COL_PartDescription = 3
Private Const COL_StdCost = 4
Private Const COL_QOH = 5
Private Const COL_LotTracked = 6
Private Const COL_LotLocation = 7


'Dim sParts(1000, 5) As String 'Location,PartRef, Number, Description, Lot Tracked
'Dim cValue(1000, 2) As Currency 'Cost, Qoh
'Dim bInclude(1000) As Byte 'Include
'Dim vCycleLots(1000, 2) As Variant 'Lotnumber, Remaining)

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCid_Click()
   bGoodCount = GetCycleCount()
   
End Sub


Private Sub cmbCid_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   cmbCid = Trim(cmbCid)
   If cmbCid.ListCount > 0 Then
      For iList = 0 To cmbCid.ListCount - 1
         If cmbCid.List(iList) = cmbCid Then b = 1
      Next
      If b = 0 Then
         Beep
         cmbCid = cmbCid.List(0)
      End If
      bGoodCount = GetCycleCount()
   End If
   
End Sub



Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdClear_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   If optSaved.Value = vbUnchecked Then
      MsgBox "This Cycle Count Has Not Been Saved (Nothing To Clear).", _
         vbInformation, Caption
      Exit Sub
   Else
      sMsg = "This Function Completely Clears Your Current List." & vbCr _
             & "It Will Be Necessary To Start Over And Re-Select " & vbCr _
             & "Items To Be Included And Counted. Continue?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbNo Then
         CancelTrans
      Else
         DeleteCurrentList
      End If
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "5403"
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdLock_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   If optSaved.Value = vbUnchecked Then
      MsgBox "The Cycle Count Hasn't Been Saved.", _
         vbInformation, Caption
   Else
      sMsg = "This Function Locks The Locations, Costs And Quantities." & vbCr _
             & "It Also Adds Lots For Lot Tracked Parts (If Set)." & vbCr _
             & "Do You Want To Lock This Count And Prepare It " & vbCr _
             & "For Inventory?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbYes Then
         If OkToLock Then
            LockThisCount
         End If
      Else
         CancelTrans
      End If
   End If
   
End Sub

Private Sub cmdAll_Click()
   MarkAll True
End Sub

Private Sub cmdSave_Click()
   Dim iCount As Integer
   Dim iRow As Integer
   
   For iRow = 1 To Grid1.Rows - 1
      Grid1.row = iRow
      Grid1.Col = COL_Include
      If Grid1.CellPicture = imgInc Then iCount = iCount + 1
      'If bInclude(iRow) = 1 Then iCount = iCount + 1
   Next
   If iCount = 0 Then
      MsgBox "There Are No Items Included On Your List.", _
         vbInformation, Caption
   Else
      iRow = MsgBox("You Have Selected " & iCount & " Items For The Count." & vbCr _
             & "After Saving This List, You Must First Clear It To " & vbCr _
             & "Add More Items. Locations, Quantities And Costs " & vbCr _
             & "Will Not Be Fixed. Continue Saving The List?", _
             ES_YESQUESTION, Caption)
      If iRow = vbYes Then
         SaveCurrentList
      Else
         CancelTrans
      End If
   End If
End Sub

Private Sub cmdSel_Click()
   FillList
   
End Sub

Private Sub cmdNone_Click()
   MarkAll False
End Sub

Private Sub cmdVew_Click()
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   On Error GoTo DiaErr1
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "CycleCountID"
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & CStr("Requested By: " & sInitials) & "'")
   aFormulaValue.Add CStr("'" & CStr(Me.cmbCid) & "'")
   sCustomReport = GetCustomReport("cclog")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init

   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.ShowGroupTree False
   
   'sSql = "{LohdTable.LOTNUMBER}='" & lblNumber & "'"
   'cCRViewer.SetReportSelectionFormula (sSql)
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   

   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   MouseCursor 0
   Exit Sub



  ' MouseCursor 13
  ' On Error GoTo DiaErr1
  ' 'SetMdiReportsize MdiSect
  ' MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
  ' MdiSect.Crw.Formulas(1) = "RequestBy = 'Requested By: " & sInitials & "'"
  ' MdiSect.Crw.Formulas(2) = "CycleCountID = '" & Me.cmbCid & "'"
  ' sCustomReport = GetCustomReport("cclog")
  ' MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
  ' 'SetCrystalAction Me
  ' MouseCursor 0
  ' Exit Sub
   
DiaErr1:
   sProcName = "cmdVew_Click"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bChanging = 0
      FillCombo
      Dim prodcode As New ClassProductCode
      prodcode.PopulateProductCodeCombo Me.cboProductCode, True
      Dim partclass As New ClassPartClass
      partclass.PopulatePartClassCombo Me.cboClass, True
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   bOnLoad = 1
   
   With Grid1
      .RowHeightMin = 300
      '.FixedRows = 1
      .Rows = 1
      .FixedCols = 0
      .Cols = 8
      .ColWidth(COL_Include) = 500
      
      .Col = COL_Location
      .Text = "Part Loc"
      .ColWidth(.Col) = 750
      
      .Col = COL_PartRef
      .Text = "Part"
      .ColWidth(.Col) = 1500
      
      .Col = COL_PartDescription
      .Text = "Description"
      .ColWidth(.Col) = 3000
      
      .Col = COL_StdCost
      .Text = "Std Cost"
      .ColWidth(.Col) = 800
      
      .Col = COL_QOH
      .Text = "Qty"
      .ColWidth(.Col) = 1000
      
      .Col = COL_LotTracked
      .Text = "Lot"
      .ColWidth(.Col) = 0
      
      .Col = COL_LotLocation
      .Text = "Lot Loc"
      .ColWidth(.Col) = 950
   End With
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtPlan.ToolTipText = "Planned Inventory Date"
   EmptyList
   
End Sub

Private Sub FillCombo()
'   On Error GoTo DiaErr1
'   cmbCid.Clear
'   sSql = "SELECT CCCOUNTLOCKED,CCREF FROM CchdTable WHERE CCCOUNTLOCKED=0 " _
'          & "ORDER BY CCREF"
'   LoadComboBox cmbCid
'   If cmbCid.ListCount > 0 Then
'      If Trim(cmbCid) = "" Then cmbCid = cmbCid.List(0)
'      ' bGoodCount = GetCycleCount()
'   Else
'      MsgBox "There Are No Unlocked Counts Recorded.", _
'         vbInformation, Caption
'      Unload Me
'   End If
'   Exit Sub
'
'DiaErr1:
'   sProcName = "fillcombo"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
   Dim cc As New ClassCycleCount
   cc.PopulateCycleCountCombo cmbCid, -1, 0
End Sub




Private Sub FillList()
   Dim RdoAbc As ADODB.Recordset
   Dim sItem As String
   bChanging = 0
   Grid1.FixedRows = 0
   On Error GoTo DiaErr1
   lblCount = "0"
   If optSaved.Value = vbUnchecked Then
      sSql = "SELECT DISTINCT PARTREF,PARTNUM,PADESC,PALOCATION,PAQOH,PAABC," & vbCrLf _
         & "PASTDCOST,PALOTTRACK,LOTLOCATION FROM PartTable left outer join lohdTable " & vbCrLf _
         & " ON PARTREF = lotpartref " & vbCrLf _
         & "WHERE PAABC='" & lblCabc & "'" & vbCrLf _
         & "AND PALEVEL<5 AND LOTREMAININGQTY > 0 " & vbCrLf
      
      If Trim(txtStart) <> "" Then
         sSql = sSql & "AND PARTREF >= '" & txtStart & "'" & vbCrLf
      End If
      If Trim(txtEnd) <> "" Then
         sSql = sSql & "AND PARTREF < '" & txtEnd & "'" & vbCrLf
      End If
      
      If Trim(txtStartLoc) <> "" Then
         sSql = sSql & "AND PALOCATION >= '" & txtStartLoc & "'" & vbCrLf
      End If
      If Trim(txtEndLoc) <> "" Then
         sSql = sSql & "AND PALOCATION < '" & txtEndLoc & "'" & vbCrLf
      End If
      
      If Trim(txtLotStartLoc) <> "" Then
         sSql = sSql & "AND LOTLOCATION >= '" & txtLotStartLoc & "'" & vbCrLf
      End If
      If Trim(txtLotEndLoc) <> "" Then
         sSql = sSql & "AND LOTLOCATION < '" & txtLotEndLoc & "'" & vbCrLf
      End If
      
      If Me.cboClass <> "<ALL>" Then
         sSql = sSql & "AND PACLASS = '" & cboClass & "'" & vbCrLf
      End If
      If Me.cboProductCode <> "<ALL>" Then
         sSql = sSql & "AND PAPRODCODE = '" & cboProductCode & "'" & vbCrLf
      End If
         
      sSql = sSql & "ORDER BY PALOCATION,PARTREF"
      
      Debug.Print sSql
      
'MsgBox "Beta display of part selection:" & vbCrLf & sSql
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoAbc, ES_FORWARD)
      If bSqlRows Then
         With RdoAbc
            Do Until .EOF
               'If Grid1.Rows > 25 Then Exit Do
               sItem = Chr(9) & Trim(!PALOCATION) _
                       & Chr(9) & " " & Trim(!PartNum) _
                       & Chr(9) & " " & Trim(!PADESC) _
                       & Chr(9) & Format(!PASTDCOST, "###,###,##0.000") _
                       & Chr(9) & Format(!PAQOH, ES_QuantityDataFormat) _
                       & Chr(9) & !PALOTTRACK _
                       & Chr(9) & !LOTLOCATION
                       
               Grid1.AddItem sItem
               If Grid1.FixedRows <> 1 Then
                  Grid1.FixedRows = 1
               End If
               
               Grid1.row = Grid1.Rows - 1
               Grid1.Col = COL_Include
               Grid1.CellPictureAlignment = flexAlignCenterCenter
               Set Grid1.CellPicture = imgdInc
               .MoveNext
            Loop
            ClearResultSet RdoAbc
         End With
      End If
   Else
      sSql = "SELECT PARTREF,PARTNUM,PADESC,PALOCATION,CILOTLOCATION,PAQOH,PAABC," _
             & "PASTDCOST,PALOTTRACK,CIREF,CIPARTREF FROM PartTable,CcitTable " _
             & "WHERE (CIREF='" & cmbCid & "' AND PARTREF=CIPARTREF) " _
             & "ORDER BY PALOCATION,PARTREF"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoAbc, ES_FORWARD)
      If bSqlRows Then
         With RdoAbc
            Do Until .EOF
               'If Grid1.Rows > 25 Then Exit Do
               sItem = Chr(9) & Trim(!PALOCATION) _
                       & Chr(9) & " " & Trim(!PartNum) _
                       & Chr(9) & " " & Trim(!PADESC) _
                       & Chr(9) & Format(!PASTDCOST, "###,###,##0.000") _
                       & Chr(9) & Format(!PAQOH, ES_QuantityDataFormat) _
                       & Chr(9) & !PALOTTRACK _
                       & Chr(9) & !CILOTLOCATION
                       
               Grid1.AddItem sItem
               If Grid1.FixedRows <> 1 Then
                  Grid1.FixedRows = 1
               End If
               Grid1.row = Grid1.Rows - 1
               Grid1.Col = COL_Include
               Grid1.CellPictureAlignment = flexAlignCenterCenter
               Set Grid1.CellPicture = imgInc
               .MoveNext
            Loop
            ClearResultSet RdoAbc
         End With
      End If
   End If
   If Grid1.Rows > 1 Then
      cmdAll.Enabled = True
      cmdNone.Enabled = True
      cmdClear.Enabled = True
      cmdSave.Enabled = True
      cmdLock.Enabled = True
'      iIndex = 0
'      iPage = 1
'      iMaxPages = iTotalList / 7
'      If iMaxPages > 1 Then
'         cmdDn.Enabled = True
'         cmdDn.Picture = Endn.Picture
'      End If
'      lblPages = iMaxPages
'      lblPage = iPage
'      optInc(1).SetFocus
   Else
'      iIndex = 0
'      iPage = 0
'      lblPage = ""
'      lblPages = "0"
   End If
   lblCount = Grid1.Rows - 1
   If Grid1.Rows = 1 Then MsgBox "No Matching Parts were Found.", _
                   vbInformation, Caption
   Set RdoAbc = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "filllist"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub EmptyList()
   'Dim bRow As Byte
   bChanging = 1
'   For bRow = 1 To 7
'      optInc(bRow).Caption = ""
'      optInc(bRow).Enabled = False
'      optInc(bRow).Value = vbUnchecked
'      lblLoc(bRow) = ""
'      lblPart(bRow) = ""
'      lblCost(bRow) = ""
'      lblQoh(bRow) = ""
'      lblDsc(bRow) = ""
'   Next
   Grid1.Rows = 1
   bChanging = 0
   
End Sub

Private Function GetCycleCount() As Byte
   Dim RdoCid As ADODB.Recordset
   EmptyList
   DisableLower
   On Error GoTo DiaErr1
   sSql = "Qry_GetCycleCountNotLocked '" & Trim(cmbCid) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCid, ES_FORWARD)
   If bSqlRows Then
      With RdoCid
         lblCabc = "" & Trim(!CCABCCODE)
         txtDsc = "" & Trim(!CCDESC)
         txtPlan = Format(!CCPLANDATE, "mm/dd/yy")
         optSaved.Value = !CCCOUNTSAVED
         GetCycleCount = 1
         ClearResultSet RdoCid
      End With
   Else
      GetCycleCount = 0
      MsgBox "That Count ID Wasn't Found Or Is Locked.", _
         vbInformation, Caption
   End If
   Set RdoCid = Nothing
   Exit Function
   
DiaErr1:
   GetCycleCount = 0
   sProcName = "getcycleco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub DisableLower()
'   Erase sParts
'   Erase cValue
'   Erase bInclude
'   cmdDn.Enabled = False
'   cmdDn.Picture = Dsdn.Picture
'   cmdUp.Enabled = False
'   cmdUp.Picture = Dsup.Picture
   cmdAll.Enabled = False
   cmdNone.Enabled = False
   cmdClear.Enabled = False
   cmdSave.Enabled = False
   cmdLock.Enabled = False
   lblCount = 0
   optSaved.Value = vbUnchecked
   
End Sub

Private Sub SaveCurrentList()
   Dim iRow As Integer
   Dim RdoLst As ADODB.Recordset
   
   clsADOCon.BeginTrans
   
   'On Error Resume Next
   'Delete Them
   sSql = "DELETE FROM CcitTable WHERE CIREF='" & Trim(cmbCid) & "'"
   clsADOCon.ExecuteSql sSql
   'Add Them
   sSql = "SELECT * FROM CcitTable WHERE CIREF='" & Trim(cmbCid) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_KEYSET)
   With RdoLst
'      For iRow = 1 To iTotalList
'         If bInclude(iRow) = 1 Then
'            .AddNew
'            !CIREF = Trim(cmbCid)
'            !CIPARTREF = sParts(iRow, 1)
'            !CIPADESC = sParts(iRow, 3)
'            .Update
'         End If
'      Next
      
      For iRow = 1 To Grid1.Rows - 1
         Grid1.row = iRow
         Grid1.Col = COL_Include
         If Grid1.CellPicture = imgInc Then
            .AddNew
            !CIREF = Trim(cmbCid)
            
            Grid1.Col = COL_PartRef
            !CIPARTREF = Compress(Trim(Grid1.Text))
            
            Grid1.Col = COL_PartDescription
            !CIPADESC = Trim(Grid1.Text)
            
            Grid1.Col = COL_LotLocation
            !CILOTLOCATION = Trim(Grid1.Text)
            .Update
         End If
      Next
      
      ClearResultSet RdoLst
   End With
   If Err = 0 Then
      sSql = "UPDATE CchdTable SET CCCOUNTSAVED=1 WHERE " _
             & "CCREF='" & Trim(cmbCid) & "'"
      clsADOCon.ExecuteSql sSql
      
      clsADOCon.CommitTrans
      
      MsgBox "The List Was Successfully Saved.", _
         vbInformation, Caption
      bGoodCount = GetCycleCount()
      'FillList
   Else
      clsADOCon.RollbackTrans
      
      MsgBox "Could Not Successfully Save The List.", _
         vbInformation, Caption
   End If
   
   Set RdoLst = Nothing
   
   Exit Sub
   
DiaErr1:
   sProcName = "savecurren"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub DeleteCurrentList()
   'On Error Resume Next
   'Delete Them
   sSql = "DELETE FROM CcitTable WHERE CIREF='" & Trim(cmbCid) & "'"
   clsADOCon.ExecuteSql sSql
   If Err = 0 Then
      sSql = "UPDATE CchdTable SET CCCOUNTSAVED=0 WHERE " _
             & "CCREF='" & Trim(cmbCid) & "'"
      clsADOCon.ExecuteSql sSql
      
      MsgBox "The List Was Successfully Cleared.", _
         vbInformation, Caption
      bGoodCount = GetCycleCount()
   Else
      MsgBox "Could Not Successfully Clear The List.", _
         vbInformation, Caption
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "deletecurre"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Grid1_Click()
   If Grid1.MouseCol = 0 Then
      Grid1.Col = COL_Include
      If Grid1.CellPicture = imgInc Then
         Set Grid1.CellPicture = imgdInc
      Else
         Set Grid1.CellPicture = imgInc
      End If
   End If
End Sub

Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 40)
   txtDsc = StrCase(txtDsc)
   'On Error Resume Next
   sSql = "UPDATE CchdTable SET CCDESC='" & txtDsc & "' WHERE " _
          & "CCREF='" & Trim(cmbCid) & "'"
   clsADOCon.ExecuteSql sSql
   
End Sub



Private Sub LockThisCount()
   Dim RdoLst As ADODB.Recordset
   Dim RdoLot As ADODB.Recordset
   Dim bLots As Byte
   Dim iRow As Integer
   'Dim iLots As Integer
   'Dim iCount As Integer
   
   'bLots = CheckLotTracking()
   'On Error Resume Next
   On Error GoTo DiaErr1
   clsADOCon.BeginTrans
   'Delete Them
   sSql = "DELETE FROM CcitTable WHERE CIREF='" & Trim(cmbCid) & "'"
   clsADOCon.ExecuteSql sSql
   
   sSql = "DELETE FROM CcltTable WHERE CLREF='" & Trim(cmbCid) & "'"
   clsADOCon.ExecuteSql sSql
   
   'Add Them to the list
   sSql = "SELECT * FROM CcitTable WHERE CIREF='" & Trim(cmbCid) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_KEYSET)
   With RdoLst
      
      For iRow = 1 To Grid1.Rows - 1
         Grid1.row = iRow
         Grid1.Col = COL_Include
         If Grid1.CellPicture = imgInc Then
            .AddNew
            !CIREF = Trim(cmbCid)
            !CIPARTREF = Compress(Grid1.TextMatrix(iRow, COL_PartRef))
            !CIPADESC = Grid1.TextMatrix(iRow, COL_PartDescription)
            !CIPALOCATION = Grid1.TextMatrix(iRow, COL_Location)
            !CILOTLOCATION = Grid1.TextMatrix(iRow, COL_LotLocation)
            
            '!CIPASTDCOST = CCur("0" & Grid1.TextMatrix(iRow, COL_StdCost))
            !CIPASTDCOST = String2Currency(Grid1.TextMatrix(iRow, COL_StdCost))
            
            '!CIPAQOH = CCur("0" & Grid1.TextMatrix(iRow, COL_QOH))
            !CIPAQOH = String2Currency(Grid1.TextMatrix(iRow, COL_QOH))
            
            !CILOTTRACK = Grid1.TextMatrix(iRow, COL_LotTracked)
            
            .Update
         End If
      Next
      
      ClearResultSet RdoLst
   End With
  
   'Add Lots where part is lot tracked
   bLots = CheckLotTracking()
   If bLots <> 0 Then
      sSql = "insert CcltTable(CLREF, CLPARTREF, CLLOTNUMBER, CLLOTREMAININGQTY)" & vbCrLf _
         & "select CIREF, CIPARTREF, LOTNUMBER, LOTREMAININGQTY" & vbCrLf _
         & "from CcitTable" & vbCrLf _
         & "join LohdTable on CIPARTREF = LOTPARTREF" & vbCrLf _
         & "join PartTable on CIPARTREF = PARTREF" & vbCrLf _
         & "where CIREF = '" & cmbCid & "'" & vbCrLf _
         & "and LOTREMAININGQTY > 0 AND LOTAVAILABLE = 1" & vbCrLf _
         & "and PALOTTRACK = 1" & vbCrLf _
         & "and CILOTLOCATION = lotlocation" & vbCrLf _
         & "order by CIPARTREF, LOTNUMBER"
         
         Debug.Print sSql
      clsADOCon.ExecuteSql sSql
   End If
   
   'add non-lot CcltTable records where none exist for part in cycle count
   sSql = "insert into CcltTable (CLREF,CLPARTREF,CLLOTNUMBER,CLLOTREMAININGQTY)" & vbCrLf _
      & "select CIREF, CIPARTREF, '', 0 from CcitTable" & vbCrLf _
      & "where not exists (select CLREF from CcltTable" & vbCrLf _
      & "  where CLREF = '" & cmbCid & "' and CLPARTREF = CIPARTREF)" & vbCrLf _
      & "and CIREF = '" & cmbCid & "'"

Debug.Print sSql

   clsADOCon.ExecuteSql sSql
   
   'double-check that each cclttable row has a ccittable row, and vice-versa.
   Dim msg As String
   Dim rdo As ADODB.Recordset
   sSql = "select CIPARTREF from CcitTable" & vbCrLf _
      & "left join cclttable on CIREF = CLREF and CIPARTREF = CLPARTREF" & vbCrLf _
      & "Where CLREF Is Null and CIREF = '" & cmbCid & "'"
   If clsADOCon.GetDataSet(sSql, rdo) Then
      Do Until rdo.EOF
         msg = msg & "Part " & rdo!CIPARTREF & " has no CcltTable record." & vbCrLf

         rdo.MoveNext
      Loop
   End If
   
   sSql = "select CLPARTREF from CcltTable" & vbCrLf _
      & "left join ccittable on CIREF = CLREF and CIPARTREF = CLPARTREF" & vbCrLf _
      & "Where CIREF Is Null and CLREF = '" & cmbCid & "'"
   If clsADOCon.GetDataSet(sSql, rdo) Then
      Do Until rdo.EOF
         msg = msg & "Part " & rdo!CLPARTREF & " has no CcitTable record." & vbCrLf

         rdo.MoveNext
      Loop
   End If
   
  'update the cycle count record status
   sSql = "UPDATE CchdTable SET CCCOUNTLOCKEDDATE='" _
          & Format(ES_SYSDATE, "mm/dd/yy") & "'," _
          & "CCCOUNTLOCKED=1 WHERE " _
          & "CCREF='" & Trim(cmbCid) & "'"
   clsADOCon.ExecuteSql sSql
   
   'if mismatched Ccit and Cclt, say so, and rollback
   If msg <> "" Then
      MsgBox msg & "Unable to proceed.", vbCritical
      clsADOCon.RollbackTrans
   Else
      clsADOCon.CommitTrans
      MsgBox msg & "The List Was Successfully Locked And Readied For Inventory.", _
         vbInformation, Caption
      FillCombo
   End If
   Set RdoLst = Nothing
   Set RdoLot = Nothing
   Exit Sub
   
DiaErr1:
   clsADOCon.RollbackTrans
   sProcName = "LockThisCount"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtPlan_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtPlan_LostFocus()
   txtPlan = CheckDate(txtPlan)
   'On Error Resume Next
   sSql = "UPDATE CchdTable SET CCPLANDATE='" & txtPlan & "' WHERE " _
          & "CCREF='" & Trim(cmbCid) & "'"
   clsADOCon.ExecuteSql sSql
   
End Sub



Private Sub MarkAll(Checked As Boolean)
   ' True = Select All
   ' False = Remove All
   
   Dim i As Long
   With Grid1
      For i = 1 To .Rows - 1
         .row = i
         .Col = COL_Include
         If Checked Then
            Set .CellPicture = imgInc
         Else
            Set .CellPicture = imgdInc
         End If
      Next
   End With
End Sub

Private Function OkToLock() As Boolean
'   sSql = "select RTRIM(CIPARTREF) as CIPARTREF, RTRIM(CIREF) as CIREF from CcitTable" & vbCrLf _
'      & "where CIREF = '" & Me.cmbCid & "'" & vbCrLf _
'      & "and CIPARTREF in (select CIPARTREF from CcitTable" & vbCrLf _
'      & "join CchdTable on CIREF = CCREF" & vbCrLf _
'      & "where CIREF <> '" & cmbCid & "'" & vbCrLf _
'      & "and CCCOUNTLOCKED = 1" & vbCrLf _
'      & "and CCUPDATED = 0)"

   sSql = "delete from CCLog where CCREF = '" & Me.cmbCid & "'"
   clsADOCon.ExecuteSql sSql
   
'   sSql = "select CIPARTREF, CIREF" & vbCrLf _
'      & "from CcitTable" & vbCrLf _
'      & "join CchdTable on CCREF = CIREF" & vbCrLf _
'      & "where CIREF <> '" & cmbCid & "'" & vbCrLf _
'      & "and CCCOUNTLOCKED = 1" & vbCrLf _
'      & "and CCUPDATED = 0" & vbCrLf _
'      & "and CIPARTREF in (select CIPARTREF from CcitTable" & vbCrLf _
'      & "where CIREF = '" & cmbCid & "')" & vbCrLf _
'      & "order by CIPARTREF, CIREF"
'
'   Dim rdo As rdoResultset
'   If GetDataSet(rdo) Then
'      Dim msg As String
'      Dim ct As Long
'      While Not rdo.EOF
'         ct = ct + 1
'         If ct <= 10 Then
'            msg = msg & "part " & Trim(rdo!CIPARTREF) & " is already on count sheet " _
'               & Trim(rdo!CIREF) & vbCrLf
'         End If
'
'         rdo.MoveNext
'      Wend
'      msg = msg & "There are a total of " & ct & " parts on other locked count sheets." & vbCrLf _
'         & "Unable to lock this count."
         
   sSql = "insert into CCLog (PARTREF, CCREF, LOTNUMBER, LOGTEXT)" & vbCrLf _
      & "select CIPARTREF, '" & Me.cmbCid & "', '', 'Part already locked on count sheet ' + rtrim(CIREF)" & vbCrLf _
      & "from CcitTable" & vbCrLf _
      & "join CchdTable on CCREF = CIREF" & vbCrLf _
      & "where CIREF <> '" & cmbCid & "'" & vbCrLf _
      & "and CCCOUNTLOCKED = 1" & vbCrLf _
      & "and CCUPDATED = 0" & vbCrLf _
      & "and CIPARTREF in (select CIPARTREF from CcitTable" & vbCrLf _
      & "where CIREF = '" & cmbCid & "')" & vbCrLf _
      & "order by CIPARTREF, CIREF"
   clsADOCon.ExecuteSql sSql

   sSql = "select count(*) from CCLog where CCREF = '" & Me.cmbCid & "'"
   Dim rdo As ADODB.Recordset
   Dim ct As Long
   If clsADOCon.GetDataSet(sSql, rdo) Then
      ct = rdo.Fields(0)
   End If
   
   If ct > 0 Then
      Dim msg As String
      msg = "There are a total of " & ct & " parts on other locked count sheets." & vbCrLf _
         & "Unable to lock this count.  See log."
      MsgBox msg
      OkToLock = False
   Else
      OkToLock = True
   End If
   Set rdo = Nothing
End Function

Private Function String2Currency(str As String) As Currency
   If IsNumeric(str) Then
      If InStr(1, str, "-") > 0 Then
         Debug.Print "negative: " & str
      End If
      String2Currency = CCur(str)
   Else
      String2Currency = 0
   End If
End Function

